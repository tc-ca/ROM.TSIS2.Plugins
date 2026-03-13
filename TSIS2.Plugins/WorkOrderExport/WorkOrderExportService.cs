using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace TSIS2.Plugins.WorkOrderExport
{
    /// <summary>
    /// Service for exporting Work Orders to PDF and bundling them into ZIP archives.
    /// This service can be used from both plugins and standalone test applications.
    /// </summary>
    public class WorkOrderExportService
    {
        private const int ERROR_MESSAGE_MAX_LENGTH = 4000;
        private const string TELEMETRY_TAG = "[WOExportTelemetry]";
        private const string TEMP_ZIP_FILE_NAME = "__WOEXPORT_TMP__.wip";
        private const string FINAL_ZIP_FILE_ATTRIBUTE_NAME = "ts_finalexportzip";
        private const string FINAL_ZIP_FILE_NAME_ATTRIBUTE_NAME = "ts_finalexportzip_name";
        private const string TEMP_ZIP_FILE_ATTRIBUTE_NAME = "ts_tempexportzip";
        private const string TEMP_ZIP_FILE_NAME_ATTRIBUTE_NAME = "ts_tempexportzip_name";
        private const int FILE_COLUMN_MAX_SIZE_KB = 131072; // ts_finalexportzip MaxSizeInKB
        private const int FILE_COLUMN_MAX_SIZE_BYTES = FILE_COLUMN_MAX_SIZE_KB * 1024;
        private const int MAX_SANDBOX_TEMP_ZIP_BYTES = FILE_COLUMN_MAX_SIZE_BYTES;

        private readonly IOrganizationService _service;
        private readonly ITracingService _tracingService;

        public WorkOrderExportService(IOrganizationService service, ITracingService tracingService)
        {
            _service = service ?? throw new ArgumentNullException(nameof(service));
            _tracingService = tracingService ?? throw new ArgumentNullException(nameof(tracingService));
        }

        public sealed class ArtifactValidationResult
        {
            public long TotalInputBytes { get; set; }
            public int TotalSurveyPdfs { get; set; }
            public int WorkOrdersValidated { get; set; }
        }

        /// <summary>
        /// Validates that MAIN and SURVEY PDF artifacts exist for all work orders using metadata-only queries.
        /// No PDF content is loaded — this is a fast pre-flight check that replaces the old merge phase.
        /// </summary>
        public ArtifactValidationResult ValidateAllArtifactsExist(Guid jobId, List<Guid> workOrderIds)
        {
            if (workOrderIds == null) throw new ArgumentNullException(nameof(workOrderIds));

            long totalInputBytes = 0;
            int totalSurveyPdfs = 0;

            for (int i = 0; i < workOrderIds.Count; i++)
            {
                var woId = workOrderIds[i];
                var estimate = EstimateMergeInputsForWorkOrder(jobId, woId);
                totalInputBytes += estimate.InputBytes;
                totalSurveyPdfs += estimate.SurveyPdfCount;
                _tracingService.Trace(
                    $"{TELEMETRY_TAG} ValidateArtifact jobId={jobId}, index={i + 1}/{workOrderIds.Count}, woId={woId}, inputBytes={estimate.InputBytes}, surveyPdfs={estimate.SurveyPdfCount}");
            }

            return new ArtifactValidationResult
            {
                TotalInputBytes = totalInputBytes,
                TotalSurveyPdfs = totalSurveyPdfs,
                WorkOrdersValidated = workOrderIds.Count
            };
        }

        public bool IsFinalZipPresent(Guid jobId)
        {
            var job = _service.Retrieve("ts_workorderexportjob", jobId, new ColumnSet(FINAL_ZIP_FILE_NAME_ATTRIBUTE_NAME));
            string fileName = (job.GetAttributeValue<string>(FINAL_ZIP_FILE_NAME_ATTRIBUTE_NAME) ?? string.Empty).Trim();
            return !string.IsNullOrWhiteSpace(fileName);
        }



        /// <summary>
        /// Processes as many work orders as the time budget allows in a single plugin invocation.
        /// Downloads the temp ZIP once, appends multiple WO merged PDFs, uploads once.
        /// This replaces the old ProcessZipBatch O(n²) loop pattern.
        /// </summary>
        public int ProcessZipBatchMulti(
            Guid jobId,
            List<Guid> workOrderIds,
            int nextZipIndex,
            IDictionary<Guid, string> workOrderNamesById,
            Stopwatch budgetStopwatch,
            int minRemainingMs)
        {
            if (workOrderIds == null) throw new ArgumentNullException(nameof(workOrderIds));

            int safeStart = Math.Max(0, Math.Min(workOrderIds.Count, nextZipIndex));
            if (safeStart >= workOrderIds.Count)
            {
                return safeStart;
            }

            EnsureUniqueZipEntryNames(workOrderIds, workOrderNamesById);

            if (budgetStopwatch != null)
            {
                long remainingMs = Math.Max(0, STAGE3_SAFE_BUDGET_MS - budgetStopwatch.ElapsedMilliseconds);
                if (remainingMs < minRemainingMs)
                {
                    TraceTelemetry(
                        $"ZipMultiYieldBeforeDownload jobId={jobId}, nextIndex={safeStart}, remainingMs={remainingMs}, minRemainingMs={minRemainingMs}");
                    return safeStart;
                }
            }

            var batchStopwatch = Stopwatch.StartNew();

            // Download temp ZIP once
            var existingTempZip = RetrieveFileFromFileColumn(jobId, TEMP_ZIP_FILE_ATTRIBUTE_NAME, TEMP_ZIP_FILE_NAME_ATTRIBUTE_NAME);
            byte[] existingBytes = existingTempZip?.Content ?? Array.Empty<byte>();

            if (existingBytes.Length > MAX_SANDBOX_TEMP_ZIP_BYTES)
            {
                throw new InvalidPluginExecutionException(
                    $"Sandbox ZIP size limit reached while building export. CurrentZIPBytes={existingBytes.Length}, LimitBytes={MAX_SANDBOX_TEMP_ZIP_BYTES}. Please export fewer work orders.");
            }

            var entryNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            int currentIndex = safeStart;
            int itemsProcessed = 0;

            using (var memoryStream = new MemoryStream())
            {
                if (existingBytes.Length > 0)
                {
                    memoryStream.Write(existingBytes, 0, existingBytes.Length);
                    memoryStream.Position = 0;
                }

                var mode = existingBytes.Length > 0 ? ZipArchiveMode.Update : ZipArchiveMode.Create;
                using (var archive = new ZipArchive(memoryStream, mode, true))
                {
                    if (mode == ZipArchiveMode.Update)
                    {
                        foreach (var existingEntry in archive.Entries)
                        {
                            entryNames.Add(existingEntry.FullName);
                        }
                    }

                    while (currentIndex < workOrderIds.Count
                        && (budgetStopwatch == null || (budgetStopwatch.ElapsedMilliseconds < STAGE3_SAFE_BUDGET_MS
                            && (STAGE3_SAFE_BUDGET_MS - budgetStopwatch.ElapsedMilliseconds) >= minRemainingMs)))
                    {
                        Guid woId = workOrderIds[currentIndex];
                        int oneBased = currentIndex + 1;
                        var itemStopwatch = Stopwatch.StartNew();

                        string workOrderName = null;
                        workOrderNamesById?.TryGetValue(woId, out workOrderName);

                        int surveyPdfCount = 0;
                        long surveyBytes = 0;
                        long mainBytes = 0;
                        long mergedBytesLength = 0;
                        bool fallbackUsed = false;
                        int fallbackPartCount = 0;

                        try
                        {
                            byte[] mergedBytes = BuildMergedPdfForWorkOrder(jobId, woId, out surveyPdfCount, out surveyBytes, out mainBytes);
                            mergedBytesLength = mergedBytes.Length;
                            string entryName = BuildDeterministicZipEntryPdfFileName(woId, workOrderName);
                            UpsertZipEntry(archive, entryNames, entryName, mergedBytes);
                        }
                        catch (OutOfMemoryException oomEx)
                        {
                            fallbackUsed = true;
                            var pdfParts = RetrievePdfPartsForWorkOrder(jobId, woId, out surveyPdfCount, out surveyBytes, out mainBytes);
                            fallbackPartCount = pdfParts.Count;

                            string basePdfName = BuildDeterministicZipEntryPdfFileName(woId, workOrderName);
                            string folderName = Path.GetFileNameWithoutExtension(basePdfName) ?? $"WO_{woId}";
                            string folderPrefix = $"{folderName}/";
                            DeleteZipEntriesByPrefix(archive, entryNames, folderPrefix);

                            for (int partIndex = 0; partIndex < pdfParts.Count; partIndex++)
                            {
                                var part = pdfParts[partIndex];
                                string safePartName = SanitizeFileName(Path.GetFileName(part.FileName) ?? $"part_{partIndex + 1:00}.pdf");
                                if (string.IsNullOrWhiteSpace(safePartName)) safePartName = $"part_{partIndex + 1:00}.pdf";
                                if (!safePartName.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase)) safePartName += ".pdf";
                                string fallbackEntryName = $"{folderPrefix}{partIndex + 1:00}_{safePartName}";
                                UpsertZipEntry(archive, entryNames, fallbackEntryName, part.Content);
                            }

                            _tracingService.Trace(
                                $"OOM_FALLBACK: JobId={jobId}, WO={woId}, Parts={fallbackPartCount}, MainBytes={mainBytes}, SurveyBytes={surveyBytes}, Message={oomEx.Message}");
                        }

                        itemStopwatch.Stop();
                        long inputBytes = mainBytes + surveyBytes;
                        currentIndex++;
                        itemsProcessed++;
                        TraceTelemetry(
                            $"ZipMultiItem jobId={jobId}, item={oneBased}/{workOrderIds.Count}, woId={woId}, surveyPdfCount={surveyPdfCount}, mainBytes={mainBytes}, surveyBytes={surveyBytes}, inputBytes={inputBytes}, mergedBytes={mergedBytesLength}, fallbackUsed={fallbackUsed}, fallbackPartCount={fallbackPartCount}, itemElapsedMs={itemStopwatch.ElapsedMilliseconds}, totalElapsedMs={batchStopwatch.ElapsedMilliseconds}");
                    }
                }

                byte[] updatedZipBytes = memoryStream.ToArray();
                if (updatedZipBytes.Length > MAX_SANDBOX_TEMP_ZIP_BYTES)
                {
                    throw new InvalidPluginExecutionException(
                        $"Sandbox ZIP size limit reached while building export. NextZIPBytes={updatedZipBytes.Length}, LimitBytes={MAX_SANDBOX_TEMP_ZIP_BYTES}. Please export fewer work orders.");
                }

                // Upload temp ZIP once
                SaveZipToFileColumn(jobId, updatedZipBytes, TEMP_ZIP_FILE_NAME, TEMP_ZIP_FILE_ATTRIBUTE_NAME);

                batchStopwatch.Stop();
                TraceTelemetry(
                    $"ZipMultiBatch jobId={jobId}, startIndex={safeStart}, endIndexExclusive={currentIndex}, itemsProcessed={itemsProcessed}, tempZipBytes={updatedZipBytes.Length}, elapsedMs={batchStopwatch.ElapsedMilliseconds}");
            }

            return currentIndex;
        }

        private const int STAGE3_SAFE_BUDGET_MS = 100000;

        public long EstimateTotalZipInputBytes(Guid jobId, IEnumerable<Guid> workOrderIds)
        {
            if (workOrderIds == null)
            {
                return 0;
            }

            long total = 0;
            foreach (var woId in workOrderIds)
            {
                total += EstimateInputBytesForWorkOrder(jobId, woId);
            }

            return total;
        }

        public int GetZipSizeLimitBytes()
        {
            return FILE_COLUMN_MAX_SIZE_BYTES;
        }

        private void UpsertZipEntry(ZipArchive archive, HashSet<string> entryNames, string entryName, byte[] content)
        {
            if (entryNames.Contains(entryName))
            {
                var existingEntry = archive.GetEntry(entryName);
                existingEntry?.Delete();
            }

            var entry = archive.CreateEntry(entryName, CompressionLevel.Fastest);
            using (var entryStream = entry.Open())
            {
                entryStream.Write(content, 0, content.Length);
            }

            entryNames.Add(entryName);
        }

        private void EnsureUniqueZipEntryNames(IEnumerable<Guid> workOrderIds, IDictionary<Guid, string> workOrderNamesById)
        {
            if (workOrderIds == null)
            {
                return;
            }

            var entryNameToWorkOrder = new Dictionary<string, Guid>(StringComparer.OrdinalIgnoreCase);
            foreach (var workOrderId in workOrderIds)
            {
                string workOrderName = null;
                workOrderNamesById?.TryGetValue(workOrderId, out workOrderName);
                string entryName = BuildDeterministicZipEntryPdfFileName(workOrderId, workOrderName);
                if (!entryNameToWorkOrder.TryGetValue(entryName, out Guid existingWorkOrderId))
                {
                    entryNameToWorkOrder[entryName] = workOrderId;
                    continue;
                }

                if (existingWorkOrderId == workOrderId)
                {
                    continue;
                }

                throw new InvalidPluginExecutionException(
                    $"Duplicate ZIP entry name detected for different work orders: '{entryName}'. WorkOrderA={existingWorkOrderId}, WorkOrderB={workOrderId}. Work Order Number must be unique.");
            }
        }

        private void DeleteZipEntriesByPrefix(ZipArchive archive, HashSet<string> entryNames, string entryPrefix)
        {
            if (string.IsNullOrWhiteSpace(entryPrefix))
            {
                return;
            }

            var toDelete = archive.Entries
                .Where(e => e.FullName.StartsWith(entryPrefix, StringComparison.OrdinalIgnoreCase))
                .Select(e => e.FullName)
                .ToList();

            foreach (var name in toDelete)
            {
                var entry = archive.GetEntry(name);
                entry?.Delete();
                entryNames.Remove(name);
            }
        }

        public void FinalizePersistedZip(Guid jobId)
        {
            var tempZip = RetrievePersistedTempZip(jobId);
            if (tempZip == null || tempZip.Value.Content == null || tempZip.Value.Content.Length == 0)
            {
                throw new InvalidPluginExecutionException("Temporary ZIP content is missing. Cannot finalize ZIP.");
            }

            string zipFileName = $"WorkOrder-Export-{DateTime.UtcNow:yyyy-MM-dd_HHmm}.zip";
            TraceTelemetry($"ZipFinalizeStart jobId={jobId}, zipBytes={tempZip.Value.Content.Length}");

            SaveZipToFileColumn(jobId, tempZip.Value.Content, zipFileName, FINAL_ZIP_FILE_ATTRIBUTE_NAME);
            ClearPersistedTempZip(jobId);
            TraceTelemetry($"ZipFinalizeDone jobId={jobId}, fileName={zipFileName}, zipBytes={tempZip.Value.Content.Length}");
        }

        public void ClearTemporaryZipStorage(Guid jobId)
        {
            ClearPersistedTempZip(jobId);
        }

        private (Guid AnnotationId, byte[] Content)? RetrievePersistedTempZip(Guid jobId)
        {
            return RetrieveFileFromFileColumn(jobId, TEMP_ZIP_FILE_ATTRIBUTE_NAME, TEMP_ZIP_FILE_NAME_ATTRIBUTE_NAME);
        }

        private void ClearPersistedTempZip(Guid jobId)
        {
            ClearFileColumn(jobId, TEMP_ZIP_FILE_ATTRIBUTE_NAME);
        }

        private byte[] BuildMergedPdfForWorkOrder(Guid jobId, Guid workOrderId, out int surveyPdfCount, out long surveyBytes, out long mainBytes)
        {
            var mainPdf = RetrieveAnnotation(jobId, $"WO_{workOrderId}_MAIN.pdf");
            if (mainPdf == null)
            {
                throw new InvalidPluginExecutionException($"Missing MAIN PDF for Work Order {workOrderId}. Expected file pattern: WO_{workOrderId}_MAIN.pdf attached to Job.");
            }

            var surveyPdfs = RetrieveSurveyPdfs(jobId, workOrderId);
            List<byte[]> pdfsToMerge = new List<byte[]> { mainPdf.Value.Content };
            pdfsToMerge.AddRange(surveyPdfs.Select(s => s.Content));
            surveyPdfCount = surveyPdfs.Count;
            surveyBytes = surveyPdfs.Sum(s => (long)s.Content.Length);
            mainBytes = mainPdf.Value.Content.Length;

            return new PdfSharpMerger().Merge(pdfsToMerge);
        }

        private long EstimateInputBytesForWorkOrder(Guid jobId, Guid workOrderId)
        {
            var query = new QueryExpression("annotation")
            {
                ColumnSet = new ColumnSet("filename", "filesize"),
                Criteria = new FilterExpression
                {
                    Conditions =
                    {
                        new ConditionExpression("objectid", ConditionOperator.Equal, jobId),
                        new ConditionExpression("isdocument", ConditionOperator.Equal, true)
                    },
                    Filters =
                    {
                        new FilterExpression(LogicalOperator.Or)
                        {
                            Conditions =
                            {
                                new ConditionExpression("filename", ConditionOperator.Equal, $"WO_{workOrderId}_MAIN.pdf"),
                                new ConditionExpression("filename", ConditionOperator.Like, $"WO_{workOrderId}_SURVEY_%")
                            }
                        }
                    }
                }
            };

            var results = _service.RetrieveMultiple(query);
            if (results.Entities.Count == 0)
            {
                throw new InvalidPluginExecutionException($"Missing PDF artifacts for Work Order {workOrderId}.");
            }

            long totalBytes = 0;
            bool foundMain = false;
            foreach (var note in results.Entities)
            {
                string fileName = note.GetAttributeValue<string>("filename") ?? string.Empty;
                if (string.Equals(fileName, $"WO_{workOrderId}_MAIN.pdf", StringComparison.OrdinalIgnoreCase))
                {
                    foundMain = true;
                }

                totalBytes += GetAnnotationSizeBytes(note);
            }

            if (!foundMain)
            {
                throw new InvalidPluginExecutionException($"Missing MAIN PDF for Work Order {workOrderId}. Expected file pattern: WO_{workOrderId}_MAIN.pdf attached to Job.");
            }

            return totalBytes;
        }

        private (long InputBytes, int SurveyPdfCount) EstimateMergeInputsForWorkOrder(Guid jobId, Guid workOrderId)
        {
            var query = new QueryExpression("annotation")
            {
                ColumnSet = new ColumnSet("filename", "filesize"),
                Criteria = new FilterExpression
                {
                    Conditions =
                    {
                        new ConditionExpression("objectid", ConditionOperator.Equal, jobId),
                        new ConditionExpression("isdocument", ConditionOperator.Equal, true)
                    },
                    Filters =
                    {
                        new FilterExpression(LogicalOperator.Or)
                        {
                            Conditions =
                            {
                                new ConditionExpression("filename", ConditionOperator.Equal, $"WO_{workOrderId}_MAIN.pdf"),
                                new ConditionExpression("filename", ConditionOperator.Like, $"WO_{workOrderId}_SURVEY_%")
                            }
                        }
                    }
                }
            };

            var results = _service.RetrieveMultiple(query);
            if (results.Entities.Count == 0)
            {
                throw new InvalidPluginExecutionException($"Missing PDF artifacts for Work Order {workOrderId}.");
            }

            bool foundMain = false;
            int surveyCount = 0;
            long totalBytes = 0;
            foreach (var note in results.Entities)
            {
                string fileName = note.GetAttributeValue<string>("filename") ?? string.Empty;
                if (string.Equals(fileName, $"WO_{workOrderId}_MAIN.pdf", StringComparison.OrdinalIgnoreCase))
                {
                    foundMain = true;
                }
                else if (fileName.StartsWith($"WO_{workOrderId}_SURVEY_", StringComparison.OrdinalIgnoreCase))
                {
                    surveyCount++;
                }

                totalBytes += GetAnnotationSizeBytes(note);
            }

            if (!foundMain)
            {
                throw new InvalidPluginExecutionException($"Missing MAIN PDF for Work Order {workOrderId}. Expected file pattern: WO_{workOrderId}_MAIN.pdf attached to Job.");
            }

            return (totalBytes, surveyCount);
        }

        private long GetAnnotationSizeBytes(Entity annotation)
        {
            if (annotation == null)
            {
                return 0;
            }

            var sizeInt = annotation.GetAttributeValue<int?>("filesize");
            if (sizeInt.HasValue && sizeInt.Value > 0)
            {
                return sizeInt.Value;
            }

            // Fallback only when filesize metadata is unavailable.
            string documentBody = annotation.GetAttributeValue<string>("documentbody");
            if (!string.IsNullOrEmpty(documentBody))
            {
                // Base64 size to bytes: floor((len * 3) / 4) minus padding.
                int padding = documentBody.EndsWith("==", StringComparison.Ordinal) ? 2
                    : documentBody.EndsWith("=", StringComparison.Ordinal) ? 1
                    : 0;
                return Math.Max(0, (documentBody.Length * 3L / 4L) - padding);
            }

            return 0;
        }

        private List<(string FileName, byte[] Content)> RetrievePdfPartsForWorkOrder(Guid jobId, Guid workOrderId, out int surveyPdfCount, out long surveyBytes, out long mainBytes)
        {
            var mainPdf = RetrieveAnnotation(jobId, $"WO_{workOrderId}_MAIN.pdf");
            if (mainPdf == null)
            {
                throw new InvalidPluginExecutionException($"Missing MAIN PDF for Work Order {workOrderId}. Expected file pattern: WO_{workOrderId}_MAIN.pdf attached to Job.");
            }

            var surveyPdfs = RetrieveSurveyPdfs(jobId, workOrderId);
            surveyPdfCount = surveyPdfs.Count;
            surveyBytes = surveyPdfs.Sum(s => (long)s.Content.Length);
            mainBytes = mainPdf.Value.Content.Length;

            var parts = new List<(string FileName, byte[] Content)>
            {
                (mainPdf.Value.FileName ?? $"WO_{workOrderId}_MAIN.pdf", mainPdf.Value.Content)
            };
            parts.AddRange(surveyPdfs);
            return parts;
        }

        public void CleanupIntermediateArtifacts(Guid jobId)
        {
            DeleteIntermediateAnnotations(jobId);
        }

        public void UpdateHeartbeat(Guid jobId, string progressMessage)
        {
            try
            {
                var update = new Entity("ts_workorderexportjob", jobId);
                update["ts_progressmessage"] = progressMessage;
                update["ts_lastheartbeat"] = DateTime.UtcNow;
                _service.Update(update);
            }
            catch (Exception ex)
            {
                _tracingService.Trace($"Warning: failed to update heartbeat/progress for job {jobId}. Message={ex.Message}");
            }
        }

        private string TryGetJobName(Guid jobId)
        {
            try
            {
                var job = _service.Retrieve("ts_workorderexportjob", jobId, new ColumnSet("ts_name"));
                return job.GetAttributeValue<string>("ts_name");
            }
            catch (Exception ex)
            {
                _tracingService.Trace($"Warning: Failed to read export job name for jobId={jobId}. Message={ex.Message}");
                return null;
            }
        }

        private static string BuildJobContext(Guid jobId, string jobName)
        {
            return string.IsNullOrWhiteSpace(jobName)
                ? $"jobId={jobId}, jobName=<not available>"
                : $"jobId={jobId}, jobName='{jobName}'";
        }

        private string TruncateForErrorField(string message)
        {
            string safeMessage = message ?? string.Empty;
            int maxLength = ERROR_MESSAGE_MAX_LENGTH;

            if (safeMessage.Length <= maxLength)
            {
                return safeMessage;
            }

            const string suffix = " ...[truncated]";
            int keepLength = Math.Max(0, maxLength - suffix.Length);
            string truncated = safeMessage.Substring(0, keepLength) + suffix;

            _tracingService.Trace($"ts_errormessage exceeded max length {maxLength}. Message was truncated.");
            return truncated;
        }

        private (string FileName, byte[] Content)? RetrieveAnnotation(Guid objectId, string exactFileName)
        {
            var query = new QueryExpression("annotation")
            {
                ColumnSet = new ColumnSet("filename", "documentbody"),
                Criteria = new FilterExpression
                {
                    Conditions =
                    {
                        new ConditionExpression("objectid", ConditionOperator.Equal, objectId),
                        new ConditionExpression("isdocument", ConditionOperator.Equal, true),
                        new ConditionExpression("filename", ConditionOperator.Equal, exactFileName)
                    }
                },
                TopCount = 1
            };

            var results = _service.RetrieveMultiple(query);
            if (results.Entities.Count > 0)
            {
                var note = results.Entities[0];
                string documentBody = note.GetAttributeValue<string>("documentbody");

                if (string.IsNullOrEmpty(documentBody))
                {
                    throw new InvalidPluginExecutionException(
                        $"Annotation '{exactFileName}' has empty documentbody.");
                }

                try
                {
                    byte[] content = Convert.FromBase64String(documentBody);
                    return (note.GetAttributeValue<string>("filename"), content);
                }
                catch (Exception ex)
                {
                    throw new InvalidPluginExecutionException(
                        $"Annotation '{exactFileName}' has invalid base64 content. {ex.Message}",
                        ex);
                }
            }
            return null;
        }

        private List<(string FileName, byte[] Content)> RetrieveSurveyPdfs(Guid jobId, Guid workOrderId)
        {
            // IMPORTANT: do not bulk-retrieve documentbody here.
            // Large note payloads can exceed sandbox message size when returned in a single RetrieveMultiple.
            // Fetch lightweight metadata first, then retrieve each note body individually.
            var query = new QueryExpression("annotation")
            {
                ColumnSet = new ColumnSet("annotationid", "filename"),
                Criteria = new FilterExpression
                {
                    Conditions =
            {
                new ConditionExpression("objectid", ConditionOperator.Equal, jobId),
                new ConditionExpression("isdocument", ConditionOperator.Equal, true),
                new ConditionExpression("filename", ConditionOperator.Like, $"WO_{workOrderId}_SURVEY_%")
            }
                }
                // No Orders here; we will sort in managed code by parsed GUID
            };

            var results = _service.RetrieveMultiple(query);
            if (results.Entities.Count == 0) return new List<(string, byte[])>();

            // Parse WOST GUID from filename and order deterministically
            var ordered = results.Entities
                .Select(n => new
                {
                    NoteId = n.Id,
                    FileName = n.GetAttributeValue<string>("filename") ?? string.Empty,
                    WostId = ExtractWostGuidFromFilename(n.GetAttributeValue<string>("filename"))
                })
                .OrderBy(x => x.WostId == Guid.Empty ? 1 : 0)
                .ThenBy(x => x.WostId)
                .ThenBy(x => x.FileName, StringComparer.OrdinalIgnoreCase)
                .ThenBy(x => x.NoteId)
                .ToList();

            var pdfs = new List<(string FileName, byte[] Content)>();
            foreach (var item in ordered)
            {
                var body = RetrieveAnnotationBody(item.NoteId);
                if (string.IsNullOrEmpty(body))
                {
                    throw new InvalidPluginExecutionException(
                        $"Survey PDF '{item.FileName}' has empty documentbody.");
                }
                try
                {
                    pdfs.Add((item.FileName, Convert.FromBase64String(body)));
                }
                catch (Exception ex)
                {
                    throw new InvalidPluginExecutionException(
                        $"Survey PDF '{item.FileName}' has invalid base64 content. {ex.Message}",
                        ex);
                }
            }

            return pdfs;
        }

        private string RetrieveAnnotationBody(Guid annotationId)
        {
            try
            {
                var note = _service.Retrieve("annotation", annotationId, new ColumnSet("documentbody"));
                return note.GetAttributeValue<string>("documentbody");
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(
                    $"Failed to retrieve annotation body for note {annotationId}. {ex.Message}",
                    ex);
            }
        }

        private static Guid ExtractWostGuidFromFilename(string filename)
        {
            // Expecting: WO_{WO}_SURVEY_{WOST}.pdf
            // Robust parse: find the last "_SURVEY_" and take the substring until ".pdf"
            if (string.IsNullOrWhiteSpace(filename)) return Guid.Empty;

            var idx = filename.LastIndexOf("_SURVEY_", StringComparison.OrdinalIgnoreCase);
            if (idx < 0) return Guid.Empty;

            var start = idx + "_SURVEY_".Length;
            var end = filename.LastIndexOf(".pdf", StringComparison.OrdinalIgnoreCase);
            if (end <= start) return Guid.Empty;

            var guidText = filename.Substring(start, end - start);
            return Guid.TryParse(guidText, out var g) ? g : Guid.Empty;
        }

        private (Guid AnnotationId, byte[] Content)? RetrieveFileFromFileColumn(Guid jobId, string fileAttributeName, string fileNameAttributeName)
        {
            var job = _service.Retrieve("ts_workorderexportjob", jobId, new ColumnSet(fileNameAttributeName));
            string fileName = (job.GetAttributeValue<string>(fileNameAttributeName) ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(fileName))
            {
                return null;
            }

            try
            {
                var initRequest = new OrganizationRequest("InitializeFileBlocksDownload");
                initRequest["Target"] = new EntityReference("ts_workorderexportjob", jobId);
                initRequest["FileAttributeName"] = fileAttributeName;
                var initResponse = _service.Execute(initRequest);
                string fileToken = initResponse["FileContinuationToken"]?.ToString();
                long fileSize = Convert.ToInt64(initResponse["FileSizeInBytes"]);

                if (string.IsNullOrWhiteSpace(fileToken) || fileSize <= 0)
                {
                    return null;
                }

                if (fileSize > MAX_SANDBOX_TEMP_ZIP_BYTES)
                {
                    throw new InvalidPluginExecutionException(
                        $"Sandbox ZIP size limit reached while resuming export. ExistingZIPBytes={fileSize}, LimitBytes={MAX_SANDBOX_TEMP_ZIP_BYTES}. Please export fewer work orders.");
                }

                const int blockSize = 4 * 1024 * 1024;
                using (var memoryStream = new MemoryStream((int)fileSize))
                {
                    long offset = 0;
                    while (offset < fileSize)
                    {
                        int length = (int)Math.Min(blockSize, fileSize - offset);
                        var downloadRequest = new OrganizationRequest("DownloadBlock");
                        downloadRequest["FileContinuationToken"] = fileToken;
                        downloadRequest["Offset"] = offset;
                        downloadRequest["BlockLength"] = (long)length;
                        var downloadResponse = _service.Execute(downloadRequest);
                        var blockData = (byte[])downloadResponse["Data"];
                        if (blockData == null || blockData.Length == 0)
                        {
                            break;
                        }

                        memoryStream.Write(blockData, 0, blockData.Length);
                        offset += blockData.Length;
                    }

                    return (Guid.Empty, memoryStream.ToArray());
                }
            }
            catch (InvalidPluginExecutionException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException($"Failed to read temporary ZIP from file column ({fileAttributeName}): {ex.Message}", ex);
            }
        }

        private void ClearFileColumn(Guid jobId, string fileAttributeName)
        {
            try
            {
                var job = _service.Retrieve("ts_workorderexportjob", jobId, new ColumnSet(fileAttributeName));
                if (!job.Attributes.TryGetValue(fileAttributeName, out object fileIdValue) || fileIdValue == null)
                {
                    return;
                }

                Guid? fileId = TryParseFileColumnId(fileIdValue);
                if (!fileId.HasValue || fileId.Value == Guid.Empty)
                {
                    _tracingService.Trace($"Warning: Could not parse file ID for '{fileAttributeName}' on job {jobId}. ValueType={fileIdValue.GetType().FullName}");
                    return;
                }

                var deleteRequest = new OrganizationRequest("DeleteFile");
                deleteRequest["FileId"] = fileId.Value;
                _service.Execute(deleteRequest);
            }
            catch (Exception ex)
            {
                _tracingService.Trace($"Warning: Failed to clear file column '{fileAttributeName}' on job {jobId}. Message={ex.Message}");
            }
        }

        private static Guid? TryParseFileColumnId(object fileIdValue)
        {
            if (fileIdValue is Guid guidValue)
            {
                return guidValue;
            }

            if (fileIdValue is string stringValue && Guid.TryParse(stringValue, out Guid parsedGuid))
            {
                return parsedGuid;
            }

            return null;
        }

        private string BuildDeterministicZipEntryPdfFileName(Guid workOrderId, string workOrderName)
        {
            // ZIP entry naming contract: use work order number only (e.g., 300-000677.pdf).
            // Keep GUID as last-resort fallback only when number is unavailable.
            string baseName = SanitizeFileName((workOrderName ?? string.Empty).Trim());
            if (string.IsNullOrWhiteSpace(baseName))
            {
                baseName = workOrderId.ToString();
            }

            return $"{baseName}.pdf";
        }

        private string SanitizeFileName(string fileName)
        {
            string invalidChars = new string(Path.GetInvalidFileNameChars());
            string sanitized = Regex.Replace(fileName, "[" + Regex.Escape(invalidChars) + "]", "");
            sanitized = sanitized.Trim();

            if (sanitized.Length > 120)
            {
                string ext = Path.GetExtension(sanitized);
                sanitized = sanitized.Substring(0, 120 - ext.Length) + ext;
            }

            return sanitized;
        }


        /// <summary>
        /// Saves ZIP file directly to the ts_finalexportzip file column using Dataverse File APIs
        /// </summary>
        private void SaveZipToFileColumn(Guid jobId, byte[] zipBytes, string fileName, string fileAttributeName)
        {
            _tracingService.Trace($"Saving ZIP to file column '{fileAttributeName}' for job {jobId}, size: {zipBytes.Length} bytes");
            TraceTelemetry($"FileUploadStart jobId={jobId}, fileAttribute={fileAttributeName}, fileName={fileName}, zipBytes={zipBytes.Length}");

            try
            {
                if (zipBytes.Length > FILE_COLUMN_MAX_SIZE_BYTES)
                {
                    throw new InvalidPluginExecutionException(
                        $"Final ZIP exceeds file column limit ({FILE_COLUMN_MAX_SIZE_BYTES} bytes). ZIPBytes={zipBytes.Length}. Please export fewer work orders.");
                }

                // Step 1: Initialize the file upload
                var initRequest = new OrganizationRequest("InitializeFileBlocksUpload");
                initRequest["Target"] = new EntityReference("ts_workorderexportjob", jobId);
                initRequest["FileAttributeName"] = fileAttributeName;
                initRequest["FileName"] = fileName;

                var initResponse = _service.Execute(initRequest);
                string fileId = initResponse["FileContinuationToken"].ToString();

                _tracingService.Trace($"File upload initialized with ID: {fileId}");
                TraceTelemetry($"FileUploadInitialized jobId={jobId}, fileTokenLength={fileId?.Length ?? 0}");

                // Step 2: Upload the file content in blocks
                // For files under 4MB, we can upload in a single block
                const int blockSize = 4 * 1024 * 1024; // 4MB chunks
                int totalBlocks = (int)Math.Ceiling((double)zipBytes.Length / blockSize);

                for (int blockNumber = 1; blockNumber <= totalBlocks; blockNumber++)
                {
                    int offset = (blockNumber - 1) * blockSize;
                    int length = Math.Min(blockSize, zipBytes.Length - offset);

                    byte[] blockData = new byte[length];
                    Array.Copy(zipBytes, offset, blockData, 0, length);

                    var uploadRequest = new OrganizationRequest("UploadBlock");
                    uploadRequest["FileContinuationToken"] = fileId;
                    uploadRequest["BlockId"] = Convert.ToBase64String(BitConverter.GetBytes(blockNumber));
                    uploadRequest["BlockData"] = blockData;

                    _service.Execute(uploadRequest);

                    _tracingService.Trace($"Uploaded block {blockNumber}/{totalBlocks} ({length} bytes)");
                    TraceTelemetry($"FileUploadBlock jobId={jobId}, block={blockNumber}/{totalBlocks}, bytes={length}");
                }

                // Step 3: Commit the upload
                var commitRequest = new OrganizationRequest("CommitFileBlocksUpload");
                commitRequest["FileContinuationToken"] = fileId;
                commitRequest["BlockList"] = GetBlockList(totalBlocks);
                commitRequest["FileName"] = fileName;
                commitRequest["MimeType"] = "application/zip";

                _service.Execute(commitRequest);

                _tracingService.Trace("File upload committed successfully");
                TraceTelemetry($"FileUploadCommit jobId={jobId}, blockCount={totalBlocks}, zipBytes={zipBytes.Length}");
            }
            catch (Exception ex)
            {
                _tracingService.Trace($"Error saving ZIP to file column: {ex.Message}");
                TraceTelemetry($"FileUploadError jobId={jobId}, message={ex.Message}");
                throw new InvalidPluginExecutionException($"Failed to save ZIP to file column: {ex.Message}", ex);
            }
        }

        private void TraceTelemetry(string message)
        {
            _tracingService.Trace($"{TELEMETRY_TAG} {message}");
        }

        /// <summary>
        /// Generates the block list for file upload commit
        /// </summary>
        private string[] GetBlockList(int totalBlocks)
        {
            var blockList = new string[totalBlocks];
            for (int i = 1; i <= totalBlocks; i++)
            {
                blockList[i - 1] = Convert.ToBase64String(BitConverter.GetBytes(i));
            }
            return blockList;
        }

        /// <summary>
        /// Deletes all intermediate PDF annotations (Survey, Main, Merged) from the job record
        /// </summary>
        private void DeleteIntermediateAnnotations(Guid jobId)
        {
            _tracingService.Trace($"Starting cleanup of intermediate annotations for job {jobId}");

            try
            {
                // Query all annotations attached to the job
                var query = new QueryExpression("annotation")
                {
                    ColumnSet = new ColumnSet("annotationid", "filename"),
                    Criteria = new FilterExpression
                    {
                        Conditions =
                        {
                            new ConditionExpression("objectid", ConditionOperator.Equal, jobId),
                            new ConditionExpression("isdocument", ConditionOperator.Equal, true)
                        }
                    }
                };

                var results = _service.RetrieveMultiple(query);
                var deleteRequests = new List<DeleteRequest>();
                int candidateCount = 0;

                foreach (var note in results.Entities)
                {
                    string fileName = note.GetAttributeValue<string>("filename") ?? string.Empty;
                    if (ShouldDeleteIntermediateAnnotation(fileName))
                    {
                        candidateCount++;
                        deleteRequests.Add(new DeleteRequest
                        {
                            Target = new EntityReference("annotation", note.Id)
                        });
                    }
                }

                if (deleteRequests.Count == 0)
                {
                    _tracingService.Trace("Cleanup completed. No intermediate annotations matched delete criteria.");
                    return;
                }

                const int chunkSize = 200;
                int deletedCount = 0;
                int failedCount = 0;
                int missingCount = 0;
                int concurrentDeleteCount = 0;

                for (int start = 0; start < deleteRequests.Count; start += chunkSize)
                {
                    var batch = new OrganizationRequestCollection();
                    int endExclusive = Math.Min(start + chunkSize, deleteRequests.Count);
                    for (int i = start; i < endExclusive; i++)
                    {
                        batch.Add(deleteRequests[i]);
                    }

                    var execute = new ExecuteMultipleRequest
                    {
                        Requests = batch,
                        Settings = new ExecuteMultipleSettings
                        {
                            ContinueOnError = true,
                            ReturnResponses = true
                        }
                    };

                    var response = (ExecuteMultipleResponse)_service.Execute(execute);
                    int batchFailures = 0;
                    foreach (var item in response.Responses)
                    {
                        if (item.Fault == null)
                        {
                            continue;
                        }

                        batchFailures++;
                        if (IsBenignMissingDeleteFault(item.Fault))
                        {
                            missingCount++;
                            continue;
                        }

                        if (IsBenignConcurrentDeleteFault(item.Fault))
                        {
                            concurrentDeleteCount++;
                            continue;
                        }

                        _tracingService.Trace(
                            $"Warning: Unexpected annotation cleanup failure. RequestIndex={item.RequestIndex}, ErrorCode={item.Fault.ErrorCode}, Message={item.Fault.Message}");
                    }

                    int batchCount = endExclusive - start;
                    failedCount += batchFailures;
                    deletedCount += Math.Max(0, batchCount - batchFailures);
                }

                _tracingService.Trace(
                    $"Cleanup completed. Candidates={candidateCount}, Deleted={deletedCount}, Missing={missingCount}, ConcurrentDelete={concurrentDeleteCount}, UnexpectedFailures={Math.Max(0, failedCount - missingCount - concurrentDeleteCount)}, Failed={failedCount}, BatchSize={chunkSize}.");
            }
            catch (Exception ex)
            {
                _tracingService.Trace($"Error during annotation cleanup: {ex.Message}");
                throw new InvalidPluginExecutionException(
                    $"Failed to cleanup intermediate annotations for job {jobId}. {ex.Message}",
                    ex);
            }
        }

        private static bool ShouldDeleteIntermediateAnnotation(string fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName))
            {
                return false;
            }

            if (fileName.Contains("_SURVEY_") && fileName.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (fileName.IndexOf("_MAIN.pdf", StringComparison.OrdinalIgnoreCase) >= 0
                || fileName.IndexOf("_MERGED.pdf", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return true;
            }

            if (string.Equals(fileName, TEMP_ZIP_FILE_NAME, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            return false;
        }

        private static bool IsBenignMissingDeleteFault(OrganizationServiceFault fault)
        {
            return fault != null && fault.ErrorCode == -2147220969;
        }

        private static bool IsBenignConcurrentDeleteFault(OrganizationServiceFault fault)
        {
            return fault != null && fault.ErrorCode == -2147188475;
        }

        public class PdfSharpMerger
        {
            public byte[] Merge(IEnumerable<byte[]> pdfs)
            {
                using (var outputDocument = new PdfSharp.Pdf.PdfDocument())
                {
                    foreach (var pdfBytes in pdfs)
                    {
                        using (var inputStream = new MemoryStream(pdfBytes))
                        {
                            var inputDocument = PdfSharp.Pdf.IO.PdfReader.Open(inputStream, PdfSharp.Pdf.IO.PdfDocumentOpenMode.Import);
                            int count = inputDocument.PageCount;
                            for (int idx = 0; idx < count; idx++)
                            {
                                var page = inputDocument.Pages[idx];
                                outputDocument.AddPage(page);
                            }
                        }
                    }

                    using (var outStream = new MemoryStream())
                    {
                        outputDocument.Save(outStream, false);
                        return outStream.ToArray();
                    }
                }
            }
        }
    }
}
