using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Xrm.Sdk;
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
        // StatusCode Values for ts_workorderexportjob
        private const int STATUS_COMPLETED = 741130006; // ZIP created
        private const int STATUS_ERROR = 741130007;

        private readonly IOrganizationService _service;
        private readonly ITracingService _tracingService;
        private readonly Guid _wordTemplateId;
        private readonly bool _testMode;

        public WorkOrderExportService(IOrganizationService service, ITracingService tracingService, Guid wordTemplateId = default, bool testMode = false)
        {
            _service = service ?? throw new ArgumentNullException(nameof(service));
            _tracingService = tracingService ?? throw new ArgumentNullException(nameof(tracingService));
            _wordTemplateId = wordTemplateId;
            _testMode = testMode;
        }

        /// <summary>
        /// Orchestrates the retrieval of MAIN and SURVEY PDFs for the given Work Orders, 
        /// merges them into a single PDF per Work Order, and then creates a ZIP archive of all merged PDFs.
        /// Update the Job status upon completion or error.
        /// </summary>
        /// <param name="jobId">The Work Order Export Job ID</param>
        /// <param name="workOrderIds">List of Work Order IDs to process</param>
        /// <param name="useFileStorage">If true, saves ZIP to ts_finalexportzip file column; if false, saves as annotation</param>
        public void ProcessMergeAndZip(Guid jobId, List<Guid> workOrderIds, bool useFileStorage = false)
        {
            _tracingService.Trace($"Starting Export Job Processing (Merge) for Job: {jobId}");
            _tracingService.Trace($"Storage mode: {(useFileStorage ? "File Column" : "Annotation")}");

            try
            {
                _tracingService.Trace($"WorkOrder count={workOrderIds.Count}");

                // 3. Process Work Orders (Merge PDFs)
                var pdfMerger = new PdfSharpMerger();

                List<(string FileName, byte[] Content)> mergedPdfs = new List<(string, byte[])>();

                foreach (var woId in workOrderIds)
                {
                    _tracingService.Trace($"[WO:{woId}] Starting merge sequence.");

                    // A. Retrieve MAIN PDF
                    var mainPdf = RetrieveAnnotation(jobId, $"WO_{woId}_MAIN.pdf");

                    if (mainPdf == null)
                    {
                        // Fail if main PDF is missing
                        throw new InvalidPluginExecutionException($"Missing MAIN PDF for Work Order {woId}. Expected file pattern: WO_{woId}_MAIN.pdf attached to Job.");
                    }

                    _tracingService.Trace($"[WO:{woId}] MAIN PDF found.");

                    // B. Retrieve SURVEY PDFs
                    var surveyPdfs = RetrieveSurveyPdfs(jobId, woId);

                    _tracingService.Trace($"[WO:{woId}] Survey PDFs count={surveyPdfs.Count}");

                    // C. Merge
                    List<byte[]> pdfsToMerge = new List<byte[]> { mainPdf.Value.Content };
                    pdfsToMerge.AddRange(surveyPdfs.Select(s => s.Content));

                    _tracingService.Trace($"Merging {pdfsToMerge.Count} PDFs for WO {woId}");
                    byte[] mergedBytes = pdfMerger.Merge(pdfsToMerge);

                    _tracingService.Trace($"[WO:{woId}] Merge completed. Size={mergedBytes.Length} bytes");

                    // D. Attach Merged PDF to Job
                    string mergedFileName = $"WO_{woId}_MERGED.pdf";
                    CreateAnnotation(jobId, mergedFileName, "application/pdf", mergedBytes);

                    mergedPdfs.Add((mergedFileName, mergedBytes));
                }

                // 4. Create ZIP
                _tracingService.Trace($"Zipping {mergedPdfs.Count} files...");
                byte[] zipBytes = CreateZip(mergedPdfs);

                _tracingService.Trace($"ZIP created successfully. Size={zipBytes.Length} bytes");

                string zipFileName = $"WorkOrderExport_{jobId}.zip";

                // 5. Save ZIP based on storage mode
                if (useFileStorage)
                {
                    _tracingService.Trace("Using File Storage mode for final ZIP");
                    SaveZipToFileColumn(jobId, zipBytes, zipFileName);
                }
                else
                {
                    _tracingService.Trace("Using Annotation mode for final ZIP");
                    CreateAnnotation(jobId, zipFileName, "application/zip", zipBytes);
                }

                // 6. Cleanup intermediate annotations (Survey, Main, Merged PDFs)
                DeleteIntermediateAnnotations(jobId, workOrderIds, useFileStorage);

                // 7. Update Status to Completed
                _tracingService.Trace($"Updating statuscode to COMPLETED ({STATUS_COMPLETED})");
                Entity updateJob = new Entity("ts_workorderexportjob", jobId);
                updateJob["statuscode"] = new OptionSetValue(STATUS_COMPLETED);
                updateJob["ts_errormessage"] = string.Empty; // Clear errors
                _service.Update(updateJob);

                _tracingService.Trace("Job Completed Successfully.");
            }
            catch (Exception ex)
            {
                _tracingService.Trace($"ERROR during ReadyForMerge processing. JobId={jobId}");
                _tracingService.Trace($"Job Failed: {ex.Message}");

                // Cleanup intermediate artifacts on failure
                try
                {
                    DeleteIntermediateAnnotations(jobId, workOrderIds, useFileStorage);
                }
                catch (Exception cleanupEx)
                {
                    _tracingService.Trace($"Warning: Cleanup failed during error handling: {cleanupEx.Message}");
                }

                Entity errorJob = new Entity("ts_workorderexportjob", jobId);
                errorJob["statuscode"] = new OptionSetValue(STATUS_ERROR);
                errorJob["ts_errormessage"] = $"Processing Failed: {ex.Message}\nStack: {ex.StackTrace}";
                _service.Update(errorJob);

                // Rethrow if you want the caller to know (like in Console App), 
                // but usually we swallow in async plugins. 
                // However, for Shared Service, let's rethrow so the caller can handle/log if needed,
                // or just rely on the job update.
                // Since we updated the Job to Error, we can swallow or rethrow. 
                // Let's rethrow to be safe for Console App debugging.
                throw;
            }
        }

        /// <summary>
        /// Backward-compatible overload that defaults to annotation storage mode
        /// </summary>
        public void ProcessMergeAndZip(Guid jobId, List<Guid> workOrderIds)
        {
            ProcessMergeAndZip(jobId, workOrderIds, useFileStorage: false);
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
                    _tracingService.Trace($"Warning: Annotation {exactFileName} has empty documentbody.");
                    return null;
                }

                try
                {
                    byte[] content = Convert.FromBase64String(documentBody);
                    return (note.GetAttributeValue<string>("filename"), content);
                }
                catch (Exception ex)
                {
                    _tracingService.Trace($"Error decoding base64 for {exactFileName}: {ex.Message}");
                    return null;
                }
            }
            return null;
        }

        private List<(string FileName, byte[] Content)> RetrieveSurveyPdfs(Guid jobId, Guid workOrderId)
        {
            // Fetch all SURVEY PDFs for this WO
            var query = new QueryExpression("annotation")
            {
                ColumnSet = new ColumnSet("filename", "documentbody"),
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
                    Note = n,
                    FileName = n.GetAttributeValue<string>("filename") ?? string.Empty,
                    WostId = ExtractWostGuidFromFilename(n.GetAttributeValue<string>("filename"))
                })
                .OrderBy(x => x.WostId) // GUID ascending
                .ToList();

            var pdfs = new List<(string FileName, byte[] Content)>();
            foreach (var item in ordered)
            {
                var body = item.Note.GetAttributeValue<string>("documentbody");
                if (string.IsNullOrEmpty(body))
                {
                    _tracingService.Trace($"Warning: Survey PDF {item.FileName} has empty documentbody. Skipping.");
                    continue;
                }
                try
                {
                    pdfs.Add((item.FileName, Convert.FromBase64String(body)));
                }
                catch (Exception ex)
                {
                    _tracingService.Trace($"Error decoding base64 for survey PDF {item.FileName}: {ex.Message}. Skipping.");
                }
            }

            return pdfs;
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


        private void CreateAnnotation(Guid parentId, string fileName, string mimeType, byte[] content)
        {
            Entity note = new Entity("annotation");
            note["objectid"] = new EntityReference("ts_workorderexportjob", parentId);
            note["subject"] = $"Generated: {fileName}";
            note["filename"] = fileName;
            note["mimetype"] = mimeType;
            note["isdocument"] = true;
            note["documentbody"] = Convert.ToBase64String(content);
            _service.Create(note);
        }

        private byte[] CreateZip(List<(string FileName, byte[] Content)> files)
        {
            using (var memoryStream = new MemoryStream())
            {
                using (var archive = new ZipArchive(memoryStream, ZipArchiveMode.Create, true))
                {
                    foreach (var file in files)
                    {
                        var entry = archive.CreateEntry(file.FileName);
                        using (var entryStream = entry.Open())
                        {
                            entryStream.Write(file.Content, 0, file.Content.Length);
                        }
                    }
                }
                return memoryStream.ToArray();
            }
        }

        /// <summary>
        /// Builds a single text blob containing all details for the specified work orders.
        /// (Legacy/Reference Implementation)
        /// </summary>
        public string BuildExportText(IEnumerable<Guid> workOrderIds)
        {
            if (workOrderIds == null || !workOrderIds.Any())
            {
                _tracingService.Trace("No work order IDs provided for text export");
                return string.Empty;
            }

            var sb = new StringBuilder();
            var retriever = new WorkOrderDataRetriever(_service, _tracingService);

            foreach (var workOrderId in workOrderIds)
            {
                try
                {
                    _tracingService.Trace($"Retrieving data for Work Order: {workOrderId}...");
                    var data = retriever.RetrieveWorkOrderData(workOrderId);

                    if (sb.Length > 0)
                    {
                        sb.AppendLine("----------------------------------------");
                        sb.AppendLine("");
                    }

                    FormatWorkOrderText(sb, data);
                }
                catch (Exception ex)
                {
                    _tracingService.Trace($"Failed to retrieve data for work order {workOrderId}: {ex.Message}");
                    sb.AppendLine($"Error retrieving Work Order {workOrderId}: {ex.Message}");
                }
            }

            return sb.ToString();
        }

        private void FormatWorkOrderText(StringBuilder sb, WorkOrderData data)
        {
            var summary = data.WorkOrderSummary;
            sb.AppendLine($"Work Order: {summary.Name} ({summary.WorkOrderId})");
            sb.AppendLine($"Region: {summary.Region?.Name ?? "N/A"}");
            sb.AppendLine($"Operation Type: {summary.OperationType?.Name ?? "N/A"}");
            sb.AppendLine($"Stakeholder: {summary.Stakeholder?.Name ?? "N/A"}");
            sb.AppendLine($"Site: {summary.Site?.Name ?? "N/A"}");

            sb.AppendLine("");
            sb.AppendLine("Findings:");
            if (data.Findings != null && data.Findings.Any())
            {
                foreach (var f in data.Findings)
                {
                    sb.AppendLine($"- {f.Finding} (Type: {f.FindingType?.Value}, Status: {f.StatusCode?.Value})");
                }
            }
            else
            {
                sb.AppendLine("- None");
            }

            sb.AppendLine("");
            sb.AppendLine("Actions:");
            if (data.Actions != null && data.Actions.Any())
            {
                foreach (var a in data.Actions)
                {
                    sb.AppendLine($"- {a.Name} (Type: {a.ActionType?.Value})");
                }
            }
            else
            {
                sb.AppendLine("- None");
            }

            sb.AppendLine("");
            sb.AppendLine("Interactions:");
            if (data.Interactions != null && data.Interactions.Any())
            {
                foreach (var i in data.Interactions)
                {
                    sb.AppendLine($"- {i.Subject} ({i.CreatedOn:yyyy-MM-dd}): {i.Description}");
                }
            }
            else
            {
                sb.AppendLine("- None");
            }

            sb.AppendLine("");
            sb.AppendLine("Attachments:");
            if (data.Documents != null)
            {
                var docs = new List<DocumentData>();
                if (data.Documents.GeneralDocuments != null) docs.AddRange(data.Documents.GeneralDocuments);
                if (data.Documents.InspectionDocuments != null) docs.AddRange(data.Documents.InspectionDocuments);

                if (docs.Any())
                {
                    foreach (var d in docs)
                    {
                        sb.AppendLine($"- {d.Name} ({d.FileCategory?.Name ?? "No Category"})");
                    }
                }
                else
                {
                    sb.AppendLine("- None");
                }
            }
            else
            {
                sb.AppendLine("- None");
            }
        }

        /// <summary>
        /// Generates a ZIP file containing PDFs for the specified work orders.
        /// (Legacy/Reference Implementation relying on Action)
        /// </summary>
        public byte[] GenerateZipForWorkOrders(List<Guid> workOrderIds)
        {
            if (workOrderIds == null || !workOrderIds.Any())
            {
                _tracingService.Trace("No work order IDs provided for export");
                return Array.Empty<byte>();
            }

            _tracingService.Trace($"Starting ZIP generation for {workOrderIds.Count} work orders");
            var pdfFiles = new List<(string FileName, byte[] Content)>();

            foreach (var workOrderId in workOrderIds)
            {
                try
                {
                    _tracingService.Trace($"Validating work order: {workOrderId}");

                    var wo = _service.Retrieve("msdyn_workorder", workOrderId,
                        new ColumnSet("msdyn_name", "ownerid", "createdon"));

                    string woName = wo.GetAttributeValue<string>("msdyn_name");
                    EntityReference ownerRef = wo.GetAttributeValue<EntityReference>("ownerid");
                    DateTime createdOn = wo.GetAttributeValue<DateTime>("createdon");

                    string ownerDisplayName = ResolveOwnerName(ownerRef);
                    string fileName = BuildPdfFileName(woName, ownerDisplayName, createdOn);

                    _tracingService.Trace($"Processing Work Order: {woName}, Owner: {ownerDisplayName}, File: {fileName}");

                    _tracingService.Trace($"Sending Work Order {workOrderId} to ExportWordDocumentToPDF action");
                    byte[] pdfContent = ExportWorkOrderToPdf(workOrderId);
                    pdfFiles.Add((fileName, pdfContent));

                    _tracingService.Trace($"Successfully exported: {fileName}");
                }
                catch (Exception ex)
                {
                    _tracingService.Trace($"Failed to export work order {workOrderId}: {ex.Message}");
                    // Continue processing other work orders
                }
            }

            if (!pdfFiles.Any())
            {
                _tracingService.Trace("No PDFs were successfully generated");
                return Array.Empty<byte>();
            }

            _tracingService.Trace($"Creating ZIP archive with {pdfFiles.Count} PDFs");
            return CreateZipArchive(pdfFiles);
        }

        public string CreateZipAnnotation(byte[] zipBytes, Guid relatedEntityId, string relatedEntityName, int workOrderCount)
        {
            if (zipBytes == null || zipBytes.Length == 0)
            {
                throw new ArgumentException("ZIP bytes cannot be null or empty", nameof(zipBytes));
            }

            string zipFileName = $"WorkOrders_{workOrderCount}_{DateTime.UtcNow:yyyy-MM-dd_HHmm}.zip";
            _tracingService.Trace($"Creating annotation with file: {zipFileName}");

            var note = new Entity("annotation");
            note["subject"] = "Work Order Export ZIP";
            note["filename"] = zipFileName;
            note["mimetype"] = "application/zip";
            note["documentbody"] = Convert.ToBase64String(zipBytes);
            note["objectid"] = new EntityReference(relatedEntityName, relatedEntityId);

            Guid noteId = _service.Create(note);
            _tracingService.Trace($"Created Annotation: {noteId}");
            return noteId.ToString();
        }

        private string ResolveOwnerName(EntityReference ownerRef)
        {
            if (ownerRef == null)
            {
                _tracingService.Trace("Owner reference is null");
                return "Unknown";
            }

            try
            {
                if (ownerRef.LogicalName == "systemuser")
                {
                    var user = _service.Retrieve("systemuser", ownerRef.Id,
                        new ColumnSet("firstname", "lastname"));
                    string firstName = user.GetAttributeValue<string>("firstname") ?? "";
                    string lastName = user.GetAttributeValue<string>("lastname") ?? "";
                    return $"{firstName} {lastName}".Trim();
                }
                else if (ownerRef.LogicalName == "team")
                {
                    var team = _service.Retrieve("team", ownerRef.Id, new ColumnSet("name"));
                    return team.GetAttributeValue<string>("name");
                }
            }
            catch (Exception ex)
            {
                _tracingService.Trace($"Failed to resolve owner name: {ex.Message}");
            }

            return "Unknown";
        }

        private string BuildPdfFileName(string woName, string ownerDisplayName, DateTime createdOn)
        {
            string raw = $"{woName}_{ownerDisplayName}_{createdOn:yyyy-MM-dd_HHmm}.pdf";
            return SanitizeFileName(raw);
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

        private byte[] ExportWorkOrderToPdf(Guid workOrderId)
        {
            if (_testMode || _wordTemplateId == Guid.Empty)
            {
                _tracingService.Trace($"Test mode or empty template ID - generating mock PDF for {workOrderId}");
                return GenerateMockPdf(workOrderId);
            }

            _tracingService.Trace($"Exporting work order {workOrderId} to PDF using template {_wordTemplateId}");

            var req = new OrganizationRequest("ExportWordDocumentToPDF");
            req["EntityTypeCode"] = "msdyn_workorder";
            req["SelectedRecords"] = new[] { workOrderId.ToString() };
            req["DocumentTemplateId"] = _wordTemplateId;

            try
            {
                var resp = _service.Execute(req);
                return (byte[])resp["PdfFile"];
            }
            catch (Exception ex)
            {
                _tracingService.Trace($"Error calling ExportWordDocumentToPDF for Work Order {workOrderId}: {ex.Message}");
                _tracingService.Trace("Falling back to mock PDF generation");
                return GenerateMockPdf(workOrderId);
            }
        }

        private byte[] GenerateMockPdf(Guid workOrderId)
        {
            string mockContent = $"MOCK PDF for Work Order {workOrderId}\nGenerated at {DateTime.UtcNow:yyyy-MM-dd HH:mm:ss} UTC\n\nThis is a test PDF placeholder.";
            return System.Text.Encoding.UTF8.GetBytes(mockContent);
        }

        private byte[] CreateZipArchive(List<(string FileName, byte[] Content)> pdfFiles)
        {
            using (var ms = new MemoryStream())
            {
                using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
                {
                    foreach (var pdf in pdfFiles)
                    {
                        _tracingService.Trace($"Adding to ZIP: {pdf.FileName}");
                        var entry = archive.CreateEntry(pdf.FileName, CompressionLevel.Fastest);
                        using (var entryStream = entry.Open())
                        {
                            entryStream.Write(pdf.Content, 0, pdf.Content.Length);
                        }
                    }
                }

                _tracingService.Trace($"ZIP archive created successfully. Size: {ms.Length} bytes");
                return ms.ToArray();
            }
        }

        /// <summary>
        /// Saves ZIP file directly to the ts_finalexportzip file column using Dataverse File APIs
        /// </summary>
        private void SaveZipToFileColumn(Guid jobId, byte[] zipBytes, string fileName)
        {
            _tracingService.Trace($"Saving ZIP to file column for job {jobId}, size: {zipBytes.Length} bytes");

            try
            {
                // Step 1: Initialize the file upload
                var initRequest = new OrganizationRequest("InitializeFileBlocksUpload");
                initRequest["Target"] = new EntityReference("ts_workorderexportjob", jobId);
                initRequest["FileAttributeName"] = "ts_finalexportzip";
                initRequest["FileName"] = fileName;

                var initResponse = _service.Execute(initRequest);
                string fileId = initResponse["FileContinuationToken"].ToString();

                _tracingService.Trace($"File upload initialized with ID: {fileId}");

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
                }

                // Step 3: Commit the upload
                var commitRequest = new OrganizationRequest("CommitFileBlocksUpload");
                commitRequest["FileContinuationToken"] = fileId;
                commitRequest["BlockList"] = GetBlockList(totalBlocks);
                commitRequest["FileName"] = fileName;
                commitRequest["MimeType"] = "application/zip";

                _service.Execute(commitRequest);

                _tracingService.Trace("File upload committed successfully");
            }
            catch (Exception ex)
            {
                _tracingService.Trace($"Error saving ZIP to file column: {ex.Message}");
                throw new InvalidPluginExecutionException($"Failed to save ZIP to file column: {ex.Message}", ex);
            }
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
        private void DeleteIntermediateAnnotations(Guid jobId, List<Guid> workOrderIds, bool useFileStorage)
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
                int deletedCount = 0;

                foreach (var note in results.Entities)
                {
                    string fileName = note.GetAttributeValue<string>("filename") ?? string.Empty;
                    bool shouldDelete = false;

                    // Delete Survey PDFs
                    if (fileName.Contains("_SURVEY_") && fileName.EndsWith(".pdf"))
                    {
                        shouldDelete = true;
                    }
                    // Delete Main PDFs
                    else if (fileName.Contains("_MAIN.pdf"))
                    {
                        shouldDelete = true;
                    }
                    // Delete Merged PDFs
                    else if (fileName.Contains("_MERGED.pdf"))
                    {
                        shouldDelete = true;
                    }
                    // If using file storage, also delete ZIP annotations (keep only file column)
                    else if (useFileStorage && fileName.StartsWith("WorkOrderExport_") && fileName.EndsWith(".zip"))
                    {
                        shouldDelete = true;
                    }

                    if (shouldDelete)
                    {
                        try
                        {
                            _service.Delete("annotation", note.Id);
                            deletedCount++;
                            _tracingService.Trace($"Deleted annotation: {fileName}");
                        }
                        catch (Exception ex)
                        {
                            _tracingService.Trace($"Warning: Failed to delete annotation {fileName}: {ex.Message}");
                        }
                    }
                }

                _tracingService.Trace($"Cleanup completed. Deleted {deletedCount} intermediate annotations.");
            }
            catch (Exception ex)
            {
                _tracingService.Trace($"Error during annotation cleanup: {ex.Message}");
                // Don't throw - cleanup failure shouldn't fail the entire process
            }
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
