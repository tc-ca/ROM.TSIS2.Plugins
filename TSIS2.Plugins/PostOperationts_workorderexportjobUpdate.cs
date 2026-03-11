using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using TSIS2.Plugins.WorkOrderExport;

namespace TSIS2.Plugins
{
    [CrmPluginRegistration(
        MessageNameEnum.Update,
        "ts_workorderexportjob",
        StageEnum.PostOperation,
        ExecutionModeEnum.Asynchronous,
        "statuscode",
        "TSIS2.Plugins.PostOperationts_workorderexportjobUpdate",
        1,
        IsolationModeEnum.Sandbox,
        Image1Type = ImageTypeEnum.PreImage,
        Image1Name = "PreImage",
        Image1Attributes = "statuscode",
        Image2Type = ImageTypeEnum.PostImage,
        Image2Name = "PostImage",
        Image2Attributes = "statuscode,ts_payloadjson,ts_name"
    )]
    public class PostOperationts_workorderexportjobUpdate : PluginBase
    {
        private const string TELEMETRY_TAG = "[WOExportTelemetry]";
        // StatusCode Values for ts_workorderexportjob
        private const int STATUS_CLIENT_PROCESSING = 741130001; // Webresource generating survey PDFs
        private const int STATUS_READY_FOR_SERVER = 741130002; // Surveys done → C# builds payload
        private const int STATUS_READY_FOR_FLOW = 741130003; // Payload ready → Flow may start
        private const int STATUS_FLOW_RUNNING = 741130004; // Flow claimed the job (lock)
        private const int STATUS_READY_FOR_MERGE = 741130005; // MAIN PDFs exist → C# merge
        private const int STATUS_COMPLETED = 741130006; // ZIP created
        private const int STATUS_ERROR = 741130007;
        private const int STATUS_MERGE_IN_PROGRESS = 741130008;
        private const int STATUS_READY_FOR_ZIP = 741130009;
        private const int STATUS_ZIP_IN_PROGRESS = 741130010;
        private const int STATUS_READY_FOR_CLEANUP = 741130011;
        private const int STATUS_CLEANUP_IN_PROGRESS = 741130012;
        private const int ERROR_MESSAGE_MAX_LENGTH = 4000;
        private const int STAGE3_SAFE_BUDGET_MS = 100000; // Keep well below sandbox 2-minute limit.
        private const int ZIP_BATCH_MIN_REMAINING_MS = 30000;
        private const int ZIP_FINALIZE_MIN_REMAINING_MS = 45000;
        private const int MAX_READY_FOR_SERVER_WORKORDERS = 100;
        private const int UILANG_ENGLISH = 1033;
        private const int UILANG_FRENCH = 1036;
        private const string EXPORT_RESX_EN = "ts_/resx/WorkOrderExport.1033.resx";
        private const string EXPORT_RESX_FR = "ts_/resx/WorkOrderExport.1036.resx";
        private static readonly HashSet<int> STAGE3_WORKER_STATUSES = new HashSet<int>
        {
            STATUS_READY_FOR_MERGE,
            STATUS_MERGE_IN_PROGRESS,
            STATUS_READY_FOR_ZIP,
            STATUS_ZIP_IN_PROGRESS,
            STATUS_READY_FOR_CLEANUP,
            STATUS_CLEANUP_IN_PROGRESS
        };
        private static readonly HashSet<int> STAGE3_REENTRY_STATUSES = new HashSet<int>
        {
            STATUS_MERGE_IN_PROGRESS,
            STATUS_ZIP_IN_PROGRESS,
            STATUS_CLEANUP_IN_PROGRESS
        };

        public PostOperationts_workorderexportjobUpdate(string unsecure, string secure)
            : base(typeof(PostOperationts_workorderexportjobUpdate))
        {
        }

        protected override void ExecuteCrmPlugin(LocalPluginContext localContext)
        {
            if (localContext == null) throw new ArgumentNullException(nameof(localContext));

            var context = localContext.PluginExecutionContext;
            var service = localContext.OrganizationService;
            var tracingService = localContext.TracingService;

            // 1. Trigger Validation
            if (context.MessageName != "Update" || context.PrimaryEntityName != "ts_workorderexportjob")
            {
                return;
            }

            Entity preImage = context.PreEntityImages.Contains("PreImage") ? context.PreEntityImages["PreImage"] : null;
            Entity postImage = context.PostEntityImages.Contains("PostImage") ? context.PostEntityImages["PostImage"] : null;

            if (preImage == null || postImage == null)
            {
                localContext.Trace("PreImage or PostImage missing. Skipping.");
                return;
            }

            int oldStatus = preImage.GetAttributeValue<OptionSetValue>("statuscode")?.Value ?? 0;
            int newStatus = postImage.GetAttributeValue<OptionSetValue>("statuscode")?.Value ?? 0;

            // Avoid trace spam from deep async re-entry loops that are already in terminal Error state.
            if (context.Depth > 1 && oldStatus == STATUS_ERROR && newStatus == STATUS_ERROR)
            {
                return;
            }

            localContext.Trace(
                $"{TELEMETRY_TAG} Trigger message={context.MessageName}, entity={context.PrimaryEntityName}, jobId={context.PrimaryEntityId}, depth={context.Depth}, mode={context.Mode}, oldStatus={oldStatus}, newStatus={newStatus}, correlationId={context.CorrelationId}, requestId={context.RequestId}");

            if (oldStatus != newStatus)
            {
                localContext.Trace(
                    $"ts_workorderexportjob statuscode transition detected. Old={oldStatus}, New={newStatus}"
                );
            }
            // Timing updates are transition-driven to avoid stale async overwrite races.
            // Same-status re-entries (e.g. Completed→Completed from concurrent workers)
            // must NOT re-run timing — they clobber stage data via read-modify-write races.
            if (oldStatus != newStatus)
            {
            }

            // Recursion guard for non-Stage3 paths only.
            // Stage 3 is a status-driven worker chain and must be able to re-enter on depth > 1.
            if (context.Depth > 1 && !IsStage3WorkerStatus(newStatus))
            {
                return;
            }

            if (newStatus == STATUS_READY_FOR_SERVER && oldStatus != STATUS_READY_FOR_SERVER)
            {
                ProcessReadyForServer(service, tracingService, context, postImage);
                return;
            }

            if (IsStage3WorkerStatus(newStatus))
            {
                if (!TryGetLiveStatusCode(service, tracingService, context.PrimaryEntityId, out int liveStatus))
                {
                    var staleStatusMessage =
                        $"Unable to verify live status before Stage 3 worker execution. oldStatus={oldStatus}, newStatus={newStatus}, depth={context.Depth}.";
                    localContext.Trace($"{TELEMETRY_TAG} Stage3LiveStatusUnavailable jobId={context.PrimaryEntityId}, details={staleStatusMessage}");
                    PersistWorkerError(service, tracingService, context.PrimaryEntityId, staleStatusMessage);
                    return;
                }

                if (liveStatus != newStatus)
                {
                    localContext.Trace(
                        $"{TELEMETRY_TAG} Stage3SkippedStale jobId={context.PrimaryEntityId}, oldStatus={oldStatus}, newStatus={newStatus}, liveStatus={liveStatus}, depth={context.Depth}");
                    return;
                }
            }

            bool shouldRunStage3Worker = IsStage3WorkerStatus(newStatus)
                && (oldStatus != newStatus
                    || IsStage3ReentryStatus(newStatus));
            if (shouldRunStage3Worker)
            {
                ProcessStage3Worker(service, tracingService, context, postImage, newStatus);
            }
        }

        private void ProcessReadyForServer(IOrganizationService service, ITracingService tracingService, IPluginExecutionContext context, Entity postImage)
        {
            Guid jobId = context.PrimaryEntityId;
            string jobName = postImage.GetAttributeValue<string>("ts_name");
            string jobContext = BuildJobContext(jobId, jobName);
            tracingService.Trace($"Starting Ready For Server Processing. {jobContext}");

            try
            {
                // 1. Retrieve ts_surveypayloadjson
                var jobEntity = service.Retrieve("ts_workorderexportjob", jobId, new ColumnSet("ts_surveypayloadjson"));
                string surveyPayload = jobEntity.GetAttributeValue<string>("ts_surveypayloadjson");

                // 2. Parse IDs
                List<Guid> workOrderIds = ParsePayload(surveyPayload, tracingService);
                if (workOrderIds == null || !workOrderIds.Any())
                {
                    throw new InvalidPluginExecutionException("No Work Order IDs found in ts_surveypayloadjson.");
                }

                // 3. Batch Guard
                if (workOrderIds.Count > MAX_READY_FOR_SERVER_WORKORDERS)
                {
                    throw new InvalidPluginExecutionException(
                        $"Too many Work Orders ({workOrderIds.Count}). Max batch size is {MAX_READY_FOR_SERVER_WORKORDERS}.");
                }

                // 4. Retrieve and Map Data
                var retriever = new WorkOrderDataRetriever(service, tracingService);
                var mapper = new WorkOrderExportMapper(tracingService);
                
                var payload = new WorkOrderExportPayload();
                
                foreach (var woId in workOrderIds)
                {
                    tracingService.Trace($"Processing WO: {woId}");
                    var data = retriever.RetrieveWorkOrderData(woId);
                    var model = mapper.Map(data);
                    payload.WorkOrders.Add(model);
                }

                // 5. Serialize
                string jsonPayload = JsonConvert.SerializeObject(payload);
                tracingService.Trace($"Serialized payload length: {jsonPayload.Length}");

                // 6. Update Job
                Entity updateJob = new Entity("ts_workorderexportjob", jobId);
                updateJob["ts_payloadjson"] = jsonPayload;
                updateJob["statuscode"] = new OptionSetValue(STATUS_READY_FOR_FLOW);
                updateJob["ts_errormessage"] = string.Empty;
                service.Update(updateJob);

                tracingService.Trace("Completed Ready For Server processing. ts_payloadjson set. Status set to Ready For Flow.");
            }
            catch (Exception ex)
            {
                tracingService.Trace($"ERROR in Ready For Server. {jobContext}. Message={ex.Message}");
                Entity errorJob = new Entity("ts_workorderexportjob", jobId);
                errorJob["statuscode"] = new OptionSetValue(STATUS_ERROR);
                errorJob["ts_lastheartbeat"] = DateTime.UtcNow;
                string serverFailedMsg = GetExportMessage(service, tracingService, context.InitiatingUserId, "ServerProcessingFailed", ex.Message);
                errorJob["ts_errormessage"] = TruncateForErrorField(
                    tracingService,
                    serverFailedMsg + "\nStack: " + (ex.StackTrace ?? ""));
                service.Update(errorJob);
            }
        }

        private void ProcessStage3Worker(IOrganizationService service, ITracingService tracingService, IPluginExecutionContext context, Entity postImage, int currentStatus)
        {
            Guid jobId = context.PrimaryEntityId;
            string jobName = postImage.GetAttributeValue<string>("ts_name");
            string jobContext = BuildJobContext(jobId, jobName);
            tracingService.Trace($"Starting Stage 3 worker. Status={currentStatus}. {jobContext}");
            var stage3Stopwatch = Stopwatch.StartNew();

            try
            {
                var jobEntity = service.Retrieve(
                    "ts_workorderexportjob",
                    jobId,
                    new ColumnSet("ts_surveypayloadjson", "ts_payloadjson", "ts_nextmergeindex", "ts_totalunits"));

                string surveyPayload = jobEntity.GetAttributeValue<string>("ts_surveypayloadjson");
                string payloadJson = jobEntity.GetAttributeValue<string>("ts_payloadjson");
                List<Guid> workOrderIds = ParsePayload(surveyPayload, tracingService);
                Dictionary<Guid, string> workOrderNamesById = ParseWorkOrderNames(payloadJson, tracingService);
                if (workOrderIds == null || !workOrderIds.Any())
                {
                    throw new InvalidPluginExecutionException("No Work Order IDs found in ts_surveypayloadjson.");
                }

                var exportService = new WorkOrderExportService(service, tracingService);
                int totalUnits = jobEntity.GetAttributeValue<int?>("ts_totalunits") ?? 0;
                bool hasTotalUnits = totalUnits > 0;
                int workOrderCount = Math.Max(0, workOrderIds.Count);
                int stage3DerivedStartDoneUnits = hasTotalUnits
                    ? Math.Max(0, totalUnits - ((2 * workOrderCount) + 1))
                    : 0;
                int stage3DerivedMergeDoneUnits = hasTotalUnits
                    ? Math.Max(0, totalUnits - (workOrderCount + 1))
                    : 0;
                tracingService.Trace(
                    $"{TELEMETRY_TAG} Stage3Start jobId={jobId}, status={currentStatus}, depth={context.Depth}, workOrderCount={workOrderCount}, totalUnits={totalUnits}, hasTotalUnits={hasTotalUnits}, derivedStartDoneUnits={stage3DerivedStartDoneUnits}, derivedMergeDoneUnits={stage3DerivedMergeDoneUnits}");

                if (currentStatus == STATUS_READY_FOR_MERGE || currentStatus == STATUS_MERGE_IN_PROGRESS)
                {
                    exportService.UpdateHeartbeat(jobId, GetExportMessage(service, tracingService, context.InitiatingUserId, "ValidatingArtifactsFor", workOrderIds.Count));

                    var mergeStopwatch = Stopwatch.StartNew();
                    // Validation-only: confirm all MAIN + SURVEY PDF notes exist using metadata queries.
                    // No PDF content is loaded here — the actual merge happens in the ZIP phase.
                    var validationResult = exportService.ValidateAllArtifactsExist(jobId, workOrderIds);
                    mergeStopwatch.Stop();

                    long estimatedMB = validationResult.TotalInputBytes / 1024 / 1024;
                    tracingService.Trace(
                        $"{TELEMETRY_TAG} MergeValidation jobId={jobId}, workOrders={validationResult.WorkOrdersValidated}, totalSurveyPdfs={validationResult.TotalSurveyPdfs}, estimatedInputBytes={validationResult.TotalInputBytes}, elapsedMs={mergeStopwatch.ElapsedMilliseconds}");

                    var update = new Entity("ts_workorderexportjob", jobId);
                    update["ts_nextmergeindex"] = 0; // ZIP phase starts from index 0
                    if (hasTotalUnits)
                    {
                        // Stage 3 progress is derived from total units + stage-local work-order count.
                        // It is intentionally decoupled from any Stage 2 done-units baseline value.
                        update["ts_doneunits"] = stage3DerivedMergeDoneUnits;
                    }
                    update["ts_lastheartbeat"] = DateTime.UtcNow;
                    update["ts_progressmessage"] = GetExportMessage(service, tracingService, context.InitiatingUserId, "ValidatedWorkOrders", validationResult.WorkOrdersValidated, validationResult.TotalSurveyPdfs, estimatedMB);
                    update["statuscode"] = new OptionSetValue(STATUS_READY_FOR_ZIP);
                    update["ts_errormessage"] = string.Empty;
                    service.Update(update);
                    stage3Stopwatch.Stop();
                    tracingService.Trace(
                        $"{TELEMETRY_TAG} Stage3End jobId={jobId}, status={currentStatus}, branch=Merge, resultStatus={STATUS_READY_FOR_ZIP}, elapsedMs={stage3Stopwatch.ElapsedMilliseconds}");
                    return;
                }

                if (currentStatus == STATUS_READY_FOR_ZIP || currentStatus == STATUS_ZIP_IN_PROGRESS)
                {
                    int nextZipIndex = Math.Max(0, Math.Min(workOrderIds.Count, jobEntity.GetAttributeValue<int?>("ts_nextmergeindex") ?? 0));

                    // Fresh ZIP run: clear any stale temp ZIP artifact first.
                    if (nextZipIndex == 0)
                    {
                        long estimatedTotalZipBytes = exportService.EstimateTotalZipInputBytes(jobId, workOrderIds);
                        int zipSizeLimitBytes = exportService.GetZipSizeLimitBytes();
                        tracingService.Trace(
                            $"{TELEMETRY_TAG} ZipEstimate jobId={jobId}, estimatedTotalBytes={estimatedTotalZipBytes}, limitBytes={zipSizeLimitBytes}, workOrderCount={workOrderIds.Count}");
                        // Temporary calibration mode support: estimate is informative only.
                        // Runtime hard limits in ProcessZipBatch/FinalizePersistedZip remain enforced.

                        exportService.ClearTemporaryZipStorage(jobId);
                    }

                    if (IsStage3BudgetExceeded(stage3Stopwatch) || RemainingStage3BudgetMs(stage3Stopwatch) < ZIP_BATCH_MIN_REMAINING_MS)
                    {
                        PersistZipYield(service, tracingService, context.InitiatingUserId, jobId, nextZipIndex, workOrderIds.Count, stage3Stopwatch, "Yielding before ZIP batch to avoid sandbox timeout");
                        return;
                    }

                    int nextIndex = exportService.ProcessZipBatchMulti(
                        jobId,
                        workOrderIds,
                        nextZipIndex,
                        workOrderNamesById,
                        stage3Stopwatch,
                        ZIP_BATCH_MIN_REMAINING_MS);

                    bool zipDone = nextIndex >= workOrderIds.Count;
                    tracingService.Trace(
                        $"{TELEMETRY_TAG} ZipBatch jobId={jobId}, startIndex={nextZipIndex}, endIndexExclusive={nextIndex}, zipDone={zipDone}");

                    if (!zipDone)
                    {
                        var continueZip = new Entity("ts_workorderexportjob", jobId);
                        continueZip["statuscode"] = new OptionSetValue(STATUS_ZIP_IN_PROGRESS);
                        continueZip["ts_nextmergeindex"] = nextIndex;
                        continueZip["ts_lastheartbeat"] = DateTime.UtcNow;
                        continueZip["ts_progressmessage"] = GetExportMessage(service, tracingService, context.InitiatingUserId, "BuildingFinalZIP", nextIndex, workOrderIds.Count);
                        continueZip["ts_errormessage"] = string.Empty;
                        service.Update(continueZip);
                        stage3Stopwatch.Stop();
                        tracingService.Trace(
                            $"{TELEMETRY_TAG} Stage3End jobId={jobId}, status={currentStatus}, branch=Zip, resultStatus={STATUS_ZIP_IN_PROGRESS}, zipPresent=false, elapsedMs={stage3Stopwatch.ElapsedMilliseconds}");
                        return;
                    }

                    if (RemainingStage3BudgetMs(stage3Stopwatch) < ZIP_FINALIZE_MIN_REMAINING_MS)
                    {
                        PersistZipYield(service, tracingService, context.InitiatingUserId, jobId, workOrderIds.Count, workOrderIds.Count, stage3Stopwatch, "Yielding before ZIP finalize to avoid sandbox timeout");
                        return;
                    }

                    exportService.FinalizePersistedZip(jobId);

                    // Strict gate: do not leave ZIP stage until final ZIP presence is confirmed.
                    if (!exportService.IsFinalZipPresent(jobId))
                    {
                        var waitForZip = new Entity("ts_workorderexportjob", jobId);
                        waitForZip["statuscode"] = new OptionSetValue(STATUS_ZIP_IN_PROGRESS);
                        waitForZip["ts_nextmergeindex"] = workOrderIds.Count;
                        waitForZip["ts_lastheartbeat"] = DateTime.UtcNow;
                        waitForZip["ts_progressmessage"] = GetExportMessage(service, tracingService, context.InitiatingUserId, "WaitingForZipConfirmation");
                        waitForZip["ts_errormessage"] = string.Empty;
                        service.Update(waitForZip);
                        stage3Stopwatch.Stop();
                        tracingService.Trace(
                            $"{TELEMETRY_TAG} Stage3End jobId={jobId}, status={currentStatus}, branch=Zip, resultStatus={STATUS_ZIP_IN_PROGRESS}, zipPresent=false, elapsedMs={stage3Stopwatch.ElapsedMilliseconds}");
                        return;
                    }

                    var readyForCleanup = new Entity("ts_workorderexportjob", jobId);
                    readyForCleanup["statuscode"] = new OptionSetValue(STATUS_READY_FOR_CLEANUP);
                    readyForCleanup["ts_nextmergeindex"] = workOrderIds.Count;
                    if (hasTotalUnits)
                    {
                        // Keep progress below 100% until cleanup truly completes.
                        readyForCleanup["ts_doneunits"] = stage3DerivedMergeDoneUnits;
                    }
                    readyForCleanup["ts_lastheartbeat"] = DateTime.UtcNow;
                    readyForCleanup["ts_progressmessage"] = GetExportMessage(service, tracingService, context.InitiatingUserId, "ZIPReadyStartingCleanup");
                    readyForCleanup["ts_errormessage"] = string.Empty;
                    service.Update(readyForCleanup);
                    stage3Stopwatch.Stop();
                    tracingService.Trace(
                        $"{TELEMETRY_TAG} Stage3End jobId={jobId}, status={currentStatus}, branch=Zip, resultStatus={STATUS_READY_FOR_CLEANUP}, zipPresent=true, elapsedMs={stage3Stopwatch.ElapsedMilliseconds}");
                    return;
                }

                if (currentStatus == STATUS_READY_FOR_CLEANUP || currentStatus == STATUS_CLEANUP_IN_PROGRESS)
                {
                    var inProgress = new Entity("ts_workorderexportjob", jobId);
                    inProgress["statuscode"] = new OptionSetValue(STATUS_CLEANUP_IN_PROGRESS);
                    inProgress["ts_lastheartbeat"] = DateTime.UtcNow;
                    inProgress["ts_progressmessage"] = GetExportMessage(service, tracingService, context.InitiatingUserId, "DeletingTemporaryPDFArtifacts");
                    service.Update(inProgress);

                    exportService.CleanupIntermediateArtifacts(jobId);

                    // Final safety gate: don't mark completed unless ZIP is still confirmed present.
                    if (!exportService.IsFinalZipPresent(jobId))
                    {
                        var backToZip = new Entity("ts_workorderexportjob", jobId);
                        backToZip["statuscode"] = new OptionSetValue(STATUS_READY_FOR_ZIP);
                        backToZip["ts_lastheartbeat"] = DateTime.UtcNow;
                        backToZip["ts_progressmessage"] = GetExportMessage(service, tracingService, context.InitiatingUserId, "ZIPVerificationFailedRetrying");
                        backToZip["ts_errormessage"] = string.Empty;
                        service.Update(backToZip);
                        stage3Stopwatch.Stop();
                        tracingService.Trace(
                            $"{TELEMETRY_TAG} Stage3End jobId={jobId}, status={currentStatus}, branch=Cleanup, resultStatus={STATUS_READY_FOR_ZIP}, zipPresent=false, elapsedMs={stage3Stopwatch.ElapsedMilliseconds}");
                        return;
                    }

                    var completed = new Entity("ts_workorderexportjob", jobId);
                    completed["statuscode"] = new OptionSetValue(STATUS_COMPLETED);
                    if (hasTotalUnits)
                    {
                        completed["ts_doneunits"] = totalUnits;
                    }
                    completed["ts_lastheartbeat"] = DateTime.UtcNow;
                    completed["ts_progressmessage"] = GetExportMessage(service, tracingService, context.InitiatingUserId, "CleanupCompleted");
                    completed["ts_errormessage"] = string.Empty;
                    service.Update(completed);
                    stage3Stopwatch.Stop();
                    tracingService.Trace(
                        $"{TELEMETRY_TAG} Stage3End jobId={jobId}, status={currentStatus}, branch=Cleanup, resultStatus={STATUS_COMPLETED}, zipPresent=true, elapsedMs={stage3Stopwatch.ElapsedMilliseconds}");
                    return;
                }

                stage3Stopwatch.Stop();
                tracingService.Trace(
                    $"{TELEMETRY_TAG} Stage3End jobId={jobId}, status={currentStatus}, branch=NoOp, elapsedMs={stage3Stopwatch.ElapsedMilliseconds}");
            }
            catch (Exception ex)
            {
                stage3Stopwatch.Stop();
                tracingService.Trace(
                    $"{TELEMETRY_TAG} Stage3Error jobId={jobId}, status={currentStatus}, elapsedMs={stage3Stopwatch.ElapsedMilliseconds}, error={ex.Message}");
                tracingService.Trace($"ERROR in Stage 3 worker. {jobContext}. Message={ex.Message}. Stack={ex.StackTrace}");
                string stage3ErrorMsg = GetExportMessage(service, tracingService, context.InitiatingUserId, "Stage3WorkerFailed", ex?.Message ?? "Unknown error.");
                PersistWorkerError(service, tracingService, jobId, stage3ErrorMsg);
                return;
            }
        }

        private bool IsStage3WorkerStatus(int status)
        {
            return STAGE3_WORKER_STATUSES.Contains(status);
        }

        private bool IsStage3ReentryStatus(int status)
        {
            return STAGE3_REENTRY_STATUSES.Contains(status);
        }

        private static bool IsStage3BudgetExceeded(Stopwatch stage3Stopwatch)
        {
            return stage3Stopwatch != null && stage3Stopwatch.ElapsedMilliseconds >= STAGE3_SAFE_BUDGET_MS;
        }

        private static long RemainingStage3BudgetMs(Stopwatch stage3Stopwatch)
        {
            if (stage3Stopwatch == null)
            {
                return STAGE3_SAFE_BUDGET_MS;
            }

            return Math.Max(0, STAGE3_SAFE_BUDGET_MS - stage3Stopwatch.ElapsedMilliseconds);
        }

        private void PersistZipYield(
            IOrganizationService service,
            ITracingService tracingService,
            Guid userId,
            Guid jobId,
            int nextIndex,
            int totalCount,
            Stopwatch stage3Stopwatch,
            string reason)
        {
            int displayIndex = Math.Max(0, Math.Min(totalCount, nextIndex));
            var continueZip = new Entity("ts_workorderexportjob", jobId);
            continueZip["statuscode"] = new OptionSetValue(STATUS_ZIP_IN_PROGRESS);
            continueZip["ts_nextmergeindex"] = Math.Max(0, nextIndex);
            continueZip["ts_lastheartbeat"] = DateTime.UtcNow;
            continueZip["ts_progressmessage"] = GetExportMessage(service, tracingService, userId, "BuildingFinalZIP", displayIndex, totalCount);
            continueZip["ts_errormessage"] = string.Empty;
            service.Update(continueZip);

            long elapsedMs = stage3Stopwatch?.ElapsedMilliseconds ?? 0;
            tracingService.Trace(
                $"{TELEMETRY_TAG} Stage3Yield jobId={jobId}, branch=Zip, nextIndex={nextIndex}, total={totalCount}, elapsedMs={elapsedMs}, reason={reason}");
        }

        private void PersistWorkerError(IOrganizationService service, ITracingService tracingService, Guid jobId, string message)
        {
            try
            {
                Entity errorJob = new Entity("ts_workorderexportjob", jobId);
                errorJob["statuscode"] = new OptionSetValue(STATUS_ERROR);
                errorJob["ts_lastheartbeat"] = DateTime.UtcNow;
                errorJob["ts_errormessage"] = TruncateForErrorField(tracingService, message);
                service.Update(errorJob);
            }
            catch (Exception persistEx)
            {
                tracingService.Trace(
                    $"{TELEMETRY_TAG} ErrorPersistFailed jobId={jobId}, message={persistEx.Message}");
                throw new InvalidPluginExecutionException(
                    $"Failed to persist Stage 3 worker error for job {jobId}. {persistEx.Message}",
                    persistEx);
            }
        }

        private static string BuildWorkerUserError(Exception ex)
        {
            string message = ex?.Message ?? "Unknown error.";
            if (ex is InvalidPluginExecutionException)
            {
                return message.Trim();
            }

            return $"Stage 3 worker failed: {message.Trim()}";
        }

        private static string BuildJobContext(Guid jobId, string jobName)
        {
            return string.IsNullOrWhiteSpace(jobName)
                ? $"jobId={jobId}, jobName=<not available>"
                : $"jobId={jobId}, jobName='{jobName}'";
        }

        /// <summary>
        /// Get a localized message from the Work Order Export resx. Uses initiating user's UI language.
        /// </summary>
        private static string GetExportMessage(IOrganizationService service, ITracingService tracingService, Guid userId, string resourceId, params object[] args)
        {
            try
            {
                int lang = LocalizationHelper.RetrieveUserUILanguageCode(service, userId);
                string resourceName = lang == UILANG_FRENCH ? EXPORT_RESX_FR : EXPORT_RESX_EN;
                string template = LocalizationHelper.GetMessage(tracingService, service, resourceName, resourceId);
                if (args != null && args.Length > 0)
                {
                    return string.Format(CultureInfo.CurrentCulture, template, args);
                }
                return template ?? resourceId;
            }
            catch (Exception ex)
            {
                tracingService.Trace("GetExportMessage failed for key={0}, fallback to key. Error={1}", resourceId, ex.Message);
                return args != null && args.Length > 0 ? string.Format(CultureInfo.CurrentCulture, resourceId, args) : resourceId;
            }
        }

        private static bool TryGetLiveStatusCode(IOrganizationService service, ITracingService tracingService, Guid jobId, out int statusCode)
        {
            statusCode = 0;
            try
            {
                var row = service.Retrieve("ts_workorderexportjob", jobId, new ColumnSet("statuscode"));
                statusCode = row.GetAttributeValue<OptionSetValue>("statuscode")?.Value ?? 0;
                return true;
            }
            catch (Exception ex)
            {
                tracingService.Trace($"{TELEMETRY_TAG} LiveStatusReadFailed jobId={jobId}, message={ex.Message}");
                return false;
            }
        }
        private static string TruncateForErrorField(ITracingService tracingService, string message)
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

            tracingService.Trace($"ts_errormessage exceeded max length {maxLength}. Message was truncated.");
            return truncated;
        }

        private List<Guid> ParsePayload(string json, ITracingService trace)
        {
            if (string.IsNullOrWhiteSpace(json)) return new List<Guid>();

            try
            {
                JObject obj = JObject.Parse(json);
                List<Guid> ids = new List<Guid>();

                // Primary schema: { ids: ["guid", ...] } (Client JS)
                if (obj["ids"] != null)
                {
                    foreach (var token in obj["ids"])
                    {
                        if (TryParseLooseGuid(token?.ToString(), out Guid id))
                        {
                            ids.Add(id);
                        }
                    }
                }
                // Fallback: { WorkOrders: [...] } (Flow/Server)
                else if (obj["WorkOrders"] != null)
                {
                    var woToken = obj["WorkOrders"];
                    if (woToken.Type == JTokenType.Array)
                    {
                        foreach (var token in woToken)
                        {
                            if (token.Type == JTokenType.String)
                            {
                                if (TryParseLooseGuid(token.ToString(), out Guid id))
                                {
                                    ids.Add(id);
                                }
                            }
                            else if (token.Type == JTokenType.Object)
                            {
                                var woObj = (JObject)token;
                                string idStr = woObj["msdyn_workorderid"]?.ToString() 
                                            ?? woObj["WorkOrderId"]?.ToString()
                                            ?? woObj["id"]?.ToString();

                                if (TryParseLooseGuid(idStr, out Guid id))
                                {
                                    ids.Add(id);
                                }
                            }
                        }
                    }
                }

                trace.Trace($"Parsed {ids.Count} IDs from payload.");
                return ids;
            }
            catch (Exception ex)
            {
                string preview = json.Length > 240 ? json.Substring(0, 240) + "...[truncated]" : json;
                trace.Trace($"Error parsing payload: {ex.Message}. PayloadPreview={preview}");
                throw new InvalidPluginExecutionException($"Invalid ts_surveypayloadjson. {ex.Message}", ex);
            }
        }

        private Dictionary<Guid, string> ParseWorkOrderNames(string payloadJson, ITracingService trace)
        {
            var namesById = new Dictionary<Guid, string>();
            if (string.IsNullOrWhiteSpace(payloadJson))
            {
                return namesById;
            }

            try
            {
                var payload = JsonConvert.DeserializeObject<WorkOrderExportPayload>(payloadJson);
                if (payload?.WorkOrders == null)
                {
                    return namesById;
                }

                foreach (var item in payload.WorkOrders)
                {
                    if (TryParseLooseGuid(item?.WorkOrderId, out Guid workOrderId) && !namesById.ContainsKey(workOrderId))
                    {
                        namesById[workOrderId] = item.WorkOrderNumber;
                    }
                }
            }
            catch (Exception ex)
            {
                trace.Trace($"Warning: Unable to parse Work Order names from ts_payloadjson. Message={ex.Message}");
            }

            return namesById;
        }

        private static bool TryParseLooseGuid(string rawValue, out Guid id)
        {
            id = Guid.Empty;
            if (string.IsNullOrWhiteSpace(rawValue))
            {
                return false;
            }

            string normalized = rawValue.Trim();
            if (normalized.StartsWith("{", StringComparison.Ordinal) && normalized.EndsWith("}", StringComparison.Ordinal))
            {
                normalized = normalized.Substring(1, normalized.Length - 2).Trim();
            }

            return Guid.TryParse(normalized, out id);
        }
    }
}

