using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
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
        // StatusCode Values for ts_workorderexportjob
        private const int STATUS_ACTIVE = 1;
        private const int STATUS_CLIENT_PROCESSING = 741130001; // Webresource generating survey PDFs
        private const int STATUS_READY_FOR_SERVER = 741130002; // Surveys done → C# builds payload
        private const int STATUS_READY_FOR_FLOW = 741130003; // Payload ready → Flow may start
        private const int STATUS_FLOW_RUNNING = 741130004; // Flow claimed the job (lock)
        private const int STATUS_READY_FOR_MERGE = 741130005; // MAIN PDFs exist → C# merge
        private const int STATUS_COMPLETED = 741130006; // ZIP created
        private const int STATUS_ERROR = 741130007; //Error occurred
        private const int STATUS_MERGE_IN_PROGRESS = 741130008;
        private const int STATUS_READY_FOR_ZIP = 741130009;
        private const int STATUS_ZIP_IN_PROGRESS = 741130010;
        private const int STATUS_READY_FOR_CLEANUP = 741130011;
        private const int STATUS_CLEANUP_IN_PROGRESS = 741130012;
        private const int ERROR_MESSAGE_MAX_LENGTH = 4000;
        private const int MERGE_BATCH_SIZE = 5;

        // Storage Configuration
        /// <summary>
        /// Controls where the final ZIP is stored:
        /// - false (default): ZIP saved as Dataverse Note/Annotation attached to the job
        /// - true: ZIP saved directly to the ts_finalexportzip File Column on the job record
        /// 
        /// When true (File Storage mode):
        /// - Final ZIP is uploaded to the ts_finalexportzip file column using Dataverse File APIs
        /// - ALL intermediate annotations (Survey PDFs, Main PDFs, Merged PDFs) are deleted after successful upload
        /// - More efficient for storage and reduces database size
        /// - Supports larger files (up to Dataverse File Column limits ~128MB default)
        /// 
        /// When false (Annotation mode):
        /// - Final ZIP is saved as a Note/Annotation
        /// - Intermediate PDFs are still cleaned up, but final ZIP Note is preserved
        /// - Compatible with older Dataverse instances that might not support File Columns
        /// </summary>
        private const bool USE_FILE_STORAGE = true; // true = ts_finalexportzip file column, false = annotation

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

            if (oldStatus != newStatus)
            {
                localContext.Trace(
                    $"ts_workorderexportjob statuscode transition detected. Old={oldStatus}, New={newStatus}"
                );
            }

            // Recursion guard for non-Stage3 paths only.
            // Stage 3 is a status-driven worker chain and must be able to re-enter on depth > 1.
            if (context.Depth > 1 && !IsStage3WorkerStatus(newStatus))
            {
                localContext.Trace($"Exiting due to recursion guard for non-Stage3 status. Depth={context.Depth}, Status={newStatus}");
                return;
            }

            // Trigger on transition TO Ready For Server (741130002)
            if (newStatus == STATUS_READY_FOR_SERVER && oldStatus != STATUS_READY_FOR_SERVER)
            {
                ProcessReadyForServer(service, tracingService, context, postImage);
                return;
            }

            if (IsStage3WorkerStatus(newStatus) && oldStatus != newStatus)
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
                if (workOrderIds.Count > 50)
                {
                    throw new InvalidPluginExecutionException($"Too many Work Orders ({workOrderIds.Count}). Max batch size is 50.");
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
                errorJob["ts_errormessage"] = TruncateForErrorField(
                    tracingService,
                    $"Server Processing Failed: {ex.Message}\nStack: {ex.StackTrace}");
                service.Update(errorJob);
            }
        }

        private void ProcessStage3Worker(IOrganizationService service, ITracingService tracingService, IPluginExecutionContext context, Entity postImage, int currentStatus)
        {
            Guid jobId = context.PrimaryEntityId;
            string jobName = postImage.GetAttributeValue<string>("ts_name");
            string jobContext = BuildJobContext(jobId, jobName);
            tracingService.Trace($"Starting Stage 3 worker. Status={currentStatus}. {jobContext}");

            try
            {
                var jobEntity = service.Retrieve(
                    "ts_workorderexportjob",
                    jobId,
                    new ColumnSet("ts_surveypayloadjson", "ts_nextmergeindex", "ts_doneunits", "ts_totalunits"));

                string surveyPayload = jobEntity.GetAttributeValue<string>("ts_surveypayloadjson");
                List<Guid> workOrderIds = ParsePayload(surveyPayload, tracingService);
                if (workOrderIds == null || !workOrderIds.Any())
                {
                    throw new InvalidPluginExecutionException("No Work Order IDs found in ts_surveypayloadjson.");
                }

                var exportService = new WorkOrderExportService(service, tracingService);
                int currentDoneUnits = jobEntity.GetAttributeValue<int?>("ts_doneunits") ?? 0;
                int totalUnits = jobEntity.GetAttributeValue<int?>("ts_totalunits") ?? 0;
                int doneUnitsUpperBound = totalUnits > 0 ? totalUnits : int.MaxValue;

                if (currentStatus == STATUS_READY_FOR_MERGE || currentStatus == STATUS_MERGE_IN_PROGRESS)
                {
                    int nextMergeIndex = jobEntity.GetAttributeValue<int?>("ts_nextmergeindex") ?? 0;
                    exportService.UpdateHeartbeat(jobId, $"Merge worker: processing {nextMergeIndex + 1}/{workOrderIds.Count}.");

                    int nextIndex = exportService.ProcessMergeBatch(jobId, workOrderIds, nextMergeIndex, MERGE_BATCH_SIZE);
                    bool mergeDone = nextIndex >= workOrderIds.Count;
                    int mergeUnitsProcessed = Math.Max(0, nextIndex - nextMergeIndex);
                    int updatedDoneUnits = Math.Min(doneUnitsUpperBound, currentDoneUnits + mergeUnitsProcessed);

                    var update = new Entity("ts_workorderexportjob", jobId);
                    update["ts_nextmergeindex"] = nextIndex;
                    update["ts_doneunits"] = updatedDoneUnits;
                    update["ts_lastheartbeat"] = DateTime.UtcNow;
                    update["ts_progressmessage"] = mergeDone
                        ? "Merge worker completed all work orders. Ready for ZIP."
                        : $"Merge worker progressed to {nextIndex}/{workOrderIds.Count}.";
                    update["statuscode"] = new OptionSetValue(mergeDone ? STATUS_READY_FOR_ZIP : STATUS_MERGE_IN_PROGRESS);
                    update["ts_errormessage"] = string.Empty;
                    service.Update(update);
                    return;
                }

                if (currentStatus == STATUS_READY_FOR_ZIP || currentStatus == STATUS_ZIP_IN_PROGRESS)
                {
                    var inProgress = new Entity("ts_workorderexportjob", jobId);
                    inProgress["statuscode"] = new OptionSetValue(STATUS_ZIP_IN_PROGRESS);
                    inProgress["ts_lastheartbeat"] = DateTime.UtcNow;
                    inProgress["ts_progressmessage"] = "ZIP worker: preparing final ZIP.";
                    service.Update(inProgress);

                    // Always recreate final ZIP for this run so restart semantics are overwrite/replace.
                    exportService.CreateAndPersistFinalZip(jobId, workOrderIds, USE_FILE_STORAGE);

                    // Strict gate: do not leave ZIP stage until final ZIP presence is confirmed.
                    if (!exportService.IsFinalZipPresent(jobId, USE_FILE_STORAGE))
                    {
                        var waitForZip = new Entity("ts_workorderexportjob", jobId);
                        waitForZip["statuscode"] = new OptionSetValue(STATUS_ZIP_IN_PROGRESS);
                        waitForZip["ts_lastheartbeat"] = DateTime.UtcNow;
                        waitForZip["ts_progressmessage"] = "ZIP worker: waiting for ZIP confirmation.";
                        waitForZip["ts_errormessage"] = string.Empty;
                        service.Update(waitForZip);
                        return;
                    }

                    var readyForCleanup = new Entity("ts_workorderexportjob", jobId);
                    readyForCleanup["statuscode"] = new OptionSetValue(STATUS_READY_FOR_CLEANUP);
                    readyForCleanup["ts_doneunits"] = Math.Min(doneUnitsUpperBound, currentDoneUnits + 1);
                    readyForCleanup["ts_lastheartbeat"] = DateTime.UtcNow;
                    readyForCleanup["ts_progressmessage"] = "ZIP worker completed. Ready for cleanup.";
                    readyForCleanup["ts_errormessage"] = string.Empty;
                    service.Update(readyForCleanup);
                    return;
                }

                if (currentStatus == STATUS_READY_FOR_CLEANUP || currentStatus == STATUS_CLEANUP_IN_PROGRESS)
                {
                    var inProgress = new Entity("ts_workorderexportjob", jobId);
                    inProgress["statuscode"] = new OptionSetValue(STATUS_CLEANUP_IN_PROGRESS);
                    inProgress["ts_lastheartbeat"] = DateTime.UtcNow;
                    inProgress["ts_progressmessage"] = "Cleanup worker: removing intermediate files.";
                    service.Update(inProgress);

                    exportService.CleanupIntermediateArtifacts(jobId, USE_FILE_STORAGE);

                    // Final safety gate: don't mark completed unless ZIP is still confirmed present.
                    if (!exportService.IsFinalZipPresent(jobId, USE_FILE_STORAGE))
                    {
                        var backToZip = new Entity("ts_workorderexportjob", jobId);
                        backToZip["statuscode"] = new OptionSetValue(STATUS_READY_FOR_ZIP);
                        backToZip["ts_lastheartbeat"] = DateTime.UtcNow;
                        backToZip["ts_progressmessage"] = "ZIP not confirmed after cleanup. Retrying ZIP stage.";
                        backToZip["ts_errormessage"] = string.Empty;
                        service.Update(backToZip);
                        return;
                    }

                    var completed = new Entity("ts_workorderexportjob", jobId);
                    completed["statuscode"] = new OptionSetValue(STATUS_COMPLETED);
                    completed["ts_lastheartbeat"] = DateTime.UtcNow;
                    completed["ts_progressmessage"] = "Cleanup worker completed.";
                    completed["ts_errormessage"] = string.Empty;
                    service.Update(completed);
                    return;
                }
            }
            catch (Exception ex)
            {
                tracingService.Trace($"ERROR in Stage 3 worker. {jobContext}. Message={ex.Message}");
                PersistWorkerError(service, tracingService, jobId, $"Stage 3 worker failed: {ex.Message}\nStack: {ex.StackTrace}");
                return;
            }
        }

        private bool IsStage3WorkerStatus(int status)
        {
            return status == STATUS_READY_FOR_MERGE
                || status == STATUS_MERGE_IN_PROGRESS
                || status == STATUS_READY_FOR_ZIP
                || status == STATUS_ZIP_IN_PROGRESS
                || status == STATUS_READY_FOR_CLEANUP
                || status == STATUS_CLEANUP_IN_PROGRESS;
        }

        private void PersistWorkerError(IOrganizationService service, ITracingService tracingService, Guid jobId, string message)
        {
            Entity errorJob = new Entity("ts_workorderexportjob", jobId);
            errorJob["statuscode"] = new OptionSetValue(STATUS_ERROR);
            errorJob["ts_lastheartbeat"] = DateTime.UtcNow;
            errorJob["ts_errormessage"] = TruncateForErrorField(tracingService, message);
            service.Update(errorJob);
        }

        private static string BuildJobContext(Guid jobId, string jobName)
        {
            return string.IsNullOrWhiteSpace(jobName)
                ? $"jobId={jobId}, jobName=<not available>"
                : $"jobId={jobId}, jobName='{jobName}'";
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
                        string idStr = token?.ToString()?.Trim();
                        if (string.IsNullOrEmpty(idStr)) continue;
                        
                        if (idStr.StartsWith("{") && idStr.EndsWith("}"))
                        {
                            idStr = idStr.Substring(1, idStr.Length - 2).Trim();
                        }
                        
                        if (Guid.TryParse(idStr, out Guid id))
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
                                string idStr = token.ToString().Trim();
                                if (idStr.StartsWith("{") && idStr.EndsWith("}"))
                                {
                                    idStr = idStr.Substring(1, idStr.Length - 2).Trim();
                                }
                                if (Guid.TryParse(idStr, out Guid id))
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
                                
                                if (!string.IsNullOrEmpty(idStr))
                                {
                                    idStr = idStr.Trim();
                                    if (idStr.StartsWith("{") && idStr.EndsWith("}"))
                                    {
                                        idStr = idStr.Substring(1, idStr.Length - 2).Trim();
                                    }
                                    if (Guid.TryParse(idStr, out Guid id))
                                    {
                                        ids.Add(id);
                                    }
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
                trace.Trace($"Error parsing payload: {ex.Message}");
                return new List<Guid>();
            }
        }
    }
}
