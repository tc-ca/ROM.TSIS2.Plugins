using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Workflow.Runtime.Tracking;
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

            // HARD recursion guard
            if (context.Depth > 1)
            {
                localContext.Trace($"Exiting due to recursion guard. Depth={context.Depth}");
                return;
            }

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

            // Trigger on transition TO Ready For Server (741130002)
            if (newStatus == STATUS_READY_FOR_SERVER && oldStatus != STATUS_READY_FOR_SERVER)
            {
                ProcessReadyForServer(service, tracingService, context, postImage);
                return;
            }

            // Trigger on transition TO Ready For Merge (741130005)
            if (newStatus == STATUS_READY_FOR_MERGE && oldStatus != STATUS_READY_FOR_MERGE)
            {
                ProcessReadyForMerge(service, tracingService, context, postImage);
            }
        }

        private void ProcessReadyForServer(IOrganizationService service, ITracingService tracingService, IPluginExecutionContext context, Entity postImage)
        {
            Guid jobId = context.PrimaryEntityId;
            tracingService.Trace($"Starting Ready For Server Processing for Job: {jobId}");

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
                tracingService.Trace($"ERROR in Ready For Server: {ex.Message}");
                Entity errorJob = new Entity("ts_workorderexportjob", jobId);
                errorJob["statuscode"] = new OptionSetValue(STATUS_ERROR);
                errorJob["ts_errormessage"] = $"Server Processing Failed: {ex.Message}\nStack: {ex.StackTrace}";
                service.Update(errorJob);
            }
        }

        private void ProcessReadyForMerge(IOrganizationService service, ITracingService tracingService, IPluginExecutionContext context, Entity postImage)
        {
            Guid jobId = context.PrimaryEntityId;
            string jobName = postImage.GetAttributeValue<string>("ts_name");

            tracingService.Trace($"Starting Export Job Processing (Merge) for Job: {jobId} ({jobName})");

            try
            {
                // 1. Retrieve ts_surveypayloadjson (contains Work Order IDs)
                var jobEntity = service.Retrieve("ts_workorderexportjob", jobId, new ColumnSet("ts_surveypayloadjson"));
                string surveyPayload = jobEntity.GetAttributeValue<string>("ts_surveypayloadjson");

                // 2. Parse IDs from ts_surveypayloadjson
                List<Guid> workOrderIds = ParsePayload(surveyPayload, tracingService);
                if (workOrderIds == null || !workOrderIds.Any())
                {
                    throw new InvalidPluginExecutionException("No Work Order IDs found in ts_surveypayloadjson.");
                }

                // 3. Delegate to Service
                var exportService = new WorkOrderExportService(service, tracingService);
                
                // Orchestrate the Merge and Zip process
                exportService.ProcessMergeAndZip(jobId, workOrderIds);
            }
            catch (Exception ex)
            {
                tracingService.Trace($"ERROR in ProcessReadyForMerge (Plugin wrapper): {ex.Message}");
                
                // Ensure status is Error if not already
                // (Optimistically assuming service handled it, but good to double check or just trace)
            }
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
                        
                        // Trim braces if present
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
