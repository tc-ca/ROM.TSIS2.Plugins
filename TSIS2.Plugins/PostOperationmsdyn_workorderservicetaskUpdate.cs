using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TSIS2.Plugins.QuestionnaireExtractor;

namespace TSIS2.Plugins
{
    [CrmPluginRegistration(
    MessageNameEnum.Update,
    "msdyn_workorderservicetask",
    StageEnum.PostOperation,
    ExecutionModeEnum.Synchronous,
    "",
    "TSIS2.Plugins.PostOperationmsdyn_workorderservicetaskUpdate Plugin",
    1,
    IsolationModeEnum.Sandbox,
    Image1Name = "PostImage", Image1Type = ImageTypeEnum.PostImage, Image1Attributes = "msdyn_name,msdyn_workorder",
    Image2Name = "PreImage", Image2Type = ImageTypeEnum.PreImage, Image2Attributes = "statuscode,ovs_questionnaireresponse,msdyn_workorder",
    Description = "If a Work Order Service Task has been moved to another Work Order, update the associated files with the new Work Order and Case; Also creates individual question response records from the questionnaire when a task is completed")]
    /// <summary>
    /// PostOperationmsdyn_workorderservicetaskUpdate Plugin.
    /// </summary>  
    public class PostOperationmsdyn_workorderservicetaskUpdate : IPlugin
    {
        private readonly string postImageAlias = "PostImage";
        private readonly string preImageAlias = "PreImage";

        /// <summary>
        /// Initializes a new instance of the <see cref="PostOperationmsdyn_workorderservicetaskUpdate"/> class.
        /// </summary>
        /// <param name="unsecure">Contains public (unsecured) configuration information.</param>
        /// <param name="secure">Contains non-public (secured) configuration information. 
        /// When using Microsoft Dynamics 365 for Outlook with Offline Access, 
        /// the secure string is not passed to a plug-in that executes while the client is offline.</param>
        public PostOperationmsdyn_workorderservicetaskUpdate(string unsecure, string secure)
        {
            //if (secure != null &&!secure.Equals(string.Empty))
            //{

            //}
        }

        /// <summary>
        /// Main entry point for he business logic that the plug-in is to execute.
        /// </summary>
        /// <param name="serviceProvider">The service provider.</param>
        /// <remarks>
        /// For improved performance, Microsoft Dynamics 365 caches plug-in instances.
        /// The plug-in's Execute method should be written to be stateless as the constructor
        /// is not called for every invocation of the plug-in. Also, multiple system threads
        /// could execute the plug-in at the same time. All per invocation state information
        /// is stored in the context. This means that you should not use global variables in plug-ins.
        /// </remarks>
        public void Execute(IServiceProvider serviceProvider)
        {
            if (serviceProvider == null)
            {
                throw new InvalidPluginExecutionException("serviceProvider");
            }

            // Obtain the tracing service
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

            // Return if triggered by another plugin. Prevents infinite loop.
            if (context.Depth > 2)
            {
                tracingService.Trace("Exiting - plugin depth > 1");
                return;
            }

            // Obtain the organization service
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            Entity target = (Entity)context.InputParameters["Target"];
            Entity postImageEntity = (context.PostEntityImages != null && context.PostEntityImages.Contains(this.postImageAlias)) ? context.PostEntityImages[this.postImageAlias] : null;
            Entity preImageEntity = (context.PreEntityImages != null && context.PreEntityImages.Contains(this.preImageAlias)) ? context.PreEntityImages[this.preImageAlias] : null;

            tracingService.Trace("Entering Execute method.");
            try
            {
                if (target.LogicalName.Equals(msdyn_workorderservicetask.EntityLogicalName))
                {
                    tracingService.Trace("If Work Order is Updated - Update any associated files with the new Work Order and Case.");
                    {
                        // Check if msdyn_workorder attribute exists in the update
                        if (target.Attributes.Contains("msdyn_workorder"))
                        {
                            // Check if the Work Order has actually changed by comparing with preImage
                            EntityReference previousWorkOrder = preImageEntity?.GetAttributeValue<EntityReference>("msdyn_workorder");
                            EntityReference currentWorkOrder = target.GetAttributeValue<EntityReference>("msdyn_workorder");

                            // Only proceed with the file update logic if the Work Order has changed
                            if (previousWorkOrder == null || currentWorkOrder == null ||
                                previousWorkOrder.Id != currentWorkOrder.Id)
                            {
                                tracingService.Trace("Work Order has changed. Updating associated files.");

                                using (var serviceContext = new Xrm(service))
                                {
                                    tracingService.Trace("Cast the target to the expected entity.");
                                    msdyn_workorderservicetask myWorkOrderServiceTask = target.ToEntity<msdyn_workorderservicetask>();

                                    tracingService.Trace("Get the selected Work Order Service Task.");
                                    var selectedWorkOrderServiceTask = serviceContext.msdyn_workorderservicetaskSet.Where(wost => wost.Id == myWorkOrderServiceTask.Id).FirstOrDefault();

                                    tracingService.Trace("Get the Work Order associated with the Work Order Service Task.");
                                    var selectedWorkOrder = serviceContext.msdyn_workorderSet.Where(wo => wo.Id == selectedWorkOrderServiceTask.msdyn_WorkOrder.Id).FirstOrDefault();

                                    tracingService.Trace("Retrieve all the files that are associated with the Work Order Service Task.");
                                    var allFiles = serviceContext.ts_FileSet.ToList();
                                    var workOrderServiceTaskFiles = allFiles.Where(f => f.ts_formintegrationid != null && f.ts_formintegrationid.Replace("WOST ", "").Trim() == selectedWorkOrderServiceTask.msdyn_name).ToList();

                                    if (workOrderServiceTaskFiles != null)
                                    {
                                        foreach (var file in workOrderServiceTaskFiles)
                                        {
                                            service.Update(new ts_File
                                            {
                                                Id = file.Id,
                                                ts_msdyn_workorder = selectedWorkOrder.ToEntityReference(),
                                                ts_Incident = selectedWorkOrder.msdyn_ServiceRequest
                                            });
                                        }
                                    }

                                    tracingService.Trace("Logic to handle ts_sharepointfile and ts_sharepointfilegroup.");
                                    {
                                        tracingService.Trace("Check if the Work Order has a SharePoint File.");
                                        var myWorkOrderSharePointFile = PostOperationts_sharepointfileCreate.CheckSharePointFile(serviceContext, selectedWorkOrder.Id.ToString().ToUpper().Trim(), PostOperationts_sharepointfileCreate.WORK_ORDER);

                                        tracingService.Trace("Check if the Work Order Service Task has a SharePoint File.");
                                        var myWorkOrderServiceTaskSharePointFile = PostOperationts_sharepointfileCreate.CheckSharePointFile(serviceContext, selectedWorkOrderServiceTask.Id.ToString().ToUpper().Trim(), PostOperationts_sharepointfileCreate.WORK_ORDER_SERVICE_TASK);

                                        if (myWorkOrderSharePointFile != null || myWorkOrderServiceTaskSharePointFile != null)
                                        {
                                            tracingService.Trace("Get the Owner of the Work Order.");
                                            string myOwner = PostOperationts_sharepointfileCreate.GetWorkOrderOwner(service, selectedWorkOrder.Id);

                                            tracingService.Trace("If we have a SharePoint File for the Work Order.");
                                            if (myWorkOrderSharePointFile != null)
                                            {
                                                tracingService.Trace("Check if we have a SharePoint File for the Work Order Service Task.");
                                                if (myWorkOrderServiceTaskSharePointFile == null)
                                                {
                                                    Guid myWorkOrderServiceTaskSharePointFileID = PostOperationts_sharepointfileCreate.CreateSharePointFile(myWorkOrderServiceTask.msdyn_name, PostOperationts_sharepointfileCreate.WORK_ORDER_SERVICE_TASK, PostOperationts_sharepointfileCreate.WORK_ORDER_SERVICE_TASK_FR, myWorkOrderServiceTask.Id.ToString().Trim().ToUpper(), myWorkOrderServiceTask.msdyn_name, myOwner, service);
                                                    myWorkOrderServiceTaskSharePointFile = PostOperationts_sharepointfileCreate.CheckSharePointFile(serviceContext, myWorkOrderServiceTaskSharePointFileID.ToString().ToUpper(), PostOperationts_sharepointfileCreate.WORK_ORDER_SERVICE_TASK);
                                                }

                                                tracingService.Trace("Update the Work Order Service Tasks with the SharePointFile Group.");
                                                service.Update(new ts_SharePointFile
                                                {
                                                    Id = myWorkOrderServiceTaskSharePointFile.Id,
                                                    ts_SharePointFileGroup = myWorkOrderSharePointFile.ts_SharePointFileGroup
                                                });

                                                // The assumption is that if a Work Order has a SharePoint file, then it has the SharePoint File Group from the Case
                                                tracingService.Trace("Update the Work Order Service Tasks that are related to the Work Order.");
                                                PostOperationts_sharepointfileCreate.UpdateRelatedWorkOrderServiceTasks(service, selectedWorkOrder.Id, myWorkOrderSharePointFile.ts_SharePointFileGroup.Id, myOwner);
                                            }
                                            else if (myWorkOrderSharePointFile == null && myWorkOrderServiceTaskSharePointFile != null)
                                            {
                                                tracingService.Trace("Create the SharePoint File for the Work Order.");
                                                Guid myWorkOrderSharePointFileID = PostOperationts_sharepointfileCreate.CreateSharePointFile(selectedWorkOrder.msdyn_name, PostOperationts_sharepointfileCreate.WORK_ORDER, PostOperationts_sharepointfileCreate.WORK_ORDER_FR, selectedWorkOrder.Id.ToString().Trim().ToUpper(), selectedWorkOrder.msdyn_name, myOwner, service);
                                                myWorkOrderSharePointFile = PostOperationts_sharepointfileCreate.CheckSharePointFile(serviceContext, selectedWorkOrder.Id.ToString().ToUpper(), PostOperationts_sharepointfileCreate.WORK_ORDER);

                                                tracingService.Trace("Check if the Work Order has a Case.");
                                                if (selectedWorkOrder.msdyn_ServiceRequest != null)
                                                {
                                                    EntityReference mySharePointFileGroup = null;

                                                    var myWorkOrderCase = serviceContext.IncidentSet.Where(c => c.Id == selectedWorkOrder.msdyn_ServiceRequest.Id).FirstOrDefault();

                                                    tracingService.Trace("Check if the Case has a SharePointFile.");
                                                    var myWorkOrderCaseSharePointFile = PostOperationts_sharepointfileCreate.CheckSharePointFile(serviceContext, myWorkOrderCase.Id.ToString().ToUpper(), PostOperationts_sharepointfileCreate.CASE);

                                                    if (myWorkOrderCaseSharePointFile == null)
                                                    {
                                                        tracingService.Trace("If the Case doesn't have a SharePointFile, create it.");
                                                        Guid myWorkOrderCaseSharePointFileID = PostOperationts_sharepointfileCreate.CreateSharePointFile(myWorkOrderCase.Title, PostOperationts_sharepointfileCreate.CASE, PostOperationts_sharepointfileCreate.CASE_FR, myWorkOrderCase.Id.ToString().Trim().ToUpper(), myWorkOrderCase.Title, myOwner, service);

                                                        tracingService.Trace("Get the SharePointFile.");
                                                        myWorkOrderCaseSharePointFile = PostOperationts_sharepointfileCreate.CheckSharePointFile(serviceContext, myWorkOrderCase.Id.ToString().Trim().ToUpper(), PostOperationts_sharepointfileCreate.CASE);

                                                        tracingService.Trace("Create the SharePoint File Group for the Case.");
                                                        Guid myWorkOrderCaseSharePointFileGroupID = PostOperationts_sharepointfileCreate.CreateSharePointFileGroup(myWorkOrderCaseSharePointFile, service);

                                                        tracingService.Trace("Update everything related to the Case with the SharePoint File Group.");
                                                        PostOperationts_sharepointfileCreate.UpdateRelatedWorkOrders(service, myWorkOrderCase.Id, myWorkOrderCaseSharePointFileGroupID, myOwner);
                                                    }
                                                    else
                                                    {
                                                        mySharePointFileGroup = myWorkOrderCaseSharePointFile.ts_SharePointFileGroup;

                                                        tracingService.Trace("(Else) Update everything related to the Case with the SharePoint File Group.");
                                                        PostOperationts_sharepointfileCreate.UpdateRelatedWorkOrders(service, myWorkOrderCase.Id, myWorkOrderServiceTaskSharePointFile.ts_SharePointFileGroup.Id, myOwner);
                                                    }
                                                }
                                                else
                                                {
                                                    tracingService.Trace("Create the SharePoint File Group for the Work Order.");
                                                    Guid myWorkOrderSharePointFileGroupID = PostOperationts_sharepointfileCreate.CreateSharePointFileGroup(myWorkOrderSharePointFile, service);

                                                    tracingService.Trace("Update the Work Order Service Tasks that are related to the Work Order.");
                                                    PostOperationts_sharepointfileCreate.UpdateRelatedWorkOrderServiceTasks(service, selectedWorkOrder.Id, myWorkOrderSharePointFileGroupID, myOwner);
                                                }
                                            }
                                        }
                                        else
                                        {
                                            tracingService.Trace("Work Order and Work Order Service Task don't have a SharePoint File.");
                                        }
                                    }
                                }
                            }
                        }

                        if (target.Contains("statuscode") &&
                            target.GetAttributeValue<OptionSetValue>("statuscode").Value == 918640002) //  "Complete"
                        {
                            // Check previous status
                            bool isFirstTimeCompletion = preImageEntity?.GetAttributeValue<OptionSetValue>("statuscode")?.Value == 918640004; // Was In Progress
                            bool isRecompletion = preImageEntity?.GetAttributeValue<OptionSetValue>("statuscode")?.Value == 1; // Was Active

                            if (isFirstTimeCompletion || isRecompletion)
                            {
                                tracingService.Trace("Work Order Service Task is being completed. Starting questionnaire processing.");
                                tracingService.Trace($"First time completion: {isFirstTimeCompletion}, Recompletion: {isRecompletion}");

                                // Create a logger adapter to pass the tracing service to the orchestrator
                                var logger = new TracingServiceAdapter(tracingService, LogLevel.Info);

                                // Get both questionnaire response and questionnaire reference from the WOST
                                var wost = service.Retrieve("msdyn_workorderservicetask", target.Id,
                                    new ColumnSet("ovs_questionnaireresponse", "ovs_questionnaire", "ovs_questionnairedefinition"));

                                // For recompletion check, get current response JSON
                                if (isRecompletion)
                                {
                                    string currentResponseJson = wost.GetAttributeValue<string>("ovs_questionnaireresponse");
                                    string previousResponseJson = preImageEntity.GetAttributeValue<string>("ovs_questionnaireresponse");

                                    if (string.Equals(currentResponseJson, previousResponseJson, StringComparison.Ordinal))
                                    {
                                        tracingService.Trace("Questionnaire response hasn't changed. Skipping processing.");
                                        return;
                                    }
                                    tracingService.Trace("Questionnaire response has changed. Processing updates.");
                                }

                                var questionnaireRef = wost.GetAttributeValue<EntityReference>("ovs_questionnaire");

                                if (questionnaireRef != null)
                                {
                                    tracingService.Trace($"Starting questionnaire processing for WOST: {target.Id}");

                                    // Process questionnaire using the new, self-contained orchestrator.
                                    // It handles creating, updating, and linking the response records in one go.
                                    var questionResponseIds = QuestionnaireOrchestrator.ProcessQuestionnaire(
                                        service,
                                        target.Id,
                                        questionnaireRef, // The orchestrator requires the questionnaire reference
                                        isRecompletion,
                                        false, // Simulation mode is false for plugins
                                        logger // Pass the wrapped tracing service
                                    );

                                    tracingService.Trace($"Successfully processed questionnaire. {questionResponseIds.Count} question response records were created/updated.");
                                }
                                else
                                {
                                    tracingService.Trace("No questionnaire reference found on WOST. Skipping processing.");
                                }
                            }
                        }
                        

                    }
                }
            }
            catch (Exception e)
            {
                tracingService.Trace("Exception occurred: {0}", e.ToString());
                throw new InvalidPluginExecutionException(e.Message);
            }
        }
    }
}
