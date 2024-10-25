using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TSIS2.Plugins
{
    [CrmPluginRegistration(
    MessageNameEnum.Update,
    "msdyn_workorderservicetask",
    StageEnum.PostOperation,
    ExecutionModeEnum.Synchronous,
    "msdyn_workorder",
    "TSIS2.Plugins.PostOperationmsdyn_workorderservicetaskUpdate Plugin",
    1,
    IsolationModeEnum.Sandbox,
    Image1Name = "PostImage", Image1Type = ImageTypeEnum.PostImage, Image1Attributes = "msdyn_name,msdyn_workorder",
    Description = "If a Work Order Service Task has been moved to another Work Order, update the associated files with the new Work Order and Case")]
    /// <summary>
    /// PostOperationmsdyn_workorderservicetaskUpdate Plugin.
    /// </summary>  
    public class PostOperationmsdyn_workorderservicetaskUpdate : PluginBase
    {
        private readonly string postImageAlias = "PostImage";

        /// <summary>
        /// Initializes a new instance of the <see cref="PostOperationmsdyn_workorderservicetaskUpdate"/> class.
        /// </summary>
        /// <param name="unsecure">Contains public (unsecured) configuration information.</param>
        /// <param name="secure">Contains non-public (secured) configuration information. 
        /// When using Microsoft Dynamics 365 for Outlook with Offline Access, 
        /// the secure string is not passed to a plug-in that executes while the client is offline.</param>
        public PostOperationmsdyn_workorderservicetaskUpdate(string unsecure, string secure)
            : base(typeof(PostOperationts_workorderactivitytypeUpdate))
        {
            //if (secure != null &&!secure.Equals(string.Empty))
            //{

            //}
        }

        /// <summary>
        /// Main entry point for he business logic that the plug-in is to execute.
        /// </summary>
        /// <param name="localContext">The <see cref="LocalPluginContext"/> which contains the
        /// <see cref="IPluginExecutionContext"/>,
        /// <see cref="IOrganizationService"/>
        /// and <see cref="ITracingService"/>
        /// </param>
        /// <remarks>
        /// For improved performance, Microsoft Dynamics 365 caches plug-in instances.
        /// The plug-in's Execute method should be written to be stateless as the constructor
        /// is not called for every invocation of the plug-in. Also, multiple system threads
        /// could execute the plug-in at the same time. All per invocation state information
        /// is stored in the context. This means that you should not use global variables in plug-ins.
        /// </remarks>
        protected override void ExecuteCrmPlugin(LocalPluginContext localContext)
        {
            if (localContext == null)
            {
                throw new InvalidPluginExecutionException("localContext");
            }

            IPluginExecutionContext context = localContext.PluginExecutionContext;
            ITracingService tracingService = localContext.TracingService;
            Entity target = (Entity)context.InputParameters["Target"];
            Entity postImageEntity = (context.PostEntityImages != null && context.PostEntityImages.Contains(this.postImageAlias)) ? context.PostEntityImages[this.postImageAlias] : null;

            tracingService.Trace("Entering ExecuteCrmPlugin method.");
            try
            {
                if (target.LogicalName.Equals(msdyn_workorderservicetask.EntityLogicalName))
                {
                    tracingService.Trace("If Work Order is Updated - Update any associated files with the new Work Order and Case.");
                    {
                        if (target.Attributes.Contains("msdyn_workorder"))
                        {
                            using (var serviceContext = new Xrm(localContext.OrganizationService))
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
                                        localContext.OrganizationService.Update(new ts_File
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
                                        string myOwner = PostOperationts_sharepointfileCreate.GetWorkOrderOwner(localContext.OrganizationService, selectedWorkOrder.Id);

                                        tracingService.Trace("If we have a SharePoint File for the Work Order.");
                                        if (myWorkOrderSharePointFile != null)
                                        {
                                            tracingService.Trace("Check if we have a SharePoint File for the Work Order Service Task.");
                                            if (myWorkOrderServiceTaskSharePointFile == null)
                                            {
                                                Guid myWorkOrderServiceTaskSharePointFileID = PostOperationts_sharepointfileCreate.CreateSharePointFile(myWorkOrderServiceTask.msdyn_name, PostOperationts_sharepointfileCreate.WORK_ORDER_SERVICE_TASK, PostOperationts_sharepointfileCreate.WORK_ORDER_SERVICE_TASK_FR, myWorkOrderServiceTask.Id.ToString().Trim().ToUpper(), myWorkOrderServiceTask.msdyn_name, myOwner, localContext.OrganizationService);
                                                myWorkOrderServiceTaskSharePointFile = PostOperationts_sharepointfileCreate.CheckSharePointFile(serviceContext, myWorkOrderServiceTaskSharePointFileID.ToString().ToUpper(), PostOperationts_sharepointfileCreate.WORK_ORDER_SERVICE_TASK);
                                            }

                                            tracingService.Trace("Update the Work Order Service Tasks with the SharePointFile Group.");
                                            localContext.OrganizationService.Update(new ts_SharePointFile
                                            {
                                                Id = myWorkOrderServiceTaskSharePointFile.Id,
                                                ts_SharePointFileGroup = myWorkOrderSharePointFile.ts_SharePointFileGroup
                                            });

                                            // The assumption is that if a Work Order has a SharePoint file, then it has the SharePoint File Group from the Case
                                            tracingService.Trace("Update the Work Order Service Tasks that are related to the Work Order.");
                                            PostOperationts_sharepointfileCreate.UpdateRelatedWorkOrderServiceTasks(localContext.OrganizationService, selectedWorkOrder.Id, myWorkOrderSharePointFile.ts_SharePointFileGroup.Id, myOwner);
                                        }
                                        else if (myWorkOrderSharePointFile == null && myWorkOrderServiceTaskSharePointFile != null)
                                        {
                                            tracingService.Trace("Create the SharePoint File for the Work Order.");
                                            Guid myWorkOrderSharePointFileID = PostOperationts_sharepointfileCreate.CreateSharePointFile(selectedWorkOrder.msdyn_name, PostOperationts_sharepointfileCreate.WORK_ORDER, PostOperationts_sharepointfileCreate.WORK_ORDER_FR, selectedWorkOrder.Id.ToString().Trim().ToUpper(), selectedWorkOrder.msdyn_name, myOwner, localContext.OrganizationService);
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
                                                    Guid myWorkOrderCaseSharePointFileID = PostOperationts_sharepointfileCreate.CreateSharePointFile(myWorkOrderCase.Title, PostOperationts_sharepointfileCreate.CASE, PostOperationts_sharepointfileCreate.CASE_FR, myWorkOrderCase.Id.ToString().Trim().ToUpper(), myWorkOrderCase.Title, myOwner, localContext.OrganizationService);

                                                    tracingService.Trace("Get the SharePointFile.");
                                                    myWorkOrderCaseSharePointFile = PostOperationts_sharepointfileCreate.CheckSharePointFile(serviceContext, myWorkOrderCase.Id.ToString().Trim().ToUpper(), PostOperationts_sharepointfileCreate.CASE);

                                                    tracingService.Trace("Create the SharePoint File Group for the Case.");
                                                    Guid myWorkOrderCaseSharePointFileGroupID = PostOperationts_sharepointfileCreate.CreateSharePointFileGroup(myWorkOrderCaseSharePointFile, localContext.OrganizationService);

                                                    tracingService.Trace("Update everything related to the Case with the SharePoint File Group.");
                                                    PostOperationts_sharepointfileCreate.UpdateRelatedWorkOrders(localContext.OrganizationService, myWorkOrderCase.Id, myWorkOrderCaseSharePointFileGroupID, myOwner);
                                                }
                                                else
                                                {
                                                    mySharePointFileGroup = myWorkOrderCaseSharePointFile.ts_SharePointFileGroup;

                                                    tracingService.Trace("(Else) Update everything related to the Case with the SharePoint File Group.");
                                                    PostOperationts_sharepointfileCreate.UpdateRelatedWorkOrders(localContext.OrganizationService, myWorkOrderCase.Id, myWorkOrderServiceTaskSharePointFile.ts_SharePointFileGroup.Id, myOwner);
                                                }
                                            }
                                            else
                                            {
                                                tracingService.Trace("Create the SharePoint File Group for the Work Order.");
                                                Guid myWorkOrderSharePointFileGroupID = PostOperationts_sharepointfileCreate.CreateSharePointFileGroup(myWorkOrderSharePointFile, localContext.OrganizationService);

                                                tracingService.Trace("Update the Work Order Service Tasks that are related to the Work Order.");
                                                PostOperationts_sharepointfileCreate.UpdateRelatedWorkOrderServiceTasks(localContext.OrganizationService, selectedWorkOrder.Id, myWorkOrderSharePointFileGroupID, myOwner);
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
