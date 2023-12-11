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
            Entity target = (Entity)context.InputParameters["Target"];
            Entity postImageEntity = (context.PostEntityImages != null && context.PostEntityImages.Contains(this.postImageAlias)) ? context.PostEntityImages[this.postImageAlias] : null;

            try
            {
                if (target.LogicalName.Equals(msdyn_workorderservicetask.EntityLogicalName))
                {
                    // If Work Order is Updated - Update any associated files with the new Work Order and Case
                    {
                        if (target.Attributes.Contains("msdyn_workorder"))
                        {
                            using (var serviceContext = new Xrm(localContext.OrganizationService))
                            {
                                // Cast the target to the expected entity
                                msdyn_workorderservicetask myWorkOrderServiceTask = target.ToEntity<msdyn_workorderservicetask>();

                                // Get the selected Work Order Service Task
                                var selectedWorkOrderServiceTask = serviceContext.msdyn_workorderservicetaskSet.Where(wost => wost.Id == myWorkOrderServiceTask.Id).FirstOrDefault();

                                // Get the Work Order associated with the Work Order Service Task
                                var selectedWorkOrder = serviceContext.msdyn_workorderSet.Where(wo => wo.Id == selectedWorkOrderServiceTask.msdyn_WorkOrder.Id).FirstOrDefault();

                                // Retrieve all the files that are associated with the Work Order Service Task
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

                                // Logic to handle ts_sharepointfile and ts_sharepointfilegroup
                                {
                                    // check if the Work Order has a SharePoint File
                                    var myWorkOrderSharePointFile = PostOperationts_sharepointfileCreate.CheckSharePointFile(serviceContext, selectedWorkOrder.Id.ToString().ToUpper().Trim(), PostOperationts_sharepointfileCreate.WORK_ORDER);

                                    // check if the Work Order Service Task has a SharePoint File
                                    var myWorkOrderServiceTaskSharePointFile = PostOperationts_sharepointfileCreate.CheckSharePointFile(serviceContext, selectedWorkOrderServiceTask.Id.ToString().ToUpper().Trim(), PostOperationts_sharepointfileCreate.WORK_ORDER_SERVICE_TASK);

                                    if (myWorkOrderSharePointFile != null || myWorkOrderServiceTaskSharePointFile != null)
                                    {
                                        // get the Owner of the Work Order
                                        string myOwner = "";

                                        // get the owner of the case
                                        string workOrderCaseFetchXML = $@"
                                            <fetch>
                                                <entity name='msdyn_workorder'>
                                                <link-entity name='ovs_operationtype' to='ovs_operationtypeid' from='ovs_operationtypeid' alias='ovs_operationtype' link-type='inner'>
                                                    <link-entity name='team' to='owningteam' from='teamid' alias='team' link-type='inner'>
                                                    <attribute name='name' alias='OWNER_NAME' />
                                                    </link-entity>
                                                </link-entity>
                                                <filter>
                                                    <condition attribute='msdyn_workorderid' operator='eq' value='{selectedWorkOrder.Id.ToString()}' />
                                                </filter>
                                                </entity>
                                            </fetch>                                                
                                        ";

                                        var myWorkOrderEntityCollection = localContext.OrganizationService.RetrieveMultiple(new FetchExpression(workOrderCaseFetchXML));

                                        foreach (var item in myWorkOrderEntityCollection.Entities)
                                        {
                                            if (item.Attributes["OWNER_NAME"] is AliasedValue aliasedOwner)
                                            {
                                                myOwner = aliasedOwner.Value.ToString();
                                            }
                                        }

                                        // if we have a SharePoint File for the Work Order
                                        if (myWorkOrderSharePointFile != null)
                                        {
                                            // check if we have a SharePoint File for the Work Order Service Task
                                            if (myWorkOrderServiceTaskSharePointFile == null)
                                            {
                                                Guid myWorkOrderServiceTaskSharePointFileID = PostOperationts_sharepointfileCreate.CreateSharePointFile(myWorkOrderServiceTask.msdyn_name, PostOperationts_sharepointfileCreate.WORK_ORDER_SERVICE_TASK, PostOperationts_sharepointfileCreate.WORK_ORDER_SERVICE_TASK_FR, myWorkOrderServiceTask.Id.ToString().Trim().ToUpper(), myWorkOrderServiceTask.msdyn_name, myOwner, localContext.OrganizationService);
                                                myWorkOrderServiceTaskSharePointFile = PostOperationts_sharepointfileCreate.CheckSharePointFile(serviceContext, myWorkOrderServiceTaskSharePointFileID.ToString().ToUpper(), PostOperationts_sharepointfileCreate.WORK_ORDER_SERVICE_TASK);
                                            }

                                            // update the Work Order Service Tasks with the SharePointFile Group
                                            localContext.OrganizationService.Update(new ts_SharePointFile
                                            {
                                                Id = myWorkOrderServiceTaskSharePointFile.Id,
                                                ts_SharePointFileGroup = myWorkOrderSharePointFile.ts_SharePointFileGroup
                                            });

                                            // update the Work Order Service Tasks that are related to the Work Order
                                            PostOperationts_sharepointfileCreate.UpdateRelatedWorkOrderServiceTasks(localContext.OrganizationService, selectedWorkOrder.Id, myWorkOrderSharePointFile.ts_SharePointFileGroup.Id, myOwner);
                                        }
                                        else if(myWorkOrderSharePointFile == null && myWorkOrderServiceTaskSharePointFile != null)
                                        {
                                            // check if we have a SharePoint File for the Work Order
                                            if (myWorkOrderSharePointFile == null)
                                            {
                                                Guid myWorkOrderSharePointFileID = PostOperationts_sharepointfileCreate.CreateSharePointFile(selectedWorkOrder.msdyn_name, PostOperationts_sharepointfileCreate.WORK_ORDER, PostOperationts_sharepointfileCreate.WORK_ORDER_FR, selectedWorkOrder.Id.ToString().Trim().ToUpper(), selectedWorkOrder.msdyn_name, myOwner, localContext.OrganizationService);
                                                myWorkOrderSharePointFile = PostOperationts_sharepointfileCreate.CheckSharePointFile(serviceContext, myWorkOrderSharePointFileID.ToString().ToUpper(), PostOperationts_sharepointfileCreate.WORK_ORDER);

                                                // update the SharePoint File Group for the Work Order
                                                localContext.OrganizationService.Update(new ts_SharePointFile
                                                {
                                                    Id = myWorkOrderSharePointFile.Id,
                                                    ts_SharePointFileGroup = myWorkOrderServiceTaskSharePointFile.ts_SharePointFileGroup
                                                });
                                            }

                                            // check if the Work Order has a Case
                                            if (selectedWorkOrder.msdyn_ServiceRequest != null)
                                            {
                                                EntityReference mySharePointFileGroup = null;

                                                var myWorkOrderCase = serviceContext.IncidentSet.Where(c => c.Id == selectedWorkOrder.msdyn_ServiceRequest.Id).FirstOrDefault();

                                                // check if the Case has a SharePointFile
                                                var myWorkOrderCaseSharePointFile = PostOperationts_sharepointfileCreate.CheckSharePointFile(serviceContext, myWorkOrderCase.Id.ToString().ToUpper(), PostOperationts_sharepointfileCreate.CASE);

                                                if (myWorkOrderCaseSharePointFile == null)
                                                {
                                                    // if the Case doesn't have a SharePointFile create it
                                                    Guid myWorkOrderCaseSharePointFileID = PostOperationts_sharepointfileCreate.CreateSharePointFile(myWorkOrderCase.Title, PostOperationts_sharepointfileCreate.CASE, PostOperationts_sharepointfileCreate.CASE_FR, myWorkOrderCase.Id.ToString().Trim().ToUpper(), myWorkOrderCase.Title, myOwner, localContext.OrganizationService);

                                                    // get the SharePointFile
                                                    myWorkOrderCaseSharePointFile = PostOperationts_sharepointfileCreate.CheckSharePointFile(serviceContext, myWorkOrderCase.Id.ToString().Trim().ToUpper(), PostOperationts_sharepointfileCreate.CASE);

                                                    // update the SharePointFileGroup for the Case with the SharePoint File Group that is being used for the Work Order Service Task
                                                    localContext.OrganizationService.Update(new ts_SharePointFile
                                                    {
                                                        Id = myWorkOrderCaseSharePointFileID,
                                                        ts_SharePointFileGroup = myWorkOrderServiceTaskSharePointFile.ts_SharePointFileGroup
                                                    });

                                                    // then update everything related to the Case with the SharePoint File Group
                                                    PostOperationts_sharepointfileCreate.UpdateRelatedWorkOrders(localContext.OrganizationService, myWorkOrderCase.Id, myWorkOrderServiceTaskSharePointFile.ts_SharePointFileGroup.Id, myOwner);
                                                }
                                                else
                                                {
                                                    mySharePointFileGroup = myWorkOrderCaseSharePointFile.ts_SharePointFileGroup;
                                                }
                                            }
                                            else
                                            {
                                                // if not just update the Work Order Service Tasks that are related to the Work Order
                                                PostOperationts_sharepointfileCreate.UpdateRelatedWorkOrderServiceTasks(localContext.OrganizationService, selectedWorkOrder.Id, myWorkOrderServiceTaskSharePointFile.ts_SharePointFileGroup.Id, myOwner);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        // do nothing since the Work Order, and Work Order Service Task don't have a SharePoint File
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception e)
            {
                throw new InvalidPluginExecutionException(e.Message);
            }
        }
    }
}
