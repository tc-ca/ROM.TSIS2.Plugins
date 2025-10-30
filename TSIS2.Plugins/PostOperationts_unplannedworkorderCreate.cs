using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using System;
using System.Text;
using System.Threading.Tasks;

namespace TSIS2.Plugins
{
    [CrmPluginRegistration(
    MessageNameEnum.Create,
    "ts_unplannedworkorder",
    StageEnum.PostOperation,
    ExecutionModeEnum.Synchronous,
    "",
    "TSIS2.Plugins.PostOperationts_unplannedworkorderCreate Plugin",
    1,
    IsolationModeEnum.Sandbox,
    Image1Name = "PostImage", Image1Type = ImageTypeEnum.PostImage, Image1Attributes = "",
    Description = "Happens after the Unplanned Work Order has been created")]
    public class PostOperationts_unplannedworkorderCreate : IPlugin
    {
        private readonly string postImageAlias = "PostImage";

        //public PostOperationts_unplannedworkorderCreate(string unsecure, string secure)
        //    : base(typeof(PostOperationts_unplannedworkorderUpdate))
        //{

        //}
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
        public void Execute(IServiceProvider serviceProvider)
        {
            // Services
            var context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            var serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            var service = serviceFactory.CreateOrganizationService(context.UserId);
            var tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));


            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity target)
                {
                    tracingService.Trace("Target entity logical name: {0}", target.LogicalName);
                    if (target.Attributes.Contains("ts_operationtype") || target.Attributes.Contains("ts_region"))
                    {
                        string unplannedWorkOrderId = target.Id.ToString();
                        string ownerName = "";


                        tracingService.Trace("Determine if the region is set to International.");
                        var selectedRegion = target.Attributes["ts_region"] as EntityReference;

                        if (selectedRegion != null && selectedRegion.Id.Equals(new Guid("3bf0fa88-150f-eb11-a813-000d3af3a7a7")))
                        {
                            tracingService.Trace("Setting business owner to International.");
                            target.Attributes["ts_businessowner"] = "AvSec International";

                            tracingService.Trace("Perform the update to the Unplanned Work Order.");
                            // IOrganizationService service = localContext.OrganizationService;

                            service.Update(target);

                            // return;
                        }
                        else
                        {
                            tracingService.Trace("Selected region is not International, checking operation type.");
                            // find out what business owns the Unplanned Work Order
                            string fetchXML = $@"
                                <fetch xmlns:generator='MarkMpn.SQL4CDS'>
                                    <entity name='ts_unplannedworkorder'>
                                    <link-entity name='ovs_operation' to='ts_operation' from='ovs_operationid' alias='ovs_operation' link-type='inner'>
                                    <link-entity name='ovs_operationtype' to='ovs_operationtypeid' from='ovs_operationtypeid' alias='ovs_operationtype' link-type='inner'>
                                    <link-entity name='team' to='owningteam' from='teamid' alias='team' link-type='inner'>
                                    <attribute name='name' alias='OwnerName' />
                                    </link-entity>
                                    </link-entity>
                                    </link-entity>
                                    <filter>
                                    <condition attribute='ts_unplannedworkorderid' operator='eq' value='{unplannedWorkOrderId}' />
                                    </filter>
                                    </entity>
                                </fetch>
                            ";

                            EntityCollection businessNameCollection = service.RetrieveMultiple(new FetchExpression(fetchXML));

                            if (businessNameCollection.Entities.Count == 0)
                            {
                                tracingService.Trace("No business owner found for unplanned work order ID. Exit out if no results.");
                                return;
                            }

                            foreach (Entity unplannedWorkOrder in businessNameCollection.Entities)
                            {
                                if (unplannedWorkOrder["OwnerName"] is AliasedValue aliasedValue)
                                {
                                    tracingService.Trace("Cast the AliasedValue to string (or the appropriate type).");
                                    ownerName = aliasedValue.Value as string;
                                }

                                tracingService.Trace("Set the Business Owner Label.");
                                unplannedWorkOrder["ts_businessowner"] = ownerName;

                                tracingService.Trace("Perform the update to the Unplanned Work Order.");
                                // IOrganizationService service = localContext.OrganizationService;
                                service.Update(unplannedWorkOrder);

                                target["ts_businessowner"] = ownerName;
                            }
                        }

                    }

                    //Create WO entity
                    // Ensure the ts_workordertype field exists in the target entity
                    if (target.Attributes.Contains("ts_workordertype"))
                    {
                        tracingService.Trace("ts_workordertype field found in ts_unplannedworkorder.");

                        // Retrieve the ts_workordertype value
                        EntityReference workOrderType = target.GetAttributeValue<EntityReference>("ts_workordertype");

                        // Create a new msdyn_workorder entity
                        Entity workOrder = new Entity("msdyn_workorder");

                        // Map fields from ts_unplannedworkorder to msdyn_workorder
                        tracingService.Trace("Mapping fields from ts_unplannedworkorder to msdyn_workorder.");

                        workOrder["msdyn_workordertype"] = target.GetAttributeValue<EntityReference>("ts_workordertype");
                        workOrder["ts_region"] = target.GetAttributeValue<EntityReference>("ts_region");
                        workOrder["ovs_operationtypeid"] = target.GetAttributeValue<EntityReference>("ts_operationtype");
                        workOrder["msdyn_serviceaccount"] = target.GetAttributeValue<EntityReference>("ts_stakeholder");
                        workOrder["ts_site"] = target.GetAttributeValue<EntityReference>("ts_site");
                        workOrder["ts_reason"] = target.GetAttributeValue<EntityReference>("ts_reason");
                        workOrder["ts_workorderjustification"] = target.GetAttributeValue<EntityReference>("ts_workorderjustification");
                        workOrder["ts_state"] = target.GetAttributeValue<OptionSetValue>("ts_state");
                        workOrder["msdyn_worklocation"] = target.GetAttributeValue<OptionSetValue>("ts_worklocation");
                        workOrder["ovs_rational"] = target.GetAttributeValue<EntityReference>("ts_rational");
                        workOrder["ts_businessowner"] = target.GetAttributeValue<string>("ts_businessowner");
                        workOrder["msdyn_primaryincidenttype"] = target.GetAttributeValue<EntityReference>("ts_primaryincidenttype");
                        workOrder["msdyn_primaryincidentdescription"] = target.GetAttributeValue<string>("ts_primaryincidentdescription");
                        workOrder["msdyn_primaryincidentestimatedduration"] = target.GetAttributeValue<int>("ts_primaryincidentestimatedduration");
                        workOrder["ts_overtimerequired"] = target.GetAttributeValue<bool>("ts_overtimerequired");
                        workOrder["ownerid"] = target.GetAttributeValue<EntityReference>("ownerid");
                        workOrder["ts_country"] = target.GetAttributeValue<EntityReference>("ts_country");
                        workOrder["ovs_operationid"] = target.GetAttributeValue<EntityReference>("ts_operation");

                        tracingService.Trace($"Retrieved ts_businessowner: {workOrder.GetAttributeValue<string>("ts_businessowner")}");

                        tracingService.Trace("Perform the create the Unplanned Work Order.");
                        // IOrganizationService service = localContext.OrganizationService;
                        // Create the msdyn_workorder entity 
                        Guid workOrderId = service.Create(workOrder);

                        tracingService.Trace($"msdyn_workorder created successfully with ID: {workOrderId}");

                        // Retrieve the msdyn_name and Id of the created msdyn_workorder
                        Entity createdWorkOrder = service.Retrieve("msdyn_workorder", workOrderId, new ColumnSet("msdyn_name"));
                        string workOrderName = createdWorkOrder.GetAttributeValue<string>("msdyn_name");

                        tracingService.Trace($"Retrieved msdyn_name: {workOrderName}");

                        // Update the ts_unplannedworkorder entity with the msdyn_name
                        Entity unplannedWorkOrder = new Entity("ts_unplannedworkorder", target.Id);
                        unplannedWorkOrder["ts_name"] = workOrderName;
                        unplannedWorkOrder["ts_workorder"] = new EntityReference("msdyn_workorder", workOrderId);
                        service.Update(unplannedWorkOrder);
                        tracingService.Trace($"Updated ts_unplannedworkorder with ts_name: {workOrderName} and ts_workorder: {workOrderId}");

                    }
                    else
                    {
                        tracingService.Trace("ts_workordertype field not found in ts_unplannedworkorder.");
                    }
                }
            }
            catch (Exception e)
            {
                tracingService.Trace("Exception occurred: {0}");
                throw new InvalidPluginExecutionException(e.Message);
            }

        }
    }
}
