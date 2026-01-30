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
    public class PostOperationts_unplannedworkorderCreate : PluginBase
    {
        public PostOperationts_unplannedworkorderCreate(string unsecure, string secure)
            : base(typeof(PostOperationts_unplannedworkorderCreate))
        {
        }

        protected override void ExecuteCrmPlugin(LocalPluginContext localContext)
        {
            var context = localContext.PluginExecutionContext;
            var service = localContext.OrganizationService;

            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity target)
                {
                    // Exit plugin when WO Workspace is created from WO [Edit Ribbon Button] because WO already exists.
                    if (target.Attributes.Contains("ts_skipplugin") && (bool)target.Attributes["ts_skipplugin"] == true)
                    {
                        localContext.Trace("ts_skipplugin is set to true. Exiting plugin.");
                        return;
                    }
                    localContext.Trace("Target entity logical name: {0}", target.LogicalName);
                    if (target.Attributes.Contains("ts_operationtype") || target.Attributes.Contains("ts_region"))
                    {
                        string unplannedWorkOrderId = target.Id.ToString();
                        string ownerName = "";

                        localContext.Trace("Determine if the region is set to International.");
                        var selectedRegion = target.Attributes["ts_region"] as EntityReference;

                        if (selectedRegion != null && selectedRegion.Id.Equals(new Guid("3bf0fa88-150f-eb11-a813-000d3af3a7a7")))
                        {
                            localContext.Trace("Setting business owner to International.");
                            target.Attributes["ts_businessowner"] = "AvSec International";

                            localContext.Trace("Perform the update to the Unplanned Work Order.");
                            service.Update(target);
                        }
                        else
                        {
                            localContext.Trace("Selected region is not International, checking operation type.");
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
                                localContext.Trace("No business owner found for unplanned work order ID. Exit out if no results.");
                                return;
                            }

                            foreach (Entity unplannedWorkOrder in businessNameCollection.Entities)
                            {
                                if (unplannedWorkOrder["OwnerName"] is AliasedValue aliasedValue)
                                {
                                    localContext.Trace("Cast the AliasedValue to string (or the appropriate type).");
                                    ownerName = aliasedValue.Value as string;
                                }

                                localContext.Trace("Set the Business Owner Label.");
                                unplannedWorkOrder["ts_businessowner"] = ownerName;

                                localContext.Trace("Perform the update to the Unplanned Work Order.");
                                service.Update(unplannedWorkOrder);

                                target["ts_businessowner"] = ownerName;
                            }
                        }
                    }

                    //Create WO entity
                    // Ensure the ts_workordertype field exists in the target entity
                    if (target.Attributes.Contains("ts_workordertype"))
                    {
                        localContext.Trace("ts_workordertype field found in ts_unplannedworkorder.");

                        // Retrieve the ts_workordertype value
                        EntityReference workOrderType = target.GetAttributeValue<EntityReference>("ts_workordertype");

                        // Create a new msdyn_workorder entity
                        Entity workOrder = new Entity("msdyn_workorder");

                        // Map fields from ts_unplannedworkorder to msdyn_workorder
                        localContext.Trace("Mapping fields from ts_unplannedworkorder to msdyn_workorder.");

                        workOrder["msdyn_workordertype"] = target.GetAttributeValue<EntityReference>("ts_workordertype");
                        workOrder["ts_region"] = target.GetAttributeValue<EntityReference>("ts_region");
                        workOrder["ovs_operationtypeid"] = target.GetAttributeValue<EntityReference>("ts_operationtype");
                        workOrder["ts_aircraftclassification"] = target.GetAttributeValue<OptionSetValue>("ts_aircraftclassification"); //1
                        workOrder["ts_tradenameid"] = target.GetAttributeValue<EntityReference>("ts_tradename"); //2
                        workOrder["msdyn_serviceaccount"] = target.GetAttributeValue<EntityReference>("ts_stakeholder");
                        workOrder["ts_contact"] = target.GetAttributeValue<EntityReference>("ts_contact"); //0
                        workOrder["ts_site"] = target.GetAttributeValue<EntityReference>("ts_site");
                        workOrder["msdyn_functionallocation"] = target.GetAttributeValue<EntityReference>("ts_functionallocation"); //3
                        workOrder["ts_subsubsite"] = target.GetAttributeValue<EntityReference>("ts_subsubsite"); //4
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
                        workOrder["ts_reportdetails"] = target.GetAttributeValue<string>("ts_reportdetails");
                        workOrder["ownerid"] = target.GetAttributeValue<EntityReference>("ownerid");
                        workOrder["ts_country"] = target.GetAttributeValue<EntityReference>("ts_country");
                        workOrder["ovs_operationid"] = target.GetAttributeValue<EntityReference>("ts_operation");
                        workOrder["msdyn_servicerequest"] = target.GetAttributeValue<EntityReference>("ts_servicerequest");
                        workOrder["ts_securityincident"] = target.GetAttributeValue<EntityReference>("ts_securityincident");
                        workOrder["ts_trip"] = target.GetAttributeValue<EntityReference>("ts_trip");
                        workOrder["msdyn_parentworkorder"] = target.GetAttributeValue<EntityReference>("ts_parentworkorder");
                        workOrder["ovs_fiscalyear"] = target.GetAttributeValue<EntityReference>("ts_plannedfiscalyear");
                        workOrder["ovs_fiscalquarter"] = target.GetAttributeValue<EntityReference>("ts_plannedfiscalquarter");
                        workOrder["ovs_revisedquarterid"] = target.GetAttributeValue<EntityReference>("ts_revisedquarterid");
                        workOrder["ts_canceledinspectionjustification"] = target.GetAttributeValue<EntityReference>("ts_cancelledinspectionjustification");
                        workOrder["ts_othercanceledjustification"] = target.GetAttributeValue<string>("ts_othercancelledjustification");
                        workOrder["ts_scheduledquarterjustification"] = target.GetAttributeValue<EntityReference>("ts_scheduledquarterjustification");
                        workOrder["ts_justificationcomment"] = target.GetAttributeValue<string>("ts_scheduledquarterjustificationcomment");
                        workOrder["ts_details"] = target.GetAttributeValue<string>("ts_details");
                        workOrder["msdyn_instructions"] = target.GetAttributeValue<string>("ts_instructions");
                        workOrder["ts_preparationtime"] = target.GetAttributeValue<decimal>("ts_wopreparationtime");
                        workOrder["ts_woreportinganddocumentation"] = target.GetAttributeValue<decimal>("ts_woreportinganddocumentation");
                        workOrder["ts_comments"] = target.GetAttributeValue<string>("ts_comments");
                        workOrder["ts_overtime"] = target.GetAttributeValue<decimal>("ts_overtime");
                        workOrder["ts_conductingoversight"] = target.GetAttributeValue<decimal>("ts_woconductingoversight");
                        workOrder["ts_traveltime"] = target.GetAttributeValue<decimal>("ts_wotraveltime");
                        workOrder["msdyn_systemstatus"] = target.GetAttributeValue<OptionSetValue>("ts_recordstatus");
                        workOrder["ts_accountableteam"] = target.GetAttributeValue<EntityReference>("ts_accountableteam");

                        localContext.Trace($"Retrieved ts_businessowner: {workOrder.GetAttributeValue<string>("ts_businessowner")}");

                        localContext.Trace("Perform the create the Unplanned Work Order.");
                        Guid workOrderId = service.Create(workOrder);

                        localContext.Trace($"msdyn_workorder created successfully with ID: {workOrderId}");

                        // Retrieve the msdyn_name and Id of the created msdyn_workorder
                        Entity createdWorkOrder = service.Retrieve("msdyn_workorder", workOrderId, new ColumnSet("msdyn_name"));
                        string workOrderName = createdWorkOrder.GetAttributeValue<string>("msdyn_name");

                        localContext.Trace($"Retrieved msdyn_name: {workOrderName}");

                        // Update the ts_unplannedworkorder entity with the msdyn_name
                        Entity unplannedWorkOrder = new Entity("ts_unplannedworkorder", target.Id);
                        unplannedWorkOrder["ts_name"] = workOrderName;
                        unplannedWorkOrder["ts_workorder"] = new EntityReference("msdyn_workorder", workOrderId);
                        service.Update(unplannedWorkOrder);
                        localContext.Trace($"Updated ts_unplannedworkorder with ts_name: {workOrderName} and ts_workorder: {workOrderId}");
                    }
                    else
                    {
                        localContext.Trace("ts_workordertype field not found in ts_unplannedworkorder.");
                    }
                }
            }
            catch (Exception e)
            {
                localContext.Trace("Exception occurred: {0}", e);
                throw new InvalidPluginExecutionException("PostOperationts_unplannedworkorderCreate failed.", e);
            }
        }
    }
}