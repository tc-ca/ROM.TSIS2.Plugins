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
        "ts_unplannedworkorder",
        StageEnum.PostOperation,
        ExecutionModeEnum.Synchronous,
        "",
        "PostOperationts_unplannedworkorderUpdate",
        1,
        IsolationModeEnum.Sandbox,
        Image1Name = "PreImage", Image1Type = ImageTypeEnum.PreImage, Image1Attributes = "ts_workorder,ts_workordertype,ts_region,ts_operation,ts_operationtype,ts_stakeholder,ts_site,ts_state,ts_worklocation,ts_rational,ts_businessowner,ts_primaryincidenttype,ts_primaryincidentdescription,ts_primaryincidentestimatedduration,ts_overtimerequired,ownerid,ts_country",
        Description = "Copies changed fields to the related msdyn_workorder record on update.")]
    public class PostOperationts_unplannedworkorderUpdate : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            var context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            var serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            var service = serviceFactory.CreateOrganizationService(context.UserId);
            var tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            tracingService.Trace("Plugin execution started: PostOperationts_unplannedworkorderUpdate");

            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity target)
                {
                    tracingService.Trace("Target entity logical name: {0}", target.LogicalName);

                    // Get Pre-Image
                    if (!context.PreEntityImages.Contains("PreImage") || context.PreEntityImages["PreImage"] == null)
                    {
                        tracingService.Trace("PreImage is not available. Exiting plugin.");
                        return;
                    }

                    Entity preImage = context.PreEntityImages["PreImage"];

                    // Get the related work order service task reference
                    EntityReference workOrderRef = preImage.GetAttributeValue<EntityReference>("ts_workorder")
                        ?? target.GetAttributeValue<EntityReference>("ts_workorder");

                    if (workOrderRef == null)
                    {
                        tracingService.Trace("ts_workorder reference is null. Exiting plugin.");
                        return;
                    }

                    tracingService.Trace("Updating msdyn_workorder Id: {0}", workOrderRef.Id);

                    Entity updateWorkOrder = new Entity(workOrderRef.LogicalName, workOrderRef.Id);
                    bool anyFieldChanged = false;

                    // Check and update each field if present in Target
                    if (target.Attributes.Contains("ts_workordertype"))
                    {
                        EntityReference workOrderType = target.GetAttributeValue<EntityReference>("ts_workordertype");
                        updateWorkOrder["msdyn_workordertype"] = workOrderType;
                        tracingService.Trace("msdyn_workordertype changed. New value: {0}", workOrderType != null ? workOrderType.Id.ToString() : "null");
                        anyFieldChanged = true;
                    }
                    if (target.Attributes.Contains("ts_region"))
                    {
                        EntityReference region = target.GetAttributeValue<EntityReference>("ts_region");
                        updateWorkOrder["ts_region"] = region;
                        tracingService.Trace("ts_region changed. New value: {0}", region != null ? region.Id.ToString() : "null");

                        // Special logic: if region is International, set business owner to "AvSec International"
                        tracingService.Trace("Determine if the region is set to International.");
                        if (region != null && region.Id.Equals(new Guid("3bf0fa88-150f-eb11-a813-000d3af3a7a7")))
                        {
                            tracingService.Trace("Setting business owner to International.");

                            // Update the business owner for the unplanned work order
                            target.Attributes["ts_businessowner"] = "AvSec International";
                            tracingService.Trace("Perform the update to the Unplanned Work Order.");
                            service.Update(target);

                            // Update the business owner for the related work order
                            updateWorkOrder["ts_businessowner"] = "AvSec International";
                            tracingService.Trace("Perform the update to the Work Order.");
                        }
                        else
                        {
                            tracingService.Trace("Selected region is not International, checking operation type.");
                            string ownerName = "";
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
                                    <condition attribute='ts_unplannedworkorderid' operator='eq' value='{target.Id.ToString()}' />
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


                        anyFieldChanged = true;
                    }
                    if (target.Attributes.Contains("ts_operationtype"))
                    {
                        EntityReference operationType = target.GetAttributeValue<EntityReference>("ts_operationtype");
                        updateWorkOrder["ovs_operationtypeid"] = operationType;
                        tracingService.Trace("ovs_operationtypeid changed. New value: {0}", operationType != null ? operationType.Id.ToString() : "null");
                        anyFieldChanged = true;
                    }
                    if (target.Attributes.Contains("ts_stakeholder"))
                    {
                        EntityReference stakeholder = target.GetAttributeValue<EntityReference>("ts_stakeholder");
                        updateWorkOrder["msdyn_serviceaccount"] = stakeholder;
                        tracingService.Trace("msdyn_serviceaccount changed. New value: {0}", stakeholder != null ? stakeholder.Id.ToString() : "null");
                        anyFieldChanged = true;
                    }
                    if (target.Attributes.Contains("ts_site"))
                    {
                        EntityReference site = target.GetAttributeValue<EntityReference>("ts_site");
                        updateWorkOrder["ts_site"] = site;
                        tracingService.Trace("ts_site changed. New value: {0}", site != null ? site.Id.ToString() : "null");
                        anyFieldChanged = true;
                    }
                    if (target.Attributes.Contains("ts_reason"))
                    {
                        EntityReference reason = target.GetAttributeValue<EntityReference>("ts_reason");
                        updateWorkOrder["ts_reason"] = reason;
                        tracingService.Trace("ts_reason changed. New value: {0}", reason?.Id.ToString() ?? "null");
                        anyFieldChanged = true;
                    }
                    if (target.Attributes.Contains("ts_workorderjustification"))
                    {
                        EntityReference workOrderJustification = target.GetAttributeValue<EntityReference>("ts_workorderjustification");
                        updateWorkOrder["ts_workorderjustification"] = workOrderJustification;
                        tracingService.Trace("ts_workorderjustification updated. New value: {0}", workOrderJustification?.Id.ToString() ?? "null");
                        anyFieldChanged = true;
                    }
                    if (target.Attributes.Contains("ts_state"))
                    {
                        OptionSetValue state = target.GetAttributeValue<OptionSetValue>("ts_state");
                        updateWorkOrder["ts_state"] = state;
                        tracingService.Trace("ts_state changed. New value: {0}", state);
                        anyFieldChanged = true;
                    }
                    if (target.Attributes.Contains("ts_worklocation"))
                    {
                        OptionSetValue workLocation = target.GetAttributeValue<OptionSetValue>("ts_worklocation");
                        updateWorkOrder["msdyn_worklocation"] = workLocation;
                        tracingService.Trace("msdyn_worklocation changed. New value: {0}", workLocation);
                        anyFieldChanged = true;
                    }
                    if (target.Attributes.Contains("ts_rational"))
                    {
                        EntityReference rational = target.GetAttributeValue<EntityReference>("ts_rational");
                        updateWorkOrder["ts_rational"] = rational;
                        tracingService.Trace("ts_rational changed. New value: {0}", rational != null ? rational.Id.ToString() : "null");
                        anyFieldChanged = true;
                    }                    
                    //if (target.Attributes.Contains("ts_businessowner"))
                    //{
                    //    string businessowner = target.GetAttributeValue<string>("ts_businessowner");
                    //    updateWorkOrder["ts_businessowner"] = businessowner;
                    //    tracingService.Trace("ts_businessowner changed. New value: {0}", businessowner);
                    //    anyFieldChanged = true;
                    //}
                    if (target.Attributes.Contains("ts_primaryincidenttype"))
                    {
                        EntityReference primaryIncidentType = target.GetAttributeValue<EntityReference>("ts_primaryincidenttype");
                        updateWorkOrder["msdyn_primaryincidenttype"] = primaryIncidentType;
                        tracingService.Trace("msdyn_primaryincidenttype changed. New value: {0}", primaryIncidentType != null ? primaryIncidentType.Id.ToString() : "null");
                        anyFieldChanged = true;
                    }
                    if (target.Attributes.Contains("ts_primaryincidentdescription"))
                    {
                        string primaryIncidentDescription = target.GetAttributeValue<string>("ts_primaryincidentdescription");
                        updateWorkOrder["msdyn_primaryincidentdescription"] = primaryIncidentDescription;
                        tracingService.Trace("msdyn_primaryincidentdescription changed. New value: {0}", primaryIncidentDescription);
                        anyFieldChanged = true;
                    }
                    if (target.Attributes.Contains("ts_primaryincidentestimatedduration"))
                    {
                        int primaryIncidentEstimatedDuration = target.GetAttributeValue<int>("ts_primaryincidentestimatedduration");
                        updateWorkOrder["msdyn_primaryincidentestimatedduration"] = primaryIncidentEstimatedDuration;
                        tracingService.Trace("msdyn_primaryincidentestimatedduration changed. New value: {0}", primaryIncidentEstimatedDuration);
                        anyFieldChanged = true;
                    }
                    if (target.Attributes.Contains("ts_overtimerequired"))
                    {
                        bool overtimeRequired = target.GetAttributeValue<bool>("ts_overtimerequired");
                        updateWorkOrder["ts_overtimerequired"] = overtimeRequired;
                        tracingService.Trace("ts_overtimerequired changed. New value: {0}", overtimeRequired);
                        anyFieldChanged = true;
                    }
                    if (target.Attributes.Contains("ownerid"))
                    {
                        EntityReference owner = target.GetAttributeValue<EntityReference>("ownerid");
                        updateWorkOrder["ownerid"] = owner;
                        tracingService.Trace("ownerid changed. New value: {0}", owner != null ? owner.Id.ToString() : "null");
                        anyFieldChanged = true;
                    }
                    if (target.Attributes.Contains("ts_country"))
                    {
                        EntityReference country = target.GetAttributeValue<EntityReference>("ts_country");
                        updateWorkOrder["ts_country"] = country;
                        tracingService.Trace("ts_country changed. New value: {0}", country != null ? country.Id.ToString() : "null");
                        anyFieldChanged = true;
                    }
                    if (anyFieldChanged)
                    {
                        service.Update(updateWorkOrder);
                        tracingService.Trace("Updated msdyn_workorder with changed fields.");
                    }
                    else
                    {
                        tracingService.Trace("No relevant fields changed. No update performed.");
                    }

                }
                else
                {
                    tracingService.Trace("No target entity found. Exiting plugin.");
                }
            }
            catch (Exception ex)
            {
                tracingService.Trace("An error occurred: {0}", ex.ToString());
                throw;
            }
        }
    }
}