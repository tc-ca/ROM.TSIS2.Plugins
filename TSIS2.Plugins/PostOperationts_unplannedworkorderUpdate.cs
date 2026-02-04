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
        Image1Name = "PreImage", Image1Type = ImageTypeEnum.PreImage, Image1Attributes = "ts_workorder,ts_workordertype,ts_region,ts_operation,ts_operationtype,ts_stakeholder,ts_site,ts_state,ts_worklocation,ts_rational,ts_businessowner,ts_primaryincidenttype,ts_primaryincidentdescription,ts_primaryincidentestimatedduration,ts_overtimerequired,ownerid,ts_country, ts_plannedfiscalyear, ts_plannedfiscalquarter",
        Description = "Copies changed fields to the related msdyn_workorder record on update.")]
    public class PostOperationts_unplannedworkorderUpdate : PluginBase
    {
        public PostOperationts_unplannedworkorderUpdate(string unsecure, string secure)
            : base(typeof(PostOperationts_unplannedworkorderUpdate))
        {
        }

        protected override void ExecuteCrmPlugin(LocalPluginContext localContext)
        {
            var context = localContext.PluginExecutionContext;
            var service = localContext.OrganizationService;

            localContext.Trace("Plugin execution started: PostOperationts_unplannedworkorderUpdate");

            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity target)
                {
                    localContext.Trace("Target entity logical name: {0}", target.LogicalName);

                    // Get Pre-Image
                    if (!context.PreEntityImages.Contains("PreImage") || context.PreEntityImages["PreImage"] == null)
                    {
                        localContext.Trace("PreImage is not available. Exiting plugin.");
                        return;
                    }

                    Entity preImage = context.PreEntityImages["PreImage"];

                    // Get the related work order service task reference
                    EntityReference workOrderRef = preImage.GetAttributeValue<EntityReference>("ts_workorder")
                        ?? target.GetAttributeValue<EntityReference>("ts_workorder");

                    if (workOrderRef == null)
                    {
                        localContext.Trace("ts_workorder reference is null. Exiting plugin.");
                        return;
                    }

                    localContext.Trace("Updating msdyn_workorder Id: {0}", workOrderRef.Id);

                    Entity updateWorkOrder = new Entity(workOrderRef.LogicalName, workOrderRef.Id);
                    bool anyFieldChanged = false;

                    // Check and update each field if present in Target
                    if (target.Attributes.Contains("ts_workordertype"))
                    {
                        EntityReference workOrderType = target.GetAttributeValue<EntityReference>("ts_workordertype");
                        updateWorkOrder["msdyn_workordertype"] = workOrderType;
                        localContext.Trace("msdyn_workordertype changed. New value: {0}", workOrderType != null ? workOrderType.Id.ToString() : "null");
                        anyFieldChanged = true;
                    }
                    if (target.Attributes.Contains("ts_region"))
                    {
                        EntityReference region = target.GetAttributeValue<EntityReference>("ts_region");
                        updateWorkOrder["ts_region"] = region;
                        localContext.Trace("ts_region changed. New value: {0}", region != null ? region.Id.ToString() : "null");

                        // Special logic: if region is International, set business owner to "AvSec International"
                        localContext.Trace("Determine if the region is set to International.");
                        if (region != null && region.Id.Equals(new Guid("3bf0fa88-150f-eb11-a813-000d3af3a7a7")))
                        {
                            localContext.Trace("Setting business owner to International.");

                            // Update the business owner for the unplanned work order
                            target.Attributes["ts_businessowner"] = "AvSec International";
                            localContext.Trace("Perform the update to the Unplanned Work Order.");
                            service.Update(target);

                            // Update the business owner for the related work order
                            updateWorkOrder["ts_businessowner"] = "AvSec International";
                            localContext.Trace("Perform the update to the Work Order.");
                        }
                        else
                        {
                            localContext.Trace("Selected region is not International, checking operation type.");
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

                        anyFieldChanged = true;
                    }
                    if (target.Attributes.Contains("ts_operationtype"))
                    {
                        EntityReference operationType = target.GetAttributeValue<EntityReference>("ts_operationtype");
                        updateWorkOrder["ovs_operationtypeid"] = operationType;
                        localContext.Trace("ovs_operationtypeid changed. New value: {0}", operationType != null ? operationType.Id.ToString() : "null");
                        anyFieldChanged = true;
                    }
                    if (target.Attributes.Contains("ts_stakeholder"))
                    {
                        EntityReference stakeholder = target.GetAttributeValue<EntityReference>("ts_stakeholder");
                        updateWorkOrder["msdyn_serviceaccount"] = stakeholder;
                        localContext.Trace("msdyn_serviceaccount changed. New value: {0}", stakeholder != null ? stakeholder.Id.ToString() : "null");
                        anyFieldChanged = true;
                    }
                    if (target.Attributes.Contains("ts_site"))
                    {
                        EntityReference site = target.GetAttributeValue<EntityReference>("ts_site");
                        updateWorkOrder["ts_site"] = site;
                        localContext.Trace("ts_site changed. New value: {0}", site != null ? site.Id.ToString() : "null");
                        anyFieldChanged = true;
                    }
                    if (target.Attributes.Contains("ts_reason"))
                    {
                        EntityReference reason = target.GetAttributeValue<EntityReference>("ts_reason");
                        updateWorkOrder["ts_reason"] = reason;
                        localContext.Trace("ts_reason changed. New value: {0}", reason?.Id.ToString() ?? "null");
                        anyFieldChanged = true;
                    }
                    if (target.Attributes.Contains("ts_workorderjustification"))
                    {
                        EntityReference workOrderJustification = target.GetAttributeValue<EntityReference>("ts_workorderjustification");
                        updateWorkOrder["ts_workorderjustification"] = workOrderJustification;
                        localContext.Trace("ts_workorderjustification updated. New value: {0}", workOrderJustification?.Id.ToString() ?? "null");
                        anyFieldChanged = true;
                    }
                    if (target.Attributes.Contains("ts_state"))
                    {
                        OptionSetValue state = target.GetAttributeValue<OptionSetValue>("ts_state");
                        updateWorkOrder["ts_state"] = state;
                        localContext.Trace("ts_state changed. New value: {0}", state);
                        anyFieldChanged = true;
                    }
                    if (target.Attributes.Contains("ts_worklocation"))
                    {
                        OptionSetValue workLocation = target.GetAttributeValue<OptionSetValue>("ts_worklocation");
                        updateWorkOrder["msdyn_worklocation"] = workLocation;
                        localContext.Trace("msdyn_worklocation changed. New value: {0}", workLocation);
                        anyFieldChanged = true;
                    }
                    if (target.Attributes.Contains("ts_rational"))
                    {
                        EntityReference rational = target.GetAttributeValue<EntityReference>("ts_rational");
                        updateWorkOrder["ts_rational"] = rational;
                        localContext.Trace("ts_rational changed. New value: {0}", rational != null ? rational.Id.ToString() : "null");
                        anyFieldChanged = true;
                    }                    
                    //if (target.Attributes.Contains("ts_businessowner"))
                    //{
                    //    string businessowner = target.GetAttributeValue<string>("ts_businessowner");
                    //    updateWorkOrder["ts_businessowner"] = businessowner;
                    //    localContext.Trace("ts_businessowner changed. New value: {0}", businessowner);
                    //    anyFieldChanged = true;
                    //}
                    if (target.Attributes.Contains("ts_primaryincidenttype"))
                    {
                        EntityReference primaryIncidentType = target.GetAttributeValue<EntityReference>("ts_primaryincidenttype");
                        updateWorkOrder["msdyn_primaryincidenttype"] = primaryIncidentType;
                        localContext.Trace("msdyn_primaryincidenttype changed. New value: {0}", primaryIncidentType != null ? primaryIncidentType.Id.ToString() : "null");
                        anyFieldChanged = true;
                    }
                    if (target.Attributes.Contains("ts_primaryincidentdescription"))
                    {
                        string primaryIncidentDescription = target.GetAttributeValue<string>("ts_primaryincidentdescription");
                        updateWorkOrder["msdyn_primaryincidentdescription"] = primaryIncidentDescription;
                        localContext.Trace("msdyn_primaryincidentdescription changed. New value: {0}", primaryIncidentDescription);
                        anyFieldChanged = true;
                    }
                    if (target.Attributes.Contains("ts_primaryincidentestimatedduration"))
                    {
                        int primaryIncidentEstimatedDuration = target.GetAttributeValue<int>("ts_primaryincidentestimatedduration");
                        updateWorkOrder["msdyn_primaryincidentestimatedduration"] = primaryIncidentEstimatedDuration;
                        localContext.Trace("msdyn_primaryincidentestimatedduration changed. New value: {0}", primaryIncidentEstimatedDuration);
                        anyFieldChanged = true;
                    }
                    if (target.Attributes.Contains("ts_overtimerequired"))
                    {
                        bool overtimeRequired = target.GetAttributeValue<bool>("ts_overtimerequired");
                        updateWorkOrder["ts_overtimerequired"] = overtimeRequired;
                        localContext.Trace("ts_overtimerequired changed. New value: {0}", overtimeRequired);
                        anyFieldChanged = true;
                    }
                    if (target.Attributes.Contains("ownerid"))
                    {
                        EntityReference owner = target.GetAttributeValue<EntityReference>("ownerid");
                        updateWorkOrder["ownerid"] = owner;
                        localContext.Trace("ownerid changed. New value: {0}", owner != null ? owner.Id.ToString() : "null");
                        anyFieldChanged = true;
                    }
                    if (target.Attributes.Contains("ts_country"))
                    {
                        EntityReference country = target.GetAttributeValue<EntityReference>("ts_country");
                        updateWorkOrder["ts_country"] = country;
                        localContext.Trace("ts_country changed. New value: {0}", country != null ? country.Id.ToString() : "null");
                        anyFieldChanged = true;
                    }
                    //start here for 473143
                    if (target.Attributes.Contains("ts_aircraftclassification"))
                    {
                        OptionSetValue aircraftClassification = target.GetAttributeValue<OptionSetValue>("ts_aircraftclassification");
                        updateWorkOrder["ts_aircraftclassification"] = aircraftClassification;
                        localContext.Trace("ts_aircraftclassification changed. New value: {0}", aircraftClassification);
                        anyFieldChanged = true;
                    }
                    if (target.Attributes.Contains("ts_reportdetails"))
                    {
                        string reportDetails = target.GetAttributeValue<string>("ts_reportdetails");
                        updateWorkOrder["ts_reportdetails"] = reportDetails;
                        localContext.Trace("ts_reportdetails changed. New value: {0}", reportDetails);
                        anyFieldChanged = true;
                    }
                    if (target.Attributes.Contains("ts_operation"))
                    {
                        EntityReference operationId = target.GetAttributeValue<EntityReference>("ts_operation");
                        updateWorkOrder["ovs_operationid"] = operationId;
                        localContext.Trace("ovs_operationid changed. New value: {0}", operationId != null ? operationId.Id.ToString() : "null");
                        anyFieldChanged = true;
                    }
                    if (target.Attributes.Contains("ts_plannedfiscalyear"))
                    {
                        EntityReference plannedFiscalYear = target.GetAttributeValue<EntityReference>("ts_plannedfiscalyear");
                        updateWorkOrder["ovs_fiscalyear"] = plannedFiscalYear;
                        localContext.Trace("ts_plannedfiscalyear changed. New value: {0}", plannedFiscalYear != null ? plannedFiscalYear.Id.ToString() : "null");
                        anyFieldChanged = true;
                    }
                    if (target.Attributes.Contains("ts_plannedfiscalquarter"))
                    {
                        EntityReference plannedFiscalQuarter = target.GetAttributeValue<EntityReference>("ts_plannedfiscalquarter");
                        updateWorkOrder["ovs_fiscalquarter"] = plannedFiscalQuarter;
                        localContext.Trace("ts_plannedfiscalquarter changed. New value: {0}", plannedFiscalQuarter != null ? plannedFiscalQuarter.Id.ToString() : "null");
                        anyFieldChanged = true;
                    }
                    if (target.Attributes.Contains("ts_cancelledinspectionjustification"))
                    {
                        EntityReference cancelledInspectionJustification = target.GetAttributeValue<EntityReference>("ts_cancelledinspectionjustification");
                        updateWorkOrder["ts_canceledinspectionjustification"] = cancelledInspectionJustification;
                        localContext.Trace("ts_cancelledinspectionjustification changed. New value: {0}", cancelledInspectionJustification != null ? cancelledInspectionJustification.Id.ToString() : "null");
                        anyFieldChanged = true;
                    }
                    if (target.Attributes.Contains("ts_othercancelledjustification"))
                    {
                        string otherCancelledInspectionJustification = target.GetAttributeValue<string>("ts_othercancelledjustification");
                        updateWorkOrder["ts_othercanceledjustification"] = otherCancelledInspectionJustification;
                        localContext.Trace("ts_othercancelledjustification changed. New value: {0}", otherCancelledInspectionJustification);
                        anyFieldChanged = true;
                    }
                    if (target.Attributes.Contains("ts_revisedquarterid"))
                    {
                        EntityReference revisedQuarterId = target.GetAttributeValue<EntityReference>("ts_revisedquarterid");
                        updateWorkOrder["ovs_revisedquarterid"] = revisedQuarterId;
                        localContext.Trace("ts_revisedquarterid changed. New value: {0}", revisedQuarterId != null ? revisedQuarterId.Id.ToString() : "null");
                        anyFieldChanged = true;
                    }
                    if (target.Attributes.Contains("ts_scheduledquarterjustification"))
                    {
                        EntityReference scheduledQuarterJustification = target.GetAttributeValue<EntityReference>("ts_scheduledquarterjustification");
                        updateWorkOrder["ts_scheduledquarterjustification"] = scheduledQuarterJustification;
                        localContext.Trace("ts_scheduledquarterjustification changed. New value: {0}", scheduledQuarterJustification != null ? scheduledQuarterJustification.Id.ToString() : "null");
                        anyFieldChanged = true;
                    }
                    if (target.Attributes.Contains("ts_scheduledquarterjustificationcomment"))
                    {
                        string justificationComment = target.GetAttributeValue<string>("ts_scheduledquarterjustificationcomment");
                        updateWorkOrder["ts_justificationcomment"] = justificationComment;
                        localContext.Trace("ts_scheduledquarterjustificationcomment changed. New value: {0}", justificationComment);
                        anyFieldChanged = true;
                    }
                    if (target.Attributes.Contains("ts_details"))
                    {
                        string planningComment = target.GetAttributeValue<string>("ts_details");
                        updateWorkOrder["ts_details"] = planningComment;
                        localContext.Trace("ts_details changed. New value: {0}", planningComment);
                        anyFieldChanged = true;
                    }
                    if (target.Attributes.Contains("ts_instructions"))
                    {
                        string instructions = target.GetAttributeValue<string>("ts_instructions");
                        updateWorkOrder["msdyn_instructions"] = instructions;
                        localContext.Trace("ts_instructions changed. New value: {0}", instructions);
                        anyFieldChanged = true;
                    }
                    if (target.Attributes.Contains("ts_wopreparationtime"))
                    {
                        decimal preparationTime = target.GetAttributeValue<decimal>("ts_wopreparationtime");
                        updateWorkOrder["ts_preparationtime"] = preparationTime;
                        localContext.Trace("ts_preparationtime changed. New value: {0}", preparationTime);
                        anyFieldChanged = true;
                    }
                    if (target.Attributes.Contains("ts_woreportinganddocumentation"))
                    {
                        decimal reportingAndDocumentation = target.GetAttributeValue<decimal>("ts_woreportinganddocumentation");
                        updateWorkOrder["ts_woreportinganddocumentation"] = reportingAndDocumentation;
                        localContext.Trace("ts_woreportinganddocumentation changed. New value: {0}", reportingAndDocumentation);
                        anyFieldChanged = true;
                    }
                    if (target.Attributes.Contains("ts_comments"))
                    {
                        string comments = target.GetAttributeValue<string>("ts_comments");
                        updateWorkOrder["ts_comments"] = comments;
                        localContext.Trace("ts_comments changed. New value: {0}", comments);
                        anyFieldChanged = true;
                    }
                    if (target.Attributes.Contains("ts_overtime"))
                    {
                        decimal overTime = target.GetAttributeValue<decimal>("ts_overtime");
                        updateWorkOrder["ts_overtime"] = overTime;
                        localContext.Trace("ts_overtime changed. New value: {0}", overTime);
                        anyFieldChanged = true;
                    }
                    if (target.Attributes.Contains("ts_woconductingoversight"))
                    {
                        decimal conductingOversight = target.GetAttributeValue<decimal>("ts_woconductingoversight");
                        updateWorkOrder["ts_conductingoversight"] = conductingOversight;
                        localContext.Trace("ts_conductingoversight changed. New value: {0}", conductingOversight);
                        anyFieldChanged = true;
                    }
                    if (target.Attributes.Contains("ts_wotraveltime"))
                    {
                        decimal travelTime = target.GetAttributeValue<decimal>("ts_wotraveltime");
                        updateWorkOrder["ts_traveltime"] = travelTime;
                        localContext.Trace("ts_traveltime changed. New value: {0}", travelTime);
                        anyFieldChanged = true;
                    }
                    if (target.Attributes.Contains("ts_servicerequest"))
                    {
                        EntityReference serviceRequest = target.GetAttributeValue<EntityReference>("ts_servicerequest");
                        updateWorkOrder["msdyn_servicerequest"] = serviceRequest;
                        localContext.Trace("msdyn_servicerequest changed. New value: {0}", serviceRequest != null ? serviceRequest.Id.ToString() : "null");
                        anyFieldChanged = true;
                    }
                    if (target.Attributes.Contains("ts_securityincident"))
                    {
                        EntityReference securityIncident = target.GetAttributeValue<EntityReference>("ts_securityincident");
                        updateWorkOrder["ts_securityincident"] = securityIncident;
                        localContext.Trace("ts_securityincident changed. New value: {0}", securityIncident != null ? securityIncident.Id.ToString() : "null");
                        anyFieldChanged = true;
                    }
                    if (target.Attributes.Contains("ts_trip"))
                    {
                        EntityReference trip = target.GetAttributeValue<EntityReference>("ts_trip");
                        updateWorkOrder["ts_trip"] = trip;
                        localContext.Trace("ts_trip changed. New value: {0}", trip != null ? trip.Id.ToString() : "null");
                        anyFieldChanged = true;
                    }
                    if (target.Attributes.Contains("ts_recordstatus"))
                    {
                        OptionSetValue recordStatus = target.GetAttributeValue<OptionSetValue>("ts_recordstatus");
                        updateWorkOrder["msdyn_systemstatus"] = recordStatus;
                        localContext.Trace("msdyn_systemstatus changed. New value: {0}", recordStatus?.Value);
                        anyFieldChanged = true;
                    }
                    if (target.Attributes.Contains("ts_parentworkorder"))
                    {
                        EntityReference parentWorkorder = target.GetAttributeValue<EntityReference>("ts_parentworkorder");
                        updateWorkOrder["msdyn_parentworkorder"] = parentWorkorder;
                        localContext.Trace("msdyn_parentworkorder changed. New value: {0}", parentWorkorder != null ? parentWorkorder.Id.ToString() : "null");
                        anyFieldChanged = true;
                    }
                                        if (target.Attributes.Contains("ts_tradename"))
                    {
                        var tradename = target.GetAttributeValue<EntityReference>("ts_tradename");
                        updateWorkOrder["ts_tradenameid"] = tradename;
                        localContext.Trace("ts_tradenameid changed. New value: {0}", tradename != null ? tradename.Id.ToString() : "null");
                        anyFieldChanged = true;
                    }

                    if (target.Attributes.Contains("ts_functionallocation"))
                    {
                        var funcLoc = target.GetAttributeValue<EntityReference>("ts_functionallocation");
                        updateWorkOrder["msdyn_functionallocation"] = funcLoc;
                        localContext.Trace("msdyn_functionallocation changed. New value: {0}", funcLoc != null ? funcLoc.Id.ToString() : "null");
                        anyFieldChanged = true;
                    }

                    if (target.Attributes.Contains("ts_subsubsite"))
                    {
                        var subSubSite = target.GetAttributeValue<EntityReference>("ts_subsubsite");
                        updateWorkOrder["ts_subsubsite"] = subSubSite;
                        localContext.Trace("ts_subsubsite changed. New value: {0}", subSubSite != null ? subSubSite.Id.ToString() : "null");
                        anyFieldChanged = true;
                    }

                    if (target.Attributes.Contains("ts_contact"))
                    {
                        var contact = target.GetAttributeValue<EntityReference>("ts_contact");
                        updateWorkOrder["ts_contact"] = contact;
                        localContext.Trace("ts_contact changed. New value: {0}", contact != null ? contact.Id.ToString() : "null");
                        anyFieldChanged = true;
                    }
                    if (target.Attributes.Contains("ts_accountableteam"))
                    {
                        var accountableteam = target.GetAttributeValue<EntityReference>("ts_accountableteam");
                        updateWorkOrder["ts_accountableteam"] = accountableteam;
                        localContext.Trace("ts_accountableteam changed. New value: {0}", accountableteam != null ? accountableteam.Id.ToString() : "null");
                        anyFieldChanged = true;
                    }
                    if (anyFieldChanged)
                    {
                        service.Update(updateWorkOrder);
                        localContext.Trace("Updated msdyn_workorder with changed fields.");
                    }
                    else
                    {
                        localContext.Trace("No relevant fields changed. No update performed.");
                    }

                }
                else
                {
                    localContext.Trace("No target entity found. Exiting plugin.");
                }
            }
            catch (Exception ex)
            {
                localContext.Trace("An error occurred: {0}", ex.ToString());
                throw;
            }
        }
    }
}