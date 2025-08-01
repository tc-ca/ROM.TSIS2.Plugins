using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;

namespace TSIS2.Plugins
{
    [CrmPluginRegistration(
        MessageNameEnum.Update,
        "ts_workorderservicetaskworkspace",
        StageEnum.PostOperation,
        ExecutionModeEnum.Synchronous,
        "",
        "PostOperation.ts_workorderservicetaskworkspace.CopyStartDateToWorkOrderServiceTaskOnUpdate",
        1,
        IsolationModeEnum.Sandbox,
        Image1Name = "PreImage", Image1Type = ImageTypeEnum.PreImage, Image1Attributes = "ts_workorderservicetask,ts_workorderservicetaskstartdate,ts_questionnaireresponse,ts_mandatory,ts_percentcomplete,ts_aocoperation,ts_aocstakeholder,ts_aocoperationtype,ts_aocsite,statuscode",
        Description = "Copies changed fields to the related msdyn_workorderservicetask record on update.")]
    public class PostOperation_CopyStartDateToTaskOnUpdate : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            var context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            var serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            var service = serviceFactory.CreateOrganizationService(context.UserId);
            var tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            tracingService.Trace("Plugin execution started: PostOperation_CopyStartDateToTaskOnUpdate");

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
                    EntityReference workOrderTaskRef = preImage.GetAttributeValue<EntityReference>("ts_workorderservicetask")
                        ?? target.GetAttributeValue<EntityReference>("ts_workorderservicetask");

                    if (workOrderTaskRef == null)
                    {
                        tracingService.Trace("ts_workorderservicetask reference is null. Exiting plugin.");
                        return;
                    }

                    tracingService.Trace("Updating msdyn_workorderservicetask Id: {0}", workOrderTaskRef.Id);

                    Entity updateTask = new Entity(workOrderTaskRef.LogicalName, workOrderTaskRef.Id);
                    bool anyFieldChanged = false;

                    // Check and update each field if present in Target
                    if (target.Attributes.Contains("ts_workorderservicetaskstartdate"))
                    {
                        DateTime? newStartDate = target.GetAttributeValue<DateTime?>("ts_workorderservicetaskstartdate");
                        updateTask["ts_servicetaskstartdate"] = newStartDate;
                        tracingService.Trace("ts_workorderservicetaskstartdate changed. New value: {0}", newStartDate);
                        anyFieldChanged = true;
                    }
                    if (target.Attributes.Contains("ts_questionnaireresponse"))
                    {
                        string questionnaireResponseText = target.GetAttributeValue<string>("ts_questionnaireresponse");
                        updateTask["ovs_questionnaireresponse"] = questionnaireResponseText;
                        tracingService.Trace("ts_questionnaireresponse changed. New value: {0}", questionnaireResponseText);
                        anyFieldChanged = true;
                    }
                    if (target.Attributes.Contains("ts_mandatory"))
                    {
                        bool isMandatory = target.GetAttributeValue<bool>("ts_mandatory");
                        updateTask["ts_mandatory"] = isMandatory;
                        tracingService.Trace("ts_mandatory changed. New value: {0}", isMandatory);
                        anyFieldChanged = true;
                    }
                    if (target.Attributes.Contains("ts_percentcomplete"))
                    {
                        double? percentComplete = target.GetAttributeValue<double?>("ts_percentcomplete");
                        updateTask["msdyn_percentcomplete"] = percentComplete;
                        tracingService.Trace("msdyn_percentcomplete changed. New value: {0}", percentComplete);
                        anyFieldChanged = true;
                    }
                    if (target.Attributes.Contains("ts_aocoperation"))
                    {
                        EntityReference aocoperationRef = target.GetAttributeValue<EntityReference>("ts_aocoperation");
                        updateTask["ts_aocoperation"] = aocoperationRef;
                        tracingService.Trace("ts_aocoperation changed. New value: {0}", aocoperationRef != null ? aocoperationRef.Id.ToString() : "null");
                        anyFieldChanged = true;
                    }
                    if (target.Attributes.Contains("ts_aocstakeholder"))
                    {
                        EntityReference aocstakeholderRef = target.GetAttributeValue<EntityReference>("ts_aocstakeholder");
                        updateTask["ts_aocstakeholder"] = aocstakeholderRef;
                        tracingService.Trace("ts_aocstakeholder changed. New value: {0}", aocstakeholderRef != null ? aocstakeholderRef.Id.ToString() : "null");
                        anyFieldChanged = true;
                    }
                    if (target.Attributes.Contains("ts_aocoperationtype"))
                    {
                        EntityReference aocoperationtypeRef = target.GetAttributeValue<EntityReference>("ts_aocoperationtype");
                        updateTask["ts_aocoperationtype"] = aocoperationtypeRef;
                        tracingService.Trace("ts_aocoperationtype changed. New value: {0}", aocoperationtypeRef != null ? aocoperationtypeRef.Id.ToString() : "null");
                        anyFieldChanged = true;
                    }
                    if (target.Attributes.Contains("ts_aocsite"))
                    {
                        EntityReference aocsiteRef = target.GetAttributeValue<EntityReference>("ts_aocsite");
                        updateTask["ts_aocsite"] = aocsiteRef;
                        tracingService.Trace("ts_aocsite changed. New value: {0}", aocsiteRef != null ? aocsiteRef.Id.ToString() : "null");
                        anyFieldChanged = true;
                    }
                    if (target.Attributes.Contains("ts_accesscontrol"))
                    {
                        bool accessControl = target.GetAttributeValue<bool>("ts_accesscontrol");
                        updateTask["ts_accesscontrol"] = accessControl;
                        tracingService.Trace("ts_mandatory changed. New value: {0}", accessControl);
                        anyFieldChanged = true;
                    }
                    if( target.Attributes.Contains("ts_workorderservicetaskenddate"))
                    {
                        DateTime? endDate = target.GetAttributeValue<DateTime?>("ts_workorderservicetaskenddate");
                        updateTask["ts_servicetaskenddate"] = endDate;
                        tracingService.Trace("ts_workorderservicetaskenddate changed. New value: {0}", endDate);
                        anyFieldChanged = true;
                    }
                    if (target.Attributes.Contains("statuscode"))
                    {
                        OptionSetValue statusCode = target.GetAttributeValue<OptionSetValue>("statuscode");
                        int? mappedStatusCode = null;
                        if (statusCode != null)
                        {
                            switch (statusCode.Value)
                            {
                                case 1: // Active
                                    mappedStatusCode = 1;
                                    break;
                                case 741130001: // Complete
                                    mappedStatusCode = 918640002;
                                    break;
                                case 741130002: // In Progress
                                    mappedStatusCode = 918640004;
                                    break;
                                case 741130003: // New
                                    mappedStatusCode = 918640005;
                                    break;
                                default:
                                    mappedStatusCode = 1;
                                    break;
                            }
                        }
                        if (mappedStatusCode.HasValue)
                        {
                            updateTask["statuscode"] = new OptionSetValue(mappedStatusCode.Value);
                            tracingService.Trace("statuscode changed. New mapped value: {0}", mappedStatusCode.Value);
                            anyFieldChanged = true;
                        }
                    }

                    if (anyFieldChanged)
                    {
                        service.Update(updateTask);
                        tracingService.Trace("Updated msdyn_workorderservicetask with changed fields.");
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