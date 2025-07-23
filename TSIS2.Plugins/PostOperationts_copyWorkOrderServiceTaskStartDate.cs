using System;
using Microsoft.Xrm.Sdk;

namespace TSIS2.Plugins
{
    [CrmPluginRegistration(
        MessageNameEnum.Create,
        "ts_workorderservicetaskworkspace",
        StageEnum.PostOperation,
        ExecutionModeEnum.Synchronous,
        "",
        "PostOperation.ts_workorderservicetaskworkspace.CopyStartDateToWorkOrderServiceTask",
        1,
        IsolationModeEnum.Sandbox,
        Description = "Copies ts_workorderservicetaskstartdate to the related msdyn_workorderservicetask record after creation.")]
    public class PostOperation_CopyStartDateToTask : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            // Services
            var context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            var serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            var service = serviceFactory.CreateOrganizationService(context.UserId);
            var tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            tracingService.Trace("Plugin execution started: PostOperation_CopyStartDateToTask");

            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity target)
                {
                    tracingService.Trace("Target entity logical name: {0}", target.LogicalName);

                    // Retrieve the start date
                    DateTime? startDate = target.GetAttributeValue<DateTime?>("ts_workorderservicetaskstartdate");
                    if (startDate == null)
                    {
                        tracingService.Trace("Start date is null. Plugin exiting.");
                        return;
                    }

                    // Retrieve the Work Order Service Task reference
                    EntityReference workOrderTaskRef = target.GetAttributeValue<EntityReference>("ts_workorderservicetask");
                    if (workOrderTaskRef == null)
                    {
                        tracingService.Trace("ts_workorderservicetask reference is null. Plugin exiting.");
                        return;
                    }
                    // Retrieve the questionnaire response as text
                    string questionnaireResponseText = target.GetAttributeValue<string>("ts_questionnaireresponse");
                    bool isMandatory = target.GetAttributeValue<bool>("ts_mandatory");
                    bool accessControl = target.GetAttributeValue<bool>("ts_accesscontrol");
                    double? percentComplete = target.GetAttributeValue<double?>("ts_percentcomplete");
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
                                // Optionally handle unknown status codes
                                mappedStatusCode = 1; // Default to Active or handle as needed
                                break;
                        }
                    }
                    EntityReference aocoperationRef = target.GetAttributeValue<EntityReference>("ts_aocoperation");
                    EntityReference aocstakeholderRef = target.GetAttributeValue<EntityReference>("ts_aocstakeholder");
                    EntityReference aocoperationtypeRef = target.GetAttributeValue<EntityReference>("ts_aocoperationtype");
                    EntityReference aocsiteRef = target.GetAttributeValue<EntityReference>("ts_aocsite");

                    tracingService.Trace("Start Date: {0}", startDate.Value);
                    tracingService.Trace("Updating msdyn_workorderservicetask Id: {0}", workOrderTaskRef.Id);

                    // Update the start date, status reason, and percent complete in a single update for efficiency
                    Entity updateTask = new Entity(workOrderTaskRef.LogicalName, workOrderTaskRef.Id);
                    updateTask["ts_servicetaskstartdate"] = startDate.Value;
                    updateTask["statuscode"] = new OptionSetValue(mappedStatusCode.Value);
                    updateTask["msdyn_percentcomplete"] = percentComplete;
                    updateTask["ts_mandatory"] = isMandatory;
                    updateTask["ts_accesscontrol"] = accessControl;
                    updateTask["ovs_questionnaireresponse"] = questionnaireResponseText;
                    updateTask["ts_aocoperation"] = aocoperationRef;
                    updateTask["ts_aocstakeholder"] = aocstakeholderRef;
                    updateTask["ts_aocoperationtype"] = aocoperationtypeRef;
                    updateTask["ts_aocsite"] = aocsiteRef;
                    service.Update(updateTask);

                    tracingService.Trace("Updated start date, status reason to In Progress, and % Complete to 50%.");
                }
                else
                {
                    tracingService.Trace("No target entity found. Plugin exiting.");
                }
            }
            catch (Exception ex)
            {
                tracingService.Trace("Exception occurred: {0}", ex.ToString());
                throw new InvalidPluginExecutionException("An error occurred in the PostOperation_CopyStartDateToTask plugin.", ex);
            }
        }
    }
}