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
        Description = "Copies ts_workorderservicetaskstartdate to the related msdyn_workorderservicetask record on update if it changed.")]
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

                    // Ensure the field is in the Target to confirm it was updated
                    if (!target.Attributes.Contains("ts_workorderservicetaskstartdate"))
                    {
                        tracingService.Trace("ts_workorderservicetaskstartdate was not updated. Exiting plugin.");
                        return;
                    }

                    // Get Pre-Image
                    if (!context.PreEntityImages.Contains("PreImage") || context.PreEntityImages["PreImage"] == null)
                    {
                        tracingService.Trace("PreImage is not available. Exiting plugin.");
                        return;
                    }

                    Entity preImage = context.PreEntityImages["PreImage"];

                    DateTime? oldStartDate = preImage.GetAttributeValue<DateTime?>("ts_workorderservicetaskstartdate");
                    DateTime? newStartDate = target.GetAttributeValue<DateTime?>("ts_workorderservicetaskstartdate");

                    tracingService.Trace("Old Start Date: {0}", oldStartDate.HasValue ? oldStartDate.Value.ToString("o") : "null");
                    tracingService.Trace("New Start Date: {0}", newStartDate.HasValue ? newStartDate.Value.ToString("o") : "null");

                    if (newStartDate == null || newStartDate == oldStartDate)
                    {
                        tracingService.Trace("Start date did not change or is null. Exiting plugin.");
                        return;
                    }

                    EntityReference workOrderTaskRef = preImage.GetAttributeValue<EntityReference>("ts_workorderservicetask")
                        ?? target.GetAttributeValue<EntityReference>("ts_workorderservicetask");

                    if (workOrderTaskRef == null)
                    {
                        tracingService.Trace("ts_workorderservicetask reference is null. Exiting plugin.");
                        return;
                    }

                    tracingService.Trace("Updating msdyn_workorderservicetask Id: {0}", workOrderTaskRef.Id);

                    Entity updateTask = new Entity(workOrderTaskRef.LogicalName, workOrderTaskRef.Id);
                    updateTask["ts_servicetaskstartdate"] = newStartDate.Value;
                    updateTask["statuscode"] = new OptionSetValue(918640004); // In Progress
                    updateTask["msdyn_percentcomplete"] = 50.0; // 50%

                    service.Update(updateTask);

                    tracingService.Trace("Updated msdyn_workorderservicetask with new start date, status reason to In Progress, and % Complete to 50%.");
                }
                else
                {
                    tracingService.Trace("No target entity found. Exiting plugin.");
                }
            }
            catch (Exception ex)
            {
                tracingService.Trace("Exception occurred: {0}", ex.ToString());
                throw new InvalidPluginExecutionException("An error occurred in the PostOperation_CopyStartDateToTaskOnUpdate plugin.", ex);
            }
        }
    }
}