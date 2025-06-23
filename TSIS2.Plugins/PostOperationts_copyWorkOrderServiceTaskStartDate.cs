using System;
using System.Linq;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;

namespace TSIS2.Plugins
{
    [CrmPluginRegistration(
        MessageNameEnum.Create,
        "ts_wostsupplementaryrecord",
        StageEnum.PostOperation,
        ExecutionModeEnum.Synchronous,
        "",
        "PostOperation.ts_wostsupplementaryrecord.CopyStartDateToTask",
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
                // Ensure we have a target entity
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity target)
                {
                    tracingService.Trace("Target entity logical name: {0}", target.LogicalName);

                    // Ensure start date is present
                    DateTime? startDate = target.GetAttributeValue<DateTime?>("ts_workorderservicetaskstartdate");
                    if (startDate == null)
                    {
                        tracingService.Trace("Start date is null. Plugin exiting.");
                        return;
                    }

                    // Ensure the work order service task lookup is present
                    EntityReference workOrderTaskRef = target.GetAttributeValue<EntityReference>("ts_workorderservicetask");
                    if (workOrderTaskRef == null)
                    {
                        tracingService.Trace("ts_workorderservicetask reference is null. Plugin exiting.");
                        return;
                    }

                    tracingService.Trace("Start Date: {0}", startDate.Value);
                    tracingService.Trace("Updating msdyn_workorderservicetask Id: {0}", workOrderTaskRef.Id);

                    // Perform the update
                    Entity updateTask = new Entity(workOrderTaskRef.LogicalName, workOrderTaskRef.Id);
                    updateTask["ts_servicetaskstartdate"] = startDate.Value;

                    service.Update(updateTask);
                    tracingService.Trace("Update complete.");
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