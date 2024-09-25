using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Security;

namespace TSIS2.Plugins
{
    [CrmPluginRegistration(
       MessageNameEnum.Delete,
       "ts_tripinspection",
       StageEnum.PreOperation,
       ExecutionModeEnum.Synchronous,
       "",
       "TSIS2.Plugins.PreOperation_TripInspectionDelete Plugin",
       1,
       IsolationModeEnum.Sandbox,       
       Description = "Happens before the Trip Inspection has been deleted")]
    public class PreOperation_TripInspectionDelete : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            ITracingService tracingService =
            (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is EntityReference)
            {
                EntityReference entityRef = (EntityReference)context.InputParameters["Target"];
                if (context.MessageName.ToLower() != "delete" || entityRef.LogicalName.ToLower() != "ts_tripinspection")
                {
                    return;
                }
                try
                {
                    Entity inspectEnt = service.Retrieve("ts_tripinspection", entityRef.Id, new ColumnSet("ts_inspection", "ts_trip"));
                    if (inspectEnt.Contains("ts_inspection") && inspectEnt.Contains("ts_trip"))
                    {
                        Entity woEnt = service.Retrieve("msdyn_workorder", inspectEnt.GetAttributeValue<EntityReference>("ts_inspection").Id, new ColumnSet("ts_trip"));
                        if (woEnt.Contains("ts_trip"))
                        {
                            tracingService.Trace("Remove trip from WO: " + woEnt.GetAttributeValue<EntityReference>("ts_trip").Id.ToString());
                            var tripId = woEnt.GetAttributeValue<EntityReference>("ts_trip").Id;
                            if (tripId == inspectEnt.GetAttributeValue<EntityReference>("ts_trip").Id)
                            {
                                Entity updWO = new Entity("msdyn_workorder", woEnt.Id);
                                updWO["ts_trip"] = null;
                                updWO["ts_ignoreupdate"] = true;
                                service.Update(updWO);
                            }

                        }
                    }

                    //service.Delete(entityRef.LogicalName, entityRef.Id);

                }
                catch (Exception ex)
                {
                    throw new InvalidPluginExecutionException($"An error occurred in the TripInspectionDelete Plugin: {ex.Message}");
                }
            }
        }
    }

}
