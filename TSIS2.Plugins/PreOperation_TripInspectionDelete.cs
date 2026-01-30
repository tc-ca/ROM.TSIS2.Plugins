using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

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
    public class PreOperation_TripInspectionDelete : PluginBase
    {
        public PreOperation_TripInspectionDelete(string unsecure, string secure)
            : base(typeof(PreOperation_TripInspectionDelete))
        {
        }

        protected override void ExecuteCrmPlugin(LocalPluginContext localContext)
        {
            if (localContext == null)
            {
                throw new InvalidPluginExecutionException("localContext");
            }

            IPluginExecutionContext context = localContext.PluginExecutionContext;
            IOrganizationService service = localContext.OrganizationService;

            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is EntityReference)
            {
                EntityReference entityRef = (EntityReference)context.InputParameters["Target"];
                if (!string.Equals(context.MessageName, "Delete", StringComparison.OrdinalIgnoreCase) ||
                    !string.Equals(entityRef.LogicalName, "ts_tripinspection", StringComparison.OrdinalIgnoreCase))
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
                            localContext.Trace("Remove trip from WO: " + woEnt.GetAttributeValue<EntityReference>("ts_trip").Id.ToString());
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
                catch (Exception e)
                {
                    localContext.Trace("TripInspectionDelete Plugin: {0}", e);
                    throw new InvalidPluginExecutionException($"An error occurred in the TripInspectionDelete Plugin: {e}");
                }
            }
        }
    }

}
