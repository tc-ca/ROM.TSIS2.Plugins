using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TSIS2.Plugins
{
    [CrmPluginRegistration(
       MessageNameEnum.Update,
       "ts_planningdata",
       StageEnum.PostOperation,
       ExecutionModeEnum.Synchronous,
       "",
       "TSIS2.Plugins.PostOperationts_planningdataUpdate Plugin",
       1,
       IsolationModeEnum.Sandbox,
       Description = "After Planning Data updatedd, automatically calculate value for Team Planning Data related record")]
    public class PostOperationts_planningdataUpdate : IPlugin
    {

        public void Execute(IServiceProvider serviceProvider)
        {
            // Obtain the tracing service
            ITracingService tracingService =
            (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)
                serviceProvider.GetService(typeof(IPluginExecutionContext));

            // The InputParameters collection contains all the data passed in the message request.
            if (context.InputParameters.Contains("Target") &&
                context.InputParameters["Target"] is Entity)
            {
                // Obtain the target entity from the input parameters.
                Entity target = (Entity)context.InputParameters["Target"];

                // Obtain the preimage entity
                Entity preImageEntity = (context.PreEntityImages != null && context.PreEntityImages.Contains("PreImage")) ? context.PreEntityImages["PreImage"] : null;

                // Obtain the organization service reference which you will need for
                // web service calls.
                IOrganizationServiceFactory serviceFactory =
                    (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

                try
                {
                    if (target.LogicalName.Equals(ts_PlanningData.EntityLogicalName))
                    {
                        ts_PlanningData planningDataTarget = target.ToEntity<ts_PlanningData>();

                        using (var serviceContext = new Xrm(service))
                        {
                            ts_PlanningData planningData = serviceContext.ts_PlanningDataSet.FirstOrDefault(pd => pd.Id == planningDataTarget.Id);
                            Guid teamPlanningDataId = planningData.ts_TeamPlanningData.Id;
                            ts_TeamPlanningData teamPlanningData = serviceContext.ts_TeamPlanningDataSet.FirstOrDefault(tpd => tpd.Id == teamPlanningDataId);
                            var planningDataList = serviceContext.ts_PlanningDataSet.Where(pd => pd.ts_TeamPlanningData.Id == teamPlanningDataId);

                            ts_BaselineHours baselineHours = serviceContext.ts_BaselineHoursSet.FirstOrDefault(blh => blh.ts_Team.Id == teamPlanningData.ts_Team.Id);

                            int plannedQ1 = 0;
                            int plannedQ2 = 0;
                            int plannedQ3 = 0;
                            int plannedQ4 = 0;

                            decimal teamEstimatedDurationQ1 = 0;
                            decimal teamEstimatedDurationQ2 = 0;
                            decimal teamEstimatedDurationQ3 = 0;
                            decimal teamEstimatedDurationQ4 = 0;

                            foreach (var pd in planningDataList)
                            {
                                plannedQ1 += (int)pd.ts_PlannedQ1;
                                plannedQ2 += (int)pd.ts_PlannedQ2;
                                plannedQ3 += (int)pd.ts_PlannedQ3;
                                plannedQ4 += (int)pd.ts_PlannedQ4;
                                teamEstimatedDurationQ1 += (int)pd.ts_PlannedQ1 * (decimal)pd.ts_TeamEstimatedDuration;
                                teamEstimatedDurationQ2 += (int)pd.ts_PlannedQ2 * (decimal)pd.ts_TeamEstimatedDuration;
                                teamEstimatedDurationQ3 += (int)pd.ts_PlannedQ3 * (decimal)pd.ts_TeamEstimatedDuration;
                                teamEstimatedDurationQ4 += (int)pd.ts_PlannedQ4 * (decimal)pd.ts_TeamEstimatedDuration;
                            }

                            if (baselineHours != null)
                            {

                                service.Update(new ts_TeamPlanningData
                                {
                                    Id = teamPlanningData.Id,
                                    ts_PlannedActivityQ1 = plannedQ1,
                                    ts_PlannedActivityQ2 = plannedQ2,
                                    ts_PlannedActivityQ3 = plannedQ3,
                                    ts_PlannedActivityQ4 = plannedQ4,
                                    ts_TeamestimateddurationQ1 = teamEstimatedDurationQ1,
                                    ts_TeamestimateddurationQ2 = teamEstimatedDurationQ2,
                                    ts_TeamestimateddurationQ3 = teamEstimatedDurationQ3,
                                    ts_TeamestimateddurationQ4 = teamEstimatedDurationQ4,
                                    ts_ResidualinspectorhoursQ1 = baselineHours.ts_PlannedQ1 - teamEstimatedDurationQ1,
                                    ts_ResidualinspectorhoursQ2 = baselineHours.ts_PlannedQ2 - teamEstimatedDurationQ2,
                                    ts_ResidualinspectorhoursQ3 = baselineHours.ts_PlannedQ3 - teamEstimatedDurationQ3,
                                    ts_ResidualinspectorhoursQ4 = baselineHours.ts_PlannedQ4 - teamEstimatedDurationQ4,

                                });
                            }
                            else
                            {
                                service.Update(new ts_TeamPlanningData
                                {
                                    Id = teamPlanningData.Id,
                                    ts_PlannedActivityQ1 = plannedQ1,
                                    ts_PlannedActivityQ2 = plannedQ2,
                                    ts_PlannedActivityQ3 = plannedQ3,
                                    ts_PlannedActivityQ4 = plannedQ4,
                                    ts_TeamestimateddurationQ1 = teamEstimatedDurationQ1,
                                    ts_TeamestimateddurationQ2 = teamEstimatedDurationQ2,
                                    ts_TeamestimateddurationQ3 = teamEstimatedDurationQ3,
                                    ts_TeamestimateddurationQ4 = teamEstimatedDurationQ4,
                                });
                            }
                        }
                    }
                }
                catch (Exception e)
                {
                    throw new InvalidPluginExecutionException(e.Message);
                }
            }
        }
    }
}
