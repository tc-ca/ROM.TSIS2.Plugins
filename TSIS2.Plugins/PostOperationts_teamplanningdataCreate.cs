using System;
using System.Linq;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;

namespace TSIS2.Plugins
{

    [CrmPluginRegistration(
        MessageNameEnum.Create,
        "ts_teamplanningdata",
        StageEnum.PostOperation,
        ExecutionModeEnum.Synchronous,
        "",
        "TSIS2.Plugins.PostOperationts_teamplanningdataCreate Plugin",
        1,
        IsolationModeEnum.Sandbox,
        Description = "After Team Planning Data is created, automatically create Planning Data")]
    public class PostOperationts_teamplanningdataCreate : IPlugin
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
                    if (target.LogicalName.Equals(ts_TeamPlanningData.EntityLogicalName))
                    {
                        using (var serviceContext = new Xrm(service))
                        {
                            ts_TeamPlanningData targetTeamPlanningData = target.ToEntity<ts_TeamPlanningData>();
                            ts_TeamPlanningData teamPlanningData = serviceContext.ts_TeamPlanningDataSet.FirstOrDefault(tpd => tpd.Id == target.Id);

                            //Retrieve all Operations where OPI Team equals Team Planning Data Team
                            var operations = serviceContext.ovs_operationSet.Where(op => op.ts_OPITeam.Id == teamPlanningData.ts_Team.Id);

                            int teamPlanningDataPlannedQ1 = 0;
                            int teamPlanningDataPlannedQ2 = 0;
                            int teamPlanningDataPlannedQ3 = 0;
                            int teamPlanningDataPlannedQ4 = 0;

                            decimal teamPlanningDataAvailableInspectorHoursQ1 = 0;
                            decimal teamPlanningDataAvailableInspectorHoursQ2 = 0;
                            decimal teamPlanningDataAvailableInspectorHoursQ3 = 0;
                            decimal teamPlanningDataAvailableInspectorHoursQ4 = 0;

                            decimal teamPlanningDataTeamEstimatedDurationQ1 = 0;
                            decimal teamPlanningDataTeamEstimatedDurationQ2 = 0;
                            decimal teamPlanningDataTeamEstimatedDurationQ3 = 0;
                            decimal teamPlanningDataTeamEstimatedDurationQ4 = 0;

                            decimal ts_teamPlanningDataResidualinspectorhoursQ1 = 0;
                            decimal ts_teamPlanningDataResidualinspectorhoursQ2 = 0;
                            decimal ts_teamPlanningDataResidualinspectorhoursQ3 = 0;
                            decimal ts_teamPlanningDataResidualinspectorhoursQ4 = 0;

                            foreach (ovs_operation operation in operations)
                            {
                                var operationActivities = serviceContext.ts_OperationActivitySet.Where(oa => oa.ts_Operation.Id == operation.ovs_operationId);
                                foreach (ts_OperationActivity operationActivity in operationActivities)
                                {
                                    string generationLog = "";
                                    bool isMissingData = false;
                                    string planningDataName = "";
                                    string planningDataEnglishName = "";
                                    string planningDataFrenchName = "";
                                    Guid planningDataFiscalYearId = targetTeamPlanningData.ts_FiscalYear.Id;
                                    Guid planningDataStakeholderId = new Guid();
                                    Guid planningDataOperationTypeId = new Guid();
                                    Guid planningDataSiteId = new Guid();
                                    Guid planningDataActivityTypeId = new Guid();
                                    int planningDataTarget = 0;
                                    int planningDataEstimatedDuration = 0;
                                    int[] planningDataQuarters = new int[4];

                                    if (operationActivity.ts_Activity != null && operation.ts_stakeholder != null && operation.ovs_OperationTypeId != null && operation.ts_site != null)
                                    {
                                        planningDataStakeholderId = operation.ts_stakeholder.Id;
                                        planningDataOperationTypeId = operation.ovs_OperationTypeId.Id;
                                        planningDataSiteId = operation.ts_site.Id;
                                        planningDataActivityTypeId = operationActivity.ts_Activity.Id;
                                        msdyn_incidenttype incidentType = serviceContext.msdyn_incidenttypeSet.FirstOrDefault(it => it.Id == operationActivity.ts_Activity.Id);
                                        msdyn_FunctionalLocation functionalLocation = serviceContext.msdyn_FunctionalLocationSet.FirstOrDefault(fl => fl.Id == operation.ts_site.Id);
                                        if (incidentType != null && incidentType.ts_RiskScore != null && incidentType.msdyn_EstimatedDuration != null)
                                        {
                                            ts_RecurrenceFrequencies recurrenceFrequency = serviceContext.ts_RecurrenceFrequenciesSet.FirstOrDefault(rf => rf.Id == incidentType.ts_RiskScore.Id);

                                            if (incidentType.ovs_IncidentTypeNameEnglish != null && incidentType.ovs_IncidentTypeNameFrench != null)
                                            {
                                                planningDataEnglishName = operation.ovs_name + " | " + incidentType.ovs_IncidentTypeNameEnglish + " | " + teamPlanningData.ts_FiscalYear.Name;
                                                planningDataFrenchName = operation.ovs_name + " | " + incidentType.ovs_IncidentTypeNameFrench + " | " + teamPlanningData.ts_FiscalYear.Name;
                                                planningDataName = planningDataEnglishName + "::" + planningDataFrenchName;

                                                ts_TeamActivityTypeEstimatedDuration teamActivityTypeEstimatedDuration = serviceContext.ts_TeamActivityTypeEstimatedDurationSet.FirstOrDefault(ed => ed.ts_Team.Id == teamPlanningData.ts_Team.Id && ed.ts_ActivityType.Id == incidentType.Id);
                                                if (teamActivityTypeEstimatedDuration != null && teamActivityTypeEstimatedDuration.ts_EstimatedDuration != null)
                                                {
                                                    planningDataEstimatedDuration = ((int)teamActivityTypeEstimatedDuration.ts_EstimatedDuration) / 60;
                                                }
                                                else
                                                {
                                                    planningDataEstimatedDuration = ((int)incidentType.msdyn_EstimatedDuration) / 60;
                                                    generationLog += "Missing Team Estimated Duration for this Team and Activity Type. Using Activity Type Estimated Duration. \n";
                                                }

                                                if (operationActivity.ts_OperationalStatus == ts_operationalstatus.Operational)
                                                {
                                                    int interval = 0;

                                                    if (recurrenceFrequency != null)
                                                    {
                                                        if (functionalLocation.ts_Class == msdyn_FunctionalLocation_ts_Class._1)
                                                        {
                                                            interval = (int)recurrenceFrequency.ts_Class1Interval;
                                                        }
                                                        else //Class 2 or 3
                                                        {
                                                            if (functionalLocation.ts_RiskScore > 5)
                                                            {
                                                                interval = (int)recurrenceFrequency.ts_Class2and3HighRiskInterval;
                                                            }
                                                            else
                                                            {
                                                                interval = (int)recurrenceFrequency.ts_Class2and3LowRiskInterval;
                                                            }
                                                        }
                                                    }
                                                    else
                                                    {
                                                        generationLog += "Could not retrieve Risk Score from the Activity Type of the Operation Activity. \n";
                                                    }

                                                    for (int i = 0; i < 4; i += interval)
                                                    {
                                                        planningDataQuarters[i]++;
                                                        planningDataTarget++;
                                                    }
                                                    teamPlanningDataPlannedQ1 += planningDataQuarters[0];
                                                    teamPlanningDataPlannedQ2 += planningDataQuarters[1];
                                                    teamPlanningDataPlannedQ3 += planningDataQuarters[2];
                                                    teamPlanningDataPlannedQ4 += planningDataQuarters[3];

                                                    teamPlanningDataTeamEstimatedDurationQ1 += planningDataQuarters[0] * planningDataEstimatedDuration;
                                                    teamPlanningDataTeamEstimatedDurationQ2 += planningDataQuarters[1] * planningDataEstimatedDuration;
                                                    teamPlanningDataTeamEstimatedDurationQ3 += planningDataQuarters[2] * planningDataEstimatedDuration;
                                                    teamPlanningDataTeamEstimatedDurationQ4 += planningDataQuarters[3] * planningDataEstimatedDuration;
                                                } 
                                                else
                                                {
                                                    generationLog += "The Operational Status is Non-Operational.";
                                                }
                                                
                                            }
                                            else
                                            {
                                                if (incidentType.ovs_IncidentTypeNameEnglish == null)
                                                {
                                                    generationLog += "There is no English Name value for the Activty Type of the Operation Activity. \n";
                                                }
                                                if (incidentType.ovs_IncidentTypeNameFrench == null)
                                                {
                                                    generationLog += "There is no French Name value for the Activty Type of the Operation Activity. \n";
                                                }
                                                isMissingData = true;
                                            }
                                        } 
                                        else
                                        {
                                            if (incidentType == null)
                                            {
                                                generationLog += "Could not retrieve Incident Type from Operation Activity. \n";
                                            }
                                            if (incidentType.ts_RiskScore == null)
                                            {
                                                generationLog += "The Incident Type does not have a Risk Score. \n";
                                            }
                                            if (incidentType.msdyn_EstimatedDuration == null)
                                            {
                                                generationLog += "The Incident Type does not have an Estimated Duration. \n";
                                            }
                                            isMissingData = true;
                                        }
                                    } 
                                    else
                                    {
                                        if (operationActivity.ts_Activity == null)
                                        {
                                            generationLog += "The Operation Activity is missing an Activity Type. \n";
                                        }
                                        if (operation.ts_stakeholder == null)
                                        {
                                            generationLog += "The Operation of the Operation Activity is missing a Stakeholder. \n";
                                        }
                                        if (operation.ovs_OperationTypeId == null)
                                        {
                                            generationLog += "The Operation of the Operation Activity is missing an Operation Type. \n";
                                        }
                                        if (operation.ts_site == null)
                                        {
                                            generationLog += "The Operation of the Operation Activity is missing a Site. \n";
                                        }
                                        isMissingData = true;
                                    }
                                    service.Create(new ts_PlanningData
                                    {
                                        ts_Name = (isMissingData) ? "ERROR " + planningDataName : planningDataName,
                                        ts_EnglishName = (isMissingData) ? "ERROR " + planningDataEnglishName : planningDataEnglishName,
                                        ts_FrenchName = (isMissingData) ? "ERREUR " + planningDataFrenchName : planningDataFrenchName,
                                        ts_OperationActivity = new EntityReference(ts_OperationActivity.EntityLogicalName, operationActivity.Id),
                                        ts_FiscalYear = new EntityReference(ts_TeamPlanningData.EntityLogicalName, planningDataFiscalYearId),
                                        ts_TeamPlanningData = new EntityReference(ts_TeamPlanningData.EntityLogicalName, teamPlanningData.Id),
                                        ts_Stakeholder = new EntityReference(Account.EntityLogicalName, planningDataStakeholderId),
                                        ts_OperationType = new EntityReference(ovs_operationtype.EntityLogicalName, planningDataOperationTypeId),
                                        ts_Site = new EntityReference(msdyn_FunctionalLocation.EntityLogicalName, planningDataSiteId),
                                        ts_ActivityType = new EntityReference(msdyn_incidenttype.EntityLogicalName, planningDataActivityTypeId),
                                        ts_Target = planningDataTarget,
                                        ts_TeamEstimatedDuration = planningDataEstimatedDuration,
                                        ts_DueQ1 = planningDataQuarters[0],
                                        ts_DueQ2 = planningDataQuarters[1],
                                        ts_DueQ3 = planningDataQuarters[2],
                                        ts_DueQ4 = planningDataQuarters[3],
                                        ts_PlannedQ1 = planningDataQuarters[0],
                                        ts_PlannedQ2 = planningDataQuarters[1],
                                        ts_PlannedQ3 = planningDataQuarters[2],
                                        ts_PlannedQ4 = planningDataQuarters[3],
                                        ts_GenerationLog = generationLog
                                    });
                                }
                            }

                            ts_BaselineHours baselineHours = serviceContext.ts_BaselineHoursSet.FirstOrDefault(blh => blh.ts_Team.Id == teamPlanningData.ts_Team.Id);

                            if (baselineHours != null)
                            {
                                teamPlanningDataAvailableInspectorHoursQ1 = (decimal)baselineHours.ts_PlannedQ1;
                                teamPlanningDataAvailableInspectorHoursQ2 = (decimal)baselineHours.ts_PlannedQ2;
                                teamPlanningDataAvailableInspectorHoursQ3 = (decimal)baselineHours.ts_PlannedQ3;
                                teamPlanningDataAvailableInspectorHoursQ4 = (decimal)baselineHours.ts_PlannedQ4;

                                ts_teamPlanningDataResidualinspectorhoursQ1 = teamPlanningDataAvailableInspectorHoursQ1 - teamPlanningDataTeamEstimatedDurationQ1;
                                ts_teamPlanningDataResidualinspectorhoursQ2 = teamPlanningDataAvailableInspectorHoursQ2 - teamPlanningDataTeamEstimatedDurationQ2;
                                ts_teamPlanningDataResidualinspectorhoursQ3 = teamPlanningDataAvailableInspectorHoursQ3 - teamPlanningDataTeamEstimatedDurationQ3;
                                ts_teamPlanningDataResidualinspectorhoursQ4 = teamPlanningDataAvailableInspectorHoursQ4 - teamPlanningDataTeamEstimatedDurationQ4;
                            }

                            service.Update(new ts_TeamPlanningData
                            {
                                Id = targetTeamPlanningData.Id,
                                ts_Name = teamPlanningData.ts_Team.Name + " | " + teamPlanningData.ts_FiscalYear.Name,
                                ts_PlannedActivityQ1 = teamPlanningDataPlannedQ1,
                                ts_PlannedActivityQ2 = teamPlanningDataPlannedQ2,
                                ts_PlannedActivityQ3 = teamPlanningDataPlannedQ3,
                                ts_PlannedActivityQ4 = teamPlanningDataPlannedQ4,
                                ts_AvailableHoursQ1 = teamPlanningDataAvailableInspectorHoursQ1,
                                ts_AvailableHoursQ2 = teamPlanningDataAvailableInspectorHoursQ2,
                                ts_AvailableHoursQ3 = teamPlanningDataAvailableInspectorHoursQ3,
                                ts_AvailableHoursQ4 = teamPlanningDataAvailableInspectorHoursQ4,
                                ts_TeamestimateddurationQ1 = teamPlanningDataTeamEstimatedDurationQ1,
                                ts_TeamestimateddurationQ2 = teamPlanningDataTeamEstimatedDurationQ2,
                                ts_TeamestimateddurationQ3 = teamPlanningDataTeamEstimatedDurationQ3,
                                ts_TeamestimateddurationQ4 = teamPlanningDataTeamEstimatedDurationQ4,
                                ts_ResidualinspectorhoursQ1 = ts_teamPlanningDataResidualinspectorhoursQ1,
                                ts_ResidualinspectorhoursQ2 = ts_teamPlanningDataResidualinspectorhoursQ2,
                                ts_ResidualinspectorhoursQ3 = ts_teamPlanningDataResidualinspectorhoursQ3,
                                ts_ResidualinspectorhoursQ4 = ts_teamPlanningDataResidualinspectorhoursQ4,
                            });
                        }
                    }
                }
                catch (NotImplementedException ex)
                {
                    // If exceptions from mocking library, just continue. If not from there, we should still throw the error.
                    if (ex.Source == "XrmMockup365" && ex.Message == "No implementation for expression operator 'Count'")
                    {
                        // continue
                    }
                    else
                    {
                        tracingService.Trace("PostOperationts_teamplanningdataCreate Plugin: {0}", ex.ToString());
                        throw;
                    }
                }
                catch (Exception ex)
                {
                    tracingService.Trace("PostOperationts_teamplanningdataCreate Plugin: {0}", ex.ToString());
                    throw;
                }
            }
        }
    }
}