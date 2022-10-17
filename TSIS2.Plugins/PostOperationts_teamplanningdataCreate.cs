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
                            ts_TeamPlanningData targetTeamPlanngingData = target.ToEntity<ts_TeamPlanningData>();
                            ts_TeamPlanningData teamPlanningData = serviceContext.ts_TeamPlanningDataSet.FirstOrDefault(tpd => tpd.Id == target.Id);

                            //Retrieve all Operations owned by the same Owner
                            var operations = serviceContext.ovs_operationSet.Where(op => op.OwnerId.Id == teamPlanningData.ts_Team.Id);

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
                                    tc_TCFiscalYear planningDataFiscalYear = null;
                                    int planningDataTarget = 0;
                                    int[] planningDataQuarters = new int[4];

                                    if (operationActivity.ts_Activity != null && operation.ts_site != null)
                                    {
                                        msdyn_incidenttype incidentType = serviceContext.msdyn_incidenttypeSet.FirstOrDefault(it => it.Id == operationActivity.ts_Activity.Id);
                                        msdyn_FunctionalLocation functionalLocation = serviceContext.msdyn_FunctionalLocationSet.FirstOrDefault(fl => fl.Id == operation.ts_site.Id);
                                        if (incidentType != null && incidentType.ts_RiskScore != null)
                                        {
                                            ts_RecurrenceFrequencies recurrenceFrequency = serviceContext.ts_RecurrenceFrequenciesSet.FirstOrDefault(rf => rf.Id == incidentType.ts_RiskScore.Id);
                                            planningDataFiscalYear = serviceContext.tc_TCFiscalYearSet.FirstOrDefault(fy => fy.Id == targetTeamPlanngingData.ts_FiscalYear.Id);
                                            if (incidentType.ovs_IncidentTypeNameEnglish != null && incidentType.ovs_IncidentTypeNameFrench != null)
                                            {
                                                planningDataEnglishName = operation.ovs_name + " | " + incidentType.ovs_IncidentTypeNameEnglish + " | " + teamPlanningData.ts_FiscalYear.Name;
                                                planningDataFrenchName = operation.ovs_name + " | " + incidentType.ovs_IncidentTypeNameFrench + " | " + teamPlanningData.ts_FiscalYear.Name;
                                                planningDataName = planningDataEnglishName + "::" + planningDataFrenchName;

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
                                                    generationLog += "Could not retrieve Risk Score from the Activity Type of the Operation Activity \n";
                                                }

                                                for (int i = 0; i < 4; i += interval)
                                                {
                                                    planningDataQuarters[i]++;
                                                    planningDataTarget++;
                                                }
                                            }
                                            else
                                            {
                                                if (incidentType.ovs_IncidentTypeNameEnglish == null)
                                                {
                                                    generationLog += "There is no English Name value for the Activty Type of the Operation Activity \n";
                                                }
                                                if (incidentType.ovs_IncidentTypeNameFrench != null)
                                                {
                                                    generationLog += "There is no French Name value for the Activty Type of the Operation Activity \n";
                                                }
                                                isMissingData = true;
                                            }
                                        } 
                                        else
                                        {
                                            if (incidentType == null)
                                            {
                                                generationLog += "Could not retrieve Incident Type from Operation Type \n";
                                            }
                                            if (incidentType.ts_RiskScore == null)
                                            {
                                                generationLog += "The Incident Type does not have a Risk Score \n";
                                            }
                                            isMissingData = true;
                                        }
                                    } 
                                    else
                                    {
                                        if (operationActivity.ts_Activity == null)
                                        {
                                            generationLog += "The Operation Activity is missing an Activity Type \n";
                                        }
                                        if (operation.ts_site == null)
                                        {
                                            generationLog += "The Operation of the Operation Activity is missing a Site \n";
                                        }
                                        isMissingData = true;
                                    }
                                    service.Create(new ts_PlanningData
                                    {
                                        ts_Name = (isMissingData) ? "ERROR " + planningDataName : planningDataName,
                                        ts_EnglishName = planningDataEnglishName,
                                        ts_FrenchName = planningDataFrenchName,
                                        ts_FiscalYear = new EntityReference(ts_TeamPlanningData.EntityLogicalName, planningDataFiscalYear.Id),
                                        ts_TeamPlanningData = new EntityReference(ts_TeamPlanningData.EntityLogicalName, teamPlanningData.Id),
                                        ts_Target = planningDataTarget,
                                        ts_DueQ1 = planningDataQuarters[0],
                                        ts_DueQ2 = planningDataQuarters[1],
                                        ts_DueQ3 = planningDataQuarters[2],
                                        ts_DueQ4 = planningDataQuarters[3],
                                        ts_GenerationLog = generationLog
                                    });
                                }
                            }
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
        //Returns to most recent fiscal quarter of a list of Work Orders
        private tc_TCFiscalQuarter GetLatestWorkOrderFiscalQuarter(IQueryable<msdyn_workorder> workOrders, Xrm serviceContext)
        {
            msdyn_workorder latestWorkOrder = null;
            tc_TCFiscalQuarter latestWorkOrderFiscalQuarter = null;
            foreach(msdyn_workorder workOrder in workOrders)
            {
                if (latestWorkOrder == null)
                {
                    latestWorkOrder = workOrder;
                    latestWorkOrderFiscalQuarter = serviceContext.tc_TCFiscalQuarterSet.FirstOrDefault(fc => fc.Id == workOrder.ovs_FiscalQuarter.Id);
                } 
                else 
                {
                    var currentWorkOrderFiscalQuarter = serviceContext.tc_TCFiscalQuarterSet.FirstOrDefault(fc => fc.Id == workOrder.ovs_FiscalQuarter.Id);
                    if (currentWorkOrderFiscalQuarter.tc_QuarterEnd < latestWorkOrderFiscalQuarter.tc_QuarterEnd)
                    {
                        latestWorkOrder = workOrder;
                        latestWorkOrderFiscalQuarter = currentWorkOrderFiscalQuarter;
                    }
                }
            }
            return latestWorkOrderFiscalQuarter;
        }

        //Return the Fiscal Quarter numberOfQuarters after the startingQuarter
        private tc_TCFiscalQuarter JumpQuarters(tc_TCFiscalQuarter startingQuarter, int numberOfQuarters, Xrm serviceContext)
        {
            DateTime startingDate = (DateTime)startingQuarter.tc_QuarterEnd;
            DateTime jumpDate = startingDate.AddMonths(numberOfQuarters * 3).AddDays(-1);
            return serviceContext.tc_TCFiscalQuarterSet.FirstOrDefault(fq => fq.tc_QuarterStart <= jumpDate && fq.tc_QuarterEnd >= jumpDate);
        }
    }
}