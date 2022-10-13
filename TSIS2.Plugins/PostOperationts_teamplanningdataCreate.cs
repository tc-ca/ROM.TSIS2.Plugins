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
                            var operations = serviceContext.ovs_operationSet.Where(op => op.OwnerId == teamPlanningData.OwnerId);

                            foreach (ovs_operation operation in operations)
                            {
                                var operationActivities = serviceContext.ts_OperationActivitySet.Where(oa => oa.ts_Operation.Id == operation.ovs_operationId);
                                foreach (ts_OperationActivity operationActivity in operationActivities)
                                {
                                    //TODO null check

                                    msdyn_incidenttype incidentType = serviceContext.msdyn_incidenttypeSet.FirstOrDefault(it => it.Id == operationActivity.ts_Activity.Id);
                                    ts_RecurrenceFrequencies recurrenceFrequency = serviceContext.ts_RecurrenceFrequenciesSet.FirstOrDefault(rf => rf.Id == incidentType.ts_RiskScore.Id);
                                    msdyn_FunctionalLocation functionalLocation = serviceContext.msdyn_FunctionalLocationSet.FirstOrDefault(fl => fl.Id == operation.ts_site.Id);
                                    tc_TCFiscalYear fiscalYear = serviceContext.tc_TCFiscalYearSet.FirstOrDefault(fy => fy.Id == targetTeamPlanngingData.ts_FiscalYear.Id);

                                    string englishName = operation.ovs_name + " | " + incidentType.ovs_IncidentTypeNameEnglish + " | " + teamPlanningData.ts_FiscalYear.Name;
                                    string frenchName = operation.ovs_name + " | " + incidentType.ovs_IncidentTypeNameFrench + " | " + teamPlanningData.ts_FiscalYear.Name;

                                    var inspections = serviceContext.msdyn_workorderSet.Where(wo => wo.msdyn_PrimaryIncidentType.Id == incidentType.Id && wo.ovs_OperationId.Id == operation.Id);
                                    tc_TCFiscalQuarter latestFiscalQuarter = GetLatestWorkOrderFiscalQuarter(inspections, serviceContext);

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
                                    int planningDataTarget = 0;
                                    int planningDataDueQ1 = 0;
                                    int planningDataDueQ2 = 0;
                                    int planningDataDueQ3 = 0;
                                    int planningDataDueQ4 = 0;

                                    tc_TCFiscalQuarter nextExpectedInspectionQuarter = JumpQuarters(latestFiscalQuarter, interval, serviceContext);

                                    //If the next expected inspection occurs before the current fiscal year, it's overdue and needs to occur Q1
                                    if (((DateTime)nextExpectedInspectionQuarter.tc_QuarterStart).AddDays(1) <= fiscalYear.tc_FiscalStart)
                                    {
                                        planningDataTarget++;
                                        planningDataDueQ1++;
                                    }

                                    service.Create(new ts_PlanningData
                                    {
                                        ts_Name = englishName + "::" + frenchName,
                                        ts_EnglishName = englishName,
                                        ts_FrenchName = frenchName,
                                        ts_FiscalYear = teamPlanningData.ts_FiscalYear,
                                        ts_TeamPlanningData = new EntityReference(ts_TeamPlanningData.EntityLogicalName, teamPlanningData.Id),

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
            DateTime jumpDate = startingDate.AddMonths(numberOfQuarters * 3).AddDays(1);
            return serviceContext.tc_TCFiscalQuarterSet.FirstOrDefault(fq => fq.tc_QuarterStart <= jumpDate && fq.tc_QuarterEnd >= jumpDate);
        }
    }
}