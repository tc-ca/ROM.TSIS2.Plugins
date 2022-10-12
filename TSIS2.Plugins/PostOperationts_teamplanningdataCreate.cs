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
                                    msdyn_incidenttype incidentType = serviceContext.msdyn_incidenttypeSet.FirstOrDefault(it => it.Id == operationActivity.ts_Activity.Id);

                                    string englishName = operation.ovs_name + " | " + incidentType.ovs_IncidentTypeNameEnglish + " | " + teamPlanningData.ts_FiscalYear.Name;
                                    string frenchName = operation.ovs_name + " | " + incidentType.ovs_IncidentTypeNameFrench + " | " + teamPlanningData.ts_FiscalYear.Name;
                                    

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
    }
}