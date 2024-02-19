using System;
using System.Linq;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;

namespace TSIS2.Plugins
{

    [CrmPluginRegistration(
        MessageNameEnum.Create,
        "ts_operationriskassessment",
        StageEnum.PostOperation,
        ExecutionModeEnum.Synchronous,
        "",
        "TSIS2.Plugins.PostOperationts_operationriskassessmentCreate Plugin",
        1,
        IsolationModeEnum.Sandbox,
        Description = "After Operation Risk Assessment created, populate Risk Criteria Responses and Discretionary Score Responses")]
    public class PostOperationts_operationriskassessmentCreate : IPlugin
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
                    if (target.LogicalName.Equals(ts_operationriskassessment.EntityLogicalName))
                    {
                        ts_operationriskassessment operationRiskAssessment = target.ToEntity<ts_operationriskassessment>();

                        using (var serviceContext = new Xrm(service))
                        {
                            //Retrieve operation of risk assessment
                            ovs_operation operation = serviceContext.ovs_operationSet.FirstOrDefault(o => o.Id == operationRiskAssessment.ts_operation.Id);

                            //Retrieve Risk Criteria Operation Type M:M records of Operation's Operation Type
                            var riskCriteriaOperationTypes = serviceContext.ts_riskcriteria_ovs_operationtypeSet.Where(rkot => rkot.ovs_operationtypeid == operation.ovs_OperationTypeId.Id);

                            //For each Risk Criteria Operation Type, create Risk Criteria Response
                            foreach (var riskCriteriaOperationType in riskCriteriaOperationTypes)
                            {
                                //Retrieve Risk Criteria of Risk Criteria Operation Type
                                ts_riskcriteria riskCriteria = serviceContext.ts_riskcriteriaSet.FirstOrDefault(rk => rk.Id == riskCriteriaOperationType.ts_riskcriteriaid);

                                //Create Risk Criteria Response
                                ts_riskcriteriaresponse riskCriteriaResponse = new ts_riskcriteriaresponse();
                                riskCriteriaResponse.ts_Name = riskCriteria.ts_Name;
                                riskCriteriaResponse.ts_description = riskCriteria.ts_description;
                                riskCriteriaResponse.ts_operationriskassessment = new EntityReference(ts_operationriskassessment.EntityLogicalName, operationRiskAssessment.Id);
                                riskCriteriaResponse.ts_riskcriteria = new EntityReference(ts_riskcriteria.EntityLogicalName, riskCriteria.Id);
                                service.Create(riskCriteriaResponse);
                            }

                            //Retrieve Discretionary Factor Groupings of Operation Type
                            var discretionaryFactorGroupings = serviceContext.ts_discretionaryfactorgroupingSet.Where(dfg => dfg.ts_operationtype.Id == operation.ovs_OperationTypeId.Id);

                            //For each Discretionary Factor Grouping, create Discretionary Factor Response
                            foreach (var discretionaryFactorGrouping in discretionaryFactorGroupings)
                            {
                                //Retrieve Discretionary Factors of Discretionary Factor Grouping
                                var discretionaryFactors = serviceContext.ts_discretionaryfactorSet.Where(df => df.ts_discretionaryfactorgrouping.Id == discretionaryFactorGrouping.Id);

                                //For each Discretionary Factor, create Discretionary Factor Response
                                foreach (var discretionaryFactor in discretionaryFactors)
                                {
                                    //Create Discretionary Factor Response
                                    ts_discretionaryfactorresponse discretionaryFactorResponse = new ts_discretionaryfactorresponse();
                                    discretionaryFactorResponse.ts_Name = discretionaryFactor.ts_Name;
                                    discretionaryFactorResponse.ts_description = discretionaryFactor.ts_description;
                                    discretionaryFactorResponse.ts_operationriskassessment = new EntityReference(ts_operationriskassessment.EntityLogicalName, operationRiskAssessment.Id);
                                    discretionaryFactorResponse.ts_discretionaryfactor = new EntityReference(ts_discretionaryfactor.EntityLogicalName, discretionaryFactor.Id);
                                    discretionaryFactorResponse.ts_discretionaryfactorgrouping = new EntityReference(ts_discretionaryfactorgrouping.EntityLogicalName, discretionaryFactorGrouping.Id);
                                    service.Create(discretionaryFactorResponse);
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
                        tracingService.Trace("PostOperationts_operationriskassessmentCreate Plugin: {0}", ex.ToString());
                        throw;
                    }
                }
                catch (Exception ex)
                {
                    tracingService.Trace("PostOperationts_operationriskassessmentCreate Plugin: {0}", ex.ToString());
                    throw;
                }
            }
        }
    }
}