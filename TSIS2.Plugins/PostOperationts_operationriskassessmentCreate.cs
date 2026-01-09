using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;

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
    public class PostOperationts_operationriskassessmentCreate : PluginBase
    {
        public PostOperationts_operationriskassessmentCreate() : base(typeof(PostOperationts_operationriskassessmentCreate))
        {
        }

        protected override void ExecuteCrmPlugin(LocalPluginContext localContext)
        {
            var context = localContext.PluginExecutionContext;
            var service = localContext.OrganizationService;

            // The InputParameters collection contains all the data passed in the message request.
            if (context.InputParameters.Contains("Target") &&
                context.InputParameters["Target"] is Entity)
            {
                // Obtain the target entity from the input parameters.
                Entity target = (Entity)context.InputParameters["Target"];

                // Obtain the preimage entity
                Entity preImageEntity = (context.PreEntityImages != null && context.PreEntityImages.Contains("PreImage")) ? context.PreEntityImages["PreImage"] : null;

                try
                {
                    if (target.LogicalName.Equals(ts_operationriskassessment.EntityLogicalName))
                    {
                        ts_operationriskassessment operationRiskAssessment = target.ToEntity<ts_operationriskassessment>();

                        using (var serviceContext = new Xrm(service))
                        {
                            if (operationRiskAssessment.ts_operation == null) return;

                            //Retrieve operation of risk assessment
                            ovs_operation operation = serviceContext.ovs_operationSet.FirstOrDefault(o => o.Id == operationRiskAssessment.ts_operation.Id);
                            
                            if (operation.ovs_OperationTypeId != null) {
                                //Retrieve Risk Criteria Operation Type M:M records of Operation's Operation Type
                                var riskCriteriaOperationTypes = serviceContext.ts_riskcriteria_ovs_operationtypeSet.Where(rkot => rkot.ovs_operationtypeid == operation.ovs_OperationTypeId.Id).ToList();

                                // this is used for debugging purposes
                                List<FoundRiskCriteraOption> foundRiskCriteriaOptions = new List<FoundRiskCriteraOption>();

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

                                    //Get the Risk Critera Response ID
                                    Guid newRiskCriteriaResponseID = service.Create(riskCriteriaResponse);

                                    // Create the Risk Criteria Option for the Risk Critera Response
                                    if (operationRiskAssessment.ts_RiskCriteriaOptionsImport != null && newRiskCriteriaResponseID != null)
                                    {
                                        try
                                        {
                                            // Deserialize the JSON string into a dynamic object
                                            dynamic riskCriteriaOptions = JsonConvert.DeserializeObject(operationRiskAssessment.ts_RiskCriteriaOptionsImport);

                                            // Get all the related Risk Criteria Options that are available
                                            string riskCriteriaOptionFetchXML = $@"
                                                <fetch xmlns:generator='MarkMpn.SQL4CDS'>
                                                  <entity name='ts_riskcriteriaoption'>
                                                    <attribute name='ts_riskcriteriaoptionid' />
                                                    <attribute name='ts_name' />
                                                    <link-entity name='ts_riskcriteria' to='ts_riskcriteria' from='ts_riskcriteriaid' alias='ts_riskcriteria' link-type='inner'>
                                                      <attribute name='ts_englishtext' />
                                                      <filter>
                                                        <condition attribute='ts_riskcriteriaid' operator='eq' value='{riskCriteria.Id.ToString()}' />
                                                      </filter>
                                                    </link-entity>
                                                  </entity>
                                                </fetch>
                                            ";

                                            EntityCollection myRiskCriteriaOptionEntityCollection = service.RetrieveMultiple(new FetchExpression(riskCriteriaOptionFetchXML));

                                            foreach (var importedRiskCriteriaOption in riskCriteriaOptions)
                                            {
                                                PropertyInfo namePropertyInfo = importedRiskCriteriaOption.GetType().GetProperty("Name");
                                                string importedRiskCriteriaOptionName = namePropertyInfo.GetValue(importedRiskCriteriaOption, null).ToString().Trim().ToUpper();

                                                PropertyInfo valuePropertyInfo = importedRiskCriteriaOption.GetType().GetProperty("Value");
                                                string importedRiskCriteriaOptionValue = valuePropertyInfo.GetValue(importedRiskCriteriaOption, null).ToString();
                                                importedRiskCriteriaOptionValue = importedRiskCriteriaOptionValue.Trim().ToUpper();

                                                // Iterate through each risk criteria option
                                                foreach (var myRiskCriteriaOption in myRiskCriteriaOptionEntityCollection.Entities)
                                                {
                                                    string riskCriteriaOptionEnglishText = "";

                                                    if (myRiskCriteriaOption.Attributes["ts_riskcriteria.ts_englishtext"] is AliasedValue ts_englishtext)
                                                    {
                                                        riskCriteriaOptionEnglishText = ts_englishtext.Value.ToString().Trim().ToUpper();
                                                    }

                                                    string riskCriteriaOptionName = myRiskCriteriaOption.Attributes["ts_name"].ToString().Trim().ToUpper();

                                                    // Match the Risk Criteria Option with the Imported Risk Criteria Option (JSON)
                                                    // This code is used when records are manually being imported with SSIS through an Excel file
                                                    if (importedRiskCriteriaOptionName == riskCriteriaOptionEnglishText &&
                                                        importedRiskCriteriaOptionValue == riskCriteriaOptionName)
                                                    {
                                                        // this is used for debugging purposes
                                                        foundRiskCriteriaOptions.Add(new FoundRiskCriteraOption { 
                                                            Id = myRiskCriteriaOption.Attributes["ts_riskcriteriaoptionid"].ToString() ,
                                                            RiskCriteriaName = riskCriteriaOptionEnglishText ,
                                                            RiskCriteriaOptionName = myRiskCriteriaOption.Attributes["ts_name"].ToString()
                                                        });

                                                        Guid foundRiskCriteriaOptionId = new Guid(myRiskCriteriaOption.Attributes["ts_riskcriteriaoptionid"].ToString());

                                                        // Update the Risk Criteria Option of the Risk Criteria that is part of the Operation Risk Assessment
                                                        ts_riskcriteriaoption myUpdatedRiskCriteriaOption = serviceContext.ts_riskcriteriaoptionSet.Where(rco => rco.Id == foundRiskCriteriaOptionId).FirstOrDefault();

                                                        service.Update(new ts_riskcriteriaresponse { 
                                                            Id = newRiskCriteriaResponseID,
                                                            ts_riskcriteriaoption = myUpdatedRiskCriteriaOption.ToEntityReference()
                                                        });
                                                    }
                                                }
                                            }
                                        }
                                        catch (JsonReaderException)
                                        {
                                            localContext.Trace($"PostOperationts_operationriskassessmentCreate: Invalid JSON format for operationRiskAssessment {operationRiskAssessment.Id.ToString()} ts_RiskCriteriaOptionsImport value:   {operationRiskAssessment.ts_RiskCriteriaOptionsImport}");
                                        }
                                    }
                                }

                                //Retrieve Discretionary Factor Groupings of Operation Type
                                var discretionaryFactorGroupings = serviceContext.ts_discretionaryfactorgroupingSet.Where(dfg => dfg.ts_operationtype.Id == operation.ovs_OperationTypeId.Id).ToList();

                                //For each Discretionary Factor Grouping, create Discretionary Factor Response
                                foreach (var discretionaryFactorGrouping in discretionaryFactorGroupings)
                                {
                                    //Retrieve Discretionary Factors of Discretionary Factor Grouping
                                    var discretionaryFactors = serviceContext.ts_discretionaryfactorSet.Where(df => df.ts_discretionaryfactorgrouping.Id == discretionaryFactorGrouping.Id).ToList();

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

                            

                            //Retrieve any other active operation risk assessments and set to inactive
                            var activeOperationRiskAssessments = serviceContext.ts_operationriskassessmentSet.Where(ora => ora.ts_operation.Id == operation.Id && ora.Id != operationRiskAssessment.Id && ora.statecode == ts_operationriskassessmentState.Active).ToList();
                            foreach (var activeOperationRiskAssessment in activeOperationRiskAssessments)
                            {
                                activeOperationRiskAssessment.statecode = ts_operationriskassessmentState.Inactive;
                                activeOperationRiskAssessment.statuscode = ts_operationriskassessment_statuscode.Inactive;
                                serviceContext.UpdateObject(activeOperationRiskAssessment);
                            }
                            serviceContext.SaveChanges();
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
                        localContext.Trace("PostOperationts_operationriskassessmentCreate Plugin: {0}", ex.ToString());
                        throw;
                    }
                }
                catch (Exception ex)
                {
                    localContext.TraceWithContext("PostOperationts_operationriskassessmentCreate Plugin: {0}", ex.ToString());
                    throw new InvalidPluginExecutionException("PostOperationts_operationriskassessmentCreate failed.", ex);
                }
            }
        }
    }

    public class FoundRiskCriteraOption
    {
        public string Id { get; set; }
        public string RiskCriteriaName { get; set; }
        public string RiskCriteriaOptionName { get; set; }
    }
}