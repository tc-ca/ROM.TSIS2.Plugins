using System;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;

namespace TSIS2.Plugins
{

    [CrmPluginRegistration(
        MessageNameEnum.Create,
        "ovs_questionnaire",
        StageEnum.PostOperation,
        ExecutionModeEnum.Synchronous,
        "",
        "TSIS2.Plugins.PostOperationovs_questionnaireCreate Plugin",
        1,
        IsolationModeEnum.Sandbox,
        Description = "After Questionnaire created, automatically create one Questionnaire Version")]
    public class PostOperationovs_questionnaireCreate : PluginBase
    {
        public PostOperationovs_questionnaireCreate(string unsecure, string secure)
            : base(typeof(PostOperationovs_questionnaireCreate))
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

            // The InputParameters collection contains all the data passed in the message request.
            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
            {
                // Obtain the target entity from the input parameters.
                Entity target = (Entity)context.InputParameters["Target"];

                try
                {
                    if (target.LogicalName.Equals(ovs_Questionnaire.EntityLogicalName))
                    {
                        ovs_Questionnaire questionnaire = target.ToEntity<ovs_Questionnaire>();

                        // Setup a new questionnaire version
                        var questionnaireVersion = new ts_questionnaireversion();
                        questionnaireVersion.ts_name = questionnaire.ovs_Name + " - Version 1";

                        // Create the new questionnaire version
                        Guid newQuestionnaireVersionId = service.Create(questionnaireVersion);

                        // Associate the newly created questionnaire version to the target questionnaire
                        service.Associate(
                            ovs_Questionnaire.EntityLogicalName,
                            questionnaire.Id,
                            new Relationship("ts_ovs_questionnaire_ovs_questionnaire"),
                            new EntityReferenceCollection {
                                new EntityReference(ts_questionnaireversion.EntityLogicalName, newQuestionnaireVersionId)
                            }
                        );

                        // Refresh the rollup field to make sure the number of versions is correct
                        // NOT SUPPORTED BY XRMMOCKUP TO TEST
                        CalculateRollupFieldRequest calculateRollupFieldRequest = new CalculateRollupFieldRequest
                        {
                            Target = questionnaire.ToEntityReference(),
                            FieldName = "ts_numberofversions"
                        };
                        CalculateRollupFieldResponse calculateRollupFieldResponse =
                            (CalculateRollupFieldResponse)service.Execute(calculateRollupFieldRequest);
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
                        localContext.Trace($"PostOperationovs_questionnaireCreate Plugin: {ex}");
                        throw new InvalidPluginExecutionException("PostOperationovs_questionnaireCreate failed", ex);
                    }
                }
                catch (Exception ex)
                {
                    localContext.Trace($"PostOperationovs_questionnaireCreate Plugin: {ex}");
                    throw new InvalidPluginExecutionException("PostOperationovs_questionnaireCreate failed", ex);
                }
            }
        }
    }
}