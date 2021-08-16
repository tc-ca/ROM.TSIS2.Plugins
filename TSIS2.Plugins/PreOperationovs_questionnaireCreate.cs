using System;
using System.Linq;
using Microsoft.Xrm.Sdk;

namespace TSIS2.Plugins
{

    [CrmPluginRegistration(
        MessageNameEnum.Create,
        "ovs_questionnaire",
        StageEnum.PreOperation,
        ExecutionModeEnum.Synchronous,
        "",
        "TSIS2.Plugins.PreOperationovs_questionnaireCreate Plugin",
        1,
        IsolationModeEnum.Sandbox,
        Description = "On Questionnaire Create, automatically create one Questionnaire Version")]
    public class PreOperationovs_questionnaireCreate : IPlugin
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
                    if (target.LogicalName.Equals(ovs_Questionnaire.EntityLogicalName))
                    {
                        // Cast the target to the expected entity
                        ovs_Questionnaire questionnaire = target.ToEntity<ovs_Questionnaire>();

                        // Check but there should always be zero questionnaire versions when first creating a questionnaire
                        if (questionnaire.ts_ovs_questionnaire_ovs_questionnaire.Count() == 0)
                        {
                            // Setup a new questionnaire version
                            var questionnaireVersion = new ts_questionnaireversion();
                            questionnaireVersion.ts_name = "Version 1";

                            // Assign the parent of this questionnaire version to be the new questionnaire
                            questionnaireVersion.ts_ovs_questionnaire = new EntityReference(ovs_Questionnaire.EntityLogicalName, target.Id);

                            // Create the new questionnaire version
                            Guid newQuestionnaireVersionId = service.Create(questionnaireVersion);
                        }
                    }
                }
                catch (Exception ex)
                {
                    tracingService.Trace("PreOperationovs_questionnaireCreate Plugin: {0}", ex.ToString());
                    throw;
                }
            }
        }
    }
}