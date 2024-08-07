﻿using System;
using System.Linq;
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
    public class PostOperationovs_questionnaireCreate : IPlugin
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
                        CalculateRollupFieldResponse calculateRollupFieldResponse = (CalculateRollupFieldResponse)service.Execute(calculateRollupFieldRequest);
                    }
                }
                catch (NotImplementedException ex)
                {
                    // If exceptions from mocking library, just continue. If not from there, we should still throw the error.
                    if (ex.Source == "XrmMockup365" && ex.Message == "No implementation for expression operator 'Count'")
                    {
                        // continue
                    } else
                    {
                        tracingService.Trace("PostOperationovs_questionnaireCreate Plugin: {0}", ex.ToString());
                        throw;
                    }
                }
                catch (Exception ex)
                {
                    tracingService.Trace("PostOperationovs_questionnaireCreate Plugin: {0}", ex.ToString());
                    throw;
                }
            }
        }
    }
}