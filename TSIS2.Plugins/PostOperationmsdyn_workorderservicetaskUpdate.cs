using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Runtime.Serialization.Json;
using System.ServiceModel;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Xrm.Sdk;
using TSIS2.Common;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace TSIS2.Plugins
{

    [CrmPluginRegistration(
        MessageNameEnum.Update,
        "msdyn_workorderservicetask",
        StageEnum.PostOperation,
        ExecutionModeEnum.Asynchronous,
        "",
        "PostOperationmsdyn_workorderservicetaskUpdate Plugin",
        1,
        IsolationModeEnum.Sandbox)]
    public class PostOperationmsdyn_workorderservicetaskUpdate : IPlugin
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
                Entity entity = (Entity)context.InputParameters["Target"];

                // Obtain the organization service reference which you will need for  
                // web service calls.  
                IOrganizationServiceFactory serviceFactory =
                    (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

                try
                {
                    // Plug-in business logic goes here.  
                    /*
                     * Check if questionnaire json is there
                     *  - if yes, parse saved questionnaire json response
                     *      - if "finding" found and it doesn't already exist
                     *          - create ovs_finding
                     *          - reference case
                     *          - reference work order service task
                    */

                    msdyn_workorderservicetask workOrderServiceTask = entity.ToEntity<msdyn_workorderservicetask>();
                    if (workOrderServiceTask.ovs_QuestionnaireReponse != null)
                    {
                        // parse json response
                        var jsonResponse = workOrderServiceTask.ovs_QuestionnaireReponse;
                        JObject o = JObject.Parse(jsonResponse);

                        using (var serviceContext = new CrmServiceContext(service))
                        {
                            // loop through each root property in the json object
                            foreach (var rootProperty in o)
                            {
                                // Check if the root property starts with finding
                                if (rootProperty.Key.StartsWith("finding"))
                                {
                                    var finding = rootProperty.Value;

                                    // if finding, does it already exist?
                                    var existingFinding = serviceContext.ovs_FindingSet.FirstOrDefault(f => f.ovs_FindingProvisionReference == finding["provisionReference"].ToString());
                                    if (existingFinding == null)
                                    {
                                        // if no, initialize new ovs_finding
                                        ovs_Finding newFinding = new ovs_Finding();
                                        newFinding.ovs_FindingProvisionReference = finding["provisionReference"].ToString();
                                        newFinding.ovs_FindingProvisionText = finding["provisionText"].ToString();
                                        newFinding.ovs_FindingComments = finding["comments"].ToString();
                                        newFinding.ovs_FindingFile = finding["documentaryEvidence"].ToString();

                                        // reference work order service task
                                        newFinding.ovs_WorkOrderServiceTaskId = new EntityReference(msdyn_workorderservicetask.EntityLogicalName, workOrderServiceTask.Id);

                                        // reference case (should already be saved in the work order service task)
                                        newFinding.ovs_CaseId = new EntityReference(Incident.EntityLogicalName, workOrderServiceTask.ovs_CaseId.Id);

                                        // Create new ovs_finding
                                        Guid newFindingId = service.Create(newFinding);
                                    }
                                }
                            }

                        }
                    }

                }

                catch (FaultException<OrganizationServiceFault> ex)
                {
                    throw new InvalidPluginExecutionException("An error occurred in FollowUpPlugin.", ex);
                }

                catch (Exception ex)
                {
                    tracingService.Trace("FollowUpPlugin: {0}", ex.ToString());
                    throw;
                }
            }
        }
    }
}
