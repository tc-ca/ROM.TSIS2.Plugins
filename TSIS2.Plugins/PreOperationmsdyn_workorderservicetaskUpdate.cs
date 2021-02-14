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
using System.Json;

namespace TSIS2.Plugins
{

    [CrmPluginRegistration(
        MessageNameEnum.Update,
        "msdyn_workorderservicetask",
        StageEnum.PreOperation,
        ExecutionModeEnum.Synchronous,
        "",
        "TSIS2.Plugins.PreOperationmsdyn_workorderservicetaskUpdate Plugin",
        1,
        IsolationModeEnum.Sandbox,
        Image1Name = "PreImage", Image1Type = ImageTypeEnum.PreImage, Image1Attributes = "msdyn_workorder",
        Description = "On Work Order Service Task Update, create findings in order to display them in a case.")]
    public class PreOperationmsdyn_workorderservicetaskUpdate : IPlugin
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
                    // Cast the target to the expected entity
                    msdyn_workorderservicetask workOrderServiceTask = target.ToEntity<msdyn_workorderservicetask>();

                    // Get the referenced work order from the preImage
                    EntityReference workOrderReference = (EntityReference)preImageEntity.Attributes["msdyn_workorder"];

                    // Check if there is a questionnaire response in this update
                    if (!String.IsNullOrWhiteSpace(workOrderServiceTask.ovs_QuestionnaireReponse))
                    {
                        using (var serviceContext = new CrmServiceContext(service))
                        {

                            // Lookup the referenced work order
                            msdyn_workorder workOrder = serviceContext.msdyn_workorderSet.Where(wo => wo.Id == workOrderReference.Id).FirstOrDefault();

                            // parse json response
                            string jsonResponse = workOrderServiceTask.ovs_QuestionnaireReponse;
                            JsonValue jsonValue = JsonValue.Parse(jsonResponse);
                            JsonObject jsonObject = jsonValue as JsonObject;

                            // If there was at least one finding found
                            // - Create a case (if work order service task doesn't already belong to a case)
                            // - Mark the inspection result to fail
                            if (jsonObject.Keys.Any(k => k.StartsWith("finding")))
                            {
                                // If the work order is not null and is not already part of a case
                                if (workOrder != null && workOrder.msdyn_ServiceRequest == null)
                                {
                                    Incident newIncident = new Incident();
                                    newIncident.CustomerId = workOrder.msdyn_BillingAccount;
                                    newIncident.Title = workOrder.msdyn_BillingAccount.Name + " Work Order " + workOrder.msdyn_name + " Inspection Failed on " + DateTime.Now.ToString("dd-MM-yy");
                                    Guid newIncidentId = service.Create(newIncident);
                                    msdyn_workorder uWorkOrder = new msdyn_workorder();
                                    uWorkOrder.Id = workOrderReference.Id;
                                    uWorkOrder.msdyn_ServiceRequest = new EntityReference(Incident.EntityLogicalName, newIncidentId);
                                    service.Update(uWorkOrder);
                                    workOrderServiceTask.ovs_CaseId = new EntityReference(Incident.EntityLogicalName, newIncidentId);
                                }
                                // Already part of a case, just assign the work order case to the work order service task case
                                else
                                {
                                    workOrderServiceTask.ovs_CaseId = workOrder.msdyn_ServiceRequest;
                                }

                                // Mark the inspection result to fail
                                workOrderServiceTask.msdyn_inspectiontaskresult = msdyn_InspectionResult.Fail;


                                // loop through each root property in the json object
                                foreach (var rootProperty in jsonObject)
                                {
                                    // Check if the root property starts with finding
                                    if (rootProperty.Key.StartsWith("finding"))
                                    {
                                        var finding = rootProperty.Value;

                                        // if finding, does it already exist?
                                        var uniqueFindingName = workOrderServiceTask.Id.ToString() + "-" + rootProperty.Key.ToString();
                                        var existingFinding = serviceContext.ovs_FindingSet.FirstOrDefault(f => f.ovs_Finding1 == uniqueFindingName);
                                        if (existingFinding == null)
                                        {
                                            // if no, initialize new ovs_finding
                                            ovs_Finding newFinding = new ovs_Finding();
                                            newFinding.ovs_FindingProvisionReference = (string)finding["provisionReference"];
                                            newFinding.ovs_FindingProvisionText = (string)finding["provisionText"];
                                            newFinding.ovs_FindingComments = finding.ContainsKey("comments") ? (string)finding["comments"] : "";
                                            newFinding.ovs_FindingFile = finding.ContainsKey("documentaryEvidence") ? (string)finding["documentaryEvidence"] : "";
                                            newFinding.ovs_Finding1 = uniqueFindingName;

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

                }

                catch (FaultException<OrganizationServiceFault> ex)
                {
                    throw new InvalidPluginExecutionException("An error occurred in PostOperationmsdyn_workorderservicetaskUpdate Plugin.", ex);
                }

                catch (Exception ex)
                {
                    tracingService.Trace("PostOperationmsdyn_workorderservicetaskUpdate Plugin: {0}", ex.ToString());
                    throw;
                }
            }
        }
    }
}
