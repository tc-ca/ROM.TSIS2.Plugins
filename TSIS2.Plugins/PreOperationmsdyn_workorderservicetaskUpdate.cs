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
using Microsoft.Xrm.Sdk.Query;
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
        Image1Name = "PreImage", Image1Type = ImageTypeEnum.PreImage, Image1Attributes = "msdyn_workorder,msdyn_percentcomplete,ovs_questionnaireresponse",
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
                    // Cast the target and preimage to the expected entity
                    msdyn_workorderservicetask workOrderServiceTask = target.ToEntity<msdyn_workorderservicetask>();
                    msdyn_workorderservicetask workOrderServiceTaskPreImage = preImageEntity.ToEntity<msdyn_workorderservicetask>();

                    // Only perform the updates if the work order service task is 100% complete on update
                    // or work order service task was already 100% complete (from pre-image)
                    if (workOrderServiceTask.msdyn_PercentComplete == 100.00 || (workOrderServiceTask.msdyn_PercentComplete == null && workOrderServiceTaskPreImage.msdyn_PercentComplete == 100.00))
                    {

                        // Get the referenced work order from the preImage
                        EntityReference workOrderReference = (EntityReference)preImageEntity.Attributes["msdyn_workorder"];

                        // Determine if we use the questionnaire response from this update or from the pre-image since it is not always passed in the update
                        var questionnaireResponse = !String.IsNullOrEmpty(workOrderServiceTask.ovs_QuestionnaireResponse) ? workOrderServiceTask.ovs_QuestionnaireResponse : workOrderServiceTaskPreImage.ovs_QuestionnaireResponse;
                        if (!String.IsNullOrWhiteSpace(questionnaireResponse))
                        {
                            using (var serviceContext = new CrmServiceContext(service))
                            {

                                // Lookup the referenced work order
                                msdyn_workorder workOrder = serviceContext.msdyn_workorderSet.Where(wo => wo.Id == workOrderReference.Id).FirstOrDefault();

                                // parse json response
                                JsonValue jsonValue = JsonValue.Parse(questionnaireResponse);
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
                                        newIncident.CustomerId = workOrder.ovs_regulatedentity;

                                        // Regulated Entity is a mandatory field on work order but, just in case, throw an error
                                        if (workOrder.ovs_regulatedentity == null) throw new ArgumentNullException("msdyn_workorder.ovs_regulatedentity");

                                        newIncident.Title = workOrder.ovs_regulatedentity.Name + " Work Order " + workOrder.msdyn_name + " Inspection Failed on " + DateTime.Now.ToString("dd-MM-yy");
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

                                                // Don't do anything with the files yet until we have the proper infrastructure decision
                                                //newFinding.ovs_FindingFile = finding.ContainsKey("documentaryEvidence") ? (string)finding["documentaryEvidence"] : "";

                                                newFinding.ovs_Finding1 = uniqueFindingName;

                                                // reference work order service task
                                                newFinding.ovs_WorkOrderServiceTaskId = new EntityReference(msdyn_workorderservicetask.EntityLogicalName, workOrderServiceTask.Id);

                                                // reference case (should already be saved in the work order service task)
                                                newFinding.ovs_CaseId = new EntityReference(Incident.EntityLogicalName, workOrderServiceTask.ovs_CaseId.Id);

                                                // Create new ovs_finding
                                                Guid newFindingId = service.Create(newFinding);
                                            }
                                            else
                                            {
                                                // Retrieve the account containing several of its attributes.
                                                //ColumnSet cols = new ColumnSet(new String[] { });

                                                // Update existing finding
                                                existingFinding.ovs_FindingComments = finding.ContainsKey("comments") ? (string)finding["comments"] : "";

                                                // Don't do anything with the files yet until we have the proper infrastructure decision
                                                //existingFinding.ovs_FindingFile = finding.ContainsKey("documentaryEvidence") ? (string)finding["documentaryEvidence"] : "";

                                                serviceContext.UpdateObject(existingFinding);
                                            }
                                        }
                                    }

                                }
                                else
                                {
                                    // Mark the inspection result to Pass if there are no findings found
                                    workOrderServiceTask.msdyn_inspectiontaskresult = msdyn_InspectionResult.Pass;
                                }

                                // Need to deactivate any old referenced findings in the work order service task and case
                                // that no longer exist in the questionnaire response.
                                // Retrieve all the findings belonging to this work order service task
                                var findings = serviceContext.ovs_FindingSet.Where(f => f.ovs_WorkOrderServiceTaskId.Id == workOrderServiceTask.Id).ToList();
                                // Get a list of unique finding names from the JSON response
                                var uniqueFindingNames = jsonObject.Keys.Select(k => workOrderServiceTask.Id.ToString() + "-" + k);

                                foreach (var finding in findings)
                                {
                                    // If the existing finding is not in the JSON response, we need to disable it
                                    if (!uniqueFindingNames.Contains(finding.ovs_Finding1))
                                    {
                                        finding.StatusCode = ovs_Finding_StatusCode.Inactive;
                                        finding.StateCode = ovs_FindingState.Inactive;
                                    } 
                                    // Otherwise, re-enable it
                                    else
                                    {
                                        finding.StatusCode = ovs_Finding_StatusCode.Active;
                                        finding.StateCode = ovs_FindingState.Active;
                                    }
                                    serviceContext.UpdateObject(finding);
                                }

                                // Save all the changes made in the service context
                                serviceContext.SaveChanges();
                            }
                        }

                    }
                }

                // Seems to be a bug if exception variables have the same name. 
                // Make sure the name of each exception variable is different.
                catch (FaultException<OrganizationServiceFault> orgServiceEx)
                {
                    throw new InvalidPluginExecutionException("An error occurred in PreOperationmsdyn_workorderservicetaskUpdate Plugin.", orgServiceEx);
                }

                catch (Exception ex)
                {
                    tracingService.Trace("PreOperationmsdyn_workorderservicetaskUpdate Plugin: {0}", ex.ToString());
                    throw;
                }
            }
        }
    }
}
