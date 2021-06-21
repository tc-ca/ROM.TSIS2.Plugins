using System;
using System.Json;
using System.Linq;
using System.ServiceModel;
using Microsoft.Xrm.Sdk;

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
        Image1Name = "PreImage", Image1Type = ImageTypeEnum.PreImage, Image1Attributes = "msdyn_name,msdyn_workorder,msdyn_percentcomplete,ovs_questionnaireresponse",
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
                        using (var serviceContext = new Xrm(service))
                        {
                            // Get the referenced work order from the preImage
                            EntityReference workOrderReference = (EntityReference)preImageEntity.Attributes["msdyn_workorder"];

                            // Lookup the referenced work order
                            msdyn_workorder workOrder = serviceContext.msdyn_workorderSet.Where(wo => wo.Id == workOrderReference.Id).FirstOrDefault();

                            // Determine if we use the questionnaire response from this update or from the pre-image since it is not always passed in the update
                            var questionnaireResponse = !String.IsNullOrEmpty(workOrderServiceTask.ovs_QuestionnaireResponse) ? workOrderServiceTask.ovs_QuestionnaireResponse : workOrderServiceTaskPreImage.ovs_QuestionnaireResponse;
                            if (!String.IsNullOrWhiteSpace(questionnaireResponse))
                            {
                                // parse json response
                                JsonValue jsonValue = JsonValue.Parse(questionnaireResponse);
                                JsonObject jsonObject = jsonValue as JsonObject;

                                // Retrieve all the findings belonging to this work order service task
                                var findings = serviceContext.ovs_FindingSet.Where(f => f.ovs_WorkOrderServiceTaskId.Id == workOrderServiceTask.Id).ToList();
                                var findingsCount = findings.Count();

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
                                        if (workOrder.ts_Site != null) newIncident.msdyn_FunctionalLocation = workOrder.ts_Site;
                                        if (workOrder.ts_Region != null) newIncident.ovs_Region = workOrder.ts_Region;
                                        if (workOrder.ts_Country != null) newIncident.ts_Country = workOrder.ts_Country;
                                        if (workOrder.msdyn_ServiceAccount != null) newIncident.CustomerId = workOrder.msdyn_ServiceAccount;
                                        if (workOrder.msdyn_ServiceAccount != null) newIncident.ts_Stakeholder = workOrder.msdyn_ServiceAccount;
                                        // Stakeholder is a mandatory field on work order but, just in case, throw an error
                                        if (workOrder.msdyn_ServiceAccount == null) throw new ArgumentNullException("msdyn_workorder.msdyn_ServiceAccount");

                                        Guid newIncidentId = service.Create(newIncident);
                                        service.Update(new msdyn_workorder
                                        {
                                            Id = workOrderReference.Id,
                                            msdyn_ServiceRequest = new EntityReference(Incident.EntityLogicalName, newIncidentId)
                                        });
                                        workOrderServiceTask.ovs_CaseId = new EntityReference(Incident.EntityLogicalName, newIncidentId);
                                    }
                                    // Already part of a case, just assign the work order case to the work order service task case
                                    else
                                    {
                                        workOrderServiceTask.ovs_CaseId = workOrder.msdyn_ServiceRequest;
                                    }

                                    // Mark the inspection result to fail
                                    workOrderServiceTask.msdyn_inspectiontaskresult = msdyn_inspectionresult.Fail;

                                    // loop through each root property in the json object
                                    foreach (var rootProperty in jsonObject)
                                    {
                                        // Check if the root property starts with finding
                                        if (rootProperty.Key.StartsWith("finding"))
                                        {
                                            var finding = rootProperty.Value;

                                            // if finding, does it already exist?
                                            var findingMappingKey = workOrderServiceTask.Id.ToString() + "-" + rootProperty.Key.ToString() + "-" + workOrder.ovs_asset.Id.ToString();
                                            var existingFinding = serviceContext.ovs_FindingSet.FirstOrDefault(f => f.ts_findingmappingkey == findingMappingKey);
                                            
                                            if (existingFinding == null)
                                            {
                                                // if no, initialize new ovs_finding
                                                ovs_Finding newFinding = new ovs_Finding();
                                                newFinding.ovs_FindingProvisionReference = finding.ContainsKey("provisionReference") ? (string)finding["provisionReference"] : "";
                                                newFinding.ts_findingProvisionTextEn = finding.ContainsKey("provisionTextEn") ? (string)finding["provisionTextEn"] : "";
                                                newFinding.ts_findingProvisionTextFr = finding.ContainsKey("provisionTextFr") ? (string)finding["provisionTextFr"] : "";
                                                newFinding.ovs_FindingComments = finding.ContainsKey("comments") ? (string)finding["comments"] : "";

                                                // Don't do anything with the files yet until we have the proper infrastructure decision
                                                //newFinding.ovs_FindingFile = finding.ContainsKey("documentaryEvidence") ? (string)finding["documentaryEvidence"] : "";

                                                // Setup the finding name
                                                // Findings are at the 100 level
                                                var wostName = preImageEntity.Attributes["msdyn_name"].ToString();
                                                var prefix = wostName.Replace("200-", "100-");
                                                var infix = (findingsCount > 0) ? findingsCount + 1 : 1;
                                                var suffix = 1;
                                                newFinding.ovs_Finding_1 = string.Format("{0}-{1}-{2}", prefix, infix, suffix);

                                                // Store the mapping key to keep track of mapping between finding and surveyjs questionnaire.
                                                newFinding.ts_findingmappingkey = findingMappingKey;

                                                // reference work order service task
                                                newFinding.ovs_WorkOrderServiceTaskId = new EntityReference(msdyn_workorderservicetask.EntityLogicalName, workOrderServiceTask.Id);
                                                newFinding.ts_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrder.Id);

                                                // reference case (should already be saved in the work order service task)
                                                newFinding.ovs_CaseId = new EntityReference(Incident.EntityLogicalName, workOrderServiceTask.ovs_CaseId.Id);

                                                // reference Work Order's Stakeholder (Account Entity) (Lookup logical name: msdyn_serviceaccount)
                                                newFinding.ts_accountid = new EntityReference(Account.EntityLogicalName, workOrder.msdyn_ServiceAccount.Id);

                                                // reference Work Order's Operation (Customer Asset Entity) (Lookup logical name: ovs_asset)
                                                newFinding.ts_Assetid = new EntityReference(msdyn_customerasset.EntityLogicalName, workOrder.ovs_asset.Id);

                                                // Create new ovs_finding
                                                Guid newFindingId = service.Create(newFinding);

                                                // The finding JSON may contain an array of string Id's of operation records
                                                var operations = finding.ContainsKey("operations") ? finding["operations"] : new JsonArray();

                                                //Iterate through operations array and create a copy of the finding for each operation, associated to their opperation's parent account
                                                foreach (System.Json.JsonPrimitive operation in operations)
                                                {
                                                    suffix++;
                                                    var operationid = (string)operation;
                                                    string newMappingKey = workOrderServiceTask.Id.ToString() + "-" + rootProperty.Key.ToString() + "-" + operationid;
                                                    EntityReference operationReference = new EntityReference(msdyn_customerasset.EntityLogicalName, new Guid(operationid));

                                                    // Lookup the operation (Customer Asset Entity) to know its parent Account's id
                                                    msdyn_customerasset operationEntity = serviceContext.msdyn_customerassetSet.Where(op => op.Id == operationReference.Id).FirstOrDefault();

                                                    // Create entity reference to the operation's parent account
                                                    EntityReference parentAccountReference = new EntityReference(Account.EntityLogicalName, operationEntity.msdyn_Account.Id);

                                                    //EntityReference parentAccountReference = new EntityReference(Account.EntityLogicalName, operationReference.)
                                                    ovs_Finding newFindingCopy = new ovs_Finding()
                                                    {
                                                        ovs_FindingProvisionReference = newFinding.ovs_FindingProvisionReference,
                                                        ts_findingProvisionTextEn = newFinding.ts_findingProvisionTextEn,
                                                        ts_findingProvisionTextFr = newFinding.ts_findingProvisionTextFr,
                                                        ovs_FindingComments = newFinding.ovs_FindingComments,
                                                        ts_findingmappingkey = newMappingKey,
                                                        ts_WorkOrder = newFinding.ts_WorkOrder,
                                                        ovs_WorkOrderServiceTaskId = newFinding.ovs_WorkOrderServiceTaskId,
                                                        ovs_CaseId = newFinding.ovs_CaseId,
                                                        ts_Assetid = operationReference,
                                                        ts_accountid = new EntityReference(Account.EntityLogicalName, operationEntity.msdyn_Account.Id),
                                                        ovs_Finding_1 = string.Format("{0}-{1}-{2}", prefix, infix, suffix),
                                                    };

                                                    service.Create(newFindingCopy);
                                                }

                                                // Increment findings count for next finding name
                                                findingsCount++;
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
                                    workOrderServiceTask.msdyn_inspectiontaskresult = msdyn_inspectionresult.Pass;
                                }

                                // Need to deactivate any old referenced findings in the work order service task and case
                                // that no longer exist in the questionnaire response.

                                // Get a list of unique finding names from the JSON response
                                var findingMappingKeys = jsonObject.Keys.Select(k => workOrderServiceTask.Id.ToString() + "-" + k + "-" + workOrder.ovs_asset.Id.ToString());

                                foreach (var finding in findings)
                                {
                                    // If the existing finding is not in the JSON response, we need to disable it
                                    if (!findingMappingKeys.Contains(finding.ts_findingmappingkey))
                                    {
                                        finding.statuscode = ovs_Finding_statuscode.Inactive;
                                        finding.statecode = ovs_FindingState.Inactive;
                                    }
                                    // Otherwise, re-enable it
                                    else
                                    {
                                        finding.statuscode = ovs_Finding_statuscode.Active;
                                        finding.statecode = ovs_FindingState.Active;
                                    }
                                    serviceContext.UpdateObject(finding);
                                }
                            }

                            // If the work order is not already "complete" or "closed" and all other work order service tasks are already completed as well, mark the parent work order system status to Open - Completed
                            var otherWorkOrderServiceTasks = serviceContext.msdyn_workorderservicetaskSet.Where(wost => wost.msdyn_WorkOrder == workOrderReference && wost.Id != workOrderServiceTask.Id).ToList<msdyn_workorderservicetask>();
                            if(workOrder.msdyn_SystemStatus != msdyn_wosystemstatus.ClosedPosted && workOrder.msdyn_SystemStatus != msdyn_wosystemstatus.OpenCompleted)
                            {
                                if (otherWorkOrderServiceTasks.All(x => x.statuscode == msdyn_workorderservicetask_statuscode.Complete))
                                {
                                    service.Update(new msdyn_workorder
                                    {
                                        Id = workOrderReference.Id,
                                        msdyn_SystemStatus = msdyn_wosystemstatus.OpenCompleted
                                    });
                                }
                            }
                            // Save all the changes in the context as well.
                            serviceContext.SaveChanges();
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
