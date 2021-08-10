using System;
using System.Collections.Generic;
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

                                // Start a list of all the used mapping keys
                                var findingMappingKeys = new List<string>();

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

                                            // The finding JSON may contain an array of string Id's of operation records
                                            var operations = finding.ContainsKey("operations") ? finding["operations"] : new JsonArray();

                                            //Loop through the operations. Check if a finding already exists for that operation. Update the comment if it exists, or make a new finding if it doesn't
                                            foreach (JsonObject operation in operations)
                                            {
                                                string operationid;
                                                //Grab operationid from Json operation object. If it wasn't populated for some reason, continue to next operation
                                                if (operation.ContainsKey("operationID"))
                                                {
                                                    operationid = operation["operationID"];
                                                } else
                                                {
                                                    continue;
                                                }
                                                var findingMappingKey = workOrderServiceTask.Id.ToString() + "-" + rootProperty.Key.ToString() + "-" + operationid;
                                                findingMappingKeys.Add(findingMappingKey);
                                                var existingFinding = serviceContext.ovs_FindingSet.FirstOrDefault(f => f.ts_findingmappingkey == findingMappingKey);
                                                if (existingFinding == null)
                                                {
                                                    // if no, initialize new ovs_finding
                                                    ovs_Finding newFinding = new ovs_Finding();
                                                    newFinding.ovs_FindingProvisionReference = finding.ContainsKey("provisionReference") ? (string)finding["provisionReference"] : "";
                                                    newFinding.ts_findingProvisionTextEn = finding.ContainsKey("provisionTextEn") ? (string)finding["provisionTextEn"] : "";
                                                    newFinding.ts_findingProvisionTextFr = finding.ContainsKey("provisionTextFr") ? (string)finding["provisionTextFr"] : "";
                                                    newFinding.ovs_FindingComments = finding.ContainsKey("comments") ? (string)finding["comments"] : "";
                                                    OptionSetValue findingType = operation.ContainsKey("findingType") ? new OptionSetValue(operation["findingType"]) : new OptionSetValue(717750000); //717750000 is Undecided
                                                    newFinding.Attributes.Add("ts_findingtype", findingType);

                                                    //Update the list of findings for this service task in case a finding was added in a previous loop
                                                    findings = serviceContext.ovs_FindingSet.Where(f => f.ovs_WorkOrderServiceTaskId.Id == workOrderServiceTask.Id).ToList();

                                                    //Find all the findings that share the same root property key and add to findingCopies list
                                                    //The count of copies is needed to determine the suffix for the next finding record created
                                                    var findingCopies = new List<ovs_Finding>();
                                                    foreach (var f in findings) {
                                                        if (f.ts_findingmappingkey.StartsWith(workOrderServiceTask.Id.ToString() + "-" + rootProperty.Key.ToString())) findingCopies.Add(f);                                                       
                                                    }

                                                    // Determine the current highest infix of all the findings for the service task
                                                    var highestInfix = 0;
                                                    foreach (ovs_Finding f in findings)
                                                    {
                                                        var currentInfix = Int32.Parse(f.ovs_Finding_1.Split('-')[3]);
                                                        if (currentInfix > highestInfix) highestInfix = currentInfix;
                                                    }

                                                    // Setup the finding name
                                                    // Findings are at the 100 level
                                                    var wostName = preImageEntity.Attributes["msdyn_name"].ToString();
                                                    var prefix = wostName.Replace("200-", "100-");
                                                    var infix = highestInfix + 1;
                                                    //If there are copies for this finding, use their infix instead
                                                    if (findingCopies.Count > 0)
                                                    {
                                                        infix = Int32.Parse(findingCopies[0].ovs_Finding_1.Split('-')[3]);
                                                    }
                                                    var suffix = findingCopies.Count + 1;
                                                    newFinding.ovs_Finding_1 = string.Format("{0}-{1}-{2}", prefix, infix, suffix);

                                                    // Store the mapping key to keep track of mapping between finding and surveyjs questionnaire.
                                                    newFinding.ts_findingmappingkey = findingMappingKey;

                                                    // reference work order service task
                                                    newFinding.ovs_WorkOrderServiceTaskId = new EntityReference(msdyn_workorderservicetask.EntityLogicalName, workOrderServiceTask.Id);
                                                    newFinding.ts_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrder.Id);

                                                    // reference case (should already be saved in the work order service task)
                                                    newFinding.ovs_CaseId = new EntityReference(Incident.EntityLogicalName, workOrderServiceTask.ovs_CaseId.Id);

                                                    EntityReference operationReference = new EntityReference(ovs_operation.EntityLogicalName, new Guid(operationid));

                                                    // Lookup the operation (Customer Asset Entity) to know its parent Account's id
                                                    ovs_operation operationEntity = serviceContext.ovs_operationSet.Where(ca => ca.Id == operationReference.Id).FirstOrDefault();

                                                    // Create entity reference to the operation's parent account
                                                    EntityReference parentAccountReference = new EntityReference(Account.EntityLogicalName, operationEntity.ts_stakeholder.Id);

                                                    // reference the current operation's stakeholder (Account Entity) (Lookup logical name: ts_stakeholder)
                                                    newFinding.ts_accountid = new EntityReference(Account.EntityLogicalName, operationEntity.ts_stakeholder.Id);

                                                    // reference current operation (Customer Asset Entity) (Lookup logical name: ovs_asset)
                                                    newFinding.ts_operationid = operationReference;

                                                    // Create new ovs_finding
                                                    Guid newFindingId = service.Create(newFinding);

                                                }
                                                else
                                                {
                                                    //Update the Finding record's comment
                                                    existingFinding.ovs_FindingComments = finding.ContainsKey("comments") ? (string)finding["comments"] : "";
                                                    //Update the Finding record's Finding Type
                                                    OptionSetValue findingType = operation.ContainsKey("findingType") ? new OptionSetValue(operation["findingType"]) : new OptionSetValue(717750000); //717750000 is Undecided
                                                    existingFinding["ts_findingtype"] = findingType;
                                                    serviceContext.UpdateObject(existingFinding);

                                                }
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
