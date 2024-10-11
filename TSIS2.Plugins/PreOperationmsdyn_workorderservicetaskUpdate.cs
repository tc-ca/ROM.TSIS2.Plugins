using System;
using System.Collections.Generic;
using System.Json;
using System.Linq;
using System.ServiceModel;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

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

            tracingService.Trace("Tracking Service Started.");

            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)
                serviceProvider.GetService(typeof(IPluginExecutionContext));

            tracingService.Trace("Return if triggered by another plugin. Prevents infinite loop.");

            if (context.Depth > 1)
                return;

            tracingService.Trace("The InputParameters collection contains all the data passed in the message request.");
            if (context.InputParameters.Contains("Target") &&
                context.InputParameters["Target"] is Entity)
            {
                tracingService.Trace("Obtain the target entity from the input parameters.");
                Entity target = (Entity)context.InputParameters["Target"];

                tracingService.Trace("Obtain the preimage entity");
                Entity preImageEntity = (context.PreEntityImages != null && context.PreEntityImages.Contains("PreImage")) ? context.PreEntityImages["PreImage"] : null;

                tracingService.Trace("Obtain the organization service reference which you will need for web service calls");
                IOrganizationServiceFactory serviceFactory =
                    (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

                try
                {
                    tracingService.Trace("Cast the target and preimage to the expected entity");
                    msdyn_workorderservicetask workOrderServiceTask = target.ToEntity<msdyn_workorderservicetask>();
                    msdyn_workorderservicetask workOrderServiceTaskPreImage = preImageEntity.ToEntity<msdyn_workorderservicetask>();

                    tracingService.Trace("Only perform the updates if the work order service task is 100% complete on update or work order service task was already 100% complete (from pre-image)");
                    if (workOrderServiceTask.msdyn_PercentComplete == 100.00 || (workOrderServiceTask.msdyn_PercentComplete == null && workOrderServiceTaskPreImage.msdyn_PercentComplete == 100.00))
                    {
                        using (var serviceContext = new Xrm(service))
                        {
                            tracingService.Trace("Get the referenced work order from the preImage");
                            EntityReference workOrderReference = (EntityReference)preImageEntity.Attributes["msdyn_workorder"];

                            tracingService.Trace("Lookup the referenced work order");
                            msdyn_workorder workOrder = serviceContext.msdyn_workorderSet.Where(wo => wo.Id == workOrderReference.Id).FirstOrDefault();
                            // msdyn_workorder parentWorkOrder = serviceContext.msdyn_workorderSet.Where(pwo => pwo.Id == workOrder.msdyn_ParentWorkOrder.Id).FirstOrDefault();

                            tracingService.Trace("Determine if we use the questionnaire response from this update or from the pre-image since it is not always passed in the update");
                            var questionnaireResponse = !String.IsNullOrEmpty(workOrderServiceTask.ovs_QuestionnaireResponse) ? workOrderServiceTask.ovs_QuestionnaireResponse : workOrderServiceTaskPreImage.ovs_QuestionnaireResponse;
                            if (!String.IsNullOrWhiteSpace(questionnaireResponse))
                            {
                                tracingService.Trace("parse json response");
                                JsonValue jsonValue = JsonValue.Parse(questionnaireResponse);
                                JsonObject jsonObject = jsonValue as JsonObject;

                                tracingService.Trace("Retrieve all the findings belonging to this work order service task");
                                var findings = serviceContext.ovs_FindingSet.Where(f => f.ovs_WorkOrderServiceTaskId.Id == workOrderServiceTask.Id).ToList();

                                tracingService.Trace("List of findingType");
                                var findingTypeList = new List<string>();
                                tracingService.Trace("Start a list of all the used mapping keys");
                                var findingMappingKeys = new List<string>();

                                tracingService.Trace("If there was at least one finding found, - Create a case (if work order service task doesn't already belong to a case) - Mark the inspection result to fail");
                                if (jsonObject.Keys.Any(k => k.StartsWith("finding")))
                                {
                                    tracingService.Trace("If the work order is not null and is not already part of a case");
                                    if (workOrder != null && workOrder.msdyn_ServiceRequest == null)

                                    {
                                        Incident newIncident = new Incident();
                                        newIncident.ts_TradeNameId = workOrder.ts_tradenameId;
                                        newIncident.CustomerId = workOrder.msdyn_ServiceAccount;
                                        if (workOrder.ts_Site != null) newIncident.msdyn_FunctionalLocation = workOrder.ts_Site;

                                        tracingService.Trace("do environtment variable check");
                                        //string environmentVariableName = "ts_usenewregiontable";
                                        //string environmentVariableValue = EnvironmentVariableHelper.GetEnvironmentVariableValue(service, environmentVariableName);
                                        //if (environmentVariableValue == "yes")
                                        //{
                                        //    //ts_RegionDoNotUsez
                                        //    tracingService.Trace("if we are using the new field service table, do the if statement");
                                        //    if (workOrder.ts_RegionDoNotUse != null) newIncident.ovs_Region = workOrder.ts_RegionDoNotUse;
                                        //}
                                        //else
                                        //{
                                            if (workOrder.ts_Region != null) newIncident.ovs_Region = workOrder.ts_Region;
                                        //}

                                        
                                        if (workOrder.ts_Country != null) newIncident.ts_Country = workOrder.ts_Country;
                                        tracingService.Trace(" Stakeholder is a mandatory field on work order but, just in case, throw an error");
                                        if (workOrder.msdyn_ServiceAccount == null) throw new ArgumentNullException("msdyn_workorder.msdyn_ServiceAccount");

                                        Guid newIncidentId = service.Create(newIncident);

                                        tracingService.Trace("Update the Work Order with the Case");
                                        service.Update(new msdyn_workorder
                                        {
                                            Id = workOrderReference.Id,
                                            msdyn_ServiceRequest = new EntityReference(Incident.EntityLogicalName, newIncidentId)
                                        });

                                        tracingService.Trace("Set the Case ID for the Work Order Service Task");
                                        workOrderServiceTask.ovs_CaseId = new EntityReference(Incident.EntityLogicalName, newIncidentId);
                                    }
                                    else
                                    {
                                        tracingService.Trace("Already part of a case, just assign the work order case to the work order service task case");
                                        workOrderServiceTask.ovs_CaseId = workOrder.msdyn_ServiceRequest;
                                    }

                                    tracingService.Trace("loop through each root property in the json object");
                                    foreach (var rootProperty in jsonObject)
                                    {
                                        tracingService.Trace("Check if the root property starts with finding");
                                        if (rootProperty.Key.StartsWith("finding"))
                                        {
                                            var finding = rootProperty.Value;

                                            tracingService.Trace("The finding JSON may contain an array of string Id's of operation records");
                                            var operations = finding.ContainsKey("operations") ? finding["operations"] : new JsonArray();

                                            tracingService.Trace("retrieve the provision data containing the legislationID and provisionCategory");
                                            var provisionData = finding.ContainsKey("provisionData") ? finding["provisionData"] : new JsonObject();

                                            tracingService.Trace("Retrieve the provision name, might be stored in different keys");
                                            string provisionReferenceName = finding.ContainsKey("reference") && finding["reference"] != null && finding["reference"] != "" ? finding["reference"].ToString().Trim('"') : null;
                                            if (provisionReferenceName == null)
                                            {
                                                provisionReferenceName = finding.ContainsKey("provisionReference") && finding["provisionReference"] != null && finding["provisionReference"] != "" ? finding["provisionReference"].ToString().Trim('"') : null;
                                                if (provisionReferenceName == null)
                                                {
                                                    provisionReferenceName = finding.ContainsKey("provision") && finding["provision"] != null && finding["provision"] != "" ? finding["provision"].ToString().Trim('"') : null;
                                                }
                                            }

                                            tracingService.Trace("retrieve the provision category");
                                            Guid provisionCategoryId = provisionData != null && provisionData.ContainsKey("provisioncategoryid") && provisionData["provisioncategoryid"] != null ? Guid.Parse(provisionData["provisioncategoryid"]) : Guid.Empty;

                                            tracingService.Trace("Loop through the operations. Check if a finding already exists for that operation. Update the comment if it exists, or make a new finding if it doesn't");
                                            foreach (JsonObject operation in operations)
                                            {
                                                string operationid;
                                                tracingService.Trace("Grab operationid from Json operation object. If it wasn't populated for some reason, continue to next operation");
                                                if (operation.ContainsKey("operationID"))
                                                {
                                                    operationid = operation["operationID"];
                                                    findingTypeList.Add(operation["findingType"]);
                                                }
                                                else
                                                {
                                                    continue;
                                                }

                                                var findingMappingKey = workOrderServiceTask.Id.ToString() + "-" + rootProperty.Key.ToString() + "-" + operationid;
                                                findingMappingKeys.Add(findingMappingKey);
                                                var existingFinding = serviceContext.ovs_FindingSet.FirstOrDefault(f => f.ts_findingmappingkey == findingMappingKey);
                                                if (existingFinding == null)
                                                {
                                                    tracingService.Trace("if no, initialize new ovs_finding");
                                                    ovs_Finding newFinding = new ovs_Finding();
                                                    newFinding.ovs_FindingProvisionReference = finding.ContainsKey("provisionReference") ? (string)finding["provisionReference"] : "";
                                                    newFinding.ts_findingProvisionTextEn = finding.ContainsKey("provisionTextEn") ? (string)finding["provisionTextEn"] : "";
                                                    newFinding.ts_findingProvisionTextFr = finding.ContainsKey("provisionTextFr") ? (string)finding["provisionTextFr"] : "";
                                                    newFinding.ovs_FindingComments = finding.ContainsKey("comments") ? (string)finding["comments"] : "";
                                                    newFinding.ts_NotetoStakeholder = finding.ContainsKey("comments") ? (string)finding["comments"] : "";
                                                    OptionSetValue findingType = operation.ContainsKey("findingType") ? new OptionSetValue(operation["findingType"]) : new OptionSetValue(717750000); //717750000 is Undecided
                                                    newFinding.Attributes.Add("ts_findingtype", findingType);

                                                    tracingService.Trace("Update the list of findings for this service task in case a finding was added in a previous loop");
                                                    findings = serviceContext.ovs_FindingSet.Where(f => f.ovs_WorkOrderServiceTaskId.Id == workOrderServiceTask.Id).ToList();

                                                    tracingService.Trace("Find all the findings that share the same root property key and add to findingCopies list");
                                                    tracingService.Trace("The count of copies is needed to determine the suffix for the next finding record created");
                                                    var findingCopies = new List<ovs_Finding>();
                                                    foreach (var f in findings)
                                                    {
                                                        if (f.ts_findingmappingkey.StartsWith(workOrderServiceTask.Id.ToString() + "-" + rootProperty.Key.ToString())) findingCopies.Add(f);
                                                    }

                                                    tracingService.Trace("Determine the current highest infix of all the findings for the service task");
                                                    var highestInfix = 0;
                                                    foreach (ovs_Finding f in findings)
                                                    {
                                                        var currentInfix = Int32.Parse(f.ovs_Finding_1.Split('-')[3]);
                                                        if (currentInfix > highestInfix) highestInfix = currentInfix;
                                                    }

                                                    tracingService.Trace("Setup the finding name");
                                                    tracingService.Trace("Findings are at the 100 level");
                                                    var wostName = preImageEntity.Attributes["msdyn_name"].ToString();
                                                    var prefix = wostName.Replace("200-", "100-");
                                                    var infix = highestInfix + 1;
                                                    tracingService.Trace("If there are copies for this finding, use their infix instead");
                                                    if (findingCopies.Count > 0)
                                                    {
                                                        infix = Int32.Parse(findingCopies[0].ovs_Finding_1.Split('-')[3]);
                                                    }
                                                    var suffix = findingCopies.Count + 1;
                                                    newFinding.ovs_Finding_1 = string.Format("{0}-{1}-{2}", prefix, infix, suffix);

                                                    tracingService.Trace("Store the mapping key to keep track of mapping between finding and surveyjs questionnaire.");
                                                    newFinding.ts_findingmappingkey = findingMappingKey;

                                                    tracingService.Trace("reference work order service task");
                                                    newFinding.ovs_WorkOrderServiceTaskId = new EntityReference(msdyn_workorderservicetask.EntityLogicalName, workOrderServiceTask.Id);
                                                    newFinding.ts_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrder.Id);

                                                    tracingService.Trace("reference case (should already be saved in the work order service task)");


                                                     newFinding.ovs_CaseId = new EntityReference(Incident.EntityLogicalName, workOrderServiceTask.ovs_CaseId.Id);

                                                    EntityReference operationReference = new EntityReference(ovs_operation.EntityLogicalName, new Guid(operationid));

                                                    tracingService.Trace("Lookup the operation (Customer Asset Entity) to know its parent Account's id");
                                                    ovs_operation operationEntity = serviceContext.ovs_operationSet.Where(ca => ca.Id == operationReference.Id).FirstOrDefault();

                                                    tracingService.Trace("Create entity reference to the operation's parent account");
                                                    EntityReference parentAccountReference = new EntityReference(Account.EntityLogicalName, operationEntity.ts_stakeholder.Id);

                                                    tracingService.Trace("reference the current operation's stakeholder (Account Entity) (Lookup logical name: ts_stakeholder)");
                                                    newFinding.ts_accountid = new EntityReference(Account.EntityLogicalName, operationEntity.ts_stakeholder.Id);

                                                    tracingService.Trace("reference current operation (Customer Asset Entity) (Lookup logical name: ovs_asset)");
                                                    newFinding.ts_operationid = operationReference;

                                                    tracingService.Trace("reference operation type");
                                                    newFinding.ts_ovs_operationtype = operationEntity.ovs_OperationTypeId;

                                                    tracingService.Trace("reference site (functional location)");
                                                    newFinding.ts_functionallocation = operationEntity.ts_site;

                                                    tracingService.Trace("Operation Type is Person");
                                                    if (operationEntity.ovs_OperationTypeId != null && operationEntity.ovs_OperationTypeId.Id == new Guid("{BE8B0910-C751-EB11-A812-000D3AF3AC0D}")) //Operation Type is Person
                                                    {
                                                        if (workOrder.ts_Contact != null && workOrder.ts_Contact.Id != null)
                                                        {
                                                            newFinding.ts_Contact = new EntityReference(Contact.EntityLogicalName, workOrder.ts_Contact.Id);
                                                        }
                                                    }

                                                    if (provisionCategoryId != Guid.Empty)
                                                    {
                                                        ts_ProvisionCategory provisionCategory = serviceContext.ts_ProvisionCategorySet.Where(provCat => provCat.Id == provisionCategoryId).FirstOrDefault();
                                                        newFinding.ts_ProvisionCategory = new EntityReference(ts_ProvisionCategory.EntityLogicalName, provisionCategory.Id);
                                                    }

                                                    tracingService.Trace("reference legislation/provision");
                                                    qm_rclegislation legislation;

                                                    if (provisionReferenceName != null)
                                                    {
                                                        legislation = serviceContext.qm_rclegislationSet.Where(leg => (leg.ts_NameEnglish.Equals(provisionReferenceName) || leg.ts_NameFrench.Equals(provisionReferenceName) || leg.qm_name.Equals(provisionReferenceName))).FirstOrDefault();
                                                        if (legislation != null)
                                                        {
                                                            newFinding.ts_qm_rclegislation = new EntityReference(qm_rclegislation.EntityLogicalName, legislation.Id);
                                                        }
                                                    }

                                                    tracingService.Trace("Create new ovs_finding");
                                                    Guid newFindingId = service.Create(newFinding);

                                                }
                                                else
                                                {
                                                    tracingService.Trace("Update the Finding record's comment");
                                                    existingFinding.ovs_FindingComments = finding.ContainsKey("comments") ? (string)finding["comments"] : "";
                                                    tracingService.Trace("Update the Finding record's Finding Type");
                                                    OptionSetValue findingType = operation.ContainsKey("findingType") ? new OptionSetValue(operation["findingType"]) : new OptionSetValue(717750000); //717750000 is Undecided
                                                    existingFinding["ts_findingtype"] = findingType;
                                                    serviceContext.UpdateObject(existingFinding);
                                                }
                                            }
                                        }
                                    }
                                    tracingService.Trace("Mark the inspection result to Fail if there are non-compliance or Undecided found");
                                    tracingService.Trace("update documents for parent work order and case");
                                    if (findingTypeList.Contains("717750002") || findingTypeList.Contains("717750000"))
                                    {
                                        workOrderServiceTask.msdyn_inspectiontaskresult = msdyn_inspectionresult.Fail;
                                    }
                                    else
                                        workOrderServiceTask.msdyn_inspectiontaskresult = msdyn_inspectionresult.Observations;
                                }
                                else
                                {
                                    tracingService.Trace("Mark the inspection result to Pass if there are no findings found");
                                    workOrderServiceTask.msdyn_inspectiontaskresult = msdyn_inspectionresult.Pass;
                                }

                                tracingService.Trace("Need to deactivate any old referenced findings in the work order service task and case");
                                tracingService.Trace("that no longer exist in the questionnaire response.");

                                foreach (var finding in findings)
                                {
                                    tracingService.Trace("If the existing finding is not in the JSON response, we need to disable it");
                                    if (!findingMappingKeys.Contains(finding.ts_findingmappingkey))
                                    {
                                        finding.statuscode = ovs_Finding_statuscode.Inactive;
                                        finding.statecode = ovs_FindingState.Inactive;
                                    }
                                    else
                                    {
                                        tracingService.Trace("Otherwise, re-enable it");
                                        finding.statuscode = ovs_Finding_statuscode.New;
                                        finding.statecode = ovs_FindingState.Active;
                                    }
                                    serviceContext.UpdateObject(finding);
                                }
                            }

                            tracingService.Trace("If the work order is not already complete or closed and all other work order service tasks are already completed as well, mark the parent work order system status to Open - Completed");
                            var otherWorkOrderServiceTasks = serviceContext.msdyn_workorderservicetaskSet.Where(wost => wost.msdyn_WorkOrder == workOrderReference && wost.Id != workOrderServiceTask.Id).ToList<msdyn_workorderservicetask>();

                            // Took this out because we don't use the 'Completed' status for Work Orders
                            //if (workOrder.msdyn_SystemStatus != msdyn_wosystemstatus.Closed && workOrder.msdyn_SystemStatus != msdyn_wosystemstatus.Completed)
                            //{
                            //    if (otherWorkOrderServiceTasks.All(x => x.statuscode == msdyn_workorderservicetask_statuscode.Complete))
                            //    {
                            //        service.Update(new msdyn_workorder
                            //        {
                            //            Id = workOrderReference.Id,
                            //            msdyn_SystemStatus = msdyn_wosystemstatus.Completed
                            //        });
                            //    }
                            //}
                            tracingService.Trace("Save all the changes in the context as well.");
                            serviceContext.SaveChanges();

                            tracingService.Trace("Avoid updating the rollup field when in the mockup environment");
                            if (context.ParentContext == null || (context.ParentContext != null && context.ParentContext.OrganizationName != "MockupOrganization") && (workOrderServiceTask.msdyn_inspectiontaskresult == msdyn_inspectionresult.Fail || workOrderServiceTask.msdyn_inspectiontaskresult == msdyn_inspectionresult.Observations))
                            {
                                tracingService.Trace("Update Rollup Fields Number Of Findings for Work Order and Case");
                                CalculateRollupFieldRequest request;
                                CalculateRollupFieldResponse response;
                                tracingService.Trace("Update Rollup field Number Of Findings for Work Order");
                                request = new CalculateRollupFieldRequest
                                {
                                    Target = new EntityReference("msdyn_workorder", workOrder.Id),
                                    FieldName = "ts_numberoffindings" // Rollup Field Name
                                };

                                response = (CalculateRollupFieldResponse)service.Execute(request);

                                tracingService.Trace("Update Rollup field Number Of Findings for Case");
                                request = new CalculateRollupFieldRequest
                                {
                                    Target = new EntityReference("incident", workOrderServiceTask.ovs_CaseId.Id),
                                    FieldName = "ts_numberoffindings" // Rollup Field Name
                                };
                                response = (CalculateRollupFieldResponse)service.Execute(request);
                            }
                        }
                    }
                }

                // Seems to be a bug if exception variables have the same name.
                // Make sure the name of each exception variable is different.
                // This will show the message in a dialog box for the user
                catch (FaultException<OrganizationServiceFault> orgServiceEx)
                {
                    tracingService.Trace($"MESSAGE: { orgServiceEx.Message}");
                    tracingService.Trace($"CODE: {orgServiceEx.Code}");
                    tracingService.Trace($"DETAIL: {orgServiceEx.Detail}");
                    tracingService.Trace($"INNER FAULT: {orgServiceEx.Detail?.InnerFault}");
                    tracingService.Trace($"TRACE: {orgServiceEx.Detail?.TraceText}");

                    throw new InvalidPluginExecutionException("An error occurred in PreOperationmsdyn_workorderservicetaskUpdate Plugin.", orgServiceEx);
                }

                catch (Exception ex)
                {
                    tracingService.Trace("PreOperationmsdyn_workorderservicetaskUpdate Plugin: {0}", ex.ToString());
                    tracingService.Trace($"MESSAGE: {ex.Message}");
                    tracingService.Trace($"STACK TRACE: {ex.StackTrace}");
                    tracingService.Trace($"INNER EXCEPTION: {ex.InnerException?.Message}");
                    tracingService.Trace($"SOURCE: {ex.Source}");

                    throw;
                }
            }
        }
    }
}
