using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using DG.XrmContext;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json.Linq;

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
    public class PreOperationmsdyn_workorderservicetaskUpdate : PluginBase
    {
        public PreOperationmsdyn_workorderservicetaskUpdate(string unsecure, string secure)
            : base(typeof(PreOperationmsdyn_workorderservicetaskUpdate))
        {
        }

        protected override void ExecuteCrmPlugin(LocalPluginContext localContext)
        {
            if (localContext == null)
            {
                throw new InvalidPluginExecutionException("localContext");
            }

            var tracingService = localContext.TracingService;
            var context = localContext.PluginExecutionContext;
            var service = localContext.OrganizationService;

            localContext.Trace("Tracking Service Started.");

            localContext.Trace("Return if triggered by another plugin. Prevents infinite loop.");
            if (context.Depth > 2)
                return;

            localContext.Trace("The InputParameters collection contains all the data passed in the message request.");
            if (context.InputParameters.Contains("Target") &&
                context.InputParameters["Target"] is Entity)
            {
                localContext.Trace("Obtain the target entity from the input parameters.");
                Entity target = (Entity)context.InputParameters["Target"];

                localContext.Trace("Obtain the preimage entity");
                Entity preImageEntity = (context.PreEntityImages != null && context.PreEntityImages.Contains("PreImage")) ? context.PreEntityImages["PreImage"] : null;

                try
                {
                    // Log the system username and Work Order Service Task at the start
                    var systemUser = service.Retrieve("systemuser", context.InitiatingUserId, new ColumnSet("fullname"));
                    var WOST = service.Retrieve("msdyn_workorderservicetask", context.PrimaryEntityId, new ColumnSet("msdyn_name"));
                    localContext.Trace("Plugin executed by user: {0}", systemUser.GetAttributeValue<string>("fullname"));
                    localContext.Trace("Work Order Service Task GUID: {0}", context.PrimaryEntityId);
                    localContext.Trace("Work Order Service Task Name: {0}", WOST.GetAttributeValue<string>("msdyn_name"));

                    localContext.Trace("Cast the target and preimage to the expected entity");
                    msdyn_workorderservicetask workOrderServiceTask = target.ToEntity<msdyn_workorderservicetask>();
                    msdyn_workorderservicetask workOrderServiceTaskPreImage = preImageEntity.ToEntity<msdyn_workorderservicetask>();

                    localContext.Trace("Only perform the updates if the work order service task is 100% complete on update or work order service task was already 100% complete (from pre-image)");
                    if (workOrderServiceTask.msdyn_PercentComplete == 100.00 || (workOrderServiceTask.msdyn_PercentComplete == null && workOrderServiceTaskPreImage.msdyn_PercentComplete == 100.00))
                    {
                        using (var serviceContext = new Xrm(service))
                        {
                            localContext.Trace("Get the referenced work order from the preImage");
                            EntityReference workOrderReference = (EntityReference)preImageEntity.Attributes["msdyn_workorder"];
                            // Retrieve the msdyn_workorder record to get the msdyn_name
                            var workOrderName = service.Retrieve("msdyn_workorder", workOrderReference.Id, new ColumnSet("msdyn_name"));
                            localContext.Trace("Work Order GUID: {0}", workOrderReference.Id);
                            localContext.Trace("Work Order Name: {0}", workOrderName.GetAttributeValue<string>("msdyn_name"));

                            localContext.Trace("Lookup the referenced work order");
                            msdyn_workorder workOrder = serviceContext.msdyn_workorderSet.Where(wo => wo.Id == workOrderReference.Id).FirstOrDefault();
                            // msdyn_workorder parentWorkOrder = serviceContext.msdyn_workorderSet.Where(pwo => pwo.Id == workOrder.msdyn_ParentWorkOrder.Id).FirstOrDefault();

                            localContext.Trace("Determine if we use the questionnaire response from this update or from the pre-image since it is not always passed in the update");
                            var questionnaireResponse = !String.IsNullOrEmpty(workOrderServiceTask.ovs_QuestionnaireResponse) ? workOrderServiceTask.ovs_QuestionnaireResponse : workOrderServiceTaskPreImage.ovs_QuestionnaireResponse;
                            if (!String.IsNullOrWhiteSpace(questionnaireResponse))
                            {
                                localContext.Trace("parse json response");
                                JObject jsonObject = JObject.Parse(questionnaireResponse);

                                localContext.Trace("Retrieve all the findings belonging to this work order service task");
                                var findings = serviceContext.ovs_FindingSet.Where(f => f.ovs_WorkOrderServiceTaskId.Id == workOrderServiceTask.Id).ToList();

                                localContext.Trace("List of findingType");
                                var findingTypeList = new List<string>();
                                localContext.Trace("Start a list of all the used mapping keys");
                                var findingMappingKeys = new List<string>();

                                localContext.Trace("If there was at least one finding found, - Create a case (if work order service task doesn't already belong to a case) - Mark the inspection result to fail");
                                if (jsonObject.Properties().Any(p => p.Name.StartsWith("finding")))
                                {
                                    localContext.Trace("If the work order is not null and is not already part of a case");
                                    if (workOrder != null && workOrder.msdyn_ServiceRequest == null)

                                    {
                                        Incident newIncident = new Incident();
                                        newIncident.ts_TradeNameId = workOrder.ts_tradenameId;
                                        newIncident.CustomerId = workOrder.msdyn_ServiceAccount;
                                        if (workOrder.ts_Site != null) newIncident.msdyn_FunctionalLocation = workOrder.ts_Site;

                                        localContext.Trace("do environtment variable check");
                                        //string environmentVariableName = "ts_usenewregiontable";
                                        //string environmentVariableValue = OrganizationConfig.GetEnvironmentVariableValue(service, environmentVariableName);
                                        //if (environmentVariableValue == "yes")
                                        //{
                                        //    //ts_RegionDoNotUsez
                                        //    localContext.Trace("if we are using the new field service table, do the if statement");
                                        //    if (workOrder.ts_RegionDoNotUse != null) newIncident.ovs_Region = workOrder.ts_RegionDoNotUse;
                                        //}
                                        //else
                                        //{
                                        if (workOrder.ts_Region != null) newIncident.ovs_Region = workOrder.ts_Region;
                                        //}


                                        if (workOrder.ts_Country != null) newIncident.ts_Country = workOrder.ts_Country;
                                        localContext.Trace(" Stakeholder is a mandatory field on work order but, just in case, throw an error");
                                        if (workOrder.msdyn_ServiceAccount == null) throw new ArgumentNullException("msdyn_workorder.msdyn_ServiceAccount");

                                        // Determine owner from the related Work Order Service Task Workspace (if any)
                                        var workspace = serviceContext.ts_WorkOrderServiceTaskWorkspaceSet.FirstOrDefault(ws => ws.ts_WorkOrderServiceTask != null && ws.ts_WorkOrderServiceTask.Id == workOrderServiceTask.Id);

                                        if (workspace != null && workspace.ModifiedBy != null)
                                        {
                                            // Owner can be a user or a team
                                            newIncident.OwnerId = workspace.ModifiedBy;
                                        }
                                        Guid newIncidentId = service.Create(newIncident);

                                        localContext.Trace("Update the Work Order with the Case");
                                        service.Update(new msdyn_workorder
                                        {
                                            Id = workOrderReference.Id,
                                            msdyn_ServiceRequest = new EntityReference(Incident.EntityLogicalName, newIncidentId)
                                        });

                                        localContext.Trace("Set the Case ID for the Work Order Service Task");
                                        workOrderServiceTask.ovs_CaseId = new EntityReference(Incident.EntityLogicalName, newIncidentId);
                                        // After setting workOrderServiceTask.ovs_CaseId
                                        if (workspace != null && workOrderServiceTask.ovs_CaseId != null && (workspace.crc77_Incident == null || workspace.crc77_Incident.Id != workOrderServiceTask.ovs_CaseId.Id))
                                        {
                                            service.Update(new ts_WorkOrderServiceTaskWorkspace
                                            {
                                                Id = workspace.Id,
                                                crc77_Incident = workOrderServiceTask.ovs_CaseId
                                            });
                                        }
                                    }
                                    else
                                    {
                                        localContext.Trace("Already part of a case, just assign the work order case to the work order service task case");
                                        workOrderServiceTask.ovs_CaseId = workOrder.msdyn_ServiceRequest;
                                        // After setting workOrderServiceTask.ovs_CaseId
                                        var workspace = serviceContext.ts_WorkOrderServiceTaskWorkspaceSet.FirstOrDefault(ws => ws.ts_WorkOrderServiceTask != null && ws.ts_WorkOrderServiceTask.Id == workOrderServiceTask.Id);

                                        if (workspace != null && workOrder.msdyn_ServiceRequest != null && (workspace.crc77_Incident == null || workspace.crc77_Incident.Id != workOrder.msdyn_ServiceRequest.Id))
                                        {
                                            service.Update(new ts_WorkOrderServiceTaskWorkspace
                                            {
                                                Id = workspace.Id,
                                                crc77_Incident = workOrder.msdyn_ServiceRequest
                                            });
                                        }
                                    }

                                    localContext.Trace("loop through each root property in the json object");
                                    foreach (var rootProperty in jsonObject.Properties())
                                    {
                                        localContext.Trace("Check if the root property starts with finding");
                                        if (rootProperty.Name.StartsWith("finding"))
                                        {
                                            var finding = rootProperty.Value as JObject;

                                            localContext.Trace("The finding JSON may contain an array of string Id's of operation records");
                                            var operations = finding["operations"] as JArray ?? new JArray();

                                            localContext.Trace("retrieve the provision data containing the legislationID and provisionCategory");
                                            var provisionData = finding["provisionData"] as JObject ?? new JObject();

                                            localContext.Trace("Retrieve the provision name, might be stored in different keys");
                                            string provisionReferenceName = finding["reference"]?.ToString();
                                            if (string.IsNullOrEmpty(provisionReferenceName))
                                            {
                                                provisionReferenceName = finding["provisionReference"]?.ToString();
                                                if (string.IsNullOrEmpty(provisionReferenceName))
                                                {
                                                    provisionReferenceName = finding["provision"]?.ToString();
                                                }
                                            }

                                            localContext.Trace("retrieve the provision category");
                                            Guid provisionCategoryId = Guid.Empty;
                                            var provisionCategoryIdToken = provisionData?["provisioncategoryid"];
                                            if (provisionCategoryIdToken != null && !string.IsNullOrEmpty(provisionCategoryIdToken.ToString()))
                                            {
                                                Guid.TryParse(provisionCategoryIdToken.ToString(), out provisionCategoryId);
                                            }

                                            localContext.Trace("Loop through the operations. Check if a finding already exists for that operation. Update the comment if it exists, or make a new finding if it doesn't");
                                            foreach (JObject operation in operations)
                                            {
                                                string operationid;
                                                localContext.Trace("Grab operationid from Json operation object. If it wasn't populated for some reason, continue to next operation");
                                                if (operation["operationID"] != null)
                                                {
                                                    operationid = operation["operationID"].ToString();
                                                    findingTypeList.Add(operation["findingType"]?.ToString());
                                                }
                                                else
                                                {
                                                    continue;
                                                }

                                                var findingMappingKey = workOrderServiceTask.Id.ToString() + "-" + rootProperty.Name + "-" + operationid;
                                                findingMappingKeys.Add(findingMappingKey);
                                                var existingFinding = serviceContext.ovs_FindingSet.FirstOrDefault(f => f.ts_findingmappingkey == findingMappingKey);
                                                if (existingFinding == null)
                                                {
                                                    localContext.Trace("if no, initialize new ovs_finding");
                                                    ovs_Finding newFinding = new ovs_Finding();
                                                    newFinding.ovs_FindingProvisionReference = finding["provisionReference"]?.ToString() ?? "";
                                                    newFinding.ts_findingProvisionTextEn = finding["provisionTextEn"]?.ToString() ?? "";
                                                    newFinding.ts_findingProvisionTextFr = finding["provisionTextFr"]?.ToString() ?? "";
                                                    newFinding.ovs_FindingComments = finding["comments"]?.ToString() ?? "";
                                                    newFinding.ts_NotetoStakeholder = finding["comments"]?.ToString() ?? "";
                                                    OptionSetValue findingType = operation["findingType"] != null ? new OptionSetValue((int)operation["findingType"]) : new OptionSetValue(717750000); //717750000 is Undecided
                                                    newFinding.Attributes.Add("ts_findingtype", findingType);

                                                    localContext.Trace("Update the list of findings for this service task in case a finding was added in a previous loop");
                                                    findings = serviceContext.ovs_FindingSet.Where(f => f.ovs_WorkOrderServiceTaskId.Id == workOrderServiceTask.Id).ToList();

                                                    localContext.Trace("Find all the findings that share the same root property key and add to findingCopies list");
                                                    localContext.Trace("The count of copies is needed to determine the suffix for the next finding record created");
                                                    var findingCopies = new List<ovs_Finding>();
                                                    foreach (var f in findings)
                                                    {
                                                        if (f.ts_findingmappingkey.StartsWith(workOrderServiceTask.Id.ToString() + "-" + rootProperty.Name + "-")) findingCopies.Add(f);
                                                    }

                                                    localContext.Trace("Determine the current highest infix of all the findings for the service task");
                                                    var highestInfix = 0;
                                                    foreach (ovs_Finding f in findings)
                                                    {
                                                        var currentInfix = Int32.Parse(f.ovs_Finding_1.Split('-')[3]);
                                                        if (currentInfix > highestInfix) highestInfix = currentInfix;
                                                    }

                                                    localContext.Trace("Setup the finding name");
                                                    localContext.Trace("Findings are at the 100 level");
                                                    var wostName = preImageEntity.Attributes["msdyn_name"].ToString();
                                                    var prefix = wostName.Replace("200-", "100-");
                                                    var infix = highestInfix + 1;
                                                    localContext.Trace("If there are copies for this finding, use their infix instead");
                                                    if (findingCopies.Count > 0)
                                                    {
                                                        infix = Int32.Parse(findingCopies[0].ovs_Finding_1.Split('-')[3]);
                                                    }
                                                    var suffix = findingCopies.Count + 1;
                                                    newFinding.ovs_Finding_1 = string.Format("{0}-{1}-{2}", prefix, infix, suffix);

                                                    localContext.Trace("Store the mapping key to keep track of mapping between finding and surveyjs questionnaire.");
                                                    newFinding.ts_findingmappingkey = findingMappingKey;

                                                    localContext.Trace("reference work order service task");
                                                    newFinding.ovs_WorkOrderServiceTaskId = new EntityReference(msdyn_workorderservicetask.EntityLogicalName, workOrderServiceTask.Id);
                                                    newFinding.ts_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrder.Id);

                                                    localContext.Trace("reference case (should already be saved in the work order service task)");


                                                    newFinding.ovs_CaseId = new EntityReference(Incident.EntityLogicalName, workOrderServiceTask.ovs_CaseId.Id);

                                                    EntityReference operationReference = new EntityReference(ovs_operation.EntityLogicalName, new Guid(operationid));

                                                    localContext.Trace("Lookup the operation (Customer Asset Entity) to know its parent Account's id");
                                                    ovs_operation operationEntity = serviceContext.ovs_operationSet.Where(ca => ca.Id == operationReference.Id).FirstOrDefault();

                                                    localContext.Trace("Create entity reference to the operation's parent account");
                                                    EntityReference parentAccountReference = new EntityReference(Account.EntityLogicalName, operationEntity.ts_stakeholder.Id);

                                                    localContext.Trace("reference the current operation's stakeholder (Account Entity) (Lookup logical name: ts_stakeholder)");
                                                    newFinding.ts_accountid = new EntityReference(Account.EntityLogicalName, operationEntity.ts_stakeholder.Id);

                                                    localContext.Trace("reference current operation (Customer Asset Entity) (Lookup logical name: ovs_asset)");
                                                    newFinding.ts_operationid = operationReference;

                                                    localContext.Trace("reference operation type");
                                                    newFinding.ts_ovs_operationtype = operationEntity.ovs_OperationTypeId;

                                                    localContext.Trace("reference site (functional location)");
                                                    newFinding.ts_functionallocation = operationEntity.ts_site;

                                                    localContext.Trace("Operation Type is Person");
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

                                                    localContext.Trace("reference legislation/provision");
                                                    qm_rclegislation legislation;

                                                    if (provisionReferenceName != null)
                                                    {
                                                        legislation = serviceContext.qm_rclegislationSet.Where(leg => (leg.ts_NameEnglish.Equals(provisionReferenceName) || leg.ts_NameFrench.Equals(provisionReferenceName) || leg.qm_name.Equals(provisionReferenceName))).FirstOrDefault();
                                                        if (legislation != null)
                                                        {
                                                            newFinding.ts_qm_rclegislation = new EntityReference(qm_rclegislation.EntityLogicalName, legislation.Id);
                                                        }
                                                    }

                                                    // Determine owner from the related Work Order Service Task Workspace (if any)
                                                    var workspace = serviceContext.ts_WorkOrderServiceTaskWorkspaceSet.FirstOrDefault(ws => ws.ts_WorkOrderServiceTask != null && ws.ts_WorkOrderServiceTask.Id == workOrderServiceTask.Id);

                                                    if (workspace != null && workspace.ModifiedBy != null)
                                                    {
                                                        // Owner can be a user or a team
                                                        newFinding.OwnerId = workspace.ModifiedBy;
                                                    }
                                                    localContext.Trace("Create new ovs_finding");
                                                    Guid newFindingId = service.Create(newFinding);

                                                }
                                                else
                                                {
                                                    localContext.Trace("Update the Finding record's comment");
                                                    existingFinding.ovs_FindingComments = finding["comments"]?.ToString() ?? "";
                                                    localContext.Trace("Update the Finding record's Finding Type");
                                                    OptionSetValue findingType = operation["findingType"] != null ? new OptionSetValue((int)operation["findingType"]) : new OptionSetValue(717750000); //717750000 is Undecided
                                                    existingFinding["ts_findingtype"] = findingType;
                                                    serviceContext.UpdateObject(existingFinding);
                                                }
                                            }
                                        }
                                    }
                                    localContext.Trace("Mark the inspection result to Fail if there are non-compliance or Undecided found");
                                    localContext.Trace("update documents for parent work order and case");
                                    if (findingTypeList.Contains("717750002") || findingTypeList.Contains("717750000"))
                                    {
                                        workOrderServiceTask.msdyn_inspectiontaskresult = msdyn_inspectionresult.Fail;
                                    }
                                    else
                                        workOrderServiceTask.msdyn_inspectiontaskresult = msdyn_inspectionresult.Observations;
                                }
                                else
                                {
                                    localContext.Trace("Mark the inspection result to Pass if there are no findings found");
                                    workOrderServiceTask.msdyn_inspectiontaskresult = msdyn_inspectionresult.Pass;
                                }

                                localContext.Trace("Need to deactivate any old referenced findings in the work order service task and case");
                                localContext.Trace("that no longer exist in the questionnaire response.");

                                foreach (var finding in findings)
                                {
                                    localContext.Trace("If the existing finding is not in the JSON response, we need to disable it");
                                    if (!findingMappingKeys.Contains(finding.ts_findingmappingkey))
                                    {
                                        finding.statuscode = ovs_Finding_statuscode.Inactive;
                                        finding.statecode = ovs_FindingState.Inactive;
                                    }
                                    else
                                    {
                                        localContext.Trace("Otherwise, re-enable it");
                                        finding.statuscode = ovs_Finding_statuscode.New;
                                        finding.statecode = ovs_FindingState.Active;
                                    }
                                    serviceContext.UpdateObject(finding);
                                }
                            }

                            localContext.Trace("If the work order is not already complete or closed and all other work order service tasks are already completed as well, mark the parent work order system status to Open - Completed");
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
                            localContext.Trace("Save all the changes in the context as well.");
                            serviceContext.SaveChanges();

                            localContext.Trace("Avoid updating the rollup field when in the mockup environment");
                            if (context.ParentContext == null || (context.ParentContext != null && context.ParentContext.OrganizationName != "MockupOrganization") && (workOrderServiceTask.msdyn_inspectiontaskresult == msdyn_inspectionresult.Fail || workOrderServiceTask.msdyn_inspectiontaskresult == msdyn_inspectionresult.Observations))
                            {
                                localContext.Trace("Update Rollup Fields Number Of Findings for Work Order and Case");
                                CalculateRollupFieldRequest request;
                                CalculateRollupFieldResponse response;
                                localContext.Trace("Update Rollup field Number Of Findings for Work Order");
                                request = new CalculateRollupFieldRequest
                                {
                                    Target = new EntityReference("msdyn_workorder", workOrder.Id),
                                    FieldName = "ts_numberoffindings" // Rollup Field Name
                                };

                                response = (CalculateRollupFieldResponse)service.Execute(request);

                                localContext.Trace("Update Rollup field Number Of Findings for Case");
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
                    localContext.Trace($"MESSAGE: {orgServiceEx.Message}");
                    localContext.Trace($"CODE: {orgServiceEx.Code}");
                    localContext.Trace($"DETAIL: {orgServiceEx.Detail}");
                    localContext.Trace($"INNER FAULT: {orgServiceEx.Detail?.InnerFault}");
                    localContext.Trace($"TRACE: {orgServiceEx.Detail?.TraceText}");

                    throw new InvalidPluginExecutionException("An error occurred in PreOperationmsdyn_workorderservicetaskUpdate Plugin.", orgServiceEx);
                }

                catch (Exception ex)
                {
                    localContext.Trace("PreOperationmsdyn_workorderservicetaskUpdate Plugin: {0}", ex);
                    throw new InvalidPluginExecutionException("PreOperationmsdyn_workorderservicetaskUpdate failed.", ex);
                }
            }
        }
    }
}
