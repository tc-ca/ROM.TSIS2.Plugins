using System;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace TSIS2.Plugins
{

    [CrmPluginRegistration(
    MessageNameEnum.Update,
    "incident",
    StageEnum.PreOperation,
    ExecutionModeEnum.Synchronous,
    "",
    "TSIS2.Plugins.PreOperationincidentUpdate Plugin",
    1,
    IsolationModeEnum.Sandbox,
    Description = "On Closing Case, validate if there is any Finding is not set to complete status")]
    /// <summary>
    /// PreOperationincidentUpdate Plugin.
    /// </summary>    
    public class PreOperationincidentUpdate : PluginBase
    {
        public PreOperationincidentUpdate() : base(typeof(PreOperationincidentUpdate))
        {
        }

        protected override void ExecuteCrmPlugin(LocalPluginContext localContext)
        {
            var tracingService = localContext.TracingService;
            var context = localContext.PluginExecutionContext;
            var service = localContext.OrganizationService;

            // The InputParameters collection contains all the data passed in the message request.
            if (context.InputParameters.Contains("Target") &&
                context.InputParameters["Target"] is Entity)
            {
                // Obtain the target entity from the input parameters.
                Entity target = (Entity)context.InputParameters["Target"];

                // Log the system username and Case ID at the start
                var systemUser = service.Retrieve("systemuser", context.InitiatingUserId, new ColumnSet("fullname"));
                // Retrieve the ticketnumber from the database
                Entity incident = service.Retrieve("incident", context.PrimaryEntityId, new ColumnSet("ticketnumber"));
                
                localContext.Trace("Success obtaining the system username and case (incident) id");

                try
                {
                    localContext.Trace("Plugin executed by user: {0}", systemUser.GetAttributeValue<string>("fullname"));
                    localContext.Trace("Case GUID: {0}", context.PrimaryEntityId);  
                    localContext.Trace("Case Number: {0}", incident.GetAttributeValue<string>("ticketnumber"));

                    if (target.LogicalName.Equals(Incident.EntityLogicalName))
                    {
                        int UserLanguage = LocalizationHelper.RetrieveUserUILanguageCode(service, context.InitiatingUserId);
                        string ResourceFile = "ovs_/resx/Incident.1033.resx";
                        if (UserLanguage == 1036) //French
                        {
                            ResourceFile = "ovs_/resx/Incident.1036.resx";
                            localContext.Trace("Load French resource file");
                        }
                        if (target.Attributes.Contains("statecode") && (int)(target.GetAttributeValue<OptionSetValue>("statecode").Value) == 1)
                        {
                            localContext.Trace("If the status is set to Resolved (1)");
                            using (var servicecontext = new Xrm(service))
                            {                              
                                var findingEntities = servicecontext.ovs_FindingSet.Where(f => f.ovs_CaseId.Id == target.Id && f.ts_findingtype == ts_findingtype.Noncompliance && f.statuscode != ovs_Finding_statuscode.Complete).ToList();
                                if (findingEntities != null && findingEntities.Count>0)
                                {
                                    localContext.Trace("Non-Completed Findings count {0}", findingEntities.Count);
                                    throw new InvalidPluginExecutionException(LocalizationHelper.GetMessage(tracingService, service, ResourceFile, "ClosingCaseErrorMsg"));
                                }
                            }
                        }

                        if (target.Attributes.Contains("ownerid") && target.GetAttributeValue<EntityReference>("ownerid").Id  != context.InitiatingUserId)
                        {
                            EntityReference ownerRef = target.GetAttributeValue<EntityReference>("ownerid");
                            localContext.Trace("Retrieve the owner EntityReference");
                            string ownerNameAttribute = ownerRef.LogicalName == "systemuser" ? "fullname" : "name";
                            Entity owner = service.Retrieve(ownerRef.LogicalName, ownerRef.Id, new ColumnSet(ownerNameAttribute));
                            localContext.Trace("Retrieve the owner's name using IOrganizationService");
                            string ownerName = owner.GetAttributeValue<string>(ownerNameAttribute);
                            localContext.Trace("Ownerid is changing to {0} by {1} ", ownerName, systemUser.GetAttributeValue<string>("fullname"));
                            using (var servicecontext = new Xrm(service))
                            {
                                var currentUser = servicecontext.SystemUserSet.Where(u => u.Id == context.InitiatingUserId).FirstOrDefault();
                                var currentUserBUId = currentUser.GetAttributeValue<EntityReference>("businessunitid").Id;
                                var updatedOwnerUser = servicecontext.SystemUserSet.Where(u => u.Id == target.GetAttributeValue<EntityReference>("ownerid").Id).FirstOrDefault();
                                localContext.Trace("Initialize variables: currentUser, updatedOwnerUser");

                                if (updatedOwnerUser != null)
                                {
                                    localContext.Trace("If updatedOwnerUser is not null then execute");
                                    var updatedOwnerUserBUId = updatedOwnerUser.GetAttributeValue<EntityReference>("businessunitid").Id;

                                    if (OrganizationConfig.IsAvSecBU(service, currentUserBUId, tracingService))
                                    {
                                        localContext.Trace("If business unit is AvSec then execute");
                                        if (!OrganizationConfig.IsAvSecBU(service, updatedOwnerUserBUId, tracingService) ||
                                        OrganizationConfig.IsAvSecPPPBU(service, updatedOwnerUserBUId, tracingService))
                                        {
                                            localContext.Trace("Reassign case error due to business unit mismatch1.");
                                            throw new InvalidPluginExecutionException(LocalizationHelper.GetMessage(tracingService, service, ResourceFile, "ReassignCaseErrorMsg"));
                                        }
                                    }
                                    else
                                    {
                                        localContext.Trace("Else, start check to see if currentUser businessunitid is not equal to updateOwnerUser businessunitid AND currentUser businessunitid is Transport Canada");
                                        if (currentUserBUId != updatedOwnerUserBUId &&
                                        !OrganizationConfig.IsTCBU(service, currentUserBUId, tracingService))
                                        {
                                            localContext.Trace("Reassign case error due to business unit mismatch2.");
                                            throw new InvalidPluginExecutionException(LocalizationHelper.GetMessage(tracingService, service, ResourceFile, "ReassignCaseErrorMsg"));
                                        }
                                    }
                                }
                                else
                                {
                                    var updatedOwnerTeam = servicecontext.TeamSet.Where(u => u.Id == target.GetAttributeValue<EntityReference>("ownerid").Id).FirstOrDefault();
                                    localContext.Trace("Initialize variables: updatedOwnerTeam");
                                    if (updatedOwnerTeam != null)
                                    {
                                        var updatedOwnerTeamBUId = updatedOwnerTeam.GetAttributeValue<EntityReference>("businessunitid").Id;

                                        if (currentUserBUId != updatedOwnerTeamBUId &&
                                        !OrganizationConfig.IsTCBU(service, currentUserBUId, tracingService))
                                        {
                                            localContext.Trace("Reassign case error due to team mismatch.");
                                            throw new InvalidPluginExecutionException(LocalizationHelper.GetMessage(tracingService, service, ResourceFile, "ReassignCaseErrorMsg"));
                                        }
                                    }
                                }

                            }
                        }
                    }
                }
                catch (Exception e)
                {
                    // Log detailed error information
                    localContext.Trace("Error occurred: {0}, Case GUID: {1}, Ticket Number: {2}, User: {3}",
                        e.Message,
                        context.PrimaryEntityId,
                        incident.GetAttributeValue<string>("ticketnumber"),
                        systemUser.GetAttributeValue<string>("fullname"));

                    throw new InvalidPluginExecutionException(e.Message);
                }

            }
        }
    }
}