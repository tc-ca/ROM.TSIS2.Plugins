using System;
using System.Linq;
using Microsoft.Xrm.Sdk;

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
    public class PreOperationincidentUpdate : IPlugin
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

                // Obtain the organization service reference which you will need for
                // web service calls.
                IOrganizationServiceFactory serviceFactory =
                    (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

                try
                {
                    if (target.LogicalName.Equals(Incident.EntityLogicalName))
                    {
                        int UserLanguage = LocalizationHelper.RetrieveUserUILanguageCode(service, context.InitiatingUserId);
                        string ResourceFile = "ovs_/resx/Incident.1033.resx";
                        if (UserLanguage == 1036) //French
                        {
                            ResourceFile = "ovs_/resx/Incident.1036.resx";
                        }
                        if (target.Attributes.Contains("statecode") && (int)(target.GetAttributeValue<OptionSetValue>("statecode").Value) == 1)
                        {
                            using (var servicecontext = new Xrm(service))
                            {                              
                                var findingEntities = servicecontext.ovs_FindingSet.Where(f => f.ovs_CaseId.Id == target.Id && f.ts_findingtype == ts_findingtype.Noncompliance && f.statuscode != ovs_Finding_statuscode.Complete).ToList();
                                if (findingEntities != null && findingEntities.Count>0)
                                {
                                    tracingService.Trace("Non-Completed Findings count {0}", findingEntities.Count);
                                    throw new InvalidPluginExecutionException(LocalizationHelper.GetMessage(tracingService, service, ResourceFile, "ClosingCaseErrorMsg"));
                                }
                            }
                        }

                        if (target.Attributes.Contains("ownerid") && target.GetAttributeValue<EntityReference>("ownerid").Id  != context.InitiatingUserId)
                        {
                            tracingService.Trace("Ownerid is changing to {0} by {1} ", target.GetAttributeValue<EntityReference>("ownerid").Id, context.InitiatingUserId);
                            using (var servicecontext = new Xrm(service))
                            {
                                var currentUser = servicecontext.SystemUserSet.Where(u => u.Id == context.InitiatingUserId).FirstOrDefault();
                                var updatedOwnerUser = servicecontext.SystemUserSet.Where(u => u.Id == target.GetAttributeValue<EntityReference>("ownerid").Id).FirstOrDefault();

                                if (updatedOwnerUser != null)
                                {
                                    if (currentUser.GetAttributeValue<EntityReference>("businessunitid").Id != updatedOwnerUser.GetAttributeValue<EntityReference>("businessunitid").Id &&
                                    !currentUser.GetAttributeValue<EntityReference>("businessunitid").Name.Equals("Transport Canada", StringComparison.OrdinalIgnoreCase))
                                    {
                                        throw new InvalidPluginExecutionException(LocalizationHelper.GetMessage(tracingService, service, ResourceFile, "ReassignCaseErrorMsg"));
                                    }
                                }
                                else
                                {
                                    var updatedOwnerTeam = servicecontext.TeamSet.Where(u => u.Id == target.GetAttributeValue<EntityReference>("ownerid").Id).FirstOrDefault();
                                    if (updatedOwnerTeam != null && currentUser.GetAttributeValue<EntityReference>("businessunitid").Id != updatedOwnerTeam.GetAttributeValue<EntityReference>("businessunitid").Id &&
                                    !currentUser.GetAttributeValue<EntityReference>("businessunitid").Name.Equals("Transport Canada", StringComparison.OrdinalIgnoreCase))
                                    {
                                        throw new InvalidPluginExecutionException(LocalizationHelper.GetMessage(tracingService, service, ResourceFile, "ReassignCaseErrorMsg"));
                                    }
                                }

                            }
                        }
                    }
                }
                catch (Exception e)
                {
                    throw new InvalidPluginExecutionException(e.Message);
                }

            }
        }
    }
}