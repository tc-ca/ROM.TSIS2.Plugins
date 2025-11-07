using System;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System.Xml.Linq;

namespace TSIS2.Plugins
{

    [CrmPluginRegistration(
    MessageNameEnum.Create,
    "ovs_operation",
    StageEnum.PreValidation,
    ExecutionModeEnum.Synchronous,
    "",
    "TSIS2.Plugins.PreValidationnovs_operationCreate Plugin",
    1,
    IsolationModeEnum.Sandbox,
    Description = "Change ownership of operation to business units team and rename ISSO operations to follow this format : Stakeholder | Operation Type | Site ")]

    public class PreValidationovs_operationCreate : IPlugin
    {
        private IOrganizationService service;
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
                    if (target.LogicalName.Equals(ovs_operation.EntityLogicalName))
                    {
                        if (target.Attributes.Contains("ownerid") && target.Attributes["ownerid"] != null &&
                            target.Attributes.Contains("ovs_name") && target.Attributes["ovs_name"] != null &&
                            target.Attributes.Contains("ts_stakeholder") && target.Attributes["ts_stakeholder"] is EntityReference stakeHolderRef &&
                            target.Attributes.Contains("ts_site") && target.Attributes["ts_site"] is EntityReference siteRef &&
                            target.Attributes.Contains("ovs_operationtypeid") && target.Attributes["ovs_operationtypeid"] is EntityReference operationTypeRef)
                        {
                            int userLang = RetrieveUserLang(service, context.UserId);

                            EntityReference owner = (EntityReference)target.Attributes["ownerid"];
                            var userBUId = ((EntityReference)(service.Retrieve("systemuser", context.InitiatingUserId, new ColumnSet("businessunitid"))).Attributes["businessunitid"]).Id;

                            if ((owner.LogicalName == "team" && EnvironmentVariableHelper.IsOwnedByISSO(service, owner, tracingService)) || EnvironmentVariableHelper.IsISSOBU(service, userBUId, tracingService))
                            {
                                string stakeholderAltLangValue = RetrieveAltLangName(service, stakeHolderRef, userLang == 1033 ? "ovs_accountnamefrench" : "ovs_accountnameenglish", "name");
                                string operationTypeAltLangValue = RetrieveAltLangName(service, operationTypeRef, userLang == 1033 ? "ovs_operationtypenamefrench" : "ovs_operationtypenameenglish", "ovs_name");
                                string siteAltLangValue = RetrieveAltLangName(service, siteRef, userLang == 1033 ? "ts_functionallocationnamefrench" : "ts_functionallocationnameenglish", "msdyn_name");

                                string stakeholderName = RetrieveLookupName(service, stakeHolderRef, "name");
                                string operationTypeName = RetrieveLookupName(service, operationTypeRef, "ovs_name");
                                string siteName = RetrieveLookupName(service, siteRef, "msdyn_name");

                                target.Attributes["ovs_name"] = $"{stakeholderName} | {operationTypeName} | {siteName}";
                                target.Attributes[userLang == 1033 ? "ts_operationnamefrench" : "ts_operationnameenglish"] = $"{stakeholderAltLangValue} | {operationTypeAltLangValue} | {siteAltLangValue}";
                                target.Attributes[userLang == 1033 ? "ts_operationnameenglish" : "ts_operationnamefrench"] = $"{stakeholderName} | {operationTypeName} | {siteName}";
                            }
                            changeOwnerToUserBusinessUnit(context, service);
                        }
                    }
                }
                catch (Exception e)
                {
                    throw new InvalidPluginExecutionException(e.Message);
                }
            }
        }

        private void changeOwnerToUserBusinessUnit(IPluginExecutionContext context, IOrganizationService service)
        {
            Guid userId = context.InitiatingUserId;
            Guid userBusinessUnitId = ((EntityReference)(service.Retrieve("systemuser", userId, new ColumnSet("businessunitid"))).Attributes["businessunitid"]).Id;

            if (!EnvironmentVariableHelper.IsTCBU(service, userBusinessUnitId))
            {
                EntityReference teamReference = EnvironmentVariableHelper.RetrieveTeamByBusinessUnitId(service, userBusinessUnitId);
                if (teamReference != null)
                {
                    Entity target = (Entity)context.InputParameters["Target"];
                    target.Attributes["ownerid"] = teamReference;
                }
            }
        }

        private string RetrieveLookupName(IOrganizationService service, EntityReference ownerRef, string attributeName)
        {
            try
            {
                ColumnSet columnSet = new ColumnSet(attributeName);
                Entity ownerEntity = service.Retrieve(ownerRef.LogicalName, ownerRef.Id, columnSet);

                if (ownerEntity.Contains(attributeName))
                {
                    return ownerEntity.GetAttributeValue<string>(attributeName);
                }
                return string.Empty;
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException($"Error retrieving owner name ({ownerRef.LogicalName} - {ownerRef.Id}): {ex.Message}", ex);
            }
        }

        private string RetrieveAltLangName(IOrganizationService service, EntityReference entityRef, string alternateLangLookupName, string LookupName)
        {
            try
            {
                string alternateLangName = RetrieveLookupName(service, entityRef, alternateLangLookupName);

                if (string.IsNullOrEmpty(alternateLangName))
                {
                    return RetrieveLookupName(service, entityRef, LookupName);
                }

                return alternateLangName;
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException($"Error retrieving alternate language value ({entityRef.Name}): {ex.Message}", ex);
            }
        }

        private int RetrieveUserLang(IOrganizationService service, Guid userId)
        {
            ColumnSet columnSet = new ColumnSet("uilanguageid");
            Entity userSettings = service.Retrieve("usersettings", userId, columnSet);

            if (userSettings.Contains("uilanguageid") && userSettings["uilanguageid"] is int languageCode)
            {
                return languageCode;
            }
            return 1033;
        }
    }
}