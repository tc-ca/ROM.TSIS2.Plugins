using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

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

    public class PreValidationovs_operationCreate : PluginBase
    {

        public PreValidationovs_operationCreate(string unsecure, string secure)
            : base(typeof(PreValidationovs_operationCreate))
        {
        }

        protected override void ExecuteCrmPlugin(LocalPluginContext localContext)
        {
            if (localContext == null)
            {
                throw new InvalidPluginExecutionException("localContext");
            }

            IPluginExecutionContext context = localContext.PluginExecutionContext;
            ITracingService tracingService = localContext.TracingService;
            IOrganizationService service = localContext.OrganizationService;

            if (context.Depth > 1)
            {
                tracingService.Trace("[OperationCreate] Exiting due to Depth > 1");
                return;
            }

            if (context.InputParameters.Contains("Target") &&
                context.InputParameters["Target"] is Entity)
            {
                Entity target = (Entity)context.InputParameters["Target"];

                try
                {
                    if (target.LogicalName.Equals(ovs_operation.EntityLogicalName))
                    {
                        var owner = target.GetAttributeValue<EntityReference>("ownerid");
                        var stakeholderRef = target.GetAttributeValue<EntityReference>("ts_stakeholder");
                        var siteRef = target.GetAttributeValue<EntityReference>("ts_site");
                        var operationTypeRef = target.GetAttributeValue<EntityReference>("ovs_operationtypeid");
                        var name = target.GetAttributeValue<string>("ovs_name");

                        if (owner == null || stakeholderRef == null || siteRef == null || operationTypeRef == null || string.IsNullOrEmpty(name))
                        {
                            tracingService.Trace("Missing required attributes; skipping processing");
                            return;
                        }

                        int userLang = RetrieveUserLang(service, context.UserId);
                        var userBUId = context.BusinessUnitId;

                        bool isISSOOwner = owner.LogicalName == "team" && OrganizationConfig.IsOwnedByISSO(service, owner, tracingService);
                        bool isUserInISSOBU = OrganizationConfig.IsISSOBU(service, userBUId, tracingService);
                        bool shouldRenameForISSO = isISSOOwner || isUserInISSOBU;

                        string ownerDisplay = $"{owner.LogicalName}:{owner.Id}" + (!string.IsNullOrEmpty(owner.Name) ? $" ; {owner.Name}" : "");

                        if (shouldRenameForISSO)
                        {
                            string stakeholderAltLangValue = RetrieveAltLangName(service, stakeholderRef, userLang == 1033 ? "ovs_accountnamefrench" : "ovs_accountnameenglish", "name", tracingService);
                            string operationTypeAltLangValue = RetrieveAltLangName(service, operationTypeRef, userLang == 1033 ? "ovs_operationtypenamefrench" : "ovs_operationtypenameenglish", "ovs_name", tracingService);
                            string siteAltLangValue = RetrieveAltLangName(service, siteRef, userLang == 1033 ? "ts_functionallocationnamefrench" : "ts_functionallocationnameenglish", "msdyn_name", tracingService);

                            string stakeholderName = RetrieveLookupName(service, stakeholderRef, "name", tracingService);
                            string operationTypeName = RetrieveLookupName(service, operationTypeRef, "ovs_name", tracingService);
                            string siteName = RetrieveLookupName(service, siteRef, "msdyn_name", tracingService);

                            string newName = $"{stakeholderName} | {operationTypeName} | {siteName}";
                            target.Attributes["ovs_name"] = newName;
                            target.Attributes[userLang == 1033 ? "ts_operationnamefrench" : "ts_operationnameenglish"] = $"{stakeholderAltLangValue} | {operationTypeAltLangValue} | {siteAltLangValue}";
                            target.Attributes[userLang == 1033 ? "ts_operationnameenglish" : "ts_operationnamefrench"] = newName;
                        }

                        ChangeOwnerToOperationTypeBusinessUnit(localContext, operationTypeRef);
                    }
                }
                catch (Exception e)
                {
                    localContext.Trace($"Exception: {e.ToString()}");
                    throw new InvalidPluginExecutionException("PreValidationovs_operationCreate failed", e);
                }
            }
        }

        private void ChangeOwnerToOperationTypeBusinessUnit(LocalPluginContext localContext, EntityReference operationTypeRef)
        {
            if (localContext == null) throw new ArgumentNullException(nameof(localContext));

            var service = localContext.OrganizationService;
            var tracingService = localContext.TracingService;
            var context = localContext.PluginExecutionContext;

            const string PREFIX = "OWNER (OP TYPE BU)";

            if (operationTypeRef == null)
            {
                tracingService?.Trace($"{PREFIX}: ✖ missingOperationTypeRef");
                return;
            }

            // 1) Retrieve operation type owner (team/user)
            Entity opType;
            try
            {
                opType = service.Retrieve(operationTypeRef.LogicalName, operationTypeRef.Id, new ColumnSet("ownerid"));
            }
            catch (Exception ex)
            {
                tracingService?.Trace($"{PREFIX}: ✖ retrieveOpTypeFailed opTypeId={operationTypeRef.Id} ex={ex}");
                return;
            }

            var opTypeOwner = opType.GetAttributeValue<EntityReference>("ownerid");
            if (opTypeOwner == null)
            {
                tracingService?.Trace($"{PREFIX}: ✖ opTypeHasNoOwner opTypeId={operationTypeRef.Id}");
                return;
            }

            // 2) Resolve BU from op type owner
            var buId = OrganizationConfig.TryGetBusinessUnitId(service, opTypeOwner, tracingService);
            if (!buId.HasValue)
            {
                tracingService?.Trace($"{PREFIX}: ✖ couldNotResolveBU opTypeOwner={opTypeOwner.LogicalName}:{opTypeOwner.Id}");
                return;
            }

            // 3) Assign operation owner to BU default team
            var teamRef = OrganizationConfig.GetTeamByBusinessUnitId(service, buId.Value);
            if (teamRef == null)
            {
                tracingService?.Trace($"{PREFIX}: ✖ noTeamForBU buId={buId.Value} opTypeId={operationTypeRef.Id}");
                return;
            }

            var target = (Entity)context.InputParameters["Target"];
            target["ownerid"] = teamRef;

            var opTypeOwnerDisplay =
                $"{opTypeOwner.LogicalName}:{opTypeOwner.Id}" +
                (!string.IsNullOrEmpty(opTypeOwner.Name) ? $" ; {opTypeOwner.Name}" : "");

            var teamDisplay =
                $"{teamRef.Id}" +
                (!string.IsNullOrEmpty(teamRef.Name) ? $" ; {teamRef.Name}" : "");
        }


        private string RetrieveLookupName(IOrganizationService service, EntityReference lookupRef, string attributeName, ITracingService tracingService)
        {
            if (lookupRef == null) return "Unknown";


            try
            {
                Entity entity = service.Retrieve(lookupRef.LogicalName, lookupRef.Id, new ColumnSet(attributeName));
                return entity.GetAttributeValue<string>(attributeName) ?? "N/A";

            }
            catch (Exception ex)
            {
                // Log to tracer and return a fallback instead of crashing the whole save
                tracingService?.Trace($"Error retrieving {attributeName}: {ex}");
                return "Error";
            }
        }

        private string RetrieveAltLangName(IOrganizationService service, EntityReference entityRef, string alternateLangLookupName, string LookupName, ITracingService tracingService)
        {
            if (entityRef == null)
            {
                tracingService?.Trace($"RetrieveAltLangName: null entityRef, returning fallback");
                return "N/A";
            }

            try
            {
                string alternateLangName = RetrieveLookupName(service, entityRef, alternateLangLookupName, tracingService);

                if (string.IsNullOrEmpty(alternateLangName) || alternateLangName == "Error" || alternateLangName == "Unknown")
                {
                    tracingService?.Trace($"RetrieveAltLangName: alternateLang empty/error, falling back to primary language");
                    return RetrieveLookupName(service, entityRef, LookupName, tracingService);
                }

                return alternateLangName;
            }
            catch (Exception ex)
            {
                tracingService?.Trace($"RetrieveAltLangName: exception for {entityRef?.LogicalName}:{entityRef?.Id}, falling back to primary: {ex}");
                try
                {
                    return RetrieveLookupName(service, entityRef, LookupName, tracingService);
                }
                catch
                {
                    return "N/A";
                }
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