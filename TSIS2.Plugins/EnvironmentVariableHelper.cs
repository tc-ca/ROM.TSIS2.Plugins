using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace TSIS2.Plugins
{
    public class EnvironmentVariableHelper
    {
        // Environment Variable Schema Name Constants
        public static class TeamSchemaNames
        {
            public const string AVIATION_SECURITY_DOMESTIC = "ts_AviationSecurityDirectorateDomesticTeamGUID";
            public const string AVIATION_SECURITY_INTERNATIONAL = "ts_AviationSecurityInternationalTeamGUID";
            public const string AVIATION_SECURITY_DOMESTIC_INSPECTORS = "ts_AviationSecurityDomesticInspectorsTeamGUID";
            public const string AVIATION_SECURITY_INTERNATIONAL_INSPECTORS = "ts_AviationSecurityInternationalInspectorsTeamGUID";
            public const string ISSO_TEAM = "ts_IntermodalSurfaceSecurityOversightISSOTeamGUID";
            public const string ISSO_INSPECTORS = "ts_IntermodalSurfaceSecurityOversightInspectorsTeamGUID";
        }

        public static class BUSchemaNames
        {
            public const string AVIATION_SECURITY_DOMESTIC = "ts_AviationSecurityDirectorateDomesticBusinessUnitGUID";
            public const string AVIATION_SECURITY_INTERNATIONAL = "ts_AviationSecurityInternationalBusinessUnitGUID";
            public const string AVIATION_SECURITY_DIRECTORATE = "ts_AviationSecurityDirectorateBusinessUnitGUID";
            public const string AVIATION_SECURITY_PPP = "ts_AviationSecurityPPPBusinessUnitGUID";
            public const string ISSO = "ts_IntermodalSurfaceSecurityOversightISSOBusinessUnitGUID";
            public const string TRANSPORT_CANADA = "ts_TransportCanadaBusinessUnitGUID";
        }

        private static void Trace(ITracingService tracer, string message)
        {
            if (tracer != null && !string.IsNullOrEmpty(message))
            {
                tracer.Trace(message);
            }
        }

        public static string GetEnvironmentVariableValue(IOrganizationService service, string schemaName, ITracingService tracer = null)
        {
            Trace(tracer, "GetEnvironmentVariableValue schemaName=" + schemaName);

            // Create a query to retrieve the environment variable definition
            QueryExpression query = new QueryExpression("environmentvariabledefinition")
            {
                //The value we want to receive
                ColumnSet = new ColumnSet("defaultvalue")
            };
            //filters the result by the schema name
            query.Criteria.AddCondition("schemaname", ConditionOperator.Equal, schemaName);

            // get the environment variable definition
            Entity definition = service.RetrieveMultiple(query).Entities.FirstOrDefault();

            if (definition != null)
            {
                QueryExpression query2 = new QueryExpression("environmentvariablevalue")
                {
                    ColumnSet = new ColumnSet("value")
                };
                //filter by environment variable id
                query2.Criteria.AddCondition("environmentvariabledefinitionid", ConditionOperator.Equal, definition.Id);

                //get the environment variable value
                Entity value = service.RetrieveMultiple(query2).Entities.FirstOrDefault();

                if (value != null)
                {
                    string result = value.GetAttributeValue<string>("value");
                    Trace(tracer, "GetEnvironmentVariableValue result=" + result);
                    return result;
                }

                string defaultValue = definition.GetAttributeValue<string>("defaultvalue");
                Trace(tracer, "GetEnvironmentVariableValue result=defaultValue=" + defaultValue);
                return defaultValue;
            }
            Trace(tracer, "GetEnvironmentVariableValue result=null");
            return null;
        }

        private static string NormalizeGuid(string guid)
        {
            if (string.IsNullOrEmpty(guid)) return null;
            return guid.Replace("{", "").Replace("}", "").ToLower();
        }

        private static string NormalizeGuid(Guid guid)
        {
            return guid.ToString("D").ToLower();
        }

        // Generic method to check if a GUID matches any of the provided environment variable schema names
        public static bool IsInList(IOrganizationService service, Guid id, string[] schemaNames, ITracingService tracer = null)
        {
            Trace(tracer, "IsInList id=" + id);
            if (schemaNames == null || schemaNames.Length == 0)
            {
                Trace(tracer, "IsInList result=false (schemaNames null or empty)");
                return false;
            }
            string normalizedId = NormalizeGuid(id);

            foreach (string schemaName in schemaNames)
            {
                string envValue = GetEnvironmentVariableValue(service, schemaName, tracer);
                if (!string.IsNullOrEmpty(envValue))
                {
                    string normalizedEnvValue = NormalizeGuid(envValue);
                    if (normalizedId == normalizedEnvValue)
                    {
                        Trace(tracer, "IsInList result=true (match found for " + schemaName + ")");
                        return true;
                    }
                }
            }
            Trace(tracer, "IsInList result=false");
            return false;
        }

        // Check if a Business Unit ID is AvSec
        public static bool IsAvSecBU(IOrganizationService service, Guid buId, ITracingService tracer = null)
        {
            Trace(tracer, "IsAvSecBU buId=" + buId);
            bool result = IsInList(service, buId, new[]
            {
                BUSchemaNames.AVIATION_SECURITY_DOMESTIC,
                BUSchemaNames.AVIATION_SECURITY_INTERNATIONAL,
                BUSchemaNames.AVIATION_SECURITY_DIRECTORATE
            }, tracer);
            Trace(tracer, "IsAvSecBU result=" + result);
            return result;
        }

        // Check if a Business Unit ID is AvSec PPP
        public static bool IsAvSecPPPBU(IOrganizationService service, Guid buId)
        {
            return IsInList(service, buId, new[] { BUSchemaNames.AVIATION_SECURITY_PPP });
        }

        // Check if a Business Unit ID is ISSO
        public static bool IsISSOBU(IOrganizationService service, Guid buId, ITracingService tracer = null)
        {
            Trace(tracer, "IsISSOBU buId=" + buId);
            bool result = IsInList(service, buId, new[] { BUSchemaNames.ISSO }, tracer);
            Trace(tracer, "IsISSOBU result=" + result);
            return result;
        }

        // Check if a Business Unit ID is Transport Canada
        public static bool IsTCBU(IOrganizationService service, Guid buId)
        {
            return IsInList(service, buId, new[] { BUSchemaNames.TRANSPORT_CANADA });
        }

        // Check if a Team ID is AvSec Team
        public static bool IsAvSecTeam(IOrganizationService service, Guid teamId)
        {
            return IsInList(service, teamId, new[]
            {
                TeamSchemaNames.AVIATION_SECURITY_DOMESTIC,
                TeamSchemaNames.AVIATION_SECURITY_INTERNATIONAL
            });
        }

        // Check if a Team ID is ISSO Team
        public static bool IsISSOTeam(IOrganizationService service, Guid teamId)
        {
            return IsInList(service, teamId, new[] { TeamSchemaNames.ISSO_TEAM });
        }

        // Check if a Team ID is ISSO Inspectors
        public static bool IsISSOInspectorsTeam(IOrganizationService service, Guid teamId)
        {
            return IsInList(service, teamId, new[] { TeamSchemaNames.ISSO_INSPECTORS });
        }

        // Check if a Team ID is AvSec Inspectors (Domestic or International)
        public static bool IsAvSecInspectorsTeam(IOrganizationService service, Guid teamId)
        {
            return IsInList(service, teamId, new[]
            {
                TeamSchemaNames.AVIATION_SECURITY_DOMESTIC_INSPECTORS,
                TeamSchemaNames.AVIATION_SECURITY_INTERNATIONAL_INSPECTORS
            });
        }

        // Check if owner (EntityReference) is AvSec team
        public static bool IsOwnedByAvSec(IOrganizationService service, EntityReference owner, ITracingService tracer = null)
        {
            if (owner == null)
            {
                Trace(tracer, "IsOwnedByAvSec result=false (owner is null)");
                return false;
            }

            Trace(tracer, "IsOwnedByAvSec owner=" + owner.LogicalName + " id=" + owner.Id);

            if (owner.LogicalName == "team")
            {
                bool result = IsAvSecTeam(service, owner.Id) || IsAvSecInspectorsTeam(service, owner.Id);
                Trace(tracer, "IsOwnedByAvSec result=" + result + " (team check)");
                return result;
            }

            if (owner.LogicalName == "systemuser")
            {
                Entity user = service.Retrieve("systemuser", owner.Id, new ColumnSet("businessunitid"));
                EntityReference buRef = user.GetAttributeValue<EntityReference>("businessunitid");
                if (buRef == null)
                {
                    Trace(tracer, "IsOwnedByAvSec result=false (user has no BU)");
                    return false;
                }
                bool result = IsAvSecBU(service, buRef.Id, tracer);
                Trace(tracer, "IsOwnedByAvSec result=" + result + " (systemuser BU check)");
                return result;
            }

            Trace(tracer, "IsOwnedByAvSec result=false (unknown owner type)");
            return false;
        }

        // Check if owner (EntityReference) is ISSO team
        public static bool IsOwnedByISSO(IOrganizationService service, EntityReference owner, ITracingService tracer = null)
        {
            if (owner == null)
            {
                Trace(tracer, "IsOwnedByISSO result=false (owner is null)");
                return false;
            }

            Trace(tracer, "IsOwnedByISSO owner=" + owner.LogicalName + " id=" + owner.Id);

            if (owner.LogicalName == "team")
            {
                bool result = IsISSOTeam(service, owner.Id) || IsISSOInspectorsTeam(service, owner.Id);
                Trace(tracer, "IsOwnedByISSO result=" + result + " (team check)");
                return result;
            }

            if (owner.LogicalName == "systemuser")
            {
                Entity user = service.Retrieve("systemuser", owner.Id, new ColumnSet("businessunitid"));
                EntityReference buRef = user.GetAttributeValue<EntityReference>("businessunitid");
                if (buRef == null)
                {
                    Trace(tracer, "IsOwnedByISSO result=false (user has no BU)");
                    return false;
                }
                bool result = IsISSOBU(service, buRef.Id, tracer);
                Trace(tracer, "IsOwnedByISSO result=" + result + " (systemuser BU check)");
                return result;
            }

            Trace(tracer, "IsOwnedByISSO result=false (unknown owner type)");
            return false;
        }

        // Generic function to check if owner matches any of the provided environment variable schema names
        public static bool IsOwnedBy(IOrganizationService service, EntityReference owner, string[] schemaNames)
        {
            if (owner == null || schemaNames == null || schemaNames.Length == 0) return false;
            if (owner.LogicalName != "team") return false;
            return IsInList(service, owner.Id, schemaNames);
        }

        // Generic function to check if a BU ID (string) matches any of the provided environment variable schema names
        public static bool IsBusinessUnit(IOrganizationService service, string buId, string[] schemaNames)
        {
            if (string.IsNullOrEmpty(buId) || schemaNames == null || schemaNames.Length == 0) return false;
            var normalizedId = NormalizeGuid(buId);
            if (normalizedId == null) return false;

            foreach (var schemaName in schemaNames)
            {
                var envValue = GetEnvironmentVariableValue(service, schemaName);
                if (!string.IsNullOrEmpty(envValue))
                {
                    var normalizedEnvValue = NormalizeGuid(envValue);
                    if (normalizedId == normalizedEnvValue) return true;
                }
            }
            return false;
        }

        // Returns an array of AvSec Business Unit GUIDs (normalized)
        public static List<string> GetAvSecBUGUIDs(IOrganizationService service)
        {
            var guids = new List<string>();
            var guid1 = GetEnvironmentVariableValue(service, BUSchemaNames.AVIATION_SECURITY_DOMESTIC);
            if (!string.IsNullOrEmpty(guid1)) guids.Add(NormalizeGuid(guid1));
            var guid2 = GetEnvironmentVariableValue(service, BUSchemaNames.AVIATION_SECURITY_INTERNATIONAL);
            if (!string.IsNullOrEmpty(guid2)) guids.Add(NormalizeGuid(guid2));
            var guid3 = GetEnvironmentVariableValue(service, BUSchemaNames.AVIATION_SECURITY_DIRECTORATE);
            if (!string.IsNullOrEmpty(guid3)) guids.Add(NormalizeGuid(guid3));
            return guids;
        }

        // Returns an array of ISSO Business Unit GUIDs (normalized)
        public static List<string> GetISSOBUGUIDs(IOrganizationService service)
        {
            var guids = new List<string>();
            var guid1 = GetEnvironmentVariableValue(service, BUSchemaNames.ISSO);
            if (!string.IsNullOrEmpty(guid1)) guids.Add(NormalizeGuid(guid1));
            return guids;
        }

        // Helper wrappers for checking if a user's BU matches specific BUs
        public static bool IsUserInAvSecBU(IOrganizationService service, Guid userBuId)
        {
            return IsAvSecBU(service, userBuId);
        }

        public static bool IsUserInISSOBU(IOrganizationService service, Guid userBuId)
        {
            return IsISSOBU(service, userBuId);
        }

        public static bool IsUserInTCBU(IOrganizationService service, Guid userBuId)
        {
            return IsTCBU(service, userBuId);
        }

        // Retrieve actual AvSec BU name from the database for display purposes
        public static string GetAvSecBUName(IOrganizationService service)
        {
            var guid = GetEnvironmentVariableValue(service, BUSchemaNames.AVIATION_SECURITY_DOMESTIC);
            if (string.IsNullOrEmpty(guid))
            {
                guid = GetEnvironmentVariableValue(service, BUSchemaNames.AVIATION_SECURITY_INTERNATIONAL);
            }
            if (string.IsNullOrEmpty(guid))
            {
                guid = GetEnvironmentVariableValue(service, BUSchemaNames.AVIATION_SECURITY_DIRECTORATE);
            }
            if (!string.IsNullOrEmpty(guid))
            {
                try
                {
                    var normalizedGuid = NormalizeGuid(guid);
                    if (Guid.TryParse(normalizedGuid, out var buGuid))
                    {
                        var bu = service.Retrieve("businessunit", buGuid, new ColumnSet("name"));
                        return bu.GetAttributeValue<string>("name") ?? "Aviation Security";
                    }
                }
                catch
                {
                    return "Aviation Security";
                }
            }
            return "Aviation Security";
        }

        // Retrieve actual ISSO BU name from the database for display purposes
        public static string GetISSOBUName(IOrganizationService service)
        {
            var guid = GetEnvironmentVariableValue(service, BUSchemaNames.ISSO);
            if (!string.IsNullOrEmpty(guid))
            {
                try
                {
                    var normalizedGuid = NormalizeGuid(guid);
                    if (Guid.TryParse(normalizedGuid, out var buGuid))
                    {
                        var bu = service.Retrieve("businessunit", buGuid, new ColumnSet("name"));
                        return bu.GetAttributeValue<string>("name") ?? "Intermodal Surface Security Oversight";
                    }
                }
                catch
                {
                    return "Intermodal Surface Security Oversight";
                }
            }
            return "Intermodal Surface Security Oversight";
        }

        // Retrieve team by BU ID (looks for a team with matching BU)
        public static EntityReference RetrieveTeamByBusinessUnitId(IOrganizationService service, Guid businessUnitId)
        {
            var query = new QueryExpression("team")
            {
                ColumnSet = new ColumnSet("teamid"),
                Criteria = new FilterExpression
                {
                    Conditions =
                    {
                        new ConditionExpression("businessunitid", ConditionOperator.Equal, businessUnitId)
                    }
                },
                TopCount = 1
            };

            var result = service.RetrieveMultiple(query);
            return result.Entities.Count > 0 ? result.Entities[0].ToEntityReference() : null;
        }

        // Check if a user is in a specific team
        public static bool IsUserInTeam(IOrganizationService service, Guid userId, Guid teamId)
        {
            var fetchXml = $@"
                <fetch distinct='false' mapping='logical'>
                  <entity name='team'>
                    <attribute name='teamid' />
                    <filter type='and'>
                      <condition attribute='teamtype' operator='ne' value='1' />
                      <condition attribute='teamid' operator='eq' value='{teamId}' />
                    </filter>
                    <link-entity name='teammembership' intersect='true' visible='false' to='teamid' from='teamid'>
                      <link-entity name='systemuser' from='systemuserid' to='systemuserid' alias='bb'>
                        <filter type='and'>
                          <condition attribute='systemuserid' operator='eq' value='{userId}' />
                        </filter>
                      </link-entity>
                    </link-entity>
                  </entity>
                </fetch>";

            var result = service.RetrieveMultiple(new FetchExpression(fetchXml));
            return result.Entities.Count > 0;
        }
    }
}
