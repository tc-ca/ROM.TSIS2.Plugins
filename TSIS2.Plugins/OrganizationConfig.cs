using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace TSIS2.Plugins
{
    // Centralizes organization structure IDs (teams/BUs) from environment variables and provides helper checks
    // for ownership and membership used across plugins.
    public static class OrganizationConfig
    {
        public static class TeamEnvVarKeys
        {
            public const string AVIATION_SECURITY = "ts_AviationSecurityTeamGUID";
            public const string AVIATION_SECURITY_DOMESTIC = "ts_AviationSecurityDirectorateDomesticTeamGUID";
            public const string AVIATION_SECURITY_INTERNATIONAL = "ts_AviationSecurityInternationalTeamGUID";
            public const string AVIATION_SECURITY_INTERNATIONAL_DEV = "ts_AviationSecurityInternationalTeamGUID_DEV";
            public const string AVIATION_SECURITY_DOMESTIC_INSPECTORS = "ts_AviationSecurityDomesticInspectorsTeamGUID";
            public const string AVIATION_SECURITY_INTERNATIONAL_INSPECTORS = "ts_AviationSecurityInternationalInspectorsTeamGUID";
            public const string ISSO_TEAM = "ts_IntermodalSurfaceSecurityOversightISSOTeamGUID";
            public const string ISSO_TEAM_DEV = "ts_IntermodalSurfaceSecurityOversightISSOTeamGUID_DEV";
            public const string ISSO_INSPECTORS = "ts_IntermodalSurfaceSecurityOversightInspectorsTeamGUID";
            public const string RAIL_SAFETY = "ts_RailSafetyTeamGUID";
            public const string RAIL_SAFETY_ADMIN = "ts_ROMRailSafetyAdministratorGUID";
        }

        public static class BusinessUnitEnvVarKeys
        {
            public const string AVIATION_SECURITY_DOMESTIC = "ts_AviationSecurityDirectorateDomesticBusinessUnitGUID";
            public const string AVIATION_SECURITY_INTERNATIONAL = "ts_AviationSecurityInternationalBusinessUnitGUID";
            public const string AVIATION_SECURITY_INTERNATIONAL_DEV = "ts_AviationSecurityInternationalBusinessUnitGUID_DEV";
            public const string AVIATION_SECURITY_DIRECTORATE = "ts_AviationSecurityDirectorateBusinessUnitGUID";
            public const string AVIATION_SECURITY_PPP = "ts_AviationSecurityPPPBusinessUnitGUID";
            public const string ISSO = "ts_IntermodalSurfaceSecurityOversightISSOBusinessUnitGUID";
            public const string TRANSPORT_CANADA = "ts_TransportCanadaBusinessUnitGUID";
        }

        private static Dictionary<string, string> _configValues;
        private static readonly object _lock = new object();

        private static readonly string[] AllEnvVarKeys = new[]
        {
            TeamEnvVarKeys.AVIATION_SECURITY,
            TeamEnvVarKeys.AVIATION_SECURITY_DOMESTIC,
            TeamEnvVarKeys.AVIATION_SECURITY_INTERNATIONAL,
            TeamEnvVarKeys.AVIATION_SECURITY_INTERNATIONAL_DEV,
            TeamEnvVarKeys.AVIATION_SECURITY_DOMESTIC_INSPECTORS,
            TeamEnvVarKeys.AVIATION_SECURITY_INTERNATIONAL_INSPECTORS,
            TeamEnvVarKeys.ISSO_TEAM,
            TeamEnvVarKeys.ISSO_TEAM_DEV,
            TeamEnvVarKeys.ISSO_INSPECTORS,
            TeamEnvVarKeys.RAIL_SAFETY,
            TeamEnvVarKeys.RAIL_SAFETY_ADMIN,
            BusinessUnitEnvVarKeys.AVIATION_SECURITY_DOMESTIC,
            BusinessUnitEnvVarKeys.AVIATION_SECURITY_INTERNATIONAL,
            BusinessUnitEnvVarKeys.AVIATION_SECURITY_INTERNATIONAL_DEV,
            BusinessUnitEnvVarKeys.AVIATION_SECURITY_DIRECTORATE,
            BusinessUnitEnvVarKeys.AVIATION_SECURITY_PPP,
            BusinessUnitEnvVarKeys.ISSO,
            BusinessUnitEnvVarKeys.TRANSPORT_CANADA
        };

        private static readonly string[] AvSecBuEnvVarKeys = new[]
        {
            BusinessUnitEnvVarKeys.AVIATION_SECURITY_DOMESTIC,
            BusinessUnitEnvVarKeys.AVIATION_SECURITY_INTERNATIONAL,
            BusinessUnitEnvVarKeys.AVIATION_SECURITY_INTERNATIONAL_DEV,
            BusinessUnitEnvVarKeys.AVIATION_SECURITY_DIRECTORATE
        };

        private static readonly string[] AvSecPPPBUKeys = new[]
        {
            BusinessUnitEnvVarKeys.AVIATION_SECURITY_PPP
        };

        private static readonly string[] ISSOBUKeys = new[]
        {
            BusinessUnitEnvVarKeys.ISSO
        };

        private static readonly string[] TCBUEnvVarKeys = new[]
        {
            BusinessUnitEnvVarKeys.TRANSPORT_CANADA
        };

        private static readonly string[] AvSecTeamEnvVarKeys = new[]
        {
            TeamEnvVarKeys.AVIATION_SECURITY,
            TeamEnvVarKeys.AVIATION_SECURITY_DOMESTIC,
            TeamEnvVarKeys.AVIATION_SECURITY_INTERNATIONAL,
            TeamEnvVarKeys.AVIATION_SECURITY_INTERNATIONAL_DEV
        };

        private static readonly string[] ISSOTeamEnvVarKeys = new[]
        {
            TeamEnvVarKeys.ISSO_TEAM,
            TeamEnvVarKeys.ISSO_TEAM_DEV
        };

        private static readonly string[] IssoInspectorsTeamEnvVarKeys = new[]
        {
            TeamEnvVarKeys.ISSO_INSPECTORS
        };

        private static readonly string[] AvSecInspectorsTeamKeys = new[]
        {
            TeamEnvVarKeys.AVIATION_SECURITY_DOMESTIC_INSPECTORS,
            TeamEnvVarKeys.AVIATION_SECURITY_INTERNATIONAL_INSPECTORS
        };

        private static readonly string[] AvSecOwnerTeamEnvVarKeys = new[]
        {
            TeamEnvVarKeys.AVIATION_SECURITY,
            TeamEnvVarKeys.AVIATION_SECURITY_DOMESTIC,
            TeamEnvVarKeys.AVIATION_SECURITY_INTERNATIONAL,
            TeamEnvVarKeys.AVIATION_SECURITY_INTERNATIONAL_DEV,
            TeamEnvVarKeys.AVIATION_SECURITY_DOMESTIC_INSPECTORS,
            TeamEnvVarKeys.AVIATION_SECURITY_INTERNATIONAL_INSPECTORS
        };

        private static readonly string[] ISSOOwnerTeamEnvVarKeys = new[]
        {
            TeamEnvVarKeys.ISSO_TEAM,
            TeamEnvVarKeys.ISSO_TEAM_DEV,
            TeamEnvVarKeys.ISSO_INSPECTORS
        };

        // Ensures environment variables are loaded (once per AppDomain/sandbox worker life).
        private static void EnsureConfigLoaded(IOrganizationService service)
        {
            if (_configValues != null) return;

            lock (_lock)
            {
                if (_configValues == null)
                {
                    LoadAllEnvironmentVariables(service);
                }
            }
        }

        // Loads environment variables in one query.
        // Uses override value when present; falls back to defaultvalue when not.
        private static void LoadAllEnvironmentVariables(IOrganizationService service)
        {
            var config = new Dictionary<string, string>();
            var query = new QueryExpression("environmentvariabledefinition")
            {
                ColumnSet = new ColumnSet("schemaname", "defaultvalue"),
                NoLock = true
            };

            query.Criteria.AddCondition("schemaname", ConditionOperator.In, AllEnvVarKeys);

            var valueLink = query.AddLink(
                "environmentvariablevalue",
                "environmentvariabledefinitionid",
                "environmentvariabledefinitionid",
                JoinOperator.LeftOuter
            );

            valueLink.EntityAlias = "val";
            valueLink.Columns = new ColumnSet("value");

            var results = service.RetrieveMultiple(query);

            foreach (var entity in results.Entities)
            {
                var schemaName = entity.GetAttributeValue<string>("schemaname");
                if (string.IsNullOrWhiteSpace(schemaName)) continue;

                // Prefer override value if it's non-blank; otherwise fallback to defaultvalue
                var overrideVal = entity.GetAttributeValue<AliasedValue>("val.value")?.Value?.ToString();
                var defaultVal = entity.GetAttributeValue<string>("defaultvalue");

                var chosen = !string.IsNullOrWhiteSpace(overrideVal) ? overrideVal : defaultVal;
                if (string.IsNullOrWhiteSpace(chosen)) continue;

                config[schemaName] = chosen.Trim();
            }

            _configValues = config;
        }


        private static Guid? TryGetEnvVarGuid(string key)
        {
            if (_configValues != null && _configValues.TryGetValue(key, out var value) && Guid.TryParse(value, out var guid))
            {
                return guid;
            }
            return null;
        }

        private static bool MatchesAnyEnvVarGuid(Guid id, string[] keys, ITracingService tracer)
        {
            if (keys == null || keys.Length == 0) return false;

            foreach (var key in keys)
            {
                var guid = TryGetEnvVarGuid(key);
                if (guid.HasValue && guid.Value == id)
                {
                    tracer?.Trace($"SecurityConfig: Match found for {key} with id={id}");
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Resolves businessunitid for any EntityReference that exposes a businessunitid column.
        /// Works for "team" and "systemuser" (and any other entity with businessunitid).
        /// Returns null if ref is null, entity has no BU, or retrieve fails.
        /// </summary>
        public static Guid? TryGetBusinessUnitId(IOrganizationService service, EntityReference entityRef, ITracingService tracer)
        {
            if (entityRef == null) return null;
            // Removed per-instance cache as we are now static and stateless.
            
            try
            {
                var entity = service.Retrieve(entityRef.LogicalName, entityRef.Id, new ColumnSet("businessunitid"));
                var buId = entity.GetAttributeValue<EntityReference>("businessunitid")?.Id;

                return buId;
            }
            catch (Exception ex)
            {
                tracer?.Trace(
                    $"SecurityConfig: TryGetBusinessUnitId failed for {entityRef.LogicalName}:{entityRef.Id}. {ex}"
                );
                return null;
            }
        }

        // Ownership rule used by plugins:
        // - If owner is a team: match by configured team GUIDs
        // - If owner is a user: match by the user's business unit GUIDs
        private static bool IsOwnedByGroup(IOrganizationService service, EntityReference owner, string[] teamKeys, string[] buKeys, ITracingService tracer)
        {
            if (owner == null) return false;
            EnsureConfigLoaded(service);

            if (owner.LogicalName == "team")
                return MatchesAnyEnvVarGuid(owner.Id, teamKeys, tracer);

            if (owner.LogicalName == "systemuser")
            {
                var buId = TryGetBusinessUnitId(service, owner, tracer);
                return buId.HasValue && MatchesAnyEnvVarGuid(buId.Value, buKeys, tracer);
            }

            return false;
        }

        public static bool IsAvSecBU(IOrganizationService service, Guid buId, ITracingService tracer = null)
        {
            EnsureConfigLoaded(service);
            return MatchesAnyEnvVarGuid(buId, AvSecBuEnvVarKeys, tracer);
        }

        public static bool IsAvSecPPPBU(IOrganizationService service, Guid buId, ITracingService tracer = null)
        {
            EnsureConfigLoaded(service);
            return MatchesAnyEnvVarGuid(buId, AvSecPPPBUKeys, tracer);
        }

        public static bool IsISSOBU(IOrganizationService service, Guid buId, ITracingService tracer = null)
        {
            EnsureConfigLoaded(service);
            return MatchesAnyEnvVarGuid(buId, ISSOBUKeys, tracer);
        }

        public static bool IsTCBU(IOrganizationService service, Guid buId, ITracingService tracer = null)
        {
            EnsureConfigLoaded(service);
            return MatchesAnyEnvVarGuid(buId, TCBUEnvVarKeys, tracer);
        }

        public static bool IsAvSecTeam(IOrganizationService service, Guid teamId, ITracingService tracer = null)
        {
            EnsureConfigLoaded(service);
            return MatchesAnyEnvVarGuid(teamId, AvSecTeamEnvVarKeys, tracer);
        }

        public static bool IsISSOTeam(IOrganizationService service, Guid teamId, ITracingService tracer = null)
        {
            EnsureConfigLoaded(service);
            return MatchesAnyEnvVarGuid(teamId, ISSOTeamEnvVarKeys, tracer);
        }

        public static bool IsISSOInspectorsTeam(IOrganizationService service, Guid teamId, ITracingService tracer = null)
        {
            EnsureConfigLoaded(service);
            return MatchesAnyEnvVarGuid(teamId, IssoInspectorsTeamEnvVarKeys, tracer);
        }

        public static bool IsAvSecInspectorsTeam(IOrganizationService service, Guid teamId, ITracingService tracer = null)
        {
            EnsureConfigLoaded(service);
            return MatchesAnyEnvVarGuid(teamId, AvSecInspectorsTeamKeys, tracer);
        }

        public static bool IsOwnedByAvSec(IOrganizationService service, EntityReference owner, ITracingService tracer = null)
        {
            if (owner == null) return false;

            var result = IsOwnedByGroup(service, owner, AvSecOwnerTeamEnvVarKeys, AvSecBuEnvVarKeys, tracer);
            if (tracer != null && result)
            {
                tracer.Trace($"SecurityConfig: IsOwnedByAvSec=true for {owner.LogicalName}:{owner.Id}");
            }
            return result;
        }

        public static bool IsOwnedByISSO(IOrganizationService service, EntityReference owner, ITracingService tracer = null)
        {
            if (owner == null) return false;

            var result = IsOwnedByGroup(service, owner, ISSOOwnerTeamEnvVarKeys, ISSOBUKeys, tracer);
            if (tracer != null && result)
            {
                tracer.Trace($"SecurityConfig: IsOwnedByISSO=true for {owner.LogicalName}:{owner.Id}");
            }
            return result;
        }

        public static bool IsOwnedByRailSafety(IOrganizationService service, EntityReference owner, ITracingService tracer = null)
        {
            if (owner == null || owner.LogicalName != "team") return false;
            EnsureConfigLoaded(service);

            var result = MatchesAnyEnvVarGuid(owner.Id, new[] { TeamEnvVarKeys.RAIL_SAFETY }, tracer);
            if (tracer != null && result)
            {
                tracer.Trace($"SecurityConfig: IsOwnedByRailSafety=true for team:{owner.Id}");
            }
            return result;
        }

        public static bool IsOwnedByRailSafetyAdministrator(IOrganizationService service, EntityReference owner, ITracingService tracer = null)
        {
            if (owner == null || owner.LogicalName != "team") return false;
            EnsureConfigLoaded(service);

            var result = MatchesAnyEnvVarGuid(owner.Id, new[] { TeamEnvVarKeys.RAIL_SAFETY_ADMIN }, tracer);
            if (tracer != null && result)
            {
                tracer.Trace($"SecurityConfig: IsOwnedByRailSafetyAdministrator=true for team:{owner.Id}");
            }
            return result;
        }

        public static List<Guid> GetAvSecBUGuids(IOrganizationService service)
        {
            EnsureConfigLoaded(service);
            var guids = new List<Guid>();
            foreach (var key in AvSecBuEnvVarKeys)
            {
                var guid = TryGetEnvVarGuid(key);
                if (guid.HasValue)
                {
                    guids.Add(guid.Value);
                }
            }
            return guids;
        }

        public static List<Guid> GetISSOBUGuids(IOrganizationService service)
        {
            EnsureConfigLoaded(service);
            var guids = new List<Guid>();
            foreach (var key in ISSOBUKeys)
            {
                var guid = TryGetEnvVarGuid(key);
                if (guid.HasValue)
                {
                    guids.Add(guid.Value);
                }
            }
            return guids;
        }

        // Returns the Business Unit's *default* team (isdefault = true). Ownership logic assumes default team.
        public static EntityReference GetTeamByBusinessUnitId(IOrganizationService service, Guid businessUnitId)
        {
            var query = new QueryExpression("team")
            {
                ColumnSet = new ColumnSet("teamid"),
                Criteria = new FilterExpression
                {
                    Conditions =
                    {
                        new ConditionExpression("businessunitid", ConditionOperator.Equal, businessUnitId),
                        new ConditionExpression("isdefault", ConditionOperator.Equal, true)
                    }
                },
                TopCount = 1
            };

            var result = service.RetrieveMultiple(query);
            return result.Entities.Count > 0 ? result.Entities[0].ToEntityReference() : null;
        }

        public static bool IsUserMemberOfTeam(IOrganizationService service, Guid userId, Guid teamId)
        {
            var query = new QueryExpression("team")
            {
                ColumnSet = new ColumnSet("teamid"),
                Criteria = new FilterExpression
                {
                    Conditions =
                    {
                        // Exclude access teams (teamtype = 1); we only care about owner/standard teams.
                        new ConditionExpression("teamtype", ConditionOperator.NotEqual, 1),
                        new ConditionExpression("teamid", ConditionOperator.Equal, teamId)
                    }
                }
            };

            var membershipLink = query.AddLink("teammembership", "teamid", "teamid");
            var userLink = membershipLink.AddLink("systemuser", "systemuserid", "systemuserid");
            userLink.LinkCriteria.AddCondition("systemuserid", ConditionOperator.Equal, userId);

            var result = service.RetrieveMultiple(query);
            return result.Entities.Count > 0;
        }

    }
}
