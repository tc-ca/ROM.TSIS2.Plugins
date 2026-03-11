using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json.Linq;
using TSIS2.Plugins.QuestionnaireProcessor;

namespace TSIS2.QuestionnaireProcessorConsole
{
    /// <summary>
    /// One-off utility to backfill ts_workorder and ts_exemptions on existing ts_questionresponse records.
    /// </summary>
    public class QuestionResponseBackfiller
    {
        private readonly IOrganizationService _service;
        private readonly ILoggingService _logger;
        private readonly QuestionnaireRepository _repository;

        public QuestionResponseBackfiller(IOrganizationService service, ILoggingService logger)
        {
            _service = service ?? throw new ArgumentNullException(nameof(service));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _repository = new QuestionnaireRepository(service, logger);
        }

        public BackfillResult BackfillWorkOrderReference(bool simulationMode)
        {
            var result = new BackfillResult();

            _logger.Info("Starting backfill of ts_workorder on ts_questionresponse records...");
            _logger.Info($"Simulation mode: {(simulationMode ? "ON (no updates will be committed)" : "OFF")}");

            // First, count how many records match the criteria
            var countQuery = new QueryExpression("ts_questionresponse")
            {
                ColumnSet = new ColumnSet(false), // No columns needed for count
                Criteria = new FilterExpression
                {
                    Conditions =
                    {
                        new ConditionExpression("ts_msdyn_workorderservicetask", ConditionOperator.NotNull),
                        new ConditionExpression("ts_workorder", ConditionOperator.Null)
                    }
                },
                PageInfo = new PagingInfo
                {
                    PageNumber = 1,
                    Count = 5000,
                    PagingCookie = null
                }
            };

            int totalCount = 0;
            while (true)
            {
                var countCollection = _service.RetrieveMultiple(countQuery);
                totalCount += countCollection.Entities.Count;
                if (!countCollection.MoreRecords)
                    break;
                countQuery.PageInfo.PageNumber++;
                countQuery.PageInfo.PagingCookie = countCollection.PagingCookie;
            }

            _logger.Info($"Found {totalCount} ts_questionresponse record(s) that need ts_workorder backfilled.");
            if (totalCount == 0)
            {
                _logger.Info("No records to process. Exiting.");
                return result;
            }

            var query = new QueryExpression("ts_questionresponse")
            {
                ColumnSet = new ColumnSet("ts_questionresponseid", "ts_msdyn_workorderservicetask", "ts_workorder", "ts_name"),
                Criteria = new FilterExpression
                {
                    Conditions =
                    {
                        new ConditionExpression("ts_msdyn_workorderservicetask", ConditionOperator.NotNull),
                        new ConditionExpression("ts_workorder", ConditionOperator.Null)
                    }
                },
                PageInfo = new PagingInfo
                {
                    PageNumber = 1,
                    Count = 5000,
                    PagingCookie = null
                }
            };

            var wostCache = new Dictionary<Guid, EntityReference>();

            try
            {
                while (true)
                {
                    var collection = _service.RetrieveMultiple(query);
                    result.TotalScanned += collection.Entities.Count;

                    if (collection.Entities.Count == 0)
                    {
                        break;
                    }

                    _logger.Info($"Scanning page {query.PageInfo.PageNumber}: {collection.Entities.Count} ts_questionresponse records.");

                    foreach (var qr in collection.Entities)
                    {
                        var qrId = qr.Id;
                        var name = qr.GetAttributeValue<string>("ts_name");
                        var wostRef = qr.GetAttributeValue<EntityReference>("ts_msdyn_workorderservicetask");

                        if (wostRef == null)
                        {
                            _logger.Warning($"QuestionResponse {qrId} ({name}) has no ts_msdyn_workorderservicetask, skipping.");
                            continue;
                        }

                        if (!wostCache.TryGetValue(wostRef.Id, out var workOrderRef))
                        {
                            var wost = _repository.GetWorkOrderServiceTask(wostRef.Id);
                            workOrderRef = wost.GetAttributeValue<EntityReference>("msdyn_workorder");
                            wostCache[wostRef.Id] = workOrderRef;
                        }

                        if (workOrderRef == null)
                        {
                            _logger.Warning($"WOST {wostRef.Id} referenced by QuestionResponse {qrId} ({name}) has no msdyn_workorder. Nothing to backfill.");
                            result.SkippedNoWorkOrder++;
                            continue;
                        }

                        _logger.Debug($"Backfilling ts_workorder on ts_questionresponse {qrId} ({name}) with Work Order {workOrderRef.Id}.");

                        if (!simulationMode)
                        {
                            var update = new Entity("ts_questionresponse", qrId)
                            {
                                ["ts_workorder"] = workOrderRef
                            };

                            _service.Update(update);
                        }

                        result.Updated++;
                    }

                    if (!collection.MoreRecords)
                    {
                        break;
                    }

                    query.PageInfo.PageNumber++;
                    query.PageInfo.PagingCookie = collection.PagingCookie;
                }

                _logger.Info($"Backfill completed. Total scanned: {result.TotalScanned}, updated: {result.Updated}, skipped (no work order): {result.SkippedNoWorkOrder}.");
                return result;
            }
            catch (Exception ex)
            {
                _logger.Error($"Error during backfill of ts_workorder: {ex}");
                throw;
            }
        }

        /// <summary>
        /// Backfills ts_exemptions from WOST ovs_questionnaireresponse. Key in JSON is exactly "questionName-Exemptions"
        /// (e.g. "SATR 3 (1) (a)-radiogroup-Exemptions"); the question name comes from the survey definition and is stored in ts_questionname.
        /// Only processes ts_questionresponse for WOSTs whose questionnaire response actually contains at least one non-empty "-Exemptions" key.
        /// Skips Update when the value to write is [] and the record already has null/empty/[] to avoid unnecessary writes.
        /// </summary>
        public ExemptionBackfillResult BackfillExemptions(bool simulationMode)
        {
            var result = new ExemptionBackfillResult();
            _logger.Info("Starting backfill of ts_exemptions on ts_questionresponse records...");
            _logger.Info($"Simulation mode: {(simulationMode ? "ON (no updates will be committed)" : "OFF")}");

            const int pageSize = 5000;
            const int inClauseMax = 1000;

            // Phase 1: Find WOSTs that have any exemption data in ovs_questionnaireresponse (so we only scan their question responses).
            _logger.Info("Finding WOSTs that have exemption data in questionnaire response...");
            var wostResponseCache = new Dictionary<Guid, JObject>();
            var wostIdsWithExemptions = new List<Guid>();
            var wostQuery = new QueryExpression("msdyn_workorderservicetask")
            {
                ColumnSet = new ColumnSet("msdyn_workorderservicetaskid", "ovs_questionnaireresponse"),
                Criteria = new FilterExpression { Conditions = { new ConditionExpression("ovs_questionnaireresponse", ConditionOperator.NotNull), new ConditionExpression("ovs_questionnaireresponse", ConditionOperator.Like, "%-Exemptions%") } },
                PageInfo = new PagingInfo { PageNumber = 1, Count = pageSize, PagingCookie = null }
            };
            int wostPages = 0;
            while (true)
            {
                var wostCollection = _service.RetrieveMultiple(wostQuery);
                if (wostCollection.Entities.Count == 0) break;
                wostPages++;
                _logger.Debug($"Checking WOST page {wostPages}: {wostCollection.Entities.Count} records.");
                foreach (var wost in wostCollection.Entities)
                {
                    var responseJson = wost.GetAttributeValue<string>("ovs_questionnaireresponse");
                    if (string.IsNullOrWhiteSpace(responseJson)) continue;
                    JObject responseObj;
                    try { responseObj = JObject.Parse(responseJson); }
                    catch { continue; }
                    bool hasExemptions = false;
                    foreach (var prop in responseObj.Properties())
                    {
                        if (prop.Name != null && prop.Name.EndsWith("-Exemptions", StringComparison.Ordinal) && prop.Value is JArray arr && arr.Count > 0)
                        {
                            hasExemptions = true;
                            result.KeysWithExemptionsInJson++;
                            _logger.Verbose($"WOST {wost.Id} JSON has key '{prop.Name}' with {arr.Count} exemption(s).");
                            break;
                        }
                    }
                    if (hasExemptions)
                    {
                        wostIdsWithExemptions.Add(wost.Id);
                        wostResponseCache[wost.Id] = responseObj;
                    }
                }
                if (!wostCollection.MoreRecords) break;
                wostQuery.PageInfo.PageNumber++;
                wostQuery.PageInfo.PagingCookie = wostCollection.PagingCookie;
            }
            _logger.Info($"Found {wostIdsWithExemptions.Count} WOST(s) with exemption data in questionnaire response.");
            if (wostIdsWithExemptions.Count == 0)
            {
                _logger.Info("No WOSTs have exemption data. Nothing to backfill.");
                return result;
            }

            // Phase 2: Count ts_questionresponse records for those WOSTs only.
            _logger.Info("Counting ts_questionresponse records for those WOSTs...");
            int totalToScan = 0;
            for (int i = 0; i < wostIdsWithExemptions.Count; i += inClauseMax)
            {
                var chunk = wostIdsWithExemptions.Skip(i).Take(inClauseMax).ToList();
                var countQuery = new QueryExpression("ts_questionresponse")
                {
                    ColumnSet = new ColumnSet(false),
                    Criteria = new FilterExpression { Conditions = { new ConditionExpression("ts_msdyn_workorderservicetask", ConditionOperator.In, chunk) } },
                    PageInfo = new PagingInfo { PageNumber = 1, Count = pageSize, PagingCookie = null }
                };
                while (true)
                {
                    var countPage = _service.RetrieveMultiple(countQuery);
                    totalToScan += countPage.Entities.Count;
                    if (!countPage.MoreRecords) break;
                    countQuery.PageInfo.PageNumber++;
                    countQuery.PageInfo.PagingCookie = countPage.PagingCookie;
                }
            }
            _logger.Info($"Found {totalToScan} ts_questionresponse records to scan (in pages of {pageSize}).");

            var skippedCountByWost = new Dictionary<Guid, (string DisplayName, int Count)>();

            try
            {
                // Phase 3: Process ts_questionresponse only for WOSTs that have exemption data (response already in cache).
                for (int i = 0; i < wostIdsWithExemptions.Count; i += inClauseMax)
                {
                    var chunk = wostIdsWithExemptions.Skip(i).Take(inClauseMax).ToList();
                    var query = new QueryExpression("ts_questionresponse")
                    {
                        ColumnSet = new ColumnSet("ts_questionresponseid", "ts_msdyn_workorderservicetask", "ts_questionname", "ts_exemptions", "ts_name"),
                        Criteria = new FilterExpression { Conditions = { new ConditionExpression("ts_msdyn_workorderservicetask", ConditionOperator.In, chunk) } },
                        PageInfo = new PagingInfo { PageNumber = 1, Count = pageSize, PagingCookie = null }
                    };
                    while (true)
                    {
                        var collection = _service.RetrieveMultiple(query);
                        if (collection.Entities.Count == 0) break;

                        result.TotalScanned += collection.Entities.Count;
                        _logger.Info($"Scanning page {query.PageInfo.PageNumber}: {collection.Entities.Count} ts_questionresponse records (total so far: {result.TotalScanned}/{totalToScan}).");

                        foreach (var qr in collection.Entities)
                        {
                            var qrId = qr.Id;
                            var name = qr.GetAttributeValue<string>("ts_name");
                            var questionName = qr.GetAttributeValue<string>("ts_questionname");
                            var existingExemptions = qr.GetAttributeValue<string>("ts_exemptions");
                            var wostRef = qr.GetAttributeValue<EntityReference>("ts_msdyn_workorderservicetask");
                            if (wostRef == null) { result.SkippedNoWost++; continue; }
                            if (string.IsNullOrEmpty(questionName)) { result.SkippedNoQuestionName++; continue; }

                            if (!wostResponseCache.TryGetValue(wostRef.Id, out var responseObj))
                                continue; // Should not happen: we only query WOSTs in cache.

                            var newValue = GetExemptionValueFromResponse(responseObj, questionName);

                            bool isEmptyNew = string.IsNullOrWhiteSpace(newValue) || newValue.Trim() == "[]";
                            if (!isEmptyNew)
                                result.SourceHadNonEmptyExemptions++;

                            bool existingIsEmpty = string.IsNullOrWhiteSpace(existingExemptions) || existingExemptions.Trim() == "[]";
                            if (isEmptyNew && existingIsEmpty)
                            {
                                result.SkippedNoChange++;
                                var bracketIdx = name != null ? name.IndexOf(" [", StringComparison.Ordinal) : -1;
                                var wostDisplayName = (bracketIdx > 0) ? name.Substring(0, bracketIdx) : (name ?? wostRef.Id.ToString());
                                if (!skippedCountByWost.TryGetValue(wostRef.Id, out var pair))
                                    skippedCountByWost[wostRef.Id] = (wostDisplayName, 1);
                                else
                                    skippedCountByWost[wostRef.Id] = (pair.DisplayName, pair.Count + 1);
                                continue;
                            }

                            var summary = SummarizeExemptionsForLog(newValue);
                            _logger.Verbose($"Backfilling ts_exemptions on ts_questionresponse {qrId} ({name}) with {summary}.");

                            if (!simulationMode)
                            {
                                var update = new Entity("ts_questionresponse", qrId)
                                {
                                    ["ts_exemptions"] = newValue.Trim() == "[]" ? "[]" : newValue
                                };
                                _service.Update(update);
                            }
                            result.Updated++;
                        }

                        if (!collection.MoreRecords) break;
                        query.PageInfo.PageNumber++;
                        query.PageInfo.PagingCookie = collection.PagingCookie;
                    }
                }

                foreach (var kv in skippedCountByWost)
                    _logger.Info($"Skipped {kv.Value.Count} ts_exemptions update(s) for WOST {kv.Value.DisplayName}: no exemptions and records already empty.");

                _logger.Info($"Exemption backfill completed. Scanned: {result.TotalScanned}, keys with exemptions in WOST JSON: {result.KeysWithExemptionsInJson}, matched non-empty: {result.SourceHadNonEmptyExemptions}, updated: {result.Updated}, skipped (no change/empty): {result.SkippedNoChange}, skipped (no WOST): {result.SkippedNoWost}, skipped (no question name): {result.SkippedNoQuestionName}.");
                return result;
            }
            catch (Exception ex)
            {
                _logger.Error($"Error during backfill of ts_exemptions: {ex}");
                throw;
            }
        }

        /// <summary>
        /// Returns a short summary of exemption JSON for logging (e.g. "EXM-001222, EXM-001324 (4 exemption(s))") instead of full JSON.
        /// </summary>
        private static string SummarizeExemptionsForLog(string exemptionJson)
        {
            if (string.IsNullOrWhiteSpace(exemptionJson) || exemptionJson.Trim() == "[]")
                return "[]";
            try
            {
                var arr = JArray.Parse(exemptionJson);
                if (arr.Count == 0) return "[]";
                var names = new List<string>();
                foreach (var item in arr)
                {
                    if (item is JObject obj && obj["exemptionName"] != null)
                        names.Add(obj["exemptionName"].ToString());
                }
                return names.Count > 0
                    ? $"{string.Join(", ", names)} ({arr.Count} exemption(s))"
                    : $"{arr.Count} exemption(s)";
            }
            catch
            {
                return $"{exemptionJson.Length} chars";
            }
        }

        /// <summary>
        /// Gets exemption JSON for a question from WOST response. Key is exactly "questionName-Exemptions" per survey client.
        /// </summary>
        private static string GetExemptionValueFromResponse(JObject responseObj, string questionName)
        {
            var token = responseObj[questionName + "-Exemptions"];
            if (token != null && token is JArray arr)
                return arr.ToString();
            return "[]";
        }
    }

    public class BackfillResult
    {
        public int TotalScanned { get; set; }
        public int Updated { get; set; }
        public int SkippedNoWorkOrder { get; set; }
    }

    public class ExemptionBackfillResult
    {
        public int TotalScanned { get; set; }
        /// <summary>Number of question responses where the WOST ovs_questionnaireresponse had a non-empty "questionName-Exemptions" array.</summary>
        public int SourceHadNonEmptyExemptions { get; set; }
        /// <summary>Total count of keys ending with "-Exemptions" with non-empty array found in WOST JSON (diagnostic; may be &gt; 0 even when updates are 0 if key doesn't match any ts_questionname).</summary>
        public int KeysWithExemptionsInJson { get; set; }
        public int Updated { get; set; }
        public int SkippedNoChange { get; set; }
        public int SkippedNoWost { get; set; }
        public int SkippedNoQuestionName { get; set; }
    }
}

