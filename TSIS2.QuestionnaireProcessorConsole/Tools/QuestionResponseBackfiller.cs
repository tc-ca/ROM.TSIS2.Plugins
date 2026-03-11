using System;
using System.Collections.Generic;
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
        /// Backfills ts_exemptions from WOST ovs_questionnaireresponse (questionName-Exemptions).
        /// Skips Update when the value to write is [] and the record already has null/empty/[] to avoid unnecessary writes.
        /// </summary>
        public ExemptionBackfillResult BackfillExemptions(bool simulationMode)
        {
            var result = new ExemptionBackfillResult();
            _logger.Info("Starting backfill of ts_exemptions on ts_questionresponse records...");
            _logger.Info($"Simulation mode: {(simulationMode ? "ON (no updates will be committed)" : "OFF")}");

            var query = new QueryExpression("ts_questionresponse")
            {
                ColumnSet = new ColumnSet("ts_questionresponseid", "ts_msdyn_workorderservicetask", "ts_questionname", "ts_exemptions", "ts_name"),
                Criteria = new FilterExpression
                {
                    Conditions = { new ConditionExpression("ts_msdyn_workorderservicetask", ConditionOperator.NotNull) }
                },
                PageInfo = new PagingInfo { PageNumber = 1, Count = 5000, PagingCookie = null }
            };

            var wostResponseCache = new Dictionary<Guid, JObject>();
            var skippedCountByWost = new Dictionary<Guid, (string DisplayName, int Count)>();

            try
            {
                while (true)
                {
                    var collection = _service.RetrieveMultiple(query);
                    if (collection.Entities.Count == 0) break;

                    result.TotalScanned += collection.Entities.Count;
                    _logger.Info($"Scanning page {query.PageInfo.PageNumber}: {collection.Entities.Count} ts_questionresponse records.");

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
                        {
                            var wost = _repository.GetWorkOrderServiceTask(wostRef.Id);
                            var responseJson = wost.GetAttributeValue<string>("ovs_questionnaireresponse");
                            responseObj = !string.IsNullOrWhiteSpace(responseJson) ? JObject.Parse(responseJson) : new JObject();
                            wostResponseCache[wostRef.Id] = responseObj;
                        }

                        var exemptionKey = questionName + "-Exemptions";
                        var exemptionToken = responseObj[exemptionKey];
                        var newValue = (exemptionToken != null && exemptionToken is JArray arr)
                            ? arr.ToString()
                            : "[]";

                        bool isEmptyNew = string.IsNullOrWhiteSpace(newValue) || newValue.Trim() == "[]";
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

                        _logger.Verbose($"Backfilling ts_exemptions on ts_questionresponse {qrId} ({name}) with {newValue}.");

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

                foreach (var kv in skippedCountByWost)
                    _logger.Info($"Skipped {kv.Value.Count} ts_exemptions update(s) for WOST {kv.Value.DisplayName}: no exemptions and records already empty.");

                _logger.Info($"Exemption backfill completed. Scanned: {result.TotalScanned}, updated: {result.Updated}, skipped (no change/empty): {result.SkippedNoChange}, skipped (no WOST): {result.SkippedNoWost}, skipped (no question name): {result.SkippedNoQuestionName}.");
                return result;
            }
            catch (Exception ex)
            {
                _logger.Error($"Error during backfill of ts_exemptions: {ex}");
                throw;
            }
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
        public int Updated { get; set; }
        public int SkippedNoChange { get; set; }
        public int SkippedNoWost { get; set; }
        public int SkippedNoQuestionName { get; set; }
    }
}

