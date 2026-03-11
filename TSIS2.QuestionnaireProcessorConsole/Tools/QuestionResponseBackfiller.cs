using System;
using System.Collections.Generic;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using TSIS2.Plugins.QuestionnaireProcessor;

namespace TSIS2.QuestionnaireProcessorConsole
{
    /// <summary>
    /// One-off utility to backfill ts_workorder on existing ts_questionresponse records.
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

        public BackfillResult BackfillExemptions(bool simulationMode)
        {
            var result = new BackfillResult();

            _logger.Info("Starting backfill of ts_exemptions on ts_questionresponse records...");
            _logger.Info($"Simulation mode: {(simulationMode ? "ON (no updates will be committed)" : "OFF")}");

            if (!_repository.HasQuestionResponseExemptionsColumn())
            {
                _logger.Warning("Column ts_questionresponse.ts_exemptions is not available. Create the exemption field before running this backfill.");
                result.SkippedMissingField = 1;
                return result;
            }

            var query = new QueryExpression("ts_questionresponse")
            {
                ColumnSet = new ColumnSet("ts_questionresponseid", "ts_name", "ts_questionname", "ts_msdyn_workorderservicetask"),
                Criteria = new FilterExpression
                {
                    Conditions =
                    {
                        new ConditionExpression("ts_msdyn_workorderservicetask", ConditionOperator.NotNull)
                    },
                    Filters =
                    {
                        new FilterExpression(LogicalOperator.Or)
                        {
                            Conditions =
                            {
                                new ConditionExpression("ts_exemptions", ConditionOperator.Null),
                                new ConditionExpression("ts_exemptions", ConditionOperator.Equal, string.Empty)
                            }
                        }
                    }
                },
                PageInfo = new PagingInfo
                {
                    PageNumber = 1,
                    Count = 5000,
                    PagingCookie = null
                }
            };

            var wostResponseCache = new Dictionary<Guid, QuestionnaireResponse>();

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

                    _logger.Info($"Scanning page {query.PageInfo.PageNumber}: {collection.Entities.Count} ts_questionresponse records for exemption backfill.");

                    foreach (var qr in collection.Entities)
                    {
                        var qrId = qr.Id;
                        var name = qr.GetAttributeValue<string>("ts_name");
                        var questionName = qr.GetAttributeValue<string>("ts_questionname");
                        var wostRef = qr.GetAttributeValue<EntityReference>("ts_msdyn_workorderservicetask");

                        if (wostRef == null)
                        {
                            _logger.Warning($"QuestionResponse {qrId} ({name}) has no ts_msdyn_workorderservicetask, skipping.");
                            result.SkippedNoWost++;
                            continue;
                        }

                        if (string.IsNullOrWhiteSpace(questionName))
                        {
                            _logger.Warning($"QuestionResponse {qrId} ({name}) has no ts_questionname, skipping.");
                            result.SkippedNoQuestionName++;
                            continue;
                        }

                        if (!wostResponseCache.TryGetValue(wostRef.Id, out var questionnaireResponse))
                        {
                            var wost = _repository.GetWorkOrderServiceTask(wostRef.Id);
                            var responseJson = wost.GetAttributeValue<string>("ovs_questionnaireresponse");
                            questionnaireResponse = string.IsNullOrWhiteSpace(responseJson)
                                ? null
                                : new QuestionnaireResponse(responseJson, _logger);
                            wostResponseCache[wostRef.Id] = questionnaireResponse;
                        }

                        var exemptionsJson = QuestionnaireExemptionSerializer.SerializeCompact(
                            questionnaireResponse?.GetExemptionValues(questionName),
                            _logger);

                        if (exemptionsJson == null)
                        {
                            exemptionsJson = "[]";
                            result.DefaultedEmptyExemptions++;
                        }

                        _logger.Debug($"Backfilling ts_exemptions on ts_questionresponse {qrId} ({name}) with {exemptionsJson}.");

                        if (!simulationMode)
                        {
                            var update = new Entity("ts_questionresponse", qrId)
                            {
                                ["ts_exemptions"] = exemptionsJson
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

                _logger.Info($"Exemption backfill completed. Total scanned: {result.TotalScanned}, updated: {result.Updated}, defaulted to empty JSON: {result.DefaultedEmptyExemptions}, skipped (no WOST): {result.SkippedNoWost}, skipped (no question name): {result.SkippedNoQuestionName}.");
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
        public int SkippedNoWost { get; set; }
        public int SkippedNoQuestionName { get; set; }
        public int DefaultedEmptyExemptions { get; set; }
        public int SkippedMissingField { get; set; }
    }
}

