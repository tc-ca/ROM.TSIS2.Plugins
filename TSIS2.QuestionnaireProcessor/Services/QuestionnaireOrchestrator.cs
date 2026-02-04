using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.PluginTelemetry;
using Newtonsoft.Json.Linq;

namespace TSIS2.Plugins.QuestionnaireProcessor
{
    /// <summary>
    /// Represents a single row that will be shown in the verbose question‑inventory report.
    /// </summary>
    public class QuestionLogEntry
    {
        public string QuestionName { get; set; }
        public string DynamicsName { get; set; }
        public string Status { get; set; } // e.g., Created, Updated, Up-to-date, MergedTruncateToCrmTextLimitTruncateToCrmTextLimit
        public bool IsMerged { get; set; }
        public int? SequenceNumber { get; set; }
    }

    /// <summary>
    /// Represents the results of processing a questionnaire.
    /// </summary>
    public class QuestionnaireProcessResult
    {
        public List<Guid> CreatedResponseIds { get; set; } = new List<Guid>();
        public int UpdatedRecordsCount { get; set; }
        public int UpToDateRecordsCount { get; set; }
        public int HiddenMergedCount { get; set; }

        public int VisibleQuestionCount { get; set; }

        public int TotalCrmRecords => CreatedResponseIds.Count + UpdatedRecordsCount + UpToDateRecordsCount;
        public int TotalCreatedOrUpdatedRecords => CreatedResponseIds.Count + UpdatedRecordsCount;

        public List<QuestionLogEntry> Inventory { get; set; } = new List<QuestionLogEntry>();
    }

    /// <summary>
    /// Orchestrates the questionnaire processing workflow, managing the multi-pass process.
    /// </summary>
    public static class QuestionnaireOrchestrator
    {
        private const int CrmTextMaxLength = 4000;
        private const int CrmNameMaxLength = 100;

        /// <summary>
        /// Processes a single questionnaire for a Work Order Service Task.
        /// </summary>
        /// <param name="service">The CRM organization service.</param>
        /// <param name="workOrderServiceTaskId">The ID of the Work Order Service Task.</param>
        /// <param name="questionnaireRef">The questionnaire reference.</param>
        /// <param name="isRecompletion">Whether this is a recompletion.</param>
        /// <param name="simulationMode">Whether to run in simulation mode.</param>
        /// <param name="logger">The logging service.</param>
        /// <returns>A list of created question response IDs.</returns>
        public static QuestionnaireProcessResult ProcessQuestionnaire(IOrganizationService service, Guid workOrderServiceTaskId, EntityReference questionnaireRef, bool isRecompletion, bool simulationMode = false, ILoggingService logger = null, bool includeQuestionInventory = false)
        {
            if (logger == null)
            {
                logger = new LoggerAdapter();
            }

            var repository = new QuestionnaireRepository(service, logger);
            var formatter = new QuestionnaireResponseFormatter(logger);

            return ProcessSingle(repository, formatter, workOrderServiceTaskId, questionnaireRef, isRecompletion, simulationMode, includeQuestionInventory);
        }

        /// <summary>
        /// Internal method that processes a single questionnaire.
        /// </summary>
        /// <param name="repository">The questionnaire repository.</param>
        /// <param name="formatter">The response formatter.</param>
        /// <param name="workOrderServiceTaskId">The ID of the Work Order Service Task.</param>
        /// <param name="questionnaireRef">The questionnaire reference.</param>
        /// <param name="isRecompletion">Whether this is a recompletion.</param>
        /// <param name="simulationMode">Whether to run in simulation mode.</param>
        /// <returns>A list of created question response IDs.</returns>
        private static QuestionnaireProcessResult ProcessSingle(QuestionnaireRepository repository, QuestionnaireResponseFormatter formatter, Guid workOrderServiceTaskId, EntityReference questionnaireRef, bool isRecompletion, bool simulationMode, bool includeQuestionInventory)
        {
            var result = new QuestionnaireProcessResult();

            try
            {
                // Step 1: Get the WOST data using the repository
                var wost = repository.GetWorkOrderServiceTask(workOrderServiceTaskId);
                string responseJson = wost.GetAttributeValue<string>("ovs_questionnaireresponse");
                string definitionJson = wost.GetAttributeValue<string>("ovs_questionnairedefinition");
                string wostName = wost.GetAttributeValue<string>("msdyn_name");
                var workOrderRef = wost.GetAttributeValue<EntityReference>("msdyn_workorder");

                if (string.IsNullOrEmpty(responseJson) || string.IsNullOrEmpty(definitionJson))
                {
                    formatter.Logger.Info($"No questionnaire data found for WOST: {wostName}");
                    return result;
                }

                // Step 2: Create instances of QuestionnaireDefinition and QuestionnaireResponse
                var questionnaireDefinition = new QuestionnaireDefinition(definitionJson, formatter.Logger);
                var questionnaireResponse = new QuestionnaireResponse(responseJson, formatter.Logger);

                // Step 3: Get existing responses using the repository
                var existingByName = repository.GetExistingResponses(workOrderServiceTaskId);

                // Step 4: Get visible questions from the definition
                var visibleQuestions = questionnaireDefinition.CollectVisibleQuestions(questionnaireResponse.Response);

                // Step 5: Pre-calculate all merged details (comments + hidden sub-questions)
                // This allows Pass 1 to perform correct 'is dirty' checks for the ts_details field.
                var mergedDetailsMap = CalculateMergedDetails(formatter, questionnaireResponse, questionnaireDefinition, visibleQuestions);

                // Step 6: Execute Pass 1 (Create/Update)
                var (questionIdMap, createdIds, processedIds, updatedCount) = ExecutePass1_CreateUpdate(
                    repository,
                    formatter,
                    questionnaireResponse,
                    questionnaireDefinition,
                    visibleQuestions,
                    wostName,
                    workOrderServiceTaskId,
                    workOrderRef,
                    questionnaireRef,
                    isRecompletion,
                    simulationMode,
                    existingByName,
                    mergedDetailsMap,
                    result.Inventory,
                    includeQuestionInventory);

                result.CreatedResponseIds.AddRange(createdIds);
                result.VisibleQuestionCount = visibleQuestions.Count;
                result.UpdatedRecordsCount = updatedCount;
                result.HiddenMergedCount = result.Inventory.Count(i => i.IsMerged);
                result.UpToDateRecordsCount = processedIds.Count - createdIds.Count - updatedCount;

                // Step 7: Execute Pass 2 (Link Findings)
                // We no longer need to collect hidden responses here as they are processed in Pass 1.
                ExecutePass2_LinkFindings(repository, formatter, visibleQuestions, questionIdMap, simulationMode, existingByName);

                // Step 8: Execute Final Pass (Deactivate)
                if (isRecompletion && !simulationMode)
                {
                    ExecuteFinalPass_Deactivate(repository, existingByName, processedIds);
                }

                return result;
            }
            catch (Exception ex)
            {
                formatter.Logger.Error($"Error processing questionnaire: {ex}");
                throw;
            }
        }

        /// <summary>
        /// Executes Pass 1: Create/Update question responses.
        /// </summary>
        private static (Dictionary<string, Guid> questionIdMap, List<Guid> createdIds, HashSet<Guid> processedIds, int updatedCount) ExecutePass1_CreateUpdate(
            QuestionnaireRepository repository,
            QuestionnaireResponseFormatter formatter,
            QuestionnaireResponse response,
            QuestionnaireDefinition definition,
            List<JObject> visibleQuestions,
            string wostName,
            Guid workOrderServiceTaskId,
            EntityReference workOrderRef,
            EntityReference questionnaireRef,
            bool isRecompletion,
            bool simulationMode,
            Dictionary<string, Entity> existingByName,
            Dictionary<string, string> mergedDetailsMap,
            List<QuestionLogEntry> inventory,
            bool includeQuestionInventory)
        {
            var questionIdMap = new Dictionary<string, Guid>();
            var createdIds = new List<Guid>();
            var processedIds = new HashSet<Guid>();
            int updatedCount = 0;
            int questionNumber = 1;

            formatter.Logger.Processing($"Starting first pass - {(isRecompletion ? "updating/creating" : "creating")} question responses based on visibility");

            // Only questions used in the current questionnaire are updated.
            // Questions no longer used are deactivated (not deleted) and may be reused later.
            foreach (var questionDefinition in visibleQuestions)
            {
                var questionName = questionDefinition["name"]?.ToString();
                if (string.IsNullOrEmpty(questionName)) continue;

                bool isHidden = questionDefinition["hideNumber"]?.ToObject<bool>() == true;
                bool hasDependency = !string.IsNullOrEmpty(questionDefinition["visibleIf"]?.ToString());

                if (isHidden)
                {
                    if (!hasDependency)
                    {
                        // Root hidden question - create non-numbered record
                        formatter.Logger.Trace($"Found visible root hidden question '{questionName}'. Processing as a non-numbered record.");
                        var (uCount, status) = CreateNonNumberedRecord(
                            repository,
                            formatter,
                            response,
                            questionDefinition, // Changed from questionName to questionDefinition
                            wostName,
                            workOrderServiceTaskId,
                            workOrderRef,
                            questionnaireRef,
                            simulationMode,
                            questionIdMap,
                            createdIds,
                            processedIds,
                            existingByName,
                            mergedDetailsMap);

                        updatedCount += uCount;
                        if (includeQuestionInventory)
                        {
                            inventory.Add(new QuestionLogEntry
                            {
                                QuestionName = questionName,
                                DynamicsName = $"{wostName} [{questionName}]",
                                Status = status,
                                IsMerged = false
                            });
                        }
                    }
                    else
                    {
                        // Dependent hidden question - data was already aggregated into parents in CalculateMergedDetails
                        formatter.Logger.Trace($"Skipping visible dependent hidden question '{questionName}'. Its data has already been merged into its parent's details.");
                        if (includeQuestionInventory)
                        {
                            inventory.Add(new QuestionLogEntry
                            {
                                QuestionName = questionName,
                                DynamicsName = "[Merged into parent]",
                                Status = "Merged",
                                IsMerged = true
                            });
                        }
                    }
                    continue;
                }

                int currentNumber = questionNumber++;
                string itemStatus = "Up-to-date";

                // Collect New Values
                var responseValue = response.GetValue(questionName);
                string questionType = questionDefinition["type"]?.ToString();
                string responseText = responseValue != null
                    ? formatter.FormatResponse(responseValue, questionType, questionDefinition)
                    : string.Empty;

                // Details are now pre-calculated (includes comments and merged hidden questions)
                mergedDetailsMap.TryGetValue(questionName, out string currentDetailsToSet);

                var titleEn = formatter.RemoveHtmlTags(QuestionnaireDefinition.GetTextFieldValue(questionDefinition["title"], "default"));
                var titleFr = formatter.RemoveHtmlTags(QuestionnaireDefinition.GetTextFieldValue(questionDefinition["title"], "fr"));
                var descriptionEn = formatter.RemoveHtmlTags(QuestionnaireDefinition.GetTextFieldValue(questionDefinition["description"], "default"));
                var descriptionFr = formatter.RemoveHtmlTags(QuestionnaireDefinition.GetTextFieldValue(questionDefinition["description"], "fr"));
                var provisionRef = questionDefinition["provision"]?.ToString();

                // Check if record already exists
                if (existingByName.TryGetValue(questionName, out var existingResponse))
                {
                    if (!questionIdMap.ContainsKey(questionName))
                        questionIdMap[questionName] = existingResponse.Id;

                    processedIds.Add(existingResponse.Id);

                    if (!simulationMode)
                    {
                        // Comparison Logic - use StringsAreEqual to handle null vs empty string
                        bool answerChanged = !StringsAreEqual(existingResponse.GetAttributeValue<string>("ts_response"), responseText);
                        bool detailsChanged = !StringsAreEqual(existingResponse.GetAttributeValue<string>("ts_details"), currentDetailsToSet);
                        bool numberChanged = existingResponse.GetAttributeValue<int?>("ts_questionnumber") != currentNumber;

                        // Metadata changes (Title, Description, Provision)
                        bool metadataChanged = !StringsAreEqual(existingResponse.GetAttributeValue<string>("ts_questiontextenglish"), titleEn)
                                            || !StringsAreEqual(existingResponse.GetAttributeValue<string>("ts_questiontextfrench"), titleFr)
                                            || !StringsAreEqual(existingResponse.GetAttributeValue<string>("ts_provisionreference"), provisionRef);

                        if (answerChanged || detailsChanged || numberChanged || metadataChanged)
                        {
                            updatedCount++;
                            int currentVersion = existingResponse.GetAttributeValue<int?>("ts_version") ?? 1;
                            bool bumpVersion = answerChanged || detailsChanged;

                            // Build a concise change summary
                            var changes = new List<string>();
                            if (answerChanged) changes.Add("Answer");
                            if (detailsChanged) changes.Add("Details");
                            if (numberChanged) changes.Add("Number");
                            if (metadataChanged) changes.Add("Metadata");

                            // Build detailed status for inventory
                            itemStatus = bumpVersion
                                ? $"Updated ({string.Join("+", changes)}) v{currentVersion}→v{currentVersion + 1}"
                                : $"Updated ({string.Join("+", changes)})";

                            formatter.Logger.Verbose($"Updating '{questionName}': {string.Join(", ", changes)} changed{(bumpVersion ? " (version bump)" : "")}");

                            CreateQuestionResponseRecord(
                                repository,
                                formatter,
                                questionName,
                                titleEn,
                                titleFr,
                                responseText,
                                provisionRef,
                                descriptionEn,
                                descriptionFr,
                                currentDetailsToSet,
                                wostName,
                                workOrderServiceTaskId,
                                workOrderRef,
                                questionnaireRef,
                                currentNumber,
                                responseValue,
                                true, // isUpdate
                                existingResponse.Id,
                                bumpVersion ? currentVersion : -1
                            );
                        }
                    }

                    if (includeQuestionInventory)
                    {
                        inventory.Add(new QuestionLogEntry
                        {
                            QuestionName = questionName,
                            DynamicsName = $"{wostName} [{currentNumber}]",
                            Status = itemStatus,
                            IsMerged = false,
                            SequenceNumber = currentNumber
                        });
                    }
                    continue;
                }

                // Create new record
                formatter.Logger.Verbose($"Creating question response record: {wostName} [{currentNumber}] | '{questionName}'");
                itemStatus = "Created";

                Guid questionResponseId;
                if (simulationMode)
                {
                    questionResponseId = Guid.NewGuid();
                }
                else
                {
                    questionResponseId = CreateQuestionResponseRecord(
                        repository,
                        formatter,
                        questionName,
                        titleEn,
                        titleFr,
                        responseText,
                        provisionRef,
                        descriptionEn,
                        descriptionFr,
                        currentDetailsToSet,
                        wostName,
                        workOrderServiceTaskId,
                        workOrderRef,
                        questionnaireRef,
                        currentNumber,
                        responseValue,
                        false,
                        default,
                        0);
                }

                processedIds.Add(questionResponseId);
                questionIdMap[questionName] = questionResponseId;
                createdIds.Add(questionResponseId);

                if (includeQuestionInventory)
                {
                    inventory.Add(new QuestionLogEntry
                    {
                        QuestionName = questionName,
                        DynamicsName = $"{wostName} [{currentNumber}]",
                        Status = itemStatus,
                        IsMerged = false,
                        SequenceNumber = currentNumber
                    });
                }
            }

            return (questionIdMap, createdIds, processedIds, updatedCount);
        }

        /// <summary>
        /// Creates a non-numbered record for root hidden questions.
        /// </summary>
        private static (int uCount, string status) CreateNonNumberedRecord(
            QuestionnaireRepository repository,
            QuestionnaireResponseFormatter formatter,
            QuestionnaireResponse response,
            JObject questionDefinition,
            string wostName,
            Guid workOrderServiceTaskId,
            EntityReference workOrderRef,
            EntityReference questionnaireRef,
            bool simulationMode,
            Dictionary<string, Guid> questionIdMap,
            List<Guid> createdIds,
            HashSet<Guid> processedIds,
            Dictionary<string, Entity> existingByName,
            Dictionary<string, string> mergedDetailsMap)
        {
            var questionName = questionDefinition["name"]?.ToString();
            mergedDetailsMap.TryGetValue(questionName, out string currentDetailsToSet);

            if (existingByName.TryGetValue(questionName, out var existingRecord))
            {
                formatter.Logger.Trace($"Record for root hidden question '{questionName}' already exists.");
                if (!questionIdMap.ContainsKey(questionName))
                {
                    questionIdMap[questionName] = existingRecord.Id;
                }
                processedIds.Add(existingRecord.Id);

                // Root hidden questions also need 'Is Dirty' check if recompletion
                if (!simulationMode)
                {
                    var respValue = response.GetValue(questionName);
                    string qType = questionDefinition["type"]?.ToString();
                    string respText = respValue != null
                        ? formatter.FormatResponse(respValue, qType, questionDefinition)
                        : string.Empty;

                    var tEn = formatter.RemoveHtmlTags(QuestionnaireDefinition.GetTextFieldValue(questionDefinition["title"], "default"));
                    var tFr = formatter.RemoveHtmlTags(QuestionnaireDefinition.GetTextFieldValue(questionDefinition["title"], "fr"));

                    // Use StringsAreEqual to handle null vs empty string
                    bool answerChanged = !StringsAreEqual(existingRecord.GetAttributeValue<string>("ts_response"), respText);
                    bool detailsChanged = !StringsAreEqual(existingRecord.GetAttributeValue<string>("ts_details"), currentDetailsToSet);
                    bool metadataChanged = !StringsAreEqual(existingRecord.GetAttributeValue<string>("ts_questiontextenglish"), tEn)
                                        || !StringsAreEqual(existingRecord.GetAttributeValue<string>("ts_questiontextfrench"), tFr);

                    if (answerChanged || detailsChanged || metadataChanged)
                    {
                        int currentVersion = existingRecord.GetAttributeValue<int>("ts_version");
                        bool bumpVersion = answerChanged || detailsChanged;

                        // Build a concise change summary
                        var changes = new List<string>();
                        if (answerChanged) changes.Add("Answer");
                        if (detailsChanged) changes.Add("Details");
                        if (metadataChanged) changes.Add("Metadata");

                        formatter.Logger.Verbose($"Updating hidden '{questionName}': {string.Join(", ", changes)} changed{(bumpVersion ? " (version bump)" : "")}");

                        CreateQuestionResponseRecord(
                            repository,
                            formatter,
                            questionName,
                            tEn,
                            tFr,
                            respText,
                            null,
                            null,
                            null,
                            currentDetailsToSet,
                            wostName,
                            workOrderServiceTaskId,
                            workOrderRef,
                            questionnaireRef,
                            null,
                            null,
                            true,
                            existingRecord.Id,
                            bumpVersion ? currentVersion : -1
                        );
                        // Build detailed status for inventory
                        var status = bumpVersion
                            ? $"Updated ({string.Join("+", changes)}) v{currentVersion}→v{currentVersion + 1}"
                            : $"Updated ({string.Join("+", changes)})";
                        return (1, status);
                    }
                }
                return (0, "Up-to-date");
            }

            var responseValue = response.GetValue(questionName);

            string questionType = questionDefinition["type"]?.ToString();
            string responseText = responseValue != null
                ? formatter.FormatResponse(responseValue, questionType, questionDefinition)
                : string.Empty;

            var titleEn = formatter.RemoveHtmlTags(QuestionnaireDefinition.GetTextFieldValue(questionDefinition["title"], "default"));
            var titleFr = formatter.RemoveHtmlTags(QuestionnaireDefinition.GetTextFieldValue(questionDefinition["title"], "fr"));

            Guid questionResponseId;
            if (simulationMode)
            {
                questionResponseId = Guid.NewGuid();
            }
            else
            {
                questionResponseId = CreateQuestionResponseRecord(
                    repository,
                    formatter,
                    questionName,
                    titleEn,
                    titleFr,
                    responseText,
                    null,
                    null,
                    null,
                    currentDetailsToSet,
                    wostName,
                    workOrderServiceTaskId,
                    workOrderRef,
                    questionnaireRef,
                    null,
                    null,
                    false,
                    default,
                    0);
            }

            processedIds.Add(questionResponseId);
            questionIdMap[questionName] = questionResponseId;
            createdIds.Add(questionResponseId);
            return (0, "Created");
        }

        /// <summary>
        /// Calculates all merged details (comments and hidden dependent questions) up front.
        /// This allows Pass 1 to perform correct versioning for hidden question changes.
        /// </summary>
        private static Dictionary<string, string> CalculateMergedDetails(
            QuestionnaireResponseFormatter formatter,
            QuestionnaireResponse response,
            QuestionnaireDefinition definition,
            List<JObject> visibleQuestions)
        {
            var detailMap = new Dictionary<string, List<JObject>>(StringComparer.OrdinalIgnoreCase);

            // 1. Process all visible questions for primary details (comments, findings, etc.)
            foreach (var qDef in visibleQuestions)
            {
                var qName = qDef["name"]?.ToString();
                if (string.IsNullOrEmpty(qName)) continue;

                var qType = qDef["type"]?.ToString();
                var hasDetail = response.HasDetailValue(qName);

                string comment = null;
                if (string.Equals(qType, "finding", StringComparison.OrdinalIgnoreCase))
                {
                    var respValue = response.GetValue(qName);
                    comment = respValue?["comments"]?.ToString();
                }
                else if (string.Equals(qType, "multipletext", StringComparison.OrdinalIgnoreCase))
                {
                    comment = formatter.FindMultipletextComment(definition, qDef, response);
                }

                if (comment == null && hasDetail)
                {
                    comment = response.GetDetailValue(qName);
                }

                if (!string.IsNullOrWhiteSpace(comment))
                {
                    // Clean up the comment text (convert newlines/tabs to spaces, remove HTML, etc.)
                    comment = formatter.RemoveHtmlTags(comment);
                    
                    if (!detailMap.ContainsKey(qName)) detailMap[qName] = new List<JObject>();
                    detailMap[qName].Add(new JObject
                    {
                        ["question"] = "Comment",
                        ["answer"] = comment
                    });
                }
            }

            // 2. Process dependent hidden questions and aggregate them into their parents' details
            foreach (var qDef in visibleQuestions.Where(q => !string.IsNullOrEmpty(q["visibleIf"]?.ToString())))
            {
                // Only merge questions that are flagged as 'hideNumber' (hidden dependent questions)
                if (qDef["hideNumber"]?.ToObject<bool>() != true) continue;

                var qName = qDef["name"]?.ToString();
                var visibleIf = qDef["visibleIf"].ToString();
                var parentQuestionName = QuestionnaireDefinition.ParseParentQuestionName(visibleIf);
                if (string.IsNullOrEmpty(parentQuestionName)) continue;

                var responseValue = response.GetValue(qName);
                if (responseValue != null && responseValue.Type != JTokenType.Null)
                {
                    string questionType = qDef["type"]?.ToString();
                    string responseText = formatter.FormatResponse(responseValue, questionType, qDef);
                    var title = formatter.RemoveHtmlTags(QuestionnaireDefinition.GetTextFieldValue(qDef["title"], "default"));

                    if (!string.IsNullOrEmpty(responseText))
                    {
                        var detailObject = new JObject
                        {
                            ["question"] = title.TrimEnd(' ', ':'),
                            ["answer"] = responseText
                        };

                        if (!detailMap.ContainsKey(parentQuestionName))
                        {
                            detailMap[parentQuestionName] = new List<JObject>();
                        }
                        detailMap[parentQuestionName].Add(detailObject);
                    }
                }
            }

            // 3. Convert JObject lists to JSON strings
            return detailMap.ToDictionary(
                kvp => kvp.Key,
                kvp => new JArray(kvp.Value).ToString(Newtonsoft.Json.Formatting.None),
                StringComparer.OrdinalIgnoreCase
            );
        }

        /// <summary>
        /// Executes Pass 2 Linking: Links findings to their parent questions.
        /// Merging is no longer handled here as it's been unified into Pass 1.
        /// </summary>
        private static void ExecutePass2_LinkFindings(
            QuestionnaireRepository repository, QuestionnaireResponseFormatter formatter,
            List<JObject> visibleQuestions, Dictionary<string, Guid> questionIdMap,
            bool simulationMode, Dictionary<string, Entity> existingByName)
        {
            formatter.Logger.Processing($"Starting second pass - linking finding relationships");

            foreach (var questionDefinition in visibleQuestions.Where(q => !string.IsNullOrEmpty(q["visibleIf"]?.ToString())))
            {
                // We only care about findings here; hidden dependent questions are already merged
                if (questionDefinition["hideNumber"]?.ToObject<bool>() == true) continue;

                var questionName = questionDefinition["name"]?.ToString();
                if (string.IsNullOrEmpty(questionName)) continue;

                var visibleIf = questionDefinition["visibleIf"].ToString();
                var parentQuestionName = QuestionnaireDefinition.ParseParentQuestionName(visibleIf);
                if (string.IsNullOrEmpty(parentQuestionName)) continue;

                // Find parent ID
                if (!questionIdMap.TryGetValue(parentQuestionName, out Guid parentId))
                {
                    if (existingByName.TryGetValue(parentQuestionName, out Entity parentEntity))
                    {
                        parentId = parentEntity.Id;
                    }
				}

                if (parentId == Guid.Empty)
                {
                    formatter.Logger.Warning($"Could not find parent record for '{parentQuestionName}' to link dependent question '{questionName}'.");
                    continue;
                }

                if (questionIdMap.TryGetValue(questionName, out Guid childId))
                {
                    if (simulationMode)
                    {
                        formatter.Logger.Verbose($"Linking finding {questionName} to parent question {parentQuestionName}");
                    }
                    else
                    {
                        formatter.Logger.Verbose($"Linking finding {questionName} to parent question {parentQuestionName}");
                        repository.LinkFindingToParent(childId, parentId);
                    }
                }
            }
        }

        /// <summary>
        /// Executes the final pass to deactivate orphaned records.
        /// </summary>
        private static void ExecuteFinalPass_Deactivate(QuestionnaireRepository repository, Dictionary<string, Entity> existingByName, HashSet<Guid> processedIds)
        {
            foreach (var kvp in existingByName)
            {
                Guid existingId = kvp.Value.Id;

                if (!processedIds.Contains(existingId))
                {
                    repository.DeactivateQuestionResponse(existingId);
                }
            }
        }

        private static Guid CreateQuestionResponseRecord(
            QuestionnaireRepository repository,
            QuestionnaireResponseFormatter formatter,
            string questionName,
            string questionTextEn,
            string questionTextFr,
            string responseText,
            string provisionReference,
            string provisionTextEn,
            string provisionTextFr,
            string details,
            string wostName,
            Guid workOrderServiceTaskId,
            EntityReference workOrderRef,
            EntityReference questionnaireRef,
            int? questionNumber,
            JToken findingObject,
            bool isUpdate,
            Guid existingId,
            int versionToSet)
        {
            var questionResponse = new Entity("ts_questionresponse");

            if (isUpdate)
            {
                questionResponse.Id = existingId;
                if (versionToSet != -1)
                {
                    questionResponse["ts_version"] = versionToSet + 1;
                }
                questionResponse["statecode"] = new OptionSetValue(0); // 0 = Active
            }
            else
            {
                questionResponse["ts_msdyn_workorderservicetask"] = new EntityReference("msdyn_workorderservicetask", workOrderServiceTaskId);
                if (workOrderRef != null)
                {
                    questionResponse["ts_workorder"] = workOrderRef;
                }
                questionResponse["ts_questionnaire"] = questionnaireRef;
                questionResponse["ts_version"] = 1;
                questionResponse["statecode"] = new OptionSetValue(0); // 0 = Active
            }

            // Common field mapping
            if (questionNumber.HasValue)
            {
                var name = $"{wostName} [{questionNumber.Value}]";
                questionResponse["ts_name"] = TruncateToCrmNameLimit(name, "ts_name", formatter.Logger);
                questionResponse["ts_questionnumber"] = questionNumber.Value;
            }
            else
            {
                var name = $"{wostName} [{questionName}]";
                questionResponse["ts_name"] = TruncateToCrmNameLimit(name, "ts_name", formatter.Logger);
                // ts_questionnumber is null for non-numbered
            }

            var provEnClean = formatter.RemoveHtmlTags(provisionTextEn);
            var provFrClean = formatter.RemoveHtmlTags(provisionTextFr);

            questionResponse["ts_questionname"] = TruncateToCrmTextLimit(questionName, "ts_questionname", formatter.Logger);
            questionResponse["ts_questiontextenglish"] = TruncateToCrmTextLimit(questionTextEn, "ts_questiontextenglish", formatter.Logger);
            questionResponse["ts_questiontextfrench"] = TruncateToCrmTextLimit(questionTextFr, "ts_questiontextfrench", formatter.Logger);
            questionResponse["ts_provisionreference"] = TruncateToCrmTextLimit(provisionReference, "ts_provisionreference", formatter.Logger);
            questionResponse["ts_provisiontextenglish"] = TruncateToCrmTextLimit(provEnClean, "ts_provisiontextenglish", formatter.Logger);
            questionResponse["ts_provisiontextfrench"] = TruncateToCrmTextLimit(provFrClean, "ts_provisiontextfrench", formatter.Logger);
            // Normalize empty/whitespace details to null to avoid false updates & version bumps.
            if (string.IsNullOrWhiteSpace(details))
            {
                questionResponse["ts_details"] = null;
            }
            else
            {
                questionResponse["ts_details"] =
                    TruncateToCrmTextLimit(details, $"ts_details ({questionName})", formatter.Logger);
            }

            if (!string.IsNullOrEmpty(responseText) &&
                !questionName.StartsWith("finding-", StringComparison.OrdinalIgnoreCase))
            {
                questionResponse["ts_response"] = TruncateToCrmTextLimit(responseText, $"ts_response ({questionName})", formatter.Logger);
            }

            if (!string.IsNullOrEmpty(responseText) &&
                !questionName.StartsWith("finding-", StringComparison.OrdinalIgnoreCase))
            {
                questionResponse["ts_response"] = TruncateToCrmTextLimit(responseText, $"ts_response ({questionName})", formatter.Logger);
            }

            // Handle findings
            if (questionName.StartsWith("finding-", StringComparison.OrdinalIgnoreCase))
            {
                var operations = findingObject?["operations"] as JArray;
                if (operations != null && operations.Count > 0)
                {
                    questionResponse["ts_operations"] =
                        string.Join(",", operations.Select(op => op["operationID"]?.ToString()).Where(s => !string.IsNullOrWhiteSpace(s)));

                    var findingTypeRaw = operations[0]?["findingType"]?.ToString();
                    if (!string.IsNullOrWhiteSpace(findingTypeRaw))
                    {
                        if (int.TryParse(findingTypeRaw, out var findingTypeInt))
                            questionResponse["ts_findingtype"] = new OptionSetValue(findingTypeInt);
                        else
                            formatter.Logger.Warning($"Invalid findingType '{findingTypeRaw}' for finding {questionName}");
                    }
                }
            }

            if (isUpdate)
            {
                repository.UpdateQuestionResponse(questionResponse);
                return existingId;
            }
            else
            {
                return repository.CreateQuestionResponse(questionResponse);
            }
        }

        private static string TruncateToCrmTextLimit(string value, string fieldName, ILoggingService logger)
        {
            return Truncate(value, CrmTextMaxLength, fieldName, logger);
        }

        private static string TruncateToCrmNameLimit(string value, string fieldName, ILoggingService logger)
        {
            return Truncate(value, CrmNameMaxLength, fieldName, logger);
        }

        /// <summary>
        /// Helpers: truncates text to Dynamics CRM max length
        /// </summary>
        private static string Truncate(string value, int maxLength, string fieldName, ILoggingService logger)
        {
            if (string.IsNullOrEmpty(value))
                return value;

            if (value.Length <= maxLength)
                return value;

            logger?.Warning($"CRM field '{fieldName}' truncated. OriginalLength={value.Length}, MaxLength={maxLength}");

            // Basic truncation
            var truncated = value.Substring(0, maxLength);

            // If the last char is a high surrogate (emoji)
            if (truncated.Length > 0 && char.IsHighSurrogate(truncated[truncated.Length - 1]))
                truncated = truncated.Substring(0, truncated.Length - 1);

            return truncated;
        }

        /// <summary>
        /// Compares two strings for equality, treating null and empty string as equivalent.
        /// This prevents false-positive "changed" detection when CRM returns null but code computes empty string.
        /// </summary>
        private static bool StringsAreEqual(string a, string b)
        {
            // Normalize: treat null and empty as the same
            var normalizedA = string.IsNullOrEmpty(a) ? null : a;
            var normalizedB = string.IsNullOrEmpty(b) ? null : b;
            return string.Equals(normalizedA, normalizedB, StringComparison.Ordinal);
        }

    }
}