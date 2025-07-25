using Microsoft.Xrm.Sdk;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;

namespace TSIS2.Plugins.QuestionnaireExtractor
{
    /// <summary>
    /// Orchestrates the entire questionnaire processing workflow.
    /// This class manages the multi-pass process in the correct sequence.
    /// </summary>
    public static class QuestionnaireOrchestrator
    {
        /// <summary>
        /// Processes a single questionnaire for a Work Order Service Task.
        /// This is the main entry point that maintains backward compatibility.
        /// </summary>
        /// <param name="service">The CRM organization service.</param>
        /// <param name="workOrderServiceTaskId">The ID of the Work Order Service Task.</param>
        /// <param name="questionnaireRef">The questionnaire reference.</param>
        /// <param name="isRecompletion">Whether this is a recompletion.</param>
        /// <param name="simulationMode">Whether to run in simulation mode.</param>
        /// <param name="logger">The logging service for logging.</param>
        /// <returns>A list of created question response IDs.</returns>
        public static List<Guid> ProcessQuestionnaire(IOrganizationService service, Guid workOrderServiceTaskId, EntityReference questionnaireRef, bool isRecompletion, bool simulationMode = false, ILoggingService logger = null)
        {
            if (logger == null)
            {
                throw new ArgumentNullException(nameof(logger), "Logger must be provided. Use TracingServiceAdapter to wrap ITracingService or provide another ILoggingService implementation.");
            }

            var repository = new QuestionnaireRepository(service, logger);
            var formatter = new QuestionnaireResponseFormatter(logger);
            
            return ProcessSingle(repository, formatter, workOrderServiceTaskId, questionnaireRef, isRecompletion, simulationMode);
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
        private static List<Guid> ProcessSingle(QuestionnaireRepository repository, QuestionnaireResponseFormatter formatter, Guid workOrderServiceTaskId, EntityReference questionnaireRef, bool isRecompletion, bool simulationMode)
        {
            var questionResponseIds = new List<Guid>();

            try
            {
                // Step 1: Get the WOST data using the repository
                var wost = repository.GetWorkOrderServiceTask(workOrderServiceTaskId);
                string responseJson = wost.GetAttributeValue<string>("ovs_questionnaireresponse");
                string definitionJson = wost.GetAttributeValue<string>("ovs_questionnairedefinition");
                string wostName = wost.GetAttributeValue<string>("msdyn_name");

                if (string.IsNullOrEmpty(responseJson) || string.IsNullOrEmpty(definitionJson))
                {
                    formatter.Logger.Trace($"No questionnaire data found for WOST: {wostName}");
                    return questionResponseIds;
                }

                // Step 2: Create instances of QuestionnaireDefinition and QuestionnaireResponse
                var questionnaireDefinition = new QuestionnaireDefinition(definitionJson, formatter.Logger);
                var questionnaireResponse = new QuestionnaireResponse(responseJson, formatter.Logger);

                // Step 3: Get existing responses using the repository
                var (existingByNumber, existingByNameAndNumber) = repository.GetExistingResponses(workOrderServiceTaskId);

                // Step 4: Get visible questions from the definition
                var visibleQuestions = questionnaireDefinition.CollectVisibleQuestions(questionnaireResponse.Response);

                // Step 5: Execute Pass 1 (Create/Update)
                var (questionIdMap, createdIds, processedIds) = ExecutePass1_CreateUpdate(
                    repository, formatter, questionnaireResponse, questionnaireDefinition, 
                    visibleQuestions, wostName, workOrderServiceTaskId, questionnaireRef, 
                    isRecompletion, simulationMode, existingByNumber, existingByNameAndNumber);

                questionResponseIds.AddRange(createdIds);

                // Step 6: Execute Pass 2 (Link/Merge)
                var detailResponsesToAppend = ExecutePass2_LinkMerge(
                    repository, formatter, questionnaireResponse, questionnaireDefinition,
                    visibleQuestions, questionIdMap, simulationMode, existingByNameAndNumber);

                // Step 7: Execute Merge Pass
                ExecuteMergePass(repository, detailResponsesToAppend, simulationMode);

                // Step 8: Execute Final Pass (Deactivate)
                if (isRecompletion && !simulationMode)
                {
                    ExecuteFinalPass_Deactivate(repository, existingByNumber, processedIds);
                }

                formatter.Logger.Trace($"Completed processing all questions and relationships");

                // Log summary
                if (questionResponseIds.Count == 0)
                {
                    if (existingByNumber.Any() || existingByNameAndNumber.Any())
                    {
                        formatter.Logger.Trace($"No new question responses generated for WOST {wostName} - all necessary records already exist.");
                    }
                    else
                    {
                        string reason = questionnaireResponse.DetermineNoResponseReason(questionnaireDefinition.GetAllQuestionNames());
                        formatter.Logger.Trace($"No question responses generated for WOST {wostName} - Reason: {reason}");
                    }
                }

                return questionResponseIds;
            }
            catch (Exception ex)
            {
                formatter.Logger.Error($"Error processing questionnaire: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Executes Pass 1: Create/Update question responses.
        /// </summary>
        private static (Dictionary<string, Guid> questionIdMap, List<Guid> createdIds, HashSet<Guid> processedIds) ExecutePass1_CreateUpdate(
            QuestionnaireRepository repository, QuestionnaireResponseFormatter formatter, 
            QuestionnaireResponse response, QuestionnaireDefinition definition, 
            List<JObject> visibleQuestions, string wostName, Guid workOrderServiceTaskId, 
            EntityReference questionnaireRef, bool isRecompletion, bool simulationMode,
            Dictionary<int, Entity> existingByNumber, Dictionary<QuestionnaireRepository.QuestionKey, Entity> existingByNameAndNumber)
        {
            var questionIdMap = new Dictionary<string, Guid>();
            var createdIds = new List<Guid>();
            var processedIds = new HashSet<Guid>();
            int questionNumber = 1;

            formatter.Logger.Trace($"Starting first pass - {(isRecompletion ? "updating/creating" : "creating")} question responses based on visibility");

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
                        CreateNonNumberedRecord(repository, formatter, response, definition, questionName, wostName, 
                            workOrderServiceTaskId, questionnaireRef, simulationMode, questionIdMap, createdIds, processedIds, existingByNameAndNumber);
                    }
                    else
                    {
                        // Dependent hidden question - will be merged later
                        formatter.Logger.Trace($"Skipping visible dependent hidden question '{questionName}'. It will be merged in Pass 2.");
                    }
                    continue;
                }

                // Check if record already exists
                var key = new QuestionnaireRepository.QuestionKey(questionName, questionNumber);
                if (existingByNameAndNumber.TryGetValue(key, out var existingResponse))
                {
                    formatter.Logger.Trace($"Record for visible question '{questionName}' #{questionNumber} already exists. Skipping creation.");
                    if (!questionIdMap.ContainsKey(questionName))
                    {
                        questionIdMap[questionName] = existingResponse.Id;
                    }
                    processedIds.Add(existingResponse.Id);
                    questionNumber++;
                    continue;
                }

                // Create new record
                formatter.Logger.Trace($"Creating new record for visible question '{questionName}' with number {questionNumber}.");
                
                var responseValue = response.GetValue(questionName);
                var hasDetail = response.HasDetailValue(questionName);

                string questionType = questionDefinition["type"]?.ToString();
                string responseText = responseValue != null
                    ? formatter.FormatResponse(responseValue, questionType, questionDefinition)
                    : string.Empty;

                string commentOrDetails = null;
                if (string.Equals(questionType, "finding", StringComparison.OrdinalIgnoreCase))
                    commentOrDetails = responseValue?["comments"]?.ToString();
                else if (string.Equals(questionType, "multipletext", StringComparison.OrdinalIgnoreCase))
                    commentOrDetails = formatter.FindMultipletextComment(definition, questionName, response);
                if (commentOrDetails == null && hasDetail)
                    commentOrDetails = response.GetDetailValue(questionName);

                // Get text field values
                var titleEn = formatter.RemoveHtmlTags(QuestionnaireDefinition.GetTextFieldValue(questionDefinition["title"], "default"));
                var titleFr = formatter.RemoveHtmlTags(QuestionnaireDefinition.GetTextFieldValue(questionDefinition["title"], "fr"));
                var descriptionEn = formatter.RemoveHtmlTags(QuestionnaireDefinition.GetTextFieldValue(questionDefinition["description"], "default"));
                var descriptionFr = formatter.RemoveHtmlTags(QuestionnaireDefinition.GetTextFieldValue(questionDefinition["description"], "fr"));
                var provisionRef = questionDefinition["provision"]?.ToString();

                Guid questionResponseId;
                if (simulationMode)
                {
                    questionResponseId = Guid.NewGuid();
                }
                else
                {
                    questionResponseId = CreateQuestionResponseRecord(repository, formatter, questionName, titleEn, titleFr, 
                        responseText, provisionRef, descriptionEn, descriptionFr, commentOrDetails, wostName, 
                        workOrderServiceTaskId, questionnaireRef, questionNumber, responseValue, false, default, 0);
                }

                processedIds.Add(questionResponseId);
                questionIdMap[questionName] = questionResponseId;
                createdIds.Add(questionResponseId);
                questionNumber++;
            }

            return (questionIdMap, createdIds, processedIds);
        }

        /// <summary>
        /// Creates a non-numbered record for root hidden questions.
        /// </summary>
        private static void CreateNonNumberedRecord(QuestionnaireRepository repository, QuestionnaireResponseFormatter formatter,
            QuestionnaireResponse response, QuestionnaireDefinition definition, string questionName, string wostName,
            Guid workOrderServiceTaskId, EntityReference questionnaireRef, bool simulationMode,
            Dictionary<string, Guid> questionIdMap, List<Guid> createdIds, HashSet<Guid> processedIds,
            Dictionary<QuestionnaireRepository.QuestionKey, Entity> existingByNameAndNumber)
        {
            var key = new QuestionnaireRepository.QuestionKey(questionName, null);
            if (existingByNameAndNumber.ContainsKey(key))
            {
                formatter.Logger.Trace($"Record for root hidden question '{questionName}' already exists. Skipping creation.");
                var existingRecord = existingByNameAndNumber[key];
                if (!questionIdMap.ContainsKey(questionName))
                {
                    questionIdMap[questionName] = existingRecord.Id;
                }
                return;
            }

            var questionDefinition = definition.FindQuestionDefinition(questionName);
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
                questionResponseId = CreateQuestionResponseRecord(repository, formatter, questionName, titleEn, titleFr,
                    responseText, null, null, null, null, wostName, workOrderServiceTaskId, questionnaireRef,
                    null, null, false, default, 0);
            }

            processedIds.Add(questionResponseId);
            questionIdMap[questionName] = questionResponseId;
            createdIds.Add(questionResponseId);
        }

        /// <summary>
        /// Executes Pass 2: Link findings and prepare hidden question data for merging.
        /// </summary>
        private static Dictionary<Guid, List<JObject>> ExecutePass2_LinkMerge(
            QuestionnaireRepository repository, QuestionnaireResponseFormatter formatter,
            QuestionnaireResponse response, QuestionnaireDefinition definition, List<JObject> visibleQuestions,
            Dictionary<string, Guid> questionIdMap, bool simulationMode, Dictionary<QuestionnaireRepository.QuestionKey, Entity> existingByNameAndNumber)
        {
            formatter.Logger.Trace($"Starting second pass - processing relationships and merging hidden questions");
            var detailResponsesToAppend = new Dictionary<Guid, List<JObject>>();

            foreach (var questionDefinition in visibleQuestions.Where(q => !string.IsNullOrEmpty(q["visibleIf"]?.ToString())))
            {
                var questionName = questionDefinition["name"]?.ToString();
                if (string.IsNullOrEmpty(questionName)) continue;

                var visibleIf = questionDefinition["visibleIf"].ToString();
                var parentQuestionName = QuestionnaireDefinition.ParseParentQuestionName(visibleIf);
                if (string.IsNullOrEmpty(parentQuestionName)) continue;

                // Find parent ID
                if (!questionIdMap.TryGetValue(parentQuestionName, out Guid parentId))
                {
                    var parentDef = definition.FindQuestionDefinition(parentQuestionName);
                    if (parentDef?["hideNumber"]?.ToObject<bool>() == true)
                    {
                        var key = new QuestionnaireRepository.QuestionKey(parentQuestionName, null);
                        if (existingByNameAndNumber.TryGetValue(key, out Entity parentEntity))
                        {
                            parentId = parentEntity.Id;
                        }
                    }
                }

                if (parentId == Guid.Empty)
                {
                    formatter.Logger.Trace($"Could not find parent record for '{parentQuestionName}' to link dependent question '{questionName}'. Skipping relationship.");
                    continue;
                }

                // Case 1: Hidden question to be merged
                if (questionDefinition["hideNumber"]?.ToObject<bool>() == true)
                {
                    var responseValue = response.GetValue(questionName);
                    if (responseValue != null && responseValue.Type != JTokenType.Null)
                    {
                        string questionType = questionDefinition["type"]?.ToString();
                        string responseText = formatter.FormatResponse(responseValue, questionType, questionDefinition);
                        var title = formatter.RemoveHtmlTags(QuestionnaireDefinition.GetTextFieldValue(questionDefinition["title"], "default"));

                        if (!string.IsNullOrEmpty(responseText))
                        {
                            var detailObject = new JObject
                            {
                                ["question"] = title.TrimEnd(' ', ':'),
                                ["answer"] = responseText
                            };

                            if (!detailResponsesToAppend.ContainsKey(parentId))
                            {
                                detailResponsesToAppend[parentId] = new List<JObject>();
                            }
                            detailResponsesToAppend[parentId].Add(detailObject);
                            formatter.Logger.Trace($"Queued detail JSON for parent record ID {parentId}");
                        }
                    }
                }
                // Case 2: Finding that needs to be linked
                else
                {
                    if (questionIdMap.TryGetValue(questionName, out Guid childId))
                    {
                        if (simulationMode)
                        {
                            formatter.Logger.Trace($"[S] Linking finding {questionName} to parent question {parentQuestionName}");
                        }
                        else
                        {
                            formatter.Logger.Trace($"Linking finding {questionName} to parent question {parentQuestionName}");
                            repository.LinkFindingToParent(childId, parentId);
                        }
                    }
                    else
                    {
                        formatter.Logger.Trace($"Could not find child finding record '{questionName}' to link to parent '{parentQuestionName}'.");
                    }
                }
            }

            return detailResponsesToAppend;
        }

        /// <summary>
        /// Executes the merge pass to update parent records with hidden question details.
        /// </summary>
        private static void ExecuteMergePass(QuestionnaireRepository repository, Dictionary<Guid, List<JObject>> detailResponsesToAppend, bool simulationMode)
        {
            if (!detailResponsesToAppend.Any()) return;

            foreach (var kvp in detailResponsesToAppend)
            {
                Guid parentId = kvp.Key;
                List<JObject> details = kvp.Value;

                if (simulationMode)
                {
                    string simJson = new JArray(details).ToString(Newtonsoft.Json.Formatting.None);
                    // Note: We don't have access to Logger here, so we'll use the repository's tracer
                    continue;
                }

                try
                {
                    string newResponseJson = new JArray(details).ToString(Newtonsoft.Json.Formatting.None);
                    repository.UpdateResponseWithMergedDetails(parentId, newResponseJson);
                }
                catch (Exception)
                {
                    // Note: We don't have access to Logger here, so we'll use the repository's tracer
                    // The repository will handle its own error logging
                }
            }
        }

        /// <summary>
        /// Executes the final pass to deactivate orphaned records.
        /// </summary>
        private static void ExecuteFinalPass_Deactivate(QuestionnaireRepository repository, Dictionary<int, Entity> existingByNumber, HashSet<Guid> processedIds)
        {
            foreach (var kvp in existingByNumber)
            {
                Guid existingId = kvp.Value.Id;
                
                if (!processedIds.Contains(existingId))
                {
                    repository.DeactivateQuestionResponse(existingId);
                }
            }
        }

        /// <summary>
        /// Creates a question response record using the repository.
        /// </summary>
        private static Guid CreateQuestionResponseRecord(QuestionnaireRepository repository, QuestionnaireResponseFormatter formatter,
            string questionName, string questionTextEn, string questionTextFr, string responseText,
            string provisionReference, string provisionTextEn, string provisionTextFr, string details,
            string wostName, Guid workOrderServiceTaskId, EntityReference questionnaireRef,
            int? questionNumber, JToken findingObject, bool isUpdate, Guid existingId, int existingVersion)
        {
            var questionResponse = new Entity("ts_questionresponse");
            
            if (isUpdate)
            {
                questionResponse.Id = existingId;
                int newVersion = existingVersion + 1;
                questionResponse["ts_version"] = newVersion;
                questionResponse["statecode"] = new OptionSetValue(0); // 0 = Active
            }
            else
            {
                if (questionNumber.HasValue)
                {
                    var name = $"{wostName} [{questionNumber.Value}]";
                    questionResponse["ts_name"] = name.Length > 100 ? name.Substring(0, 100) : name;
                    questionResponse["ts_questionnumber"] = questionNumber.Value;
                }
                else
                {
                    var name = $"{wostName} [{questionName}]";
                    questionResponse["ts_name"] = name.Length > 100 ? name.Substring(0, 100) : name;
                }
                
                questionResponse["ts_msdyn_workorderservicetask"] = new EntityReference("msdyn_workorderservicetask", workOrderServiceTaskId);
                questionResponse["ts_questionnaire"] = questionnaireRef;
                questionResponse["ts_version"] = 1;
                questionResponse["statecode"] = new OptionSetValue(0); // 0 = Active
            }
            
            // Common properties
            questionResponse["ts_questionname"] = questionName?.Length > 4000 ? questionName.Substring(0, 4000) : questionName;
            questionResponse["ts_questiontextenglish"] = questionTextEn?.Length > 4000 ? questionTextEn.Substring(0, 4000) : questionTextEn;
            questionResponse["ts_questiontextfrench"] = questionTextFr?.Length > 4000 ? questionTextFr.Substring(0, 4000) : questionTextFr;
            questionResponse["ts_provisionreference"] = provisionReference?.Length > 4000 ? provisionReference.Substring(0, 4000) : provisionReference;
            questionResponse["ts_provisiontextenglish"] = formatter.RemoveHtmlTags(provisionTextEn)?.Length > 4000 ? formatter.RemoveHtmlTags(provisionTextEn).Substring(0, 4000) : formatter.RemoveHtmlTags(provisionTextEn);
            questionResponse["ts_provisiontextfrench"] = formatter.RemoveHtmlTags(provisionTextFr)?.Length > 4000 ? formatter.RemoveHtmlTags(provisionTextFr).Substring(0, 4000) : formatter.RemoveHtmlTags(provisionTextFr);
            questionResponse["ts_details"] = details?.Length > 4000 ? details.Substring(0, 4000) : details;

            if (!string.IsNullOrEmpty(responseText) && !questionName.StartsWith("finding-"))
            {
                questionResponse["ts_response"] = responseText.Length > 4000 ? responseText.Substring(0, 4000) : responseText;
            }

            // Handle findings
            if (questionName.StartsWith("finding-"))
            {
                try
                {
                    if (findingObject != null)
                    {
                        var operations = findingObject["operations"] as JArray;
                        if (operations != null && operations.Count > 0)
                        {
                            var operationIds = string.Join(",", operations.Select(op => op["operationID"]?.ToString()));
                            questionResponse["ts_operations"] = operationIds;

                            var findingType = operations[0]["findingType"]?.ToString();
                            if (!string.IsNullOrEmpty(findingType))
                            {
                                questionResponse["ts_findingtype"] = new OptionSetValue(int.Parse(findingType));
                            }
                        }
                    }
                }
                catch (Newtonsoft.Json.JsonReaderException ex)
                {
                    formatter.Logger.Error($"Error parsing operations for finding {questionName}: {ex.Message}");
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
    }
} 