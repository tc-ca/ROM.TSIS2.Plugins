using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json;

namespace TSIS2.Plugins.Services
{
    public static class QuestionnaireProcessor
    {
        public static List<Guid> ProcessQuestionnaire(IOrganizationService service, Guid workOrderServiceTaskId, ITracingService tracingService, bool isRecompletion)
        {
            var questionResponseIds = new List<Guid>();
            tracingService.Trace("Starting questionnaire processing for WOST: {0}", workOrderServiceTaskId);

            // Retrieve the WOST with questionnaire data
            var wost = service.Retrieve("msdyn_workorderservicetask",
                workOrderServiceTaskId,
                new Microsoft.Xrm.Sdk.Query.ColumnSet(
                    "msdyn_name",
                    "ovs_questionnaireresponse",
                    "ovs_questionnairedefinition"
                ));

            string responseJson = wost.GetAttributeValue<string>("ovs_questionnaireresponse");
            string definitionJson = wost.GetAttributeValue<string>("ovs_questionnairedefinition");
            string wostName = wost.GetAttributeValue<string>("msdyn_name");

            if (string.IsNullOrEmpty(responseJson) || string.IsNullOrEmpty(definitionJson))
            {
                tracingService.Trace("No questionnaire data found for WOST");
                return questionResponseIds;
            }

            try
            {
                tracingService.Trace("Parsing questionnaire JSON data");
                JObject response = JObject.Parse(responseJson);
                JObject definition = JObject.Parse(definitionJson);

                tracingService.Trace("Processing questionnaire responses");
                questionResponseIds = ProcessResponses(service, response, definition, wostName, workOrderServiceTaskId, isRecompletion, tracingService);

                tracingService.Trace("Completed questionnaire processing. Created/Updated {0} responses", questionResponseIds.Count);
                return questionResponseIds;
            }
            catch (Exception ex)
            {
                tracingService.Trace("Error processing questionnaire: {0}", ex.Message);
                throw new InvalidPluginExecutionException($"Error processing questionnaire: {ex.Message}");
            }
        }

        private static List<Guid> ProcessResponses(IOrganizationService service, JObject response, JObject definition, string wostName, Guid workOrderServiceTaskId, bool isRecompletion, ITracingService tracingService)
        {
            var questionResponseIds = new List<Guid>();
            var questionNames = GetAllQuestionNames(definition);
            int questionNumber = 1;

            tracingService.Trace($"Found {questionNames.Count} questions to process: {string.Join(", ", questionNames)} ====");

            // If this is a recompletion, fetch existing question responses
            Dictionary<int, Entity> existingResponses = new Dictionary<int, Entity>();
            
            // Create HashSet to track processed IDs
            HashSet<Guid> processedIds = new HashSet<Guid>();
            
            if (isRecompletion)
            {
                tracingService.Trace("Fetching existing question responses for WOST");
                var query = new QueryExpression("ts_questionresponse")
                {
                    ColumnSet = new ColumnSet(true),
                    Criteria = new FilterExpression
                    {
                        Conditions =
                        {
                            new ConditionExpression("ts_msdyn_workorderservicetask", ConditionOperator.Equal, workOrderServiceTaskId)
                        }
                    }
                };

                var results = service.RetrieveMultiple(query);
                foreach (var existingResponse in results.Entities)
                {
                    var questionNum = existingResponse.GetAttributeValue<int>("ts_questionnumber");
                    existingResponses[questionNum] = existingResponse;
                    tracingService.Trace($"Found existing response for question {questionNum}");
                }
            }

            // First pass - create/update all primary question responses and store the question response id in a map for second pass (relationships)
            var questionIdMap = new Dictionary<string, Guid>();

            tracingService.Trace(isRecompletion ? "Starting first pass - updating/creating question responses" : "Starting first pass - creating primary question responses");
            foreach (var questionName in questionNames)
            {
                tracingService.Trace($"\n=== Processing Question: {questionName} ===");

                // Get the response value for the question
                var responseValue = response.TryGetValue(questionName.Trim(), out JToken value) ? value : null;

                var questionDefinition = FindQuestionDefinition(definition, questionName);
                
                // Check for conditional visibility dependency
                var visibleIfDebug = questionDefinition?["visibleIf"]?.ToString();
                if (!string.IsNullOrEmpty(visibleIfDebug))
                {
                    tracingService.Trace($"DEBUG - Question '{questionName}' has visibility dependency on: '{visibleIfDebug}'");
                }

                if (questionDefinition == null)
                {
                    tracingService.Trace($"No definition found for {questionName}, skipping.");
                    continue;
                }

                // Check if question has either response value or detail value
                var detailKey = $"{questionName}-Detail";
                var hasDetail = response.TryGetValue(detailKey, out JToken detailToken);
                var hasResponse = responseValue != null;
                
                if (!hasResponse && !hasDetail)
                {
                    tracingService.Trace($"DEBUG - Skipping {questionName} - no response or details provided");
                    continue;
                }

                try
                {
                    tracingService.Trace("Processing question: {0}", questionName);
                    string questionType = questionDefinition["type"]?.ToString();
                    
                    string responseText = hasResponse 
                        ? GetResponseText(responseValue, questionType, questionDefinition, tracingService)
                        : string.Empty;

                    // Get comment or details
                    string commentOrDetails = null;
                    if (hasResponse && string.Equals(questionType, "finding", StringComparison.OrdinalIgnoreCase))
                    {
                        commentOrDetails = responseValue?["comments"]?.ToString();
                    }
                    commentOrDetails = commentOrDetails ?? (hasDetail ? detailToken.ToString() : null);

                    //Sanitize the text fields
                    var titleEn = RemoveHtmlTags(GetTextFieldValue(questionDefinition["title"], "default"));
                    var titleFr = RemoveHtmlTags(GetTextFieldValue(questionDefinition["title"], "fr"));
                    var descriptionEn = RemoveHtmlTags(GetTextFieldValue(questionDefinition["description"], "default"));
                    var descriptionFr = RemoveHtmlTags(GetTextFieldValue(questionDefinition["description"], "fr"));
                    var provisionRef = questionDefinition["provision"]?.ToString();

                    Entity existingResponse = null;
                    var isUpdate = isRecompletion && existingResponses.TryGetValue(questionNumber, out existingResponse);
                    
                    //Call function to either create or update the question response record
                    Guid questionResponseId;
                    questionResponseId = ProcessQuestionResponseRecord(
                        service: service,
                        questionName: questionName,
                        questionTextEn: titleEn,
                        questionTextFr: titleFr,
                        responseText: responseText,
                        provisionReference: provisionRef,
                        provisionTextEn: descriptionEn,
                        provisionTextFr: descriptionFr,
                        details: commentOrDetails,
                        wostName: wostName,
                        questionNumber: questionNumber,
                        findingObject: responseValue,
                        tracingService: tracingService,
                        isUpdate: isUpdate,
                        existingId: isUpdate ? existingResponse.Id : default,
                        existingVersion: isUpdate ? existingResponse.GetAttributeValue<int>("ts_version") : 0
                    );

                    // Add to list of processed IDs
                    processedIds.Add(questionResponseId);
                    
                    questionIdMap[questionName] = questionResponseId;
                    questionResponseIds.Add(questionResponseId);
                    questionNumber++;
                }
                catch (Exception ex)
                {
                    tracingService.Trace($"Error processing question {questionName}: {ex.Message}");
                    throw;
                }
            }

            // Second pass - relationship linking.  Link findings to their parent questions
            tracingService.Trace("\n\n====== Starting second pass - processing findings and parent relationships ======");
            tracingService.Trace($"questionIdMap contains {questionIdMap.Count} entries: {string.Join(", ", questionIdMap.Keys)}");
            
            foreach (var questionName in questionNames)
            {
                var questionDefinition = FindQuestionDefinition(definition, questionName);
                var visibleIf = questionDefinition?["visibleIf"]?.ToString();

                // We go through conditional visibility questions to find the parent question and link the findings to it
                if (!string.IsNullOrEmpty(visibleIf))
                {
                    tracingService.Trace($"\nDEBUG - Processing finding relationship for: {questionName}");
                    tracingService.Trace($"DEBUG - visibleIf condition: {visibleIf}");
                    
                    //Get the parent question name from the visibleIf condition
                    var parentQuestionName = ParseParentQuestionName(visibleIf);

                    tracingService.Trace($"DEBUG - Parent question: '{parentQuestionName}'");
                    tracingService.Trace($"DEBUG - Parent ID exists in map: {questionIdMap.ContainsKey(parentQuestionName)}");
                    tracingService.Trace($"DEBUG - Current ID exists in map: {questionIdMap.ContainsKey(questionName)}");

                    // If the parent question is not in the map, we skip the relationship linking
                    if (!questionIdMap.ContainsKey(parentQuestionName))
                    {
                        tracingService.Trace($"DEBUG -  Parent question '{parentQuestionName}' not found in questionIdMap. Available keys: {string.Join(", ", questionIdMap.Keys)}");
                    }

                    // If the current question is not in the map, we skip the relationship linking
                    if (!questionIdMap.ContainsKey(questionName))
                    {
                        tracingService.Trace($"DEBUG - Current question '{questionName}' not found in questionIdMap.");
                    }

                    // If both the parent and current question are in the map, we can link the finding to the parent question
                    if (questionIdMap.ContainsKey(parentQuestionName) && questionIdMap.ContainsKey(questionName))
                    {
                        tracingService.Trace($"DEBUG - Linking finding {questionName} to parent question {parentQuestionName}");
                        tracingService.Trace($"DEBUG - Parent ID: {questionIdMap[parentQuestionName]}");
                        tracingService.Trace($"DEBUG - Current ID: {questionIdMap[questionName]}");

                        var updateEntity = new Entity("ts_questionresponse")
                        {
                            Id = questionIdMap[questionName],
                            ["ts_questionresponse"] = new EntityReference("ts_questionresponse", questionIdMap[parentQuestionName])
                        };

                        try
                        {
                            service.Update(updateEntity);
                            tracingService.Trace("DEBUG - Successfully updated relationship");
                        }
                        catch (Exception ex)
                        {
                            tracingService.Trace($"ERROR - Failed to update relationship: {ex.Message}");
                            tracingService.Trace($"ERROR - Exception details: {ex.ToString()}");
                            throw;
                        }
                    }
                }
            }

            // Final pass - handle orphaned records
            if (isRecompletion)
            {
                tracingService.Trace("\n\n====== Starting final pass - handling orphaned records ======");
                foreach (var kvp in existingResponses)
                {
                    Guid existingId = kvp.Value.Id;
                    
                    if (!processedIds.Contains(existingId))
                    {
                        tracingService.Trace($"Found orphaned record: {existingId} for question number {kvp.Key}");
                        
                        // Mark orphaned record as inactive
                        Entity updateEntity = new Entity("ts_questionresponse")
                        {
                            Id = existingId,
                            ["statecode"] = new OptionSetValue(1) // 1 = Inactive
                        };
                        
                        try
                        {
                            service.Update(updateEntity);
                            tracingService.Trace($"Successfully marked orphaned record {existingId} as inactive");
                        }
                        catch (Exception ex)
                        {
                            tracingService.Trace($"ERROR - Failed to update orphaned record: {ex.Message}");
                            tracingService.Trace($"ERROR - Exception details: {ex}");
                        }
                    }
                }
            }

            tracingService.Trace("Completed processing all questions and relationships");
            return questionResponseIds;
        }

        private static Guid ProcessQuestionResponseRecord(
            IOrganizationService service,
            string questionName,
            string questionTextEn,
            string questionTextFr,
            string responseText,
            string provisionReference,
            string provisionTextEn,
            string provisionTextFr,
            string details,
            string wostName,
            int questionNumber,
            JToken findingObject,
            ITracingService tracingService,
            bool isUpdate = false,
            Guid existingId = default,
            int existingVersion = 0)  // Added parameter for existing version
        {
            // Create entity or prepare update entity
            var questionResponse = new Entity("ts_questionresponse");
            
            if (isUpdate)
            {
                tracingService.Trace($"Updating question response record: {existingId}");
                questionResponse.Id = existingId;
                
                // Handle null or 0 version - treat as version 1
                if (existingVersion <= 0)
                {
                    existingVersion = 1;
                    tracingService.Trace($"Existing version was null or 0, treating as version 1");
                }
                
                // Increment version by 1
                int newVersion = existingVersion + 1;
                questionResponse["ts_version"] = newVersion;
                tracingService.Trace($"Incrementing version from {existingVersion} to {newVersion}");
                
                // Ensure record is active
                questionResponse["statecode"] = new OptionSetValue(0); // 0 = Active
            }
            else
            {
                // These fields are only needed for new records
                questionResponse["ts_name"] = $"{wostName} [{questionNumber}]";
                questionResponse["ts_questionnumber"] = questionNumber;
                
                // Initialize version to 1 for new records
                questionResponse["ts_version"] = 1;
                tracingService.Trace($"Setting initial version to 1 for new record");
                
                // Set as active
                questionResponse["statecode"] = new OptionSetValue(0); // 0 = Active
            }
            
            // Common properties for both create and update
            questionResponse["ts_questionname"] = questionName;
            questionResponse["ts_questiontextenglish"] = questionTextEn;
            questionResponse["ts_questiontextfrench"] = questionTextFr;
            questionResponse["ts_provisionreference"] = provisionReference;
            questionResponse["ts_provisiontextenglish"] = provisionTextEn;
            questionResponse["ts_provisiontextfrench"] = provisionTextFr;
            questionResponse["ts_details"] = details;

            // Only set response if it's not empty and not a finding
            if (!string.IsNullOrEmpty(responseText) && !questionName.StartsWith("finding-"))
            {
                questionResponse["ts_response"] = responseText;
            }

            // Handle findings, we store the operations for now but that might not be needed
            if (questionName.StartsWith("finding-"))
            {
                try
                {
                    // Only process operations if findingObject is not null
                    if (findingObject != null)
                    {
                        // Get the operations array
                        var operations = findingObject["operations"] as JArray;
                        if (operations != null && operations.Count > 0)
                        {
                            // Store operation IDs as comma-separated string
                            var operationIds = string.Join(",",
                                operations.Select(op => op["operationID"]?.ToString()));
                            questionResponse["ts_operations"] = operationIds;

                            // Store the finding type (observation, non-compliance)
                            var findingType = operations[0]["findingType"]?.ToString();
                            if (!string.IsNullOrEmpty(findingType))
                            {
                                questionResponse["ts_findingtype"] = new OptionSetValue(int.Parse(findingType));
                            }
                        }
                    }
                }
                catch (JsonReaderException ex)
                {
                    tracingService.Trace($"Error parsing operations for finding {questionName}: {ex.Message}");
                }
            }

            // Perform create or update based on isUpdate flag
            if (isUpdate)
            {
                service.Update(questionResponse);
                return existingId;
            }
            else
            {
                return service.Create(questionResponse);
            }
        }

        private static List<string> GetAllQuestionNames(JObject definition)
        {
            List<string> questionNames = new List<string>();

            // Get all pages
            var pages = definition["pages"] as JArray;
            if (pages != null)
            {
                foreach (var page in pages)
                {
                    // Only look at elements array - this is where questions are
                    var elements = page["elements"] as JArray;
                    if (elements != null)
                    {
                        foreach (var element in elements)
                        {
                            // Each element in the elements array is a question/finding
                            if (element["name"] != null)
                            {
                                questionNames.Add(element["name"].ToString());
                            }
                        }
                    }
                }
            }
            return questionNames;
        }

        // Find the question definition in the definition JSON from the question name in the response JSON
        private static JToken FindQuestionDefinition(JObject definition, string questionName)
        {
            // For matrix questions, the questionName might contain a dot (e.g., "AOSP – TC Review.Result")
            // We only need the part before the dot to find the question definition
            if (questionName.Contains("."))
            {
                var parts = questionName.Split('.');
                string parentName = parts[0];
                var parentQuestion = FindElementByName(definition, parentName);
                if (parentQuestion != null && parentQuestion["type"]?.ToString() == "matrix")
                {
                    return parentQuestion;
                }
            }
            return FindElementByName(definition, questionName);
        }

        //Recursively find a question definition in the JSON using the question name
        private static JToken FindElementByName(JToken token, string questionName)
        {
            if (token == null)
                return null;
            if (token.Type == JTokenType.Object)
            {
                JObject obj = (JObject)token;
                //direct match, check if object is the question we want (name: questioName)
                if (obj["name"]?.ToString() == questionName)
                {
                    return obj;
                }
                //check elements array
                if (obj["elements"] != null)
                {
                    var found = FindElementByName(obj["elements"], questionName);
                    if (found != null) return found;
                }
                //check pages array
                else if (obj["pages"] != null)
                {
                    var found = FindElementByName(obj["pages"], questionName);
                    if (found != null) return found;
                }
            }
            //token is an array (pages/elements or rows/columns)
            else if (token.Type == JTokenType.Array)
            {
                foreach (var item in token)
                {
                    var found = FindElementByName(item, questionName);
                    if (found != null) return found;
                }
            }
            return null;
        }

        private static string GetTextFieldValue(JToken token, string localeKey = "default")
        {
            if (token == null) return null;

            // if text is an array ->  {"default": "eng text?", "fr": "fr text"} return text in localeKey
            if (token.Type == JTokenType.Object)
            {
                var obj = (JObject)token;
                return obj[localeKey]?.ToString();
            }
            // get the string if not an object
            return token.ToString();
        }

        private static string RemoveHtmlTags(string input)
        {
            if (string.IsNullOrEmpty(input))
                return input;

            // Replace <br> tags with new lines
            input = System.Text.RegularExpressions.Regex.Replace(input, "<br\\s*/?>", "\n", System.Text.RegularExpressions.RegexOptions.IgnoreCase);

            // Remove other HTML tags including self-closing tags
            var withoutTags = System.Text.RegularExpressions.Regex.Replace(input, "<[^>]*>", string.Empty);
            
            // Replace multiple consecutive spaces with a single space, but preserve newlines
            withoutTags = System.Text.RegularExpressions.Regex.Replace(withoutTags, " {2,}", " ");

            // Replace common HTML entities
            return withoutTags
                // Space and basic punctuation
                .Replace("&nbsp;", " ")
                .Replace("&rsquo;", "'")
                .Replace("&lsquo;", "'")
                .Replace("&ndash;", "–")
                .Replace("&mdash;", "—")
                .Replace("&amp;", "&")
                .Replace("&lt;", "<")
                .Replace("&gt;", ">")
                .Replace("&quot;", "\"")
                .Replace("&#39;", "'")
                // French specific characters
                .Replace("&eacute;", "é")
                .Replace("&egrave;", "è")
                .Replace("&agrave;", "à")
                .Replace("&ecirc;", "ê")
                .Replace("&ucirc;", "û")
                .Replace("&icirc;", "î")
                .Replace("&ocirc;", "ô")
                .Replace("&acirc;", "â")
                .Replace("&ccedil;", "ç")
                .Replace("&euml;", "ë")
                .Replace("&iuml;", "ï")
                .Replace("&uuml;", "ü")
                .Trim(); // Remove leading/trailing spaces
        }

        private static string ParseParentQuestionName(string visibleIf)
        {
            // Handle matrix question format: "{AOSP – TC Review.Result} = 'Not Satisfactory'"
            // Need to extract "AOSP – TC Review"
            var match = System.Text.RegularExpressions.Regex.Match(visibleIf, @"\{([^}.]+)");
            var result = match.Success ? match.Groups[1].Value.Trim() : string.Empty;
            return result;
        }

        private static string GetResponseText(JToken responseValue, string questionType, JToken questionDefinition, ITracingService tracingService)
        {
            try
            {
                if (responseValue == null)
                {
                    tracingService.Trace("Response value is null, returning empty string");
                    return string.Empty;
                }

                // Findings don't have response text
                if (questionType?.ToLower() == "finding")
                {
                    tracingService.Trace("Question is a finding, returning empty string");
                    return string.Empty;
                }

                // Handle primitive types
                if (responseValue.Type == JTokenType.Boolean ||
                    responseValue.Type == JTokenType.String ||
                    responseValue.Type == JTokenType.Integer ||
                    responseValue.Type == JTokenType.Float)
                {
                    return responseValue.ToString().ToLower();
                }

                // Handle different question types
                switch (questionType?.ToLower())
                {
                    case "finding":
                        return string.Empty;

                    case "checkbox":
                        if (responseValue.Type == JTokenType.Array)
                        {
                            var selectedValues = new List<string>();
                            var choices = questionDefinition["choices"] as JArray;
                            foreach (var value in responseValue)
                            {
                                string valueStr = value.ToString();
                                var choice = choices?.FirstOrDefault(c => c["value"]?.ToString() == valueStr);
                                if (choice != null)
                                {
                                    // Handle both formats: when 'text' is a complex object or a simple string
                                    string choiceText;
                                    var textToken = choice["text"];
                                    if (textToken.Type == JTokenType.Object)
                                    {
                                        // Handle localized text object with 'default' property
                                        choiceText = textToken["default"]?.ToString() ?? valueStr;
                                    }
                                    else
                                    {
                                        // Handle simple string text
                                        choiceText = textToken.ToString();
                                    }
                                    selectedValues.Add(RemoveHtmlTags(choiceText));
                                }
                                else
                                {
                                    selectedValues.Add(valueStr);
                                }
                            }
                            return string.Join(",", selectedValues);
                        }
                        return responseValue.ToString();

                    // Matrix response processing
                    case "matrix":
                        if (responseValue.Type == JTokenType.Object)
                        {
                            var results = new List<string>();
                            
                            foreach (JProperty row in responseValue.Children())
                            {
                                // Matrix response format: "RowQuestion.AnswerColumn"
                                var rowName = row.Name;  // Format: "QuestionText.ColumnValue"
                                var isResultRow = rowName.EndsWith(".Result", StringComparison.OrdinalIgnoreCase);

                                //Find which column (answer) was selected for this row (e.g.: "Satisfactory", "Not Satisfactory")
                                var selectedColumnValue = row.Value.ToString();
                                
                                // Look up column display text with proper handling of both text formats
                                string columnDisplayText = selectedColumnValue;
                                var columnDef = questionDefinition["columns"]?.FirstOrDefault(c => c["value"]?.ToString() == selectedColumnValue);
                                if (columnDef != null && columnDef["text"] != null)
                                {
                                    var textToken = columnDef["text"];
                                    if (textToken.Type == JTokenType.Object)
                                    {
                                        // Handle localized text object with 'default' property
                                        columnDisplayText = textToken["default"]?.ToString() ?? selectedColumnValue;
                                    }
                                    else
                                    {
                                        // Handle simple string text
                                        columnDisplayText = textToken.ToString();
                                    }
                                }

                                // Find the original question text for this row with proper handling of both formats
                                string rowQuestionText = rowName;
                                
                                // Handle two formats of rows array - objects with properties or simple strings
                                var rowsArray = questionDefinition["rows"] as JArray;
                                if (rowsArray != null)
                                {
                                    // Try to find row by value property first (object format)
                                    var rowDefObj = rowsArray.FirstOrDefault(r => r.Type == JTokenType.Object && r["value"]?.ToString() == rowName);
                                    if (rowDefObj != null && rowDefObj["text"] != null)
                                    {
                                        var textToken = rowDefObj["text"];
                                        if (textToken.Type == JTokenType.Object)
                                        {
                                            rowQuestionText = textToken["default"]?.ToString() ?? rowName;
                                        }
                                        else
                                        {
                                            rowQuestionText = textToken.ToString();
                                        }
                                    }
                                    else
                                    {
                                        // Try to find row as simple string
                                        var rowDefString = rowsArray.FirstOrDefault(r => r.Type == JTokenType.String && r.ToString() == rowName);
                                        if (rowDefString != null)
                                        {
                                            rowQuestionText = rowDefString.ToString();
                                        }
                                    }
                                }

                                // Special handling for Result row to avoid redundant responses like "Result: Not Satisfactory" (When there is only 1 row)
                                results.Add(isResultRow 
                                    ? RemoveHtmlTags(columnDisplayText)
                                    : $"{RemoveHtmlTags(rowQuestionText)}: {RemoveHtmlTags(columnDisplayText)}");
                            }
                            
                            return string.Join(", ", results);
                        }
                        return responseValue.ToString();

                    default:
                        return responseValue.ToString();
                }
            }
            catch (Exception ex)
            {
                ex.Data["DebugInfo"] = $"Failed processing response value: {responseValue?.ToString() ?? "null"}" +
                                      $" of type {responseValue?.Type.ToString() ?? "null"}" +
                                      $" for question type {questionType ?? "null"}";
                throw;
            }
        }
    }
}