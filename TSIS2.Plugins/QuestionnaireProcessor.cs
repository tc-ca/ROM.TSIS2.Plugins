using Microsoft.Xrm.Sdk;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json;

namespace TSIS2.Plugins.Services
{
    public static class QuestionnaireProcessor
    {
        public static List<Guid> ProcessQuestionnaire(IOrganizationService service, Guid workOrderServiceTaskId, ITracingService tracingService)
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
                questionResponseIds = ProcessResponses(service, response, definition, wostName, tracingService);

                tracingService.Trace("Completed questionnaire processing. Created {0} responses", questionResponseIds.Count);
                return questionResponseIds;
            }
            catch (Exception ex)
            {
                tracingService.Trace("Error processing questionnaire: {0}", ex.Message);
                throw new InvalidPluginExecutionException($"Error processing questionnaire: {ex.Message}");
            }
        }

        private static List<Guid> ProcessResponses(IOrganizationService service, JObject response, JObject definition, string wostName, ITracingService tracingService)
        {
            var questionResponseIds = new List<Guid>();
            var questionNames = GetAllQuestionNames(definition);
            int questionNumber = 1;
            var pageQuestionMap = BuildPageQuestionMap(definition);

            tracingService.Trace($"Full Response JSON: {response.ToString()}");
            tracingService.Trace($"Full Definition JSON: {definition.ToString()}");
            tracingService.Trace("Found {0} questions to process", questionNames.Count);

            // First pass - create all primary question responses
            var questionIdMap = new Dictionary<string, Guid>();

            tracingService.Trace("Starting first pass - creating primary question responses");
            foreach (var questionName in questionNames)
            {
                tracingService.Trace($"\n=== Processing Question: {questionName} ===");
                var responseValue = GetResponseValue(response, questionName);
                tracingService.Trace($"Response Value: {responseValue?.ToString() ?? "null"}");

                var questionDefinition = FindQuestionDefinition(definition, questionName);
                tracingService.Trace($"Question Definition: {questionDefinition?.ToString() ?? "null"}");

                if (questionDefinition != null)
                {
                    tracingService.Trace("Processing question: {0}", questionName);
                    string questionType = questionDefinition["type"]?.ToString();
                    tracingService.Trace($"Question Type: {questionType}");

                    // Get page number from our map
                    int pageNumber = 1;
                    int questionInPageNumber = questionNumber;
                    if (pageQuestionMap.TryGetValue(questionName, out (int page, int question) value))
                    {
                        pageNumber = value.page;
                        questionInPageNumber = value.question;
                    }

                    // Create the question response
                    var questionResponseId = CreateQuestionResponseRecord(
                        service,
                        questionName,
                        RemoveHtmlTags(questionDefinition["title"]?["default"]?.ToString()),
                        RemoveHtmlTags(questionDefinition["title"]?["fr"]?.ToString()),
                        GetResponseText(responseValue, questionType, questionDefinition, tracingService),
                        questionDefinition["provision"]?.ToString(),
                        RemoveHtmlTags(questionDefinition["description"]?["default"]?.ToString()),
                        RemoveHtmlTags(questionDefinition["description"]?["fr"]?.ToString()),
                        responseValue?["comments"]?.ToString(),
                        wostName,
                        pageNumber,
                        questionInPageNumber
                    );

                    questionIdMap[questionName] = questionResponseId;
                    questionResponseIds.Add(questionResponseId);
                    questionNumber++;
                }
            }

            // Second pass - link findings to their parent questions
            tracingService.Trace("Starting second pass - processing findings and parent relationships");
            foreach (var questionName in questionNames)
            {
                var questionDefinition = FindQuestionDefinition(definition, questionName);
                var visibleIf = questionDefinition?["visibleIf"]?.ToString();

                if (!string.IsNullOrEmpty(visibleIf))
                {
                    tracingService.Trace("Processing finding relationship for: {0}", questionName);
                    // Parse visibleIf to get parent question name
                    var parentQuestionName = ParseParentQuestionName(visibleIf);

                    if (questionIdMap.ContainsKey(parentQuestionName) && questionIdMap.ContainsKey(questionName))
                    {
                        tracingService.Trace("Linking finding {0} to parent question {1}", questionName, parentQuestionName);
                        // Update the finding's record to link to parent
                        var updateEntity = new Entity("ts_questionresponse")
                        {
                            Id = questionIdMap[questionName],
                            ["ts_QuestionResponse"] = new EntityReference("ts_questionresponse", questionIdMap[parentQuestionName])
                        };
                        service.Update(updateEntity);
                    }
                }
            }

            tracingService.Trace("Completed processing all questions and relationships");
            return questionResponseIds;
        }

        private static Dictionary<string, (int pageNumber, int questionNumber)> BuildPageQuestionMap(JObject definition)
        {
            var map = new Dictionary<string, (int pageNumber, int questionNumber)>();
            var pages = definition["pages"] as JArray;

            if (pages != null)
            {
                int pageNumber = 1;
                int cumulativeQuestionNumber = 1;  // This will keep counting across pages

                foreach (var page in pages)
                {
                    var elements = page["elements"] as JArray;

                    if (elements != null)
                    {
                        foreach (var element in elements)
                        {
                            string questionName = element["name"]?.ToString();
                            if (!string.IsNullOrEmpty(questionName))
                            {
                                map[questionName] = (pageNumber, cumulativeQuestionNumber);
                                cumulativeQuestionNumber++;
                            }
                        }
                    }
                    pageNumber++;
                }
            }

            return map;
        }

        private static Guid CreateQuestionResponseRecord(
            IOrganizationService service,
            string questionName,
            string questionTextEn,
            string questionTextFr,
            string responseText,
            string provisionReference,
            string provisionTextEn,
            string provisionTextFr,
            string comments,
            string wostName,
            int pageNumber,
            int questionNumber)
        {
            var questionResponse = new Entity("ts_questionresponse")
            {
                ["ts_questionname"] = questionName,
                ["ts_questiontextenglish"] = questionTextEn,
                ["ts_questiontextfrench"] = questionTextFr,
                ["ts_response"] = responseText,
                ["ts_provisionreference"] = provisionReference,
                ["ts_provisiontextenglish"] = provisionTextEn,
                ["ts_provisiontextfrench"] = provisionTextFr,
                ["ts_comments"] = comments,
                ["ts_name"] = $"{wostName} [{questionNumber}]",
                ["ts_pagenumber"] = pageNumber,
                ["ts_questionnumber"] = questionNumber
            };

            if (questionName.StartsWith("finding-"))
            {
                try
                {
                    var operations = JArray.Parse(responseText);
                    if (operations.Count > 0)
                    {
                        var firstOperation = operations[0];
                        questionResponse["ts_operation"] = firstOperation["operationID"]?.ToString();
                        questionResponse["ts_findingtype"] = firstOperation["findingType"]?.ToString();
                    }
                }
                catch (JsonReaderException)
                {
                    // Handle JSON parsing error if needed
                }
            }

            return service.Create(questionResponse);
        }

        // Helper methods for JSON processing
        private static JToken GetResponseValue(JObject response, string questionName)
        {
            var parts = questionName.Split('.');
            JToken currentToken = response;
            foreach (var part in parts)
            {
                if (currentToken == null)
                    return null;

                if (currentToken.Type == JTokenType.Object)
                {
                    JObject obj = (JObject)currentToken;
                    if (obj.TryGetValue(part, out JToken nextToken))
                    {
                        currentToken = nextToken;
                    }
                    else
                    {
                        return null;
                    }
                }
                else
                {
                    return null;
                }
            }
            return currentToken;
        }

        private static List<string> GetAllQuestionNames(JObject definition)
        {
            List<string> questionNames = new List<string>();
            if (definition["pages"] != null)
            {
                CollectQuestionNames(definition["pages"], questionNames);
            }
            return questionNames;
        }

        private static void CollectQuestionNames(JToken token, List<string> questionNames)
        {
            if (token == null)
                return;

            if (token.Type == JTokenType.Object)
            {
                JObject obj = (JObject)token;
                if (obj["name"] != null)
                {
                    questionNames.Add(obj["name"].ToString());
                }
                if (obj["elements"] != null)
                {
                    CollectQuestionNames(obj["elements"], questionNames);
                }
                else if (obj["pages"] != null)
                {
                    CollectQuestionNames(obj["pages"], questionNames);
                }
            }
            else if (token.Type == JTokenType.Array)
            {
                foreach (var item in token)
                {
                    CollectQuestionNames(item, questionNames);
                }
            }
        }

        private static JToken FindQuestionDefinition(JObject definition, string questionName)
        {
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

        private static JToken FindElementByName(JToken token, string questionName)
        {
            if (token == null)
                return null;

            if (token.Type == JTokenType.Object)
            {
                JObject obj = (JObject)token;
                if (obj["name"]?.ToString() == questionName)
                {
                    return obj;
                }
                if (obj["elements"] != null)
                {
                    var found = FindElementByName(obj["elements"], questionName);
                    if (found != null) return found;
                }
                else if (obj["pages"] != null)
                {
                    var found = FindElementByName(obj["pages"], questionName);
                    if (found != null) return found;
                }
            }
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

        private static string RemoveHtmlTags(string input)
        {
            if (string.IsNullOrEmpty(input))
                return input;

            // Remove HTML tags including self-closing tags and line breaks
            var withoutTags = System.Text.RegularExpressions.Regex.Replace(input, "<[^>]*>", string.Empty)
                .Replace("\n", " ")
                .Replace("\r", " ");

            // Replace multiple spaces with a single space
            withoutTags = System.Text.RegularExpressions.Regex.Replace(withoutTags, @"\s+", " ");

            // Replace common HTML entities
            return withoutTags
                .Replace("&nbsp;", " ")
                .Replace("&amp;", "&")
                .Replace("&lt;", "<")
                .Replace("&gt;", ">")
                .Replace("&quot;", "\"")
                .Replace("&#39;", "'")
                .Replace("&ndash;", "–")
                .Replace("&mdash;", "—")
                .Trim(); // Remove leading/trailing spaces
        }

        private static (string enText, string frText) GetCleanText(JToken token)
        {
            if (token == null)
                return (null, null);

            string enText = null;
            string frText = null;

            if (token.Type == JTokenType.String)
            {
                string text = RemoveHtmlTags(token.ToString());
                enText = text;
                frText = text;
            }
            else if (token.Type == JTokenType.Object)
            {
                enText = token["default"]?.ToString() ?? token["en"]?.ToString();
                frText = token["fr"]?.ToString();

                enText = RemoveHtmlTags(enText);
                frText = RemoveHtmlTags(frText);
            }

            return (enText, frText);
        }

        private static string ParseParentQuestionName(string visibleIf)
        {
            // Handle matrix question format: "{AOSP – TC Review.Result} = 'Not Satisfactory'"
            var match = System.Text.RegularExpressions.Regex.Match(visibleIf, @"\{([^}.]+)");
            return match.Success ? match.Groups[1].Value.Trim() : string.Empty;
        }

        private static string GetResponseText(JToken responseValue, string questionType, JToken questionDefinition, ITracingService tracingService)
        {
            try
            {
                if (responseValue == null) return string.Empty;

                // Debug tracing
                var debugInfo = $"\nProcessing Response:" +
                               $"\nValue Type: {responseValue.Type}" +
                               $"\nQuestion Type: {questionType}" +
                               $"\nActual Value: {responseValue}" +
                               $"\nDefinition: {(questionDefinition?.ToString() ?? "null")}";

                tracingService.Trace(debugInfo);

                // Add diagnostic information
                var ex = new Exception();
                ex.Data["ResponseValueType"] = responseValue.Type.ToString();
                ex.Data["QuestionType"] = questionType;
                ex.Data["ResponseValue"] = responseValue.ToString();

                // Handle primitive types first - this prevents trying to access child properties on simple values
                if (responseValue.Type == JTokenType.Boolean ||
                    responseValue.Type == JTokenType.String ||
                    responseValue.Type == JTokenType.Integer ||
                    responseValue.Type == JTokenType.Float)
                {
                    return responseValue.ToString().ToLower();
                }

                switch (questionType?.ToLower())
                {
                    case "finding":
                        return responseValue["operations"]?.ToString() ?? string.Empty;

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
                                    var choiceText = choice["text"]?["default"]?.ToString() ?? valueStr;
                                    selectedValues.Add($"{valueStr} ({RemoveHtmlTags(choiceText)})");
                                }
                                else
                                {
                                    selectedValues.Add(valueStr);
                                }
                            }
                            return string.Join(", ", selectedValues);
                        }
                        return responseValue.ToString();

                    case "matrix":
                        if (responseValue.Type == JTokenType.Object)
                        {
                            var result = new List<string>();
                            foreach (JProperty row in responseValue.Children())
                            {
                                string rowName = row.Name;
                                string selectedValue = row.Value.ToString();

                                var columns = questionDefinition["columns"] as JArray;
                                var column = columns?.FirstOrDefault(c => c["value"]?.ToString() == selectedValue);
                                var columnText = column?["text"]?["default"]?.ToString() ?? selectedValue;

                                var rows = questionDefinition["rows"] as JArray;
                                var rowDef = rows?.FirstOrDefault(r => r["value"]?.ToString() == rowName);
                                var rowText = rowDef?["text"]?["default"]?.ToString() ?? rowName;

                                result.Add($"{RemoveHtmlTags(rowText)}: {RemoveHtmlTags(columnText)}");
                            }
                            return string.Join(", ", result);
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
