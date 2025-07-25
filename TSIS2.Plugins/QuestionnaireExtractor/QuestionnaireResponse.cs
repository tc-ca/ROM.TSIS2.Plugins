using Microsoft.Xrm.Sdk;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;

namespace TSIS2.Plugins.QuestionnaireExtractor
{
    /// <summary>
    /// Represents a questionnaire response with methods to access and validate response data.
    /// </summary>
    public class QuestionnaireResponse
    {
        private readonly JObject _response;
        private readonly ILoggingService _logger;

        /// <summary>
        /// Gets the raw response data.
        /// </summary>
        public JObject Response => _response;

        /// <summary>
        /// Initializes a new instance of the QuestionnaireResponse class.
        /// </summary>
        /// <param name="responseJson">The JSON response data.</param>
        /// <param name="logger">The logging service for logging.</param>
        public QuestionnaireResponse(string responseJson, ILoggingService logger = null)
        {
            if (string.IsNullOrEmpty(responseJson))
            {
                _response = new JObject();
            }
            else
            {
                _response = JObject.Parse(responseJson);
            }
            _logger = logger ?? new LoggerAdapter();
        }

        /// <summary>
        /// Gets the response value for a specific question.
        /// </summary>
        /// <param name="questionName">The name of the question to retrieve.</param>
        /// <returns>The JToken containing the response value, or null if not found.</returns>
        public JToken GetValue(string questionName)
        {
            return _response.TryGetValue(questionName.Trim(), out JToken value) ? value : null;
        }

        /// <summary>
        /// Checks if a question has a response value.
        /// </summary>
        /// <param name="questionName">The name of the question to check.</param>
        /// <returns>True if the question has a response, false otherwise.</returns>
        public bool HasValue(string questionName)
        {
            return _response.ContainsKey(questionName.Trim());
        }

        /// <summary>
        /// Gets the response value for a specific question as a string.
        /// </summary>
        /// <param name="questionName">The name of the question to retrieve.</param>
        /// <returns>The response as a string, or null if not found.</returns>
        public string GetStringValue(string questionName)
        {
            var value = GetValue(questionName);
            return value?.ToString();
        }

        /// <summary>
        /// Gets the response value for a specific question as a boolean.
        /// </summary>
        /// <param name="questionName">The name of the question to retrieve.</param>
        /// <returns>The response as a boolean, or null if not found or not a boolean.</returns>
        public bool? GetBooleanValue(string questionName)
        {
            var value = GetValue(questionName);
            if (value?.Type == JTokenType.Boolean)
            {
                return value.ToObject<bool>();
            }
            return null;
        }

        /// <summary>
        /// Gets the response value for a specific question as an array of strings.
        /// </summary>
        /// <param name="questionName">The name of the question to retrieve.</param>
        /// <returns>The response as an array of strings, or null if not found or not an array.</returns>
        public string[] GetArrayValue(string questionName)
        {
            var value = GetValue(questionName);
            if (value?.Type == JTokenType.Array)
            {
                return value.Select(v => v.ToString()).ToArray();
            }
            return null;
        }

        /// <summary>
        /// Gets the response value for a specific question as a JObject.
        /// </summary>
        /// <param name="questionName">The name of the question to retrieve.</param>
        /// <returns>The response as a JObject, or null if not found or not an object.</returns>
        public JObject GetObjectValue(string questionName)
        {
            var value = GetValue(questionName);
            return value?.Type == JTokenType.Object ? (JObject)value : null;
        }

        /// <summary>
        /// Gets the detail value for a specific question (e.g., "questionName-Detail").
        /// </summary>
        /// <param name="questionName">The name of the question to retrieve details for.</param>
        /// <returns>The detail value as a string, or null if not found.</returns>
        public string GetDetailValue(string questionName)
        {
            var detailKey = $"{questionName}-Detail";
            return GetStringValue(detailKey);
        }

        /// <summary>
        /// Checks if a question has a detail value.
        /// </summary>
        /// <param name="questionName">The name of the question to check.</param>
        /// <returns>True if the question has a detail value, false otherwise.</returns>
        public bool HasDetailValue(string questionName)
        {
            var detailKey = $"{questionName}-Detail";
            return HasValue(detailKey);
        }

        /// <summary>
        /// Gets all question names that have responses.
        /// </summary>
        /// <returns>A list of question names that have response values.</returns>
        public List<string> GetAllQuestionNames()
        {
            return _response.Properties()
                .Select(p => p.Name)
                .Where(name => !name.EndsWith("-Detail")) // Exclude detail fields
                .ToList();
        }

        /// <summary>
        /// Gets all question names including detail fields.
        /// </summary>
        /// <returns>A list of all question names including detail fields.</returns>
        public List<string> GetAllFieldNames()
        {
            return _response.Properties()
                .Select(p => p.Name)
                .ToList();
        }

        /// <summary>
        /// Checks if the response is empty (no questions answered).
        /// </summary>
        /// <returns>True if no questions have responses, false otherwise.</returns>
        public bool IsEmpty()
        {
            return !_response.Properties().Any();
        }

        /// <summary>
        /// Gets the count of questions that have responses.
        /// </summary>
        /// <returns>The number of questions with responses.</returns>
        public int GetResponseCount()
        {
            return _response.Properties()
                .Count(p => !p.Name.EndsWith("-Detail")); // Exclude detail fields
        }

        /// <summary>
        /// Determines the reason why no responses were generated for a set of questions.
        /// </summary>
        /// <param name="questionNames">The list of question names to check.</param>
        /// <returns>A descriptive reason for why no responses were generated.</returns>
        public string DetermineNoResponseReason(List<string> questionNames)
        {
            if (questionNames.Count == 0)
                return "No questions found in definition";
            
            int questionsWithData = 0;
            foreach (var questionName in questionNames)
            {
                if (HasValue(questionName) || HasDetailValue(questionName))
                    questionsWithData++;
            }
            
            return questionsWithData == 0 
                ? "All questions have empty responses" 
                : "Questions had data but processing failed";
        }

        /// <summary>
        /// Finds the comment associated with a multipletext question by looking for the next comment question in the definition.
        /// </summary>
        /// <param name="definition">The questionnaire definition.</param>
        /// <param name="multipletextQuestionName">The name of the multipletext question.</param>
        /// <returns>The comment text if found, null otherwise.</returns>
        public string FindMultipletextComment(QuestionnaireDefinition definition, string multipletextQuestionName)
        {
            try
            {
                // Find all elements in the definition
                var allElements = new List<JToken>();
                var pages = definition.Definition["pages"] as JArray;
                if (pages != null)
                {
                    foreach (var page in pages)
                    {
                        var elements = page["elements"] as JArray;
                        if (elements != null)
                        {
                            allElements.AddRange(elements);
                        }
                    }
                }

                // Find the index of the multipletext question
                int multipletextIndex = -1;
                for (int i = 0; i < allElements.Count; i++)
                {
                    if (allElements[i]["name"]?.ToString() == multipletextQuestionName)
                    {
                        multipletextIndex = i;
                        break;
                    }
                }

                if (multipletextIndex == -1)
                {
                    _logger.Debug($"Multipletext question {multipletextQuestionName} not found in definition");
                    return null;
                }

                // Look for the next comment question after the multipletext question
                for (int i = multipletextIndex + 1; i < allElements.Count; i++)
                {
                    var element = allElements[i];
                    string elementType = element["type"]?.ToString();
                    string elementName = element["name"]?.ToString();

                    if (string.Equals(elementType, "comment", StringComparison.OrdinalIgnoreCase) && 
                        !string.IsNullOrEmpty(elementName))
                    {
                        // Found a comment question, check if it has a response value
                        if (HasValue(elementName))
                        {
                            string commentText = GetStringValue(elementName);
                            if (!string.IsNullOrEmpty(commentText))
                            {
                                _logger.Debug($"Found comment for multipletext {multipletextQuestionName}: {elementName} = {commentText}");
                                return commentText;
                            }
                        }
                        break; // Stop at the first comment question found, even if it's empty
                    }
                    
                    // Stop looking if we hit another question type that's not a comment
                    if (!string.IsNullOrEmpty(elementType) && 
                        !string.Equals(elementType, "comment", StringComparison.OrdinalIgnoreCase))
                    {
                        break;
                    }
                }

                return null;
            }
            catch (Exception ex)
            {
                _logger.Debug($"Error finding multipletext comment for {multipletextQuestionName}: {ex.Message}");
                return null;
            }
        }
    }
} 