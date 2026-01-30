using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Newtonsoft.Json.Linq;

namespace TSIS2.Plugins.QuestionnaireProcessor
{
    /// <summary>
    /// Represents a questionnaire response with methods to access and validate response data.
    /// </summary>
    public class QuestionnaireResponse
    {
        private readonly JObject _response;
        private readonly ILoggingService _logger;

        public JObject Response => _response;

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

        public JToken GetValue(string questionName)
        {
            if (string.IsNullOrEmpty(questionName))
                return null;

            return _response.TryGetValue(questionName.Trim(), out JToken value) ? value : null;
        }

        public bool HasValue(string questionName)
        {
            if (string.IsNullOrEmpty(questionName))
                return false;

            return _response.ContainsKey(questionName.Trim());
        }

        public string GetStringValue(string questionName)
        {
            return GetValue(questionName)?.ToString();
        }

        public bool? GetBooleanValue(string questionName)
        {
            var value = GetValue(questionName);
            return value?.Type == JTokenType.Boolean ? value.ToObject<bool>() : (bool?)null;
        }

        public string[] GetArrayValue(string questionName)
        {
            var value = GetValue(questionName);
            return value?.Type == JTokenType.Array ? value.Select(v => v.ToString()).ToArray() : null;
        }

        public JObject GetObjectValue(string questionName)
        {
            var value = GetValue(questionName);
            return value?.Type == JTokenType.Object ? (JObject)value : null;
        }

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
            if (definition == null || string.IsNullOrEmpty(multipletextQuestionName))
                return null;

            try
            {
                // Use the definition's GetAllQuestionNames which properly handles nested elements (panels, containers)
                var allQuestions = definition.GetAllQuestionNames();

                int multipletextIndex = allQuestions.IndexOf(multipletextQuestionName);
                if (multipletextIndex == -1)
                {
                    _logger.Debug($"Multipletext question {multipletextQuestionName} not found in definition");
                    return null;
                }

                // Look for the next comment question after the multipletext question
                for (int i = multipletextIndex + 1; i < allQuestions.Count; i++)
                {
                    string elementName = allQuestions[i];
                    var questionDef = definition.FindQuestionDefinition(elementName);
                    string elementType = questionDef?["type"]?.ToString();

                    if (string.Equals(elementType, "comment", StringComparison.OrdinalIgnoreCase))
                    {
                        // Found a comment question, check if it has a response value
                        if (HasValue(elementName))
                        {
                            string commentText = GetStringValue(elementName);
                            if (!string.IsNullOrEmpty(commentText))
                            {
                                _logger.Debug($"Found comment for multipletext {multipletextQuestionName}: {elementName}");
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
                _logger.Debug($"Error finding multipletext comment for {multipletextQuestionName}: {ex}");
                return null;
            }
        }
    }
}