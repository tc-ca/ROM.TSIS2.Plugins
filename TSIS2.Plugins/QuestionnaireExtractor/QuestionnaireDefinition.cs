using Microsoft.Xrm.Sdk;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;

namespace TSIS2.Plugins.QuestionnaireExtractor
{
    /// <summary>
    /// Represents a questionnaire definition with methods to parse and access question metadata.
    /// This class encapsulates the questionnaire structure and provides easy access to question definitions.
    /// </summary>
    public class QuestionnaireDefinition
    {
        private readonly JObject _definition;
        private readonly ILoggingService _logger;

        public JObject Definition => _definition;

        public QuestionnaireDefinition(string definitionJson, ILoggingService logger = null)
        {
            if (string.IsNullOrEmpty(definitionJson))
            {
                throw new ArgumentNullException(nameof(definitionJson), "Questionnaire definition JSON cannot be null or empty.");
            }
            _definition = JObject.Parse(definitionJson);
            _logger = logger ?? new LoggerAdapter();
        }

        /// <summary>
        /// Finds a specific question's definition JToken within the questionnaire structure.
        /// It handles standard questions and matrix-based questions.
        /// </summary>
        /// <param name="questionName">The name of the question to find.</param>
        /// <returns>The JToken for the question definition, or null if not found.</returns>
        public JToken FindQuestionDefinition(string questionName)
        {
            // For matrix questions, the questionName might contain a dot (e.g., "AOSP – TC Review.Result")
            // We only need the part before the dot to find the question definition
            if (questionName.Contains("."))
            {
                var parts = questionName.Split('.');
                string parentName = parts[0];
                var parentQuestion = FindElementByName(_definition, parentName);
                if (parentQuestion != null && parentQuestion["type"]?.ToString() == "matrix")
                {
                    return parentQuestion;
                }
            }
            return FindElementByName(_definition, questionName);
        }

        private JToken FindElementByName(JToken token, string questionName)
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

        public List<string> GetAllQuestionNames()
        {
            var questionNames = new List<string>();
            
            void CollectNames(JToken token)
            {
                if (token == null) return;

                if (token.Type == JTokenType.Object)
                {
                    var obj = (JObject)token;
                    string name = obj["name"]?.ToString();
                    bool isContainer = obj["elements"] != null;

                    if (!string.IsNullOrEmpty(name) && !isContainer)
                    {
                        questionNames.Add(name);
                    }
                    
                    if (obj["pages"] != null)
                    {
                        foreach (var child in obj["pages"])
                            CollectNames(child);
                    }
                    if (isContainer)
                    {
                        foreach (var child in obj["elements"])
                            CollectNames(child);
                    }
                }
                else if (token.Type == JTokenType.Array)
                {
                    foreach (var item in token)
                        CollectNames(item);
                }
            }

            CollectNames(_definition["pages"]);
            return questionNames;
        }

        /// <summary>
        /// A utility method to get a localized string from a JToken that might be
        /// a simple string or a JSON object with locale keys.
        /// </summary>
        public static string GetTextFieldValue(JToken token, string localeKey = "default")
        {
            if (token == null) return null;

            if (token.Type == JTokenType.Object)
            {
                var obj = (JObject)token;
                return obj[localeKey]?.ToString();
            }
            return token.ToString();
        }

        public static string ParseParentQuestionName(string visibleIf)
        {
            var match = System.Text.RegularExpressions.Regex.Match(visibleIf, @"\{([^}.]+)");
            var result = match.Success ? match.Groups[1].Value.Trim() : string.Empty;
            return result;
        }

        /// <summary>
        /// Traverses the definition and collects all question elements that are visible
        /// based on the answers provided in the response JSON.
        /// </summary>
        public List<JObject> CollectVisibleQuestions(JObject response)
        {
            var visibleQuestions = new List<JObject>();
            CollectVisibleQuestionsRecursive(_definition["pages"], response, visibleQuestions);
            return visibleQuestions;
        }

        private void CollectVisibleQuestionsRecursive(JToken node, JObject response, List<JObject> visibleQuestions)
        {
            if (node == null) return;

            if (node.Type == JTokenType.Object)
            {
                var obj = (JObject)node;
                var visibleIf = obj["visibleIf"]?.ToString();
                if (!string.IsNullOrEmpty(visibleIf) && !IsConditionMet(response, visibleIf))
                {
                    return;
                }

                string name = obj["name"]?.ToString();
                bool isContainer = obj["elements"] != null;
                bool isPageContainer = obj["pages"] != null;

                if (!isContainer && !isPageContainer && !string.IsNullOrEmpty(name))
                {
                    visibleQuestions.Add(obj);
                }

                if (isPageContainer)
                {
                    CollectVisibleQuestionsRecursive(obj["pages"], response, visibleQuestions);
                }
                if (isContainer)
                {
                    CollectVisibleQuestionsRecursive(obj["elements"], response, visibleQuestions);
                }
            }
            else if (node.Type == JTokenType.Array)
            {
                foreach (var item in node)
                {
                    CollectVisibleQuestionsRecursive(item, response, visibleQuestions);
                }
            }
        }

        private bool IsConditionMet(JObject response, string visibleIf)
        {
            if (string.IsNullOrEmpty(visibleIf))
                return false;
            
            var orConditions = visibleIf.Split(new[] { " or " }, StringSplitOptions.None);
            foreach (var condition in orConditions)
            {
                if (EvaluateSingleCondition(response, condition.Trim()))
                {
                    return true;
                }
            }
            return false;
        }

        private bool EvaluateSingleCondition(JObject response, string condition)
        {
            if (string.IsNullOrEmpty(condition)) return false;

            var match = System.Text.RegularExpressions.Regex.Match(condition, @"\{([^}]+)\}\s*(?<op>anyof|contains|strarequals|equals|=)\s*(?<val>\[[^\]]*\]|'[^']*'|true|false)", System.Text.RegularExpressions.RegexOptions.IgnoreCase);

            if (!match.Success)
            {
                _logger.Debug($"Could not parse single visibleIf sub-condition: '{condition}'");
                return false;
            }

            string questionName = match.Groups[1].Value.Trim();
            string op = match.Groups["op"].Value.ToLower().Trim();
            string valStr = match.Groups["val"].Value;

            if (!response.TryGetValue(questionName, out JToken responseValue) || responseValue.Type == JTokenType.Null)
            {
                return false;
            }

            switch (op)
            {
                case "anyof":
                    if (responseValue.Type != JTokenType.Array)
                    {
                        _logger.Debug($"'anyof' condition failed for '{questionName}': response was not an array.");
                        return false;
                    }
                    var expectedValues = new HashSet<string>();
                    var valueMatches = System.Text.RegularExpressions.Regex.Matches(valStr, @"'([^']*)'");
                    foreach (System.Text.RegularExpressions.Match valueMatch in valueMatches)
                    {
                        expectedValues.Add(valueMatch.Groups[1].Value);
                    }
                    if (!expectedValues.Any())
                    {
                        _logger.Debug($"'anyof' condition for '{questionName}' has no values to check against in '{valStr}'.");
                        return false;
                    }
                    foreach (var value in responseValue)
                    {
                        if (expectedValues.Contains(value.ToString()))
                        {
                            return true;
                        }
                    }
                    return false;

                case "contains":
                    {
                        string expectedValue = valStr.Trim('\'');
                        if (responseValue.Type == JTokenType.Array)
                        {
                            return responseValue.Values<string>().Contains(expectedValue);
                        }
                        return responseValue.ToString().Contains(expectedValue);
                    }

                case "=":
                case "equals":
                case "strarequals":
                    {
                        if (responseValue.Type == JTokenType.Array && valStr.StartsWith("[") && valStr.EndsWith("]"))
                        {
                            var valueMatch = System.Text.RegularExpressions.Regex.Match(valStr, @"\['([^']*)'\]");
                            if (valueMatch.Success)
                            {
                                string expectedValueInArray = valueMatch.Groups[1].Value;
                                return responseValue.Values<string>().Contains(expectedValueInArray);
                            }
                        }
                        string expectedValue = valStr.Trim('\'');
                        return string.Equals(responseValue.ToString(), expectedValue, StringComparison.OrdinalIgnoreCase);
                    }

                default:
                    _logger.Warning($"Unsupported operator '{op}' in visibleIf condition: '{condition}'");
                    return false;
            }
        }
    }
}