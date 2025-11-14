using Microsoft.Xrm.Sdk;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;

namespace TSIS2.Plugins.QuestionnaireExtractor
{
    /// <summary>
    /// Handles the complex task of converting raw JSON answers into human-readable text.
    /// This class transforms response data using the questionnaire definition for context.
    /// </summary>
    public class QuestionnaireResponseFormatter
    {
        private readonly ILoggingService _logger;

        /// <summary>
        /// Initializes a new instance of the QuestionnaireResponseFormatter class.
        /// </summary>
        /// <param name="logger">The logging service for logging.</param>
        public QuestionnaireResponseFormatter(ILoggingService logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        /// <summary>
        /// Gets the logging service for logging.
        /// </summary>
        public ILoggingService Logger => _logger;

        /// <summary>
        /// Formats a response value into human-readable text based on the question type and definition.
        /// </summary>
        /// <param name="responseValue">The raw response value from the questionnaire.</param>
        /// <param name="questionType">The type of question (e.g., radiogroup, checkbox, matrix).</param>
        /// <param name="questionDefinition">The question definition containing choices, columns, etc.</param>
        /// <returns>Formatted human-readable text.</returns>
        public string FormatResponse(JToken responseValue, string questionType, JToken questionDefinition)
        {
            try
            {
                if (responseValue == null || responseValue.Type == JTokenType.Null)
                {
                    _logger.Trace("Response value is null, returning empty string");
                    return string.Empty;
                }

                // Handle different question types
                switch (questionType?.ToLower())
                {
                    case "finding":
                        return string.Empty;

                    case "radiogroup":
                    case "dropdown": // Assuming dropdown behaves similarly
                        return FormatRadiogroupResponse(responseValue, questionDefinition);

                    case "checkbox":
                        return FormatCheckboxResponse(responseValue, questionDefinition);

                    case "multipletext":
                        return FormatMultipletextResponse(responseValue);

                    case "matrix":
                        return FormatMatrixResponse(responseValue, questionDefinition);

                    default:
                        // This covers primitive types like boolean, string (for text questions), integer, float.
                        return RemoveHtmlTags(CleanupEscapedCharacters(responseValue.ToString()));
                }
            }
            catch (Exception ex)
            {
                _logger.Error($"Failed processing response value: {responseValue?.ToString() ?? "null"}" +
                          $" of type {responseValue?.Type.ToString() ?? "null"}" +
                          $" for question type {questionType ?? "null"}" +
                          $" - Error: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Formats a radiogroup or dropdown response by finding the display text for the selected choice.
        /// </summary>
        private string FormatRadiogroupResponse(JToken responseValue, JToken questionDefinition)
        {
            var selectedValue = responseValue.ToString();
            var choices = questionDefinition["choices"] as JArray;

            if (choices != null)
            {
                var choice = choices.FirstOrDefault(c =>
                {
                    if (c.Type == JTokenType.Object)
                        return c["value"]?.ToString() == selectedValue;
                    if (c.Type == JTokenType.String)
                        return c.ToString() == selectedValue;
                    return false;
                });

                if (choice != null)
                {
                    JToken textSource = (choice.Type == JTokenType.Object) ? choice["text"] : choice;

                    if (textSource != null)
                    {
                        string choiceText;
                        if (textSource.Type == JTokenType.Object)
                        {
                            choiceText = textSource["default"]?.ToString() ?? selectedValue;
                        }
                        else
                        {
                            // Handle simple string text
                            choiceText = textSource.ToString();
                        }
                        return RemoveHtmlTags(CleanupEscapedCharacters(choiceText));
                    }
                }
            }
            // Fallback to the value itself if no matching choice text is found.
            return RemoveHtmlTags(CleanupEscapedCharacters(selectedValue));
        }

        /// <summary>
        /// Formats a checkbox response by finding display text for each selected choice.
        /// </summary>
        private string FormatCheckboxResponse(JToken responseValue, JToken questionDefinition)
        {
            if (responseValue.Type == JTokenType.Array)
            {
                var selectedValues = new List<string>();
                var choices = questionDefinition["choices"] as JArray;
                
                foreach (var value in responseValue)
                {
                    string valueStr = value.ToString();
                    string choiceText = valueStr; // Default to the value itself
                    
                    if (choices != null)
                    {
                        // Handle both formats of choices
                        // Format 1: Objects with value/text properties
                        var choiceObj = choices.FirstOrDefault(c => c.Type == JTokenType.Object && c["value"]?.ToString() == valueStr);
                        if (choiceObj != null && choiceObj["text"] != null)
                        {
                            var textToken = choiceObj["text"];
                            if (textToken.Type == JTokenType.Object)
                            {
                                // Handle localized text
                                choiceText = textToken["default"]?.ToString() ?? valueStr;
                            }
                            else
                            {
                                // Handle simple string text
                                choiceText = textToken.ToString();
                            }
                        }
                        else
                        {
                            // Format 2: Simple strings
                            var choiceString = choices.FirstOrDefault(c => c.Type == JTokenType.String && c.ToString() == valueStr);
                            if (choiceString != null)
                            {
                                choiceText = choiceString.ToString();
                            }
                        }
                    }
                    
                    selectedValues.Add(RemoveHtmlTags(CleanupEscapedCharacters(choiceText)));
                }
                
                return string.Join(",", selectedValues);
            }
            return CleanupEscapedCharacters(responseValue.ToString());
        }

        /// <summary>
        /// Formats a multipletext response by combining all field names and values.
        /// </summary>
        private string FormatMultipletextResponse(JToken responseValue)
        {
            if (responseValue.Type == JTokenType.Object)
            {
                var results = new List<string>();
                
                foreach (JProperty item in responseValue.Children())
                {
                    string itemName = item.Name;
                    string itemValue = item.Value?.ToString() ?? string.Empty;
                    
                    // Clean up the item name and value
                    string cleanItemName = RemoveHtmlTags(CleanupEscapedCharacters(itemName));
                    string cleanItemValue = RemoveHtmlTags(CleanupEscapedCharacters(itemValue));
                    
                    string quotedItemName = NeedsQuoting(cleanItemName) ? $"\"{cleanItemName}\"" : cleanItemName;
                    results.Add($"{quotedItemName}: \"{cleanItemValue}\"");
                }
                
                return string.Join("; ", results);
            }
            return CleanupEscapedCharacters(responseValue.ToString());
        }

        /// <summary>
        /// Formats a matrix response by combining row questions with their selected column answers.
        /// </summary>
        private string FormatMatrixResponse(JToken responseValue, JToken questionDefinition)
        {
            if (responseValue.Type == JTokenType.Object)
            {
                var results = new List<string>();
                
                foreach (JProperty row in responseValue.Children())
                {
                    // Matrix response format: "RowQuestion.AnswerColumn"
                    var rowName = row.Name;  // Format: "QuestionText.ColumnValue"
                    var isResultRow = rowName.EndsWith(".Result", StringComparison.OrdinalIgnoreCase) ||
                                     string.Equals(rowName, "Result", StringComparison.OrdinalIgnoreCase);

                    // Find which column (answer) was selected for this row (e.g.: "Satisfactory", "Not Satisfactory")
                    var selectedColumnValue = row.Value.ToString();
                    
                    // Look up column display text with proper handling of both text formats
                    string columnDisplayText = selectedColumnValue;
                    var columnDef = questionDefinition["columns"]?.FirstOrDefault(c => c.Type == JTokenType.Object && c["value"]?.ToString() == selectedColumnValue);
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
                    if (isResultRow)
                    {
                        results.Add($"\"{RemoveHtmlTags(CleanupEscapedCharacters(columnDisplayText))}\"");
                    }
                    else
                    {
                        string cleanRowText = RemoveHtmlTags(CleanupEscapedCharacters(rowQuestionText));
                        string cleanColumnText = RemoveHtmlTags(CleanupEscapedCharacters(columnDisplayText));
                        
                        // Add quotes around field names for Power BI compatibility
                        string quotedRowText = NeedsQuoting(cleanRowText) ? $"\"{cleanRowText}\"" : cleanRowText;
                        results.Add($"{quotedRowText}: \"{cleanColumnText}\"");
                    }
                }
                
                return string.Join("; ", results);
            }
            return CleanupEscapedCharacters(responseValue.ToString());
        }

        /// <summary>
        /// Finds the comment associated with a multipletext question by looking for the next comment question in the definition.
        /// </summary>
        /// <param name="definition">The questionnaire definition.</param>
        /// <param name="multipletextQuestionName">The name of the multipletext question.</param>
        /// <param name="response">The questionnaire response.</param>
        /// <returns>The comment text if found, null otherwise.</returns>
        public string FindMultipletextComment(QuestionnaireDefinition definition, string multipletextQuestionName, QuestionnaireResponse response)
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
                    _logger.Trace($"Multipletext question {multipletextQuestionName} not found in definition");
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
                        if (response.HasValue(elementName))
                        {
                            string commentText = response.GetStringValue(elementName);
                            if (!string.IsNullOrEmpty(commentText))
                            {
                                _logger.Trace($"Found comment for multipletext {multipletextQuestionName}: {elementName} = {commentText}");
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
                _logger.Error($"Error finding multipletext comment for {multipletextQuestionName}: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Removes HTML tags and entities from text while preserving operators.
        /// </summary>
        public string RemoveHtmlTags(string input)
        {
            if (string.IsNullOrEmpty(input))
                return input;

            // Strip invalid XML characters that can cause CRM API errors.
            // This removes non-printable control characters (like 0x02) that are not valid in XML.
            input = System.Text.RegularExpressions.Regex.Replace(input, @"[\x00-\x08\x0B\x0C\x0E-\x1F]", string.Empty);

            // First, remove the "ite...m[digit]" pattern
            input = System.Text.RegularExpressions.Regex.Replace(input, @"^ite(.+?)m\d$", "$1", System.Text.RegularExpressions.RegexOptions.Singleline);

            // Replace <br> tags with new lines
            input = System.Text.RegularExpressions.Regex.Replace(input, "<br\\s*/?>", "\n", System.Text.RegularExpressions.RegexOptions.IgnoreCase);

            // Remove other HTML tags (e.g., <p>, <span>) but not < or > used as operators
            var withoutTags = System.Text.RegularExpressions.Regex.Replace(input, @"</?\w+((\s+\w+(\s*=\s*(?:"".*?""|'.*?'|[^'"">\s]+))?)+\s*|\s*)/?>", string.Empty);
            
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
                // Mathematical symbols
                .Replace("&ge;", ">=")
                .Replace("&le;", "<=")
                .Replace("&ne;", "!=")
                .Replace("&plusmn;", "±")
                .Replace("&times;", "×")
                .Replace("&divide;", "÷")
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

        /// <summary>
        /// Cleans up escaped characters that are artifacts of JSON escaping.
        /// </summary>
        private string CleanupEscapedCharacters(string input)
        {
            if (string.IsNullOrEmpty(input))
                return input;
            
            // Remove leading backslashes that are artifacts of JSON escaping
            // This handles cases like "\>" becoming ">"
            return input
                .Replace("\\>", ">")
                .Replace("\\<", "<")
                .Replace("\\&", "&")
                .Replace("\\\"", "\"")
                .Replace("\\'", "'");
        }

        /// <summary>
        /// Determines if text needs to be quoted for Power BI compatibility.
        /// </summary>
        private bool NeedsQuoting(string text)
        {
            if (string.IsNullOrEmpty(text))
                return false;
                
            // Add quotes if the text contains spaces, commas, colons, semicolons, or special characters
            return text.Contains(" ") || 
                   text.Contains(",") || 
                   text.Contains(":") || 
                   text.Contains(";") || 
                   text.Contains("/") || 
                   text.Contains("\\") || 
                   text.Contains("\"") ||
                   text.Contains("'");
        }
    }
} 