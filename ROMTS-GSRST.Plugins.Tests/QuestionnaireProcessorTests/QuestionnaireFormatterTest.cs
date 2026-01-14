using Newtonsoft.Json.Linq;
using ROMTS_GSRST.Plugins.Tests.TestData;
using TSIS2.Plugins.QuestionnaireProcessor;
using Xunit;

namespace ROMTS_GSRST.Plugins.QuestionnaireProcessor
{
    /// <summary>
    /// Tests for QuestionnaireResponseFormatter class.
    /// </summary>
    public class QuestionnaireResponseFormatterTests
    {
        private readonly QuestionnaireResponseFormatter _formatter;

        public QuestionnaireResponseFormatterTests()
        {
            _formatter = new QuestionnaireResponseFormatter(new LoggerAdapter());
        }

        #region FormatResponse - Radiogroup Tests

        [Fact]
        public void FormatResponse_Radiogroup_ReturnsDisplayText()
        {
            var scenario = QuestionnaireSamples.CheckboxRadiogroupScenario;
            var questionnaireDef = new QuestionnaireDefinition(scenario.Definition);
            var questionnaireResp = new QuestionnaireResponse(scenario.Response);
            var questionDef = questionnaireDef.FindQuestionDefinition("complianceLevel");
            var responseValue = questionnaireResp.GetValue("complianceLevel");

            var formatted = _formatter.FormatResponse(responseValue, "radiogroup", questionDef);

            Assert.Equal("Partially Compliant", formatted);
        }

        [Fact]
        public void FormatResponse_Radiogroup_NoMatchingChoice_ReturnsRawValue()
        {
            var questionDef = JObject.Parse(@"{ ""choices"": [{ ""value"": ""a"", ""text"": ""Option A"" }] }");
            var responseValue = JToken.FromObject("unknown");

            var formatted = _formatter.FormatResponse(responseValue, "radiogroup", questionDef);

            Assert.Equal("unknown", formatted);
        }

        #endregion

        #region FormatResponse - Checkbox Tests

        [Fact]
        public void FormatResponse_Checkbox_ReturnsCommaSeparatedDisplayText()
        {
            var scenario = QuestionnaireSamples.CheckboxRadiogroupScenario;
            var questionnaireDef = new QuestionnaireDefinition(scenario.Definition);
            var questionnaireResp = new QuestionnaireResponse(scenario.Response);
            var questionDef = questionnaireDef.FindQuestionDefinition("areasInspected");
            var responseValue = questionnaireResp.GetValue("areasInspected");

            var formatted = _formatter.FormatResponse(responseValue, "checkbox", questionDef);

            Assert.Equal("Storage Area,Loading Dock", formatted);
        }

        [Fact]
        public void FormatResponse_Checkbox_SingleValue_ReturnsDisplayText()
        {
            var questionDef = JObject.Parse(@"{ ""choices"": [{ ""value"": ""opt1"", ""text"": { ""default"": ""Option One"" } }] }");
            var responseValue = JArray.Parse(@"[""opt1""]");

            var formatted = _formatter.FormatResponse(responseValue, "checkbox", questionDef);

            Assert.Equal("Option One", formatted);
        }

        #endregion

        #region FormatResponse - Matrix Tests

        [Fact]
        public void FormatResponse_Matrix_ReturnsFormattedRowsAndColumns()
        {
            var scenario = QuestionnaireSamples.MatrixScenario;
            var questionnaireDef = new QuestionnaireDefinition(scenario.Definition);
            var questionnaireResp = new QuestionnaireResponse(scenario.Response);
            var questionDef = questionnaireDef.FindQuestionDefinition("facilityChecklist");
            var responseValue = questionnaireResp.GetValue("facilityChecklist");

            var formatted = _formatter.FormatResponse(responseValue, "matrix", questionDef);

            Assert.Contains("Adequate Lighting", formatted);
            Assert.Contains("Satisfactory", formatted);
        }

        [Fact]
        public void FormatResponse_Matrix_ResultRow_ReturnsOnlyColumnText()
        {
            var questionDef = JObject.Parse(@"{
                ""columns"": [{ ""value"": ""sat"", ""text"": { ""default"": ""Satisfactory"" } }],
                ""rows"": [{ ""value"": ""Result"", ""text"": { ""default"": ""Result"" } }]
            }");
            var responseValue = JObject.Parse(@"{ ""Result"": ""sat"" }");

            var formatted = _formatter.FormatResponse(responseValue, "matrix", questionDef);

            Assert.Equal("\"Satisfactory\"", formatted);
        }

        #endregion

        #region FormatResponse - Multipletext Tests

        [Fact]
        public void FormatResponse_Multipletext_ReturnsFormattedFields()
        {
            var scenario = QuestionnaireSamples.MultipletextScenario;
            var questionnaireDef = new QuestionnaireDefinition(scenario.Definition);
            var questionnaireResp = new QuestionnaireResponse(scenario.Response);
            var questionDef = questionnaireDef.FindQuestionDefinition("contactInfo");
            var responseValue = questionnaireResp.GetValue("contactInfo");

            var formatted = _formatter.FormatResponse(responseValue, "multipletext", questionDef);

            Assert.Contains("John Doe", formatted);
            Assert.Contains("555-1234", formatted);
        }

        #endregion

        #region FormatResponse - Other Types

        [Fact]
        public void FormatResponse_Boolean_ReturnsStringValue()
        {
            var responseValue = JToken.FromObject(true);
            var formatted = _formatter.FormatResponse(responseValue, "boolean", null);
            Assert.Equal("True", formatted);
        }

        [Fact]
        public void FormatResponse_Finding_ReturnsEmptyString()
        {
            var responseValue = JToken.FromObject("any value");
            var formatted = _formatter.FormatResponse(responseValue, "finding", null);
            Assert.Equal(string.Empty, formatted);
        }

        [Fact]
        public void FormatResponse_NullValue_ReturnsEmptyString()
        {
            var formatted = _formatter.FormatResponse(null, "text", null);
            Assert.Equal(string.Empty, formatted);
        }

        #endregion

        #region RemoveHtmlTags Tests

        [Fact]
        public void RemoveHtmlTags_WithBreakTags_ReplacesWithSpaces()
        {
            var result = _formatter.RemoveHtmlTags("Line 1<br>Line 2<br/>Line 3");
            Assert.Equal("Line 1 Line 2 Line 3", result);
        }

        [Fact]
        public void RemoveHtmlTags_WithHtmlTags_RemovesTags()
        {
            var result = _formatter.RemoveHtmlTags("<p>Paragraph <strong>bold</strong> text</p>");
            Assert.Equal("Paragraph bold text", result);
        }

        [Fact]
        public void RemoveHtmlTags_WithHtmlEntities_DecodesEntities()
        {
            var result = _formatter.RemoveHtmlTags("Value &ge; 10 &amp; &lt; 20");
            Assert.Equal("Value >= 10 & < 20", result);
        }

        [Fact]
        public void RemoveHtmlTags_WithFrenchEntities_DecodesCorrectly()
        {
            var result = _formatter.RemoveHtmlTags("D&eacute;j&agrave; vu &ccedil;a");
            Assert.Equal("Déjà vu ça", result);
        }

        [Fact]
        public void RemoveHtmlTags_NullInput_ReturnsNull()
        {
            var result = _formatter.RemoveHtmlTags(null);
            Assert.Null(result);
        }

        [Fact]
        public void RemoveHtmlTags_WithNewlines_ReplacesWithSpaces()
        {
            var result = _formatter.RemoveHtmlTags("Line 1\nLine 2\nLine 3");
            Assert.Equal("Line 1 Line 2 Line 3", result);
        }

        [Fact]
        public void RemoveHtmlTags_WithTabs_ReplacesWithSpaces()
        {
            var result = _formatter.RemoveHtmlTags("Col1\tCol2\tCol3");
            Assert.Equal("Col1 Col2 Col3", result);
        }

        [Fact]
        public void RemoveHtmlTags_WithCarriageReturnNewlines_ReplacesWithSpaces()
        {
            var result = _formatter.RemoveHtmlTags("Line 1\r\nLine 2\r\nLine 3");
            Assert.Equal("Line 1 Line 2 Line 3", result);
        }

        #endregion
    }
}
