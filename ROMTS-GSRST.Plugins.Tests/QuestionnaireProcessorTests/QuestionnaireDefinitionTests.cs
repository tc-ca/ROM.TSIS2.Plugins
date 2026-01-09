using System.Linq;
using ROMTS_GSRST.Plugins.Tests.TestData;
using TSIS2.Plugins.QuestionnaireProcessor;
using Xunit;

namespace ROMTS_GSRST.Plugins.QuestionnaireProcessor
{
    /// <summary>
    /// Tests for QuestionnaireDefinition class - JSON parsing and question lookup.
    /// </summary>
    public class QuestionnaireDefinitionTests
    {
        #region Constructor Tests

        [Fact]
        public void Constructor_ValidJson_ParsesSuccessfully()
        {
            var scenario = QuestionnaireSamples.SimpleBooleanScenario;
            var questionnaireDef = new QuestionnaireDefinition(scenario.Definition);
            Assert.NotNull(questionnaireDef.Definition);
        }

        [Fact]
        public void Constructor_NullJson_ThrowsArgumentNullException()
        {
            Assert.Throws<System.ArgumentNullException>(() => new QuestionnaireDefinition(null));
        }

        [Fact]
        public void Constructor_EmptyJson_ThrowsArgumentNullException()
        {
            Assert.Throws<System.ArgumentNullException>(() => new QuestionnaireDefinition(string.Empty));
        }

        [Fact]
        public void Constructor_InvalidJson_ThrowsJsonException()
        {
            Assert.ThrowsAny<Newtonsoft.Json.JsonException>(() => new QuestionnaireDefinition("{ invalid }"));
        }

        [Fact]
        public void Constructor_MissingPages_DoesNotThrowButHasNoQuestions()
        {
            var questionnaireDef = new QuestionnaireDefinition("{}");
            Assert.Empty(questionnaireDef.GetAllQuestionNames());
        }

        #endregion

        #region FindQuestionDefinition Tests

        [Fact]
        public void FindQuestionDefinition_ExistingQuestion_ReturnsDefinition()
        {
            var scenario = QuestionnaireSamples.SimpleBooleanScenario;
            var questionnaireDef = new QuestionnaireDefinition(scenario.Definition);

            var questionDef = questionnaireDef.FindQuestionDefinition("question1");

            Assert.NotNull(questionDef);
            Assert.Equal("boolean", questionDef["type"]?.ToString());
        }

        [Fact]
        public void FindQuestionDefinition_NonExistingQuestion_ReturnsNull()
        {
            var scenario = QuestionnaireSamples.SimpleBooleanScenario;
            var questionnaireDef = new QuestionnaireDefinition(scenario.Definition);

            var questionDef = questionnaireDef.FindQuestionDefinition("nonexistent");

            Assert.Null(questionDef);
        }

        [Fact]
        public void FindQuestionDefinition_MatrixQuestion_ReturnsParentDefinition()
        {
            var scenario = QuestionnaireSamples.MatrixScenario;
            var questionnaireDef = new QuestionnaireDefinition(scenario.Definition);

            var questionDef = questionnaireDef.FindQuestionDefinition("facilityChecklist.lighting");

            Assert.NotNull(questionDef);
            Assert.Equal("matrix", questionDef["type"]?.ToString());
        }

        #endregion

        #region GetAllQuestionNames Tests

        [Fact]
        public void GetAllQuestionNames_ReturnsAllQuestions()
        {
            var scenario = QuestionnaireSamples.SimpleBooleanScenario;
            var questionnaireDef = new QuestionnaireDefinition(scenario.Definition);

            var questionNames = questionnaireDef.GetAllQuestionNames();

            Assert.Equal(4, questionNames.Count);
            Assert.Contains("question1", questionNames);
            Assert.Contains("question2", questionNames);
        }

        #endregion

        #region CollectVisibleQuestions Tests

        [Fact]
        public void CollectVisibleQuestions_AllTrue_HidesFindings()
        {
            var scenario = QuestionnaireSamples.SimpleBooleanScenario;
            var questionnaireDef = new QuestionnaireDefinition(scenario.Definition);
            var questionnaireResp = new QuestionnaireResponse(scenario.Response);

            var visibleQuestions = questionnaireDef.CollectVisibleQuestions(questionnaireResp.Response);

            Assert.Equal(2, visibleQuestions.Count);
            Assert.DoesNotContain(visibleQuestions, q => q["name"]?.ToString() == "finding-q1");
        }

        [Fact]
        public void CollectVisibleQuestions_OneFalse_ShowsRelatedFinding()
        {
            var scenario = QuestionnaireSamples.SimpleBooleanScenario;
            var responseWithFalse = @"{""question1"": false, ""question2"": true}";
            var questionnaireDef = new QuestionnaireDefinition(scenario.Definition);
            var questionnaireResp = new QuestionnaireResponse(responseWithFalse);

            var visibleQuestions = questionnaireDef.CollectVisibleQuestions(questionnaireResp.Response);

            Assert.Equal(3, visibleQuestions.Count);
            Assert.Contains(visibleQuestions, q => q["name"]?.ToString() == "finding-q1");
        }

        [Fact]
        public void CollectVisibleQuestions_AnyOfOperator_Matches()
        {
            var definition = @"{""pages"":[{""name"":""p1"",""elements"":[{""type"":""checkbox"",""name"":""q1"",""choices"":[""A"",""B"",""C""]},{""type"":""text"",""name"":""q2"",""visibleIf"":""{q1} anyof ['A', 'C']""}]}]}";
            var response = @"{""q1"": [""A""]}";
            var questionnaireDef = new QuestionnaireDefinition(definition);
            var questionnaireResp = new QuestionnaireResponse(response);

            var visibleQuestions = questionnaireDef.CollectVisibleQuestions(questionnaireResp.Response);

            Assert.Equal(2, visibleQuestions.Count);
            Assert.Contains(visibleQuestions, q => q["name"]?.ToString() == "q2");
        }

        [Fact]
        public void CollectVisibleQuestions_ContainsOperator_Matches()
        {
            var definition = @"{""pages"":[{""name"":""p1"",""elements"":[{""type"":""checkbox"",""name"":""q1"",""choices"":[""A"",""B""]},{""type"":""text"",""name"":""q2"",""visibleIf"":""{q1} contains 'A'""}]}]}";
            var response = @"{""q1"": [""A"", ""B""]}";
            var questionnaireDef = new QuestionnaireDefinition(definition);
            var questionnaireResp = new QuestionnaireResponse(response);

            var visibleQuestions = questionnaireDef.CollectVisibleQuestions(questionnaireResp.Response);

            Assert.Equal(2, visibleQuestions.Count);
            Assert.Contains(visibleQuestions, q => q["name"]?.ToString() == "q2");
        }

        #endregion

        #region GetTextFieldValue Tests

        [Fact]
        public void GetTextFieldValue_LocalizedObject_ReturnsDefaultText()
        {
            var localizedToken = Newtonsoft.Json.Linq.JObject.Parse(@"{""default"":""English"",""fr"":""French""}");
            var result = QuestionnaireDefinition.GetTextFieldValue(localizedToken);
            Assert.Equal("English", result);
        }

        [Fact]
        public void GetTextFieldValue_FrenchLocale_ReturnsFrenchText()
        {
            var localizedToken = Newtonsoft.Json.Linq.JObject.Parse(@"{""default"":""English"",""fr"":""French""}");
            var result = QuestionnaireDefinition.GetTextFieldValue(localizedToken, "fr");
            Assert.Equal("French", result);
        }

        [Fact]
        public void GetTextFieldValue_SimpleString_ReturnsString()
        {
            var simpleToken = Newtonsoft.Json.Linq.JToken.FromObject("Simple Text");
            var result = QuestionnaireDefinition.GetTextFieldValue(simpleToken);
            Assert.Equal("Simple Text", result);
        }

        #endregion

        #region ParseParentQuestionName Tests

        [Fact]
        public void ParseParentQuestionName_MatrixVisibleIf_ReturnsMatrixNameOnly()
        {
            // Arrange
            var visibleIf = "{MatrixName.RowName} = 'Not Met'";

            // Act
            var result = QuestionnaireDefinition.ParseParentQuestionName(visibleIf);

            // Assert
            Assert.Equal("MatrixName", result);
        }

        [Fact]
        public void ParseParentQuestionName_SimpleVisibleIf_ReturnsQuestionName()
        {
            // Arrange
            var visibleIf = "{question1} = true";

            // Act
            var result = QuestionnaireDefinition.ParseParentQuestionName(visibleIf);

            // Assert
            Assert.Equal("question1", result);
        }

        #endregion

        #region Or Operator Tests

        [Fact]
        public void CollectVisibleQuestions_OrOperator_MatchesFirstCondition()
        {
            var definition = @"{""pages"":[{""name"":""p1"",""elements"":[
                {""type"":""radiogroup"",""name"":""q1"",""choices"":[""A"",""B"",""C""]},
                {""type"":""text"",""name"":""q2"",""visibleIf"":""{q1} = 'A' or {q1} = 'B'""}
            ]}]}";
            var response = @"{""q1"": ""A""}";
            var questionnaireDef = new QuestionnaireDefinition(definition);
            var questionnaireResp = new QuestionnaireResponse(response);

            var visibleQuestions = questionnaireDef.CollectVisibleQuestions(questionnaireResp.Response);

            Assert.Equal(2, visibleQuestions.Count);
            Assert.Contains(visibleQuestions, q => q["name"]?.ToString() == "q2");
        }

        [Fact]
        public void CollectVisibleQuestions_OrOperator_MatchesSecondCondition()
        {
            var definition = @"{""pages"":[{""name"":""p1"",""elements"":[
                {""type"":""radiogroup"",""name"":""q1"",""choices"":[""A"",""B"",""C""]},
                {""type"":""text"",""name"":""q2"",""visibleIf"":""{q1} = 'A' or {q1} = 'B'""}
            ]}]}";
            var response = @"{""q1"": ""B""}";
            var questionnaireDef = new QuestionnaireDefinition(definition);
            var questionnaireResp = new QuestionnaireResponse(response);

            var visibleQuestions = questionnaireDef.CollectVisibleQuestions(questionnaireResp.Response);

            Assert.Equal(2, visibleQuestions.Count);
            Assert.Contains(visibleQuestions, q => q["name"]?.ToString() == "q2");
        }

        [Fact]
        public void CollectVisibleQuestions_OrOperator_NeitherMatches()
        {
            var definition = @"{""pages"":[{""name"":""p1"",""elements"":[
                {""type"":""radiogroup"",""name"":""q1"",""choices"":[""A"",""B"",""C""]},
                {""type"":""text"",""name"":""q2"",""visibleIf"":""{q1} = 'A' or {q1} = 'B'""}
            ]}]}";
            var response = @"{""q1"": ""C""}";
            var questionnaireDef = new QuestionnaireDefinition(definition);
            var questionnaireResp = new QuestionnaireResponse(response);

            var visibleQuestions = questionnaireDef.CollectVisibleQuestions(questionnaireResp.Response);

            Assert.Single(visibleQuestions);
            Assert.DoesNotContain(visibleQuestions, q => q["name"]?.ToString() == "q2");
        }

        #endregion

        #region Duplicate Question Name Tests

        [Fact]
        public void Constructor_DuplicateQuestionNames_FirstWins()
        {
            // Definition with duplicate question names - first definition should win
            var definition = @"{""pages"":[
                {""name"":""p1"",""elements"":[{""type"":""text"",""name"":""q1"",""title"":""First Title""}]},
                {""name"":""p2"",""elements"":[{""type"":""boolean"",""name"":""q1"",""title"":""Second Title""}]}
            ]}";
            var questionnaireDef = new QuestionnaireDefinition(definition);

            var questionDef = questionnaireDef.FindQuestionDefinition("q1");

            Assert.NotNull(questionDef);
            Assert.Equal("text", questionDef["type"]?.ToString());
        }

        #endregion

        #region Nested Panel Tests

        [Fact]
        public void GetAllQuestionNames_NestedPanels_CollectsAllQuestions()
        {
            var definition = @"{""pages"":[{""name"":""p1"",""elements"":[
                {""type"":""panel"",""name"":""panel1"",""elements"":[
                    {""type"":""text"",""name"":""nestedQ1""},
                    {""type"":""panel"",""name"":""innerPanel"",""elements"":[
                        {""type"":""text"",""name"":""deepNestedQ""}
                    ]}
                ]},
                {""type"":""text"",""name"":""topLevelQ""}
            ]}]}";
            var questionnaireDef = new QuestionnaireDefinition(definition);

            var questionNames = questionnaireDef.GetAllQuestionNames();

            Assert.Contains("nestedQ1", questionNames);
            Assert.Contains("deepNestedQ", questionNames);
            Assert.Contains("topLevelQ", questionNames);
            Assert.DoesNotContain("panel1", questionNames);
            Assert.DoesNotContain("innerPanel", questionNames);
        }

        #endregion
    }
}
