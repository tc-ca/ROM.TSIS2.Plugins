using ROMTS_GSRST.Plugins.Tests.TestData;
using TSIS2.Plugins.QuestionnaireProcessor;
using Xunit;

namespace ROMTS_GSRST.Plugins.QuestionnaireProcessor
{
    /// <summary>
    /// Tests for QuestionnaireResponse class - response data access and validation.
    /// </summary>
    public class QuestionnaireResponseTests
    {
        #region Constructor Tests

        [Fact]
        public void Constructor_ValidJson_ParsesSuccessfully()
        {
            // Arrange
            var scenario = QuestionnaireSamples.SimpleBooleanScenario;

            // Act
            var questionnaireResp = new QuestionnaireResponse(scenario.Response);

            // Assert
            Assert.NotNull(questionnaireResp.Response);
            Assert.False(questionnaireResp.IsEmpty());
        }

        [Fact]
        public void Constructor_NullJson_CreatesEmptyResponse()
        {
            // Act
            var questionnaireResp = new QuestionnaireResponse(null);

            // Assert
            Assert.True(questionnaireResp.IsEmpty());
        }

        [Fact]
        public void Constructor_EmptyJson_CreatesEmptyResponse()
        {
            // Act
            var questionnaireResp = new QuestionnaireResponse(string.Empty);

            // Assert
            Assert.True(questionnaireResp.IsEmpty());
        }

        [Fact]
        public void Constructor_InvalidJson_ThrowsJsonException()
        {
            // Arrange
            var invalidJson = "{ invalid }";

            // Act & Assert
            Assert.ThrowsAny<Newtonsoft.Json.JsonException>(() => new QuestionnaireResponse(invalidJson));
        }

        #endregion

        #region GetValue / HasValue Tests

        [Fact]
        public void GetValue_ExistingQuestion_ReturnsValue()
        {
            // Arrange
            var scenario = QuestionnaireSamples.SimpleBooleanScenario;
            var questionnaireResp = new QuestionnaireResponse(scenario.Response);

            // Act
            var value = questionnaireResp.GetValue("question1");

            // Assert
            Assert.NotNull(value);
        }

        [Fact]
        public void GetValue_NonExistingQuestion_ReturnsNull()
        {
            // Arrange
            var scenario = QuestionnaireSamples.SimpleBooleanScenario;
            var questionnaireResp = new QuestionnaireResponse(scenario.Response);

            // Act
            var value = questionnaireResp.GetValue("nonexistent");

            // Assert
            Assert.Null(value);
        }

        [Fact]
        public void GetValue_ExistingLongString_ReturnsValue()
        {
            var longValue = new string('a', 5000);
            var response = $"{{\"q1\": \"{longValue}\"}}";
            var questionnaireResp = new QuestionnaireResponse(response);

            var result = questionnaireResp.GetValue("q1");

            Assert.Equal(longValue, result.ToString());
        }

        [Fact]
        public void HasValue_ExistingQuestion_ReturnsTrue()
        {
            // Arrange
            var scenario = QuestionnaireSamples.SimpleBooleanScenario;
            var questionnaireResp = new QuestionnaireResponse(scenario.Response);

            // Act
            var hasValue = questionnaireResp.HasValue("question1");

            // Assert
            Assert.True(hasValue);
        }

        [Fact]
        public void HasValue_NonExistingQuestion_ReturnsFalse()
        {
            // Arrange
            var scenario = QuestionnaireSamples.SimpleBooleanScenario;
            var questionnaireResp = new QuestionnaireResponse(scenario.Response);

            // Act
            var hasValue = questionnaireResp.HasValue("nonexistent");

            // Assert
            Assert.False(hasValue);
        }

        #endregion

        #region GetBooleanValue Tests

        [Fact]
        public void GetBooleanValue_TrueValue_ReturnsTrue()
        {
            // Arrange
            var scenario = QuestionnaireSamples.SimpleBooleanScenario;
            var questionnaireResp = new QuestionnaireResponse(scenario.Response);

            // Act
            var value = questionnaireResp.GetBooleanValue("question1");

            // Assert
            Assert.True(value);
        }

        [Fact]
        public void GetBooleanValue_FalseValue_ReturnsFalse()
        {
            // Arrange
            var responseWithFalse = @"{""question1"": false}";
            var questionnaireResp = new QuestionnaireResponse(responseWithFalse);

            // Act
            var value = questionnaireResp.GetBooleanValue("question1");

            // Assert
            Assert.False(value);
        }

        [Fact]
        public void GetBooleanValue_NonBooleanValue_ReturnsNull()
        {
            // Arrange
            var responseWithString = @"{""question1"": ""not a boolean""}";
            var questionnaireResp = new QuestionnaireResponse(responseWithString);

            // Act
            var value = questionnaireResp.GetBooleanValue("question1");

            // Assert
            Assert.Null(value);
        }

        #endregion

        #region GetArrayValue Tests

        [Fact]
        public void GetArrayValue_ArrayResponse_ReturnsArray()
        {
            // Arrange
            var scenario = QuestionnaireSamples.CheckboxRadiogroupScenario;
            var questionnaireResp = new QuestionnaireResponse(scenario.Response);

            // Act
            var values = questionnaireResp.GetArrayValue("areasInspected");

            // Assert
            Assert.NotNull(values);
            Assert.Equal(2, values.Length);
            Assert.Contains("storage", values);
            Assert.Contains("loading", values);
        }

        [Fact]
        public void GetArrayValue_NonArrayResponse_ReturnsNull()
        {
            // Arrange
            var scenario = QuestionnaireSamples.CheckboxRadiogroupScenario;
            var questionnaireResp = new QuestionnaireResponse(scenario.Response);

            // Act
            var values = questionnaireResp.GetArrayValue("complianceLevel");

            // Assert
            Assert.Null(values);
        }

        #endregion

        #region GetObjectValue Tests

        [Fact]
        public void GetObjectValue_ObjectResponse_ReturnsObject()
        {
            // Arrange
            var scenario = QuestionnaireSamples.MatrixScenario;
            var questionnaireResp = new QuestionnaireResponse(scenario.Response);

            // Act
            var value = questionnaireResp.GetObjectValue("facilityChecklist");

            // Assert
            Assert.NotNull(value);
            Assert.Equal("sat", value["lighting"]?.ToString());
        }

        #endregion

        #region Detail Value Tests

        [Fact]
        public void GetDetailValue_ExistingDetail_ReturnsDetail()
        {
            // Arrange
            var scenario = QuestionnaireSamples.SimpleBooleanScenario;
            var questionnaireResp = new QuestionnaireResponse(scenario.Response);

            // Act
            var detail = questionnaireResp.GetDetailValue("question1");

            // Assert
            Assert.Equal("All areas inspected", detail);
        }

        [Fact]
        public void HasDetailValue_ExistingDetail_ReturnsTrue()
        {
            // Arrange
            var scenario = QuestionnaireSamples.SimpleBooleanScenario;
            var questionnaireResp = new QuestionnaireResponse(scenario.Response);

            // Act
            var hasDetail = questionnaireResp.HasDetailValue("question1");

            // Assert
            Assert.True(hasDetail);
        }

        #endregion

        #region GetAllQuestionNames Tests

        [Fact]
        public void GetAllQuestionNames_ExcludesDetailFields()
        {
            // Arrange
            var scenario = QuestionnaireSamples.SimpleBooleanScenario;
            var questionnaireResp = new QuestionnaireResponse(scenario.Response);

            // Act
            var names = questionnaireResp.GetAllQuestionNames();

            // Assert
            Assert.Equal(2, names.Count);
            Assert.Contains("question1", names);
            Assert.Contains("question2", names);
            Assert.DoesNotContain("question1-Detail", names);
        }

        [Fact]
        public void GetResponseCount_ExcludesDetailFields()
        {
            // Arrange
            var scenario = QuestionnaireSamples.SimpleBooleanScenario;
            var questionnaireResp = new QuestionnaireResponse(scenario.Response);

            // Act
            var count = questionnaireResp.GetResponseCount();

            // Assert
            Assert.Equal(2, count);
        }

        #endregion

        #region Diagnostic Tests

        [Fact]
        public void DetermineNoResponseReason_EmptyResponse_ReturnsEmptyMessage()
        {
            var questionnaireResp = new QuestionnaireResponse("{}");
            var reason = questionnaireResp.DetermineNoResponseReason(new System.Collections.Generic.List<string> { "q1" });
            Assert.Contains("All questions have empty responses", reason);
        }

        [Fact]
        public void DetermineNoResponseReason_WithMatches_ReturnsExplanation()
        {
            var questionnaireResp = new QuestionnaireResponse(@"{""q1"": ""val""}");
            var reason = questionnaireResp.DetermineNoResponseReason(new System.Collections.Generic.List<string> { "q1" });
            Assert.Contains("Questions had data but processing failed", reason);
        }

        #endregion
    }
}
