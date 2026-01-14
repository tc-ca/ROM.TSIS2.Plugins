using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Xrm.Sdk;
using ROMTS_GSRST.Plugins.Tests;
using ROMTS_GSRST.Plugins.Tests.QuestionnaireProcessorTests;
using ROMTS_GSRST.Plugins.Tests.TestData;
using TSIS2.Plugins.QuestionnaireProcessor;
using Xunit;
using Xunit.Abstractions;

namespace ROMTS_GSRST.Plugins.QuestionnaireProcessor
{
    public class QuestionnaireIntegrationScenarioTests : UnitTestBase
    {
        private readonly QuestionnaireResponseFormatter _formatter;
        private readonly ITestOutputHelper _output;

        public QuestionnaireIntegrationScenarioTests(XrmMockupFixture fixture, ITestOutputHelper output) : base(fixture)
        {
            _output = output;
            _formatter = new QuestionnaireResponseFormatter(new TestLoggingAdapter(_output));
        }

        #region Automatic Tests for All Scenarios

        [Theory]
        [MemberData(nameof(QuestionnaireSamples.AllScenarios), MemberType = typeof(QuestionnaireSamples))]
        public void AllScenarios_ParseDefinitionSuccessfully(string scenarioName, string definition, string response)
        {
            // Ensure all theory parameters are used
            Assert.False(string.IsNullOrWhiteSpace(response));

            var questionnaireDef = new QuestionnaireDefinition(definition);
            Assert.NotNull(questionnaireDef.Definition);
            Assert.True(questionnaireDef.Definition.ContainsKey("pages") || questionnaireDef.Definition.ContainsKey("elements"), $"Scenario '{scenarioName}' should have 'pages' or 'elements' section in definition");
        }

        [Theory]
        [MemberData(nameof(QuestionnaireSamples.AllScenarios), MemberType = typeof(QuestionnaireSamples))]
        public void AllScenarios_ParseResponseSuccessfully(string scenarioName, string definition, string response)
        {
            // Ensure all theory parameters are used
            Assert.False(string.IsNullOrWhiteSpace(scenarioName));
            Assert.False(string.IsNullOrWhiteSpace(definition));

            var questionnaireResp = new QuestionnaireResponse(response);
            Assert.NotNull(questionnaireResp.Response);
            Assert.False(questionnaireResp.IsEmpty());
        }

        [Theory]
        [MemberData(nameof(QuestionnaireSamples.AllScenarios), MemberType = typeof(QuestionnaireSamples))]
        public void AllScenarios_CollectVisibleQuestionsWithoutError(string scenarioName, string definition, string response)
        {
            // Ensure all theory parameters are used
            Assert.False(string.IsNullOrWhiteSpace(scenarioName));

            var questionnaireDef = new QuestionnaireDefinition(definition);
            var questionnaireResp = new QuestionnaireResponse(response);
            var visibleQuestions = questionnaireDef.CollectVisibleQuestions(questionnaireResp.Response);
            Assert.NotNull(visibleQuestions);
        }

        [Theory]
        [MemberData(nameof(QuestionnaireSamples.AllScenarios), MemberType = typeof(QuestionnaireSamples))]
        public void AllScenarios_GetAllQuestionNamesWithoutError(string scenarioName, string definition, string response)
        {
            // Ensure all theory parameters are used
            Assert.False(string.IsNullOrWhiteSpace(response));

            var questionnaireDef = new QuestionnaireDefinition(definition);
            var questionNames = questionnaireDef.GetAllQuestionNames();
            Assert.NotNull(questionNames);
            Assert.True(questionNames.Count > 0, $"Scenario '{scenarioName}' should have at least one question");
        }

        #endregion

        #region Specific Scenario Tests

        [Fact]
        public void SimpleBooleanScenario_AllCompliant_NoFindingsVisible()
        {
            var scenario = QuestionnaireSamples.SimpleBooleanScenario;
            var questionnaireDef = new QuestionnaireDefinition(scenario.Definition);
            var questionnaireResp = new QuestionnaireResponse(scenario.Response);

            var visibleQuestions = questionnaireDef.CollectVisibleQuestions(questionnaireResp.Response);

            Assert.Equal(2, visibleQuestions.Count);
            Assert.DoesNotContain(visibleQuestions, q => q["name"]?.ToString() == "finding-q1");
        }

        [Fact]
        public void CheckboxScenario_FormatsAllSelectedValues()
        {
            var scenario = QuestionnaireSamples.CheckboxRadiogroupScenario;
            var questionnaireDef = new QuestionnaireDefinition(scenario.Definition);
            var questionnaireResp = new QuestionnaireResponse(scenario.Response);
            var questionDef = questionnaireDef.FindQuestionDefinition("areasInspected");
            var responseValue = questionnaireResp.GetValue("areasInspected");

            var formatted = _formatter.FormatResponse(responseValue, "checkbox", questionDef);

            Assert.Contains("Storage Area", formatted);
            Assert.Contains("Loading Dock", formatted);
        }

        [Fact]
        public void MatrixScenario_FormatsRowsAndColumns()
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
        public void ConditionalVisibility_CargoOnly_ShowsCargoQuestion()
        {
            var scenario = QuestionnaireSamples.ConditionalVisibilityScenario;
            var questionnaireDef = new QuestionnaireDefinition(scenario.Definition);
            var questionnaireResp = new QuestionnaireResponse(scenario.Response);

            var visibleQuestions = questionnaireDef.CollectVisibleQuestions(questionnaireResp.Response);
            var visibleNames = visibleQuestions.Select(q => q["name"]?.ToString()).ToList();

            Assert.Contains("cargoSecurityCheck", visibleNames);
            Assert.DoesNotContain("passengerScreening", visibleNames);
        }

        [Fact]
        public void RegressionScenario_MultiLayerLogic_WorksCorrectly()
        {
            var scenario = QuestionnaireSamples.Get("RegressionScenario");
            var questionnaireDef = new QuestionnaireDefinition(scenario.Definition);
            var questionnaireResp = new QuestionnaireResponse(scenario.Response);

            var visibleQuestions = questionnaireDef.CollectVisibleQuestions(questionnaireResp.Response);
            var visibleNames = visibleQuestions.Select(q => q["name"]?.ToString()).ToList();

            // q3 is visible if q2 contains 'v1' AND q1 = 'show'
            Assert.Contains("q1", visibleNames);
            Assert.Contains("q2", visibleNames);
            Assert.Contains("q3", visibleNames);
        }

        [Fact]
        public void LongQuestionnaire_CreatesExpectedNumberOfQuestions()
        {
            var scenario = QuestionnaireSamples.Get("LongQuestionnaire");
            var questionnaireDef = new QuestionnaireDefinition(scenario.Definition);
            var questionnaireResp = new QuestionnaireResponse(scenario.Response);

            var visibleQuestions = questionnaireDef.CollectVisibleQuestions(questionnaireResp.Response);

            Assert.Equal(55, visibleQuestions.Count);
        }

        [Fact]
        public void LongQuestionnaire_CalculatesExpectedCrmRecords()
        {
            var scenario = QuestionnaireSamples.Get("LongQuestionnaire");


            var wost = new Entity("msdyn_workorderservicetask")
            {
                ["msdyn_name"] = "Test Task",
                ["ovs_questionnaireresponse"] = scenario.Response,
                ["ovs_questionnairedefinition"] = scenario.Definition
            };
            var wostId = orgAdminService.Create(wost);

            var logger = new TestLoggingAdapter(_output);

            var dummyRef = new EntityReference("ovs_questionnairedefinition", Guid.NewGuid());

            var result = QuestionnaireOrchestrator.ProcessQuestionnaire(
                orgAdminService,
                wostId,
                dummyRef,
                isRecompletion: false,
                simulationMode: true,
                logger: logger,
                includeQuestionInventory: true
            );

            // 3. Assertions
            // Expected: 55 visible questions -> 3 merged -> 52 created records
            _output.WriteLine($"Result: {result.TotalCreatedOrUpdatedRecords} records, {result.HiddenMergedCount} merged.");

            Assert.Equal(55, result.VisibleQuestionCount);
            Assert.Equal(3, result.HiddenMergedCount);
            Assert.Equal(52, result.TotalCreatedOrUpdatedRecords);
        }

        #endregion
    }
}


