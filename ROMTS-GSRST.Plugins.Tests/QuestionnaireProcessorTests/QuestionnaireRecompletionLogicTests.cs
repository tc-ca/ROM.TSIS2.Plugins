using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using TSIS2.Plugins.QuestionnaireProcessor;
using Xunit;
using Xunit.Abstractions;

namespace ROMTS_GSRST.Plugins.Tests.QuestionnaireProcessorTests
{
    public class QuestionnaireRecompletionLogicTests : UnitTestBase
    {
        private readonly ILoggingService _logger;

        public QuestionnaireRecompletionLogicTests(XrmMockupFixture fixture, ITestOutputHelper output) : base(fixture)
        {
            _logger = new TestLoggingAdapter(output);
        }

        [Fact]
        public void Recompletion_SameNameDifferentNumber_UpdatesExistingRecordWithoutVersionBump()
        {
            // 1. Mock WOST (Parent)
            var wost = new Entity("msdyn_workorderservicetask")
            {
                ["msdyn_name"] = "200-000987-1"
            };
            var wostId = orgAdminService.Create(wost);

            // 2. Setup existing record in CRM (Child)
            var existingResponse = new Entity("ts_questionresponse")
            {
                ["ts_questionname"] = "q1",
                ["ts_questionnumber"] = 1,
                ["ts_msdyn_workorderservicetask"] = new EntityReference("msdyn_workorderservicetask", wostId),
                ["ts_response"] = "Old Answer",
                ["ts_version"] = 1,
                ["statecode"] = new OptionSetValue(0)
            };
            orgAdminService.Create(existingResponse);

            // 3. Update WOST with new questionnaire data
            var definition = @"{""pages"":[{""name"":""p1"",""elements"":[{""type"":""text"",""name"":""filler""},{""type"":""text"",""name"":""q1""}]}]}";
            var response = @"{""filler"": ""val"", ""q1"": ""Old Answer""}"; // Answer same, only number changed in definition

            var wostUpdate = new Entity("msdyn_workorderservicetask", wostId)
            {
                ["ovs_questionnairedefinition"] = definition,
                ["ovs_questionnaireresponse"] = response
            };
            orgAdminService.Update(wostUpdate);

            // 4. Run orchestrator
            QuestionnaireOrchestrator.ProcessQuestionnaire(orgAdminService, wostId, null, true, false, _logger);

            // 5. Verify
            var results = orgAdminService.RetrieveMultiple(new QueryExpression("ts_questionresponse") { ColumnSet = new ColumnSet(true) });
            var q1Record = results.Entities.FirstOrDefault(e => e.GetAttributeValue<string>("ts_questionname") == "q1");

            Assert.NotNull(q1Record);
            Assert.Equal(2, q1Record.GetAttributeValue<int>("ts_questionnumber"));
            Assert.Equal(1, q1Record.GetAttributeValue<int>("ts_version")); // Should NOT bump version for number change
        }

        [Fact]
        public void Recompletion_AnswerChange_BumpsVersion()
        {
            // 1. Mock WOST (Parent)
            var wost = new Entity("msdyn_workorderservicetask")
            {
                ["msdyn_name"] = "200-000987-1"
            };
            var wostId = orgAdminService.Create(wost);

            // 2. Setup existing record in CRM (Child)
            var existingResponse = new Entity("ts_questionresponse")
            {
                ["ts_questionname"] = "q1",
                ["ts_questionnumber"] = 1,
                ["ts_msdyn_workorderservicetask"] = new EntityReference("msdyn_workorderservicetask", wostId),
                ["ts_response"] = "Old Answer",
                ["ts_version"] = 5,
                ["statecode"] = new OptionSetValue(0)
            };
            orgAdminService.Create(existingResponse);

            // 3. Update WOST with new questionnaire data
            var definition = @"{""pages"":[{""name"":""p1"",""elements"":[{""type"":""text"",""name"":""q1""}]}]}";
            var response = @"{""q1"": ""New Answer""}";

            var wostUpdate = new Entity("msdyn_workorderservicetask", wostId)
            {
                ["ovs_questionnairedefinition"] = definition,
                ["ovs_questionnaireresponse"] = response
            };
            orgAdminService.Update(wostUpdate);

            // 4. Run orchestrator
            QuestionnaireOrchestrator.ProcessQuestionnaire(orgAdminService, wostId, null, true, false, _logger);

            // 5. Verify
            var results = orgAdminService.RetrieveMultiple(new QueryExpression("ts_questionresponse") { ColumnSet = new ColumnSet(true) });
            var qRecord = results.Entities.FirstOrDefault(e => e.GetAttributeValue<string>("ts_questionname") == "q1");

            Assert.NotNull(qRecord);
            Assert.Equal("New Answer", qRecord.GetAttributeValue<string>("ts_response"));
            Assert.Equal(6, qRecord.GetAttributeValue<int>("ts_version")); // Should bump from 5 to 6
        }

        [Fact]
        public void Recompletion_HiddenQuestionRemoval_ClearsDetails()
        {
            // 1. Mock WOST (Parent)
            var wost = new Entity("msdyn_workorderservicetask")
            {
                ["msdyn_name"] = "200-000987-1"
            };
            var wostId = orgAdminService.Create(wost);

            // 2. Setup existing parent record in CRM (Child of WOST)
            var parentRecord = new Entity("ts_questionresponse")
            {
                ["ts_questionname"] = "parentQ",
                ["ts_questionnumber"] = 1,
                ["ts_msdyn_workorderservicetask"] = new EntityReference("msdyn_workorderservicetask", wostId),
                ["ts_response"] = "No",
                ["ts_details"] = @"[{""question"":""Explain"",""answer"":""Old Explanation""}]",
                ["ts_version"] = 1,
                ["statecode"] = new OptionSetValue(0)
            };
            var parentId = orgAdminService.Create(parentRecord);

            // 3. Update WOST with new questionnaire data
            var definition = @"{""pages"":[{""name"":""p1"",""elements"":[
                {""type"":""radiogroup"",""name"":""parentQ"",""choices"":[""Yes"",""No""]},
                {""type"":""text"",""name"":""explainQ"",""visibleIf"":""{parentQ} = 'No'"", ""hideNumber"": true}
            ]}]}";
            var response = @"{""parentQ"": ""Yes""}"; // explainQ is gone

            var wostUpdate = new Entity("msdyn_workorderservicetask", wostId)
            {
                ["ovs_questionnairedefinition"] = definition,
                ["ovs_questionnaireresponse"] = response
            };
            orgAdminService.Update(wostUpdate);

            // 4. Run orchestrator
            QuestionnaireOrchestrator.ProcessQuestionnaire(orgAdminService, wostId, null, true, false, _logger);

            // 5. Verify parent record details are cleared
            var qRecord = orgAdminService.Retrieve("ts_questionresponse", parentId, new ColumnSet("ts_details"));
            Assert.Null(qRecord.GetAttributeValue<string>("ts_details"));
        }

        [Fact]
        public void Recompletion_HiddenQuestionContentChange_BumpsVersion()
        {
            // 1. Mock WOST (Parent)
            var wost = new Entity("msdyn_workorderservicetask")
            {
                ["msdyn_name"] = "200-000987-1"
            };
            var wostId = orgAdminService.Create(wost);

            // 2. Setup existing parent record in CRM (Child of WOST)
            var parentRecord = new Entity("ts_questionresponse")
            {
                ["ts_questionname"] = "parentQ",
                ["ts_questionnumber"] = 1,
                ["ts_msdyn_workorderservicetask"] = new EntityReference("msdyn_workorderservicetask", wostId),
                ["ts_response"] = "No",
                ["ts_details"] = @"[{""question"":""Explain"",""answer"":""Old Explanation""}]",
                ["ts_version"] = 1,
                ["statecode"] = new OptionSetValue(0)
            };
            var parentId = orgAdminService.Create(parentRecord);

            // 3. Update WOST with new questionnaire data where the hidden answer is modified
            var definition = @"{""pages"":[{""name"":""p1"",""elements"":[
                {""type"":""radiogroup"",""name"":""parentQ"",""choices"":[""Yes"",""No""]},
                {""type"":""text"",""name"":""explainQ"",""visibleIf"":""{parentQ} = 'No'"", ""hideNumber"": true, ""title"": ""Explain""}
            ]}]}";
            var response = @"{""parentQ"": ""No"", ""explainQ"": ""NEW Better Explanation""}";

            var wostUpdate = new Entity("msdyn_workorderservicetask", wostId)
            {
                ["ovs_questionnairedefinition"] = definition,
                ["ovs_questionnaireresponse"] = response
            };
            orgAdminService.Update(wostUpdate);

            // 4. Run orchestrator
            QuestionnaireOrchestrator.ProcessQuestionnaire(orgAdminService, wostId, null, true, false, _logger);

            // 5. Verify parent record version is bumped even though parent answer stayed "No"
            var qRecord = orgAdminService.Retrieve("ts_questionresponse", parentId, new ColumnSet("ts_version", "ts_details"));
            Assert.Equal(2, qRecord.GetAttributeValue<int>("ts_version"));
            Assert.Contains("NEW Better Explanation", qRecord.GetAttributeValue<string>("ts_details"));
        }


        [Fact]
        public void Recompletion_OrphanedRecord_GetsDeactivated()
        {
            // 1. Mock WOST (Parent)
            var wost = new Entity("msdyn_workorderservicetask")
            {
                ["msdyn_name"] = "200-000987-1"
            };
            var wostId = orgAdminService.Create(wost);

            // 2. Setup two existing records - one will become orphaned
            var record1 = new Entity("ts_questionresponse")
            {
                ["ts_name"] = "200-000987-1 [1]", // Required for XrmMockup
                ["ts_questionname"] = "q1",
                ["ts_questionnumber"] = 1,
                ["ts_msdyn_workorderservicetask"] = new EntityReference("msdyn_workorderservicetask", wostId),
                ["ts_response"] = "Answer 1",
                ["ts_version"] = 1,
                ["statecode"] = new OptionSetValue(0)
            };
            var record1Id = orgAdminService.Create(record1);

            var record2 = new Entity("ts_questionresponse")
            {
                ["ts_name"] = "200-000987-1 [2]", // Required for XrmMockup
                ["ts_questionname"] = "q2",
                ["ts_questionnumber"] = 2,
                ["ts_msdyn_workorderservicetask"] = new EntityReference("msdyn_workorderservicetask", wostId),
                ["ts_response"] = "Answer 2",
                ["ts_version"] = 1,
                ["statecode"] = new OptionSetValue(0)
            };
            var record2Id = orgAdminService.Create(record2);

            // 3. Update WOST - q2 no longer exists in definition
            var definition = @"{""pages"":[{""name"":""p1"",""elements"":[{""type"":""text"",""name"":""q1""}]}]}";
            var response = @"{""q1"": ""Answer 1""}";

            var wostUpdate = new Entity("msdyn_workorderservicetask", wostId)
            {
                ["ovs_questionnairedefinition"] = definition,
                ["ovs_questionnaireresponse"] = response
            };
            orgAdminService.Update(wostUpdate);

            // 4. Run orchestrator with recompletion=true
            QuestionnaireOrchestrator.ProcessQuestionnaire(orgAdminService, wostId, null, true, false, _logger);

            // 5. Verify q2 is deactivated
            var q2Record = orgAdminService.Retrieve("ts_questionresponse", record2Id, new ColumnSet("statecode"));
            Assert.Equal(1, q2Record.GetAttributeValue<OptionSetValue>("statecode").Value); // 1 = Inactive
        }
    }
}
