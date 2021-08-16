using System;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Xunit;

namespace ROMTS_GSRST.Plugins.Tests
{
    public class PreOperationovs_questionnaireCreateTests : UnitTestBase
    {

        public PreOperationovs_questionnaireCreateTests(XrmMockupFixture fixture) : base(fixture)
        {
        }

        [Fact]
        public void When_questionnaire_created_expect_one_questionnaire_version()
        {
            // ACT
            var questionnaireId = orgAdminService.Create(new ovs_Questionnaire()
            {
                ovs_Name = "New Test Questionnaire"
            });

            // ASSERT
            var questionnaire = orgAdminUIService.Retrieve(ovs_Questionnaire.EntityLogicalName, questionnaireId, new ColumnSet("ts_ovs_questionnaire_ovs_questionnaire")).ToEntity<ovs_Questionnaire>(); ;

            // Expect one questionnaire version with name "Version 1"
            Assert.Single(questionnaire.ts_ovs_questionnaire_ovs_questionnaire);
            var questionnaireVersion = questionnaire.ts_ovs_questionnaire_ovs_questionnaire.First().ToEntity<ts_questionnaireversion>();
            Assert.Equal("Version 1", questionnaireVersion.ts_name);
        }
    }
}
