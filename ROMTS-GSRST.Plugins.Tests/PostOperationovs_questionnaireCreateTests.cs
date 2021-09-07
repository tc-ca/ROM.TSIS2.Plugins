using System;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Xunit;

namespace ROMTS_GSRST.Plugins.Tests
{
    public class PostOperationovs_questionnaireCreateTests : UnitTestBase
    {

        public PostOperationovs_questionnaireCreateTests(XrmMockupFixture fixture) : base(fixture)
        {
        }

        [Fact]
        public void When_questionnaire_created_expect_one_questionnaire_version()
        {
            using (var context = new Xrm(orgAdminUIService))
            {
                // ACT
                var questionnaire = new ovs_Questionnaire();
                questionnaire.Id = orgAdminService.Create(new ovs_Questionnaire()
                {
                    ovs_Name = "New Test Questionnaire"
                });

                // ASSERT
                // Get all questionnaire versions for the newly created questionnaire
                var versions = context.ts_questionnaireversionSet.Where(x => x.ts_ovs_questionnaire == questionnaire.ToEntityReference());

                // Expect one questionnaire version with name "Version 1"
                Assert.Single(versions);
                var questionnaireVersion = versions.First();
                Assert.Equal("New Test Questionnaire - Version 1", questionnaireVersion.ts_name);
            }
        }
    }
}
