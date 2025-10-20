// Commented out since it no longer works for .Net Framework 4.7.1
//using System;
//using System.Linq;
//using Microsoft.Xrm.Sdk;
//using Microsoft.Xrm.Sdk.Query;
//using Xunit;

//namespace ROMTS_GSRST.Plugins.Tests
//{
//    public class PreOperationts_questionnaireversionUpdateTests : UnitTestBase
//    {

//        public PreOperationts_questionnaireversionUpdateTests(XrmMockupFixture fixture) : base(fixture) { }

//        [Fact]
//        public void When_questionnaire_version_start_date_updated_update_end_date_previous_version()
//        {
//            // ARRANGE

//            var questionnaireId = orgAdminUIService.Create(new ovs_Questionnaire()
//            {
//                ovs_Name = "New Test Questionnaire"
//            });
//            var query = new QueryExpression(ts_questionnaireversion.EntityLogicalName)
//            {
//                ColumnSet = new ColumnSet()
//            };

//            var currentQuestionnaireVersion = orgAdminUIService.RetrieveMultiple(query).Entities.Cast<ts_questionnaireversion>().ToList().FirstOrDefault();
//            orgAdminUIService.Update(new ts_questionnaireversion
//            {
//                Id = currentQuestionnaireVersion.Id,
//                ts_effectivestartdate = DateTime.Now,

//            });

//            var previousQuestionnaireVersion = orgAdminUIService.Create(new ts_questionnaireversion()
//            {
//                ts_effectivestartdate = DateTime.Now.AddDays(-5),
//                ts_effectiveenddate = null,
//                ts_ovs_questionnaire = new EntityReference(ovs_Questionnaire.EntityLogicalName, questionnaireId)
//            });

//            //ACT

//            orgAdminUIService.Update(new ts_questionnaireversion
//            {
//                Id = currentQuestionnaireVersion.Id,
//                ts_effectivestartdate = DateTime.Now,

//            });
//            //ASSERT
//            var query1 = new QueryExpression(ts_questionnaireversion.EntityLogicalName)
//            {
//                ColumnSet = new ColumnSet("ts_effectivestartdate", "ts_effectiveenddate")
//            };

//            var versions = orgAdminUIService.RetrieveMultiple(query1).Entities.Cast<ts_questionnaireversion>().OrderBy(d => d.ts_effectivestartdate).ToList();

//            //Expect that end date of privious version is the day before of current version start date
//            Assert.Equal(Convert.ToDateTime(versions[0].ts_effectiveenddate), Convert.ToDateTime(versions[1].ts_effectivestartdate).AddDays(-1));
//        }
//    }
//}
