using System;
using System.Data.SqlClient;
using System.Linq;
using System.Reflection.Emit;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Xunit;

namespace ROMTS_GSRST.Plugins.Tests
{
    public class PostOperationts_operationriskassessmentTests : UnitTestBase
    {
        public PostOperationts_operationriskassessmentTests(XrmMockupFixture fixture) : base(fixture) { }

        [Fact]
        public void When_operation_risk_assessment_is_created_Then_populate_Risk_Criteria_Responses_and_Discretionary_Score_Responses()
        {
            // Arrange
            var operationType = orgAdminService.Create(new ovs_operationtype
            {
                
            });
            var operation = orgAdminUIService.Create(new ovs_operation
            {
                ovs_OperationTypeId = new EntityReference(ovs_operationtype.EntityLogicalName, operationType),
            });

            // Create 3 test risk criterias
            var riskCriteria1 = orgAdminUIService.Create(new ts_riskcriteria
            {
                ts_Name = "Risk Criteria 1"
            });

            var riskCriteria2 = orgAdminUIService.Create(new ts_riskcriteria
            {
                ts_Name = "Risk Criteria 2"
            });

            var riskCriteria3 = orgAdminUIService.Create(new ts_riskcriteria
            {
                ts_Name = "Risk Criteria 3"
            });

            //Associate risk criterias to operation type
            orgAdminUIService.Associate("ovs_operationtype", operationType, new Relationship("ts_riskcriteria_ovs_operationtype"), new EntityReferenceCollection
            {
                new EntityReference("ts_riskcriteria", riskCriteria1),
                new EntityReference("ts_riskcriteria", riskCriteria2),
                new EntityReference("ts_riskcriteria", riskCriteria3)
            });

            // Create discretionary factor grouping
            var discretionaryFactorGrouping = orgAdminUIService.Create(new ts_discretionaryfactorgrouping
            {
                ts_Name = "Discretionary Factor Grouping",
                ts_operationtype = new EntityReference(ovs_operationtype.EntityLogicalName, operationType)
            });

            // Create discretionary factors
            var discretionaryFactor1 = orgAdminUIService.Create(new ts_discretionaryfactor
            {
                ts_Name = "Discretionary Factor 1",
                ts_discretionaryfactorgrouping = new EntityReference(ts_discretionaryfactorgrouping.EntityLogicalName, discretionaryFactorGrouping),
            });
            var discretionaryFactor2 = orgAdminUIService.Create(new ts_discretionaryfactor
            {
                ts_Name = "Discretionary Factor 2",
                ts_discretionaryfactorgrouping = new EntityReference(ts_discretionaryfactorgrouping.EntityLogicalName, discretionaryFactorGrouping),
            });
            var discretionaryFactor3 = orgAdminUIService.Create(new ts_discretionaryfactor
            {
                ts_Name = "Discretionary Factor 3",
                ts_discretionaryfactorgrouping = new EntityReference(ts_discretionaryfactorgrouping.EntityLogicalName, discretionaryFactorGrouping),
            });

            // Act
            var operationRiskAssessment = orgAdminUIService.Create(new ts_operationriskassessment
            {
                ts_operation = new EntityReference(ovs_operation.EntityLogicalName, operation)
            });

            // Assert
            // Three Risk Criteria Responses should be created
            var riskCriteriaResponses = orgAdminUIService.RetrieveMultiple(new QueryExpression(ts_riskcriteriaresponse.EntityLogicalName)
            {
                ColumnSet = new ColumnSet(true),
                Criteria = new FilterExpression
                {
                    Conditions =
                    {
                        new ConditionExpression("ts_operationriskassessment", ConditionOperator.Equal, operationRiskAssessment)
                    }
                }
            });
            Assert.Equal(3, riskCriteriaResponses.Entities.Count);
            // Three Discretionary Factor Responses should be created
            var discretionaryFactorResponses = orgAdminUIService.RetrieveMultiple(new QueryExpression(ts_discretionaryfactorresponse.EntityLogicalName)
            {
                ColumnSet = new ColumnSet(true),
                Criteria = new FilterExpression
                {
                    Conditions =
                    {
                        new ConditionExpression("ts_operationriskassessment", ConditionOperator.Equal, operationRiskAssessment)
                    }
                }
            });
            Assert.Equal(3, discretionaryFactorResponses.Entities.Count);
        }

        //Write a test to verify that any other active operation risk assessments are set to inactive
        [Fact]
        public void When_operation_risk_assessment_is_created_Then_set_any_other_active_operation_risk_assessments_to_inactive()
        {
            // Arrange
            var operationType = orgAdminService.Create(new ovs_operationtype
            {
                
            });
            var operation = orgAdminUIService.Create(new ovs_operation
            {
                ovs_OperationTypeId = new EntityReference(ovs_operationtype.EntityLogicalName, operationType),
            });

            var operationRiskAssessment1 = orgAdminUIService.Create(new ts_operationriskassessment
            {
                ts_operation = new EntityReference(ovs_operation.EntityLogicalName, operation),
                statecode = ts_operationriskassessmentState.Active
            });

            var operationRiskAssessment2 = orgAdminUIService.Create(new ts_operationriskassessment
            {
                ts_operation = new EntityReference(ovs_operation.EntityLogicalName, operation),
                statecode = ts_operationriskassessmentState.Active
            });

            // Act
            var operationRiskAssessment3 = orgAdminUIService.Create(new ts_operationriskassessment
            {
                ts_operation = new EntityReference(ovs_operation.EntityLogicalName, operation),
                statecode = ts_operationriskassessmentState.Active
            });

            // Assert
            var inactiveOperationRiskAssessments = orgAdminUIService.RetrieveMultiple(new QueryExpression(ts_operationriskassessment.EntityLogicalName)
            {
                ColumnSet = new ColumnSet(true),
                Criteria = new FilterExpression
                {
                    Conditions =
                    {
                        new ConditionExpression("ts_operation", ConditionOperator.Equal, operation),
                        new ConditionExpression("statecode", ConditionOperator.Equal, ts_operationriskassessmentState.Inactive)
                    }
                }
            });
            Assert.Equal(2, inactiveOperationRiskAssessments.Entities.Count);
        }
    }
}
