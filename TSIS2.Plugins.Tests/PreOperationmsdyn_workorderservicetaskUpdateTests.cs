using System;
using System.Collections.Generic;
using System.Reflection;
using System.Linq;
using Xunit;
using FakeItEasy;
using FakeXrmEasy;
using Microsoft.Xrm.Sdk;
using TSIS2.Common;

namespace TSIS2.Plugins.Tests
{
 
    public class PreOperationmsdyn_workorderservicetaskUpdateTests
    {
        [Fact]
        public void When_ovs_questionnaireresponse_contains_finding_but_work_order_service_task_is_not_100_percent_complete_expect_ovs_finding_not_created()
        {
            /**********
            * ARRANGE
            **********/
            var context = new XrmFakedContext();

            // Given a work order service task that
            var regulatedEntityId = Guid.NewGuid();
            var regulatedEntity = new Account()
            {
                Id = regulatedEntityId,
                Name = "Test Regulated Entity"
            };

            var workOrderId = Guid.NewGuid();
            var workOrder = new msdyn_workorder()
            {
                Id = workOrderId,
                msdyn_ServiceRequest = null, // does not already belong to a case (Incident)
                ovs_regulatedentity = new EntityReference(Account.EntityLogicalName, regulatedEntityId)
            };

            var workOrderServiceTaskId = Guid.NewGuid();
            var workOrderServiceTask = new msdyn_workorderservicetask()
            {
                Id = workOrderServiceTaskId,
                msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId), // belongs to a work order
                ovs_InspectionStatus = false,
                ovs_QuestionnaireResponse = @"
                {
                    ""finding-sq_162"": {
                        ""provisionReference"": ""Section 2"",
                        ""provisionText"": ""<strong>Application</strong></br><strong><mark>Section 2</mark></strong>: Sections 3 to 15 apply in respect of the following passenger-carrying flights — or in respect of air carriers conducting such flights — if the passengers, the property in the possession or control of the passengers and the belongings or baggage that the passengers give to the air carrier for transport are subject to screening that is carried out — in Canada under the Aeronautics Act or in another country by the person or entity responsible for the screening of such persons, property and belongings or baggage — before boarding:</br><ul style='list-style-type:none;'><li><strong>(a)</strong> domestic flights that depart from Canadian aerodromes and that are conducted by air carriers under Subpart 5 of Part VII of the Canadian Aviation Regulations ; and</li><li><strong>(b)</strong> international flights that depart from or will arrive at Canadian aerodromes and that are conducted by air carriers</li><ul style='list-style-type:none;'><li><strong>(i)</strong> under Subpart 1 of Part VII of the Canadian Aviation Regulations using aircraft that have a maximum certificated take-off weight of more than 8 618 kg (19,000 pounds) or have a seating configuration, excluding crew seats, of 20 or more, or</li><li><strong>(ii)</strong> under Subpart 5 of Part VII of the Canadian Aviation Regulations.</li></ul></ul>"",
                        ""comments"": ""new comments"",
                        ""documentaryEvidence"": ""C:\\fakepath\\newfile.png""
                    }
                }
                "
            };

            context.Initialize(
                new List<Entity>() {
                    regulatedEntity,
                    workOrder,
                    workOrderServiceTask
                }
            );

            ParameterCollection inputParams = new ParameterCollection();
            inputParams.Add("Target", workOrderServiceTask);
            ParameterCollection outputParams = new ParameterCollection();
            outputParams.Add("id", workOrderServiceTaskId);
            EntityImageCollection preEntityImages = new EntityImageCollection();
            preEntityImages.Add("PreImage", workOrderServiceTask);

            /**********
            * ACT
            **********/
            // Execute the PreOperationmsdyn_workorderservicetaskUpdate plugin with the defined workOrderServiceTask as a target
            context.ExecutePluginWith<PreOperationmsdyn_workorderservicetaskUpdate>(inputParams, outputParams, preEntityImages, null);

            /**********
             * ASSERT
             **********/
            var findings = context.CreateQuery<ovs_Finding>().ToList();

            // Expect no findings created because questionnaire is not completed 
            Assert.True(findings.Count == 0, "Expected no findings");
        }

        [Fact]
        public void When_ovs_questionnaireresponse_contains_finding_expect_do_not_recreate_existing_ovs_finding()
        {

            /**********
             * ARRANGE
             **********/
            var context = new XrmFakedContext();

            // Given a work order service task that
            // - belongs to a work order
            // - belongs to a case (Incident)
            // - has a SurveyJS questionnaire response saved in the ovs_QuestionnaireResponse field
            var regulatedEntityId = Guid.NewGuid();
            var regulatedEntity = new Account()
            {
                Id = regulatedEntityId,
                Name = "Test Regulated Entity"
            };

            var incidentId = Guid.NewGuid();
            var incident = new Incident()
            {
                Id = incidentId
            };

            var workOrderId = Guid.NewGuid();
            var workOrder = new msdyn_workorder()
            {
                Id = workOrderId,
                msdyn_ServiceRequest = new EntityReference(Incident.EntityLogicalName, incidentId),
                ovs_regulatedentity = new EntityReference(Account.EntityLogicalName, regulatedEntityId)
            };

            var workOrderServiceTaskId = Guid.NewGuid();
            var workOrderServiceTask = new msdyn_workorderservicetask()
            {
                Id = workOrderServiceTaskId,
                msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId), // belongs to a work order
                ovs_InspectionStatus = true,
                ovs_QuestionnaireResponse = @"
                {
                    ""finding-sq_162"": {
                        ""provisionReference"": ""Section 2"",
                        ""provisionText"": ""<strong>Application</strong></br><strong><mark>Section 2</mark></strong>: Sections 3 to 15 apply in respect of the following passenger-carrying flights — or in respect of air carriers conducting such flights — if the passengers, the property in the possession or control of the passengers and the belongings or baggage that the passengers give to the air carrier for transport are subject to screening that is carried out — in Canada under the Aeronautics Act or in another country by the person or entity responsible for the screening of such persons, property and belongings or baggage — before boarding:</br><ul style='list-style-type:none;'><li><strong>(a)</strong> domestic flights that depart from Canadian aerodromes and that are conducted by air carriers under Subpart 5 of Part VII of the Canadian Aviation Regulations ; and</li><li><strong>(b)</strong> international flights that depart from or will arrive at Canadian aerodromes and that are conducted by air carriers</li><ul style='list-style-type:none;'><li><strong>(i)</strong> under Subpart 1 of Part VII of the Canadian Aviation Regulations using aircraft that have a maximum certificated take-off weight of more than 8 618 kg (19,000 pounds) or have a seating configuration, excluding crew seats, of 20 or more, or</li><li><strong>(ii)</strong> under Subpart 5 of Part VII of the Canadian Aviation Regulations.</li></ul></ul>"",
                        ""comments"": ""test"",
                        ""documentaryEvidence"": ""C:\\fakepath\\Acc2.PNG""
                    }
                }
                "
            };

            var findingId = Guid.NewGuid();
            var finding = new ovs_Finding()
            {
                Id = findingId,
                ovs_Finding1 = workOrderServiceTaskId + "-finding-sq_162", // unique finding names are created using the work order service task ID and the reference ID in the questionnaire
                ovs_FindingProvisionReference = "Section 2",
                ovs_FindingComments = "original comments",
                ovs_FindingFile = "C:\\fakepath\\originalfile.png",
                ovs_CaseId = new EntityReference(Incident.EntityLogicalName, incidentId), // this finding already belongs to the case
                ovs_WorkOrderServiceTaskId = new EntityReference(msdyn_workorderservicetask.EntityLogicalName, workOrderServiceTaskId), // this finding already belongs to a work order service task
                StatusCode = ovs_Finding_StatusCode.Active, // finding is also already active
                StateCode = ovs_FindingState.Active, // finding is also already active
            };

            context.Initialize(
                new List<Entity>() {
                    regulatedEntity,
                    incident,
                    workOrder,
                    workOrderServiceTask,
                    finding
                }
            );

            ParameterCollection inputParams = new ParameterCollection();
            inputParams.Add("Target", workOrderServiceTask);
            ParameterCollection outputParams = new ParameterCollection();
            outputParams.Add("id", workOrderServiceTaskId);
            EntityImageCollection preEntityImages = new EntityImageCollection();
            preEntityImages.Add("PreImage", workOrderServiceTask);

            /**********
            * ACT
            **********/
            // Execute the PreOperationmsdyn_workorderservicetaskUpdate plugin with the defined workOrderServiceTask as a target
            context.ExecutePluginWith<PreOperationmsdyn_workorderservicetaskUpdate>(inputParams, outputParams, preEntityImages, null);

            /**********
             * ASSERT
             **********/
            var findings = context.CreateQuery<ovs_Finding>().ToList();

            // Expect target to still only contain 1 ovs_finding reference
            Assert.True(findings.Count == 1, "Expecting 1 finding");

            // Expect ovs_finding already in the context to still have the same provision reference
            var first = findings.First();
            Assert.Equal(finding.ovs_FindingProvisionReference, first.ovs_FindingProvisionReference);

            // Expect finding to still be in an active state
            Assert.Equal(ovs_FindingState.Active, first.StateCode);
            Assert.Equal(ovs_Finding_StatusCode.Active, first.StatusCode);

            // Expect work order service task result to be Fail
            Assert.Equal(msdyn_InspectionResult.Fail, workOrderServiceTask.msdyn_inspectiontaskresult);

        }

        [Fact]
        public void When_ovs_questionnaireresponse_contains_finding_that_already_exists_expect_existing_ovs_finding_record_to_be_updated()
        {
            /**********
            * ARRANGE
            **********/
            var context = new XrmFakedContext();

            // Given a work order service task that
            // - belongs to a work order
            // - belongs to a case (Incident)
            // - has a SurveyJS questionnaire response saved in the ovs_QuestionnaireResponse field
            var regulatedEntityId = Guid.NewGuid();
            var regulatedEntity = new Account()
            {
                Id = regulatedEntityId,
                Name = "Test Regulated Entity"
            };

            var incidentId = Guid.NewGuid();
            var incident = new Incident()
            {
                Id = incidentId
            };

            var workOrderId = Guid.NewGuid();
            var workOrder = new msdyn_workorder()
            {
                Id = workOrderId,
                msdyn_ServiceRequest = new EntityReference(Incident.EntityLogicalName, incidentId),
                ovs_regulatedentity = new EntityReference(Account.EntityLogicalName, regulatedEntityId)
            };

            var workOrderServiceTaskId = Guid.NewGuid();
            var workOrderServiceTask = new msdyn_workorderservicetask()
            {
                Id = workOrderServiceTaskId,
                ovs_InspectionStatus = true,
                msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId), // belongs to a work order
                ovs_QuestionnaireResponse = @"
                {
                    ""finding-sq_162"": {
                        ""provisionReference"": ""Section 2"",
                        ""provisionText"": ""<strong>Application</strong></br><strong><mark>Section 2</mark></strong>: Sections 3 to 15 apply in respect of the following passenger-carrying flights — or in respect of air carriers conducting such flights — if the passengers, the property in the possession or control of the passengers and the belongings or baggage that the passengers give to the air carrier for transport are subject to screening that is carried out — in Canada under the Aeronautics Act or in another country by the person or entity responsible for the screening of such persons, property and belongings or baggage — before boarding:</br><ul style='list-style-type:none;'><li><strong>(a)</strong> domestic flights that depart from Canadian aerodromes and that are conducted by air carriers under Subpart 5 of Part VII of the Canadian Aviation Regulations ; and</li><li><strong>(b)</strong> international flights that depart from or will arrive at Canadian aerodromes and that are conducted by air carriers</li><ul style='list-style-type:none;'><li><strong>(i)</strong> under Subpart 1 of Part VII of the Canadian Aviation Regulations using aircraft that have a maximum certificated take-off weight of more than 8 618 kg (19,000 pounds) or have a seating configuration, excluding crew seats, of 20 or more, or</li><li><strong>(ii)</strong> under Subpart 5 of Part VII of the Canadian Aviation Regulations.</li></ul></ul>"",
                        ""comments"": ""new comments"",
                        ""documentaryEvidence"": ""C:\\fakepath\\newfile.png""
                    }
                }
                "
            };

            var findingId = Guid.NewGuid();
            var finding = new ovs_Finding()
            {
                Id = findingId,
                ovs_Finding1 = workOrderServiceTaskId + "-finding-sq_162", // unique finding names are created using the work order service task ID and the reference ID in the questionnaire
                ovs_FindingProvisionReference = "Section 2",
                ovs_FindingComments = "original comments",
                ovs_FindingFile = "C:\\fakepath\\originalfile.png"
            };

            context.Initialize(
                new List<Entity>() {
                    regulatedEntity,
                    incident,
                    workOrder,
                    workOrderServiceTask,
                    finding
                }
            );

            ParameterCollection inputParams = new ParameterCollection();
            inputParams.Add("Target", workOrderServiceTask);
            ParameterCollection outputParams = new ParameterCollection();
            outputParams.Add("id", workOrderServiceTaskId);
            EntityImageCollection preEntityImages = new EntityImageCollection();
            preEntityImages.Add("PreImage", workOrderServiceTask);

            /**********
            * ACT
            **********/
            // Execute the PreOperationmsdyn_workorderservicetaskUpdate plugin with the defined workOrderServiceTask as a target
            context.ExecutePluginWith<PreOperationmsdyn_workorderservicetaskUpdate>(inputParams, outputParams, preEntityImages, null);

            /**********
             * ASSERT
             **********/
            var findings = context.CreateQuery<ovs_Finding>().ToList();

            // Expect target to still only contain 1 ovs_finding reference
            Assert.True(findings.Count == 1, "Expecting 1 finding");

            // Expect ovs_finding already in the context to now have new updated values
            var first = findings.First();
            Assert.Equal("new comments", first.ovs_FindingComments);

            // Expect work order service task result to be Fail
            Assert.Equal(msdyn_InspectionResult.Fail, workOrderServiceTask.msdyn_inspectiontaskresult);

        }

        [Fact]
        public void When_ovs_questionnaireresponse_contains_finding_that_already_exists_but_is_deactivated_expect_existing_ovs_finding_record_to_be_activated_in_case()
        {
            /**********
            * ARRANGE
            **********/
            var context = new XrmFakedContext();

            // Given a work order service task that
            // - belongs to a work order
            // - belongs to a case (Incident)
            // - has a SurveyJS questionnaire response saved in the ovs_QuestionnaireResponse field
            var regulatedEntityId = Guid.NewGuid();
            var regulatedEntity = new Account()
            {
                Id = regulatedEntityId,
                Name = "Test Regulated Entity"
            };

            var incidentId = Guid.NewGuid();
            var incident = new Incident()
            {
                Id = incidentId
            };

            var workOrderId = Guid.NewGuid();
            var workOrder = new msdyn_workorder()
            {
                Id = workOrderId,
                msdyn_ServiceRequest = new EntityReference(Incident.EntityLogicalName, incidentId),
                ovs_regulatedentity = new EntityReference(Account.EntityLogicalName, regulatedEntityId),
            };

            var workOrderServiceTaskId = Guid.NewGuid();
            var workOrderServiceTask = new msdyn_workorderservicetask()
            {
                Id = workOrderServiceTaskId,
                ovs_InspectionStatus = true,
                msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId), // belongs to a work order
                ovs_QuestionnaireResponse = @"
                {
                    ""finding-sq_162"": {
                        ""provisionReference"": ""Section 2"",
                        ""provisionText"": ""<strong>Application</strong></br><strong><mark>Section 2</mark></strong>: Sections 3 to 15 apply in respect of the following passenger-carrying flights — or in respect of air carriers conducting such flights — if the passengers, the property in the possession or control of the passengers and the belongings or baggage that the passengers give to the air carrier for transport are subject to screening that is carried out — in Canada under the Aeronautics Act or in another country by the person or entity responsible for the screening of such persons, property and belongings or baggage — before boarding:</br><ul style='list-style-type:none;'><li><strong>(a)</strong> domestic flights that depart from Canadian aerodromes and that are conducted by air carriers under Subpart 5 of Part VII of the Canadian Aviation Regulations ; and</li><li><strong>(b)</strong> international flights that depart from or will arrive at Canadian aerodromes and that are conducted by air carriers</li><ul style='list-style-type:none;'><li><strong>(i)</strong> under Subpart 1 of Part VII of the Canadian Aviation Regulations using aircraft that have a maximum certificated take-off weight of more than 8 618 kg (19,000 pounds) or have a seating configuration, excluding crew seats, of 20 or more, or</li><li><strong>(ii)</strong> under Subpart 5 of Part VII of the Canadian Aviation Regulations.</li></ul></ul>"",
                        ""comments"": ""new comments"",
                        ""documentaryEvidence"": ""C:\\fakepath\\newfile.png""
                    }
                }
                "
            };

            var findingId = Guid.NewGuid();
            var finding = new ovs_Finding()
            {
                Id = findingId,
                ovs_Finding1 = workOrderServiceTaskId + "-finding-sq_162", // unique finding names are created using the work order service task ID and the reference ID in the questionnaire
                ovs_FindingProvisionReference = "Section 2",
                ovs_FindingComments = "original comments",
                ovs_FindingFile = "C:\\fakepath\\originalfile.png",
                ovs_CaseId = new EntityReference(Incident.EntityLogicalName, incidentId), // this finding already belongs to the case
                ovs_WorkOrderServiceTaskId = new EntityReference(msdyn_workorderservicetask.EntityLogicalName, workOrderServiceTaskId), // this finding already belongs to a work order service task
                StatusCode = ovs_Finding_StatusCode.Inactive, // finding is also already deactivated
                StateCode = ovs_FindingState.Inactive, // finding is also already deactivated
            };



            context.Initialize(
                new List<Entity>() {
                    regulatedEntity,
                    incident,
                    workOrder,
                    workOrderServiceTask,
                    finding
                }
            );

            ParameterCollection inputParams = new ParameterCollection();
            inputParams.Add("Target", workOrderServiceTask);
            ParameterCollection outputParams = new ParameterCollection();
            outputParams.Add("id", workOrderServiceTaskId);
            EntityImageCollection preEntityImages = new EntityImageCollection();
            preEntityImages.Add("PreImage", workOrderServiceTask);

            /**********
            * ACT
            **********/
            // Execute the PreOperationmsdyn_workorderservicetaskUpdate plugin with the defined workOrderServiceTask as a target
            context.ExecutePluginWith<PreOperationmsdyn_workorderservicetaskUpdate>(inputParams, outputParams, preEntityImages, null);

            /**********
             * ASSERT
             **********/
            var findings = context.CreateQuery<ovs_Finding>().ToList();

            // Expect target to still only contain 1 ovs_finding reference
            Assert.True(findings.Count == 1, "Expecting 1 finding");

            // However, expect ovs_finding already in the context to now be activated
            var first = findings.First();
            Assert.Equal(ovs_Finding_StatusCode.Active, first.StatusCode);
            Assert.Equal(ovs_FindingState.Active, first.StateCode);

            // Expect work order service task result to be Fail
            Assert.Equal(msdyn_InspectionResult.Fail, workOrderServiceTask.msdyn_inspectiontaskresult);

        }

        [Fact]
        public void When_ovs_questionnaireresponse_no_longer_contains_finding_that_already_exists_expect_existing_ovs_finding_record_to_be_deactived_in_case()
        {
            /**********
            * ARRANGE
            **********/
            var context = new XrmFakedContext();

            // Given a work order service task that
            // - belongs to a work order
            // - belongs to a case (Incident)
            // - has a SurveyJS questionnaire response saved in the ovs_QuestionnaireResponse field
            var regulatedEntityId = Guid.NewGuid();
            var regulatedEntity = new Account()
            {
                Id = regulatedEntityId,
                Name = "Test Regulated Entity"
            };

            var incidentId = Guid.NewGuid();
            var incident = new Incident()
            {
                Id = incidentId
            };

            var workOrderId = Guid.NewGuid();
            var workOrder = new msdyn_workorder()
            {
                Id = workOrderId,
                msdyn_ServiceRequest = new EntityReference(Incident.EntityLogicalName, incidentId),
                ovs_regulatedentity = new EntityReference(Account.EntityLogicalName, regulatedEntityId),
            };

            var workOrderServiceTaskId = Guid.NewGuid();
            var workOrderServiceTask = new msdyn_workorderservicetask()
            {
                Id = workOrderServiceTaskId,
                ovs_InspectionStatus = true,
                msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId), // belongs to a work order
                ovs_QuestionnaireResponse = @"
                {
                }
                "
            };

            var findingId = Guid.NewGuid();
            var finding = new ovs_Finding()
            {
                Id = findingId,
                ovs_Finding1 = workOrderServiceTaskId + "-finding-sq_162", // unique finding names are created using the work order service task ID and the reference ID in the questionnaire
                ovs_FindingProvisionReference = "Section 2",
                ovs_FindingComments = "original comments",
                ovs_FindingFile = "C:\\fakepath\\originalfile.png",
                ovs_CaseId = new EntityReference(Incident.EntityLogicalName, incidentId), // this finding already belongs to the case
                ovs_WorkOrderServiceTaskId = new EntityReference(msdyn_workorderservicetask.EntityLogicalName, workOrderServiceTaskId), // this finding already belongs to a work order service task
                StatusCode = ovs_Finding_StatusCode.Active, // finding is also already active
                StateCode = ovs_FindingState.Active, // finding is also already active
            };



            context.Initialize(
                new List<Entity>() {
                    regulatedEntity,
                    incident,
                    workOrder,
                    workOrderServiceTask,
                    finding
                }
            );

            ParameterCollection inputParams = new ParameterCollection();
            inputParams.Add("Target", workOrderServiceTask);
            ParameterCollection outputParams = new ParameterCollection();
            outputParams.Add("id", workOrderServiceTaskId);
            EntityImageCollection preEntityImages = new EntityImageCollection();
            preEntityImages.Add("PreImage", workOrderServiceTask);

            /**********
            * ACT
            **********/
            // Execute the PreOperationmsdyn_workorderservicetaskUpdate plugin with the defined workOrderServiceTask as a target
            context.ExecutePluginWith<PreOperationmsdyn_workorderservicetaskUpdate>(inputParams, outputParams, preEntityImages, null);

            /**********
             * ASSERT
             **********/
            var findings = context.CreateQuery<ovs_Finding>().ToList();

            // Expect target to still only contain 1 ovs_finding reference
            Assert.True(findings.Count == 1, "Expecting 1 finding");

            // However, expect ovs_finding already in the context to now be deactivated
            var first = findings.First();
            Assert.Equal(ovs_Finding_StatusCode.Inactive, first.StatusCode);
            Assert.Equal(ovs_FindingState.Inactive, first.StateCode);

            // Expect work order service task result to be Pass
            Assert.Equal(msdyn_InspectionResult.Pass, workOrderServiceTask.msdyn_inspectiontaskresult);

        }

        [Fact] 
        public void When_ovs_questionnaireresponse_contains_finding_with_missing_optional_values_expect_ovs_finding_record_to_still_be_created()
        {
            /**********
            * ARRANGE
            **********/
            var context = new XrmFakedContext();

            // Given a work order service task that
            // - belongs to a work order
            // - belongs to a case (Incident)
            // - has a SurveyJS questionnaire response saved in the ovs_QuestionnaireResponse field that does not have the comments or file provided
            var regulatedEntityId = Guid.NewGuid();
            var regulatedEntity = new Account()
            {
                Id = regulatedEntityId,
                Name = "Test Regulated Entity"
            };

            var incidentId = Guid.NewGuid();
            var incident = new Incident()
            {
                Id = incidentId
            };

            var workOrderId = Guid.NewGuid();
            var workOrder = new msdyn_workorder()
            {
                Id = workOrderId,
                msdyn_ServiceRequest = new EntityReference(Incident.EntityLogicalName, incidentId),
                ovs_regulatedentity = new EntityReference(Account.EntityLogicalName, regulatedEntityId)
            };

            var workOrderServiceTaskId = Guid.NewGuid();
            var workOrderServiceTask = new msdyn_workorderservicetask()
            {
                Id = workOrderServiceTaskId,
                msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId), // belongs to a work order
                ovs_InspectionStatus = true,
                ovs_QuestionnaireResponse = @"
                {
                    ""finding-sq_162"": {
                        ""provisionReference"": ""Section 2"",
                        ""provisionText"": ""<strong>Application</strong></br><strong><mark>Section 2</mark></strong>: Sections 3 to 15 apply in respect of the following passenger-carrying flights — or in respect of air carriers conducting such flights — if the passengers, the property in the possession or control of the passengers and the belongings or baggage that the passengers give to the air carrier for transport are subject to screening that is carried out — in Canada under the Aeronautics Act or in another country by the person or entity responsible for the screening of such persons, property and belongings or baggage — before boarding:</br><ul style='list-style-type:none;'><li><strong>(a)</strong> domestic flights that depart from Canadian aerodromes and that are conducted by air carriers under Subpart 5 of Part VII of the Canadian Aviation Regulations ; and</li><li><strong>(b)</strong> international flights that depart from or will arrive at Canadian aerodromes and that are conducted by air carriers</li><ul style='list-style-type:none;'><li><strong>(i)</strong> under Subpart 1 of Part VII of the Canadian Aviation Regulations using aircraft that have a maximum certificated take-off weight of more than 8 618 kg (19,000 pounds) or have a seating configuration, excluding crew seats, of 20 or more, or</li><li><strong>(ii)</strong> under Subpart 5 of Part VII of the Canadian Aviation Regulations.</li></ul></ul>""
                    }
                }
                "
            };

            context.Initialize(
                new List<Entity>() {
                    regulatedEntity,
                    incident,
                    workOrder,
                    workOrderServiceTask
                }
            );

            ParameterCollection inputParams = new ParameterCollection();
            inputParams.Add("Target", workOrderServiceTask);
            ParameterCollection outputParams = new ParameterCollection();
            outputParams.Add("id", workOrderServiceTaskId);
            EntityImageCollection preEntityImages = new EntityImageCollection();
            preEntityImages.Add("PreImage", workOrderServiceTask);

            /**********
            * ACT
            **********/
            // Execute the PreOperationmsdyn_workorderservicetaskUpdate plugin with the defined workOrderServiceTask as a target
            context.ExecutePluginWith<PreOperationmsdyn_workorderservicetaskUpdate>(inputParams, outputParams, preEntityImages, null);

            /**********
             * ASSERT
             **********/
            var findings = context.CreateQuery<ovs_Finding>().ToList();

            // Expect target to still only contain 1 ovs_finding reference
            Assert.True(findings.Count == 1, "Expecting 1 finding");

            // Expect ovs_finding already in the context to still have the same provision reference
            var first = findings.First();
            Assert.Equal("Section 2", first.ovs_FindingProvisionReference);

            // Expect work order service task result to be Fail
            Assert.Equal(msdyn_InspectionResult.Fail, workOrderServiceTask.msdyn_inspectiontaskresult);

        }

        [Fact]
        public void When_ovs_questionnaireresponse_contains_finding_expect_msdyn_inspectiontaskresult_fail()
        {
            /**********
             * ARRANGE
             **********/
            var context = new XrmFakedContext();

            // Given a work order service task that
            // - inspection task rseult is no already in a failed state
            // - belongs to a work order
            // - does not already belongs to a case (Incident)
            // - has a SurveyJS questionnaire response saved in the ovs_QuestionnaireResponse field
            var regulatedEntityId = Guid.NewGuid();
            var regulatedEntity = new Account()
            {
                Id = regulatedEntityId,
                Name = "Test Regulated Entity"
            };

            var workOrderId = Guid.NewGuid();
            var workOrder = new msdyn_workorder()
            {
                Id = workOrderId,
                msdyn_ServiceRequest = null, // does not already belong to a case (Incident)
                ovs_regulatedentity = new EntityReference(Account.EntityLogicalName, regulatedEntityId)
            };

            var workOrderServiceTaskId = Guid.NewGuid();
            var workOrderServiceTask = new msdyn_workorderservicetask()
            {
                Id = workOrderServiceTaskId,
                msdyn_inspectiontaskresult = msdyn_InspectionResult.NA, 
                msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId), // belongs to a work order
                ovs_InspectionStatus = true,
                ovs_QuestionnaireResponse = @"
                {
                    ""finding-sq_162"": {
                        ""provisionReference"": ""Section 2"",
                        ""provisionText"": ""<strong>Application</strong></br><strong><mark>Section 2</mark></strong>: Sections 3 to 15 apply in respect of the following passenger-carrying flights — or in respect of air carriers conducting such flights — if the passengers, the property in the possession or control of the passengers and the belongings or baggage that the passengers give to the air carrier for transport are subject to screening that is carried out — in Canada under the Aeronautics Act or in another country by the person or entity responsible for the screening of such persons, property and belongings or baggage — before boarding:</br><ul style='list-style-type:none;'><li><strong>(a)</strong> domestic flights that depart from Canadian aerodromes and that are conducted by air carriers under Subpart 5 of Part VII of the Canadian Aviation Regulations ; and</li><li><strong>(b)</strong> international flights that depart from or will arrive at Canadian aerodromes and that are conducted by air carriers</li><ul style='list-style-type:none;'><li><strong>(i)</strong> under Subpart 1 of Part VII of the Canadian Aviation Regulations using aircraft that have a maximum certificated take-off weight of more than 8 618 kg (19,000 pounds) or have a seating configuration, excluding crew seats, of 20 or more, or</li><li><strong>(ii)</strong> under Subpart 5 of Part VII of the Canadian Aviation Regulations.</li></ul></ul>"",
                        ""comments"": ""new comments"",
                        ""documentaryEvidence"": ""C:\\fakepath\\newfile.png""
                    }
                }
                "
            };

            context.Initialize(
                new List<Entity>() {
                    regulatedEntity,
                    workOrder,
                    workOrderServiceTask
                }
            );

            ParameterCollection inputParams = new ParameterCollection();
            inputParams.Add("Target", workOrderServiceTask);
            ParameterCollection outputParams = new ParameterCollection();
            outputParams.Add("id", workOrderServiceTaskId);
            EntityImageCollection preEntityImages = new EntityImageCollection();
            preEntityImages.Add("PreImage", workOrderServiceTask);

            /**********
            * ACT
            **********/
            // Execute the PreOperationmsdyn_workorderservicetaskUpdate plugin with the defined workOrderServiceTask as a target
            context.ExecutePluginWith<PreOperationmsdyn_workorderservicetaskUpdate>(inputParams, outputParams, preEntityImages, null);

            /**********
             * ASSERT
             **********/
            Assert.True(workOrderServiceTask.msdyn_inspectiontaskresult.Equals(msdyn_InspectionResult.Fail), "Expected inspection result to be fail");
        }

        [Fact]
        public void When_ovs_questionnaireresponse_contains_finding_expect_create_incident_and_ovs_finding_if_they_do_not_exist()
        {
            /**********
             * ARRANGE
             **********/
            var context = new XrmFakedContext();

            // Given a work order service task that
            // - belongs to a work order
            // - does not already belongs to a case (Incident)
            // - has a SurveyJS questionnaire response saved in the ovs_QuestionnaireResponse field
            var regulatedEntityId = Guid.NewGuid();
            var regulatedEntity = new Account()
            {
                Id = regulatedEntityId,
                Name = "Test Regulated Entity"
            };

            var workOrderId = Guid.NewGuid();
            var workOrder = new msdyn_workorder()
            {
                Id = workOrderId,
                msdyn_ServiceRequest = null, // does not already belong to a case (Incident)
                ovs_regulatedentity = new EntityReference(Account.EntityLogicalName, regulatedEntityId)
            };

            var workOrderServiceTaskId = Guid.NewGuid();
            var workOrderServiceTask = new msdyn_workorderservicetask()
            {
                Id = workOrderServiceTaskId,
                msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId), // belongs to a work order
                ovs_InspectionStatus = true,
                ovs_QuestionnaireResponse = @"
                {
                    ""finding-sq_162"": {
                        ""provisionReference"": ""Section 2"",
                        ""provisionText"": ""<strong>Application</strong></br><strong><mark>Section 2</mark></strong>: Sections 3 to 15 apply in respect of the following passenger-carrying flights — or in respect of air carriers conducting such flights — if the passengers, the property in the possession or control of the passengers and the belongings or baggage that the passengers give to the air carrier for transport are subject to screening that is carried out — in Canada under the Aeronautics Act or in another country by the person or entity responsible for the screening of such persons, property and belongings or baggage — before boarding:</br><ul style='list-style-type:none;'><li><strong>(a)</strong> domestic flights that depart from Canadian aerodromes and that are conducted by air carriers under Subpart 5 of Part VII of the Canadian Aviation Regulations ; and</li><li><strong>(b)</strong> international flights that depart from or will arrive at Canadian aerodromes and that are conducted by air carriers</li><ul style='list-style-type:none;'><li><strong>(i)</strong> under Subpart 1 of Part VII of the Canadian Aviation Regulations using aircraft that have a maximum certificated take-off weight of more than 8 618 kg (19,000 pounds) or have a seating configuration, excluding crew seats, of 20 or more, or</li><li><strong>(ii)</strong> under Subpart 5 of Part VII of the Canadian Aviation Regulations.</li></ul></ul>"",
                        ""comments"": ""new comments"",
                        ""documentaryEvidence"": ""C:\\fakepath\\newfile.png""
                    }
                }
                "
            };

            context.Initialize(
                new List<Entity>() {
                    regulatedEntity,
                    workOrder,
                    workOrderServiceTask
                }
            );

            ParameterCollection inputParams = new ParameterCollection();
            inputParams.Add("Target", workOrderServiceTask);
            ParameterCollection outputParams = new ParameterCollection();
            outputParams.Add("id", workOrderServiceTaskId);
            EntityImageCollection preEntityImages = new EntityImageCollection();
            preEntityImages.Add("PreImage", workOrderServiceTask);

            /**********
            * ACT
            **********/
            // Execute the PreOperationmsdyn_workorderservicetaskUpdate plugin with the defined workOrderServiceTask as a target
            context.ExecutePluginWith<PreOperationmsdyn_workorderservicetaskUpdate>(inputParams, outputParams, preEntityImages, null);

            /**********
             * ASSERT
             **********/
            var incidents = context.CreateQuery<Incident>().ToList();
            var findings = context.CreateQuery<ovs_Finding>().ToList();

            // Expect there is only one incident created in the context
            var oneIncident = incidents.First();
            Assert.True(incidents.Count == 1, "Expected 1 incident");

            // Expect there is only one finding created in the context
            var oneFinding = findings.First();
            Assert.True(findings.Count == 1, "Expected 1 finding");

            // Expect ovs_finding name to be combination of Work Order Number and Finding Name
            Assert.True(oneFinding.ovs_Finding1 == workOrderServiceTask.Id.ToString() + "-finding-sq_162");

            // Expect ovs_finding to contain same content as what is in JSON response
            Assert.Equal("Section 2", oneFinding.ovs_FindingProvisionReference);
            Assert.Equal("<strong>Application</strong></br><strong><mark>Section 2</mark></strong>: Sections 3 to 15 apply in respect of the following passenger-carrying flights — or in respect of air carriers conducting such flights — if the passengers, the property in the possession or control of the passengers and the belongings or baggage that the passengers give to the air carrier for transport are subject to screening that is carried out — in Canada under the Aeronautics Act or in another country by the person or entity responsible for the screening of such persons, property and belongings or baggage — before boarding:</br><ul style='list-style-type:none;'><li><strong>(a)</strong> domestic flights that depart from Canadian aerodromes and that are conducted by air carriers under Subpart 5 of Part VII of the Canadian Aviation Regulations ; and</li><li><strong>(b)</strong> international flights that depart from or will arrive at Canadian aerodromes and that are conducted by air carriers</li><ul style='list-style-type:none;'><li><strong>(i)</strong> under Subpart 1 of Part VII of the Canadian Aviation Regulations using aircraft that have a maximum certificated take-off weight of more than 8 618 kg (19,000 pounds) or have a seating configuration, excluding crew seats, of 20 or more, or</li><li><strong>(ii)</strong> under Subpart 5 of Part VII of the Canadian Aviation Regulations.</li></ul></ul>", oneFinding.ovs_FindingProvisionText);
            Assert.Equal("new comments", oneFinding.ovs_FindingComments);

            // Expect ovs_finding to reference the proper entities
            Assert.Equal(workOrderServiceTaskId, oneFinding.ovs_WorkOrderServiceTaskId.Id);
            Assert.Equal(oneIncident.Id, oneFinding.ovs_CaseId.Id);

            // Expect work order service task result to be Fail
            Assert.Equal(msdyn_InspectionResult.Fail, workOrderServiceTask.msdyn_inspectiontaskresult);
        }

        [Fact]
        public void When_ovs_questionnaireresponse_does_not_contain_finding_expect_no_ovs_finding_record_created()
        {
            /**********
             * ARRANGE
             **********/
            var context = new XrmFakedContext();

            // Given a work order service task that
            // - belongs to a work order
            // - belongs to a case (Incident)
            // - ovs_questionnaireresponse is empty
            var regulatedEntityId = Guid.NewGuid();
            var regulatedEntity = new Account()
            {
                Id = regulatedEntityId,
                Name = "Test Regulated Entity"
            };

            var incidentId = Guid.NewGuid();
            var incident = new Incident()
            {
                Id = incidentId
            };

            var workOrderId = Guid.NewGuid();
            var workOrder = new msdyn_workorder()
            {
                Id = workOrderId,
                msdyn_ServiceRequest = new EntityReference(Incident.EntityLogicalName, incidentId),
                ovs_regulatedentity = new EntityReference(Account.EntityLogicalName, regulatedEntityId)
            };

            var workOrderServiceTaskId = Guid.NewGuid();
            var workOrderServiceTask = new msdyn_workorderservicetask()
            {
                Id = workOrderServiceTaskId,
                msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId), // belongs to a work order
                ovs_InspectionStatus = true,
                ovs_QuestionnaireResponse = ""
            };

            context.Initialize(
                new List<Entity>() {
                    regulatedEntity,
                    incident,
                    workOrder,
                    workOrderServiceTask
                }
            );

            ParameterCollection inputParams = new ParameterCollection();
            inputParams.Add("Target", workOrderServiceTask);
            ParameterCollection outputParams = new ParameterCollection();
            outputParams.Add("id", workOrderServiceTaskId);
            EntityImageCollection preEntityImages = new EntityImageCollection();
            preEntityImages.Add("PreImage", workOrderServiceTask);

            /**********
            * ACT
            **********/
            // Execute the PreOperationmsdyn_workorderservicetaskUpdate plugin with the defined workOrderServiceTask as a target
            context.ExecutePluginWith<PreOperationmsdyn_workorderservicetaskUpdate>(inputParams, outputParams, preEntityImages, null);

            /**********
             * ASSERT
             **********/
            var findings = context.CreateQuery<ovs_Finding>().ToList();

            // Expect target to now contain 1 ovs_finding reference
            Assert.True(findings.Count == 0, "Expected no findings");
        }

    }
}
