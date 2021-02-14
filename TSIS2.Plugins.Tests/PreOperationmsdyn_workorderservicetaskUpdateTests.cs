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
            var billingAccountId = Guid.NewGuid();
            var billingAccount = new Account()
            {
                Id = billingAccountId,
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
                msdyn_BillingAccount = new EntityReference(Account.EntityLogicalName, billingAccountId)
            };

            var workOrderServiceTaskId = Guid.NewGuid();
            var workOrderServiceTask = new msdyn_workorderservicetask()
            {
                Id = workOrderServiceTaskId,
                msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId), // belongs to a work order
                ovs_QuestionnaireReponse = @"
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
                ovs_FindingProvisionReference = "Section 2"
            };

            context.Initialize(
                new List<Entity>() {
                    billingAccount,
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
        }

        [Fact]
        public void When_ovs_questionnaireresponse_contains_finding_that_already_exists_expect_existing_ovs_finding_record_to_be_updated()
        {
            throw new NotImplementedException();
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
            var billingAccountId = Guid.NewGuid();
            var billingAccount = new Account()
            {
                Id = billingAccountId,
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
                msdyn_BillingAccount = new EntityReference(Account.EntityLogicalName, billingAccountId)
            };

            var workOrderServiceTaskId = Guid.NewGuid();
            var workOrderServiceTask = new msdyn_workorderservicetask()
            {
                Id = workOrderServiceTaskId,
                msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId), // belongs to a work order
                ovs_QuestionnaireReponse = @"
                {
                    ""finding - Reg15"": {
                        ""provisionReference"": ""1.3 (2) (d) (i)"",
                        ""provisionText"": ""**1.3 (2) (d)** shipping names listed in Schedule 1 may be<br/>&#160;&#160;&#160;&#160;**(i)** written in the singular or plural,<br/>&#160;&#160;&#160;&#160;**ii)** written in upper or lower case letters, except that when the shipping name is followed by the descriptive text associated with the shipping name the descriptive text must be in lower case letters and the shipping name must be in upper case letters (capitals),<br/>&#160;&#160;&#160;&#160;**ii)** in English only, put in a different word order as long as the full shipping name is used and the word order is a commonly used one,<br/>&#160;&#160;&#160;&#160;**iv)** for solutions and mixtures, followed by the word “SOLUTION” or “MIXTURE”, as appropriate, and may include the concentration of the solution or mixture, and<br/>&#160;&#160;&#160;&#160;**(v)** for waste, preceded or followed by the word “WASTE” or “DÉCHET”;<br/>""
                    }
                }
                "
            };

            context.Initialize(
                new List<Entity>() {
                    billingAccount,
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
            Assert.Equal("1.3 (2) (d) (i)", first.ovs_FindingProvisionReference);
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
            var billingAccountId = Guid.NewGuid();
            var billingAccount = new Account()
            {
                Id = billingAccountId,
                Name = "Test Regulated Entity"
            };

            var workOrderId = Guid.NewGuid();
            var workOrder = new msdyn_workorder()
            {
                Id = workOrderId,
                msdyn_ServiceRequest = null, // does not already belong to a case (Incident)
                msdyn_BillingAccount = new EntityReference(Account.EntityLogicalName, billingAccountId)
            };

            var workOrderServiceTaskId = Guid.NewGuid();
            var workOrderServiceTask = new msdyn_workorderservicetask()
            {
                Id = workOrderServiceTaskId,
                msdyn_inspectiontaskresult = msdyn_InspectionResult.NA, 
                msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId), // belongs to a work order
                ovs_QuestionnaireReponse = @"
                {
                    ""finding - Reg15"": {
                        ""provisionReference"": ""1.3 (2) (d) (i)"",
                        ""provisionText"": ""**1.3 (2) (d)** shipping names listed in Schedule 1 may be<br/>&#160;&#160;&#160;&#160;**(i)** written in the singular or plural,<br/>&#160;&#160;&#160;&#160;**ii)** written in upper or lower case letters, except that when the shipping name is followed by the descriptive text associated with the shipping name the descriptive text must be in lower case letters and the shipping name must be in upper case letters (capitals),<br/>&#160;&#160;&#160;&#160;**ii)** in English only, put in a different word order as long as the full shipping name is used and the word order is a commonly used one,<br/>&#160;&#160;&#160;&#160;**iv)** for solutions and mixtures, followed by the word “SOLUTION” or “MIXTURE”, as appropriate, and may include the concentration of the solution or mixture, and<br/>&#160;&#160;&#160;&#160;**(v)** for waste, preceded or followed by the word “WASTE” or “DÉCHET”;<br/>"",
                        ""comments"": ""blah 1"",
                        ""documentaryEvidence"": ""C:\\fakepath\\Untitled.png""
                    }
                }
                "
            };

            context.Initialize(
                new List<Entity>() {
                    billingAccount,
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
            var billingAccountId = Guid.NewGuid();
            var billingAccount = new Account()
            {
                Id = billingAccountId,
                Name = "Test Regulated Entity"
            };

            var workOrderId = Guid.NewGuid();
            var workOrder = new msdyn_workorder()
            {
                Id = workOrderId,
                msdyn_ServiceRequest = null, // does not already belong to a case (Incident)
                msdyn_BillingAccount = new EntityReference(Account.EntityLogicalName, billingAccountId)
            };

            var workOrderServiceTaskId = Guid.NewGuid();
            var workOrderServiceTask = new msdyn_workorderservicetask()
            {
                Id = workOrderServiceTaskId,
                msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId), // belongs to a work order
                ovs_QuestionnaireReponse = @"
                {
                    ""finding - Reg15"": {
                        ""provisionReference"": ""1.3 (2) (d) (i)"",
                        ""provisionText"": ""**1.3 (2) (d)** shipping names listed in Schedule 1 may be<br/>&#160;&#160;&#160;&#160;**(i)** written in the singular or plural,<br/>&#160;&#160;&#160;&#160;**ii)** written in upper or lower case letters, except that when the shipping name is followed by the descriptive text associated with the shipping name the descriptive text must be in lower case letters and the shipping name must be in upper case letters (capitals),<br/>&#160;&#160;&#160;&#160;**ii)** in English only, put in a different word order as long as the full shipping name is used and the word order is a commonly used one,<br/>&#160;&#160;&#160;&#160;**iv)** for solutions and mixtures, followed by the word “SOLUTION” or “MIXTURE”, as appropriate, and may include the concentration of the solution or mixture, and<br/>&#160;&#160;&#160;&#160;**(v)** for waste, preceded or followed by the word “WASTE” or “DÉCHET”;<br/>"",
                        ""comments"": ""blah 1"",
                        ""documentaryEvidence"": ""C:\\fakepath\\Untitled.png""
                    }
                }
                "
            };

            context.Initialize(
                new List<Entity>() {
                    billingAccount,
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
            throw new NotImplementedException();

            // Expect ovs_finding to contain same content as what is in JSON response
            Assert.Equal("1.3 (2) (d) (i)", oneFinding.ovs_FindingProvisionReference);
            Assert.Equal("**1.3 (2) (d)** shipping names listed in Schedule 1 may be<br/>&#160;&#160;&#160;&#160;**(i)** written in the singular or plural,<br/>&#160;&#160;&#160;&#160;**ii)** written in upper or lower case letters, except that when the shipping name is followed by the descriptive text associated with the shipping name the descriptive text must be in lower case letters and the shipping name must be in upper case letters (capitals),<br/>&#160;&#160;&#160;&#160;**ii)** in English only, put in a different word order as long as the full shipping name is used and the word order is a commonly used one,<br/>&#160;&#160;&#160;&#160;**iv)** for solutions and mixtures, followed by the word “SOLUTION” or “MIXTURE”, as appropriate, and may include the concentration of the solution or mixture, and<br/>&#160;&#160;&#160;&#160;**(v)** for waste, preceded or followed by the word “WASTE” or “DÉCHET”;<br/>", oneFinding.ovs_FindingProvisionText);
            Assert.Equal("blah 1", oneFinding.ovs_FindingComments);
            Assert.Equal("C:\\fakepath\\Untitled.png", oneFinding.ovs_FindingFile);

            // Expect ovs_finding to reference the proper entities
            Assert.Equal(workOrderServiceTaskId, oneFinding.ovs_WorkOrderServiceTaskId.Id);
            Assert.Equal(oneIncident.Id, oneFinding.ovs_CaseId.Id);
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
            var billingAccountId = Guid.NewGuid();
            var billingAccount = new Account()
            {
                Id = billingAccountId,
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
                msdyn_BillingAccount = new EntityReference(Account.EntityLogicalName, billingAccountId)
            };

            var workOrderServiceTaskId = Guid.NewGuid();
            var workOrderServiceTask = new msdyn_workorderservicetask()
            {
                Id = workOrderServiceTaskId,
                msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId), // belongs to a work order
                ovs_QuestionnaireReponse = ""
            };

            context.Initialize(
                new List<Entity>() {
                    billingAccount,
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
