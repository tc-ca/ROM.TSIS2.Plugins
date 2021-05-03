using System;
using System.Collections.Generic;
using System.Linq;
using Xunit;
using FakeXrmEasy;
using Microsoft.Xrm.Sdk;
using TSIS2.Common;

namespace TSIS2.Plugins.Tests
{
 
    public class PreOperationmsdyn_workorderservicetaskUpdateTests
    {
        [Fact]
        public void When_work_order_service_task_already_has_finding_expect_next_findings_to_have_name_with_incremented_suffix()
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
                msdyn_name = "200-345678-1",
                msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId), // belongs to a work order
                msdyn_PercentComplete = 100.00,
                ovs_QuestionnaireResponse = @"
                {
                    ""finding-sq_142"": {
                        ""provisionReference"": ""SATR 4"",
                        ""provisionTextEn"": ""<strong>Verification of Identity</strong></br><strong><mark><mark>SATR 4</mark></mark></strong>: An air carrier must, at the boarding gate for an international flight, verify the identity of each passenger who appears to be 18 years of age or older using</br><ul style='list-style-type:none;'><li><strong>(a)</strong> one of the following pieces of photo identification issued by a government authority that shows the passenger’s surname, first name and any middle names, their date of birth and gender and that is valid:</li><ul style='list-style-type:none;'><li><strong>(i)</strong> a passport issued by the country of which the passenger is a citizen or a national,</li><li><strong>(ii)</strong> a NEXUS card,</li><li><strong>(iii)</strong> any document referred to in subsection 50(1) or 52(1) of the Immigration and Refugee Protection Regulations; or</li></ul><li><strong>(b)</strong> a valid restricted area identity card, as defined in section 3 of the Canadian Aviation Security Regulations, 2012.</li></ul>"",
                        ""provisionTextFr"": ""<strong>Verification of Identity</strong></br><strong><mark><mark><mark>SATR 4</mark></mark></mark></strong>: Tout transporteur aérien vérifie, à la porte d’embarquement pour un vol international, l’identité de chaque passager qui semble âgé de 18 ans ou plus au moyen :</br><ul style='list-style-type:none;'><li><strong>(a)</strong> soit de l’une des pièces d’identité avec photo ci-après qui est délivrée par une autorité gouvernementale, qui indique les nom et prénoms, date de naissance et genre du passager et qui est valide :</li><ul style='list-style-type:none;'><li><strong>(i)</strong> un passeport délivré au passager par le pays dont il est citoyen ou ressortissant,</li><li><strong>(ii)</strong> une carte NEXUS,</li><li><strong>(iii)</strong> un document visé au paragraphe 50(1) ou 52(1) du Règlement sur l’immigration et la protection des réfugiés;</li></ul><li><strong>(b)</strong> soit d’une carte d’identité de zone réglementée au sens de l’article 3 du Règlement canadien de 2012 sur la sûreté aérienne qui est valide.</li></ul>"",
                        ""comments"": ""existing comments"",
                        ""documentaryEvidence"": ""C:\\fakepath\\exitingfile.png""
                    },
                    ""finding-sq_155"": {
                        ""provisionReference"": ""SATR 5"",
                        ""provisionTextEn"": ""<strong>Verification of Identity</strong></br><strong><mark><mark>SATR 4</mark></mark></strong>: An air carrier must, at the boarding gate for an international flight, verify the identity of each passenger who appears to be 18 years of age or older using</br><ul style='list-style-type:none;'><li><strong>(a)</strong> one of the following pieces of photo identification issued by a government authority that shows the passenger’s surname, first name and any middle names, their date of birth and gender and that is valid:</li><ul style='list-style-type:none;'><li><strong>(i)</strong> a passport issued by the country of which the passenger is a citizen or a national,</li><li><strong>(ii)</strong> a NEXUS card,</li><li><strong>(iii)</strong> any document referred to in subsection 50(1) or 52(1) of the Immigration and Refugee Protection Regulations; or</li></ul><li><strong>(b)</strong> a valid restricted area identity card, as defined in section 3 of the Canadian Aviation Security Regulations, 2012.</li></ul>"",
                        ""provisionTextFr"": ""<strong>Verification of Identity</strong></br><strong><mark><mark><mark>SATR 4</mark></mark></mark></strong>: Tout transporteur aérien vérifie, à la porte d’embarquement pour un vol international, l’identité de chaque passager qui semble âgé de 18 ans ou plus au moyen :</br><ul style='list-style-type:none;'><li><strong>(a)</strong> soit de l’une des pièces d’identité avec photo ci-après qui est délivrée par une autorité gouvernementale, qui indique les nom et prénoms, date de naissance et genre du passager et qui est valide :</li><ul style='list-style-type:none;'><li><strong>(i)</strong> un passeport délivré au passager par le pays dont il est citoyen ou ressortissant,</li><li><strong>(ii)</strong> une carte NEXUS,</li><li><strong>(iii)</strong> un document visé au paragraphe 50(1) ou 52(1) du Règlement sur l’immigration et la protection des réfugiés;</li></ul><li><strong>(b)</strong> soit d’une carte d’identité de zone réglementée au sens de l’article 3 du Règlement canadien de 2012 sur la sûreté aérienne qui est valide.</li></ul>"",
                        ""comments"": ""some new comment"",
                        ""documentaryEvidence"": ""C:\\fakepath\\somenewfile.png""
                    },
                    ""finding-sq_166"": {
                        ""provisionReference"": ""SATR 6"",
                        ""provisionTextEn"": ""<strong>Verification of Identity</strong></br><strong><mark><mark>SATR 4</mark></mark></strong>: An air carrier must, at the boarding gate for an international flight, verify the identity of each passenger who appears to be 18 years of age or older using</br><ul style='list-style-type:none;'><li><strong>(a)</strong> one of the following pieces of photo identification issued by a government authority that shows the passenger’s surname, first name and any middle names, their date of birth and gender and that is valid:</li><ul style='list-style-type:none;'><li><strong>(i)</strong> a passport issued by the country of which the passenger is a citizen or a national,</li><li><strong>(ii)</strong> a NEXUS card,</li><li><strong>(iii)</strong> any document referred to in subsection 50(1) or 52(1) of the Immigration and Refugee Protection Regulations; or</li></ul><li><strong>(b)</strong> a valid restricted area identity card, as defined in section 3 of the Canadian Aviation Security Regulations, 2012.</li></ul>"",
                        ""provisionTextFr"": ""<strong>Verification of Identity</strong></br><strong><mark><mark><mark>SATR 4</mark></mark></mark></strong>: Tout transporteur aérien vérifie, à la porte d’embarquement pour un vol international, l’identité de chaque passager qui semble âgé de 18 ans ou plus au moyen :</br><ul style='list-style-type:none;'><li><strong>(a)</strong> soit de l’une des pièces d’identité avec photo ci-après qui est délivrée par une autorité gouvernementale, qui indique les nom et prénoms, date de naissance et genre du passager et qui est valide :</li><ul style='list-style-type:none;'><li><strong>(i)</strong> un passeport délivré au passager par le pays dont il est citoyen ou ressortissant,</li><li><strong>(ii)</strong> une carte NEXUS,</li><li><strong>(iii)</strong> un document visé au paragraphe 50(1) ou 52(1) du Règlement sur l’immigration et la protection des réfugiés;</li></ul><li><strong>(b)</strong> soit d’une carte d’identité de zone réglementée au sens de l’article 3 du Règlement canadien de 2012 sur la sûreté aérienne qui est valide.</li></ul>"",
                        ""comments"": ""some new comment 2"",
                        ""documentaryEvidence"": ""C:\\fakepath\\somenewfile2.png""
                    }
                }
                "
            };

            var findingId = Guid.NewGuid();
            var finding = new ovs_Finding()
            {
                Id = findingId,
                ovs_Finding1 = "100-345678-1-1", // Already named based of parent work order service task
                ts_findingmappingkey = workOrderServiceTaskId + "-finding-sq_142",
                ovs_FindingProvisionReference = "SATR 4",
                ovs_FindingComments = "existing comments",
                ovs_FindingFile = "C:\\fakepath\\existingfile.png",
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
            Assert.True(findings.Count == 3, "Expecting 3 findings");

            // Expect first ovs_finding to still have the same name
            var first = findings[0];
            Assert.Equal(finding.ovs_FindingProvisionReference, first.ovs_FindingProvisionReference);
            Assert.Equal(finding.ovs_Finding1, first.ovs_Finding1);

            // Expect newly created second ovs_finding to have the proper name
            var second = findings[1];
            Assert.Equal("100-345678-1-2", second.ovs_Finding1);

            // Expect newly created third ovs_finding to have the proper name
            var third = findings[2];
            Assert.Equal("100-345678-1-3", third.ovs_Finding1);
        }

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
                msdyn_name = "200-34567-1",
                msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId), // belongs to a work order
                msdyn_PercentComplete = 0.00,
                ovs_QuestionnaireResponse = @"
                {
                    ""finding-sq_142"": {
                        ""provisionReference"": ""SATR 4"",
                        ""provisionTextEn"": ""<strong>Verification of Identity</strong></br><strong><mark><mark>SATR 4</mark></mark></strong>: An air carrier must, at the boarding gate for an international flight, verify the identity of each passenger who appears to be 18 years of age or older using</br><ul style='list-style-type:none;'><li><strong>(a)</strong> one of the following pieces of photo identification issued by a government authority that shows the passenger’s surname, first name and any middle names, their date of birth and gender and that is valid:</li><ul style='list-style-type:none;'><li><strong>(i)</strong> a passport issued by the country of which the passenger is a citizen or a national,</li><li><strong>(ii)</strong> a NEXUS card,</li><li><strong>(iii)</strong> any document referred to in subsection 50(1) or 52(1) of the Immigration and Refugee Protection Regulations; or</li></ul><li><strong>(b)</strong> a valid restricted area identity card, as defined in section 3 of the Canadian Aviation Security Regulations, 2012.</li></ul>"",
                        ""provisionTextFr"": ""<strong>Verification of Identity</strong></br><strong><mark><mark><mark>SATR 4</mark></mark></mark></strong>: Tout transporteur aérien vérifie, à la porte d’embarquement pour un vol international, l’identité de chaque passager qui semble âgé de 18 ans ou plus au moyen :</br><ul style='list-style-type:none;'><li><strong>(a)</strong> soit de l’une des pièces d’identité avec photo ci-après qui est délivrée par une autorité gouvernementale, qui indique les nom et prénoms, date de naissance et genre du passager et qui est valide :</li><ul style='list-style-type:none;'><li><strong>(i)</strong> un passeport délivré au passager par le pays dont il est citoyen ou ressortissant,</li><li><strong>(ii)</strong> une carte NEXUS,</li><li><strong>(iii)</strong> un document visé au paragraphe 50(1) ou 52(1) du Règlement sur l’immigration et la protection des réfugiés;</li></ul><li><strong>(b)</strong> soit d’une carte d’identité de zone réglementée au sens de l’article 3 du Règlement canadien de 2012 sur la sûreté aérienne qui est valide.</li></ul>"",
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
                msdyn_name = "200-34567-1",
                msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId), // belongs to a work order
                msdyn_PercentComplete = 100.00,
                ovs_QuestionnaireResponse = @"
                {
                    ""finding-sq_142"": {
                        ""provisionReference"": ""SATR 4"",
                        ""provisionTextEn"": ""<strong>Verification of Identity</strong></br><strong><mark><mark>SATR 4</mark></mark></strong>: An air carrier must, at the boarding gate for an international flight, verify the identity of each passenger who appears to be 18 years of age or older using</br><ul style='list-style-type:none;'><li><strong>(a)</strong> one of the following pieces of photo identification issued by a government authority that shows the passenger’s surname, first name and any middle names, their date of birth and gender and that is valid:</li><ul style='list-style-type:none;'><li><strong>(i)</strong> a passport issued by the country of which the passenger is a citizen or a national,</li><li><strong>(ii)</strong> a NEXUS card,</li><li><strong>(iii)</strong> any document referred to in subsection 50(1) or 52(1) of the Immigration and Refugee Protection Regulations; or</li></ul><li><strong>(b)</strong> a valid restricted area identity card, as defined in section 3 of the Canadian Aviation Security Regulations, 2012.</li></ul>"",
                        ""provisionTextFr"": ""<strong>Verification of Identity</strong></br><strong><mark><mark><mark>SATR 4</mark></mark></mark></strong>: Tout transporteur aérien vérifie, à la porte d’embarquement pour un vol international, l’identité de chaque passager qui semble âgé de 18 ans ou plus au moyen :</br><ul style='list-style-type:none;'><li><strong>(a)</strong> soit de l’une des pièces d’identité avec photo ci-après qui est délivrée par une autorité gouvernementale, qui indique les nom et prénoms, date de naissance et genre du passager et qui est valide :</li><ul style='list-style-type:none;'><li><strong>(i)</strong> un passeport délivré au passager par le pays dont il est citoyen ou ressortissant,</li><li><strong>(ii)</strong> une carte NEXUS,</li><li><strong>(iii)</strong> un document visé au paragraphe 50(1) ou 52(1) du Règlement sur l’immigration et la protection des réfugiés;</li></ul><li><strong>(b)</strong> soit d’une carte d’identité de zone réglementée au sens de l’article 3 du Règlement canadien de 2012 sur la sûreté aérienne qui est valide.</li></ul>"",
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
                ts_findingmappingkey = workOrderServiceTaskId + "-finding-sq_142",
                ovs_FindingProvisionReference = "SATR 4",
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
                msdyn_name = "200-34567-1",
                msdyn_PercentComplete = 100.00,
                msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId), // belongs to a work order
                ovs_QuestionnaireResponse = @"
                {
                    ""finding-sq_142"": {
                        ""provisionReference"": ""SATR 4"",
                        ""provisionTextEn"": ""<strong>Verification of Identity</strong></br><strong><mark><mark>SATR 4</mark></mark></strong>: An air carrier must, at the boarding gate for an international flight, verify the identity of each passenger who appears to be 18 years of age or older using</br><ul style='list-style-type:none;'><li><strong>(a)</strong> one of the following pieces of photo identification issued by a government authority that shows the passenger’s surname, first name and any middle names, their date of birth and gender and that is valid:</li><ul style='list-style-type:none;'><li><strong>(i)</strong> a passport issued by the country of which the passenger is a citizen or a national,</li><li><strong>(ii)</strong> a NEXUS card,</li><li><strong>(iii)</strong> any document referred to in subsection 50(1) or 52(1) of the Immigration and Refugee Protection Regulations; or</li></ul><li><strong>(b)</strong> a valid restricted area identity card, as defined in section 3 of the Canadian Aviation Security Regulations, 2012.</li></ul>"",
                        ""provisionTextFr"": ""<strong>Verification of Identity</strong></br><strong><mark><mark><mark>SATR 4</mark></mark></mark></strong>: Tout transporteur aérien vérifie, à la porte d’embarquement pour un vol international, l’identité de chaque passager qui semble âgé de 18 ans ou plus au moyen :</br><ul style='list-style-type:none;'><li><strong>(a)</strong> soit de l’une des pièces d’identité avec photo ci-après qui est délivrée par une autorité gouvernementale, qui indique les nom et prénoms, date de naissance et genre du passager et qui est valide :</li><ul style='list-style-type:none;'><li><strong>(i)</strong> un passeport délivré au passager par le pays dont il est citoyen ou ressortissant,</li><li><strong>(ii)</strong> une carte NEXUS,</li><li><strong>(iii)</strong> un document visé au paragraphe 50(1) ou 52(1) du Règlement sur l’immigration et la protection des réfugiés;</li></ul><li><strong>(b)</strong> soit d’une carte d’identité de zone réglementée au sens de l’article 3 du Règlement canadien de 2012 sur la sûreté aérienne qui est valide.</li></ul>"",
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
                ts_findingmappingkey = workOrderServiceTaskId + "-finding-sq_142", // unique finding names are created using the work order service task ID and the reference ID in the questionnaire
                ovs_FindingProvisionReference = "SATR 4",
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
                msdyn_name = "200-34567-1",
                msdyn_PercentComplete = 100.00,
                msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId), // belongs to a work order
                ovs_QuestionnaireResponse = @"
                {
                    ""finding-sq_142"": {
                        ""provisionReference"": ""SATR 4"",
                        ""provisionTextEn"": ""<strong>Verification of Identity</strong></br><strong><mark><mark>SATR 4</mark></mark></strong>: An air carrier must, at the boarding gate for an international flight, verify the identity of each passenger who appears to be 18 years of age or older using</br><ul style='list-style-type:none;'><li><strong>(a)</strong> one of the following pieces of photo identification issued by a government authority that shows the passenger’s surname, first name and any middle names, their date of birth and gender and that is valid:</li><ul style='list-style-type:none;'><li><strong>(i)</strong> a passport issued by the country of which the passenger is a citizen or a national,</li><li><strong>(ii)</strong> a NEXUS card,</li><li><strong>(iii)</strong> any document referred to in subsection 50(1) or 52(1) of the Immigration and Refugee Protection Regulations; or</li></ul><li><strong>(b)</strong> a valid restricted area identity card, as defined in section 3 of the Canadian Aviation Security Regulations, 2012.</li></ul>"",
                        ""provisionTextFr"": ""<strong>Verification of Identity</strong></br><strong><mark><mark><mark>SATR 4</mark></mark></mark></strong>: Tout transporteur aérien vérifie, à la porte d’embarquement pour un vol international, l’identité de chaque passager qui semble âgé de 18 ans ou plus au moyen :</br><ul style='list-style-type:none;'><li><strong>(a)</strong> soit de l’une des pièces d’identité avec photo ci-après qui est délivrée par une autorité gouvernementale, qui indique les nom et prénoms, date de naissance et genre du passager et qui est valide :</li><ul style='list-style-type:none;'><li><strong>(i)</strong> un passeport délivré au passager par le pays dont il est citoyen ou ressortissant,</li><li><strong>(ii)</strong> une carte NEXUS,</li><li><strong>(iii)</strong> un document visé au paragraphe 50(1) ou 52(1) du Règlement sur l’immigration et la protection des réfugiés;</li></ul><li><strong>(b)</strong> soit d’une carte d’identité de zone réglementée au sens de l’article 3 du Règlement canadien de 2012 sur la sûreté aérienne qui est valide.</li></ul>"",
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
                ts_findingmappingkey = workOrderServiceTaskId + "-finding-sq_142", // unique finding names are created using the work order service task ID and the reference ID in the questionnaire
                ovs_FindingProvisionReference = "SATR 4",
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
                msdyn_name = "200-34567-1",
                msdyn_PercentComplete = 100.00,
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
                ts_findingmappingkey = workOrderServiceTaskId + "-finding-sq_142",
                ovs_FindingProvisionReference = "SATR 4",
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
                msdyn_name = "200-34567-1",
                msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId), // belongs to a work order
                msdyn_PercentComplete = 100.00,
                ovs_QuestionnaireResponse = @"
                {
                    ""finding-sq_142"": {
                        ""provisionReference"": ""SATR 4"",
                        ""provisionTextEn"": ""<strong>Verification of Identity</strong></br><strong><mark><mark>SATR 4</mark></mark></strong>: An air carrier must, at the boarding gate for an international flight, verify the identity of each passenger who appears to be 18 years of age or older using</br><ul style='list-style-type:none;'><li><strong>(a)</strong> one of the following pieces of photo identification issued by a government authority that shows the passenger’s surname, first name and any middle names, their date of birth and gender and that is valid:</li><ul style='list-style-type:none;'><li><strong>(i)</strong> a passport issued by the country of which the passenger is a citizen or a national,</li><li><strong>(ii)</strong> a NEXUS card,</li><li><strong>(iii)</strong> any document referred to in subsection 50(1) or 52(1) of the Immigration and Refugee Protection Regulations; or</li></ul><li><strong>(b)</strong> a valid restricted area identity card, as defined in section 3 of the Canadian Aviation Security Regulations, 2012.</li></ul>"",
                        ""provisionTextFr"": ""<strong>Verification of Identity</strong></br><strong><mark><mark><mark>SATR 4</mark></mark></mark></strong>: Tout transporteur aérien vérifie, à la porte d’embarquement pour un vol international, l’identité de chaque passager qui semble âgé de 18 ans ou plus au moyen :</br><ul style='list-style-type:none;'><li><strong>(a)</strong> soit de l’une des pièces d’identité avec photo ci-après qui est délivrée par une autorité gouvernementale, qui indique les nom et prénoms, date de naissance et genre du passager et qui est valide :</li><ul style='list-style-type:none;'><li><strong>(i)</strong> un passeport délivré au passager par le pays dont il est citoyen ou ressortissant,</li><li><strong>(ii)</strong> une carte NEXUS,</li><li><strong>(iii)</strong> un document visé au paragraphe 50(1) ou 52(1) du Règlement sur l’immigration et la protection des réfugiés;</li></ul><li><strong>(b)</strong> soit d’une carte d’identité de zone réglementée au sens de l’article 3 du Règlement canadien de 2012 sur la sûreté aérienne qui est valide.</li></ul>"",
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
            Assert.Equal("SATR 4", first.ovs_FindingProvisionReference);

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
                msdyn_name = "200-34567-1",
                msdyn_inspectiontaskresult = msdyn_InspectionResult.NA, 
                msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId), // belongs to a work order
                msdyn_PercentComplete = 100.00,
                ovs_QuestionnaireResponse = @"
                {
                    ""finding-sq_142"": {
                        ""provisionReference"": ""SATR 4"",
                        ""provisionTextEn"": ""<strong>Verification of Identity</strong></br><strong><mark><mark>SATR 4</mark></mark></strong>: An air carrier must, at the boarding gate for an international flight, verify the identity of each passenger who appears to be 18 years of age or older using</br><ul style='list-style-type:none;'><li><strong>(a)</strong> one of the following pieces of photo identification issued by a government authority that shows the passenger’s surname, first name and any middle names, their date of birth and gender and that is valid:</li><ul style='list-style-type:none;'><li><strong>(i)</strong> a passport issued by the country of which the passenger is a citizen or a national,</li><li><strong>(ii)</strong> a NEXUS card,</li><li><strong>(iii)</strong> any document referred to in subsection 50(1) or 52(1) of the Immigration and Refugee Protection Regulations; or</li></ul><li><strong>(b)</strong> a valid restricted area identity card, as defined in section 3 of the Canadian Aviation Security Regulations, 2012.</li></ul>"",
                        ""provisionTextFr"": ""<strong>Verification of Identity</strong></br><strong><mark><mark><mark>SATR 4</mark></mark></mark></strong>: Tout transporteur aérien vérifie, à la porte d’embarquement pour un vol international, l’identité de chaque passager qui semble âgé de 18 ans ou plus au moyen :</br><ul style='list-style-type:none;'><li><strong>(a)</strong> soit de l’une des pièces d’identité avec photo ci-après qui est délivrée par une autorité gouvernementale, qui indique les nom et prénoms, date de naissance et genre du passager et qui est valide :</li><ul style='list-style-type:none;'><li><strong>(i)</strong> un passeport délivré au passager par le pays dont il est citoyen ou ressortissant,</li><li><strong>(ii)</strong> une carte NEXUS,</li><li><strong>(iii)</strong> un document visé au paragraphe 50(1) ou 52(1) du Règlement sur l’immigration et la protection des réfugiés;</li></ul><li><strong>(b)</strong> soit d’une carte d’identité de zone réglementée au sens de l’article 3 du Règlement canadien de 2012 sur la sûreté aérienne qui est valide.</li></ul>"",
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
                msdyn_name = "200-34567-1",
                msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId), // belongs to a work order
                msdyn_PercentComplete = 100.00,
                ovs_QuestionnaireResponse = @"
                {
                    ""finding-sq_142"": {
                        ""provisionReference"": ""SATR 4"",
                        ""provisionTextEn"": ""<strong>Verification of Identity</strong></br><strong><mark><mark>SATR 4</mark></mark></strong>: An air carrier must, at the boarding gate for an international flight, verify the identity of each passenger who appears to be 18 years of age or older using</br><ul style='list-style-type:none;'><li><strong>(a)</strong> one of the following pieces of photo identification issued by a government authority that shows the passenger’s surname, first name and any middle names, their date of birth and gender and that is valid:</li><ul style='list-style-type:none;'><li><strong>(i)</strong> a passport issued by the country of which the passenger is a citizen or a national,</li><li><strong>(ii)</strong> a NEXUS card,</li><li><strong>(iii)</strong> any document referred to in subsection 50(1) or 52(1) of the Immigration and Refugee Protection Regulations; or</li></ul><li><strong>(b)</strong> a valid restricted area identity card, as defined in section 3 of the Canadian Aviation Security Regulations, 2012.</li></ul>"",
                        ""provisionTextFr"": ""<strong>Verification of Identity</strong></br><strong><mark><mark><mark>SATR 4</mark></mark></mark></strong>: Tout transporteur aérien vérifie, à la porte d’embarquement pour un vol international, l’identité de chaque passager qui semble âgé de 18 ans ou plus au moyen :</br><ul style='list-style-type:none;'><li><strong>(a)</strong> soit de l’une des pièces d’identité avec photo ci-après qui est délivrée par une autorité gouvernementale, qui indique les nom et prénoms, date de naissance et genre du passager et qui est valide :</li><ul style='list-style-type:none;'><li><strong>(i)</strong> un passeport délivré au passager par le pays dont il est citoyen ou ressortissant,</li><li><strong>(ii)</strong> une carte NEXUS,</li><li><strong>(iii)</strong> un document visé au paragraphe 50(1) ou 52(1) du Règlement sur l’immigration et la protection des réfugiés;</li></ul><li><strong>(b)</strong> soit d’une carte d’identité de zone réglementée au sens de l’article 3 du Règlement canadien de 2012 sur la sûreté aérienne qui est valide.</li></ul>"",
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

            // Expect ovs_finding name to be combination of Work Order Service Task name as suffix and prefix of 1
            Assert.Equal("100-34567-1-1", oneFinding.ovs_Finding1);

            // Expect ovs_finding name to be combination of Work Order Number and Finding Name
            Assert.Equal(workOrderServiceTask.Id.ToString() + "-finding-sq_142", oneFinding.ts_findingmappingkey);

            // Expect ovs_finding to contain same content as what is in JSON response
            Assert.Equal("SATR 4", oneFinding.ovs_FindingProvisionReference);
            Assert.Equal("<strong>Verification of Identity</strong></br><strong><mark><mark>SATR 4</mark></mark></strong>: An air carrier must, at the boarding gate for an international flight, verify the identity of each passenger who appears to be 18 years of age or older using</br><ul style='list-style-type:none;'><li><strong>(a)</strong> one of the following pieces of photo identification issued by a government authority that shows the passenger’s surname, first name and any middle names, their date of birth and gender and that is valid:</li><ul style='list-style-type:none;'><li><strong>(i)</strong> a passport issued by the country of which the passenger is a citizen or a national,</li><li><strong>(ii)</strong> a NEXUS card,</li><li><strong>(iii)</strong> any document referred to in subsection 50(1) or 52(1) of the Immigration and Refugee Protection Regulations; or</li></ul><li><strong>(b)</strong> a valid restricted area identity card, as defined in section 3 of the Canadian Aviation Security Regulations, 2012.</li></ul>", oneFinding.ts_findingProvisionTextEn);
            Assert.Equal("<strong>Verification of Identity</strong></br><strong><mark><mark><mark>SATR 4</mark></mark></mark></strong>: Tout transporteur aérien vérifie, à la porte d’embarquement pour un vol international, l’identité de chaque passager qui semble âgé de 18 ans ou plus au moyen :</br><ul style='list-style-type:none;'><li><strong>(a)</strong> soit de l’une des pièces d’identité avec photo ci-après qui est délivrée par une autorité gouvernementale, qui indique les nom et prénoms, date de naissance et genre du passager et qui est valide :</li><ul style='list-style-type:none;'><li><strong>(i)</strong> un passeport délivré au passager par le pays dont il est citoyen ou ressortissant,</li><li><strong>(ii)</strong> une carte NEXUS,</li><li><strong>(iii)</strong> un document visé au paragraphe 50(1) ou 52(1) du Règlement sur l’immigration et la protection des réfugiés;</li></ul><li><strong>(b)</strong> soit d’une carte d’identité de zone réglementée au sens de l’article 3 du Règlement canadien de 2012 sur la sûreté aérienne qui est valide.</li></ul>", oneFinding.ts_findingProvisionTextFr);
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
                msdyn_name = "200-34567-1",
                msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId), // belongs to a work order
                msdyn_PercentComplete = 100.00,
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

        [Fact]
        public void When_parent_work_order_does_not_have_regulated_entity_throw_argument_null_exception()
        {
            /**********
            * ARRANGE
            **********/
            var context = new XrmFakedContext();

            var workOrderId = Guid.NewGuid();
            var workOrder = new msdyn_workorder()
            {
                Id = workOrderId,
                msdyn_ServiceRequest = null, //Not part of a case
                ovs_regulatedentity = null //No regulated entity set
            };

            var workOrderServiceTaskId = Guid.NewGuid();
            var workOrderServiceTask = new msdyn_workorderservicetask()
            {
                Id = workOrderServiceTaskId,
                msdyn_name = "200-34567-1",
                msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId), // belongs to a work order
                msdyn_PercentComplete = 100.00,
                ovs_QuestionnaireResponse = @"
                {
                    ""finding-sq_142"": {
                        ""provisionReference"": ""SATR 4"",
                        ""provisionTextEn"": ""<strong>Verification of Identity</strong></br><strong><mark><mark>SATR 4</mark></mark></strong>: An air carrier must, at the boarding gate for an international flight, verify the identity of each passenger who appears to be 18 years of age or older using</br><ul style='list-style-type:none;'><li><strong>(a)</strong> one of the following pieces of photo identification issued by a government authority that shows the passenger’s surname, first name and any middle names, their date of birth and gender and that is valid:</li><ul style='list-style-type:none;'><li><strong>(i)</strong> a passport issued by the country of which the passenger is a citizen or a national,</li><li><strong>(ii)</strong> a NEXUS card,</li><li><strong>(iii)</strong> any document referred to in subsection 50(1) or 52(1) of the Immigration and Refugee Protection Regulations; or</li></ul><li><strong>(b)</strong> a valid restricted area identity card, as defined in section 3 of the Canadian Aviation Security Regulations, 2012.</li></ul>"",
                        ""provisionTextFr"": ""<strong>Verification of Identity</strong></br><strong><mark><mark><mark>SATR 4</mark></mark></mark></strong>: Tout transporteur aérien vérifie, à la porte d’embarquement pour un vol international, l’identité de chaque passager qui semble âgé de 18 ans ou plus au moyen :</br><ul style='list-style-type:none;'><li><strong>(a)</strong> soit de l’une des pièces d’identité avec photo ci-après qui est délivrée par une autorité gouvernementale, qui indique les nom et prénoms, date de naissance et genre du passager et qui est valide :</li><ul style='list-style-type:none;'><li><strong>(i)</strong> un passeport délivré au passager par le pays dont il est citoyen ou ressortissant,</li><li><strong>(ii)</strong> une carte NEXUS,</li><li><strong>(iii)</strong> un document visé au paragraphe 50(1) ou 52(1) du Règlement sur l’immigration et la protection des réfugiés;</li></ul><li><strong>(b)</strong> soit d’une carte d’identité de zone réglementée au sens de l’article 3 du Règlement canadien de 2012 sur la sûreté aérienne qui est valide.</li></ul>"",
                        ""comments"": ""new comments"",
                        ""documentaryEvidence"": ""C:\\fakepath\\newfile.png""
                    }
                }
                "
            };

            context.Initialize(
                new List<Entity>() {
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

            /***************
            * ACT and ASSERT
            ****************/
            // Execute the PreOperationmsdyn_workorderservicetaskUpdate plugin with the defined workOrderServiceTask as a target
            Assert.Throws<ArgumentNullException>(() => context.ExecutePluginWith<PreOperationmsdyn_workorderservicetaskUpdate>(inputParams, outputParams, preEntityImages, null));
        }
    }
}
