﻿using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using Xunit;

namespace ROMTS_GSRST.Plugins.Tests
{
    public class PreOperationmsdyn_workorderservicetaskUpdateTests : UnitTestBase
    {
        public PreOperationmsdyn_workorderservicetaskUpdateTests(XrmMockupFixture fixture) : base(fixture) { }

        [Fact]
        public void When_not_all_work_order_service_task_is_complete_expect_parent_work_order_to_not_change()
        {
            // ARRANGE
            var serviceAccountId = orgAdminUIService.Create(new Account { Name = "Test Service Account" });
            var incidentId = orgAdminUIService.Create(new Incident { });
            var workOrderId = orgAdminUIService.Create(new msdyn_workorder
            {
                msdyn_name = "300-345678",
                msdyn_SystemStatus = msdyn_wosystemstatus.New,
                msdyn_ServiceRequest = new EntityReference(Incident.EntityLogicalName, incidentId),
                msdyn_ServiceAccount = new EntityReference(Account.EntityLogicalName, serviceAccountId)
            });
            var existingWorkOrderServiceTask1Id = orgAdminUIService.Create(new msdyn_workorderservicetask
            {
                msdyn_name = "200-345678-1",
                msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId), // belongs to a work order
                msdyn_PercentComplete = 0,
                ovs_QuestionnaireResponse = "",
                statuscode = msdyn_workorderservicetask_statuscode.InProgress
            });
            var existingWorkOrderServiceTask2Id = orgAdminUIService.Create(new msdyn_workorderservicetask
            {
                msdyn_name = "200-345678-2",
                msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId), // belongs to a work order
                msdyn_PercentComplete = 0,
                ovs_QuestionnaireResponse = "",
                statuscode = msdyn_workorderservicetask_statuscode.New
            });

            // ACT
            orgAdminUIService.Update(new msdyn_workorderservicetask { Id = existingWorkOrderServiceTask2Id, msdyn_PercentComplete = 100.00, ovs_QuestionnaireResponse = "" });

            // ASSERT
            var workOrder = orgAdminUIService.Retrieve(msdyn_workorder.EntityLogicalName, workOrderId, new ColumnSet("msdyn_systemstatus")).ToEntity<msdyn_workorder>();
            Assert.Equal(msdyn_wosystemstatus.New, workOrder.msdyn_SystemStatus);
        }

        [Fact]
        public void When_all_work_order_service_tasks_are_complete_expect_parent_work_order_to_be_open_completed()
        {
            // ARRANGE
            var serviceAccountId = orgAdminUIService.Create(new Account { Name = "Test Service Account" });
            var incidentId = orgAdminUIService.Create(new Incident { });
            var workOrderId = orgAdminUIService.Create(new msdyn_workorder
            {
                msdyn_name = "300-345678",
                msdyn_SystemStatus = msdyn_wosystemstatus.New,
                msdyn_ServiceRequest = new EntityReference(Incident.EntityLogicalName, incidentId),
                msdyn_ServiceAccount = new EntityReference(Account.EntityLogicalName, serviceAccountId)
            });
            var existingWorkOrderServiceTask1Id = orgAdminUIService.Create(new msdyn_workorderservicetask
            {
                msdyn_name = "200-345678-1",
                msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId), // belongs to a work order
                msdyn_PercentComplete = 100.00,
                ovs_QuestionnaireResponse = "",
                statuscode = msdyn_workorderservicetask_statuscode.Complete
            });
            var existingWorkOrderServiceTask2Id = orgAdminUIService.Create(new msdyn_workorderservicetask
            {
                msdyn_name = "200-345678-2",
                msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId), // belongs to a work order
                msdyn_PercentComplete = 0,
                ovs_QuestionnaireResponse = "",
                statuscode = msdyn_workorderservicetask_statuscode.InProgress
            });

            // ACT
            orgAdminUIService.Update(new msdyn_workorderservicetask { Id = existingWorkOrderServiceTask2Id, msdyn_PercentComplete = 100.00, ovs_QuestionnaireResponse = "" });

            // ASSERT
            var workOrder = orgAdminUIService.Retrieve(msdyn_workorder.EntityLogicalName, workOrderId, new ColumnSet("msdyn_systemstatus")).ToEntity<msdyn_workorder>();
            Assert.Equal(msdyn_wosystemstatus.New, workOrder.msdyn_SystemStatus);
        }

        [Fact]
        public void When_work_order_service_task_already_has_finding_expect_next_findings_to_have_name_with_incremented_infix()
        {
            // ARRANGE
            var serviceAccountId = orgAdminUIService.Create(new Account { Name = "Test Service Account" });
            var incidentId = orgAdminUIService.Create(new Incident { });
            var workOrderId = orgAdminUIService.Create(new msdyn_workorder
            {
                msdyn_name = "300-345678",
                msdyn_ServiceRequest = new EntityReference(Incident.EntityLogicalName, incidentId),
                msdyn_ServiceAccount = new EntityReference(Account.EntityLogicalName, serviceAccountId)
            });
            var account = orgAdminUIService.Create(new Account());
            var accountReference = new EntityReference(Account.EntityLogicalName, account);
            var operationId = orgAdminUIService.Create(new ovs_operation()
            {
                Id = new Guid("9de3a6e3-c4ad-eb11-8236-000d3ae8b866"),
                ts_stakeholder = accountReference
            });

            orgAdminUIService.Create(new qm_rclegislation()
            {
                Id = new Guid("417d6241-386a-eb11-a812-000d3a84129e"),
                qm_name = "SATR 4",
                ts_NameEnglish = "SATR 4",
                ts_NameFrench = "RSDA 4",
            });
            orgAdminUIService.Create(new qm_rclegislation()
            {
                Id = new Guid("517d6241-386a-eb11-a812-000d3a84129e"),
                qm_name = "SATR 5",
                ts_NameEnglish = "SATR 5",
                ts_NameFrench = "RSDA 5",
            });
            orgAdminUIService.Create(new qm_rclegislation()
            {
                Id = new Guid("617d6241-386a-eb11-a812-000d3a84129e"),
                qm_name = "SATR 6",
                ts_NameEnglish = "SATR 6",
                ts_NameFrench = "RSDA 6",
            });

            var existingWorkOrderServiceTaskId = orgAdminUIService.Create(new msdyn_workorderservicetask
            {
                msdyn_name = "200-345678-1",
                msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId), // belongs to a work order
                msdyn_PercentComplete = 0,
                ovs_QuestionnaireResponse = @"
                {
                    ""finding-sq_142"": {
                        ""provisionReference"": ""SATR 4"",
                        ""provisionTextEn"": ""<strong>Verification of Identity</strong></br><strong><mark><mark>SATR 4</mark></mark></strong>: An air carrier must, at the boarding gate for an international flight, verify the identity of each passenger who appears to be 18 years of age or older using</br><ul style='list-style-type:none;'><li><strong>(a)</strong> one of the following pieces of photo identification issued by a government authority that shows the passenger’s surname, first name and any middle names, their date of birth and gender and that is valid:</li><ul style='list-style-type:none;'><li><strong>(i)</strong> a passport issued by the country of which the passenger is a citizen or a national,</li><li><strong>(ii)</strong> a NEXUS card,</li><li><strong>(iii)</strong> any document referred to in subsection 50(1) or 52(1) of the Immigration and Refugee Protection Regulations; or</li></ul><li><strong>(b)</strong> a valid restricted area identity card, as defined in section 3 of the Canadian Aviation Security Regulations, 2012.</li></ul>"",
                        ""provisionTextFr"": ""<strong>Verification of Identity</strong></br><strong><mark><mark><mark>SATR 4</mark></mark></mark></strong>: Tout transporteur aérien vérifie, à la porte d’embarquement pour un vol international, l’identité de chaque passager qui semble âgé de 18 ans ou plus au moyen :</br><ul style='list-style-type:none;'><li><strong>(a)</strong> soit de l’une des pièces d’identité avec photo ci-après qui est délivrée par une autorité gouvernementale, qui indique les nom et prénoms, date de naissance et genre du passager et qui est valide :</li><ul style='list-style-type:none;'><li><strong>(i)</strong> un passeport délivré au passager par le pays dont il est citoyen ou ressortissant,</li><li><strong>(ii)</strong> une carte NEXUS,</li><li><strong>(iii)</strong> un document visé au paragraphe 50(1) ou 52(1) du Règlement sur l’immigration et la protection des réfugiés;</li></ul><li><strong>(b)</strong> soit d’une carte d’identité de zone réglementée au sens de l’article 3 du Règlement canadien de 2012 sur la sûreté aérienne qui est valide.</li></ul>"",
                        ""comments"": ""existing comments"",
                        ""operations"": [{""operationID"": ""9de3a6e3-c4ad-eb11-8236-000d3ae8b866"",""findingType"": ""717750000""}],
                        ""documentaryEvidence"": ""C:\\fakepath\\exitingfile.png""
                    }
                }
                "
            });
            var findingId1 = orgAdminUIService.Create(new ovs_Finding
            {
                ovs_Finding_1 = "100-345678-1-1-1",
                ts_findingmappingkey = existingWorkOrderServiceTaskId + "-finding-sq_142-" + operationId, // unique finding names are created using the work order service task ID and the reference ID in the questionnaire
                ovs_FindingProvisionReference = "SATR 4",
                ovs_FindingComments = "existing comments",
                ovs_FindingFile = "C:\\fakepath\\originalfile.png",
                ovs_CaseId = new EntityReference(Incident.EntityLogicalName, incidentId), // this finding already belongs to the case
                ovs_WorkOrderServiceTaskId = new EntityReference(msdyn_workorderservicetask.EntityLogicalName, existingWorkOrderServiceTaskId), // this finding already belongs to a work order service task
                statuscode = ovs_Finding_statuscode.New, // finding is also already active
                statecode = ovs_FindingState.Active, // finding is also already active
            });

            // ACT
            orgAdminUIService.Update(new msdyn_workorderservicetask
            {
                Id = existingWorkOrderServiceTaskId,
                msdyn_PercentComplete = 100.00,
                ovs_QuestionnaireResponse = @"
                {
                    ""finding-sq_142"": {
                        ""provisionReference"": ""SATR 4"",
                        ""provisionTextEn"": ""<strong>Verification of Identity</strong></br><strong><mark><mark>SATR 4</mark></mark></strong>: An air carrier must, at the boarding gate for an international flight, verify the identity of each passenger who appears to be 18 years of age or older using</br><ul style='list-style-type:none;'><li><strong>(a)</strong> one of the following pieces of photo identification issued by a government authority that shows the passenger’s surname, first name and any middle names, their date of birth and gender and that is valid:</li><ul style='list-style-type:none;'><li><strong>(i)</strong> a passport issued by the country of which the passenger is a citizen or a national,</li><li><strong>(ii)</strong> a NEXUS card,</li><li><strong>(iii)</strong> any document referred to in subsection 50(1) or 52(1) of the Immigration and Refugee Protection Regulations; or</li></ul><li><strong>(b)</strong> a valid restricted area identity card, as defined in section 3 of the Canadian Aviation Security Regulations, 2012.</li></ul>"",
                        ""provisionTextFr"": ""<strong>Verification of Identity</strong></br><strong><mark><mark><mark>SATR 4</mark></mark></mark></strong>: Tout transporteur aérien vérifie, à la porte d’embarquement pour un vol international, l’identité de chaque passager qui semble âgé de 18 ans ou plus au moyen :</br><ul style='list-style-type:none;'><li><strong>(a)</strong> soit de l’une des pièces d’identité avec photo ci-après qui est délivrée par une autorité gouvernementale, qui indique les nom et prénoms, date de naissance et genre du passager et qui est valide :</li><ul style='list-style-type:none;'><li><strong>(i)</strong> un passeport délivré au passager par le pays dont il est citoyen ou ressortissant,</li><li><strong>(ii)</strong> une carte NEXUS,</li><li><strong>(iii)</strong> un document visé au paragraphe 50(1) ou 52(1) du Règlement sur l’immigration et la protection des réfugiés;</li></ul><li><strong>(b)</strong> soit d’une carte d’identité de zone réglementée au sens de l’article 3 du Règlement canadien de 2012 sur la sûreté aérienne qui est valide.</li></ul>"",
                        ""comments"": ""existing comments"",
                        ""operations"": [{""operationID"": ""9de3a6e3-c4ad-eb11-8236-000d3ae8b866"",""findingType"": ""717750001""}],
                        ""documentaryEvidence"": ""C:\\fakepath\\exitingfile.png""
                    },
                    ""finding-sq_155"": {
                        ""provisionReference"": ""SATR 5"",
                        ""provisionTextEn"": ""<strong>Verification of Identity</strong></br><strong><mark><mark>SATR 4</mark></mark></strong>: An air carrier must, at the boarding gate for an international flight, verify the identity of each passenger who appears to be 18 years of age or older using</br><ul style='list-style-type:none;'><li><strong>(a)</strong> one of the following pieces of photo identification issued by a government authority that shows the passenger’s surname, first name and any middle names, their date of birth and gender and that is valid:</li><ul style='list-style-type:none;'><li><strong>(i)</strong> a passport issued by the country of which the passenger is a citizen or a national,</li><li><strong>(ii)</strong> a NEXUS card,</li><li><strong>(iii)</strong> any document referred to in subsection 50(1) or 52(1) of the Immigration and Refugee Protection Regulations; or</li></ul><li><strong>(b)</strong> a valid restricted area identity card, as defined in section 3 of the Canadian Aviation Security Regulations, 2012.</li></ul>"",
                        ""provisionTextFr"": ""<strong>Verification of Identity</strong></br><strong><mark><mark><mark>SATR 4</mark></mark></mark></strong>: Tout transporteur aérien vérifie, à la porte d’embarquement pour un vol international, l’identité de chaque passager qui semble âgé de 18 ans ou plus au moyen :</br><ul style='list-style-type:none;'><li><strong>(a)</strong> soit de l’une des pièces d’identité avec photo ci-après qui est délivrée par une autorité gouvernementale, qui indique les nom et prénoms, date de naissance et genre du passager et qui est valide :</li><ul style='list-style-type:none;'><li><strong>(i)</strong> un passeport délivré au passager par le pays dont il est citoyen ou ressortissant,</li><li><strong>(ii)</strong> une carte NEXUS,</li><li><strong>(iii)</strong> un document visé au paragraphe 50(1) ou 52(1) du Règlement sur l’immigration et la protection des réfugiés;</li></ul><li><strong>(b)</strong> soit d’une carte d’identité de zone réglementée au sens de l’article 3 du Règlement canadien de 2012 sur la sûreté aérienne qui est valide.</li></ul>"",
                        ""comments"": ""some new comment"",
                        ""operations"": [{""operationID"": ""9de3a6e3-c4ad-eb11-8236-000d3ae8b866"",""findingType"": ""717750001""}],
                        ""documentaryEvidence"": ""C:\\fakepath\\somenewfile.png""
                    },
                    ""finding-sq_166"": {
                        ""provisionReference"": ""SATR 6"",
                        ""provisionTextEn"": ""<strong>Verification of Identity</strong></br><strong><mark><mark>SATR 4</mark></mark></strong>: An air carrier must, at the boarding gate for an international flight, verify the identity of each passenger who appears to be 18 years of age or older using</br><ul style='list-style-type:none;'><li><strong>(a)</strong> one of the following pieces of photo identification issued by a government authority that shows the passenger’s surname, first name and any middle names, their date of birth and gender and that is valid:</li><ul style='list-style-type:none;'><li><strong>(i)</strong> a passport issued by the country of which the passenger is a citizen or a national,</li><li><strong>(ii)</strong> a NEXUS card,</li><li><strong>(iii)</strong> any document referred to in subsection 50(1) or 52(1) of the Immigration and Refugee Protection Regulations; or</li></ul><li><strong>(b)</strong> a valid restricted area identity card, as defined in section 3 of the Canadian Aviation Security Regulations, 2012.</li></ul>"",
                        ""provisionTextFr"": ""<strong>Verification of Identity</strong></br><strong><mark><mark><mark>SATR 4</mark></mark></mark></strong>: Tout transporteur aérien vérifie, à la porte d’embarquement pour un vol international, l’identité de chaque passager qui semble âgé de 18 ans ou plus au moyen :</br><ul style='list-style-type:none;'><li><strong>(a)</strong> soit de l’une des pièces d’identité avec photo ci-après qui est délivrée par une autorité gouvernementale, qui indique les nom et prénoms, date de naissance et genre du passager et qui est valide :</li><ul style='list-style-type:none;'><li><strong>(i)</strong> un passeport délivré au passager par le pays dont il est citoyen ou ressortissant,</li><li><strong>(ii)</strong> une carte NEXUS,</li><li><strong>(iii)</strong> un document visé au paragraphe 50(1) ou 52(1) du Règlement sur l’immigration et la protection des réfugiés;</li></ul><li><strong>(b)</strong> soit d’une carte d’identité de zone réglementée au sens de l’article 3 du Règlement canadien de 2012 sur la sûreté aérienne qui est valide.</li></ul>"",
                        ""comments"": ""some new comment 2"",
                        ""operations"": [{""operationID"": ""9de3a6e3-c4ad-eb11-8236-000d3ae8b866"",""findingType"": ""717750001""}],
                        ""documentaryEvidence"": ""C:\\fakepath\\somenewfile2.png""
                    }
                }
                "
            });

            // ASSERT
            var query = new QueryExpression(ovs_Finding.EntityLogicalName)
            {
                ColumnSet = new ColumnSet("ovs_finding")
            };
            //Retrieve findings ordering by infix
            var findings = orgAdminUIService.RetrieveMultiple(query).Entities.Cast<ovs_Finding>().OrderBy(f => f.ovs_Finding_1.Split('-')[3]).ToList();

            // Expect target to contain 3 ovs_finding references
            Assert.Equal(3, findings.Count());

            // Expect first ovs_finding to still have the same name
            var first = findings[0];
            Assert.Equal("100-345678-1-1-1", first.ovs_Finding_1);

            // Expect newly created second ovs_finding to have the proper name
            var second = findings[1];
            Assert.Equal("100-345678-1-2-1", second.ovs_Finding_1);

            // Expect newly created third ovs_finding to have the proper name
            var third = findings[2];
            Assert.Equal("100-345678-1-3-1", third.ovs_Finding_1);
        }

        [Fact]
        public void When_ovs_questionnaireresponse_contains_finding_containing_operation_array_expect_findings_for_each_operation_to_be_created_with_incremented_name_suffix()
        {

            // ARRANGE
            var serviceAccountId = orgAdminUIService.Create(new Account() { Name = "Test Service Account" });
            var incidentId = orgAdminUIService.Create(new Incident());
            var workOrderId = orgAdminUIService.Create(new msdyn_workorder()
            {
                msdyn_name = "300-345678",
                msdyn_ServiceRequest = new EntityReference(Incident.EntityLogicalName, incidentId),
                msdyn_ServiceAccount = new EntityReference(Account.EntityLogicalName, serviceAccountId),
            });

            var testAccount1Id = orgAdminUIService.Create(new Account());
            var testAccount2Id = orgAdminUIService.Create(new Account());
            var testAccount3Id = orgAdminUIService.Create(new Account());
            var testAccount4Id = orgAdminUIService.Create(new Account());

            var testAccountReference1 = new EntityReference(Account.EntityLogicalName, testAccount1Id);
            var testAccountReference2 = new EntityReference(Account.EntityLogicalName, testAccount2Id);
            var testAccountReference3 = new EntityReference(Account.EntityLogicalName, testAccount3Id);
            var testAccountReference4 = new EntityReference(Account.EntityLogicalName, testAccount4Id);

            var testOperation1Id = orgAdminUIService.Create(new ovs_operation()
            {
                Id = new Guid("9de3a6e3-c4ad-eb11-8236-000d3ae8b866"),
                ts_stakeholder = testAccountReference1
            });
            var testOperation2Id = orgAdminUIService.Create(new ovs_operation()
            {
                Id = new Guid("6b796de3-b3a4-eb11-9442-000d3a8410dc"),
                ts_stakeholder = testAccountReference2
            });
            var testOperation3Id = orgAdminUIService.Create(new ovs_operation()
            {
                Id = new Guid("7d085d54-c2a9-eb11-9442-000d3a8410dc"),
                ts_stakeholder = testAccountReference3
            });
            var testOperation4Id = orgAdminUIService.Create(new ovs_operation()
            {
                Id = new Guid("22364b7e-e1ce-eb11-bacc-0022483c068d"),
                ts_stakeholder = testAccountReference4
            });

            orgAdminUIService.Create(new qm_rclegislation()
            {
                Id = new Guid("417d6241-386a-eb11-a812-000d3a84129e"),
                qm_name = "SATR 4",
                ts_NameEnglish = "SATR 4",
                ts_NameFrench = "RSDA 4",
            });

            var workOrderServiceTaskId = orgAdminUIService.Create(new msdyn_workorderservicetask()
            {
                msdyn_name = "200-345678-1",
                msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId), // belongs to a work order
                msdyn_PercentComplete = 100.00,
                ovs_QuestionnaireResponse = @"
                {
                    ""finding-sq_142"": {
                        ""provisionReference"": ""SATR 4"",
                        ""provisionTextEn"": ""<strong>Verification of Identity</strong></br><strong><mark><mark>SATR 4</mark></mark></strong>: An air carrier must, at the boarding gate for an international flight, verify the identity of each passenger who appears to be 18 years of age or older using</br><ul style='list-style-type:none;'><li><strong>(a)</strong> one of the following pieces of photo identification issued by a government authority that shows the passenger’s surname, first name and any middle names, their date of birth and gender and that is valid:</li><ul style='list-style-type:none;'><li><strong>(i)</strong> a passport issued by the country of which the passenger is a citizen or a national,</li><li><strong>(ii)</strong> a NEXUS card,</li><li><strong>(iii)</strong> any document referred to in subsection 50(1) or 52(1) of the Immigration and Refugee Protection Regulations; or</li></ul><li><strong>(b)</strong> a valid restricted area identity card, as defined in section 3 of the Canadian Aviation Security Regulations, 2012.</li></ul>"",
                        ""provisionTextFr"": ""<strong>Verification of Identity</strong></br><strong><mark><mark><mark>SATR 4</mark></mark></mark></strong>: Tout transporteur aérien vérifie, à la porte d’embarquement pour un vol international, l’identité de chaque passager qui semble âgé de 18 ans ou plus au moyen :</br><ul style='list-style-type:none;'><li><strong>(a)</strong> soit de l’une des pièces d’identité avec photo ci-après qui est délivrée par une autorité gouvernementale, qui indique les nom et prénoms, date de naissance et genre du passager et qui est valide :</li><ul style='list-style-type:none;'><li><strong>(i)</strong> un passeport délivré au passager par le pays dont il est citoyen ou ressortissant,</li><li><strong>(ii)</strong> une carte NEXUS,</li><li><strong>(iii)</strong> un document visé au paragraphe 50(1) ou 52(1) du Règlement sur l’immigration et la protection des réfugiés;</li></ul><li><strong>(b)</strong> soit d’une carte d’identité de zone réglementée au sens de l’article 3 du Règlement canadien de 2012 sur la sûreté aérienne qui est valide.</li></ul>"",
                        ""comments"": ""old comments"",
                        ""operations"": [
	                        {""operationID"": ""9de3a6e3-c4ad-eb11-8236-000d3ae8b866"",""findingType"": ""717750001""},
	                        {""operationID"": ""6b796de3-b3a4-eb11-9442-000d3a8410dc"",""findingType"": ""717750001""},
	                        {""operationID"": ""7d085d54-c2a9-eb11-9442-000d3a8410dc"",""findingType"": ""717750001""},
	                        {""operationID"": ""22364b7e-e1ce-eb11-bacc-0022483c068d"",""findingType"": ""717750001""}
                        ],
                        ""documentaryEvidence"": ""C:\\fakepath\\oldfile.png""
                    }
                }
                "
            });

            // ACT

            orgAdminUIService.Update(new msdyn_workorderservicetask { Id = workOrderServiceTaskId });

            // ASSERT
            var query = new QueryExpression(ovs_Finding.EntityLogicalName)
            {
                ColumnSet = new ColumnSet("ovs_finding")
            };
            //Retrieve findings, sorting by suffix in case they aren't retrieved in order
            var findings = orgAdminUIService.RetrieveMultiple(query).Entities.Cast<ovs_Finding>().OrderBy(f => f.ovs_Finding_1.Split('-')[4]).ToList();

            // Expect 4 findings to be created
            Assert.Equal(4, findings.Count);

            // Expect first ovs_finding to still have the same name
            var first = findings[0];
            Assert.Equal("100-345678-1-1-1", first.ovs_Finding_1);

            // Expect newly created second ovs_finding to have the proper name
            var second = findings[1];
            Assert.Equal("100-345678-1-1-2", second.ovs_Finding_1);

            // Expect newly created third ovs_finding to have the proper name
            var third = findings[2];
            Assert.Equal("100-345678-1-1-3", third.ovs_Finding_1);

            // Expect newly created fourth ovs_finding to have the proper name
            var fourth = findings[3];
            Assert.Equal("100-345678-1-1-4", fourth.ovs_Finding_1);

        }

        [Fact]
        public void When_ovs_questionnaireresponse_contains_finding_containing_operation_array_expect_findings_for_each_operation_to_be_created_with_Entity_References_to_Account_and_Operation()
        {

            // ARRANGE
            var serviceAccountId = orgAdminUIService.Create(new Account() { Name = "Test Service Account" });
            var incidentId = orgAdminUIService.Create(new Incident());
            // var siteId = orgAdminUIService.Create(new msdyn_FunctionalLocation());
            var workOrderId = orgAdminUIService.Create(new msdyn_workorder()
            {
                msdyn_name = "300-345678",
                msdyn_ServiceRequest = new EntityReference(Incident.EntityLogicalName, incidentId),
                msdyn_ServiceAccount = new EntityReference(Account.EntityLogicalName, serviceAccountId)
            });

            var testSite1Id = orgAdminUIService.Create(new msdyn_FunctionalLocation());
            var testSite2Id = orgAdminUIService.Create(new msdyn_FunctionalLocation());
            var testSite3Id = orgAdminUIService.Create(new msdyn_FunctionalLocation());
            var testSite4Id = orgAdminUIService.Create(new msdyn_FunctionalLocation());

            var testOperationType1Id = orgAdminUIService.Create(new ovs_operationtype());
            var testOperationType2Id = orgAdminUIService.Create(new ovs_operationtype());
            var testOperationType3Id = orgAdminUIService.Create(new ovs_operationtype());
            var testOperationType4Id = orgAdminUIService.Create(new ovs_operationtype());

            var testAccount1Id = orgAdminUIService.Create(new Account());
            var testAccount2Id = orgAdminUIService.Create(new Account());
            var testAccount3Id = orgAdminUIService.Create(new Account());
            var testAccount4Id = orgAdminUIService.Create(new Account());

            var testSiteReference1 = new EntityReference(msdyn_FunctionalLocation.EntityLogicalName, testSite1Id);
            var testSiteReference2 = new EntityReference(msdyn_FunctionalLocation.EntityLogicalName, testSite2Id);
            var testSiteReference3 = new EntityReference(msdyn_FunctionalLocation.EntityLogicalName, testSite3Id);
            var testSiteReference4 = new EntityReference(msdyn_FunctionalLocation.EntityLogicalName, testSite4Id);

            var testOperationTypeReference1 = new EntityReference(ovs_operationtype.EntityLogicalName, testOperationType1Id);
            var testOperationTypeReference2 = new EntityReference(ovs_operationtype.EntityLogicalName, testOperationType2Id);
            var testOperationTypeReference3 = new EntityReference(ovs_operationtype.EntityLogicalName, testOperationType3Id);
            var testOperationTypeReference4 = new EntityReference(ovs_operationtype.EntityLogicalName, testOperationType4Id);

            var testAccountReference1 = new EntityReference(Account.EntityLogicalName, testAccount1Id);
            var testAccountReference2 = new EntityReference(Account.EntityLogicalName, testAccount2Id);
            var testAccountReference3 = new EntityReference(Account.EntityLogicalName, testAccount3Id);
            var testAccountReference4 = new EntityReference(Account.EntityLogicalName, testAccount4Id);

            var testOperation1Id = orgAdminUIService.Create(new ovs_operation()
            {
                Id = new Guid("9de3a6e3-c4ad-eb11-8236-000d3ae8b866"),
                ts_stakeholder = testAccountReference1,
                ts_site = testSiteReference1,
                ovs_OperationTypeId = testOperationTypeReference1
            });
            var testOperation2Id = orgAdminUIService.Create(new ovs_operation()
            {
                Id = new Guid("6b796de3-b3a4-eb11-9442-000d3a8410dc"),
                ts_stakeholder = testAccountReference2,
                ts_site = testSiteReference2,
                ovs_OperationTypeId = testOperationTypeReference2
            });
            var testOperation3Id = orgAdminUIService.Create(new ovs_operation()
            {
                Id = new Guid("7d085d54-c2a9-eb11-9442-000d3a8410dc"),
                ts_stakeholder = testAccountReference3,
                ts_site = testSiteReference3,
                ovs_OperationTypeId = testOperationTypeReference3
            });
            var testOperation4Id = orgAdminUIService.Create(new ovs_operation()
            {
                Id = new Guid("22364b7e-e1ce-eb11-bacc-0022483c068d"),
                ts_stakeholder = testAccountReference4,
                ts_site = testSiteReference4,
                ovs_OperationTypeId = testOperationTypeReference4
            });

           orgAdminUIService.Create(new qm_rclegislation()
            {
                Id = new Guid("417d6241-386a-eb11-a812-000d3a84129e"),
                qm_name = "SATR 4",
                ts_NameEnglish = "SATR 4",
                ts_NameFrench = "RSDA 4",
            });

            var workOrderServiceTaskId = orgAdminUIService.Create(new msdyn_workorderservicetask()
            {
                msdyn_name = "200-345678-1",
                msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId), // belongs to a work order
                msdyn_PercentComplete = 100.00,
                ovs_QuestionnaireResponse = @"
                {
                    ""finding-sq_142"": {
                        ""provisionReference"": ""SATR 4"",
                        ""provisionTextEn"": ""<strong>Verification of Identity</strong></br><strong><mark><mark>SATR 4</mark></mark></strong>: An air carrier must, at the boarding gate for an international flight, verify the identity of each passenger who appears to be 18 years of age or older using</br><ul style='list-style-type:none;'><li><strong>(a)</strong> one of the following pieces of photo identification issued by a government authority that shows the passenger’s surname, first name and any middle names, their date of birth and gender and that is valid:</li><ul style='list-style-type:none;'><li><strong>(i)</strong> a passport issued by the country of which the passenger is a citizen or a national,</li><li><strong>(ii)</strong> a NEXUS card,</li><li><strong>(iii)</strong> any document referred to in subsection 50(1) or 52(1) of the Immigration and Refugee Protection Regulations; or</li></ul><li><strong>(b)</strong> a valid restricted area identity card, as defined in section 3 of the Canadian Aviation Security Regulations, 2012.</li></ul>"",
                        ""provisionTextFr"": ""<strong>Verification of Identity</strong></br><strong><mark><mark><mark>SATR 4</mark></mark></mark></strong>: Tout transporteur aérien vérifie, à la porte d’embarquement pour un vol international, l’identité de chaque passager qui semble âgé de 18 ans ou plus au moyen :</br><ul style='list-style-type:none;'><li><strong>(a)</strong> soit de l’une des pièces d’identité avec photo ci-après qui est délivrée par une autorité gouvernementale, qui indique les nom et prénoms, date de naissance et genre du passager et qui est valide :</li><ul style='list-style-type:none;'><li><strong>(i)</strong> un passeport délivré au passager par le pays dont il est citoyen ou ressortissant,</li><li><strong>(ii)</strong> une carte NEXUS,</li><li><strong>(iii)</strong> un document visé au paragraphe 50(1) ou 52(1) du Règlement sur l’immigration et la protection des réfugiés;</li></ul><li><strong>(b)</strong> soit d’une carte d’identité de zone réglementée au sens de l’article 3 du Règlement canadien de 2012 sur la sûreté aérienne qui est valide.</li></ul>"",
                        ""comments"": ""old comments"",
                        ""operations"": [
	                        {""operationID"": ""9de3a6e3-c4ad-eb11-8236-000d3ae8b866"",""findingType"": ""717750001""},
	                        {""operationID"": ""6b796de3-b3a4-eb11-9442-000d3a8410dc"",""findingType"": ""717750001""},
	                        {""operationID"": ""7d085d54-c2a9-eb11-9442-000d3a8410dc"",""findingType"": ""717750001""},
	                        {""operationID"": ""22364b7e-e1ce-eb11-bacc-0022483c068d"",""findingType"": ""717750001""}
                        ],
                        ""documentaryEvidence"": ""C:\\fakepath\\oldfile.png""
                    }
                }
                "
            });

            // ACT

            orgAdminUIService.Update(new msdyn_workorderservicetask { Id = workOrderServiceTaskId });

            // ASSERT
            var query = new QueryExpression(ovs_Finding.EntityLogicalName)
            {
                ColumnSet = new ColumnSet("ovs_finding", "ts_accountid", "ts_operationid", "ts_functionallocation", "ts_ovs_operationtype")
            };
            var findings = orgAdminUIService.RetrieveMultiple(query).Entities.Cast<ovs_Finding>().OrderBy(f => f.ovs_Finding_1.Split('-')[4]).ToList();

            // Expect 4 findings to be created
            Assert.Equal(4, findings.Count);



            // Expect first ovs_finding to have the correct account and asset references
            var first = findings[0];
            Assert.Equal(testAccount1Id, first.ts_accountid.Id);
            Assert.Equal(testOperation1Id, first.ts_operationid.Id);
            Assert.Equal(testSite1Id, first.ts_functionallocation.Id);
            Assert.Equal(testOperationType1Id, first.ts_ovs_operationtype.Id);

            // Expect second ovs_finding to have the correct account and asset references
            var second = findings[1];
            Assert.Equal(testAccount2Id, second.ts_accountid.Id);
            Assert.Equal(testOperation2Id, second.ts_operationid.Id);
            Assert.Equal(testSite2Id, second.ts_functionallocation.Id);
            Assert.Equal(testOperationType2Id, second.ts_ovs_operationtype.Id);

            // Expect third ovs_finding to have the correct account and asset references
            var third = findings[2];
            Assert.Equal(testAccount3Id, third.ts_accountid.Id);
            Assert.Equal(testOperation3Id, third.ts_operationid.Id);
            Assert.Equal(testSite3Id, third.ts_functionallocation.Id);
            Assert.Equal(testOperationType3Id, third.ts_ovs_operationtype.Id);

            // Expect fourth ovs_finding to have the correct account and asset references
            var fourth = findings[3];
            Assert.Equal(testAccount4Id, fourth.ts_accountid.Id);
            Assert.Equal(testOperation4Id, fourth.ts_operationid.Id);
            Assert.Equal(testSite4Id, fourth.ts_functionallocation.Id);
            Assert.Equal(testOperationType4Id, fourth.ts_ovs_operationtype.Id);
        }

        [Fact]
        public void When_ovs_questionnaireresponse_contains_findings_with_findingtypes_expect_findings_for_each_operation_to_be_created_with_findingtypes_set()
        {

            // ARRANGE
            var serviceAccountId = orgAdminUIService.Create(new Account() { Name = "Test Service Account" });
            var incidentId = orgAdminUIService.Create(new Incident());
            var workOrderId = orgAdminUIService.Create(new msdyn_workorder()
            {
                msdyn_name = "300-345678",
                msdyn_ServiceRequest = new EntityReference(Incident.EntityLogicalName, incidentId),
                msdyn_ServiceAccount = new EntityReference(Account.EntityLogicalName, serviceAccountId)
            });

            var testAccount1Id = orgAdminUIService.Create(new Account());

            var testAccountReference1 = new EntityReference(Account.EntityLogicalName, testAccount1Id);

            var testOperation1Id = orgAdminUIService.Create(new ovs_operation()
            {
                Id = new Guid("9de3a6e3-c4ad-eb11-8236-000d3ae8b866"),
                ts_stakeholder = testAccountReference1
            });

            orgAdminUIService.Create(new qm_rclegislation()
            {
                Id = new Guid("417d6241-386a-eb11-a812-000d3a84125e"),
                qm_name = "SATR 2",
                ts_NameEnglish = "SATR 2",
                ts_NameFrench = "RSDA 2",
            });
            orgAdminUIService.Create(new qm_rclegislation()
            {
                Id = new Guid("417d6241-386a-eb11-a812-000d3a84157e"),
                qm_name = "SATR 7 (1) (a)",
                ts_NameEnglish = "SATR 7 (1) (a)",
                ts_NameFrench = "RSDA 7 (1) (a)",
            });
            orgAdminUIService.Create(new qm_rclegislation()
            {
                Id = new Guid("417d6241-386a-eb11-a812-000d3a84126e"),
                qm_name = "SATR 2",
                ts_NameEnglish = "SATR 2 (b) (ii)",
                ts_NameFrench = "RSDA 2 (b) (ii)",
            });

            var workOrderServiceTaskId = orgAdminUIService.Create(new msdyn_workorderservicetask()
            {
                msdyn_name = "200-345678-1",
                msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId), // belongs to a work order
                msdyn_PercentComplete = 100.00,
                ovs_QuestionnaireResponse = @"
                {
                    ""finding-sq_141"": {
                        ""provisionReference"": ""SATR 2"",
                        ""provisionTextEn"": ""<strong>Verification of Identity</strong></br><strong><mark><mark>SATR 4</mark></mark></strong>: An air carrier must, at the boarding gate for an international flight, verify the identity of each passenger who appears to be 18 years of age or older using</br><ul style='list-style-type:none;'><li><strong>(a)</strong> one of the following pieces of photo identification issued by a government authority that shows the passenger’s surname, first name and any middle names, their date of birth and gender and that is valid:</li><ul style='list-style-type:none;'><li><strong>(i)</strong> a passport issued by the country of which the passenger is a citizen or a national,</li><li><strong>(ii)</strong> a NEXUS card,</li><li><strong>(iii)</strong> any document referred to in subsection 50(1) or 52(1) of the Immigration and Refugee Protection Regulations; or</li></ul><li><strong>(b)</strong> a valid restricted area identity card, as defined in section 3 of the Canadian Aviation Security Regulations, 2012.</li></ul>"",
                        ""provisionTextFr"": ""<strong>Verification of Identity</strong></br><strong><mark><mark><mark>SATR 4</mark></mark></mark></strong>: Tout transporteur aérien vérifie, à la porte d’embarquement pour un vol international, l’identité de chaque passager qui semble âgé de 18 ans ou plus au moyen :</br><ul style='list-style-type:none;'><li><strong>(a)</strong> soit de l’une des pièces d’identité avec photo ci-après qui est délivrée par une autorité gouvernementale, qui indique les nom et prénoms, date de naissance et genre du passager et qui est valide :</li><ul style='list-style-type:none;'><li><strong>(i)</strong> un passeport délivré au passager par le pays dont il est citoyen ou ressortissant,</li><li><strong>(ii)</strong> une carte NEXUS,</li><li><strong>(iii)</strong> un document visé au paragraphe 50(1) ou 52(1) du Règlement sur l’immigration et la protection des réfugiés;</li></ul><li><strong>(b)</strong> soit d’une carte d’identité de zone réglementée au sens de l’article 3 du Règlement canadien de 2012 sur la sûreté aérienne qui est valide.</li></ul>"",
                        ""comments"": ""comment text"",
                        ""operations"": [{""operationID"": ""9de3a6e3-c4ad-eb11-8236-000d3ae8b866"",""findingType"": ""717750000""}]
                    },
                    ""finding-sq_142"": {
                        ""provisionReference"": ""SATR 7 (1) (a)"",
                        ""provisionTextEn"": ""<strong>Verification of Identity</strong></br><strong><mark><mark>SATR 4</mark></mark></strong>: An air carrier must, at the boarding gate for an international flight, verify the identity of each passenger who appears to be 18 years of age or older using</br><ul style='list-style-type:none;'><li><strong>(a)</strong> one of the following pieces of photo identification issued by a government authority that shows the passenger’s surname, first name and any middle names, their date of birth and gender and that is valid:</li><ul style='list-style-type:none;'><li><strong>(i)</strong> a passport issued by the country of which the passenger is a citizen or a national,</li><li><strong>(ii)</strong> a NEXUS card,</li><li><strong>(iii)</strong> any document referred to in subsection 50(1) or 52(1) of the Immigration and Refugee Protection Regulations; or</li></ul><li><strong>(b)</strong> a valid restricted area identity card, as defined in section 3 of the Canadian Aviation Security Regulations, 2012.</li></ul>"",
                        ""provisionTextFr"": ""<strong>Verification of Identity</strong></br><strong><mark><mark><mark>SATR 4</mark></mark></mark></strong>: Tout transporteur aérien vérifie, à la porte d’embarquement pour un vol international, l’identité de chaque passager qui semble âgé de 18 ans ou plus au moyen :</br><ul style='list-style-type:none;'><li><strong>(a)</strong> soit de l’une des pièces d’identité avec photo ci-après qui est délivrée par une autorité gouvernementale, qui indique les nom et prénoms, date de naissance et genre du passager et qui est valide :</li><ul style='list-style-type:none;'><li><strong>(i)</strong> un passeport délivré au passager par le pays dont il est citoyen ou ressortissant,</li><li><strong>(ii)</strong> une carte NEXUS,</li><li><strong>(iii)</strong> un document visé au paragraphe 50(1) ou 52(1) du Règlement sur l’immigration et la protection des réfugiés;</li></ul><li><strong>(b)</strong> soit d’une carte d’identité de zone réglementée au sens de l’article 3 du Règlement canadien de 2012 sur la sûreté aérienne qui est valide.</li></ul>"",
                        ""comments"": ""comment text"",
                        ""operations"": [{""operationID"": ""9de3a6e3-c4ad-eb11-8236-000d3ae8b866"",""findingType"": ""717750000""}]
                    },
                    ""finding-sq_143"": {
                        ""provisionReference"": ""SATR 7 (1) (a)"",
                        ""provisionTextEn"": ""<strong>Verification of Identity</strong></br><strong><mark><mark>SATR 4</mark></mark></strong>: An air carrier must, at the boarding gate for an international flight, verify the identity of each passenger who appears to be 18 years of age or older using</br><ul style='list-style-type:none;'><li><strong>(a)</strong> one of the following pieces of photo identification issued by a government authority that shows the passenger’s surname, first name and any middle names, their date of birth and gender and that is valid:</li><ul style='list-style-type:none;'><li><strong>(i)</strong> a passport issued by the country of which the passenger is a citizen or a national,</li><li><strong>(ii)</strong> a NEXUS card,</li><li><strong>(iii)</strong> any document referred to in subsection 50(1) or 52(1) of the Immigration and Refugee Protection Regulations; or</li></ul><li><strong>(b)</strong> a valid restricted area identity card, as defined in section 3 of the Canadian Aviation Security Regulations, 2012.</li></ul>"",
                        ""provisionTextFr"": ""<strong>Verification of Identity</strong></br><strong><mark><mark><mark>SATR 4</mark></mark></mark></strong>: Tout transporteur aérien vérifie, à la porte d’embarquement pour un vol international, l’identité de chaque passager qui semble âgé de 18 ans ou plus au moyen :</br><ul style='list-style-type:none;'><li><strong>(a)</strong> soit de l’une des pièces d’identité avec photo ci-après qui est délivrée par une autorité gouvernementale, qui indique les nom et prénoms, date de naissance et genre du passager et qui est valide :</li><ul style='list-style-type:none;'><li><strong>(i)</strong> un passeport délivré au passager par le pays dont il est citoyen ou ressortissant,</li><li><strong>(ii)</strong> une carte NEXUS,</li><li><strong>(iii)</strong> un document visé au paragraphe 50(1) ou 52(1) du Règlement sur l’immigration et la protection des réfugiés;</li></ul><li><strong>(b)</strong> soit d’une carte d’identité de zone réglementée au sens de l’article 3 du Règlement canadien de 2012 sur la sûreté aérienne qui est valide.</li></ul>"",
                        ""comments"": ""comment text"",
                        ""operations"": [{""operationID"": ""9de3a6e3-c4ad-eb11-8236-000d3ae8b866"",""findingType"": ""717750001""}]
                    },
                    ""finding-sq_144"": {
                        ""provisionReference"": ""SATR 2 (b) (ii)"",
                        ""provisionTextEn"": ""<strong>Verification of Identity</strong></br><strong><mark><mark>SATR 4</mark></mark></strong>: An air carrier must, at the boarding gate for an international flight, verify the identity of each passenger who appears to be 18 years of age or older using</br><ul style='list-style-type:none;'><li><strong>(a)</strong> one of the following pieces of photo identification issued by a government authority that shows the passenger’s surname, first name and any middle names, their date of birth and gender and that is valid:</li><ul style='list-style-type:none;'><li><strong>(i)</strong> a passport issued by the country of which the passenger is a citizen or a national,</li><li><strong>(ii)</strong> a NEXUS card,</li><li><strong>(iii)</strong> any document referred to in subsection 50(1) or 52(1) of the Immigration and Refugee Protection Regulations; or</li></ul><li><strong>(b)</strong> a valid restricted area identity card, as defined in section 3 of the Canadian Aviation Security Regulations, 2012.</li></ul>"",
                        ""provisionTextFr"": ""<strong>Verification of Identity</strong></br><strong><mark><mark><mark>SATR 4</mark></mark></mark></strong>: Tout transporteur aérien vérifie, à la porte d’embarquement pour un vol international, l’identité de chaque passager qui semble âgé de 18 ans ou plus au moyen :</br><ul style='list-style-type:none;'><li><strong>(a)</strong> soit de l’une des pièces d’identité avec photo ci-après qui est délivrée par une autorité gouvernementale, qui indique les nom et prénoms, date de naissance et genre du passager et qui est valide :</li><ul style='list-style-type:none;'><li><strong>(i)</strong> un passeport délivré au passager par le pays dont il est citoyen ou ressortissant,</li><li><strong>(ii)</strong> une carte NEXUS,</li><li><strong>(iii)</strong> un document visé au paragraphe 50(1) ou 52(1) du Règlement sur l’immigration et la protection des réfugiés;</li></ul><li><strong>(b)</strong> soit d’une carte d’identité de zone réglementée au sens de l’article 3 du Règlement canadien de 2012 sur la sûreté aérienne qui est valide.</li></ul>"",
                        ""comments"": ""comment text"",
                        ""operations"": [{""operationID"": ""9de3a6e3-c4ad-eb11-8236-000d3ae8b866"",""findingType"": ""717750002""}]
                    }
                }
                "
            });

            // ACT

            orgAdminUIService.Update(new msdyn_workorderservicetask { Id = workOrderServiceTaskId });

            // ASSERT
            var query = new QueryExpression(ovs_Finding.EntityLogicalName)
            {
                ColumnSet = new ColumnSet("ovs_finding", "ts_findingtype", "ts_findingmappingkey")
            };
            var findings = orgAdminUIService.RetrieveMultiple(query).Entities.Cast<ovs_Finding>().OrderBy(f => f.ovs_Finding_1.Split('-')[3]).ToList();

            // Expect 4 findings to be created
            Assert.Equal(4, findings.Count);

            // Expect first ovs_finding to have an Undecided finding type
            var first = findings[0];
            Assert.Equal(ts_findingtype.Undecided, first.ts_findingtype);

            // Expect second ovs_finding to have an Undecided finding type
            var second = findings[1];
            Assert.Equal(ts_findingtype.Undecided, second.ts_findingtype);

            // Expect third ovs_finding to have an Observation finding type
            var third = findings[2];
            Assert.Equal(ts_findingtype.Observation, third.ts_findingtype);

            // Expect fourth ovs_finding to have an Non-compliance finding type
            var fourth = findings[3];
            Assert.Equal(ts_findingtype.Noncompliance, fourth.ts_findingtype);
        }

        [Fact]
        public void When_ovs_questionnaireresponse_contains_finding_but_work_order_service_task_is_not_100_percent_complete_expect_ovs_finding_not_created()
        {
            // ARRANGE
            var serviceAccountId = orgAdminUIService.Create(new Account() { Name = "Test Service Account" });
            var workOrderId = orgAdminUIService.Create(new msdyn_workorder()
            {
                msdyn_name = "300-345678",
                msdyn_ServiceRequest = null,
                msdyn_ServiceAccount = new EntityReference(Account.EntityLogicalName, serviceAccountId)

            });
            var workOrderServiceTaskId = orgAdminUIService.Create(new msdyn_workorderservicetask()
            {
                msdyn_name = "200-345678-1",
                msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId), // belongs to a work order
                msdyn_PercentComplete = 0.00,
            });
            var account = orgAdminUIService.Create(new Account());
            var accountReference = new EntityReference(Account.EntityLogicalName, account);
            var operation = orgAdminUIService.Create(new ovs_operation()
            {
                Id = new Guid("9de3a6e3-c4ad-eb11-8236-000d3ae8b866"),
                ts_stakeholder = accountReference
            });

            // ACT
            orgAdminUIService.Update(new msdyn_workorderservicetask
            {
                Id = workOrderServiceTaskId,
                msdyn_PercentComplete = 50.00,
                ovs_QuestionnaireResponse = @"
                {
                    ""finding-sq_142"": {
                        ""provisionReference"": ""SATR 4"",
                        ""provisionTextEn"": ""<strong>Verification of Identity</strong></br><strong><mark><mark>SATR 4</mark></mark></strong>: An air carrier must, at the boarding gate for an international flight, verify the identity of each passenger who appears to be 18 years of age or older using</br><ul style='list-style-type:none;'><li><strong>(a)</strong> one of the following pieces of photo identification issued by a government authority that shows the passenger’s surname, first name and any middle names, their date of birth and gender and that is valid:</li><ul style='list-style-type:none;'><li><strong>(i)</strong> a passport issued by the country of which the passenger is a citizen or a national,</li><li><strong>(ii)</strong> a NEXUS card,</li><li><strong>(iii)</strong> any document referred to in subsection 50(1) or 52(1) of the Immigration and Refugee Protection Regulations; or</li></ul><li><strong>(b)</strong> a valid restricted area identity card, as defined in section 3 of the Canadian Aviation Security Regulations, 2012.</li></ul>"",
                        ""provisionTextFr"": ""<strong>Verification of Identity</strong></br><strong><mark><mark><mark>SATR 4</mark></mark></mark></strong>: Tout transporteur aérien vérifie, à la porte d’embarquement pour un vol international, l’identité de chaque passager qui semble âgé de 18 ans ou plus au moyen :</br><ul style='list-style-type:none;'><li><strong>(a)</strong> soit de l’une des pièces d’identité avec photo ci-après qui est délivrée par une autorité gouvernementale, qui indique les nom et prénoms, date de naissance et genre du passager et qui est valide :</li><ul style='list-style-type:none;'><li><strong>(i)</strong> un passeport délivré au passager par le pays dont il est citoyen ou ressortissant,</li><li><strong>(ii)</strong> une carte NEXUS,</li><li><strong>(iii)</strong> un document visé au paragraphe 50(1) ou 52(1) du Règlement sur l’immigration et la protection des réfugiés;</li></ul><li><strong>(b)</strong> soit d’une carte d’identité de zone réglementée au sens de l’article 3 du Règlement canadien de 2012 sur la sûreté aérienne qui est valide.</li></ul>"",
                        ""comments"": ""new comments"",
                        ""operations"": ""[{""operationID"": ""9de3a6e3-c4ad-eb11-8236-000d3ae8b866"",""findingType"": ""717750000""}]""
                        ""documentaryEvidence"": ""C:\\fakepath\\newfile.png""
                    }
                }
                "
            });

            // ASSERT
            var query = new QueryExpression(ovs_Finding.EntityLogicalName)
            {
                ColumnSet = new ColumnSet()
            };
            var findings = orgAdminUIService.RetrieveMultiple(query).Entities.Cast<ovs_Finding>().ToList();

            // Expect no findings created because questionnaire is not completed
            Assert.Empty(findings);
        }

        [Fact]
        public void When_ovs_questionnaireresponse_contains_finding_expect_do_not_recreate_existing_ovs_finding()
        {

            // ARRANGE
            var serviceAccountId = orgAdminUIService.Create(new Account() { Name = "Test Service Account" });
            var incidentId = orgAdminUIService.Create(new Incident());
            var workOrderId = orgAdminUIService.Create(new msdyn_workorder()
            {
                msdyn_name = "300-345678",
                msdyn_ServiceRequest = new EntityReference(Incident.EntityLogicalName, incidentId),
                msdyn_ServiceAccount = new EntityReference(Account.EntityLogicalName, serviceAccountId)
            });
            var account = orgAdminUIService.Create(new Account());
            var accountReference = new EntityReference(Account.EntityLogicalName, account);
            var operation = orgAdminUIService.Create(new ovs_operation()
            {
                Id = new Guid("9de3a6e3-c4ad-eb11-8236-000d3ae8b866"),
                ts_stakeholder = accountReference
            });
            var workOrderServiceTaskId = orgAdminUIService.Create(new msdyn_workorderservicetask()
            {
                msdyn_name = "200-345678-1",
                msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId), // belongs to a work order
                msdyn_PercentComplete = 100.00,
                ovs_QuestionnaireResponse = @"
                {
                    ""finding-sq_142"": {
                        ""provisionReference"": ""SATR 4"",
                        ""provisionTextEn"": ""<strong>Verification of Identity</strong></br><strong><mark><mark>SATR 4</mark></mark></strong>: An air carrier must, at the boarding gate for an international flight, verify the identity of each passenger who appears to be 18 years of age or older using</br><ul style='list-style-type:none;'><li><strong>(a)</strong> one of the following pieces of photo identification issued by a government authority that shows the passenger’s surname, first name and any middle names, their date of birth and gender and that is valid:</li><ul style='list-style-type:none;'><li><strong>(i)</strong> a passport issued by the country of which the passenger is a citizen or a national,</li><li><strong>(ii)</strong> a NEXUS card,</li><li><strong>(iii)</strong> any document referred to in subsection 50(1) or 52(1) of the Immigration and Refugee Protection Regulations; or</li></ul><li><strong>(b)</strong> a valid restricted area identity card, as defined in section 3 of the Canadian Aviation Security Regulations, 2012.</li></ul>"",
                        ""provisionTextFr"": ""<strong>Verification of Identity</strong></br><strong><mark><mark><mark>SATR 4</mark></mark></mark></strong>: Tout transporteur aérien vérifie, à la porte d’embarquement pour un vol international, l’identité de chaque passager qui semble âgé de 18 ans ou plus au moyen :</br><ul style='list-style-type:none;'><li><strong>(a)</strong> soit de l’une des pièces d’identité avec photo ci-après qui est délivrée par une autorité gouvernementale, qui indique les nom et prénoms, date de naissance et genre du passager et qui est valide :</li><ul style='list-style-type:none;'><li><strong>(i)</strong> un passeport délivré au passager par le pays dont il est citoyen ou ressortissant,</li><li><strong>(ii)</strong> une carte NEXUS,</li><li><strong>(iii)</strong> un document visé au paragraphe 50(1) ou 52(1) du Règlement sur l’immigration et la protection des réfugiés;</li></ul><li><strong>(b)</strong> soit d’une carte d’identité de zone réglementée au sens de l’article 3 du Règlement canadien de 2012 sur la sûreté aérienne qui est valide.</li></ul>"",
                        ""comments"": ""old comments"",
                        ""operations"": [{""operationID"": ""9de3a6e3-c4ad-eb11-8236-000d3ae8b866"",""findingType"": ""717750001""}]
                    }
                }
                "
            });
            var findingId = orgAdminUIService.Create(new ovs_Finding()
            {
                ovs_Finding_1 = "100-345678-1-3-1",
                ts_findingmappingkey = workOrderServiceTaskId + "-finding-sq_142-9de3a6e3-c4ad-eb11-8236-000d3ae8b866",
                ovs_FindingProvisionReference = "SATR 4",
                ovs_FindingComments = "original comments",
                ovs_FindingFile = "C:\\fakepath\\originalfile.png",
                ovs_CaseId = new EntityReference(Incident.EntityLogicalName, incidentId), // this finding already belongs to the case
                ovs_WorkOrderServiceTaskId = new EntityReference(msdyn_workorderservicetask.EntityLogicalName, workOrderServiceTaskId), // this finding already belongs to a work order service task
                statuscode = ovs_Finding_statuscode.New, // finding is also already active
                statecode = ovs_FindingState.Active, // finding is also already active
            });

            // ACT
            var newQuestionnaireResponse = @"
            {
                ""finding-sq_142"": {
                    ""provisionReference"": ""SATR 4"",
                    ""provisionTextEn"": ""<strong>Verification of Identity</strong></br><strong><mark><mark>SATR 4</mark></mark></strong>: An air carrier must, at the boarding gate for an international flight, verify the identity of each passenger who appears to be 18 years of age or older using</br><ul style='list-style-type:none;'><li><strong>(a)</strong> one of the following pieces of photo identification issued by a government authority that shows the passenger’s surname, first name and any middle names, their date of birth and gender and that is valid:</li><ul style='list-style-type:none;'><li><strong>(i)</strong> a passport issued by the country of which the passenger is a citizen or a national,</li><li><strong>(ii)</strong> a NEXUS card,</li><li><strong>(iii)</strong> any document referred to in subsection 50(1) or 52(1) of the Immigration and Refugee Protection Regulations; or</li></ul><li><strong>(b)</strong> a valid restricted area identity card, as defined in section 3 of the Canadian Aviation Security Regulations, 2012.</li></ul>"",
                    ""provisionTextFr"": ""<strong>Verification of Identity</strong></br><strong><mark><mark><mark>SATR 4</mark></mark></mark></strong>: Tout transporteur aérien vérifie, à la porte d’embarquement pour un vol international, l’identité de chaque passager qui semble âgé de 18 ans ou plus au moyen :</br><ul style='list-style-type:none;'><li><strong>(a)</strong> soit de l’une des pièces d’identité avec photo ci-après qui est délivrée par une autorité gouvernementale, qui indique les nom et prénoms, date de naissance et genre du passager et qui est valide :</li><ul style='list-style-type:none;'><li><strong>(i)</strong> un passeport délivré au passager par le pays dont il est citoyen ou ressortissant,</li><li><strong>(ii)</strong> une carte NEXUS,</li><li><strong>(iii)</strong> un document visé au paragraphe 50(1) ou 52(1) du Règlement sur l’immigration et la protection des réfugiés;</li></ul><li><strong>(b)</strong> soit d’une carte d’identité de zone réglementée au sens de l’article 3 du Règlement canadien de 2012 sur la sûreté aérienne qui est valide.</li></ul>"",
                    ""comments"": ""new comments"",
                    ""operations"": [{""operationID"": ""9de3a6e3-c4ad-eb11-8236-000d3ae8b866"",""findingType"": ""717750001""}]
                }
            }
            ";
            orgAdminUIService.Update(new msdyn_workorderservicetask { Id = workOrderServiceTaskId, ovs_QuestionnaireResponse = newQuestionnaireResponse });

            // ASSERT
            var query = new QueryExpression(ovs_Finding.EntityLogicalName)
            {
                ColumnSet = new ColumnSet()
            };
            var findings = orgAdminUIService.RetrieveMultiple(query).Entities.Cast<ovs_Finding>().ToList();

            // Expect target to still only contain the same finding
            Assert.Single(findings);
            Assert.Equal(findings.First().Id, findingId);
        }

        [Fact]
        public void When_ovs_questionnaireresponse_contains_finding_that_already_exists_expect_existing_ovs_finding_record_to_be_updated()
        {
            // ARRANGE
            var serviceAccountId = orgAdminUIService.Create(new Account() { Name = "Test Service Account" });
            var incidentId = orgAdminUIService.Create(new Incident());
            var workOrderId = orgAdminUIService.Create(new msdyn_workorder()
            {
                msdyn_name = "300-345678",
                msdyn_ServiceRequest = new EntityReference(Incident.EntityLogicalName, incidentId),
                msdyn_ServiceAccount = new EntityReference(Account.EntityLogicalName, serviceAccountId)
            });
            var account = orgAdminUIService.Create(new Account());
            var accountReference = new EntityReference(Account.EntityLogicalName, account);
            var operation = orgAdminUIService.Create(new ovs_operation()
            {
                Id = new Guid("9de3a6e3-c4ad-eb11-8236-000d3ae8b866"),
                ts_stakeholder = accountReference
            });

            orgAdminUIService.Create(new qm_rclegislation()
            {
                Id = new Guid("417d6241-386a-eb11-a812-000d3a84129e"),
                qm_name = "SATR 4",
                ts_NameEnglish = "SATR 4",
                ts_NameFrench = "RSDA 4",
            });

            var workOrderServiceTaskId = orgAdminUIService.Create(new msdyn_workorderservicetask()
            {
                msdyn_name = "200-345678-1",
                msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId), // belongs to a work order
                msdyn_PercentComplete = 100.00,
                ovs_QuestionnaireResponse = @"
                {
                    ""finding-sq_142"": {
                        ""provisionReference"": ""SATR 4"",
                        ""provisionTextEn"": ""<strong>Verification of Identity</strong></br><strong><mark><mark>SATR 4</mark></mark></strong>: An air carrier must, at the boarding gate for an international flight, verify the identity of each passenger who appears to be 18 years of age or older using</br><ul style='list-style-type:none;'><li><strong>(a)</strong> one of the following pieces of photo identification issued by a government authority that shows the passenger’s surname, first name and any middle names, their date of birth and gender and that is valid:</li><ul style='list-style-type:none;'><li><strong>(i)</strong> a passport issued by the country of which the passenger is a citizen or a national,</li><li><strong>(ii)</strong> a NEXUS card,</li><li><strong>(iii)</strong> any document referred to in subsection 50(1) or 52(1) of the Immigration and Refugee Protection Regulations; or</li></ul><li><strong>(b)</strong> a valid restricted area identity card, as defined in section 3 of the Canadian Aviation Security Regulations, 2012.</li></ul>"",
                        ""provisionTextFr"": ""<strong>Verification of Identity</strong></br><strong><mark><mark><mark>SATR 4</mark></mark></mark></strong>: Tout transporteur aérien vérifie, à la porte d’embarquement pour un vol international, l’identité de chaque passager qui semble âgé de 18 ans ou plus au moyen :</br><ul style='list-style-type:none;'><li><strong>(a)</strong> soit de l’une des pièces d’identité avec photo ci-après qui est délivrée par une autorité gouvernementale, qui indique les nom et prénoms, date de naissance et genre du passager et qui est valide :</li><ul style='list-style-type:none;'><li><strong>(i)</strong> un passeport délivré au passager par le pays dont il est citoyen ou ressortissant,</li><li><strong>(ii)</strong> une carte NEXUS,</li><li><strong>(iii)</strong> un document visé au paragraphe 50(1) ou 52(1) du Règlement sur l’immigration et la protection des réfugiés;</li></ul><li><strong>(b)</strong> soit d’une carte d’identité de zone réglementée au sens de l’article 3 du Règlement canadien de 2012 sur la sûreté aérienne qui est valide.</li></ul>"",
                        ""comments"": ""original comments"",
                        ""operations"": [{""operationID"": ""9de3a6e3-c4ad-eb11-8236-000d3ae8b866"",""findingType"": ""717750001""}]
                    }
                }
                "
            });

            // ACT
            var newQuestionnaireResponse = @"
            {
                ""finding-sq_142"": {
                    ""provisionReference"": ""SATR 4"",
                    ""provisionTextEn"": ""<strong>Verification of Identity</strong></br><strong><mark><mark>SATR 4</mark></mark></strong>: An air carrier must, at the boarding gate for an international flight, verify the identity of each passenger who appears to be 18 years of age or older using</br><ul style='list-style-type:none;'><li><strong>(a)</strong> one of the following pieces of photo identification issued by a government authority that shows the passenger’s surname, first name and any middle names, their date of birth and gender and that is valid:</li><ul style='list-style-type:none;'><li><strong>(i)</strong> a passport issued by the country of which the passenger is a citizen or a national,</li><li><strong>(ii)</strong> a NEXUS card,</li><li><strong>(iii)</strong> any document referred to in subsection 50(1) or 52(1) of the Immigration and Refugee Protection Regulations; or</li></ul><li><strong>(b)</strong> a valid restricted area identity card, as defined in section 3 of the Canadian Aviation Security Regulations, 2012.</li></ul>"",
                    ""provisionTextFr"": ""<strong>Verification of Identity</strong></br><strong><mark><mark><mark>SATR 4</mark></mark></mark></strong>: Tout transporteur aérien vérifie, à la porte d’embarquement pour un vol international, l’identité de chaque passager qui semble âgé de 18 ans ou plus au moyen :</br><ul style='list-style-type:none;'><li><strong>(a)</strong> soit de l’une des pièces d’identité avec photo ci-après qui est délivrée par une autorité gouvernementale, qui indique les nom et prénoms, date de naissance et genre du passager et qui est valide :</li><ul style='list-style-type:none;'><li><strong>(i)</strong> un passeport délivré au passager par le pays dont il est citoyen ou ressortissant,</li><li><strong>(ii)</strong> une carte NEXUS,</li><li><strong>(iii)</strong> un document visé au paragraphe 50(1) ou 52(1) du Règlement sur l’immigration et la protection des réfugiés;</li></ul><li><strong>(b)</strong> soit d’une carte d’identité de zone réglementée au sens de l’article 3 du Règlement canadien de 2012 sur la sûreté aérienne qui est valide.</li></ul>"",
                    ""comments"": ""new comments"",
                    ""operations"": [{""operationID"": ""9de3a6e3-c4ad-eb11-8236-000d3ae8b866"",""findingType"": ""717750001""}]
                }
            }
            ";
            orgAdminUIService.Update(new msdyn_workorderservicetask { Id = workOrderServiceTaskId, ovs_QuestionnaireResponse = newQuestionnaireResponse });

            // ASSERT
            var query = new QueryExpression(ovs_Finding.EntityLogicalName)
            {
                ColumnSet = new ColumnSet("ovs_findingcomments")
            };
            var findings = orgAdminUIService.RetrieveMultiple(query).Entities.Cast<ovs_Finding>().ToList();

            // Expect target to still only contain 1 ovs_finding reference
            Assert.Single(findings);

            // Expect first ovs_finding to have updated comments
            var first = findings[0];
            Assert.Equal("new comments", first.ovs_FindingComments);
        }

        [Fact]
        public void When_ovs_questionnaireresponse_contains_finding_that_already_exists_but_is_deactivated_expect_existing_ovs_finding_record_to_be_activated_in_case()
        {
            // ARRANGE
            var serviceAccountId = orgAdminUIService.Create(new Account() { Name = "Test Service Account" });
            var incidentId = orgAdminUIService.Create(new Incident());
            var workOrderId = orgAdminUIService.Create(new msdyn_workorder()
            {
                msdyn_name = "300-345678",
                msdyn_ServiceRequest = new EntityReference(Incident.EntityLogicalName, incidentId),
                msdyn_ServiceAccount = new EntityReference(Account.EntityLogicalName, serviceAccountId)

            });
            var workOrderServiceTaskId = orgAdminUIService.Create(new msdyn_workorderservicetask
            {
                msdyn_name = "200-345678-1",
                msdyn_PercentComplete = 100.00,
                msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId) // belongs to a work order
            });
            var account = orgAdminUIService.Create(new Account());
            var accountReference = new EntityReference(Account.EntityLogicalName, account);
            var operation = orgAdminUIService.Create(new ovs_operation()
            {
                Id = new Guid("9de3a6e3-c4ad-eb11-8236-000d3ae8b866"),
                ts_stakeholder = accountReference
            });
            var findingId = orgAdminUIService.Create(new ovs_Finding
            {
                ovs_Finding_1 = "100-345678-1-3-1",
                ts_findingmappingkey = workOrderServiceTaskId + "-finding-sq_142-" + "9de3a6e3-c4ad-eb11-8236-000d3ae8b866", // unique finding names are created using the work order service task ID and the reference ID in the questionnaire
                ovs_FindingProvisionReference = "SATR 4",
                ovs_FindingComments = "original comments",
                ovs_FindingFile = "C:\\fakepath\\originalfile.png",
                ovs_CaseId = new EntityReference(Incident.EntityLogicalName, incidentId), // this finding already belongs to the case
                ovs_WorkOrderServiceTaskId = new EntityReference(msdyn_workorderservicetask.EntityLogicalName, workOrderServiceTaskId), // this finding already belongs to a work order service task
                statuscode = ovs_Finding_statuscode.Inactive, // finding is also already deactivated
                statecode = ovs_FindingState.Inactive, // finding is also already deactivated
            });

            // ACT
            orgAdminUIService.Update(new msdyn_workorderservicetask
            {
                Id = workOrderServiceTaskId,
                ovs_QuestionnaireResponse = @"
                {
                    ""finding-sq_142"": {
                        ""provisionReference"": ""SATR 4"",
                        ""provisionTextEn"": ""<strong>Verification of Identity</strong></br><strong><mark><mark>SATR 4</mark></mark></strong>: An air carrier must, at the boarding gate for an international flight, verify the identity of each passenger who appears to be 18 years of age or older using</br><ul style='list-style-type:none;'><li><strong>(a)</strong> one of the following pieces of photo identification issued by a government authority that shows the passenger’s surname, first name and any middle names, their date of birth and gender and that is valid:</li><ul style='list-style-type:none;'><li><strong>(i)</strong> a passport issued by the country of which the passenger is a citizen or a national,</li><li><strong>(ii)</strong> a NEXUS card,</li><li><strong>(iii)</strong> any document referred to in subsection 50(1) or 52(1) of the Immigration and Refugee Protection Regulations; or</li></ul><li><strong>(b)</strong> a valid restricted area identity card, as defined in section 3 of the Canadian Aviation Security Regulations, 2012.</li></ul>"",
                        ""provisionTextFr"": ""<strong>Verification of Identity</strong></br><strong><mark><mark><mark>SATR 4</mark></mark></mark></strong>: Tout transporteur aérien vérifie, à la porte d’embarquement pour un vol international, l’identité de chaque passager qui semble âgé de 18 ans ou plus au moyen :</br><ul style='list-style-type:none;'><li><strong>(a)</strong> soit de l’une des pièces d’identité avec photo ci-après qui est délivrée par une autorité gouvernementale, qui indique les nom et prénoms, date de naissance et genre du passager et qui est valide :</li><ul style='list-style-type:none;'><li><strong>(i)</strong> un passeport délivré au passager par le pays dont il est citoyen ou ressortissant,</li><li><strong>(ii)</strong> une carte NEXUS,</li><li><strong>(iii)</strong> un document visé au paragraphe 50(1) ou 52(1) du Règlement sur l’immigration et la protection des réfugiés;</li></ul><li><strong>(b)</strong> soit d’une carte d’identité de zone réglementée au sens de l’article 3 du Règlement canadien de 2012 sur la sûreté aérienne qui est valide.</li></ul>"",
                        ""comments"": ""new comments"",
                        ""operations"": [{""operationID"": ""9de3a6e3-c4ad-eb11-8236-000d3ae8b866"",""findingType"": ""717750001""}]
                    }
                }
                "
            });

            // ASSERT
            var query = new QueryExpression(ovs_Finding.EntityLogicalName)
            {
                ColumnSet = new ColumnSet("statuscode", "statecode")
            };
            var findings = orgAdminUIService.RetrieveMultiple(query).Entities.Cast<ovs_Finding>().ToList();

            // Expect target to still only contain 1 ovs_finding reference
            Assert.Single(findings);

            // Expect first ovs_finding to have updated comments
            var first = findings[0];
            Assert.Equal(ovs_FindingState.Active, first.statecode);
            Assert.Equal(ovs_Finding_statuscode.New, first.statuscode);
        }

        [Fact]
        public void When_ovs_questionnaireresponse_no_longer_contains_finding_that_already_exists_expect_existing_ovs_finding_record_to_be_deactived_in_case()
        {
            // ARRANGE
            var serviceAccountId = orgAdminUIService.Create(new Account() { Name = "Test Service Account" });
            var incidentId = orgAdminUIService.Create(new Incident());
            var assetId = orgAdminUIService.Create(new ovs_operation());
            var workOrderId = orgAdminUIService.Create(new msdyn_workorder()
            {
                msdyn_name = "300-345678",
                msdyn_ServiceRequest = new EntityReference(Incident.EntityLogicalName, incidentId),
                msdyn_ServiceAccount = new EntityReference(Account.EntityLogicalName, serviceAccountId)

            });
            var workOrderServiceTaskId = orgAdminUIService.Create(new msdyn_workorderservicetask
            {
                msdyn_name = "200-345678-1",
                msdyn_PercentComplete = 100.00,
                msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId) // belongs to a work order
            });
            var findingId = orgAdminUIService.Create(new ovs_Finding
            {
                ovs_Finding_1 = "100-345678-1-3-1",
                ts_findingmappingkey = workOrderServiceTaskId + "-finding-sq_142-" + assetId.ToString(), // unique finding names are created using the work order service task ID and the reference ID in the questionnaire
                ovs_FindingProvisionReference = "SATR 4",
                ovs_FindingComments = "original comments",
                ovs_FindingFile = "C:\\fakepath\\originalfile.png",
                ovs_CaseId = new EntityReference(Incident.EntityLogicalName, incidentId), // this finding already belongs to the case
                ovs_WorkOrderServiceTaskId = new EntityReference(msdyn_workorderservicetask.EntityLogicalName, workOrderServiceTaskId), // this finding already belongs to a work order service task
                statuscode = ovs_Finding_statuscode.New, // finding is also already active
                statecode = ovs_FindingState.Active, // finding is also already active
            });

            // ACT
            orgAdminUIService.Update(new msdyn_workorderservicetask
            {
                Id = workOrderServiceTaskId,
                ovs_QuestionnaireResponse = @"
                {
                }
                "
            });

            // ASSERT
            var query = new QueryExpression(ovs_Finding.EntityLogicalName)
            {
                ColumnSet = new ColumnSet("statuscode", "statecode")
            };
            var findings = orgAdminUIService.RetrieveMultiple(query).Entities.Cast<ovs_Finding>().ToList();

            // Expect target to still only contain 1 ovs_finding reference
            Assert.Single(findings);

            // Expect first ovs_finding to have updated comments
            var first = findings[0];
            Assert.Equal(ovs_FindingState.Inactive, first.statecode);
            Assert.Equal(ovs_Finding_statuscode.Inactive, first.statuscode);
        }

        [Fact]
        public void When_ovs_questionnaireresponse_contains_finding_that_already_extists_with_copies_for_other_operations_expect_all_comments_of_copies_to_be_updated()
        {
            // ARRANGE
            var serviceAccountId = orgAdminUIService.Create(new Account() { Name = "Test Service Account" });
            var incidentId = orgAdminUIService.Create(new Incident());
            var workOrderId = orgAdminUIService.Create(new msdyn_workorder()
            {
                msdyn_name = "300-345678",
                msdyn_ServiceRequest = new EntityReference(Incident.EntityLogicalName, incidentId),
                msdyn_ServiceAccount = new EntityReference(Account.EntityLogicalName, serviceAccountId),

            });
            var workOrderServiceTaskId = orgAdminUIService.Create(new msdyn_workorderservicetask
            {
                msdyn_name = "200-345678-1",
                msdyn_PercentComplete = 100.00,
                msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId) // belongs to a work order
            });
            var account = orgAdminUIService.Create(new Account());
            var accountReference = new EntityReference(Account.EntityLogicalName, account);
            var operation = orgAdminUIService.Create(new ovs_operation()
            {
                Id = new Guid("9de3a6e3-c4ad-eb11-8236-000d3ae8b866"),
                ts_stakeholder = accountReference
            });
            var findingId1 = orgAdminUIService.Create(new ovs_Finding
            {
                ovs_Finding_1 = "100-345678-1-3-1",
                ts_findingmappingkey = workOrderServiceTaskId + "-finding-sq_142-9de3a6e3-c4ad-eb11-8236-000d3ae8b866", // unique finding names are created using the work order service task ID and the reference ID in the questionnaire
                ovs_FindingProvisionReference = "SATR 4",
                ovs_FindingComments = "original comments",
                ovs_FindingFile = "C:\\fakepath\\originalfile.png",
                ovs_CaseId = new EntityReference(Incident.EntityLogicalName, incidentId), // this finding already belongs to the case
                ovs_WorkOrderServiceTaskId = new EntityReference(msdyn_workorderservicetask.EntityLogicalName, workOrderServiceTaskId), // this finding already belongs to a work order service task
                statuscode = ovs_Finding_statuscode.New, // finding is also already active
                statecode = ovs_FindingState.Active, // finding is also already active
            });
            var findingId2 = orgAdminUIService.Create(new ovs_Finding
            {
                ovs_Finding_1 = "100-345678-1-3-2",
                ts_findingmappingkey = workOrderServiceTaskId + "-finding-sq_142-6b796de3-b3a4-eb11-9442-000d3a8410dc", // unique finding names are created using the work order service task ID and the reference ID in the questionnaire
                ovs_FindingProvisionReference = "SATR 4",
                ovs_FindingComments = "original comments",
                ovs_FindingFile = "C:\\fakepath\\originalfile.png",
                ovs_CaseId = new EntityReference(Incident.EntityLogicalName, incidentId), // this finding already belongs to the case
                ovs_WorkOrderServiceTaskId = new EntityReference(msdyn_workorderservicetask.EntityLogicalName, workOrderServiceTaskId), // this finding already belongs to a work order service task
                statuscode = ovs_Finding_statuscode.New, // finding is also already active
                statecode = ovs_FindingState.Active, // finding is also already active
            });
            var findingId3 = orgAdminUIService.Create(new ovs_Finding
            {
                ovs_Finding_1 = "100-345678-1-3-3",
                ts_findingmappingkey = workOrderServiceTaskId + "-finding-sq_142-7d085d54-c2a9-eb11-9442-000d3a8410dc", // unique finding names are created using the work order service task ID and the reference ID in the questionnaire
                ovs_FindingProvisionReference = "SATR 4",
                ovs_FindingComments = "original comments",
                ovs_FindingFile = "C:\\fakepath\\originalfile.png",
                ovs_CaseId = new EntityReference(Incident.EntityLogicalName, incidentId), // this finding already belongs to the case
                ovs_WorkOrderServiceTaskId = new EntityReference(msdyn_workorderservicetask.EntityLogicalName, workOrderServiceTaskId), // this finding already belongs to a work order service task
                statuscode = ovs_Finding_statuscode.New, // finding is also already active
                statecode = ovs_FindingState.Active, // finding is also already active
            });
            // ACT
            orgAdminUIService.Update(new msdyn_workorderservicetask
            {
                Id = workOrderServiceTaskId,
                ovs_QuestionnaireResponse = @"
                {
                    ""finding-sq_142"": {
                        ""provisionReference"": ""SATR 4"",
                        ""provisionTextEn"": ""<strong>Verification of Identity</strong></br><strong><mark><mark>SATR 4</mark></mark></strong>: An air carrier must, at the boarding gate for an international flight, verify the identity of each passenger who appears to be 18 years of age or older using</br><ul style='list-style-type:none;'><li><strong>(a)</strong> one of the following pieces of photo identification issued by a government authority that shows the passenger’s surname, first name and any middle names, their date of birth and gender and that is valid:</li><ul style='list-style-type:none;'><li><strong>(i)</strong> a passport issued by the country of which the passenger is a citizen or a national,</li><li><strong>(ii)</strong> a NEXUS card,</li><li><strong>(iii)</strong> any document referred to in subsection 50(1) or 52(1) of the Immigration and Refugee Protection Regulations; or</li></ul><li><strong>(b)</strong> a valid restricted area identity card, as defined in section 3 of the Canadian Aviation Security Regulations, 2012.</li></ul>"",
                        ""provisionTextFr"": ""<strong>Verification of Identity</strong></br><strong><mark><mark><mark>SATR 4</mark></mark></mark></strong>: Tout transporteur aérien vérifie, à la porte d’embarquement pour un vol international, l’identité de chaque passager qui semble âgé de 18 ans ou plus au moyen :</br><ul style='list-style-type:none;'><li><strong>(a)</strong> soit de l’une des pièces d’identité avec photo ci-après qui est délivrée par une autorité gouvernementale, qui indique les nom et prénoms, date de naissance et genre du passager et qui est valide :</li><ul style='list-style-type:none;'><li><strong>(i)</strong> un passeport délivré au passager par le pays dont il est citoyen ou ressortissant,</li><li><strong>(ii)</strong> une carte NEXUS,</li><li><strong>(iii)</strong> un document visé au paragraphe 50(1) ou 52(1) du Règlement sur l’immigration et la protection des réfugiés;</li></ul><li><strong>(b)</strong> soit d’une carte d’identité de zone réglementée au sens de l’article 3 du Règlement canadien de 2012 sur la sûreté aérienne qui est valide.</li></ul>"",
                        ""comments"": ""new comments"",
                        ""operations"": [
	                        {""operationID"": ""9de3a6e3-c4ad-eb11-8236-000d3ae8b866"",""findingType"": ""717750001""},
	                        {""operationID"": ""6b796de3-b3a4-eb11-9442-000d3a8410dc"",""findingType"": ""717750001""},
	                        {""operationID"": ""7d085d54-c2a9-eb11-9442-000d3a8410dc"",""findingType"": ""717750001""}
                        ]
                    }
                }
                "
            });
            // ASSERT
            var query = new QueryExpression(ovs_Finding.EntityLogicalName)
            {
                ColumnSet = new ColumnSet("ovs_findingcomments")
            };
            var findings = orgAdminUIService.RetrieveMultiple(query).Entities.Cast<ovs_Finding>().ToList();

            // Expect target to still only contain 3 ovs_finding references
            Assert.Equal(3, findings.Count);

            // Expect first ovs_finding to have updated comments
            var first = findings[0];
            Assert.Equal("new comments", first.ovs_FindingComments);

            // Expect second ovs_finding to have updated comments
            var second = findings[1];
            Assert.Equal("new comments", second.ovs_FindingComments);

            // Expect third ovs_finding to have updated comments
            var third = findings[2];
            Assert.Equal("new comments", third.ovs_FindingComments);
        }

        [Fact]
        public void When_ovs_questionnaireresponse_contains_finding_that_already_extists_with_a_findingtype_expect_findingtype_to_be_updated()
        {
            // ARRANGE
            var serviceAccountId = orgAdminUIService.Create(new Account() { Name = "Test Service Account" });
            var incidentId = orgAdminUIService.Create(new Incident());
            var workOrderId = orgAdminUIService.Create(new msdyn_workorder()
            {
                msdyn_name = "300-345678",
                msdyn_ServiceRequest = new EntityReference(Incident.EntityLogicalName, incidentId),
                msdyn_ServiceAccount = new EntityReference(Account.EntityLogicalName, serviceAccountId),

            });
            var workOrderServiceTaskId = orgAdminUIService.Create(new msdyn_workorderservicetask
            {
                msdyn_name = "200-345678-1",
                msdyn_PercentComplete = 100.00,
                msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId) // belongs to a work order
            });
            var account = orgAdminUIService.Create(new Account());
            var accountReference = new EntityReference(Account.EntityLogicalName, account);
            var operation = orgAdminUIService.Create(new ovs_operation()
            {
                Id = new Guid("9de3a6e3-c4ad-eb11-8236-000d3ae8b866"),
                ts_stakeholder = accountReference
            });
            var findingId1 = orgAdminUIService.Create(new ovs_Finding
            {
                ovs_Finding_1 = "100-345678-1-3-1",
                ts_findingmappingkey = workOrderServiceTaskId + "-finding-sq_142-9de3a6e3-c4ad-eb11-8236-000d3ae8b866", // unique finding names are created using the work order service task ID and the reference ID in the questionnaire
                ovs_FindingProvisionReference = "SATR 4",
                ovs_FindingComments = "original comments",
                ovs_CaseId = new EntityReference(Incident.EntityLogicalName, incidentId), // this finding already belongs to the case
                ovs_WorkOrderServiceTaskId = new EntityReference(msdyn_workorderservicetask.EntityLogicalName, workOrderServiceTaskId), // this finding already belongs to a work order service task
                statuscode = ovs_Finding_statuscode.New, // finding is also already active
                statecode = ovs_FindingState.Active, // finding is also already active
                ts_findingtype = ts_findingtype.Undecided
            });
            // ACT
            orgAdminUIService.Update(new msdyn_workorderservicetask
            {
                Id = workOrderServiceTaskId,
                ovs_QuestionnaireResponse = @"
                {
                    ""finding-sq_142"": {
                        ""provisionReference"": ""SATR 4"",
                        ""provisionTextEn"": ""<strong>Verification of Identity</strong></br><strong><mark><mark>SATR 4</mark></mark></strong>: An air carrier must, at the boarding gate for an international flight, verify the identity of each passenger who appears to be 18 years of age or older using</br><ul style='list-style-type:none;'><li><strong>(a)</strong> one of the following pieces of photo identification issued by a government authority that shows the passenger’s surname, first name and any middle names, their date of birth and gender and that is valid:</li><ul style='list-style-type:none;'><li><strong>(i)</strong> a passport issued by the country of which the passenger is a citizen or a national,</li><li><strong>(ii)</strong> a NEXUS card,</li><li><strong>(iii)</strong> any document referred to in subsection 50(1) or 52(1) of the Immigration and Refugee Protection Regulations; or</li></ul><li><strong>(b)</strong> a valid restricted area identity card, as defined in section 3 of the Canadian Aviation Security Regulations, 2012.</li></ul>"",
                        ""provisionTextFr"": ""<strong>Verification of Identity</strong></br><strong><mark><mark><mark>SATR 4</mark></mark></mark></strong>: Tout transporteur aérien vérifie, à la porte d’embarquement pour un vol international, l’identité de chaque passager qui semble âgé de 18 ans ou plus au moyen :</br><ul style='list-style-type:none;'><li><strong>(a)</strong> soit de l’une des pièces d’identité avec photo ci-après qui est délivrée par une autorité gouvernementale, qui indique les nom et prénoms, date de naissance et genre du passager et qui est valide :</li><ul style='list-style-type:none;'><li><strong>(i)</strong> un passeport délivré au passager par le pays dont il est citoyen ou ressortissant,</li><li><strong>(ii)</strong> une carte NEXUS,</li><li><strong>(iii)</strong> un document visé au paragraphe 50(1) ou 52(1) du Règlement sur l’immigration et la protection des réfugiés;</li></ul><li><strong>(b)</strong> soit d’une carte d’identité de zone réglementée au sens de l’article 3 du Règlement canadien de 2012 sur la sûreté aérienne qui est valide.</li></ul>"",
                        ""comments"": ""new comments"",
                        ""operations"": [{""operationID"": ""9de3a6e3-c4ad-eb11-8236-000d3ae8b866"",""findingType"": ""717750001""}]
                    }
                }
                "
            });
            // ASSERT
            var query = new QueryExpression(ovs_Finding.EntityLogicalName)
            {
                ColumnSet = new ColumnSet("ovs_findingcomments", "ts_findingtype")
            };
            var findings = orgAdminUIService.RetrieveMultiple(query).Entities.Cast<ovs_Finding>().ToList();

            // Expect target to still only contain 1 ovs_finding references
            Assert.Single(findings);

            // Expect first ovs_finding to have updated findingType
            var first = findings[0];
            Assert.Equal(ts_findingtype.Observation, first.ts_findingtype);
        }

        [Fact]
        public void When_ovs_questionnaireresponse_contains_finding_with_missing_optional_values_expect_ovs_finding_record_to_still_be_created()
        {
            // ARRANGE
            var serviceAccountId = orgAdminUIService.Create(new Account() { Name = "Test Service Account" });
            var workOrderId = orgAdminUIService.Create(new msdyn_workorder()
            {
                msdyn_name = "300-345678",
                msdyn_ServiceRequest = null,
                msdyn_ServiceAccount = new EntityReference(Account.EntityLogicalName, serviceAccountId)

            });
            var workOrderServiceTaskId = orgAdminUIService.Create(new msdyn_workorderservicetask()
            {
                msdyn_name = "200-345678-1",
                msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId), // belongs to a work order
                msdyn_PercentComplete = 0.00,
            });
            var account = orgAdminUIService.Create(new Account());
            var accountReference = new EntityReference(Account.EntityLogicalName, account);
            var operation = orgAdminUIService.Create(new ovs_operation()
            {
                Id = new Guid("9de3a6e3-c4ad-eb11-8236-000d3ae8b866"),
                ts_stakeholder = accountReference
            });

            orgAdminUIService.Create(new qm_rclegislation()
            {
                Id = new Guid("417d6241-386a-eb11-a812-000d3a84129e"),
                qm_name = "SATR 4",
                ts_NameEnglish = "SATR 4",
                ts_NameFrench = "RSDA 4",
            });
            // ACT
            orgAdminUIService.Update(new msdyn_workorderservicetask
            {
                Id = workOrderServiceTaskId,
                msdyn_PercentComplete = 100.00,
                ovs_QuestionnaireResponse = @"
                {
                    ""finding-sq_142"": {
                        ""provisionReference"": ""SATR 4"",
                        ""provisionTextEn"": ""<strong>Verification of Identity</strong></br><strong><mark><mark>SATR 4</mark></mark></strong>: An air carrier must, at the boarding gate for an international flight, verify the identity of each passenger who appears to be 18 years of age or older using</br><ul style='list-style-type:none;'><li><strong>(a)</strong> one of the following pieces of photo identification issued by a government authority that shows the passenger’s surname, first name and any middle names, their date of birth and gender and that is valid:</li><ul style='list-style-type:none;'><li><strong>(i)</strong> a passport issued by the country of which the passenger is a citizen or a national,</li><li><strong>(ii)</strong> a NEXUS card,</li><li><strong>(iii)</strong> any document referred to in subsection 50(1) or 52(1) of the Immigration and Refugee Protection Regulations; or</li></ul><li><strong>(b)</strong> a valid restricted area identity card, as defined in section 3 of the Canadian Aviation Security Regulations, 2012.</li></ul>"",
                        ""provisionTextFr"": ""<strong>Verification of Identity</strong></br><strong><mark><mark><mark>SATR 4</mark></mark></mark></strong>: Tout transporteur aérien vérifie, à la porte d’embarquement pour un vol international, l’identité de chaque passager qui semble âgé de 18 ans ou plus au moyen :</br><ul style='list-style-type:none;'><li><strong>(a)</strong> soit de l’une des pièces d’identité avec photo ci-après qui est délivrée par une autorité gouvernementale, qui indique les nom et prénoms, date de naissance et genre du passager et qui est valide :</li><ul style='list-style-type:none;'><li><strong>(i)</strong> un passeport délivré au passager par le pays dont il est citoyen ou ressortissant,</li><li><strong>(ii)</strong> une carte NEXUS,</li><li><strong>(iii)</strong> un document visé au paragraphe 50(1) ou 52(1) du Règlement sur l’immigration et la protection des réfugiés;</li></ul><li><strong>(b)</strong> soit d’une carte d’identité de zone réglementée au sens de l’article 3 du Règlement canadien de 2012 sur la sûreté aérienne qui est valide.</li></ul>"",
                        ""operations"": [{""operationID"": ""9de3a6e3-c4ad-eb11-8236-000d3ae8b866"",""findingType"": ""717750001""}]
                    }
                }
                "
            });

            // ASSERT
            var query = new QueryExpression(ovs_Finding.EntityLogicalName)
            {
                ColumnSet = new ColumnSet()
            };
            var findings = orgAdminUIService.RetrieveMultiple(query).Entities.Cast<ovs_Finding>().ToList();

            // Expect exactly one finding
            Assert.Single(findings);
        }

        [Fact]
        public void When_ovs_questionnaireresponse_contains_non_compliance_finding_expect_msdyn_inspectiontaskresult_fail()
        {
            // ARRANGE
            var serviceAccountId = orgAdminUIService.Create(new Account() { Name = "Test Service Account" });
            var workOrderId = orgAdminUIService.Create(new msdyn_workorder()
            {
                msdyn_name = "300-345678",
                msdyn_ServiceRequest = null,
                msdyn_ServiceAccount = new EntityReference(Account.EntityLogicalName, serviceAccountId)

            });
            var workOrderServiceTaskId = orgAdminUIService.Create(new msdyn_workorderservicetask()
            {
                msdyn_name = "200-345678-1",
                msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId), // belongs to a work order
                msdyn_PercentComplete = 0.00,
            });
            var account = orgAdminUIService.Create(new Account());
            var accountReference1 = new EntityReference(Account.EntityLogicalName, account);
            var accountReference2 = new EntityReference(Account.EntityLogicalName, account);
            var accountReference3 = new EntityReference(Account.EntityLogicalName, account);

            var operation1 = orgAdminUIService.Create(new ovs_operation()
            {
                Id = new Guid("9de3a6e3-c4ad-eb11-8236-000d3ae8b866"),
                ts_stakeholder = accountReference1
            });
            var operation2 = orgAdminUIService.Create(new ovs_operation()
            {
                Id = new Guid("7d085d54-c2a9-eb11-9442-000d3a8410dc"),
                ts_stakeholder = accountReference2
            });
            var operation3 = orgAdminUIService.Create(new ovs_operation()
            {
                Id = new Guid("22364b7e-e1ce-eb11-bacc-0022483c068d"),
                ts_stakeholder = accountReference3
            });

            orgAdminUIService.Create(new qm_rclegislation()
            {
                Id = new Guid("417d6241-386a-eb11-a812-000d3a84129e"),
                qm_name = "SATR 4",
                ts_NameEnglish = "SATR 4",
                ts_NameFrench = "RSDA 4",
            });

            orgAdminUIService.Create(new qm_rclegislation()
            {
                Id = new Guid("417d6241-386a-eb11-a812-000d3a84128e"),
                qm_name = "SATR 5",
                ts_NameEnglish = "SATR 5",
                ts_NameFrench = "RSDA 5",
            });

            // ACT
            orgAdminUIService.Update(new msdyn_workorderservicetask
            {
                Id = workOrderServiceTaskId,
                msdyn_PercentComplete = 100.00,
                ovs_QuestionnaireResponse = @"
                {
                    ""finding-sq_142"": {
                        ""provisionReference"": ""SATR 4"",
                        ""provisionTextEn"": ""<strong>Verification of Identity</strong></br><strong><mark><mark>SATR 4</mark></mark></strong>: An air carrier must, at the boarding gate for an international flight, verify the identity of each passenger who appears to be 18 years of age or older using</br><ul style='list-style-type:none;'><li><strong>(a)</strong> one of the following pieces of photo identification issued by a government authority that shows the passenger’s surname, first name and any middle names, their date of birth and gender and that is valid:</li><ul style='list-style-type:none;'><li><strong>(i)</strong> a passport issued by the country of which the passenger is a citizen or a national,</li><li><strong>(ii)</strong> a NEXUS card,</li><li><strong>(iii)</strong> any document referred to in subsection 50(1) or 52(1) of the Immigration and Refugee Protection Regulations; or</li></ul><li><strong>(b)</strong> a valid restricted area identity card, as defined in section 3 of the Canadian Aviation Security Regulations, 2012.</li></ul>"",
                        ""provisionTextFr"": ""<strong>Verification of Identity</strong></br><strong><mark><mark><mark>SATR 4</mark></mark></mark></strong>: Tout transporteur aérien vérifie, à la porte d’embarquement pour un vol international, l’identité de chaque passager qui semble âgé de 18 ans ou plus au moyen :</br><ul style='list-style-type:none;'><li><strong>(a)</strong> soit de l’une des pièces d’identité avec photo ci-après qui est délivrée par une autorité gouvernementale, qui indique les nom et prénoms, date de naissance et genre du passager et qui est valide :</li><ul style='list-style-type:none;'><li><strong>(i)</strong> un passeport délivré au passager par le pays dont il est citoyen ou ressortissant,</li><li><strong>(ii)</strong> une carte NEXUS,</li><li><strong>(iii)</strong> un document visé au paragraphe 50(1) ou 52(1) du Règlement sur l’immigration et la protection des réfugiés;</li></ul><li><strong>(b)</strong> soit d’une carte d’identité de zone réglementée au sens de l’article 3 du Règlement canadien de 2012 sur la sûreté aérienne qui est valide.</li></ul>"",
                        ""comments"": ""new comments"",
                        ""operations"": [{""operationID"": ""9de3a6e3-c4ad-eb11-8236-000d3ae8b866"",""findingType"": ""717750001""},
                                         {""operationID"": ""7d085d54-c2a9-eb11-9442-000d3a8410dc"",""findingType"": ""717750001""}
                        ]                                      
                    },
                    ""finding-sq_155"": {
                        ""provisionReference"": ""SATR 5"",
                        ""provisionTextEn"": ""<strong>Verification of Identity</strong></br><strong><mark><mark>SATR 4</mark></mark></strong>: An air carrier must, at the boarding gate for an international flight, verify the identity of each passenger who appears to be 18 years of age or older using</br><ul style='list-style-type:none;'><li><strong>(a)</strong> one of the following pieces of photo identification issued by a government authority that shows the passenger’s surname, first name and any middle names, their date of birth and gender and that is valid:</li><ul style='list-style-type:none;'><li><strong>(i)</strong> a passport issued by the country of which the passenger is a citizen or a national,</li><li><strong>(ii)</strong> a NEXUS card,</li><li><strong>(iii)</strong> any document referred to in subsection 50(1) or 52(1) of the Immigration and Refugee Protection Regulations; or</li></ul><li><strong>(b)</strong> a valid restricted area identity card, as defined in section 3 of the Canadian Aviation Security Regulations, 2012.</li></ul>"",
                        ""provisionTextFr"": ""<strong>Verification of Identity</strong></br><strong><mark><mark><mark>SATR 4</mark></mark></mark></strong>: Tout transporteur aérien vérifie, à la porte d’embarquement pour un vol international, l’identité de chaque passager qui semble âgé de 18 ans ou plus au moyen :</br><ul style='list-style-type:none;'><li><strong>(a)</strong> soit de l’une des pièces d’identité avec photo ci-après qui est délivrée par une autorité gouvernementale, qui indique les nom et prénoms, date de naissance et genre du passager et qui est valide :</li><ul style='list-style-type:none;'><li><strong>(i)</strong> un passeport délivré au passager par le pays dont il est citoyen ou ressortissant,</li><li><strong>(ii)</strong> une carte NEXUS,</li><li><strong>(iii)</strong> un document visé au paragraphe 50(1) ou 52(1) du Règlement sur l’immigration et la protection des réfugiés;</li></ul><li><strong>(b)</strong> soit d’une carte d’identité de zone réglementée au sens de l’article 3 du Règlement canadien de 2012 sur la sûreté aérienne qui est valide.</li></ul>"",
                        ""comments"": ""new comment 2"",
                        ""operations"": [	                       
	                        {""operationID"": ""22364b7e-e1ce-eb11-bacc-0022483c068d"",""findingType"": ""717750002""}
                        ]
                    }
                }
                "
            });

            // ASSERT
            var query = new QueryExpression(msdyn_workorderservicetask.EntityLogicalName)
            {
                ColumnSet = new ColumnSet("msdyn_inspectiontaskresult")
            };
            var workOrderServiceTasks = orgAdminUIService.RetrieveMultiple(query).Entities.Cast<msdyn_workorderservicetask>().ToList();

            // Expect exactly one finding
            Assert.Single(workOrderServiceTasks);
            Assert.Equal(msdyn_inspectionresult.Fail, workOrderServiceTasks.Single().msdyn_inspectiontaskresult);
        }

        [Fact]
        public void When_ovs_questionnaireresponse_contains_observations_only_finding_expect_msdyn_inspectiontaskresult_observation()
        {
            // ARRANGE
            var serviceAccountId = orgAdminUIService.Create(new Account() { Name = "Test Service Account" });
            var workOrderId = orgAdminUIService.Create(new msdyn_workorder()
            {
                msdyn_name = "300-345678",
                msdyn_ServiceRequest = null,
                msdyn_ServiceAccount = new EntityReference(Account.EntityLogicalName, serviceAccountId)

            });
            var workOrderServiceTaskId = orgAdminUIService.Create(new msdyn_workorderservicetask()
            {
                msdyn_name = "200-345678-1",
                msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId), // belongs to a work order
                msdyn_PercentComplete = 0.00,
            });
            var account = orgAdminUIService.Create(new Account());
            var accountReference1 = new EntityReference(Account.EntityLogicalName, account);
            var accountReference2 = new EntityReference(Account.EntityLogicalName, account);
            var accountReference3 = new EntityReference(Account.EntityLogicalName, account);

            var operation1 = orgAdminUIService.Create(new ovs_operation()
            {
                Id = new Guid("9de3a6e3-c4ad-eb11-8236-000d3ae8b866"),
                ts_stakeholder = accountReference1
            });
            var operation2 = orgAdminUIService.Create(new ovs_operation()
            {
                Id = new Guid("7d085d54-c2a9-eb11-9442-000d3a8410dc"),
                ts_stakeholder = accountReference2
            });
            var operation3 = orgAdminUIService.Create(new ovs_operation()
            {
                Id = new Guid("22364b7e-e1ce-eb11-bacc-0022483c068d"),
                ts_stakeholder = accountReference3
            });

            orgAdminUIService.Create(new qm_rclegislation()
            {
                Id = new Guid("417d6241-386a-eb11-a812-000d3a84129e"),
                qm_name = "SATR 4",
                ts_NameEnglish = "SATR 4",
                ts_NameFrench = "RSDA 4",
            });

            orgAdminUIService.Create(new qm_rclegislation()
            {
                Id = new Guid("517d6241-386a-eb11-a812-000d3a84129e"),
                qm_name = "SATR 5",
                ts_NameEnglish = "SATR 5",
                ts_NameFrench = "RSDA 5",
            });

            // ACT
            orgAdminUIService.Update(new msdyn_workorderservicetask
            {
                Id = workOrderServiceTaskId,
                msdyn_PercentComplete = 100.00,
                ovs_QuestionnaireResponse = @"
                {
                    ""finding-sq_142"": {
                        ""provisionReference"": ""SATR 4"",
                        ""provisionTextEn"": ""<strong>Verification of Identity</strong></br><strong><mark><mark>SATR 4</mark></mark></strong>: An air carrier must, at the boarding gate for an international flight, verify the identity of each passenger who appears to be 18 years of age or older using</br><ul style='list-style-type:none;'><li><strong>(a)</strong> one of the following pieces of photo identification issued by a government authority that shows the passenger’s surname, first name and any middle names, their date of birth and gender and that is valid:</li><ul style='list-style-type:none;'><li><strong>(i)</strong> a passport issued by the country of which the passenger is a citizen or a national,</li><li><strong>(ii)</strong> a NEXUS card,</li><li><strong>(iii)</strong> any document referred to in subsection 50(1) or 52(1) of the Immigration and Refugee Protection Regulations; or</li></ul><li><strong>(b)</strong> a valid restricted area identity card, as defined in section 3 of the Canadian Aviation Security Regulations, 2012.</li></ul>"",
                        ""provisionTextFr"": ""<strong>Verification of Identity</strong></br><strong><mark><mark><mark>SATR 4</mark></mark></mark></strong>: Tout transporteur aérien vérifie, à la porte d’embarquement pour un vol international, l’identité de chaque passager qui semble âgé de 18 ans ou plus au moyen :</br><ul style='list-style-type:none;'><li><strong>(a)</strong> soit de l’une des pièces d’identité avec photo ci-après qui est délivrée par une autorité gouvernementale, qui indique les nom et prénoms, date de naissance et genre du passager et qui est valide :</li><ul style='list-style-type:none;'><li><strong>(i)</strong> un passeport délivré au passager par le pays dont il est citoyen ou ressortissant,</li><li><strong>(ii)</strong> une carte NEXUS,</li><li><strong>(iii)</strong> un document visé au paragraphe 50(1) ou 52(1) du Règlement sur l’immigration et la protection des réfugiés;</li></ul><li><strong>(b)</strong> soit d’une carte d’identité de zone réglementée au sens de l’article 3 du Règlement canadien de 2012 sur la sûreté aérienne qui est valide.</li></ul>"",
                        ""comments"": ""new comments"",
                        ""operations"": [{""operationID"": ""9de3a6e3-c4ad-eb11-8236-000d3ae8b866"",""findingType"": ""717750001""},
                                         {""operationID"": ""7d085d54-c2a9-eb11-9442-000d3a8410dc"",""findingType"": ""717750001""}
                        ]                                      
                    },
                    ""finding-sq_155"": {
                        ""provisionReference"": ""SATR 5"",
                        ""provisionTextEn"": ""<strong>Verification of Identity</strong></br><strong><mark><mark>SATR 4</mark></mark></strong>: An air carrier must, at the boarding gate for an international flight, verify the identity of each passenger who appears to be 18 years of age or older using</br><ul style='list-style-type:none;'><li><strong>(a)</strong> one of the following pieces of photo identification issued by a government authority that shows the passenger’s surname, first name and any middle names, their date of birth and gender and that is valid:</li><ul style='list-style-type:none;'><li><strong>(i)</strong> a passport issued by the country of which the passenger is a citizen or a national,</li><li><strong>(ii)</strong> a NEXUS card,</li><li><strong>(iii)</strong> any document referred to in subsection 50(1) or 52(1) of the Immigration and Refugee Protection Regulations; or</li></ul><li><strong>(b)</strong> a valid restricted area identity card, as defined in section 3 of the Canadian Aviation Security Regulations, 2012.</li></ul>"",
                        ""provisionTextFr"": ""<strong>Verification of Identity</strong></br><strong><mark><mark><mark>SATR 4</mark></mark></mark></strong>: Tout transporteur aérien vérifie, à la porte d’embarquement pour un vol international, l’identité de chaque passager qui semble âgé de 18 ans ou plus au moyen :</br><ul style='list-style-type:none;'><li><strong>(a)</strong> soit de l’une des pièces d’identité avec photo ci-après qui est délivrée par une autorité gouvernementale, qui indique les nom et prénoms, date de naissance et genre du passager et qui est valide :</li><ul style='list-style-type:none;'><li><strong>(i)</strong> un passeport délivré au passager par le pays dont il est citoyen ou ressortissant,</li><li><strong>(ii)</strong> une carte NEXUS,</li><li><strong>(iii)</strong> un document visé au paragraphe 50(1) ou 52(1) du Règlement sur l’immigration et la protection des réfugiés;</li></ul><li><strong>(b)</strong> soit d’une carte d’identité de zone réglementée au sens de l’article 3 du Règlement canadien de 2012 sur la sûreté aérienne qui est valide.</li></ul>"",
                        ""comments"": ""new comment 2"",
                        ""operations"": [	                       
	                        {""operationID"": ""22364b7e-e1ce-eb11-bacc-0022483c068d"",""findingType"": ""717750001""}
                        ]
                    }
                }
                "
            });

            // ASSERT
            var query = new QueryExpression(msdyn_workorderservicetask.EntityLogicalName)
            {
                ColumnSet = new ColumnSet("msdyn_inspectiontaskresult")
            };
            var workOrderServiceTasks = orgAdminUIService.RetrieveMultiple(query).Entities.Cast<msdyn_workorderservicetask>().ToList();

            // Expect exactly one finding
            Assert.Single(workOrderServiceTasks);
            Assert.Equal(msdyn_inspectionresult.Observations, workOrderServiceTasks.Single().msdyn_inspectiontaskresult);
        }

        [Fact]
        public void When_ovs_questionnaireresponse_contains_finding_expect_create_incident_and_ovs_finding_if_they_do_not_exist()
        {
            // ARRANGE
            var serviceAccountId = orgAdminUIService.Create(new Account() { Name = "Test Service Account" });
            var workOrderId = orgAdminUIService.Create(new msdyn_workorder()
            {
                msdyn_name = "300-345678",
                msdyn_ServiceRequest = null,
                msdyn_ServiceAccount = new EntityReference(Account.EntityLogicalName, serviceAccountId),
            });
            var workOrderServiceTaskId = orgAdminUIService.Create(new msdyn_workorderservicetask()
            {
                msdyn_name = "200-345678-1",
                msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId), // belongs to a work order
                msdyn_PercentComplete = 0.00,
            });
            var account = orgAdminUIService.Create(new Account());
            var accountReference = new EntityReference(Account.EntityLogicalName, account);
            var operation = orgAdminUIService.Create(new ovs_operation()
            {
                Id = new Guid("9de3a6e3-c4ad-eb11-8236-000d3ae8b866"),
                ts_stakeholder = accountReference
            });

            orgAdminUIService.Create(new qm_rclegislation()
            {
                Id = new Guid("417d6241-386a-eb11-a812-000d3a84129e"),
                qm_name = "SATR 4",
                ts_NameEnglish = "SATR 4",
                ts_NameFrench = "RSDA 4",
            });

            // ACT
            orgAdminUIService.Update(new msdyn_workorderservicetask
            {
                Id = workOrderServiceTaskId,
                msdyn_PercentComplete = 100.00,
                ovs_QuestionnaireResponse = @"
                {
                    ""finding-sq_142"": {
                        ""provisionReference"": ""SATR 4"",
                        ""provisionTextEn"": ""<strong>Verification of Identity</strong></br><strong><mark><mark>SATR 4</mark></mark></strong>: An air carrier must, at the boarding gate for an international flight, verify the identity of each passenger who appears to be 18 years of age or older using</br><ul style='list-style-type:none;'><li><strong>(a)</strong> one of the following pieces of photo identification issued by a government authority that shows the passenger’s surname, first name and any middle names, their date of birth and gender and that is valid:</li><ul style='list-style-type:none;'><li><strong>(i)</strong> a passport issued by the country of which the passenger is a citizen or a national,</li><li><strong>(ii)</strong> a NEXUS card,</li><li><strong>(iii)</strong> any document referred to in subsection 50(1) or 52(1) of the Immigration and Refugee Protection Regulations; or</li></ul><li><strong>(b)</strong> a valid restricted area identity card, as defined in section 3 of the Canadian Aviation Security Regulations, 2012.</li></ul>"",
                        ""provisionTextFr"": ""<strong>Verification of Identity</strong></br><strong><mark><mark><mark>SATR 4</mark></mark></mark></strong>: Tout transporteur aérien vérifie, à la porte d’embarquement pour un vol international, l’identité de chaque passager qui semble âgé de 18 ans ou plus au moyen :</br><ul style='list-style-type:none;'><li><strong>(a)</strong> soit de l’une des pièces d’identité avec photo ci-après qui est délivrée par une autorité gouvernementale, qui indique les nom et prénoms, date de naissance et genre du passager et qui est valide :</li><ul style='list-style-type:none;'><li><strong>(i)</strong> un passeport délivré au passager par le pays dont il est citoyen ou ressortissant,</li><li><strong>(ii)</strong> une carte NEXUS,</li><li><strong>(iii)</strong> un document visé au paragraphe 50(1) ou 52(1) du Règlement sur l’immigration et la protection des réfugiés;</li></ul><li><strong>(b)</strong> soit d’une carte d’identité de zone réglementée au sens de l’article 3 du Règlement canadien de 2012 sur la sûreté aérienne qui est valide.</li></ul>"",
                        ""comments"": ""new comments"",
                        ""operations"": [{""operationID"": ""9de3a6e3-c4ad-eb11-8236-000d3ae8b866"",""findingType"": ""717750001""}]
                    }
                }
                "
            });

            // ASSERT
            using (var context = new Xrm(orgAdminUIService))
            {
                var incidents = context.IncidentSet.ToList();
                var findings = context.ovs_FindingSet.ToList();

                // Expect exactly one case and one finding
                Assert.Single(incidents);
                Assert.Single(findings);
            }
        }

        [Fact]
        public void When_ovs_questionnaireresponse_does_not_contain_finding_expect_no_ovs_finding_record_created()
        {
            // ARRANGE
            var serviceAccountId = orgAdminUIService.Create(new Account() { Name = "Test Service Account" });
            var workOrderId = orgAdminUIService.Create(new msdyn_workorder()
            {
                msdyn_name = "300-345678",
                msdyn_ServiceRequest = null,
                msdyn_ServiceAccount = new EntityReference(Account.EntityLogicalName, serviceAccountId)
            });
            var workOrderServiceTaskId = orgAdminUIService.Create(new msdyn_workorderservicetask()
            {
                msdyn_name = "200-345678-1",
                msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId), // belongs to a work order
                msdyn_PercentComplete = 0.00,
            });

            // ACT
            orgAdminUIService.Update(new msdyn_workorderservicetask
            {
                Id = workOrderServiceTaskId,
                msdyn_PercentComplete = 100.00,
                ovs_QuestionnaireResponse = ""
            });

            // ASSERT
            using (var context = new Xrm(orgAdminUIService))
            {
                var findings = context.ovs_FindingSet.ToList();

                // Expect no findings
                Assert.Empty(findings);
            }
        }

        [Fact]
        public void When_ovs_questionnaireresponse_contains_finding_expect_legislation_reference_equals_ts_qm_rclegislation()
        {
            // ARRANGE
            var serviceAccountId = orgAdminUIService.Create(new Account() { Name = "Test Service Account" });
            var workOrderId = orgAdminUIService.Create(new msdyn_workorder()
            {
                msdyn_name = "300-345678",
                msdyn_ServiceRequest = null,
                msdyn_ServiceAccount = new EntityReference(Account.EntityLogicalName, serviceAccountId)

            });
            var workOrderServiceTaskId = orgAdminUIService.Create(new msdyn_workorderservicetask()
            {
                msdyn_name = "200-345678-1",
                msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId), // belongs to a work order
                msdyn_PercentComplete = 0.00,
            });
            var account = orgAdminUIService.Create(new Account());
            var accountReference = new EntityReference(Account.EntityLogicalName, account);

            Guid operationGuid = new Guid("cb84b80e-f2f6-eb11-94ef-000d3a09c1c3");
            var operation = orgAdminUIService.Create(new ovs_operation()
            {
                Id = operationGuid,
                ts_stakeholder = accountReference
            });

            Guid legislationGuid = new Guid("1de3a6e1-c2ad-eb11-1231-100d3ae8b861");
            var legislationid = orgAdminUIService.Create(new qm_rclegislation()
            {
                Id = legislationGuid,
                qm_name = "SATR 4.1 (1) (b)",
                ts_NameEnglish = "SATR 4.1 (1) (b)",
                ts_NameFrench = "RSDA 4.1 (1) (b)",
            });

            // ACT
            orgAdminUIService.Update(new msdyn_workorderservicetask
            {
                Id = workOrderServiceTaskId,
                msdyn_PercentComplete = 100.00,
                ovs_QuestionnaireResponse = @"
                {
                    ""finding-sq_141"": {
                        ""provisionReference"": ""SATR 4.1 (1) (b)"",
                        ""provisionTextEn"": ""<html><strong>Verification of Identity</strong></br><strong>SATR 4.1</strong>: </br><strong>SATR 4.1 (1)</strong>: The air carrier must carry out each verification referred to in section 3 or 4 by</br><strong><mark>SATR 4.1 (1) (b)</mark></strong>: if the passenger presents a piece of photo identification, comparing the passenger’s entire face with the face displayed in the photograph.</br></html>"",
                        ""provisionTextFr"": ""<html><strong>Vérification de l’identité</strong></br><strong>SATR 4.1</strong>: </br><strong>SATR 4.1 (1)</strong>: Le transporteur aérien effectue la vérification visée à l’article 3 ou 4 de la manière suivante :</br><strong><mark>SATR 4.1 (1) (b)</mark></strong>: si le passager présente une pièce d’identité avec photo, en comparant son visage en entier avec le visage paraissant sur la photo.</br></html>"",
                        ""operations"": [{""operationID"": """ + operationGuid + @""",""findingType"": ""717750001""}],
                        ""provisionData"": {""legislationid"": """ + legislationid + @""",""provisioncategoryid"": null},
                        ""comments"": ""test comments 1"",
                        ""reference"": ""SATR 4.1 (1) (b)""
                    },   
                    ""finding-sq_142"": {
                        ""provisionReference"": ""RSDA 4.1 (1) (b)"",
                        ""provisionTextEn"": ""<strong>Verification of Identity</strong></br><strong><mark><mark>SATR 4</mark></mark></strong>: An air carrier must, at the boarding gate for an international flight, verify the identity of each passenger who appears to be 18 years of age or older using</br><ul style='list-style-type:none;'><li><strong>(a)</strong> one of the following pieces of photo identification issued by a government authority that shows the passenger’s surname, first name and any middle names, their date of birth and gender and that is valid:</li><ul style='list-style-type:none;'><li><strong>(i)</strong> a passport issued by the country of which the passenger is a citizen or a national,</li><li><strong>(ii)</strong> a NEXUS card,</li><li><strong>(iii)</strong> any document referred to in subsection 50(1) or 52(1) of the Immigration and Refugee Protection Regulations; or</li></ul><li><strong>(b)</strong> a valid restricted area identity card, as defined in section 3 of the Canadian Aviation Security Regulations, 2012.</li></ul>"",
                        ""provisionTextFr"": ""<strong>Verification of Identity</strong></br><strong><mark><mark><mark>SATR 4</mark></mark></mark></strong>: Tout transporteur aérien vérifie, à la porte d’embarquement pour un vol international, l’identité de chaque passager qui semble âgé de 18 ans ou plus au moyen :</br><ul style='list-style-type:none;'><li><strong>(a)</strong> soit de l’une des pièces d’identité avec photo ci-après qui est délivrée par une autorité gouvernementale, qui indique les nom et prénoms, date de naissance et genre du passager et qui est valide :</li><ul style='list-style-type:none;'><li><strong>(i)</strong> un passeport délivré au passager par le pays dont il est citoyen ou ressortissant,</li><li><strong>(ii)</strong> une carte NEXUS,</li><li><strong>(iii)</strong> un document visé au paragraphe 50(1) ou 52(1) du Règlement sur l’immigration et la protection des réfugiés;</li></ul><li><strong>(b)</strong> soit d’une carte d’identité de zone réglementée au sens de l’article 3 du Règlement canadien de 2012 sur la sûreté aérienne qui est valide.</li></ul>"",
                        ""operations"": [{""operationID"": """ + operationGuid + @""",""findingType"": ""717750001""}],
                        ""provisionData"": {""legislationid"": """ + legislationid + @""",""provisioncategoryid"": null},
                        ""comments"": ""test comments 1"",
                        ""reference"": ""RSDA 4.1 (1) (b)""
                    },
                    ""finding-sq_143"": {
                        ""provisionReference"": ""SATR 4.1 (1) (b)"",
                        ""provisionTextEn"": ""<strong>Verification of Identity</strong></br><strong><mark><mark>SATR 4</mark></mark></strong>: An air carrier must, at the boarding gate for an international flight, verify the identity of each passenger who appears to be 18 years of age or older using</br><ul style='list-style-type:none;'><li><strong>(a)</strong> one of the following pieces of photo identification issued by a government authority that shows the passenger’s surname, first name and any middle names, their date of birth and gender and that is valid:</li><ul style='list-style-type:none;'><li><strong>(i)</strong> a passport issued by the country of which the passenger is a citizen or a national,</li><li><strong>(ii)</strong> a NEXUS card,</li><li><strong>(iii)</strong> any document referred to in subsection 50(1) or 52(1) of the Immigration and Refugee Protection Regulations; or</li></ul><li><strong>(b)</strong> a valid restricted area identity card, as defined in section 3 of the Canadian Aviation Security Regulations, 2012.</li></ul>"",
                        ""provisionTextFr"": ""<strong>Verification of Identity</strong></br><strong><mark><mark><mark>SATR 4</mark></mark></mark></strong>: Tout transporteur aérien vérifie, à la porte d’embarquement pour un vol international, l’identité de chaque passager qui semble âgé de 18 ans ou plus au moyen :</br><ul style='list-style-type:none;'><li><strong>(a)</strong> soit de l’une des pièces d’identité avec photo ci-après qui est délivrée par une autorité gouvernementale, qui indique les nom et prénoms, date de naissance et genre du passager et qui est valide :</li><ul style='list-style-type:none;'><li><strong>(i)</strong> un passeport délivré au passager par le pays dont il est citoyen ou ressortissant,</li><li><strong>(ii)</strong> une carte NEXUS,</li><li><strong>(iii)</strong> un document visé au paragraphe 50(1) ou 52(1) du Règlement sur l’immigration et la protection des réfugiés;</li></ul><li><strong>(b)</strong> soit d’une carte d’identité de zone réglementée au sens de l’article 3 du Règlement canadien de 2012 sur la sûreté aérienne qui est valide.</li></ul>"",
                        ""operations"": [{""operationID"": """ + operationGuid + @""",""findingType"": ""717750001""}],
                        ""comments"": ""test comments 2""
                    }
                }
                "
            });

            // ASSERT
            var query = new QueryExpression(ovs_Finding.EntityLogicalName)
            {
                ColumnSet = new ColumnSet("ts_qm_rclegislation")
            };
            var findings = orgAdminUIService.RetrieveMultiple(query).Entities.Cast<ovs_Finding>().ToList();

            Assert.Equal(legislationGuid.ToString(), findings[0].ts_qm_rclegislation.Id.ToString());
            Assert.Equal(legislationGuid.ToString(), findings[1].ts_qm_rclegislation.Id.ToString());
        }

        [Fact]
        public void When_ovs_questionnaireresponse_contains_finding_expect_NotetoStakeholder_equals_inspector_comments()
        {
            // ARRANGE
            var serviceAccountId = orgAdminUIService.Create(new Account() { Name = "Test Service Account" });
            var workOrderId = orgAdminUIService.Create(new msdyn_workorder()
            {
                msdyn_name = "300-345678",
                msdyn_ServiceRequest = null,
                msdyn_ServiceAccount = new EntityReference(Account.EntityLogicalName, serviceAccountId)

            });
            var workOrderServiceTaskId = orgAdminUIService.Create(new msdyn_workorderservicetask()
            {
                msdyn_name = "200-345678-1",
                msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId), // belongs to a work order
                msdyn_PercentComplete = 0.00,
            });
            var account = orgAdminUIService.Create(new Account());
            var accountReference = new EntityReference(Account.EntityLogicalName, account);
            var operation = orgAdminUIService.Create(new ovs_operation()
            {
                Id = new Guid("9de3a6e3-c4ad-eb11-8236-000d3ae8b866"),
                ts_stakeholder = accountReference
            });

            orgAdminUIService.Create(new qm_rclegislation()
            {
                Id = new Guid("417d6241-386a-eb11-a812-000d3a84129e"),
                qm_name = "SATR 4",
                ts_NameEnglish = "SATR 4",
                ts_NameFrench = "RSDA 4",
            });

            // ACT
            orgAdminUIService.Update(new msdyn_workorderservicetask
            {
                Id = workOrderServiceTaskId,
                msdyn_PercentComplete = 100.00,
                ovs_QuestionnaireResponse = @"
                {
                    ""finding-sq_142"": {
                        ""provisionReference"": ""SATR 4"",
                        ""provisionTextEn"": ""<strong>Verification of Identity</strong></br><strong><mark><mark>SATR 4</mark></mark></strong>: An air carrier must, at the boarding gate for an international flight, verify the identity of each passenger who appears to be 18 years of age or older using</br><ul style='list-style-type:none;'><li><strong>(a)</strong> one of the following pieces of photo identification issued by a government authority that shows the passenger’s surname, first name and any middle names, their date of birth and gender and that is valid:</li><ul style='list-style-type:none;'><li><strong>(i)</strong> a passport issued by the country of which the passenger is a citizen or a national,</li><li><strong>(ii)</strong> a NEXUS card,</li><li><strong>(iii)</strong> any document referred to in subsection 50(1) or 52(1) of the Immigration and Refugee Protection Regulations; or</li></ul><li><strong>(b)</strong> a valid restricted area identity card, as defined in section 3 of the Canadian Aviation Security Regulations, 2012.</li></ul>"",
                        ""provisionTextFr"": ""<strong>Verification of Identity</strong></br><strong><mark><mark><mark>SATR 4</mark></mark></mark></strong>: Tout transporteur aérien vérifie, à la porte d’embarquement pour un vol international, l’identité de chaque passager qui semble âgé de 18 ans ou plus au moyen :</br><ul style='list-style-type:none;'><li><strong>(a)</strong> soit de l’une des pièces d’identité avec photo ci-après qui est délivrée par une autorité gouvernementale, qui indique les nom et prénoms, date de naissance et genre du passager et qui est valide :</li><ul style='list-style-type:none;'><li><strong>(i)</strong> un passeport délivré au passager par le pays dont il est citoyen ou ressortissant,</li><li><strong>(ii)</strong> une carte NEXUS,</li><li><strong>(iii)</strong> un document visé au paragraphe 50(1) ou 52(1) du Règlement sur l’immigration et la protection des réfugiés;</li></ul><li><strong>(b)</strong> soit d’une carte d’identité de zone réglementée au sens de l’article 3 du Règlement canadien de 2012 sur la sûreté aérienne qui est valide.</li></ul>"",
                        ""operations"": [{""operationID"": ""9de3a6e3-c4ad-eb11-8236-000d3ae8b866"",""findingType"": ""717750001""}],
                        ""comments"": ""comments to be duplicated in NotetoStakeholder"",
                    }
                }
                "
            });

            // ASSERT
            var query = new QueryExpression(ovs_Finding.EntityLogicalName)
            {
                ColumnSet = new ColumnSet("ovs_findingcomments", "ts_notetostakeholder")
            };
            var findings = orgAdminUIService.RetrieveMultiple(query).Entities.Cast<ovs_Finding>().ToList();

            // Expect exactly one finding
            Assert.Equal(findings[0].ovs_FindingComments, findings[0].ts_NotetoStakeholder);
        }


        [Fact]
        public void When_work_order_service_task_already_has_findings_with_multiple_copies_and_ovs_questionnaireresponse_contains_new_finding_with_operation_array_expect_new_finding_records_to_be_created_with_correct_infix()
        {
            // ARRANGE
            var serviceAccountId = orgAdminUIService.Create(new Account { Name = "Test Service Account" });
            var incidentId = orgAdminUIService.Create(new Incident { });
            var workOrderId = orgAdminUIService.Create(new msdyn_workorder
            {
                msdyn_name = "300-345678",
                msdyn_ServiceRequest = new EntityReference(Incident.EntityLogicalName, incidentId),
                msdyn_ServiceAccount = new EntityReference(Account.EntityLogicalName, serviceAccountId)
            });

            var testAccount1Id = orgAdminUIService.Create(new Account());
            var testAccount2Id = orgAdminUIService.Create(new Account());
            var testAccount3Id = orgAdminUIService.Create(new Account());
            var testAccount4Id = orgAdminUIService.Create(new Account());

            var testAccountReference1 = new EntityReference(Account.EntityLogicalName, testAccount1Id);
            var testAccountReference2 = new EntityReference(Account.EntityLogicalName, testAccount2Id);
            var testAccountReference3 = new EntityReference(Account.EntityLogicalName, testAccount3Id);
            var testAccountReference4 = new EntityReference(Account.EntityLogicalName, testAccount4Id);

            var testOperation1Id = orgAdminUIService.Create(new ovs_operation()
            {
                Id = new Guid("9de3a6e3-c4ad-eb11-8236-000d3ae8b866"),
                ts_stakeholder = testAccountReference1
            });
            var testOperation2Id = orgAdminUIService.Create(new ovs_operation()
            {
                Id = new Guid("6b796de3-b3a4-eb11-9442-000d3a8410dc"),
                ts_stakeholder = testAccountReference2
            });
            var testOperation3Id = orgAdminUIService.Create(new ovs_operation()
            {
                Id = new Guid("7d085d54-c2a9-eb11-9442-000d3a8410dc"),
                ts_stakeholder = testAccountReference3
            });
            var testOperation4Id = orgAdminUIService.Create(new ovs_operation()
            {
                Id = new Guid("22364b7e-e1ce-eb11-bacc-0022483c068d"),
                ts_stakeholder = testAccountReference4
            });

            orgAdminUIService.Create(new qm_rclegislation()
            {
                Id = new Guid("417d6241-386a-eb11-a812-000d3a84129e"),
                qm_name = "SATR 4",
                ts_NameEnglish = "SATR 4",
                ts_NameFrench = "RSDA 4",
            });
            orgAdminUIService.Create(new qm_rclegislation()
            {
                Id = new Guid("517d6241-386a-eb11-a812-000d3a84129e"),
                qm_name = "SATR 5",
                ts_NameEnglish = "SATR 5",
                ts_NameFrench = "RSDA 5",
            });
            orgAdminUIService.Create(new qm_rclegislation()
            {
                Id = new Guid("617d6241-386a-eb11-a812-000d3a84129e"),
                qm_name = "SATR 6",
                ts_NameEnglish = "SATR 6",
                ts_NameFrench = "RSDA 6",
            });


            var existingWorkOrderServiceTaskId = orgAdminUIService.Create(new msdyn_workorderservicetask
            {
                msdyn_name = "200-345678-1",
                msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId), // belongs to a work order
                msdyn_PercentComplete = 0,
                ovs_QuestionnaireResponse = @"
                {
                    ""finding-sq_142"": {
                        ""provisionReference"": ""SATR 4"",
                        ""provisionTextEn"": ""<strong>Verification of Identity</strong></br><strong><mark><mark>SATR 4</mark></mark></strong>: An air carrier must, at the boarding gate for an international flight, verify the identity of each passenger who appears to be 18 years of age or older using</br><ul style='list-style-type:none;'><li><strong>(a)</strong> one of the following pieces of photo identification issued by a government authority that shows the passenger’s surname, first name and any middle names, their date of birth and gender and that is valid:</li><ul style='list-style-type:none;'><li><strong>(i)</strong> a passport issued by the country of which the passenger is a citizen or a national,</li><li><strong>(ii)</strong> a NEXUS card,</li><li><strong>(iii)</strong> any document referred to in subsection 50(1) or 52(1) of the Immigration and Refugee Protection Regulations; or</li></ul><li><strong>(b)</strong> a valid restricted area identity card, as defined in section 3 of the Canadian Aviation Security Regulations, 2012.</li></ul>"",
                        ""provisionTextFr"": ""<strong>Verification of Identity</strong></br><strong><mark><mark><mark>SATR 4</mark></mark></mark></strong>: Tout transporteur aérien vérifie, à la porte d’embarquement pour un vol international, l’identité de chaque passager qui semble âgé de 18 ans ou plus au moyen :</br><ul style='list-style-type:none;'><li><strong>(a)</strong> soit de l’une des pièces d’identité avec photo ci-après qui est délivrée par une autorité gouvernementale, qui indique les nom et prénoms, date de naissance et genre du passager et qui est valide :</li><ul style='list-style-type:none;'><li><strong>(i)</strong> un passeport délivré au passager par le pays dont il est citoyen ou ressortissant,</li><li><strong>(ii)</strong> une carte NEXUS,</li><li><strong>(iii)</strong> un document visé au paragraphe 50(1) ou 52(1) du Règlement sur l’immigration et la protection des réfugiés;</li></ul><li><strong>(b)</strong> soit d’une carte d’identité de zone réglementée au sens de l’article 3 du Règlement canadien de 2012 sur la sûreté aérienne qui est valide.</li></ul>"",
                        ""comments"": ""existing comments"",
                        ""operations"": [
	                        {""operationID"": ""9de3a6e3-c4ad-eb11-8236-000d3ae8b866"",""findingType"": ""717750001""},
	                        {""operationID"": ""6b796de3-b3a4-eb11-9442-000d3a8410dc"",""findingType"": ""717750001""},
	                        {""operationID"": ""7d085d54-c2a9-eb11-9442-000d3a8410dc"",""findingType"": ""717750001""}
                        ]
                    }
                }
                "
            });

            var findingId1 = orgAdminUIService.Create(new ovs_Finding
            {
                ovs_Finding_1 = "100-345678-1-1-1",
                ts_findingmappingkey = existingWorkOrderServiceTaskId + "-finding-sq_142-" + testOperation1Id, // unique finding names are created using the work order service task ID and the reference ID in the questionnaire
                ovs_FindingProvisionReference = "SATR 4",
                ovs_FindingComments = "original comments",
                ovs_FindingFile = "C:\\fakepath\\originalfile.png",
                ovs_CaseId = new EntityReference(Incident.EntityLogicalName, incidentId), // this finding already belongs to the case
                ovs_WorkOrderServiceTaskId = new EntityReference(msdyn_workorderservicetask.EntityLogicalName, existingWorkOrderServiceTaskId), // this finding already belongs to a work order service task
                statuscode = ovs_Finding_statuscode.New, // finding is also already active
                statecode = ovs_FindingState.Active, // finding is also already active
            });
            var findingId2 = orgAdminUIService.Create(new ovs_Finding
            {
                ovs_Finding_1 = "100-345678-1-1-2",
                ts_findingmappingkey = existingWorkOrderServiceTaskId + "-finding-sq_142-" + testOperation2Id, // unique finding names are created using the work order service task ID and the reference ID in the questionnaire
                ovs_FindingProvisionReference = "SATR 4",
                ovs_FindingComments = "original comments",
                ovs_FindingFile = "C:\\fakepath\\originalfile.png",
                ovs_CaseId = new EntityReference(Incident.EntityLogicalName, incidentId), // this finding already belongs to the case
                ovs_WorkOrderServiceTaskId = new EntityReference(msdyn_workorderservicetask.EntityLogicalName, existingWorkOrderServiceTaskId), // this finding already belongs to a work order service task
                statuscode = ovs_Finding_statuscode.New, // finding is also already active
                statecode = ovs_FindingState.Active, // finding is also already active
            });
            var findingId3 = orgAdminUIService.Create(new ovs_Finding
            {
                ovs_Finding_1 = "100-345678-1-1-3",
                ts_findingmappingkey = existingWorkOrderServiceTaskId + "-finding-sq_142-" + testOperation3Id, // unique finding names are created using the work order service task ID and the reference ID in the questionnaire
                ovs_FindingProvisionReference = "SATR 4",
                ovs_FindingComments = "original comments",
                ovs_FindingFile = "C:\\fakepath\\originalfile.png",
                ovs_CaseId = new EntityReference(Incident.EntityLogicalName, incidentId), // this finding already belongs to the case
                ovs_WorkOrderServiceTaskId = new EntityReference(msdyn_workorderservicetask.EntityLogicalName, existingWorkOrderServiceTaskId), // this finding already belongs to a work order service task
                statuscode = ovs_Finding_statuscode.New, // finding is also already active
                statecode = ovs_FindingState.Active, // finding is also already active
            });

            // ACT
            orgAdminUIService.Update(new msdyn_workorderservicetask
            {
                Id = existingWorkOrderServiceTaskId,
                msdyn_PercentComplete = 100.00,
                ovs_QuestionnaireResponse = @"
                {
                    ""finding-sq_142"": {
                        ""provisionReference"": ""SATR 4"",
                        ""provisionTextEn"": ""<strong>Verification of Identity</strong></br><strong><mark><mark>SATR 4</mark></mark></strong>: An air carrier must, at the boarding gate for an international flight, verify the identity of each passenger who appears to be 18 years of age or older using</br><ul style='list-style-type:none;'><li><strong>(a)</strong> one of the following pieces of photo identification issued by a government authority that shows the passenger’s surname, first name and any middle names, their date of birth and gender and that is valid:</li><ul style='list-style-type:none;'><li><strong>(i)</strong> a passport issued by the country of which the passenger is a citizen or a national,</li><li><strong>(ii)</strong> a NEXUS card,</li><li><strong>(iii)</strong> any document referred to in subsection 50(1) or 52(1) of the Immigration and Refugee Protection Regulations; or</li></ul><li><strong>(b)</strong> a valid restricted area identity card, as defined in section 3 of the Canadian Aviation Security Regulations, 2012.</li></ul>"",
                        ""provisionTextFr"": ""<strong>Verification of Identity</strong></br><strong><mark><mark><mark>SATR 4</mark></mark></mark></strong>: Tout transporteur aérien vérifie, à la porte d’embarquement pour un vol international, l’identité de chaque passager qui semble âgé de 18 ans ou plus au moyen :</br><ul style='list-style-type:none;'><li><strong>(a)</strong> soit de l’une des pièces d’identité avec photo ci-après qui est délivrée par une autorité gouvernementale, qui indique les nom et prénoms, date de naissance et genre du passager et qui est valide :</li><ul style='list-style-type:none;'><li><strong>(i)</strong> un passeport délivré au passager par le pays dont il est citoyen ou ressortissant,</li><li><strong>(ii)</strong> une carte NEXUS,</li><li><strong>(iii)</strong> un document visé au paragraphe 50(1) ou 52(1) du Règlement sur l’immigration et la protection des réfugiés;</li></ul><li><strong>(b)</strong> soit d’une carte d’identité de zone réglementée au sens de l’article 3 du Règlement canadien de 2012 sur la sûreté aérienne qui est valide.</li></ul>"",
                        ""comments"": ""new comment"",
                        ""operations"": [
	                        {""operationID"": ""9de3a6e3-c4ad-eb11-8236-000d3ae8b866"",""findingType"": ""717750001""},
	                        {""operationID"": ""6b796de3-b3a4-eb11-9442-000d3a8410dc"",""findingType"": ""717750001""},
	                        {""operationID"": ""7d085d54-c2a9-eb11-9442-000d3a8410dc"",""findingType"": ""717750001""}
                        ]
                    },
                    ""finding-sq_155"": {
                        ""provisionReference"": ""SATR 5"",
                        ""provisionTextEn"": ""<strong>Verification of Identity</strong></br><strong><mark><mark>SATR 4</mark></mark></strong>: An air carrier must, at the boarding gate for an international flight, verify the identity of each passenger who appears to be 18 years of age or older using</br><ul style='list-style-type:none;'><li><strong>(a)</strong> one of the following pieces of photo identification issued by a government authority that shows the passenger’s surname, first name and any middle names, their date of birth and gender and that is valid:</li><ul style='list-style-type:none;'><li><strong>(i)</strong> a passport issued by the country of which the passenger is a citizen or a national,</li><li><strong>(ii)</strong> a NEXUS card,</li><li><strong>(iii)</strong> any document referred to in subsection 50(1) or 52(1) of the Immigration and Refugee Protection Regulations; or</li></ul><li><strong>(b)</strong> a valid restricted area identity card, as defined in section 3 of the Canadian Aviation Security Regulations, 2012.</li></ul>"",
                        ""provisionTextFr"": ""<strong>Verification of Identity</strong></br><strong><mark><mark><mark>SATR 4</mark></mark></mark></strong>: Tout transporteur aérien vérifie, à la porte d’embarquement pour un vol international, l’identité de chaque passager qui semble âgé de 18 ans ou plus au moyen :</br><ul style='list-style-type:none;'><li><strong>(a)</strong> soit de l’une des pièces d’identité avec photo ci-après qui est délivrée par une autorité gouvernementale, qui indique les nom et prénoms, date de naissance et genre du passager et qui est valide :</li><ul style='list-style-type:none;'><li><strong>(i)</strong> un passeport délivré au passager par le pays dont il est citoyen ou ressortissant,</li><li><strong>(ii)</strong> une carte NEXUS,</li><li><strong>(iii)</strong> un document visé au paragraphe 50(1) ou 52(1) du Règlement sur l’immigration et la protection des réfugiés;</li></ul><li><strong>(b)</strong> soit d’une carte d’identité de zone réglementée au sens de l’article 3 du Règlement canadien de 2012 sur la sûreté aérienne qui est valide.</li></ul>"",
                        ""comments"": ""new comment 2"",
                        ""operations"": [
	                        {""operationID"": ""9de3a6e3-c4ad-eb11-8236-000d3ae8b866"",""findingType"": ""717750001""},
	                        {""operationID"": ""6b796de3-b3a4-eb11-9442-000d3a8410dc"",""findingType"": ""717750001""},
	                        {""operationID"": ""7d085d54-c2a9-eb11-9442-000d3a8410dc"",""findingType"": ""717750001""},
	                        {""operationID"": ""22364b7e-e1ce-eb11-bacc-0022483c068d"",""findingType"": ""717750001""}
                        ]
                    },
                    ""finding-sq_162"": {
                        ""provisionReference"": ""SATR 6"",
                        ""provisionTextEn"": ""<strong>Verification of Identity</strong></br><strong><mark><mark>SATR 4</mark></mark></strong>: An air carrier must, at the boarding gate for an international flight, verify the identity of each passenger who appears to be 18 years of age or older using</br><ul style='list-style-type:none;'><li><strong>(a)</strong> one of the following pieces of photo identification issued by a government authority that shows the passenger’s surname, first name and any middle names, their date of birth and gender and that is valid:</li><ul style='list-style-type:none;'><li><strong>(i)</strong> a passport issued by the country of which the passenger is a citizen or a national,</li><li><strong>(ii)</strong> a NEXUS card,</li><li><strong>(iii)</strong> any document referred to in subsection 50(1) or 52(1) of the Immigration and Refugee Protection Regulations; or</li></ul><li><strong>(b)</strong> a valid restricted area identity card, as defined in section 3 of the Canadian Aviation Security Regulations, 2012.</li></ul>"",
                        ""provisionTextFr"": ""<strong>Verification of Identity</strong></br><strong><mark><mark><mark>SATR 4</mark></mark></mark></strong>: Tout transporteur aérien vérifie, à la porte d’embarquement pour un vol international, l’identité de chaque passager qui semble âgé de 18 ans ou plus au moyen :</br><ul style='list-style-type:none;'><li><strong>(a)</strong> soit de l’une des pièces d’identité avec photo ci-après qui est délivrée par une autorité gouvernementale, qui indique les nom et prénoms, date de naissance et genre du passager et qui est valide :</li><ul style='list-style-type:none;'><li><strong>(i)</strong> un passeport délivré au passager par le pays dont il est citoyen ou ressortissant,</li><li><strong>(ii)</strong> une carte NEXUS,</li><li><strong>(iii)</strong> un document visé au paragraphe 50(1) ou 52(1) du Règlement sur l’immigration et la protection des réfugiés;</li></ul><li><strong>(b)</strong> soit d’une carte d’identité de zone réglementée au sens de l’article 3 du Règlement canadien de 2012 sur la sûreté aérienne qui est valide.</li></ul>"",
                        ""comments"": ""new comment 3"",
                        ""operations"": [
	                        {""operationID"": ""9de3a6e3-c4ad-eb11-8236-000d3ae8b866"",""findingType"": ""717750001""},
	                        {""operationID"": ""6b796de3-b3a4-eb11-9442-000d3a8410dc"",""findingType"": ""717750001""},
	                        {""operationID"": ""22364b7e-e1ce-eb11-bacc-0022483c068d"",""findingType"": ""717750001""}
                        ]
                    }
                }
                "
            });

            // ASSERT
            var query = new QueryExpression(ovs_Finding.EntityLogicalName)
            {
                ColumnSet = new ColumnSet("ovs_finding")
            };
            //Retrieve findings ordered by infix then suffix
            var findings = orgAdminUIService.RetrieveMultiple(query).Entities.Cast<ovs_Finding>().OrderBy(f => f.ovs_Finding_1.Split('-')[3]).ThenBy(f => f.ovs_Finding_1.Split('-')[4]).ToList();

            // Expect target to contain 7 ovs_finding references
            Assert.Equal(10, findings.Count());

            // Expect previously created first ovs_finding to have the same name
            var first = findings[0];
            Assert.Equal("100-345678-1-1-1", first.ovs_Finding_1);

            // Expect previously created second ovs_finding to have the same name
            var second = findings[1];
            Assert.Equal("100-345678-1-1-2", second.ovs_Finding_1);

            // Expect previously created third ovs_finding to have the same name
            var third = findings[2];
            Assert.Equal("100-345678-1-1-3", third.ovs_Finding_1);

            // Expect newly created fourth ovs_finding to have the proper name
            var fourth = findings[3];
            Assert.Equal("100-345678-1-2-1", fourth.ovs_Finding_1);

            // Expect newly created fifth ovs_finding to have the proper name
            var fifth = findings[4];
            Assert.Equal("100-345678-1-2-2", fifth.ovs_Finding_1);

            // Expect newly created sixth ovs_finding to have the proper name
            var sixth = findings[5];
            Assert.Equal("100-345678-1-2-3", sixth.ovs_Finding_1);

            // Expect newly created seventh ovs_finding to have the proper name
            var seventh = findings[6];
            Assert.Equal("100-345678-1-2-4", seventh.ovs_Finding_1);

            // Expect newly created eight ovs_finding to have the proper name
            var eighth = findings[7];
            Assert.Equal("100-345678-1-3-1", eighth.ovs_Finding_1);

            // Expect newly created ninth ovs_finding to have the proper name
            var ninth = findings[8];
            Assert.Equal("100-345678-1-3-2", ninth.ovs_Finding_1);

            // Expect newly created tenth ovs_finding to have the proper name
            var tenth = findings[9];
            Assert.Equal("100-345678-1-3-3", tenth.ovs_Finding_1);
        }
    }
}
