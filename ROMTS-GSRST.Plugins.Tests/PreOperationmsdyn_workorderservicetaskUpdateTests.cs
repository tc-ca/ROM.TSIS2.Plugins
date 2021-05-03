using System;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Messages;
using System.ServiceModel;
using DG.Tools.XrmMockup;
using Xunit;
using Xunit.Sdk;

namespace ROMTS_GSRST.Plugins.Tests
{
    public class PreOperationmsdyn_workorderservicetaskUpdateTests : UnitTestBase
    {
        public PreOperationmsdyn_workorderservicetaskUpdateTests(XrmMockupFixture fixture) : base(fixture) { }

        [Fact]
        public void When_ovs_questionnaireresponse_contains_finding_but_work_order_service_task_is_not_100_percent_complete_expect_ovs_finding_not_created()
        {
            /**********
            * ARRANGE
            **********/
            var regulatedEntityId = orgAdminUIService.Create(new Account() { Name = "Test Regulated Entity" });
            var workOrderId = orgAdminUIService.Create(new msdyn_workorder()
            {
                msdyn_name = "300-34567",
                msdyn_ServiceRequest = null,
                ovs_regulatedentity = new EntityReference(Account.EntityLogicalName, regulatedEntityId)

            });

            /**********
            * ACT
            **********/
            var workOrderServiceTaskId = orgAdminUIService.Create(new msdyn_workorderservicetask()
            {
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
            });

            /**********
             * ASSERT
             **********/
            using (var context = new Xrm(orgAdminUIService))
            {
                var findings = context.ovs_FindingSet.ToList();
                // Expect no findings created because questionnaire is not completed 
                Assert.Empty(findings);
            }
        }

        [Fact]
        public void When_ovs_questionnaireresponse_contains_finding_expect_do_not_recreate_existing_ovs_finding()
        {

            /**********
             * ARRANGE
             **********/
            var regulatedEntityId = orgAdminUIService.Create(new Account() { Name = "Test Regulated Entity" });
            var incidentId = orgAdminUIService.Create(new Incident());
            var workOrderId = orgAdminUIService.Create(new msdyn_workorder()
            {
                msdyn_name = "300-34567",
                msdyn_ServiceRequest = new EntityReference(Incident.EntityLogicalName, incidentId),
                ovs_regulatedentity = new EntityReference(Account.EntityLogicalName, regulatedEntityId)

            });
            var workOrderServiceTaskId = orgAdminUIService.Create(new msdyn_workorderservicetask()
            {
                msdyn_name = "200-34567-1",
                msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId), // belongs to a work order
                msdyn_PercentComplete = 100.00,
                ovs_QuestionnaireResponse = @"
                {
                    ""finding-sq_142"": {
                        ""provisionReference"": ""SATR 4"",
                        ""provisionTextEn"": ""<strong>Verification of Identity</strong></br><strong><mark><mark>SATR 4</mark></mark></strong>: An air carrier must, at the boarding gate for an international flight, verify the identity of each passenger who appears to be 18 years of age or older using</br><ul style='list-style-type:none;'><li><strong>(a)</strong> one of the following pieces of photo identification issued by a government authority that shows the passenger’s surname, first name and any middle names, their date of birth and gender and that is valid:</li><ul style='list-style-type:none;'><li><strong>(i)</strong> a passport issued by the country of which the passenger is a citizen or a national,</li><li><strong>(ii)</strong> a NEXUS card,</li><li><strong>(iii)</strong> any document referred to in subsection 50(1) or 52(1) of the Immigration and Refugee Protection Regulations; or</li></ul><li><strong>(b)</strong> a valid restricted area identity card, as defined in section 3 of the Canadian Aviation Security Regulations, 2012.</li></ul>"",
                        ""provisionTextFr"": ""<strong>Verification of Identity</strong></br><strong><mark><mark><mark>SATR 4</mark></mark></mark></strong>: Tout transporteur aérien vérifie, à la porte d’embarquement pour un vol international, l’identité de chaque passager qui semble âgé de 18 ans ou plus au moyen :</br><ul style='list-style-type:none;'><li><strong>(a)</strong> soit de l’une des pièces d’identité avec photo ci-après qui est délivrée par une autorité gouvernementale, qui indique les nom et prénoms, date de naissance et genre du passager et qui est valide :</li><ul style='list-style-type:none;'><li><strong>(i)</strong> un passeport délivré au passager par le pays dont il est citoyen ou ressortissant,</li><li><strong>(ii)</strong> une carte NEXUS,</li><li><strong>(iii)</strong> un document visé au paragraphe 50(1) ou 52(1) du Règlement sur l’immigration et la protection des réfugiés;</li></ul><li><strong>(b)</strong> soit d’une carte d’identité de zone réglementée au sens de l’article 3 du Règlement canadien de 2012 sur la sûreté aérienne qui est valide.</li></ul>"",
                        ""comments"": ""old comments"",
                        ""documentaryEvidence"": ""C:\\fakepath\\oldfile.png""
                    }
                }
                "
            });

            var findingId = orgAdminUIService.Create(new ovs_Finding()
            {
                ts_findingmappingkey = workOrderServiceTaskId + "-finding-sq_142",
                ovs_FindingProvisionReference = "SATR 4",
                ovs_FindingComments = "original comments",
                ovs_FindingFile = "C:\\fakepath\\originalfile.png",
                ovs_CaseId = new EntityReference(Incident.EntityLogicalName, incidentId), // this finding already belongs to the case
                ovs_WorkOrderServiceTaskId = new EntityReference(msdyn_workorderservicetask.EntityLogicalName, workOrderServiceTaskId), // this finding already belongs to a work order service task
                statuscode = ovs_Finding_statuscode.Active, // finding is also already active
                statecode = ovs_FindingState.Active, // finding is also already active
            });

            /**********
            * ACT
            **********/
            var newQuestionnaireResponse = @"
            {
                ""finding-sq_142"": {
                    ""provisionReference"": ""SATR 4"",
                    ""provisionTextEn"": ""<strong>Verification of Identity</strong></br><strong><mark><mark>SATR 4</mark></mark></strong>: An air carrier must, at the boarding gate for an international flight, verify the identity of each passenger who appears to be 18 years of age or older using</br><ul style='list-style-type:none;'><li><strong>(a)</strong> one of the following pieces of photo identification issued by a government authority that shows the passenger’s surname, first name and any middle names, their date of birth and gender and that is valid:</li><ul style='list-style-type:none;'><li><strong>(i)</strong> a passport issued by the country of which the passenger is a citizen or a national,</li><li><strong>(ii)</strong> a NEXUS card,</li><li><strong>(iii)</strong> any document referred to in subsection 50(1) or 52(1) of the Immigration and Refugee Protection Regulations; or</li></ul><li><strong>(b)</strong> a valid restricted area identity card, as defined in section 3 of the Canadian Aviation Security Regulations, 2012.</li></ul>"",
                    ""provisionTextFr"": ""<strong>Verification of Identity</strong></br><strong><mark><mark><mark>SATR 4</mark></mark></mark></strong>: Tout transporteur aérien vérifie, à la porte d’embarquement pour un vol international, l’identité de chaque passager qui semble âgé de 18 ans ou plus au moyen :</br><ul style='list-style-type:none;'><li><strong>(a)</strong> soit de l’une des pièces d’identité avec photo ci-après qui est délivrée par une autorité gouvernementale, qui indique les nom et prénoms, date de naissance et genre du passager et qui est valide :</li><ul style='list-style-type:none;'><li><strong>(i)</strong> un passeport délivré au passager par le pays dont il est citoyen ou ressortissant,</li><li><strong>(ii)</strong> une carte NEXUS,</li><li><strong>(iii)</strong> un document visé au paragraphe 50(1) ou 52(1) du Règlement sur l’immigration et la protection des réfugiés;</li></ul><li><strong>(b)</strong> soit d’une carte d’identité de zone réglementée au sens de l’article 3 du Règlement canadien de 2012 sur la sûreté aérienne qui est valide.</li></ul>"",
                    ""comments"": ""new comments"",
                    ""documentaryEvidence"": ""C:\\fakepath\\newfile.png""
                }
            }
            ";
            orgAdminUIService.Update(new msdyn_workorderservicetask { Id = workOrderServiceTaskId, ovs_QuestionnaireResponse = newQuestionnaireResponse });

            /**********
             * ASSERT
             **********/
            using (var context = new Xrm(orgAdminUIService))
            {
                var finding = orgAdminUIService.Retrieve(ovs_Finding.EntityLogicalName, findingId, new ColumnSet()).ToEntity<ovs_Finding>();
                var findings = context.ovs_FindingSet.ToList();

                // Expect target to still only contain the same finding
                Assert.True(findings.Count == 1, "Expecting 1 finding");
                Assert.Equal(findings.First().Id, finding.Id);
            };

        }
    }
}
