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
    public class WorkOrderServiceTaskTests
    {
        [Fact]
        public void When_PostOperation_msdyn_workorderservicetask_update_ovs_questionnaireresponse_contains_finding_expect_ovs_finding_record_created()
        {
            /**********
             * ARRANGE
             **********/
            var context = new XrmFakedContext();

            // Given a work order service task that already belongs to a case (Incident)
            // and has a SurveyJS questionnaire response saved in the ovs_QuestionnaireResponse field.
            var incidentId = Guid.NewGuid();
            var workOrderServiceTaskId = Guid.NewGuid();
            var workOrderServiceTask = new msdyn_workorderservicetask()
            {
                Id = workOrderServiceTaskId,
                ovs_CaseId = new EntityReference(Incident.EntityLogicalName, incidentId),
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

            /**********
             * ACT
             **********/
            // Execute the PostOperationmsdyn_workorderservicetaskUpdate plugin with the defined workOrderServiceTask as a target
            context.ExecutePluginWithTarget<PostOperationmsdyn_workorderservicetaskUpdate>(workOrderServiceTask);

            /**********
             * ASSERT
             **********/
            var findings = context.CreateQuery<ovs_Finding>().ToList();

            // Expect target to now contain 1 ovs_finding reference
            Assert.True(findings.Count == 1, "Findings count mismatch from input");

            // Expect ovs_finding to contain same content as what is in JSON response
            var first = findings.First();
            Assert.Equal("1.3 (2) (d) (i)", first.ovs_FindingProvisionReference);
            Assert.Equal("**1.3 (2) (d)** shipping names listed in Schedule 1 may be<br/>&#160;&#160;&#160;&#160;**(i)** written in the singular or plural,<br/>&#160;&#160;&#160;&#160;**ii)** written in upper or lower case letters, except that when the shipping name is followed by the descriptive text associated with the shipping name the descriptive text must be in lower case letters and the shipping name must be in upper case letters (capitals),<br/>&#160;&#160;&#160;&#160;**ii)** in English only, put in a different word order as long as the full shipping name is used and the word order is a commonly used one,<br/>&#160;&#160;&#160;&#160;**iv)** for solutions and mixtures, followed by the word “SOLUTION” or “MIXTURE”, as appropriate, and may include the concentration of the solution or mixture, and<br/>&#160;&#160;&#160;&#160;**(v)** for waste, preceded or followed by the word “WASTE” or “DÉCHET”;<br/>", first.ovs_FindingProvisionText);
            Assert.Equal("blah 1", first.ovs_FindingComments);
            Assert.Equal("C:\\fakepath\\Untitled.png", first.ovs_FindingFile);

            // Expect ovs_finding to reference the proper entities
            Assert.Equal(workOrderServiceTaskId, first.ovs_WorkOrderServiceTaskId.Id);
            Assert.Equal(incidentId, first.ovs_CaseId.Id);
        }
    }
}
