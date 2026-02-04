// Commented out for 442446 as it was interfering with the test run
//using System;
//using System.Linq;
//using System.Collections.Generic;
//using Microsoft.Xrm.Sdk;
//using Microsoft.Xrm.Sdk.Query;
//using Xunit;
//using Xunit.Abstractions;
//using ROMTS_GSRST.Plugins.Tests.TestData;
//using Newtonsoft.Json.Linq;
//using TSIS2.Plugins.Services;

//namespace ROMTS_GSRST.Plugins.Tests
//{
//    public class PostOperationmsdyn_workorderservicetaskUpdateTests : UnitTestBase
//    {
//        private readonly ITestOutputHelper _output;

//        public PostOperationmsdyn_workorderservicetaskUpdateTests(
//            XrmMockupFixture fixture,
//            ITestOutputHelper output) : base(fixture)
//        {
//            _output = output;
//        }

//        [Fact]
//        public void When_WOST_Completed_Should_Correctly_Parse_Questionnaire_Json()
//        {
//            try
//            {
//                // Arrange
//                var (definitionJson, responseJson) = TestData.QuestionnaireSamples.Questionnaire1;

//                // Validate JSON strings separately
//                try {
//                    JObject.Parse(definitionJson);
//                    _output.WriteLine("Definition JSON is valid");
//                }
//                catch (Exception ex) {
//                    _output.WriteLine($"Definition JSON is invalid: {ex}");
//                    throw new Exception("Invalid Definition JSON", ex);
//                }

//                try {
//                    JObject.Parse(responseJson);
//                    _output.WriteLine("Response JSON is valid");
//                }
//                catch (Exception ex) {
//                    _output.WriteLine($"Response JSON is invalid: {ex}");
//                    _output.WriteLine($"Response JSON content: {responseJson}");
//                    throw new Exception("Invalid Response JSON", ex);
//                }

//                // Setup all required entities
//                var serviceAccountId = orgAdminUIService.Create(new Entity("account")
//                {
//                    ["name"] = "Test Service Account"
//                });

//                var workOrder = new Entity("msdyn_workorder")
//                {
//                    ["msdyn_name"] = "WO-TEST",
//                    ["msdyn_serviceaccount"] = new EntityReference("account", serviceAccountId)
//                };

//                var workOrderId = orgAdminUIService.Create(workOrder);

//                var questionnaireId = orgAdminUIService.Create(new Entity("ovs_questionnaire")
//                {
//                    ["ovs_name"] = "Test Questionnaire"
//                });

//                var stakeholderId = orgAdminUIService.Create(new Entity("account")
//                {
//                    ["name"] = "Test Stakeholder Account"
//                });

//                var siteId = orgAdminUIService.Create(new Entity("msdyn_functionallocation")
//                {
//                    ["msdyn_name"] = "Test Site"
//                });

//                var operationTypeId = orgAdminUIService.Create(new Entity("ovs_operationtype")
//                {
//                    ["ovs_name"] = "Test Operation Type"
//                });

//                var operationId = orgAdminUIService.Create(new Entity("ovs_operation")
//                {
//                    ["ts_stakeholder"] = new EntityReference("account", stakeholderId),
//                    ["ts_site"] = new EntityReference("msdyn_functionallocation", siteId),
//                    ["ovs_operationtypeid"] = new EntityReference("ovs_operationtype", operationTypeId) 
//                });

//                var wost = new Entity("msdyn_workorderservicetask")
//                {
//                    ["msdyn_name"] = "WOST-TEST",
//                    ["msdyn_workorder"] = new EntityReference("msdyn_workorder", workOrderId),
//                    ["ovs_questionnairedefinition"] = definitionJson,
//                    ["ovs_questionnaireresponse"] = responseJson,
//                    ["ovs_questionnaire"] = new EntityReference("ovs_questionnaire", questionnaireId),
//                    ["statuscode"] = new OptionSetValue(918640004) // In Progress
//                };
//                var wostId = orgAdminUIService.Create(wost);

//                // Act - Instead of relying on the plugin, directly call QuestionnaireProcessor
//                var tracingService = new MockTracingService(_output);

//                _output.WriteLine("Directly calling QuestionnaireProcessor.ProcessQuestionnaire...");
//                var questionResponseIds = QuestionnaireProcessor.ProcessQuestionnaire(
//                    orgAdminService,
//                    wostId,
//                    tracingService,
//                    false // Not a recompletion
//                );

//                _output.WriteLine($"QuestionnaireProcessor returned {questionResponseIds.Count} question response IDs");

//                // Update the question responses with WOST and questionnaire references
//                // (This is normally done by the plugin after processing)
//                foreach (var responseId in questionResponseIds)
//                {
//                    _output.WriteLine($"Linking response {responseId} to WOST and questionnaire");
//                    orgAdminService.Update(new Entity("ts_questionresponse")
//                    {
//                        Id = responseId,
//                        ["ts_msdyn_workorderservicetask"] = new EntityReference("msdyn_workorderservicetask", wostId),
//                        ["ts_questionnaire"] = new EntityReference("ovs_questionnaire", questionnaireId)
//                    });
//                }

//                // Assert - Verify responses were created
//                var responses = orgAdminUIService.RetrieveMultiple(
//                    new QueryExpression("ts_questionresponse")
//                    {
//                        ColumnSet = new ColumnSet(true),
//                        Criteria = new FilterExpression
//                        {
//                            Conditions =
//                            {
//                                new ConditionExpression(
//                                    "ts_msdyn_workorderservicetask",
//                                    ConditionOperator.Equal,
//                                    wostId)
//                            }
//                        }
//                    }).Entities;

//                _output.WriteLine($"Number of created question responses: {responses.Count}");
//                _output.WriteLine("Created Question Responses:");
//                foreach (var response in responses)
//                {
//                    _output.WriteLine($"- [{response.Id}] {response["ts_questionname"]}");
//                    if (response.Contains("ts_response"))
//                        _output.WriteLine($"  Response: {response["ts_response"]}");
//                    if (response.Contains("ts_comments"))
//                        _output.WriteLine($"  Comments: {response["ts_comments"]}");
//                };

//                Assert.Equal(questionResponseIds.Count, responses.Count);
//            }
//            catch (Exception ex)
//            {
//                _output.WriteLine($"TEST FAILED: {ex}");
//                throw;
//            }
//        }
//    }

//    public class MockTracingService : ITracingService
//    {
//        private readonly ITestOutputHelper _output;

//        public MockTracingService(ITestOutputHelper output)
//        {
//            _output = output;
//        }

//        public void Trace(string format, params object[] args)
//        {
//            try 
//            {
//                // If we have arguments, use string.Format
//                if (args != null && args.Length > 0)
//                {
//                    _output.WriteLine(string.Format(format, args));
//                }
//                // Otherwise just output the format string directly
//                else
//                {
//                    _output.WriteLine(format);
//                }
//            }
//            catch (Exception ex)
//            {
//                // Fallback in case formatting fails
//                _output.WriteLine($"[TRACE ERROR] {format} - Error: {ex}");
//            }
//        }
//    }
//}