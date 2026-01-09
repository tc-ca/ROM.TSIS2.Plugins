using System;
using System.Collections.Generic;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace TSIS2.Plugins.QuestionnaireProcessor
{
    /// <summary>
    /// Handles all direct communication with the CRM database.
    /// </summary>
    public class QuestionnaireRepository
    {
        private readonly IOrganizationService _service;
        private readonly ILoggingService _logger;
        private const int SingleLineTextMaxLength = 4000;

        public QuestionnaireRepository(IOrganizationService service, ILoggingService logger)
        {
            _service = service ?? throw new ArgumentNullException(nameof(service));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        /// <summary>
        /// Retrieves the Work Order Service Task entity with the JSON fields.
        /// </summary>
        /// <param name="wostId">The ID of the Work Order Service Task.</param>
        /// <returns>The WOST entity with questionnaire data.</returns>
        public Entity GetWorkOrderServiceTask(Guid wostId)
        {
            try
            {
                _logger.Trace($"Retrieving WOST with ID: {wostId}");

                var wost = _service.Retrieve(
                    "msdyn_workorderservicetask",
                    wostId,
                    new ColumnSet(
                        "msdyn_name",
                        "ovs_questionnaireresponse",
                        "ovs_questionnairedefinition",
                        "msdyn_workorder"
                    ));

                return wost;
            }
            catch (Exception ex)
            {
                _logger.Error($"Error retrieving WOST {wostId}: {ex}");
                throw;
            }
        }


        public Dictionary<string, Entity> GetExistingResponses(Guid workOrderServiceTaskId)
        {
            var existingByName = new Dictionary<string, Entity>(StringComparer.OrdinalIgnoreCase);

            try
            {
                _logger.Trace($"Fetching existing question responses for WOST: {workOrderServiceTaskId}");

                var query = new QueryExpression("ts_questionresponse")
                {
                    ColumnSet = new ColumnSet(
                        "ts_name",
                        "ts_questionnumber",
                        "ts_msdyn_workorderservicetask",
                        "ts_workorder",
                        "ts_questionnaire",
                        "ts_version",
                        "ts_response",
                        "statecode",
                        "ts_questionname",
                        "ts_questiontextenglish",
                        "ts_questiontextfrench",
                        "ts_provisionreference",
                        "ts_provisiontextenglish",
                        "ts_provisiontextfrench",
                        "ts_details"
                    ),
                    Criteria = new FilterExpression
                    {
                        Conditions =
                        {
                            new ConditionExpression(
                                "ts_msdyn_workorderservicetask",
                                ConditionOperator.Equal,
                                workOrderServiceTaskId
                            )
                        }
                    }
                };

                var results = _service.RetrieveMultiple(query);
                _logger.Trace($"Found {results.Entities.Count} existing response records for WOST {workOrderServiceTaskId}");

                foreach (var existingResponse in results.Entities)
                {
                    if (existingResponse.Contains("ts_questionname"))
                    {
                        var questionName = existingResponse.GetAttributeValue<string>("ts_questionname");
                        if (!string.IsNullOrEmpty(questionName))
                        {
                            if (existingByName.ContainsKey(questionName))
                            {
                                _logger.Trace($"Duplicate ts_questionname '{questionName}' found. Overwriting {existingByName[questionName].Id} with {existingResponse.Id}.");
                            }

                            existingByName[questionName] = existingResponse;
                            _logger.Trace($"Found existing response: '{questionName}', Record ID: {existingResponse.Id}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Error($"Error fetching existing responses for WOST {workOrderServiceTaskId}: {ex}");
                throw;
            }

            return existingByName;
        }

        public Guid CreateQuestionResponse(Entity newRecord)
        {
            try
            {
                var recordId = _service.Create(newRecord);
                return recordId;
            }
            catch (Exception ex)
            {
                _logger.Error($"Error creating question response record: {ex}");
                throw;
            }
        }

        public void UpdateQuestionResponse(Entity updatedRecord)
        {
            try
            {
                _service.Update(updatedRecord);
            }
            catch (Exception ex)
            {
                _logger.Error($"Error updating question response record {updatedRecord.Id}: {ex}");
                throw;
            }
        }

        /// <summary>
        /// Updates a question response record with merged details from hidden questions.
        /// Targets ts_details field with REPLACE (clear-on-empty) logic.
        /// </summary>
        /// <param name="recordId">The ID of the record to update.</param>
        /// <param name="newDetailsText">The new formatted details text (e.g. key-value pairs or JSON).</param>
        public void UpdateResponseWithMergedDetails(Guid recordId, string newDetailsText)
        {
            try
            {
                // If empty, we explicitly set to null to clear stale data
                string finalValue = string.IsNullOrWhiteSpace(newDetailsText) ? null : newDetailsText;

                if (finalValue != null)
                {
                    finalValue = TruncateSingleLineSafe(finalValue);
                }

                _service.Update(new Entity("ts_questionresponse", recordId)
                {
                    ["ts_details"] = finalValue
                });
            }
            catch (Exception ex)
            {
                _logger.Error($"Failed to merge hidden question details for parent record {recordId}: {ex}");
                throw;
            }
        }

        /// <summary>
        /// Deactivates a question response record.
        /// </summary>
        /// <param name="recordId">The ID of the record to deactivate.</param>
        public void DeactivateQuestionResponse(Guid recordId)
        {
            try
            {
                _logger.Trace($"Deactivating orphaned record: {recordId}");

                var updateEntity = new Entity("ts_questionresponse", recordId)
                {
                    ["statecode"] = new OptionSetValue(1) // 1 = Inactive
                };

                _service.Update(updateEntity);
                _logger.Trace($"Successfully deactivated record: {recordId}");
            }
            catch (Exception ex)
            {
                _logger.Error($"Error deactivating record {recordId}: {ex}");
                throw;
            }
        }

        /// <summary>
        /// Links a finding record to its parent question record.
        /// </summary>
        /// <param name="findingId">The ID of the finding record.</param>
        /// <param name="parentId">The ID of the parent question record.</param>
        public void LinkFindingToParent(Guid findingId, Guid parentId)
        {
            try
            {
                var updateFinding = new Entity("ts_questionresponse", findingId)
                {
                    ["ts_questionresponse"] = new EntityReference("ts_questionresponse", parentId)
                };

                _service.Update(updateFinding);
            }
            catch (Exception ex)
            {
                _logger.Error($"Error linking finding {findingId} to parent {parentId}: {ex}");
                // Don't throw the exception - just log it and continue
                // This prevents the entire processing from failing due to missing attributes
            }
        }

        /// <summary>
        /// Truncates a string to the single-line text limit (4000),
        /// ensuring we don't have trouble with emojis/surrogate pairs.
        /// </summary>
        private static string TruncateSingleLineSafe(string value)
        {
            if (string.IsNullOrEmpty(value) || value.Length <= SingleLineTextMaxLength)
                return value;

            var truncated = value.Substring(0, SingleLineTextMaxLength);

            if (char.IsHighSurrogate(truncated[truncated.Length - 1]))
                truncated = truncated.Substring(0, truncated.Length - 1);

            return truncated;
        }

        /// <summary>
        /// Gets the organization service for direct access if needed.
        /// </summary>
        public IOrganizationService Service => _service;
    }
}