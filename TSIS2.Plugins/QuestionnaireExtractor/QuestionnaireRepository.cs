using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;

namespace TSIS2.Plugins.QuestionnaireExtractor
{
    /// <summary>
    /// Handles all direct communication with the CRM database.
    /// This is the only class that should use IOrganizationService.
    /// </summary>
    public class QuestionnaireRepository
    {
        private readonly IOrganizationService _service;
        private readonly ILoggingService _logger;

        /// <summary>
        /// Initializes a new instance of the QuestionnaireRepository class.
        /// </summary>
        /// <param name="service">The CRM organization service.</param>
        /// <param name="logger">The logging service for logging.</param>
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
                
                var wost = _service.Retrieve("msdyn_workorderservicetask",
                    wostId,
                    new ColumnSet(
                        "msdyn_name",
                        "ovs_questionnaireresponse",
                        "ovs_questionnairedefinition"
                    ));

                _logger.Trace($"Successfully retrieved WOST: {wost.GetAttributeValue<string>("msdyn_name")}");
                return wost;
            }
            catch (Exception ex)
            {
                _logger.Error($"Error retrieving WOST {wostId}: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Retrieves existing question responses for a Work Order Service Task.
        /// </summary>
        /// <param name="workOrderServiceTaskId">The ID of the Work Order Service Task.</param>
        /// <returns>A tuple containing dictionaries of existing responses by number and by name/number.</returns>
        public (Dictionary<int, Entity> byNumber, Dictionary<QuestionKey, Entity> byNameAndNumber) GetExistingResponses(Guid workOrderServiceTaskId)
        {
            var existingByNumber = new Dictionary<int, Entity>();
            var existingByNameAndNumber = new Dictionary<QuestionKey, Entity>();

            try
            {
                _logger.Trace($"Fetching existing question responses for WOST: {workOrderServiceTaskId}");
                
                var query = new QueryExpression("ts_questionresponse")
                {
                    ColumnSet = new ColumnSet(true),
                    Criteria = new FilterExpression
                    {
                        Conditions =
                        {
                            new ConditionExpression("ts_msdyn_workorderservicetask", ConditionOperator.Equal, workOrderServiceTaskId)
                        }
                    }
                };

                var results = _service.RetrieveMultiple(query);
                _logger.Trace($"Found {results.Entities.Count} existing response records for WOST {workOrderServiceTaskId}");
                
                foreach (var existingResponse in results.Entities)
                {
                    int? questionNum = null;
                    if (existingResponse.Contains("ts_questionnumber"))
                    {
                        questionNum = existingResponse.GetAttributeValue<int>("ts_questionnumber");
                        existingByNumber[questionNum.Value] = existingResponse;
                        _logger.Trace($"Found existing response by number: Q#{questionNum}, Record ID: {existingResponse.Id}");
                    }

                    if (existingResponse.Contains("ts_questionname"))
                    {
                        var questionName = existingResponse.GetAttributeValue<string>("ts_questionname");
                        if (!string.IsNullOrEmpty(questionName))
                        {
                            var questionKey = new QuestionKey(questionName, questionNum);
                            existingByNameAndNumber[questionKey] = existingResponse;
                            _logger.Trace($"Found existing response by name and number: '{questionName}' #{questionNum}, Record ID: {existingResponse.Id}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Error($"Error fetching existing responses for WOST {workOrderServiceTaskId}: {ex.Message}");
                throw;
            }

            return (existingByNumber, existingByNameAndNumber);
        }

        /// <summary>
        /// Creates a new question response record in CRM.
        /// </summary>
        /// <param name="newRecord">The entity to create.</param>
        /// <returns>The ID of the created record.</returns>
        public Guid CreateQuestionResponse(Entity newRecord)
        {
            try
            {
                _logger.Trace($"Creating question response record: {newRecord.GetAttributeValue<string>("ts_name")}");
                var recordId = _service.Create(newRecord);
                _logger.Trace($"Successfully created question response record with ID: {recordId}");
                return recordId;
            }
            catch (Exception ex)
            {
                _logger.Error($"Error creating question response record: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Updates an existing question response record in CRM.
        /// </summary>
        /// <param name="updatedRecord">The entity to update.</param>
        public void UpdateQuestionResponse(Entity updatedRecord)
        {
            try
            {
                _logger.Trace($"Updating question response record: {updatedRecord.Id}");
                _service.Update(updatedRecord);
                _logger.Trace($"Successfully updated question response record: {updatedRecord.Id}");
            }
            catch (Exception ex)
            {
                _logger.Error($"Error updating question response record {updatedRecord.Id}: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Updates a question response record with merged details from hidden questions.
        /// </summary>
        /// <param name="recordId">The ID of the record to update.</param>
        /// <param name="newResponseJson">The new response JSON with merged details.</param>
        public void UpdateResponseWithMergedDetails(Guid recordId, string newResponseJson)
        {
            try
            {
                _logger.Trace($"Updating response with merged details for record: {recordId}");
                
                // Check if the response would be too long for CRM
                if (newResponseJson.Length > 4000)
                {
                    _logger.Warning($"Response for record {recordId} was truncated to 4000 characters after merge.");
                    newResponseJson = newResponseJson.Substring(0, 4000);
                }

                var updateRecord = new Entity("ts_questionresponse", recordId)
                {
                    ["ts_answer"] = newResponseJson
                };

                _service.Update(updateRecord);
                _logger.Trace($"Successfully merged details into parent record {recordId}.");
            }
            catch (Exception ex)
            {
                _logger.Error($"Failed to merge hidden question details for parent record {recordId}: {ex.Message}");
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
                _logger.Error($"Error deactivating record {recordId}: {ex.Message}");
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
                _logger.Trace($"Linking finding {findingId} to parent question {parentId}");
                
                var updateFinding = new Entity("ts_questionresponse", findingId)
                {
                    ["ts_questionresponse"] = new EntityReference("ts_questionresponse", parentId)
                };
                
                _service.Update(updateFinding);
                _logger.Trace($"Successfully linked finding {findingId} to parent {parentId}");
            }
            catch (Exception ex)
            {
                _logger.Error($"Error linking finding {findingId} to parent {parentId}: {ex.Message}");
                // Don't throw the exception - just log it and continue
                // This prevents the entire processing from failing due to missing attributes
            }
        }

        /// <summary>
        /// Gets the organization service for direct access if needed.
        /// </summary>
        public IOrganizationService Service => _service;

        /// <summary>
        /// Represents a question key for identifying questions by name and number.
        /// </summary>
        public struct QuestionKey
        {
            public string Name { get; }
            public int? Number { get; }

            public QuestionKey(string name, int? number)
            {
                Name = name;
                Number = number;
            }

            public override bool Equals(object obj)
            {
                if (!(obj is QuestionKey other))
                    return false;

                return string.Equals(Name, other.Name, StringComparison.OrdinalIgnoreCase) &&
                       Number == other.Number;
            }

            public override int GetHashCode()
            {
                unchecked
                {
                    int hash = 17;
                    hash = hash * 23 + (Name?.ToLowerInvariant()?.GetHashCode() ?? 0);
                    hash = hash * 23 + (Number?.GetHashCode() ?? 0);
                    return hash;
                }
            }
        }
    }
} 