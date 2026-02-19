using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using TSIS2.Plugins.WorkOrderExport;

namespace TSIS2.Plugins.WorkOrderExport
{
    /// <summary>
    /// Comprehensive data retrieval service for Work Order export.
    /// Retrieves all related data needed for the PDF export including:
    /// - Work Order Summary
    /// - Service Tasks with Questionnaires
    /// - Findings and Actions
    /// - Documents
    /// - Interactions
    /// - Case details
    /// </summary>
    public class WorkOrderDataRetriever
    {
        private readonly IOrganizationService _service;
        private readonly ITracingService _tracingService;

        public WorkOrderDataRetriever(IOrganizationService service, ITracingService tracingService)
        {
            _service = service ?? throw new ArgumentNullException(nameof(service));
            _tracingService = tracingService ?? throw new ArgumentNullException(nameof(tracingService));
        }

        /// <summary>
        /// Retrieves all data for a Work Order and logs it to the console.
        /// </summary>
        public WorkOrderData RetrieveWorkOrderData(Guid workOrderId)
        {
            _tracingService.Trace($"=== RETRIEVING DATA FOR WORK ORDER: {workOrderId} ===");
            _tracingService.Trace("");

            var data = new WorkOrderData { WorkOrderId = workOrderId };

            try
            {
                // 1. Work Order Summary
                data.WorkOrderSummary = RetrieveWorkOrderSummary(workOrderId);

                // 2. Service Tasks (with Questionnaires)
                data.ServiceTasks = RetrieveServiceTasks(workOrderId);

                // 3. Case (if linked)
                data.CaseData = RetrieveCaseData(data.WorkOrderSummary);

                // 4. Findings (via Case)
                data.Findings = RetrieveFindings(data.CaseData);

                // 5. Actions (via Case)
                data.Actions = RetrieveActions(data.CaseData);

                // 6. Documents (General and Inspection)
                data.Documents = RetrieveDocuments(workOrderId);

                // 7. Interactions (Timeline activities)
                data.Interactions = RetrieveInteractions(workOrderId);

                // 8. Contacts
                data.Contacts = RetrieveContacts(workOrderId);

                // 9. Additional Inspectors (Access Team)
                data.AdditionalInspectors = RetrieveAdditionalInspectors(workOrderId);

                // 10. Supporting Regions
                data.SupportingRegions = RetrieveSupportingRegions(workOrderId);

                _tracingService.Trace("");
                _tracingService.Trace($"[INFO] COMPLETED DATA RETRIEVAL FOR WORK ORDER: {workOrderId}");
                _tracingService.Trace("");
            }
            catch (Exception ex)
            {
                _tracingService.Trace($"[INFO] Error retrieving work order data: {ex.Message}");
                _tracingService.Trace(ex.Message, ex.ToString());
                throw;
            }

            return data;
        }

        #region Work Order Summary

        private WorkOrderSummary RetrieveWorkOrderSummary(Guid workOrderId)
        {
            _tracingService.Trace("[STEP 1/10] Retrieving Work Order Summary...");

            var columns = new ColumnSet(
                // General
                "msdyn_name", "msdyn_workordertype", "ts_region",
                "ovs_operationtypeid", "ts_aircraftclassification", "msdyn_serviceaccount",
                "ts_site", "ts_state", "msdyn_worklocation", "ovs_rational",
                "ts_businessowner", "ownerid", "ts_stakeholdertcscp",
                // Activity Details
                "msdyn_primaryincidenttype", "msdyn_primaryincidentdescription",
                "msdyn_primaryincidentestimatedduration", "ts_overtimerequired",
                // Related To
                "msdyn_servicerequest", "ts_securityincident", "ts_trip",
                "msdyn_parentworkorder", "msdyn_opportunityid",
                // System Fields
                "createdon", "modifiedon", "statecode", "statuscode",
                "ts_numberoffindings",
                // Address
                "msdyn_address1", "msdyn_address2", "msdyn_city",
                "msdyn_stateorprovince", "msdyn_postalcode",
                // Schedule
                "msdyn_datewindowstart", "msdyn_datewindowend",
                "msdyn_instructions"
            // Removed: "ts_stakeholdertscp" - field doesn't exist in this environment
            // Removed: "msdyn_workorderid" - causes casting issues, use entity.Id instead
            );

            var wo = _service.Retrieve("msdyn_workorder", workOrderId, columns);

            var summary = new WorkOrderSummary
            {
                Name = wo.GetAttributeValue<string>("msdyn_name"),
                WorkOrderId = wo.Id.ToString(), // Use entity.Id instead of msdyn_workorderid field
                StateCode = wo.GetAttributeValue<OptionSetValue>("statecode")?.Value,
                StatusCode = wo.GetAttributeValue<OptionSetValue>("statuscode")?.Value,
                CreatedOn = wo.GetAttributeValue<DateTime>("createdon"),
                ModifiedOn = wo.GetAttributeValue<DateTime>("modifiedon"),

                // Lookups
                WorkOrderType = wo.GetAttributeValue<EntityReference>("msdyn_workordertype"),
                Region = wo.GetAttributeValue<EntityReference>("ts_region"),
                OperationType = wo.GetAttributeValue<EntityReference>("ovs_operationtypeid"),
                Stakeholder = wo.GetAttributeValue<EntityReference>("msdyn_serviceaccount"),
                Site = wo.GetAttributeValue<EntityReference>("ts_site"),
                Owner = wo.GetAttributeValue<EntityReference>("ownerid"),

                // Picklists
                AircraftClassification = wo.GetAttributeValue<OptionSetValue>("ts_aircraftclassification"),
                State = wo.GetAttributeValue<OptionSetValue>("ts_state"),
                WorkLocation = wo.GetAttributeValue<OptionSetValue>("msdyn_worklocation"),

                // Activity Details
                PrimaryIncidentType = wo.GetAttributeValue<EntityReference>("msdyn_primaryincidenttype"),
                PrimaryIncidentDescription = wo.GetAttributeValue<string>("msdyn_primaryincidentdescription"),
                PrimaryIncidentEstimatedDuration = wo.GetAttributeValue<int?>("msdyn_primaryincidentestimatedduration"),
                OvertimeRequired = wo.GetAttributeValue<bool?>("ts_overtimerequired"),

                // Related To
                ServiceRequest = wo.GetAttributeValue<EntityReference>("msdyn_servicerequest"),

                // Other
                BusinessOwner = wo.GetAttributeValue<string>("ts_businessowner"),
                StakeholderTCSCP = wo.GetAttributeValue<string>("ts_stakeholdertcscp"),
                NumberOfFindings = wo.GetAttributeValue<int?>("ts_numberoffindings"),

                RawEntity = wo
            };

            LogWorkOrderSummary(summary);
            return summary;
        }

        private void LogWorkOrderSummary(WorkOrderSummary summary)
        {
            _tracingService.Trace("   [INFO] Work Order Summary Retrieved");
            _tracingService.Trace($"   [DETAIL] Name: {summary.Name}");
            _tracingService.Trace($"   [DETAIL] WO ID: {summary.WorkOrderId}");
            _tracingService.Trace($"   [DETAIL] Created: {summary.CreatedOn:yyyy-MM-dd HH:mm}");
            _tracingService.Trace($"   [DETAIL] Owner: {summary.Owner?.Name ?? "N/A"} ({summary.Owner?.LogicalName})");
            _tracingService.Trace($"   [DETAIL] Stakeholder: {summary.Stakeholder?.Name ?? "N/A"}");
            _tracingService.Trace($"   [DETAIL] Stakeholder TCSCP: {summary.StakeholderTCSCP ?? "N/A"}");
            _tracingService.Trace($"   [DETAIL] Region: {summary.Region?.Name ?? "N/A"}");
            _tracingService.Trace($"   [DETAIL] Operation Type: {summary.OperationType?.Name ?? "N/A"}");
            _tracingService.Trace($"   [DETAIL] Site: {summary.Site?.Name ?? "N/A"}");
            _tracingService.Trace($"   [DETAIL] Findings Count: {summary.NumberOfFindings ?? 0}");

            if (summary.ServiceRequest != null)
            {
                _tracingService.Trace($"   [DETAIL] Linked Case: {summary.ServiceRequest.Name} ({summary.ServiceRequest.Id})");
            }

            _tracingService.Trace("");
        }

        #endregion

        #region Service Tasks

        private List<ServiceTaskData> RetrieveServiceTasks(Guid workOrderId)
        {
            _tracingService.Trace("[STEP 2/10] Retrieving Service Tasks...");

            // Using FetchXML as per requirement to include ts_fromoffline and match subgrid behaviour
            var fetchXml = $@"
                <fetch version='1.0' output-format='xml-platform' mapping='logical'>
                  <entity name='msdyn_workorderservicetask'>
                    <attribute name='msdyn_name' />
                    <attribute name='msdyn_tasktype' />
                    <attribute name='statuscode' />
                    <attribute name='statecode' />
                    <attribute name='ovs_questionnairedefinition' />
                    <attribute name='ovs_questionnaireresponse' />
                    <attribute name='ovs_questionnaire' />
                    <attribute name='msdyn_inspectiontaskresult' />
                    <attribute name='msdyn_percentcomplete' />
                    <attribute name='msdyn_estimatedduration' />
                    <attribute name='createdon' />
                    <attribute name='modifiedon' />
                    <attribute name='ts_fromoffline' />
                    <link-entity name='msdyn_workorder' from='msdyn_workorderid' to='msdyn_workorder' alias='bb'>
                      <filter type='and'>
                        <condition attribute='msdyn_workorderid' operator='eq' value='{workOrderId}' />
                      </filter>
                    </link-entity>
                  </entity>
                </fetch>";

            var results = _service.RetrieveMultiple(new FetchExpression(fetchXml));
            var tasks = new List<ServiceTaskData>();

            _tracingService.Trace($"   [INFO] Found {results.Entities.Count} Service Task(s)");

            foreach (var entity in results.Entities)
            {
                var task = new ServiceTaskData
                {
                    Id = entity.Id,
                    Name = entity.GetAttributeValue<string>("msdyn_name"),
                    TaskType = entity.GetAttributeValue<EntityReference>("msdyn_tasktype"),
                    StatusCode = entity.GetAttributeValue<OptionSetValue>("statuscode"),
                    QuestionnaireDefinition = entity.GetAttributeValue<string>("ovs_questionnairedefinition"),
                    QuestionnaireResponse = entity.GetAttributeValue<string>("ovs_questionnaireresponse"),
                    Questionnaire = entity.GetAttributeValue<EntityReference>("ovs_questionnaire"),
                    InspectionResult = entity.GetAttributeValue<OptionSetValue>("msdyn_inspectiontaskresult"),
                    RawEntity = entity
                };

                tasks.Add(task);

                _tracingService.Trace($"   [DETAIL] Task: {task.Name} (ID: {task.Id})");
                _tracingService.Trace($"      Status: {task.StatusCode?.Value}");
                _tracingService.Trace($"      Inspection Result: {task.InspectionResult?.Value}");

                if (task.Questionnaire != null)
                {
                    _tracingService.Trace($"      [DETAIL] Questionnaire: {task.Questionnaire.Name}");
                }

                if (!string.IsNullOrEmpty(task.QuestionnaireDefinition))
                {
                    _tracingService.Trace($"      [DETAIL] Questionnaire Definition Length: {task.QuestionnaireDefinition.Length} chars");
                }

                if (!string.IsNullOrEmpty(task.QuestionnaireResponse))
                {
                    _tracingService.Trace($"      [DETAIL] Response JSON Length: {task.QuestionnaireResponse.Length} chars");
                }
            }

            _tracingService.Trace("");
            return tasks;
        }

        #endregion

        #region Case Data

        private CaseData RetrieveCaseData(WorkOrderSummary woSummary)
        {
            _tracingService.Trace("[STEP 3/10] Retrieving Case Data...");

            if (woSummary.ServiceRequest == null)
            {
                _tracingService.Trace("   [DETAIL] No Case linked to this Work Order");
                _tracingService.Trace("");
                return null;
            }

            try
            {
                var caseEntity = _service.Retrieve("incident", woSummary.ServiceRequest.Id,
                    new ColumnSet(
                        "title", "ticketnumber", "statuscode", "statecode",
                        "description", "customerid",
                        "createdon", "modifiedon"
                    ));

                var caseData = new CaseData
                {
                    Id = caseEntity.Id,
                    Title = caseEntity.GetAttributeValue<string>("title"),
                    CaseNumber = caseEntity.GetAttributeValue<string>("ticketnumber"),
                    StatusCode = caseEntity.GetAttributeValue<OptionSetValue>("statuscode"),
                    Description = caseEntity.GetAttributeValue<string>("description"),
                    Customer = caseEntity.GetAttributeValue<EntityReference>("customerid"),
                    RawEntity = caseEntity
                };

                _tracingService.Trace($"   [INFO] Case Retrieved: {caseData.CaseNumber}");
                _tracingService.Trace($"   [DETAIL] Title: {caseData.Title}");
                _tracingService.Trace($"   [DETAIL] Status: {caseData.StatusCode?.Value}");
                _tracingService.Trace($"   [DETAIL] Description: {(string.IsNullOrEmpty(caseData.Description) ? "N/A" : caseData.Description.Substring(0, Math.Min(100, caseData.Description.Length)))}...");
                _tracingService.Trace("");

                return caseData;
            }
            catch (Exception ex)
            {
                _tracingService.Trace($"   [DETAIL] Could not retrieve Case: {ex.Message}");
                _tracingService.Trace("");
                return null;
            }
        }

        #endregion

        #region Findings

        private List<FindingData> RetrieveFindings(CaseData caseData)
        {
            _tracingService.Trace("[STEP 4/10] Retrieving Findings...");

            if (caseData == null)
            {
                _tracingService.Trace("   [DETAIL] No Case available for finding retrieval");
                _tracingService.Trace("");
                return new List<FindingData>();
            }

            var query = new QueryExpression("ovs_finding")
            {
                ColumnSet = new ColumnSet(
                    "ovs_finding", "ts_findingtype", "ts_finalenforcementaction",
                    "ts_sensitivitylevel", "statecode", "statuscode",
                    "createdon", "modifiedon", "ovs_caseid",
                    "ts_accountid" // Added from user FetchXML
                ),
                Criteria = new FilterExpression
                {
                    Conditions =
                    {
                        new ConditionExpression("ovs_caseid", ConditionOperator.Equal, caseData.Id),
                        new ConditionExpression("statecode", ConditionOperator.Equal, 0) // Active
                    }
                }
            };

            var results = _service.RetrieveMultiple(query);
            var allFindings = new List<FindingData>();

            _tracingService.Trace($"   [INFO] Found {results.Entities.Count} Finding(s) linked to Case");

            foreach (var entity in results.Entities)
            {
                var finding = new FindingData
                {
                    Id = entity.Id,
                    Finding = entity.GetAttributeValue<string>("ovs_finding"),
                    FindingType = entity.GetAttributeValue<OptionSetValue>("ts_findingtype"),
                    FinalEnforcementAction = entity.GetAttributeValue<OptionSetValue>("ts_finalenforcementaction"),
                    SensitivityLevel = entity.GetAttributeValue<OptionSetValue>("ts_sensitivitylevel"),
                    StatusCode = entity.GetAttributeValue<OptionSetValue>("statuscode"),
                    Incident = entity.GetAttributeValue<EntityReference>("ovs_caseid"),
                    AccountId = entity.GetAttributeValue<EntityReference>("ts_accountid"),
                    RawEntity = entity
                };

                allFindings.Add(finding);

                _tracingService.Trace($"   [DETAIL] Finding: {finding.Finding}");
                _tracingService.Trace($"      Type: {finding.FindingType?.Value}");
                _tracingService.Trace($"      Status: {finding.StatusCode?.Value}");
            }

            _tracingService.Trace("");
            return allFindings;
        }

        #endregion

        #region Actions

        private List<ActionData> RetrieveActions(CaseData caseData)
        {
            _tracingService.Trace("[STEP 5/10] Retrieving Actions...");

            if (caseData == null)
            {
                _tracingService.Trace("   [DETAIL] No Case available for Action retrieval");
                _tracingService.Trace("");
                return new List<ActionData>();
            }

            var query = new QueryExpression("ts_action")
            {
                ColumnSet = new ColumnSet(
                    "ts_name", "ts_actioncategory", "ts_actiontype",
                    "ts_case", "statecode", "statuscode", "createdon"
                ),
                Criteria = new FilterExpression
                {
                    Conditions =
                    {
                        new ConditionExpression("ts_case", ConditionOperator.Equal, caseData.Id),
                        new ConditionExpression("statecode", ConditionOperator.Equal, 0)
                    }
                },
                Orders = { new OrderExpression("createdon", OrderType.Ascending) }
            };

            var results = _service.RetrieveMultiple(query);
            var actions = new List<ActionData>();

            foreach (var entity in results.Entities)
            {
                actions.Add(new ActionData
                {
                    Id = entity.Id,
                    Name = entity.GetAttributeValue<string>("ts_name"),
                    ActionCategory = entity.GetAttributeValue<OptionSetValue>("ts_actioncategory"),
                    ActionType = entity.GetAttributeValue<OptionSetValue>("ts_actiontype"),
                    Case = entity.GetAttributeValue<EntityReference>("ts_case"),
                    RawEntity = entity
                });
            }

            _tracingService.Trace($"   [INFO] Found {actions.Count} Action(s) linked to Case");
            _tracingService.Trace("");

            return actions;
        }

        #endregion

        #region Documents

        private DocumentsData RetrieveDocuments(Guid workOrderId)
        {
            _tracingService.Trace("[STEP 6/10] Retrieving Documents...");

            var documentsData = new DocumentsData
            {
                GeneralDocuments = RetrieveGeneralDocuments(workOrderId),
                InspectionDocuments = RetrieveInspectionDocuments(workOrderId)
            };

            _tracingService.Trace($"   [INFO] Total Documents: {documentsData.GeneralDocuments.Count + documentsData.InspectionDocuments.Count}");
            _tracingService.Trace($"      [DETAIL] General: {documentsData.GeneralDocuments.Count}");
            _tracingService.Trace($"      [DETAIL] Inspection: {documentsData.InspectionDocuments.Count}");
            _tracingService.Trace("");

            return documentsData;
        }

        private List<DocumentData> RetrieveGeneralDocuments(Guid workOrderId)
        {
            // N:N relationship: ts_files_msdyn_workorders
            var fetchXml = $@"
                <fetch version='1.0' output-format='xml-platform' mapping='logical'>
                  <entity name='ts_file'>
                    <attribute name='ts_file' />
                    <attribute name='ts_filecategory' />
                    <attribute name='ts_filesubcategory' />
                    <attribute name='ts_filecontext' />
                    <attribute name='ts_description' />
                    <attribute name='createdon' />
                    <attribute name='ts_sharepointlink' />
                    <link-entity name='ts_files_msdyn_workorders' from='ts_fileid' to='ts_fileid' intersect='true' visible='false'>
                      <link-entity name='msdyn_workorder' from='msdyn_workorderid' to='msdyn_workorderid' alias='bb'>
                        <filter type='and'>
                          <condition attribute='msdyn_workorderid' operator='eq' value='{workOrderId}' />
                        </filter>
                      </link-entity>
                    </link-entity>
                  </entity>
                </fetch>";

            var results = _service.RetrieveMultiple(new FetchExpression(fetchXml));
            var documents = new List<DocumentData>();

            foreach (var entity in results.Entities)
            {
                documents.Add(MapDocumentEntity(entity));
            }

            return documents;
        }

        private List<DocumentData> RetrieveInspectionDocuments(Guid workOrderId)
        {
            // 1:N relationship via ts_msdyn_workorder lookup (using FetchXML to match requirement)
            var fetchXml = $@"
                <fetch version='1.0' output-format='xml-platform' mapping='logical'>
                  <entity name='ts_file'>
                    <attribute name='ts_file' />
                    <attribute name='ts_filecategory' />
                    <attribute name='ts_filesubcategory' />
                    <attribute name='ts_filecontext' />
                    <attribute name='ts_description' />
                    <attribute name='createdon' />
                    <attribute name='ts_sharepointlink' />
                    <link-entity name='msdyn_workorder' from='msdyn_workorderid' to='ts_msdyn_workorder' alias='bb'>
                      <filter type='and'>
                        <condition attribute='msdyn_workorderid' operator='eq' value='{workOrderId}' />
                      </filter>
                    </link-entity>
                  </entity>
                </fetch>";

            var results = _service.RetrieveMultiple(new FetchExpression(fetchXml));
            var documents = new List<DocumentData>();

            foreach (var entity in results.Entities)
            {
                documents.Add(MapDocumentEntity(entity));
            }

            return documents;
        }

        private DocumentData MapDocumentEntity(Entity entity)
        {
            return new DocumentData
            {
                Id = entity.Id,
                Name = entity.GetAttributeValue<string>("ts_file"),
                FileCategory = entity.GetAttributeValue<EntityReference>("ts_filecategory"),
                FileSubCategory = entity.GetAttributeValue<EntityReference>("ts_filesubcategory"),
                FileContext = entity.GetAttributeValue<OptionSetValue>("ts_filecontext"),
                Description = entity.GetAttributeValue<string>("ts_description"),
                CreatedOn = entity.GetAttributeValue<DateTime>("createdon"),
                SharePointLink = entity.GetAttributeValue<string>("ts_sharepointlink"),
                RawEntity = entity
            };
        }

        #endregion

        #region Interactions

        private List<InteractionData> RetrieveInteractions(Guid workOrderId)
        {
            _tracingService.Trace("[STEP 7/10] Retrieving Interactions (Timeline & Notes)...");

            var interactions = new List<InteractionData>();

            // 1. Timeline Activities (Email, Phone, Appointment, Task, etc.)
            try
            {
                var fetchXml = $@"
                    <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                      <entity name='activitypointer'>
                        <attribute name='subject' />
                        <attribute name='activitytypecode' />
                        <attribute name='statecode' />
                        <attribute name='statuscode' />
                        <attribute name='description' />
                        <attribute name='createdon' />
                        <attribute name='scheduledstart' />
                        <order attribute='createdon' descending='true' />
                        <filter type='and'>
                          <condition attribute='regardingobjectid' operator='eq' value='{workOrderId}' />
                          <condition attribute='activitytypecode' operator='in'>
                            <value>4201</value> <!-- Appointment -->
                            <value>4202</value> <!-- Email -->
                            <value>4210</value> <!-- Phone Call -->
                            <value>4212</value> <!-- Task -->
                            <value>10853</value> <!-- Custom/Unknown -->
                          </condition>
                        </filter>
                      </entity>
                    </fetch>";

                var results = _service.RetrieveMultiple(new FetchExpression(fetchXml));
                _tracingService.Trace($"   [INFO] Found {results.Entities.Count} Activity(ies)");

                foreach (var entity in results.Entities)
                {
                    interactions.Add(new InteractionData
                    {
                        Id = entity.Id,
                        Subject = entity.GetAttributeValue<string>("subject"),
                        ActivityType = entity.FormattedValues.Contains("activitytypecode") ? entity.FormattedValues["activitytypecode"] : entity.GetAttributeValue<OptionSetValue>("activitytypecode")?.Value.ToString(),
                        StateCode = entity.GetAttributeValue<OptionSetValue>("statecode"),
                        StatusCode = entity.GetAttributeValue<OptionSetValue>("statuscode"),
                        CreatedOn = entity.GetAttributeValue<DateTime>("createdon"),
                        ScheduledStart = entity.GetAttributeValue<DateTime?>("scheduledstart"),
                        Description = entity.GetAttributeValue<string>("description"),
                        IsNote = false,
                        RawEntity = entity
                    });
                }
            }
            catch (Exception ex)
            {
                _tracingService.Trace($"   [INFO] Error retrieving activities: {ex.Message}");
            }

            // 2. Notes (Annotations)
            try
            {
                var fetchXml = $@"
                    <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                      <entity name='annotation'>
                        <attribute name='subject' />
                        <attribute name='notetext' />
                        <attribute name='filename' />
                        <attribute name='isdocument' />
                        <attribute name='createdon' />
                        <order attribute='createdon' descending='true' />
                        <filter type='and'>
                          <condition attribute='objectid' operator='eq' value='{workOrderId}' />
                        </filter>
                      </entity>
                    </fetch>";

                var results = _service.RetrieveMultiple(new FetchExpression(fetchXml));
                _tracingService.Trace($"   [INFO] Found {results.Entities.Count} Note(s)");

                foreach (var entity in results.Entities)
                {
                    var isDocument = entity.GetAttributeValue<bool?>("isdocument") ?? false;
                    var subject = entity.GetAttributeValue<string>("subject");
                    var noteText = entity.GetAttributeValue<string>("notetext");
                    var filename = entity.GetAttributeValue<string>("filename");

                    interactions.Add(new InteractionData
                    {
                        Id = entity.Id,
                        Subject = !string.IsNullOrEmpty(subject) ? subject : (isDocument ? "Attachment" : "Note"),
                        ActivityType = "Note",
                        CreatedOn = entity.GetAttributeValue<DateTime>("createdon"),
                        Description = !string.IsNullOrEmpty(noteText) ? noteText : (isDocument ? $"Attachment: {filename}" : ""),
                        IsNote = true,
                        RawEntity = entity
                    });
                }
            }
            catch (Exception ex)
            {
                _tracingService.Trace($"   [INFO] Error retrieving notes: {ex.Message}");
            }

            // Sort combined list by CreatedOn descending
            interactions = interactions.OrderByDescending(i => i.CreatedOn).ToList();

            _tracingService.Trace($"   [INFO] Total Interactions: {interactions.Count}");
            _tracingService.Trace("");

            return interactions;
        }

        #endregion

        #region Contacts

        private List<ContactData> RetrieveContacts(Guid workOrderId)
        {
            _tracingService.Trace("[STEP 8/10] Retrieving Contacts...");

            var fetchXml = $@"
                <fetch version='1.0' output-format='xml-platform' mapping='logical'>
                  <entity name='contact'>
                    <attribute name='fullname' />
                    <attribute name='emailaddress1' />
                    <attribute name='telephone1' />
                    <attribute name='jobtitle' />
                    <link-entity name='ts_contact_msdyn_workorder' intersect='true' visible='false' to='contactid' from='contactid'>
                      <link-entity name='msdyn_workorder' from='msdyn_workorderid' to='msdyn_workorderid' alias='bb'>
                        <filter type='and'>
                          <condition attribute='msdyn_workorderid' operator='eq' value='{workOrderId}' />
                        </filter>
                      </link-entity>
                    </link-entity>
                  </entity>
                </fetch>";

            var results = _service.RetrieveMultiple(new FetchExpression(fetchXml));
            var contacts = new List<ContactData>();

            _tracingService.Trace($"   [INFO] Found {results.Entities.Count} Contact(s)");

            foreach (var entity in results.Entities)
            {
                var contact = new ContactData
                {
                    Id = entity.Id,
                    FullName = entity.GetAttributeValue<string>("fullname"),
                    Email = entity.GetAttributeValue<string>("emailaddress1"),
                    Phone = entity.GetAttributeValue<string>("telephone1"),
                    JobTitle = entity.GetAttributeValue<string>("jobtitle"),
                    RawEntity = entity
                };

                contacts.Add(contact);

                _tracingService.Trace($"   [DETAIL] {contact.FullName}");
                if (!string.IsNullOrEmpty(contact.Email))
                    _tracingService.Trace($"      [DETAIL] {contact.Email}");
            }

            _tracingService.Trace("");
            return contacts;
        }

        #endregion

        #region Additional Inspectors

        private List<InspectorData> RetrieveAdditionalInspectors(Guid workOrderId)
        {
            _tracingService.Trace("[STEP 9/10] Retrieving Additional Inspectors...");

            var fetchXml = $@"
                <fetch version='1.0' output-format='xml-platform' mapping='logical'>
                  <entity name='systemuser'>
                    <attribute name='fullname' />
                    <attribute name='title' />
                    <attribute name='internalemailaddress' />
                    <link-entity name='teammembership' from='systemuserid' to='systemuserid' link-type='inner' alias='TM_Internal'>
                      <link-entity name='team' from='teamid' to='teamid' link-type='inner'>
                        <filter type='and'>
                          <condition attribute='regardingobjectid' operator='eq' value='{workOrderId}' />
                          <condition attribute='teamtype' operator='eq' value='1' />
                          <condition attribute='teamtemplateid' operator='eq' value='bddf1d45-706d-ec11-8f8e-0022483da5aa' />
                        </filter>
                      </link-entity>
                    </link-entity>
                  </entity>
                </fetch>";

            var results = _service.RetrieveMultiple(new FetchExpression(fetchXml));
            var inspectors = new List<InspectorData>();

            _tracingService.Trace($"   [INFO] Found {results.Entities.Count} Additional Inspector(s)");

            foreach (var entity in results.Entities)
            {
                var inspector = new InspectorData
                {
                    Id = entity.Id,
                    FullName = entity.GetAttributeValue<string>("fullname"),
                    Title = entity.GetAttributeValue<string>("title"),
                    Email = entity.GetAttributeValue<string>("internalemailaddress"),
                    RawEntity = entity
                };

                inspectors.Add(inspector);

                _tracingService.Trace($"   [DETAIL] {inspector.FullName}");
                if (!string.IsNullOrEmpty(inspector.Title))
                    _tracingService.Trace($"      [DETAIL] {inspector.Title}");
            }

            _tracingService.Trace("");
            return inspectors;
        }

        #endregion

        #region Supporting Regions

        private List<SupportingRegionData> RetrieveSupportingRegions(Guid workOrderId)
        {
            _tracingService.Trace("[STEP 10/10] Retrieving Supporting Regions...");

            var fetchXml = $@"
                <fetch version='1.0' mapping='logical'>
                  <entity name='ts_workordertimetracking'>
                    <attribute name='ts_name' />
                    <attribute name='statecode' />
                    <attribute name='ts_region' />
                    <attribute name='ts_workordertimetrackingid' />
                    <order attribute='ts_region' descending='false' />
                    <filter type='and'>
                      <condition attribute='statecode' operator='eq' value='0' />
                    </filter>
                    <link-entity name='msdyn_workorder' from='msdyn_workorderid' to='ts_workorder' alias='bb'>
                      <filter type='and'>
                        <condition attribute='msdyn_workorderid' operator='eq' value='{workOrderId}' />
                      </filter>
                    </link-entity>
                  </entity>
                </fetch>";

            var results = _service.RetrieveMultiple(new FetchExpression(fetchXml));
            var regions = new List<SupportingRegionData>();

            _tracingService.Trace($"   [INFO] Found {results.Entities.Count} Supporting Region(s)");

            foreach (var entity in results.Entities)
            {
                var region = new SupportingRegionData
                {
                    Id = entity.Id,
                    Region = entity.GetAttributeValue<EntityReference>("ts_region"),
                    RawEntity = entity
                };

                regions.Add(region);

                _tracingService.Trace($"   [DETAIL] {region.Region?.Name ?? "N/A"}");
            }

            _tracingService.Trace("");
            return regions;
        }

        #endregion
    }

    #region Data Classes

    public class WorkOrderData
    {
        public Guid WorkOrderId { get; set; }
        public WorkOrderSummary WorkOrderSummary { get; set; }
        public List<ServiceTaskData> ServiceTasks { get; set; }
        public CaseData CaseData { get; set; }
        public List<FindingData> Findings { get; set; }
        public List<ActionData> Actions { get; set; }
        public DocumentsData Documents { get; set; }
        public List<InteractionData> Interactions { get; set; }
        public List<ContactData> Contacts { get; set; }
        public List<InspectorData> AdditionalInspectors { get; set; }
        public List<SupportingRegionData> SupportingRegions { get; set; }
    }

    public class WorkOrderSummary
    {
        public string Name { get; set; }
        public string WorkOrderId { get; set; }
        public int? StateCode { get; set; }
        public int? StatusCode { get; set; }
        public DateTime CreatedOn { get; set; }
        public DateTime ModifiedOn { get; set; }

        public EntityReference WorkOrderType { get; set; }
        public EntityReference Region { get; set; }
        public EntityReference OperationType { get; set; }
        public EntityReference Stakeholder { get; set; }
        public EntityReference Site { get; set; }
        public EntityReference Owner { get; set; }

        public OptionSetValue AircraftClassification { get; set; }
        public OptionSetValue State { get; set; }
        public OptionSetValue WorkLocation { get; set; }

        public EntityReference PrimaryIncidentType { get; set; }
        public string PrimaryIncidentDescription { get; set; }
        public int? PrimaryIncidentEstimatedDuration { get; set; }
        public bool? OvertimeRequired { get; set; }

        public EntityReference ServiceRequest { get; set; }

        public string BusinessOwner { get; set; }
        public string StakeholderTCSCP { get; set; }
        public int? NumberOfFindings { get; set; }

        public Entity RawEntity { get; set; }
    }

    public class ServiceTaskData
    {
        public Guid Id { get; set; }
        public string Name { get; set; }
        public EntityReference TaskType { get; set; }
        public OptionSetValue StatusCode { get; set; }
        public string QuestionnaireDefinition { get; set; }
        public string QuestionnaireResponse { get; set; }
        public EntityReference Questionnaire { get; set; }
        public OptionSetValue InspectionResult { get; set; }
        public Entity RawEntity { get; set; }
    }

    public class CaseData
    {
        public Guid Id { get; set; }
        public string Title { get; set; }
        public string CaseNumber { get; set; }
        public OptionSetValue StatusCode { get; set; }
        public string Description { get; set; }
        public EntityReference Customer { get; set; }
        public Entity RawEntity { get; set; }
    }

    public class FindingData
    {
        public Guid Id { get; set; }
        public string Finding { get; set; }
        public OptionSetValue FindingType { get; set; }
        public OptionSetValue FinalEnforcementAction { get; set; }
        public OptionSetValue SensitivityLevel { get; set; }
        public OptionSetValue StatusCode { get; set; }
        public EntityReference Incident { get; set; }
        public EntityReference AccountId { get; set; }
        public Entity RawEntity { get; set; }
    }

    public class ActionData
    {
        public Guid Id { get; set; }
        public string Name { get; set; }
        public OptionSetValue ActionCategory { get; set; }
        public OptionSetValue ActionType { get; set; }
        public EntityReference Case { get; set; }
        public Entity RawEntity { get; set; }
    }

    public class DocumentsData
    {
        public List<DocumentData> GeneralDocuments { get; set; }
        public List<DocumentData> InspectionDocuments { get; set; }
    }

    public class DocumentData
    {
        public Guid Id { get; set; }
        public string Name { get; set; }
        public EntityReference FileCategory { get; set; }
        public EntityReference FileSubCategory { get; set; }
        public OptionSetValue FileContext { get; set; }
        public string Description { get; set; }
        public DateTime CreatedOn { get; set; }
        public string SharePointLink { get; set; }
        public Entity RawEntity { get; set; }
    }

    public class InteractionData
    {
        public Guid Id { get; set; }
        public string Subject { get; set; }
        public string ActivityType { get; set; }
        public OptionSetValue StateCode { get; set; }
        public OptionSetValue StatusCode { get; set; }
        public DateTime CreatedOn { get; set; }
        public DateTime? ScheduledStart { get; set; }
        public string Description { get; set; }
        public bool IsNote { get; set; }
        public Entity RawEntity { get; set; }
    }

    public class ContactData
    {
        public Guid Id { get; set; }
        public string FullName { get; set; }
        public string Email { get; set; }
        public string Phone { get; set; }
        public string JobTitle { get; set; }
        public Entity RawEntity { get; set; }
    }

    public class InspectorData
    {
        public Guid Id { get; set; }
        public string FullName { get; set; }
        public string Title { get; set; }
        public string Email { get; set; }
        public Entity RawEntity { get; set; }
    }

    public class SupportingRegionData
    {
        public Guid Id { get; set; }
        public EntityReference Region { get; set; }
        public Entity RawEntity { get; set; }
    }

    #endregion
}


