using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Xrm.Sdk;

namespace TSIS2.Plugins.WorkOrderExport
{
    public class WorkOrderExportMapper
    {
        private readonly ITracingService _tracingService;

        public WorkOrderExportMapper(ITracingService tracingService)
        {
            _tracingService = tracingService ?? throw new ArgumentNullException(nameof(tracingService));
        }

        public WorkOrderWordTemplateModel Map(WorkOrderData data)
        {
            if (data == null) throw new ArgumentNullException(nameof(data));

            var summary = data.WorkOrderSummary;
            var findings = data.Findings?.Select(f => new FindingModel
            {
                Finding_Description = S(f.Finding),
                Finding_Type = LabelOrValue(f.RawEntity, "ts_findingtype", f.FindingType?.Value),
                Finding_Status = LabelOrValue(f.RawEntity, "statuscode", f.StatusCode?.Value),
                Finding_Enforcement = LabelOrValue(f.RawEntity, "ts_finalenforcementaction", f.FinalEnforcementAction?.Value),
                Finding_Sensitivity = LabelOrValue(f.RawEntity, "ts_sensitivitylevel", f.SensitivityLevel?.Value)
            }).ToList() ?? new List<FindingModel>();

            var actions = data.Actions?.Select(a => new ActionModel
            {
                Action_Name = S(a.Name),
                Action_Category = LabelOrValue(a.RawEntity, "ts_actioncategory", a.ActionCategory?.Value),
                Action_Type = LabelOrValue(a.RawEntity, "ts_actiontype", a.ActionType?.Value)
            }).ToList() ?? new List<ActionModel>();

            var interactions = data.Interactions?.Select(i => new InteractionModel
            {
                Interaction_Date = i.CreatedOn.ToString("yyyy-MM-dd HH:mm"),
                Interaction_Type = i.ActivityType ?? "Note",
                Interaction_Subject = S(i.Subject),
                Interaction_Description = S(i.Description)
            }).ToList() ?? new List<InteractionModel>();

            var workOrderDocuments = MapWorkOrderDocuments(data.Documents?.WorkOrderDocuments);
            var inspectionDocuments = MapInspectionDocuments(data.Documents?.InspectionDocuments);
            var orderedServiceTasks = data.ServiceTasks ?? new List<ServiceTaskData>();
            var serviceTasks = orderedServiceTasks.Select(MapServiceTask).ToList();
            var serviceTaskDocuments = MapServiceTaskDocuments(orderedServiceTasks);

            var contacts = data.Contacts?.Select(c => new ContactModel
            {
                Contact_Name = S(c.FullName),
                Contact_JobTitle = S(c.JobTitle),
                Contact_Email = S(c.Email),
                Contact_Phone = S(c.Phone)
            }).ToList() ?? new List<ContactModel>();

            var model = new WorkOrderWordTemplateModel
            {
                // Single Fields
                WorkOrderId = data.WorkOrderId.ToString(),
                WorkOrderNumber = summary.Name,
                WorkOrderDate = summary.CreatedOn.ToString("yyyy-MM-dd"),
                RegionName = S(summary.Region?.Name),
                OperationType = S(summary.OperationType?.Name),
                StakeholderName = S(summary.Stakeholder?.Name),
                SiteName = S(summary.Site?.Name),
                WorkLocation = LabelOrValue(summary.RawEntity, "msdyn_worklocation", summary.WorkLocation?.Value),
                PrimaryIncidentDescription = S(summary.PrimaryIncidentDescription),
                BusinessOwner = S(summary.BusinessOwner),
                Findings = findings,
                Actions = actions,
                Interactions = interactions,
                WorkOrderDocuments = workOrderDocuments,
                InspectionDocuments = inspectionDocuments,
                ServiceTasks = serviceTasks,
                ServiceTaskDocuments = serviceTaskDocuments,
                Contacts = contacts,
                HasFindings = findings.Count > 0,
                NoFindingsMessage = GetNoItemsMessage(findings.Count > 0, "Findings"),
                HasActions = actions.Count > 0,
                NoActionsMessage = GetNoItemsMessage(actions.Count > 0, "Actions"),
                HasInteractions = interactions.Count > 0,
                NoInteractionsMessage = GetNoItemsMessage(interactions.Count > 0, "Interactions"),
                HasServiceTasks = serviceTasks.Count > 0,
                NoServiceTasksMessage = GetNoItemsMessage(serviceTasks.Count > 0, "Service Tasks"),
                HasServiceTaskDocuments = serviceTaskDocuments.Count > 0,
                NoServiceTaskDocumentsMessage = GetNoItemsMessage(serviceTaskDocuments.Count > 0, "Service Task Documents"),
                HasContacts = contacts.Count > 0,
                NoContactsMessage = GetNoItemsMessage(contacts.Count > 0, "Contacts"),
                HasWorkOrderDocuments = workOrderDocuments.Count > 0,
                NoWorkOrderDocumentsMessage = GetNoItemsMessage(workOrderDocuments.Count > 0, "Work Order Documents"),
                HasInspectionDocuments = inspectionDocuments.Count > 0,
                NoInspectionDocumentsMessage = GetNoItemsMessage(inspectionDocuments.Count > 0, "Inspection Documents")
            };

            return model;
        }

        private static string S(string value)
        {
            return value ?? string.Empty;
        }

        private ServiceTaskModel MapServiceTask(ServiceTaskData task)
        {
            return new ServiceTaskModel
            {
                Task_Id = task.Id.ToString(),
                Task_Name = S(task.Name),
                Task_Status = LabelOrValue(task.RawEntity, "statuscode", task.StatusCode?.Value),
                Task_InspectionResult = LabelOrValue(task.RawEntity, "msdyn_inspectiontaskresult", task.InspectionResult?.Value)
            };
        }

        private List<WorkOrderDocumentModel> MapWorkOrderDocuments(List<DocumentData> documents)
        {
            var mapped = (documents ?? new List<DocumentData>())
                .Select(d => new WorkOrderDocumentModel
                {
                    WorkOrderDocument_Name = S(d.Name),
                    WorkOrderDocument_Category = S(d.FileCategory?.Name),
                    WorkOrderDocument_Context = LabelOrValue(d.RawEntity, "ts_filecontext", d.FileContext?.Value),
                    WorkOrderDocument_Link = S(d.SharePointLink)
                })
                .OrderBy(d => d.WorkOrderDocument_Name, StringComparer.OrdinalIgnoreCase)
                .ThenBy(d => d.WorkOrderDocument_Link, StringComparer.OrdinalIgnoreCase)
                .ToList();

            return DeduplicateDocuments(
                mapped,
                d => (d.WorkOrderDocument_Name ?? "") + "|" + (d.WorkOrderDocument_Link ?? ""));
        }

        private List<InspectionDocumentModel> MapInspectionDocuments(List<DocumentData> documents)
        {
            var mapped = (documents ?? new List<DocumentData>())
                .Select(d => new InspectionDocumentModel
                {
                    InspectionDocument_Name = S(d.Name),
                    InspectionDocument_Category = S(d.FileCategory?.Name),
                    InspectionDocument_Context = LabelOrValue(d.RawEntity, "ts_filecontext", d.FileContext?.Value),
                    InspectionDocument_Link = S(d.SharePointLink)
                })
                .OrderBy(d => d.InspectionDocument_Name, StringComparer.OrdinalIgnoreCase)
                .ThenBy(d => d.InspectionDocument_Link, StringComparer.OrdinalIgnoreCase)
                .ToList();

            return DeduplicateDocuments(
                mapped,
                d => (d.InspectionDocument_Name ?? "") + "|" + (d.InspectionDocument_Link ?? ""));
        }

        private List<ServiceTaskDocumentModel> MapServiceTaskDocuments(List<ServiceTaskData> tasks)
        {
            var flattened = new List<ServiceTaskDocumentModel>();

            foreach (var task in tasks ?? new List<ServiceTaskData>())
            {
                var orderedDocs = (task.Documents ?? new List<DocumentData>())
                    .OrderBy(d => d.CreatedOn)
                    .ThenBy(d => d.Name ?? string.Empty, StringComparer.OrdinalIgnoreCase)
                    .ThenBy(d => d.Id)
                    .ToList();

                var uniqueDocs = DeduplicateDocuments(
                    orderedDocs,
                    d => (d.Name ?? "") + "|" + (d.SharePointLink ?? ""));

                foreach (var doc in uniqueDocs)
                {
                    flattened.Add(new ServiceTaskDocumentModel
                    {
                        Task_Name = S(task.Name),
                        ServiceTaskDocument_Name = S(doc.Name),
                        ServiceTaskDocument_Category = S(doc.FileCategory?.Name),
                        ServiceTaskDocument_Context = LabelOrValue(doc.RawEntity, "ts_filecontext", doc.FileContext?.Value),
                        ServiceTaskDocument_Link = S(doc.SharePointLink)
                    });
                }
            }

            return flattened;
        }

        private static List<T> DeduplicateDocuments<T>(IEnumerable<T> documents, Func<T, string> keySelector)
        {
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var unique = new List<T>();
            foreach (var doc in documents ?? Enumerable.Empty<T>())
            {
                if (seen.Add(keySelector(doc) ?? string.Empty))
                {
                    unique.Add(doc);
                }
            }
            return unique;
        }

        private string LabelOrValue(Entity entity, string attributeName, int? value)
        {
            return GetOptionSetLabel(entity, attributeName) ?? value?.ToString() ?? string.Empty;
        }

        private static string GetNoItemsMessage(bool hasItems, string sectionName)
        {
            return hasItems ? string.Empty : $"No {sectionName}";
        }

        private string GetOptionSetLabel(Entity entity, string attributeName)
        {
            if (entity != null && entity.FormattedValues.Contains(attributeName))
            {
                return entity.FormattedValues[attributeName];
            }
            return null;
        }
    }
}
