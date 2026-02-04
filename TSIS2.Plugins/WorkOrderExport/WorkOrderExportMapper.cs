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
            
            var model = new WorkOrderWordTemplateModel
            {
                // Single Fields
                WorkOrderId = data.WorkOrderId.ToString(),
                WorkOrderNumber = summary.Name,
                WorkOrderDate = summary.CreatedOn.ToString("yyyy-MM-dd"),
                RegionName = summary.Region?.Name ?? "",
                OperationType = summary.OperationType?.Name ?? "",
                StakeholderName = summary.Stakeholder?.Name ?? "",
                SiteName = summary.Site?.Name ?? "",
                WorkLocation = GetOptionSetLabel(summary.RawEntity, "msdyn_worklocation") ?? summary.WorkLocation?.Value.ToString() ?? "",
                PrimaryIncidentDescription = summary.PrimaryIncidentDescription ?? "",
                BusinessOwner = summary.BusinessOwner ?? "",
                
                // Collections
                Findings = data.Findings?.Select(f => new FindingModel
                {
                    Finding_Description = f.Finding ?? "",
                    Finding_Type = GetOptionSetLabel(f.RawEntity, "ts_findingtype") ?? f.FindingType?.Value.ToString() ?? "",
                    Finding_Status = GetOptionSetLabel(f.RawEntity, "statuscode") ?? f.StatusCode?.Value.ToString() ?? "",
                    Finding_Enforcement = GetOptionSetLabel(f.RawEntity, "ts_finalenforcementaction") ?? f.FinalEnforcementAction?.Value.ToString() ?? "",
                    Finding_Sensitivity = GetOptionSetLabel(f.RawEntity, "ts_sensitivitylevel") ?? f.SensitivityLevel?.Value.ToString() ?? ""
                }).ToList() ?? new List<FindingModel>(),

                Actions = data.Actions?.Select(a => new ActionModel
                {
                    Action_Name = a.Name ?? "",
                    Action_Category = GetOptionSetLabel(a.RawEntity, "ts_actioncategory") ?? a.ActionCategory?.Value.ToString() ?? "",
                    Action_Type = GetOptionSetLabel(a.RawEntity, "ts_actiontype") ?? a.ActionType?.Value.ToString() ?? ""
                }).ToList() ?? new List<ActionModel>(),

                Interactions = data.Interactions?.Select(i => new InteractionModel
                {
                    Interaction_Date = i.CreatedOn.ToString("yyyy-MM-dd HH:mm"),
                    Interaction_Type = i.ActivityType ?? "Note",
                    Interaction_Subject = i.Subject ?? "",
                    Interaction_Description = i.Description ?? ""
                }).ToList() ?? new List<InteractionModel>(),

                ServiceTasks = data.ServiceTasks?.Select(t => new ServiceTaskModel
                {
                    Task_Name = t.Name ?? "",
                    Task_Status = GetOptionSetLabel(t.RawEntity, "statuscode") ?? t.StatusCode?.Value.ToString() ?? "",
                    Task_InspectionResult = GetOptionSetLabel(t.RawEntity, "msdyn_inspectiontaskresult") ?? t.InspectionResult?.Value.ToString() ?? ""
                }).ToList() ?? new List<ServiceTaskModel>(),

                Contacts = data.Contacts?.Select(c => new ContactModel
                {
                    Contact_Name = c.FullName ?? "",
                    Contact_JobTitle = c.JobTitle ?? "",
                    Contact_Email = c.Email ?? "",
                    Contact_Phone = c.Phone ?? ""
                }).ToList() ?? new List<ContactModel>()
            };

            // Combine Documents
            var allDocs = new List<DocumentModel>();
            if (data.Documents?.GeneralDocuments != null)
            {
                allDocs.AddRange(data.Documents.GeneralDocuments.Select(d => new DocumentModel
                {
                    Doc_Name = d.Name ?? "",
                    Doc_Category = d.FileCategory?.Name ?? "",
                    Doc_Context = GetOptionSetLabel(d.RawEntity, "ts_filecontext") ?? d.FileContext?.Value.ToString() ?? "",
                    Doc_Link = d.SharePointLink ?? ""
                }));
            }
            if (data.Documents?.InspectionDocuments != null)
            {
                allDocs.AddRange(data.Documents.InspectionDocuments.Select(d => new DocumentModel
                {
                    Doc_Name = d.Name ?? "",
                    Doc_Category = d.FileCategory?.Name ?? "Inspection",
                    Doc_Context = GetOptionSetLabel(d.RawEntity, "ts_filecontext") ?? d.FileContext?.Value.ToString() ?? "",
                    Doc_Link = d.SharePointLink ?? ""
                }));
            }
            model.Documents = allDocs;

            return model;
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
