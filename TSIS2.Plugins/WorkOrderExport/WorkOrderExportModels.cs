using System.Collections.Generic;

namespace TSIS2.Plugins.WorkOrderExport
{
    /// <summary>
    /// Wrapper for the final JSON payload stored in ts_payloadjson.
    /// </summary>
    public class WorkOrderExportPayload
    {
        public List<WorkOrderWordTemplateModel> WorkOrders { get; set; } = new List<WorkOrderWordTemplateModel>();
    }

    /// <summary>
    /// Represents the data structure for the Word Template JSON payload.
    /// Property names MUST match the keys used in the Power Automate Flow / Word Template.
    /// </summary>
    public class WorkOrderWordTemplateModel
    {
        // --- Single Fields ---
        public string WorkOrderId { get; set; }
        public string WorkOrderNumber { get; set; }
        public string WorkOrderDate { get; set; }
        public string RegionName { get; set; }
        public string OperationType { get; set; }
        public string StakeholderName { get; set; }
        public string SiteName { get; set; }
        public string WorkLocation { get; set; }
        public string PrimaryIncidentDescription { get; set; }
        public string BusinessOwner { get; set; }

        // --- Repeating Sections ---
        public List<FindingModel> Findings { get; set; } = new List<FindingModel>();
        public List<ActionModel> Actions { get; set; } = new List<ActionModel>();
        public List<InteractionModel> Interactions { get; set; } = new List<InteractionModel>();
        public List<ServiceTaskModel> ServiceTasks { get; set; } = new List<ServiceTaskModel>();
        public List<DocumentModel> Documents { get; set; } = new List<DocumentModel>();
        public List<ContactModel> Contacts { get; set; } = new List<ContactModel>();
    }

    public class FindingModel
    {
        public string Finding_Description { get; set; }
        public string Finding_Type { get; set; }
        public string Finding_Status { get; set; }
        public string Finding_Enforcement { get; set; }
        public string Finding_Sensitivity { get; set; }
    }

    public class ActionModel
    {
        public string Action_Name { get; set; }
        public string Action_Category { get; set; }
        public string Action_Type { get; set; }
    }

    public class InteractionModel
    {
        public string Interaction_Date { get; set; }
        public string Interaction_Type { get; set; }
        public string Interaction_Subject { get; set; }
        public string Interaction_Description { get; set; }
    }

    public class ServiceTaskModel
    {
        public string Task_Name { get; set; }
        public string Task_Status { get; set; }
        public string Task_InspectionResult { get; set; }
    }

    public class DocumentModel
    {
        public string Doc_Name { get; set; }
        public string Doc_Category { get; set; }
        public string Doc_Context { get; set; }
        public string Doc_Link { get; set; }
    }

    public class ContactModel
    {
        public string Contact_Name { get; set; }
        public string Contact_JobTitle { get; set; }
        public string Contact_Email { get; set; }
        public string Contact_Phone { get; set; }
    }
}
