using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace TSIS2.Plugins
{
    /// <summary>
    /// Synchronizes Access Team membership between:
    /// 1. Unplanned Work Order Access Team, and
    /// 2. Related Work Order Access Team
    ///
    /// Triggered when a user is added via the custom API ts_AddUserToAccessTeam.
    ///
    /// This ensures that adding a user to an Unplanned Work Order automatically
    /// adds them to the corresponding Work Order team.
    /// </summary>
    [CrmPluginRegistration(
        "ts_AddUserToAccessTeam",
        "systemuser",
        StageEnum.PostOperation,
        ExecutionModeEnum.Synchronous,
        "",
        "TSIS2.Plugins.PostOperation_SyncAccessTeam_AddUser_UnplannedWorkOrder_to_WorkOrder",
        1,
        IsolationModeEnum.Sandbox,
        Description = "Synchronizes Access Team users from Unplanned Work Order to Work Order when a user is added.")]
    public class PostOperation_SyncAccessTeam_AddUser_UnplannedWorkOrder_to_WorkOrder : IPlugin
    {
        private const string UnplannedWorkOrderEntity = "ts_unplannedworkorder";
        private const string UnplannedWorkOrder_WorkOrderLookup = "ts_workorder";
        private const string WorkOrderTeamTemplateId = "bddf1d45-706d-ec11-8f8e-0022483da5aa";

        public void Execute(IServiceProvider serviceProvider)
        {
            var tracer = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            var context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            var serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            var service = serviceFactory.CreateOrganizationService(context.UserId);

            tracer.Trace("PostOperation_SyncAccessTeam_AddUser_WO_UnplannedWO started.");

            try
            {

                if (context.PrimaryEntityId == Guid.Empty)
                    throw new InvalidPluginExecutionException("PrimaryEntityId (systemuser) is required.");

                var systemUserId = context.PrimaryEntityId;

                if (!context.InputParameters.Contains("Record") || !(context.InputParameters["Record"] is EntityReference record))
                    throw new InvalidPluginExecutionException("Record parameter is required.");

                if (!context.InputParameters.Contains("TeamTemplate") || !(context.InputParameters["TeamTemplate"] is EntityReference teamTemplate))
                    throw new InvalidPluginExecutionException("TeamTemplate parameter is required.");

                tracer.Trace($"SystemUser: {systemUserId}");
                tracer.Trace($"Record: {record.LogicalName} {record.Id}");
                tracer.Trace($"TeamTemplate: {teamTemplate.Id}");

                if (record.LogicalName == UnplannedWorkOrderEntity)
                {
                    tracer.Trace("Unplanned Work Order detected. Adding user to 2 teams.");


                    // Create AddUserToRecordTeam request for Unplanned Work Order
                    var upwoRequest = new OrganizationRequest("AddUserToRecordTeam")
                    {
                        ["SystemUserId"] = systemUserId,
                        ["Record"] = record,
                        ["TeamTemplateId"] = teamTemplate.Id   // Template passed by custom API
                    };
                    service.Execute(upwoRequest);
                    tracer.Trace("Added to Unplanned Work Order access team.");

                    var unplanned = service.Retrieve(record.LogicalName, record.Id, new ColumnSet(UnplannedWorkOrder_WorkOrderLookup));

                    if (!unplanned.Contains(UnplannedWorkOrder_WorkOrderLookup))
                        throw new InvalidPluginExecutionException("Unplanned Work Order missing required Work Order lookup.");

                    var workOrderRef = unplanned.GetAttributeValue<EntityReference>(UnplannedWorkOrder_WorkOrderLookup);
                    tracer.Trace($"Related Work Order: {workOrderRef.Id}");

                    // Create AddUserToRecordTeam request for Work Order
                    var woRequest = new OrganizationRequest("AddUserToRecordTeam")
                    {
                        ["SystemUserId"] = systemUserId,
                        ["Record"] = workOrderRef,
                        ["TeamTemplateId"] = Guid.Parse(WorkOrderTeamTemplateId)
                    };
                    service.Execute(woRequest);

                    tracer.Trace("Added to related Work Order access team.");
                }
                else
                {
                    tracer.Trace("Unsupported entity type. No action taken.");
                }
            }
            catch (Exception ex)
            {
                tracer.Trace($"Exception: {ex}");
                throw new InvalidPluginExecutionException("PostOperation_SyncAccessTeam_AddUser_UnplannedWorkOrder_to_WorkOrder failed.", ex);
            }

            tracer.Trace("PostOperation_SyncAccessTeam_AddUser_UnplannedWorkOrder_to_WorkOrder finished.");
        }
    }

    /// <summary>
    /// Ensures user removal from Unplanned Work Order team ALSO removes
    /// them from the related Work Order Access Team.
    /// </summary>
    [CrmPluginRegistration(
        MessageNameEnum.RemoveUserFromRecordTeam,
        "none",
        StageEnum.PostOperation,
        ExecutionModeEnum.Synchronous,
        "",
        "TSIS2.Plugins.PostOperation_SyncAccessTeam_RemoveUser_UnplannedWorkOrder_to_WorkOrder",
        1,
        IsolationModeEnum.Sandbox)]
    public class PostOperation_SyncAccessTeam_RemoveUser_UnplannedWorkOrder_to_WorkOrder : IPlugin
    {
        private const string UnplannedWorkOrderEntity = "ts_unplannedworkorder";
        private const string WorkOrderLookup = "ts_workorder";
        private const string WorkOrderTeamTemplateId = "bddf1d45-706d-ec11-8f8e-0022483da5aa";

        public void Execute(IServiceProvider serviceProvider)
        {
            var trace = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            var context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            var serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            var service = serviceFactory.CreateOrganizationService(context.UserId);

            trace.Trace("=== RemoveUser plugin STARTED ===");

            try
            {
                // Validate expected input parameters
                if (!context.InputParameters.Contains("SystemUserId") ||
                    !context.InputParameters.Contains("Record") ||
                    !context.InputParameters.Contains("TeamTemplateId"))
                {
                    trace.Trace("Missing required parameters. EXIT.");
                    return;
                }

                var userId = (Guid)context.InputParameters["SystemUserId"];
                var recordRef = (EntityReference)context.InputParameters["Record"];
                var templateId = (Guid)context.InputParameters["TeamTemplateId"];

                trace.Trace("Parameters received:");
                trace.Trace($" - UserId = {userId}");
                trace.Trace($" - Record = {recordRef.LogicalName} ({recordRef.Id})");
                trace.Trace($" - TeamTemplateId = {templateId}");

                // Ensure plugin only runs when removing user from UNPLANNED WORK ORDER team
                if (recordRef.LogicalName != UnplannedWorkOrderEntity)
                {
                    trace.Trace($"Record is of type {recordRef.LogicalName}. Sync only applies to ts_unplannedworkorder. EXIT.");
                    return;
                }

                // Retrieve related Work Order (ts_workorder lookup)
                var unplanned = service.Retrieve(UnplannedWorkOrderEntity, recordRef.Id,
                    new ColumnSet(WorkOrderLookup));

                var woRef = unplanned.GetAttributeValue<EntityReference>(WorkOrderLookup);
                if (woRef != null)
                {
                    trace.Trace($"Linked Work Order found: {woRef.LogicalName} ({woRef.Id})");

                    // Remove user from Work Order team
                    var removeFromWO = new OrganizationRequest("RemoveUserFromRecordTeam")
                    {
                        ["SystemUserId"] = userId,
                        ["Record"] = woRef,
                        ["TeamTemplateId"] = Guid.Parse(WorkOrderTeamTemplateId)
                    };
                    service.Execute(removeFromWO);

                    trace.Trace("User removed from Work Order access team.");
                }
            }
            catch (Exception ex)
            {
                trace.Trace("ERROR: " + ex);
                throw;
            }
        }
    }
}