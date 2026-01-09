using System;
using System.Linq;
using System.ServiceModel;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;

namespace TSIS2.Plugins
{
    [CrmPluginRegistration(
        MessageNameEnum.Associate,
        "none",
        StageEnum.PostOperation,
        ExecutionModeEnum.Synchronous,
        "",
        "TSIS2.Plugins.PostOperation_UnplannedWO-WorkOrderTimeTrackingSync_Associate",
        1,
        IsolationModeEnum.Sandbox,
        Description = "Synchronizes work order time tracking from Unplanned Work Order to Work Order on association.")]
    public class PostOperation_WorkOrderTimeTrackingSync_Associate : PluginBase
    {
        private const string UnplannedWorkOrderEntity = "ts_unplannedworkorder";
        private const string WorkOrderEntity = "msdyn_workorder";
        private const string WorkOrderTimeTrackingEntity = "ts_workordertimetracking";

        // Lookup field on unplanned work order that points to the msdyn_workorder
        private const string UnplannedWorkOrder_WorkOrderLookup = "ts_workorder";

        // Relationship schema names
        // Unplanned Work Order <-> Work Order Time Tracking M:N relationship schema name
        private const string UnplannedWorkOrder_TimeTracking_Relationship = "ts_unplannedworkorder_UnplannedWorkOrder_ts_workordertimetracking";
        // Work Order <-> Work Order Time Tracking M:N relationship schema name to mirror onto
        private const string WorkOrder_TimeTracking_Relationship = "ts_workorder_ts_workordertimetracking_WorkOrder";

        public PostOperation_WorkOrderTimeTrackingSync_Associate(string unsecure, string secure)
            : base(typeof(PostOperation_WorkOrderTimeTrackingSync_Associate))
        {
        }

        protected override void ExecuteCrmPlugin(LocalPluginContext localContext)
        {
            if (localContext == null)
                throw new InvalidPluginExecutionException("localContext");

            var context = localContext.PluginExecutionContext;
            var service = localContext.OrganizationService;
            var trace = localContext.TracingService;

            try
            {
                trace.Trace("Work Order Time Tracking Associate sync start. CorrelationId={0}, Depth={1}", context.CorrelationId, context.Depth);

                // Guard: avoid recursion
                if (context.Depth > 1)
                {
                    trace.Trace("Depth > 1 detected. Exiting to prevent recursion.");
                    return;
                }

                if (!(context.InputParameters.Contains("Relationship") &&
                      context.InputParameters["Relationship"] is Relationship relationship))
                {
                    trace.Trace("Missing Relationship parameter. Exiting.");
                    return;
                }

                if (!string.Equals(relationship.SchemaName, UnplannedWorkOrder_TimeTracking_Relationship, StringComparison.OrdinalIgnoreCase))
                {
                    trace.Trace("Relationship {0} does not match expected {1}. Exiting.",
                        relationship.SchemaName, UnplannedWorkOrder_TimeTracking_Relationship);
                    return;
                }

                if (!(context.InputParameters.Contains("Target") &&
                      context.InputParameters["Target"] is EntityReference target))
                {
                    trace.Trace("Missing Target parameter. Exiting.");
                    return;
                }

                if (!(context.InputParameters.Contains("RelatedEntities") &&
                      context.InputParameters["RelatedEntities"] is EntityReferenceCollection relatedEntities) ||
                    relatedEntities.Count == 0)
                {
                    trace.Trace("Missing RelatedEntities parameter. Exiting.");
                    return;
                }

                // Determine which side is the unplanned work order and which side is/are the time tracking record(s)
                EntityReference unplannedWorkOrderRef = null;
                var timeTrackingRefs = new EntityReferenceCollection();

                if (string.Equals(target.LogicalName, UnplannedWorkOrderEntity, StringComparison.OrdinalIgnoreCase))
                {
                    unplannedWorkOrderRef = target;
                    foreach (var re in relatedEntities.Where(e => e.LogicalName == WorkOrderTimeTrackingEntity))
                        timeTrackingRefs.Add(re);
                }
                else if (string.Equals(target.LogicalName, WorkOrderTimeTrackingEntity, StringComparison.OrdinalIgnoreCase))
                {
                    var reUnplannedWorkOrder = relatedEntities.FirstOrDefault(e => e.LogicalName == UnplannedWorkOrderEntity);
                    if (reUnplannedWorkOrder != null)
                    {
                        unplannedWorkOrderRef = reUnplannedWorkOrder;
                        timeTrackingRefs.Add(target);
                    }
                }

                if (unplannedWorkOrderRef == null || timeTrackingRefs.Count == 0)
                {
                    trace.Trace("Could not resolve unplanned work order and time tracking record(s) from the association. Exiting.");
                    return;
                }

                // Retrieve the linked msdyn_workorder from the unplanned work order
                var unplannedWorkOrder = service.Retrieve(UnplannedWorkOrderEntity, unplannedWorkOrderRef.Id, new ColumnSet(UnplannedWorkOrder_WorkOrderLookup));
                var workOrderRef = unplannedWorkOrder.GetAttributeValue<EntityReference>(UnplannedWorkOrder_WorkOrderLookup);
                if (workOrderRef == null)
                {
                    trace.Trace("Unplanned Work Order {0} does not have {1} populated. Exiting.", unplannedWorkOrderRef.Id, UnplannedWorkOrder_WorkOrderLookup);
                    return;
                }

                // Mirror the association(s) on the Work Order <-> Time Tracking relationship
                var relationshipToMirror = new Relationship(WorkOrder_TimeTracking_Relationship);

                foreach (var timeTrackingRef in timeTrackingRefs)
                {
                    try
                    {
                        var associate = new AssociateRequest
                        {
                            Target = new EntityReference(WorkOrderEntity, workOrderRef.Id),
                            Relationship = relationshipToMirror,
                            RelatedEntities = new EntityReferenceCollection { new EntityReference(WorkOrderTimeTrackingEntity, timeTrackingRef.Id) }
                        };

                        service.Execute(associate);
                        trace.Trace("Mirrored association WorkOrder({0}) <-> TimeTracking({1}) via relationship {2}.",
                            workOrderRef.Id, timeTrackingRef.Id, WorkOrder_TimeTracking_Relationship);
                    }
                    catch (FaultException<OrganizationServiceFault> ex)
                    {
                        var msg = ex.Detail?.Message ?? string.Empty;
                        if (msg.IndexOf("already associated", StringComparison.OrdinalIgnoreCase) >= 0 ||
                            msg.IndexOf("duplicate", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            trace.Trace("Association already exists for WorkOrder({0}) and TimeTracking({1}); ignoring.", workOrderRef.Id, timeTrackingRef.Id);
                        }
                        else
                        {
                            trace.Trace("Associate fault for WorkOrder({0}) and TimeTracking({1}): {2}", workOrderRef.Id, timeTrackingRef.Id, ex);
                            throw;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                localContext.TraceWithContext("Exception: {0}", ex.Message);
                throw new InvalidPluginExecutionException("PostOperation_WorkOrderTimeTrackingSync_Associate failed.", ex);
            }
        }
    }

    [CrmPluginRegistration(
        MessageNameEnum.Disassociate,
        "none",
        StageEnum.PostOperation,
        ExecutionModeEnum.Synchronous,
        "",
        "TSIS2.Plugins.PostOperation_UnplannedWO-WorkOrderTimeTrackingSync_Disassociate",
        1,
        IsolationModeEnum.Sandbox,
        Description = "Synchronizes work order time tracking removal from Unplanned Work Order to Work Order on disassociation.")]
    public class PostOperation_WorkOrderTimeTrackingSync_Disassociate : PluginBase
    {
        private const string UnplannedWorkOrderEntity = "ts_unplannedworkorder";
        private const string WorkOrderEntity = "msdyn_workorder";
        private const string WorkOrderTimeTrackingEntity = "ts_workordertimetracking";

        // Lookup field on unplanned work order that points to the msdyn_workorder
        private const string UnplannedWorkOrder_WorkOrderLookup = "ts_workorder";

        // Relationship schema names
        private const string UnplannedWorkOrder_TimeTracking_Relationship = "ts_unplannedworkorder_UnplannedWorkOrder_ts_workordertimetracking";
        private const string WorkOrder_TimeTracking_Relationship = "ts_workorder_ts_workordertimetracking_WorkOrder";

        public PostOperation_WorkOrderTimeTrackingSync_Disassociate(string unsecure, string secure)
            : base(typeof(PostOperation_WorkOrderTimeTrackingSync_Disassociate))
        {
        }

        protected override void ExecuteCrmPlugin(LocalPluginContext localContext)
        {
            if (localContext == null)
                throw new InvalidPluginExecutionException("localContext");

            var context = localContext.PluginExecutionContext;
            var service = localContext.OrganizationService;
            var trace = localContext.TracingService;

            try
            {
                trace.Trace("Work Order Time Tracking Disassociate sync start. CorrelationId={0}, Depth={1}", context.CorrelationId, context.Depth);

                // Guard: avoid recursion
                if (context.Depth > 1)
                {
                    trace.Trace("Depth > 1 detected. Exiting to prevent recursion.");
                    return;
                }

                if (!(context.InputParameters.Contains("Relationship") &&
                      context.InputParameters["Relationship"] is Relationship relationship))
                {
                    trace.Trace("Missing Relationship parameter. Exiting.");
                    return;
                }

                if (!string.Equals(relationship.SchemaName, UnplannedWorkOrder_TimeTracking_Relationship, StringComparison.OrdinalIgnoreCase))
                {
                    trace.Trace("Relationship {0} does not match expected {1}. Exiting.",
                        relationship.SchemaName, UnplannedWorkOrder_TimeTracking_Relationship);
                    return;
                }

                if (!(context.InputParameters.Contains("Target") &&
                      context.InputParameters["Target"] is EntityReference target))
                {
                    trace.Trace("Missing Target parameter. Exiting.");
                    return;
                }

                if (!(context.InputParameters.Contains("RelatedEntities") &&
                      context.InputParameters["RelatedEntities"] is EntityReferenceCollection relatedEntities) ||
                    relatedEntities.Count == 0)
                {
                    trace.Trace("Missing RelatedEntities parameter. Exiting.");
                    return;
                }

                // Determine which side is the unplanned work order and which side is/are the time tracking record(s)
                EntityReference unplannedWorkOrderRef = null;
                var timeTrackingRefs = new EntityReferenceCollection();

                if (string.Equals(target.LogicalName, UnplannedWorkOrderEntity, StringComparison.OrdinalIgnoreCase))
                {
                    unplannedWorkOrderRef = target;
                    foreach (var re in relatedEntities.Where(e => e.LogicalName == WorkOrderTimeTrackingEntity))
                        timeTrackingRefs.Add(re);
                }
                else if (string.Equals(target.LogicalName, WorkOrderTimeTrackingEntity, StringComparison.OrdinalIgnoreCase))
                {
                    var reUnplannedWorkOrder = relatedEntities.FirstOrDefault(e => e.LogicalName == UnplannedWorkOrderEntity);
                    if (reUnplannedWorkOrder != null)
                    {
                        unplannedWorkOrderRef = reUnplannedWorkOrder;
                        timeTrackingRefs.Add(target);
                    }
                }

                if (unplannedWorkOrderRef == null || timeTrackingRefs.Count == 0)
                {
                    trace.Trace("Could not resolve unplanned work order and time tracking record(s) from the disassociation. Exiting.");
                    return;
                }

                // Retrieve the linked msdyn_workorder from the unplanned work order
                var unplannedWorkOrder = service.Retrieve(UnplannedWorkOrderEntity, unplannedWorkOrderRef.Id, new ColumnSet(UnplannedWorkOrder_WorkOrderLookup));
                var workOrderRef = unplannedWorkOrder.GetAttributeValue<EntityReference>(UnplannedWorkOrder_WorkOrderLookup);
                if (workOrderRef == null)
                {
                    trace.Trace("Unplanned Work Order {0} does not have {1} populated. Exiting.", unplannedWorkOrderRef.Id, UnplannedWorkOrder_WorkOrderLookup);
                    return;
                }

                // Mirror the disassociation(s) on the Work Order <-> Time Tracking relationship
                var relationshipToMirror = new Relationship(WorkOrder_TimeTracking_Relationship);

                foreach (var timeTrackingRef in timeTrackingRefs)
                {
                    try
                    {
                        var disassociate = new DisassociateRequest
                        {
                            Target = new EntityReference(WorkOrderEntity, workOrderRef.Id),
                            Relationship = relationshipToMirror,
                            RelatedEntities = new EntityReferenceCollection { new EntityReference(WorkOrderTimeTrackingEntity, timeTrackingRef.Id) }
                        };

                        service.Execute(disassociate);
                        trace.Trace("Mirrored disassociation WorkOrder({0}) !- TimeTracking({1}) via relationship {2}.",
                            workOrderRef.Id, timeTrackingRef.Id, WorkOrder_TimeTracking_Relationship);
                    }
                    catch (FaultException<OrganizationServiceFault> ex)
                    {
                        var msg = ex.Detail?.Message ?? string.Empty;
                        if (msg.IndexOf("does not exist", StringComparison.OrdinalIgnoreCase) >= 0 ||
                            msg.IndexOf("not associated", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            trace.Trace("No existing association WorkOrder({0}) !- TimeTracking({1}); ignoring.", workOrderRef.Id, timeTrackingRef.Id);
                        }
                        else
                        {
                            trace.Trace("Disassociate fault for WorkOrder({0}) and TimeTracking({1}): {2}", workOrderRef.Id, timeTrackingRef.Id, ex);
                            throw;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                localContext.TraceWithContext("Exception: {0}", ex.Message);
                throw new InvalidPluginExecutionException("PostOperation_WorkOrderTimeTrackingSync_Disassociate failed.", ex);
            }
        }
    }
}