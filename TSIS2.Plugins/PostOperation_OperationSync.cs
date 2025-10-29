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
        "TSIS2.Plugins.PostOperation_UnplannedWO-OperationSync_Associate",
        1,
        IsolationModeEnum.Sandbox,
        Description = "Synchronizes operations from Unplanned Work Order to Work Order on association.")]
    public class PostOperation_OperationSync_Associate : IPlugin
    {
        private const string UnplannedWorkOrderEntity = "ts_unplannedworkorder";
        private const string WorkOrderEntity = "msdyn_workorder";
        private const string OperationEntity = "ovs_operation";

        // Lookup field on unplanned work order that points to the msdyn_workorder
        private const string UnplannedWorkOrder_WorkOrderLookup = "ts_workorder";

        // Relationship schema names
        // Unplanned Work Order <-> Operation M:N relationship schema name
        private const string UnplannedWorkOrder_Operation_Relationship = "ts_ovs_operation_ts_unplannedworkorder_ts_unplannedworkorder";
        // Work Order <-> Operation M:N relationship schema name to mirror onto
        private const string WorkOrder_Operation_Relationship = "ts_msdyn_workorder_ovs_operation_ovs_operati";

        public void Execute(IServiceProvider serviceProvider)
        {
            var context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            var factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            var service = factory.CreateOrganizationService(context.UserId);
            var trace = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            try
            {
                trace.Trace("Operation Associate sync start. CorrelationId={0}, Depth={1}", context.CorrelationId, context.Depth);

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

                if (!string.Equals(relationship.SchemaName, UnplannedWorkOrder_Operation_Relationship, StringComparison.OrdinalIgnoreCase))
                {
                    trace.Trace("Relationship {0} does not match expected {1}. Exiting.",
                        relationship.SchemaName, UnplannedWorkOrder_Operation_Relationship);
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

                // Determine which side is the unplanned work order and which side is/are the operation(s)
                EntityReference unplannedWorkOrderRef = null;
                var operationRefs = new EntityReferenceCollection();

                if (string.Equals(target.LogicalName, UnplannedWorkOrderEntity, StringComparison.OrdinalIgnoreCase))
                {
                    unplannedWorkOrderRef = target;
                    foreach (var re in relatedEntities.Where(e => e.LogicalName == OperationEntity))
                        operationRefs.Add(re);
                }
                else if (string.Equals(target.LogicalName, OperationEntity, StringComparison.OrdinalIgnoreCase))
                {
                    var reUnplannedWorkOrder = relatedEntities.FirstOrDefault(e => e.LogicalName == UnplannedWorkOrderEntity);
                    if (reUnplannedWorkOrder != null)
                    {
                        unplannedWorkOrderRef = reUnplannedWorkOrder;
                        operationRefs.Add(target);
                    }
                }

                if (unplannedWorkOrderRef == null || operationRefs.Count == 0)
                {
                    trace.Trace("Could not resolve unplanned work order and operation(s) from the association. Exiting.");
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

                // Mirror the association(s) on the Work Order <-> Operation relationship
                var relationshipToMirror = new Relationship(WorkOrder_Operation_Relationship);

                foreach (var operationRef in operationRefs)
                {
                    try
                    {
                        var associate = new AssociateRequest
                        {
                            Target = new EntityReference(WorkOrderEntity, workOrderRef.Id),
                            Relationship = relationshipToMirror,
                            RelatedEntities = new EntityReferenceCollection { new EntityReference(OperationEntity, operationRef.Id) }
                        };

                        service.Execute(associate);
                        trace.Trace("Mirrored association WorkOrder({0}) <-> Operation({1}) via relationship {2}.",
                            workOrderRef.Id, operationRef.Id, WorkOrder_Operation_Relationship);
                    }
                    catch (FaultException<OrganizationServiceFault> ex)
                    {
                        var msg = ex.Detail?.Message ?? string.Empty;
                        if (msg.IndexOf("already associated", StringComparison.OrdinalIgnoreCase) >= 0 ||
                            msg.IndexOf("duplicate", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            trace.Trace("Association already exists for WorkOrder({0}) and Operation({1}); ignoring.", workOrderRef.Id, operationRef.Id);
                        }
                        else
                        {
                            trace.Trace("Associate fault for WorkOrder({0}) and Operation({1}): {2}", workOrderRef.Id, operationRef.Id, ex.Message);
                            throw;
                        }
                    }
                }
            }
            catch
            {
                throw;
            }
        }
    }

    [CrmPluginRegistration(
        MessageNameEnum.Disassociate,
        "none",
        StageEnum.PostOperation,
        ExecutionModeEnum.Synchronous,
        "",
        "TSIS2.Plugins.PostOperation_UnplannedWO-OperationSync_Disassociate",
        1,
        IsolationModeEnum.Sandbox,
        Description = "Synchronizes operation removal from Unplanned Work Order to Work Order on disassociation.")]
    public class PostOperation_OperationSync_Disassociate : IPlugin
    {
        private const string UnplannedWorkOrderEntity = "ts_unplannedworkorder";
        private const string WorkOrderEntity = "msdyn_workorder";
        private const string OperationEntity = "ovs_operation";

        // Lookup field on unplanned work order that points to the msdyn_workorder
        private const string UnplannedWorkOrder_WorkOrderLookup = "ts_workorder";

        // Relationship schema names
        private const string UnplannedWorkOrder_Operation_Relationship = "ts_ovs_operation_ts_unplannedworkorder_ts_unplannedworkorder";
        private const string WorkOrder_Operation_Relationship = "ts_msdyn_workorder_ovs_operation_ovs_operati";

        public void Execute(IServiceProvider serviceProvider)
        {
            var context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            var factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            var service = factory.CreateOrganizationService(context.UserId);
            var trace = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            try
            {
                trace.Trace("Operation Disassociate sync start. CorrelationId={0}, Depth={1}", context.CorrelationId, context.Depth);

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

                if (!string.Equals(relationship.SchemaName, UnplannedWorkOrder_Operation_Relationship, StringComparison.OrdinalIgnoreCase))
                {
                    trace.Trace("Relationship {0} does not match expected {1}. Exiting.",
                        relationship.SchemaName, UnplannedWorkOrder_Operation_Relationship);
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

                // Determine which side is the unplanned work order and which side is/are the operation(s)
                EntityReference unplannedWorkOrderRef = null;
                var operationRefs = new EntityReferenceCollection();

                if (string.Equals(target.LogicalName, UnplannedWorkOrderEntity, StringComparison.OrdinalIgnoreCase))
                {
                    unplannedWorkOrderRef = target;
                    foreach (var re in relatedEntities.Where(e => e.LogicalName == OperationEntity))
                        operationRefs.Add(re);
                }
                else if (string.Equals(target.LogicalName, OperationEntity, StringComparison.OrdinalIgnoreCase))
                {
                    var reUnplannedWorkOrder = relatedEntities.FirstOrDefault(e => e.LogicalName == UnplannedWorkOrderEntity);
                    if (reUnplannedWorkOrder != null)
                    {
                        unplannedWorkOrderRef = reUnplannedWorkOrder;
                        operationRefs.Add(target);
                    }
                }

                if (unplannedWorkOrderRef == null || operationRefs.Count == 0)
                {
                    trace.Trace("Could not resolve unplanned work order and operation(s) from the disassociation. Exiting.");
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

                // Mirror the disassociation(s) on the Work Order <-> Operation relationship
                var relationshipToMirror = new Relationship(WorkOrder_Operation_Relationship);

                foreach (var operationRef in operationRefs)
                {
                    try
                    {
                        var disassociate = new DisassociateRequest
                        {
                            Target = new EntityReference(WorkOrderEntity, workOrderRef.Id),
                            Relationship = relationshipToMirror,
                            RelatedEntities = new EntityReferenceCollection { new EntityReference(OperationEntity, operationRef.Id) }
                        };

                        service.Execute(disassociate);
                        trace.Trace("Mirrored disassociation WorkOrder({0}) !- Operation({1}) via relationship {2}.",
                            workOrderRef.Id, operationRef.Id, WorkOrder_Operation_Relationship);
                    }
                    catch (FaultException<OrganizationServiceFault> ex)
                    {
                        var msg = ex.Detail?.Message ?? string.Empty;
                        if (msg.IndexOf("does not exist", StringComparison.OrdinalIgnoreCase) >= 0 ||
                            msg.IndexOf("not associated", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            trace.Trace("No existing association WorkOrder({0}) !- Operation({1}); ignoring.", workOrderRef.Id, operationRef.Id);
                        }
                        else
                        {
                            trace.Trace("Disassociate fault for WorkOrder({0}) and Operation({1}): {2}", workOrderRef.Id, operationRef.Id, ex.Message);
                            throw;
                        }
                    }
                }
            }
            catch
            {
                throw;
            }
        }
    }
}