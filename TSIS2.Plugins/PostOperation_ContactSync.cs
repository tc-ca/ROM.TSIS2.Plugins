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
        "TSIS2.Plugins.PostOperation_UnplannedWO-ContactSync_Associate",
        1,
        IsolationModeEnum.Sandbox,
        Description = "Synchronizes contacts from Unplanned Work Order to Work Order on association.")]
    public class PostOperation_ContactSync_Associate : PluginBase
    {
        private const string UnplannedWorkOrderEntity = "ts_unplannedworkorder";
        private const string WorkOrderEntity = "msdyn_workorder";
        private const string ContactEntity = "contact";

        // Lookup field on unplanned work order that points to the msdyn_workorder
        private const string UnplannedWorkOrder_WorkOrderLookup = "ts_workorder";

        // Relationship schema names
        // Unplanned Work Order <-> Contact M:N relationship schema name
        private const string UnplannedWorkOrder_Contact_Relationship = "ts_Contact_ts_unplannedworkorder_ts_unplannedworkorder";
        // Work Order <-> Contact M:N relationship schema name to mirror onto
        private const string WorkOrder_Contact_Relationship = "ts_Contact_msdyn_workorder_msdyn_workorder";

        public PostOperation_ContactSync_Associate(string unsecure, string secure)
            : base(typeof(PostOperation_ContactSync_Associate))
        {
        }

        protected override void ExecuteCrmPlugin(LocalPluginContext localContext)
        {
            if (localContext == null)
            {
                throw new InvalidPluginExecutionException("localContext");
            }

            var context = localContext.PluginExecutionContext;
            var service = localContext.OrganizationService;

            try
            {
                localContext.Trace("Contact Associate sync start. CorrelationId={0}, Depth={1}", context.CorrelationId, context.Depth);

                // Guard: avoid recursion
                if (context.Depth > 1)
                {
                    localContext.Trace("Depth > 1 detected. Exiting to prevent recursion.");
                    return;
                }

                if (!(context.InputParameters.Contains("Relationship") &&
                      context.InputParameters["Relationship"] is Relationship relationship))
                {
                    localContext.Trace("Missing Relationship parameter. Exiting.");
                    return;
                }

                if (!string.Equals(relationship.SchemaName, UnplannedWorkOrder_Contact_Relationship, StringComparison.OrdinalIgnoreCase))
                {
                    localContext.Trace("Relationship {0} does not match expected {1}. Exiting.",
                        relationship.SchemaName, UnplannedWorkOrder_Contact_Relationship);
                    return;
                }

                if (!(context.InputParameters.Contains("Target") &&
                      context.InputParameters["Target"] is EntityReference target))
                {
                    localContext.Trace("Missing Target parameter. Exiting.");
                    return;
                }

                if (!(context.InputParameters.Contains("RelatedEntities") &&
                      context.InputParameters["RelatedEntities"] is EntityReferenceCollection relatedEntities) ||
                    relatedEntities.Count == 0)
                {
                    localContext.Trace("Missing RelatedEntities parameter. Exiting.");
                    return;
                }

                // Determine which side is the unplanned work order and which side is/are the contact(s)
                EntityReference unplannedWorkOrderRef = null;
                var contactRefs = new EntityReferenceCollection();

                if (string.Equals(target.LogicalName, UnplannedWorkOrderEntity, StringComparison.OrdinalIgnoreCase))
                {
                    unplannedWorkOrderRef = target;
                    foreach (var re in relatedEntities.Where(e => e.LogicalName == ContactEntity))
                        contactRefs.Add(re);
                }
                else if (string.Equals(target.LogicalName, ContactEntity, StringComparison.OrdinalIgnoreCase))
                {
                    var reUnplannedWorkOrder = relatedEntities.FirstOrDefault(e => e.LogicalName == UnplannedWorkOrderEntity);
                    if (reUnplannedWorkOrder != null)
                    {
                        unplannedWorkOrderRef = reUnplannedWorkOrder;
                        contactRefs.Add(target);
                    }
                }

                if (unplannedWorkOrderRef == null || contactRefs.Count == 0)
                {
                    localContext.Trace("Could not resolve unplanned work order and contact(s) from the association. Exiting.");
                    return;
                }

                // Retrieve the linked msdyn_workorder from the unplanned work order
                var unplannedWorkOrder = service.Retrieve(UnplannedWorkOrderEntity, unplannedWorkOrderRef.Id, new ColumnSet(UnplannedWorkOrder_WorkOrderLookup));
                var workOrderRef = unplannedWorkOrder.GetAttributeValue<EntityReference>(UnplannedWorkOrder_WorkOrderLookup);
                if (workOrderRef == null)
                {
                    localContext.Trace("Unplanned Work Order {0} does not have {1} populated. Exiting.", unplannedWorkOrderRef.Id, UnplannedWorkOrder_WorkOrderLookup);
                    return;
                }

                // Mirror the association(s) on the Work Order <-> Contact relationship
                var relationshipToMirror = new Relationship(WorkOrder_Contact_Relationship);

                foreach (var contactRef in contactRefs)
                {
                    try
                    {
                        var associate = new AssociateRequest
                        {
                            Target = new EntityReference(WorkOrderEntity, workOrderRef.Id),
                            Relationship = relationshipToMirror,
                            RelatedEntities = new EntityReferenceCollection { new EntityReference(ContactEntity, contactRef.Id) }
                        };

                        service.Execute(associate);
                        localContext.Trace("Mirrored association WorkOrder({0}) <-> Contact({1}) via relationship {2}.",
                            workOrderRef.Id, contactRef.Id, WorkOrder_Contact_Relationship);
                    }
                    catch (FaultException<OrganizationServiceFault> ex)
                    {
                        var msg = ex.Detail?.Message ?? string.Empty;
                        if (msg.IndexOf("already associated", StringComparison.OrdinalIgnoreCase) >= 0 ||
                            msg.IndexOf("duplicate", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            localContext.Trace("Association already exists for WorkOrder({0}) and Contact({1}); ignoring.", workOrderRef.Id, contactRef.Id);
                        }
                        else
                        {
                            localContext.Trace("Associate fault for WorkOrder({0}) and Contact({1}): {2}", workOrderRef.Id, contactRef.Id, ex);
                            throw;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                localContext.TraceWithContext("Exception: {0}", ex.Message);
                throw new InvalidPluginExecutionException("PostOperation_ContactSync_Associate failed.", ex);
            }
        }
    }

    [CrmPluginRegistration(
        MessageNameEnum.Disassociate,
        "none",
        StageEnum.PostOperation,
        ExecutionModeEnum.Synchronous,
        "",
        "TSIS2.Plugins.PostOperation_UnplannedWO-ContactSync_Disassociate",
        1,
        IsolationModeEnum.Sandbox,
        Description = "Synchronizes contact removal from Unplanned Work Order to Work Order on disassociation.")]
    public class PostOperation_ContactSync_Disassociate : PluginBase
    {
        private const string UnplannedWorkOrderEntity = "ts_unplannedworkorder";
        private const string WorkOrderEntity = "msdyn_workorder";
        private const string ContactEntity = "contact";

        // Lookup field on unplanned work order that points to the msdyn_workorder
        private const string UnplannedWorkOrder_WorkOrderLookup = "ts_workorder";

        // Relationship schema names
        private const string UnplannedWorkOrder_Contact_Relationship = "ts_Contact_ts_unplannedworkorder_ts_unplannedworkorder";
        private const string WorkOrder_Contact_Relationship = "ts_Contact_msdyn_workorder_msdyn_workorder";

        public PostOperation_ContactSync_Disassociate(string unsecure, string secure)
            : base(typeof(PostOperation_ContactSync_Disassociate))
        {
        }

        protected override void ExecuteCrmPlugin(LocalPluginContext localContext)
        {
            if (localContext == null)
            {
                throw new InvalidPluginExecutionException("localContext");
            }

            var context = localContext.PluginExecutionContext;
            var service = localContext.OrganizationService;

            try
            {
                localContext.Trace("Contact Disassociate sync start. CorrelationId={0}, Depth={1}", context.CorrelationId, context.Depth);

                // Guard: avoid recursion
                if (context.Depth > 1)
                {
                    localContext.Trace("Depth > 1 detected. Exiting to prevent recursion.");
                    return;
                }

                if (!(context.InputParameters.Contains("Relationship") &&
                      context.InputParameters["Relationship"] is Relationship relationship))
                {
                    localContext.Trace("Missing Relationship parameter. Exiting.");
                    return;
                }

                if (!string.Equals(relationship.SchemaName, UnplannedWorkOrder_Contact_Relationship, StringComparison.OrdinalIgnoreCase))
                {
                    localContext.Trace("Relationship {0} does not match expected {1}. Exiting.",
                        relationship.SchemaName, UnplannedWorkOrder_Contact_Relationship);
                    return;
                }

                if (!(context.InputParameters.Contains("Target") &&
                      context.InputParameters["Target"] is EntityReference target))
                {
                    localContext.Trace("Missing Target parameter. Exiting.");
                    return;
                }

                if (!(context.InputParameters.Contains("RelatedEntities") &&
                      context.InputParameters["RelatedEntities"] is EntityReferenceCollection relatedEntities) ||
                    relatedEntities.Count == 0)
                {
                    localContext.Trace("Missing RelatedEntities parameter. Exiting.");
                    return;
                }

                // Determine which side is the unplanned work order and which side is/are the contact(s)
                EntityReference unplannedWorkOrderRef = null;
                var contactRefs = new EntityReferenceCollection();

                if (string.Equals(target.LogicalName, UnplannedWorkOrderEntity, StringComparison.OrdinalIgnoreCase))
                {
                    unplannedWorkOrderRef = target;
                    foreach (var re in relatedEntities.Where(e => e.LogicalName == ContactEntity))
                        contactRefs.Add(re);
                }
                else if (string.Equals(target.LogicalName, ContactEntity, StringComparison.OrdinalIgnoreCase))
                {
                    var reUnplannedWorkOrder = relatedEntities.FirstOrDefault(e => e.LogicalName == UnplannedWorkOrderEntity);
                    if (reUnplannedWorkOrder != null)
                    {
                        unplannedWorkOrderRef = reUnplannedWorkOrder;
                        contactRefs.Add(target);
                    }
                }

                if (unplannedWorkOrderRef == null || contactRefs.Count == 0)
                {
                    localContext.Trace("Could not resolve unplanned work order and contact(s) from the disassociation. Exiting.");
                    return;
                }

                // Retrieve the linked msdyn_workorder from the unplanned work order
                var unplannedWorkOrder = service.Retrieve(UnplannedWorkOrderEntity, unplannedWorkOrderRef.Id, new ColumnSet(UnplannedWorkOrder_WorkOrderLookup));
                var workOrderRef = unplannedWorkOrder.GetAttributeValue<EntityReference>(UnplannedWorkOrder_WorkOrderLookup);
                if (workOrderRef == null)
                {
                    localContext.Trace("Unplanned Work Order {0} does not have {1} populated. Exiting.", unplannedWorkOrderRef.Id, UnplannedWorkOrder_WorkOrderLookup);
                    return;
                }

                // Mirror the disassociation(s) on the Work Order <-> Contact relationship
                var relationshipToMirror = new Relationship(WorkOrder_Contact_Relationship);

                foreach (var contactRef in contactRefs)
                {
                    try
                    {
                        var disassociate = new DisassociateRequest
                        {
                            Target = new EntityReference(WorkOrderEntity, workOrderRef.Id),
                            Relationship = relationshipToMirror,
                            RelatedEntities = new EntityReferenceCollection { new EntityReference(ContactEntity, contactRef.Id) }
                        };

                        service.Execute(disassociate);
                        localContext.Trace("Mirrored disassociation WorkOrder({0}) !- Contact({1}) via relationship {2}.",
                            workOrderRef.Id, contactRef.Id, WorkOrder_Contact_Relationship);
                    }
                    catch (FaultException<OrganizationServiceFault> ex)
                    {
                        var msg = ex.Detail?.Message ?? string.Empty;
                        if (msg.IndexOf("does not exist", StringComparison.OrdinalIgnoreCase) >= 0 ||
                            msg.IndexOf("not associated", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            localContext.Trace("No existing association WorkOrder({0}) !- Contact({1}); ignoring.", workOrderRef.Id, contactRef.Id);
                        }
                        else
                        {
                            localContext.Trace("Disassociate fault for WorkOrder({0}) and Contact({1}): {2}", workOrderRef.Id, contactRef.Id, ex);
                            throw;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                localContext.TraceWithContext("Exception: {0}", ex.Message);
                throw new InvalidPluginExecutionException("PostOperation_ContactSync_Disassociate failed.", ex);
            }
        }
    }
}