using System;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace TSIS2.Plugins
{
    [CrmPluginRegistration(
        "ts_CreateQualityControlServiceTask",   // Custom API (unique name)
        "msdyn_workorder",                      // Bound to Work Order (extra check)
        StageEnum.PostOperation,
        ExecutionModeEnum.Synchronous,
        "",
        "TSIS2.Plugins.PostOperation_workorderservicetaskQCCreation",
        1,
        IsolationModeEnum.Sandbox,
        Description = "Creates a QC Work Order Service Task from the Work Order using the ROM Service Account.")]
    public class PostOperation_workorderservicetaskQCCreation : IPlugin
    {
        // Task Type IDs (Aviation vs Non-Aviation)
        private static readonly Guid AviationTaskTypeId = new Guid("931b334c-c55b-ee11-8df0-000d3af4f52a");
        private static readonly Guid NonAviationTaskTypeId = new Guid("765fcc32-7339-ef11-a316-6045bd5f6387");

        public void Execute(IServiceProvider serviceProvider)
        {
            var context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            var serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            var service = serviceFactory.CreateOrganizationService(context.UserId);
            var tracing = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            try
            {
                // Retrieve entity reference (also check bound entity)
                if (!context.InputParameters.TryGetValue("Target", out var boundObject))
                    throw new InvalidPluginExecutionException("Bound Work Order entity is required.");

                Guid workOrderId;

                if (boundObject is Entity entity)
                {
                    workOrderId = entity.Id;
                    tracing.Trace("Bound object is full Entity. WorkOrderId: {0}", workOrderId);
                }
                else if (boundObject is EntityReference entityRef)
                {
                    workOrderId = entityRef.Id;
                    tracing.Trace("Bound object is EntityReference. WorkOrderId: {0}", workOrderId);
                }
                else
                {
                    throw new InvalidPluginExecutionException("Unexpected type for bound entity.");
                }

                // Decide Aviation by OperationType.owningbusinessunit.name
                var isAviation = IsAviationByOperationTypeBU(service, workOrderId, tracing);
                var taskTypeId = isAviation ? AviationTaskTypeId : NonAviationTaskTypeId;

                // Create Work Order Service Task
                var wost = new Entity("msdyn_workorderservicetask");
                wost["msdyn_workorder"] = new EntityReference("msdyn_workorder", workOrderId);
                wost["msdyn_tasktype"] = new EntityReference("msdyn_servicetasktype", taskTypeId);

                var createdId = service.Create(wost);
                tracing.Trace("Created WOST: {0}", createdId);

            }
            catch (Exception ex)
            {
                tracing.Trace("CreateQualityControlServiceTask error: {0}", ex);
                throw;
            }
        }

        // WorkOrder -> Operation (ovs_operationid) -> OperationType (ovs_operationtypeid) -> owningbusinessunit.name
        private static bool IsAviationByOperationTypeBU(IOrganizationService service, Guid workOrderId, ITracingService tracing)
        {
            var wo = service.Retrieve("msdyn_workorder", workOrderId, new ColumnSet("ovs_operationid"));
            var opRef = wo.GetAttributeValue<EntityReference>("ovs_operationid");
            if (opRef == null)
            {
                tracing.Trace("Work Order has no ovs_operationid.");
                return false;
            }

            var op = service.Retrieve("ovs_operation", opRef.Id, new ColumnSet("ovs_operationtypeid"));
            var opTypeRef = op.GetAttributeValue<EntityReference>("ovs_operationtypeid");
            if (opTypeRef == null)
            {
                tracing.Trace("Operation has no ovs_operationtypeid.");
                return false;
            }

            var opType = service.Retrieve("ovs_operationtype", opTypeRef.Id, new ColumnSet("owningbusinessunit"));
            var buRef = opType.GetAttributeValue<EntityReference>("owningbusinessunit");

            if (buRef == null)
            {
                tracing.Trace("OperationType has no owningbusinessunit.");
                return false;
            }

            tracing.Trace("OperationType BU ID: {0}", buRef.Id);
            return EnvironmentVariableHelper.IsAvSecBU(service, buRef.Id, tracing);
        }
    }
}