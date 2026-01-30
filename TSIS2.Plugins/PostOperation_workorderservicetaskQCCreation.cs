using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;

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
    public class PostOperation_workorderservicetaskQCCreation : PluginBase
    {
        // Task Type IDs (Aviation vs Non-Aviation)
        private static readonly Guid AviationTaskTypeId = new Guid("931b334c-c55b-ee11-8df0-000d3af4f52a");
        private static readonly Guid NonAviationTaskTypeId = new Guid("765fcc32-7339-ef11-a316-6045bd5f6387");

        private const string QC_WOST_NAME = "Quality Control(QC) Review";

        public PostOperation_workorderservicetaskQCCreation(string unsecure, string secure)
            : base(typeof(PostOperation_workorderservicetaskQCCreation))
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
            var tracingService = localContext.TracingService;

            try
            {
                // Retrieve entity reference (also check bound entity)
                if (!context.InputParameters.TryGetValue("Target", out var boundObject))
                    throw new InvalidPluginExecutionException("Bound Work Order entity is required.");

                Guid workOrderId;

                if (boundObject is Entity entity)
                {
                    workOrderId = entity.Id;
                    tracingService.Trace("Bound object is full Entity. WorkOrderId: {0}", workOrderId);
                }
                else if (boundObject is EntityReference entityRef)
                {
                    workOrderId = entityRef.Id;
                    tracingService.Trace("Bound object is EntityReference. WorkOrderId: {0}", workOrderId);
                }
                else
                {
                    throw new InvalidPluginExecutionException("Unexpected type for bound entity.");
                }

                // Decide Aviation by OperationType.owningbusinessunit.name
                var isAviation = IsAviationByOperationTypeBU(service, workOrderId, tracingService);
                var taskTypeId = isAviation ? AviationTaskTypeId : NonAviationTaskTypeId;

                //User who created the QC Task (not Service Account)
                var currentUser = context.InitiatingUserId;

                // Create Work Order Service Task
                var wost = new Entity("msdyn_workorderservicetask");
                wost["msdyn_workorder"] = new EntityReference("msdyn_workorder", workOrderId);
                wost["msdyn_tasktype"] = new EntityReference("msdyn_servicetasktype", taskTypeId);
                wost["ownerid"] = new EntityReference("systemuser", currentUser);
                wost["msdyn_name"] = QC_WOST_NAME;

                var createdId = service.Create(wost);
                tracingService.Trace("Created WOST: {0}", createdId);
            }
            catch (Exception ex)
            {
                localContext.TraceWithContext("CreateQualityControlServiceTask error: {0}", ex.Message);
                throw new InvalidPluginExecutionException("PostOperation_workorderservicetaskQCCreation failed.", ex);
            }
        }

        // WorkOrder -> Operation (ovs_operationid) -> OperationType (ovs_operationtypeid) -> owningbusinessunit.name
        private static bool IsAviationByOperationTypeBU(IOrganizationService service, Guid workOrderId, ITracingService tracingService)
        {
            var wo = service.Retrieve("msdyn_workorder", workOrderId, new ColumnSet("ovs_operationid"));
            var opRef = wo.GetAttributeValue<EntityReference>("ovs_operationid");
            if (opRef == null)
            {
                tracingService.Trace("Work Order has no ovs_operationid.");
                return false;
            }

            var op = service.Retrieve("ovs_operation", opRef.Id, new ColumnSet("ovs_operationtypeid"));
            var opTypeRef = op.GetAttributeValue<EntityReference>("ovs_operationtypeid");
            if (opTypeRef == null)
            {
                tracingService.Trace("Operation has no ovs_operationtypeid.");
                return false;
            }

            var opType = service.Retrieve("ovs_operationtype", opTypeRef.Id, new ColumnSet("owningbusinessunit"));
            var buRef = opType.GetAttributeValue<EntityReference>("owningbusinessunit");

            if (buRef == null)
            {
                tracingService.Trace("OperationType has no owningbusinessunit.");
                return false;
            }

            tracingService.Trace("OperationType BU ID: {0}", buRef.Id);
            return OrganizationConfig.IsAvSecBU(service, buRef.Id, tracingService);
        }
    }
}