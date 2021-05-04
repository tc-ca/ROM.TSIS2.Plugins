using System;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Messages;
using System.ServiceModel;
using DG.Tools.XrmMockup;
using Xunit;
using Xunit.Sdk;

namespace ROMTS_GSRST.Plugins.Tests
{
    public class PreOperationmsdyn_workorderservicetaskCreateTests : UnitTestBase
    {

        Guid _regulatedEntityId;
        Guid _workOrderId;

        public PreOperationmsdyn_workorderservicetaskCreateTests(XrmMockupFixture fixture) : base(fixture) 
        {
            _regulatedEntityId = orgAdminUIService.Create(new Account() { Name = "Test Regulated Entity" });
            _workOrderId = orgAdminUIService.Create(new msdyn_workorder()
            {
                msdyn_name = "300-345678",
                msdyn_ServiceRequest = null,
                ovs_regulatedentity = new EntityReference(Account.EntityLogicalName, _regulatedEntityId)

            });
        }

        [Fact]
        public void When_parent_workorder_has_name_expect_child_service_task_with_prefix_equal_to_200_parent_name()
        {
            // ACT
            var workOrderServiceTaskId = orgAdminUIService.Create(new msdyn_workorderservicetask()
            {
                msdyn_name = "SATR Boarding Gate Inspection", // sample name based off service task name, it should change after creation
                msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, _workOrderId), // belongs to a work order
            });

            // ASSERT
            var workOrderServiceTask = orgAdminUIService.Retrieve(msdyn_workorderservicetask.EntityLogicalName, workOrderServiceTaskId, new ColumnSet("msdyn_name")).ToEntity<msdyn_workorderservicetask>();

            // Expect work order service task name to start with parent work order's name but prefixed with 200-
            Assert.StartsWith("200-345678", workOrderServiceTask.msdyn_name);
        }

        [Fact]
        public void When_parent_workorder_has_1_previous_service_tasks_expect_new_service_task_with_suffix_2()
        {
            // ARRANGE
            var existingWorkOrderServiceTaskId = orgAdminUIService.Create(new msdyn_workorderservicetask()
            {
                msdyn_name = "200-34567-1",
                msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, _workOrderId)
            });

            // ACT
            var targetWorkOrderServiceTaskId = orgAdminUIService.Create(new msdyn_workorderservicetask()
            {
                msdyn_name = "SATR Boarding Gate Inspection", // sample name based off service task name, it should change after creation
                msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, _workOrderId), // belongs to a work order
            });

            // ASSERT
            var targetWorkOrderServiceTask = orgAdminUIService.Retrieve(msdyn_workorderservicetask.EntityLogicalName, targetWorkOrderServiceTaskId, new ColumnSet("msdyn_name")).ToEntity<msdyn_workorderservicetask>();

            // Expect work order service task name to start with parent work order's name and suffixed with -2
            Assert.Equal("200-345678-2", targetWorkOrderServiceTask.msdyn_name);

        }
    }
}
