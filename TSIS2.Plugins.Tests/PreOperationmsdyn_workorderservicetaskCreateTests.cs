using System;
using System.Collections.Generic;
using System.Reflection;
using System.Linq;
using Xunit;
using FakeItEasy;
using FakeXrmEasy;
using Microsoft.Xrm.Sdk;
using TSIS2.Common;

namespace TSIS2.Plugins.Tests
{
 
    public class PreOperationmsdyn_workorderservicetaskCreateTests
    {
        [Fact]
        public void When_parent_workorder_has_name_expect_child_service_task_with_prefix_equal_to_200_parent_name()
        {
            /**********
            * ARRANGE
            **********/
            var context = new XrmFakedContext();

            // Setup prefixes and IDs
            var workOrderPrefix = "300";
            var workOrderServiceTaskPrefix = "200";
            var uniqueId = "34567";

            // Given a work order service task that
            var regulatedEntityId = Guid.NewGuid();
            var regulatedEntity = new Account()
            {
                Id = regulatedEntityId,
                Name = "Test Regulated Entity"
            };

            var workOrderId = Guid.NewGuid();
            var workOrderName = string.Format("{0}-{1}", workOrderPrefix, uniqueId);
            var workOrder = new msdyn_workorder()
            {
                Id = workOrderId,
                msdyn_name = workOrderName,
                msdyn_ServiceRequest = null, // does not already belong to a case (Incident)
                ovs_regulatedentity = new EntityReference(Account.EntityLogicalName, regulatedEntityId)
            };

            var workOrderServiceTaskId = Guid.NewGuid();
            var workOrderServiceTask = new msdyn_workorderservicetask()
            {
                Id = workOrderServiceTaskId,
                msdyn_name = "SATR Boarding Gate Inspection", // sample name based off service task name, it should change after creation
                msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId), // belongs to a work order
            };

            // Fake the relationship between work order and work order service task
            context.AddRelationship("msdyn_msdyn_workorder_msdyn_workorderservicetask_WorkOrder", new XrmFakedRelationship()
            {
                Entity1LogicalName = msdyn_workorderservicetask.EntityLogicalName,
                Entity1Attribute = "msdyn_workorder",
                Entity2LogicalName = msdyn_workorder.EntityLogicalName,
                Entity2Attribute = "msdyn_workorderid",
                RelationshipType = XrmFakedRelationship.enmFakeRelationshipType.OneToMany
            });

            context.Initialize(
                new List<Entity>() {
                    regulatedEntity,
                    workOrder
                }
            );

            ParameterCollection inputParams = new ParameterCollection
            {
                { "Target", workOrderServiceTask }
            };
            ParameterCollection outputParams = new ParameterCollection
            {
                { "id", workOrderServiceTaskId }
            };

            /**********
            * ACT
            **********/
            // Execute the PreOperationmsdyn_workorderservicetaskUpdate plugin with the defined workOrderServiceTask as a target
            context.ExecutePluginWith<PreOperationmsdyn_workorderservicetaskCreate>(inputParams, outputParams, null, null);

            /**********
             * ASSERT
             **********/
            // Expect work order service task name to start with parent work order's name but prefixed with 200-
            Assert.True(workOrderServiceTask.msdyn_name.StartsWith(string.Format("{0}-{1}-", workOrderServiceTaskPrefix, uniqueId)));
        }

        [Fact]
        public void When_parent_workorder_has_1_previous_service_tasks_expect_new_service_task_with_suffix_2()
        {
            /**********
            * ARRANGE
            **********/
            var context = new XrmFakedContext();

            // Setup prefixes and IDs
            var workOrderPrefix = "300";
            var workOrderServiceTaskPrefix = "200";
            var uniqueId = "34567";

            // Given a work order service task that
            var regulatedEntityId = Guid.NewGuid();
            var regulatedEntity = new Account()
            {
                Id = regulatedEntityId,
                Name = "Test Regulated Entity"
            };

            var workOrderId = Guid.NewGuid();
            var workOrderName = "300-34567";

            var workOrder = new msdyn_workorder()
            {
                Id = workOrderId,
                msdyn_name = workOrderName,
                msdyn_ServiceRequest = null, // does not already belong to a case (Incident)
                ovs_regulatedentity = new EntityReference(Account.EntityLogicalName, regulatedEntityId)
            };

            var existingWorkOrderServiceTask = new msdyn_workorderservicetask()
            {
                Id = Guid.NewGuid(),
                msdyn_name = workOrderName + "-1",
                msdyn_WorkOrder = workOrder.ToEntityReference()
            };

            var targetWorkOrderServiceTaskId = Guid.NewGuid();
            var targetWorkOrderServiceTask = new msdyn_workorderservicetask()
            {
                Id = targetWorkOrderServiceTaskId,
                msdyn_name = "SATR Boarding Gate Inspection", // sample name based off service task name, it should change after creation
                msdyn_WorkOrder = workOrder.ToEntityReference()
            };

            // Fake the relationship between work order and work order service task
            context.AddRelationship("msdyn_msdyn_workorder_msdyn_workorderservicetask_WorkOrder", new XrmFakedRelationship()
            {
                Entity1LogicalName = msdyn_workorderservicetask.EntityLogicalName,
                Entity1Attribute = "msdyn_workorder",
                Entity2LogicalName = msdyn_workorder.EntityLogicalName,
                Entity2Attribute = "msdyn_workorderid",
                RelationshipType = XrmFakedRelationship.enmFakeRelationshipType.OneToMany
            });

            context.Initialize(
                new List<Entity>() {
                    regulatedEntity,
                    workOrder,
                    existingWorkOrderServiceTask,
                }
            );

            ParameterCollection inputParams = new ParameterCollection
            {
                { "Target", targetWorkOrderServiceTask }
            };
            ParameterCollection outputParams = new ParameterCollection
            {
                { "id", targetWorkOrderServiceTaskId }
            };

            /**********
            * ACT
            **********/
            // Execute the PreOperationmsdyn_workorderservicetaskUpdate plugin with the defined workOrderServiceTask as a target
            context.ExecutePluginWith<PreOperationmsdyn_workorderservicetaskCreate>(inputParams, outputParams, null, null);

            /**********
             * ASSERT
             **********/
            // Expect work order service task name to start with parent work order's name and suffixed with -2
            Assert.Equal(string.Format("{0}-{1}-2", workOrderServiceTaskPrefix, uniqueId), targetWorkOrderServiceTask.msdyn_name);
        }

    }
}
