using System;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Xunit;

namespace ROMTS_GSRST.Plugins.Tests
{
    public class PreOperationmsdyn_workorderUpdateTests : UnitTestBase
    {

        public PreOperationmsdyn_workorderUpdateTests(XrmMockupFixture fixture) : base(fixture) { }

        [Fact]
        public void When_all_work_order_service_tasks_completed_and_work_order_posted_closed_expect_work_order_to_stay_posted_closed()
        {
            // ARRANGE
            var serviceAccountId = orgAdminUIService.Create(new Account { Name = "Test Service Account" });
            var incidentId = orgAdminUIService.Create(new Incident { });
            var workOrderId = orgAdminUIService.Create(new msdyn_workorder
            {
                msdyn_name = "300-345678",
                msdyn_SystemStatus = msdyn_wosystemstatus.OpenCompleted,
                msdyn_ServiceRequest = new EntityReference(Incident.EntityLogicalName, incidentId),
                msdyn_ServiceAccount = new EntityReference(Account.EntityLogicalName, serviceAccountId),
            });
            var existingWorkOrderServiceTask1Id = orgAdminUIService.Create(new msdyn_workorderservicetask
            {
                msdyn_name = "200-345678-1",
                msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId), // belongs to a work order
                msdyn_PercentComplete = 100.00,
                ovs_QuestionnaireResponse = "",
                statuscode = msdyn_workorderservicetask_statuscode.Complete
            });

            // ACT
            orgAdminUIService.Update(new msdyn_workorder { Id = workOrderId, msdyn_SystemStatus = msdyn_wosystemstatus.ClosedPosted });

            // ASSERT
            var workOrder = orgAdminUIService.Retrieve(msdyn_workorder.EntityLogicalName, workOrderId, new ColumnSet("msdyn_systemstatus")).ToEntity<msdyn_workorder>();
            Assert.Equal(msdyn_wosystemstatus.ClosedPosted, workOrder.msdyn_SystemStatus);
        }
        [Fact]
        public void When_case_of_work_order_updated_to_new_case_existing_findings_of_old_case_associate_to_new_case()
        {
            //ARRANGE
            var case1Id = orgAdminUIService.Create(new Incident { });
            var case2Id = orgAdminUIService.Create(new Incident { });

            var workOrderId = orgAdminService.Create(new msdyn_workorder { msdyn_ServiceRequest = new EntityReference(Incident.EntityLogicalName, case1Id) });

            var finding1 = orgAdminService.Create(new ovs_Finding { ovs_CaseId = new EntityReference(Incident.EntityLogicalName, case1Id), ts_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId) });
            var finding2 = orgAdminService.Create(new ovs_Finding { ovs_CaseId = new EntityReference(Incident.EntityLogicalName, case1Id), ts_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId) });
            var finding3 = orgAdminService.Create(new ovs_Finding { ovs_CaseId = new EntityReference(Incident.EntityLogicalName, case1Id), ts_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId) });

            

            //ACT
            orgAdminService.Update(new msdyn_workorder { Id = workOrderId, msdyn_ServiceRequest = new EntityReference(Incident.EntityLogicalName, case2Id) });

            //ASSERT
            var query = new QueryExpression(ovs_Finding.EntityLogicalName)
            {
                ColumnSet = new ColumnSet("ovs_caseid")
            };
            var findings = orgAdminUIService.RetrieveMultiple(query).Entities.Cast<ovs_Finding>().ToList();

            //Expect all findings to be associated to Case2
            Assert.Equal(case2Id, findings[0].ovs_CaseId.Id);
            Assert.Equal(case2Id, findings[1].ovs_CaseId.Id);
            Assert.Equal(case2Id, findings[2].ovs_CaseId.Id);
        }

    }
}
