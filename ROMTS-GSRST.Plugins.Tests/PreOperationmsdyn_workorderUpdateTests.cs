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
        public void When_case_of_work_order_updated_to_new_case_expect_existing_findings_of_old_case_to_be_associated_to_new_case()
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

        [Fact]
        public void When_activity_type_of_work_order_updated_to_new_activity_type_expect_new_work_order_service_task_to_be_created_with_new_task_type()
        {
            //ARRANGE
            var serviceTaskType1 = orgAdminService.Create(new msdyn_servicetasktype { });
            var serviceTaskType1Reference = new EntityReference(msdyn_servicetasktype.EntityLogicalName, serviceTaskType1);

            var incidentType1 = orgAdminService.Create(new msdyn_incidenttype { });
            var incidentType1Reference = new EntityReference(msdyn_incidenttype.EntityLogicalName, incidentType1);

            var incidentTypeServiceTask1 = orgAdminService.Create(new msdyn_incidenttypeservicetask {
                msdyn_TaskType = serviceTaskType1Reference,
                msdyn_IncidentType = incidentType1Reference
            });

            var serviceTaskType2 = orgAdminService.Create(new msdyn_servicetasktype { });
            var serviceTaskType2Reference = new EntityReference(msdyn_servicetasktype.EntityLogicalName, serviceTaskType2);

            var incidentType2 = orgAdminService.Create(new msdyn_incidenttype { });
            var incidentType2Reference = new EntityReference(msdyn_incidenttype.EntityLogicalName, incidentType2);

            var incidentTypeServiceTask2 = orgAdminService.Create(new msdyn_incidenttypeservicetask
            {
                msdyn_TaskType = serviceTaskType2Reference,
                msdyn_IncidentType = incidentType2Reference
            });

            var workOrderId = orgAdminService.Create(new msdyn_workorder {  
                msdyn_PrimaryIncidentType = incidentType1Reference
            });

            var wost = orgAdminService.Create(new msdyn_workorderservicetask {
                msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId),
                msdyn_TaskType = serviceTaskType1Reference,
                statuscode = msdyn_workorderservicetask_statuscode.New
            });

            //ACT
            orgAdminService.Update(new msdyn_workorder { Id = workOrderId, msdyn_PrimaryIncidentType = incidentType2Reference });

            //ASSERT
            var query = new QueryExpression(msdyn_workorderservicetask.EntityLogicalName)
            {
                ColumnSet = new ColumnSet("msdyn_tasktype")
            };
            var wosts = orgAdminUIService.RetrieveMultiple(query).Entities.Cast<msdyn_workorderservicetask>().ToList();

            //Expect the Task Type of the new Work Order Service Task to be the second Service Task Type
            Assert.Equal(serviceTaskType2Reference.Id, wosts[1].msdyn_TaskType.Id);
        }

        [Fact]
        public void When_activity_type_of_work_order_updated_to_new_activity_type_expect_related_work_order_service_tasks_with_status_not_new_to_not_change()
        {
            //ARRANGE
            var serviceTaskType1 = orgAdminService.Create(new msdyn_servicetasktype { });
            var serviceTaskType1Reference = new EntityReference(msdyn_servicetasktype.EntityLogicalName, serviceTaskType1);

            var incidentType1 = orgAdminService.Create(new msdyn_incidenttype { });
            var incidentType1Reference = new EntityReference(msdyn_incidenttype.EntityLogicalName, incidentType1);

            var incidentTypeServiceTask1 = orgAdminService.Create(new msdyn_incidenttypeservicetask
            {
                msdyn_TaskType = serviceTaskType1Reference,
                msdyn_IncidentType = incidentType1Reference
            });

            var serviceTaskType2 = orgAdminService.Create(new msdyn_servicetasktype { });
            var serviceTaskType2Reference = new EntityReference(msdyn_servicetasktype.EntityLogicalName, serviceTaskType2);

            var incidentType2 = orgAdminService.Create(new msdyn_incidenttype { });
            var incidentType2Reference = new EntityReference(msdyn_incidenttype.EntityLogicalName, incidentType2);

            var incidentTypeServiceTask2 = orgAdminService.Create(new msdyn_incidenttypeservicetask
            {
                msdyn_TaskType = serviceTaskType2Reference,
                msdyn_IncidentType = incidentType2Reference
            });

            var workOrderId = orgAdminService.Create(new msdyn_workorder
            {
                msdyn_PrimaryIncidentType = incidentType1Reference
            });

            var wost = orgAdminService.Create(new msdyn_workorderservicetask
            {
                msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId),
                msdyn_TaskType = serviceTaskType1Reference,
                statuscode = msdyn_workorderservicetask_statuscode.InProgress
            });

            //ACT
            orgAdminService.Update(new msdyn_workorder { Id = workOrderId, msdyn_PrimaryIncidentType = incidentType2Reference });

            //ASSERT
            var query = new QueryExpression(msdyn_workorderservicetask.EntityLogicalName)
            {
                ColumnSet = new ColumnSet("msdyn_tasktype")
            };
            var wosts = orgAdminUIService.RetrieveMultiple(query).Entities.Cast<msdyn_workorderservicetask>().ToList();

            //Expect the Task Type of the Work Order Service Task to remain the first Task Type Reference
            Assert.Equal(serviceTaskType1Reference.Id, wosts[0].msdyn_TaskType.Id);
        }

        [Fact]
        public void When_activity_type_of_work_order_updated_to_new_activity_type_expect_only_related_work_order_service_tasks_with_status_new_to_be_replaced()
        {
            //ARRANGE
            var serviceTaskType1 = orgAdminService.Create(new msdyn_servicetasktype { });
            var serviceTaskType1Reference = new EntityReference(msdyn_servicetasktype.EntityLogicalName, serviceTaskType1);

            var incidentType1 = orgAdminService.Create(new msdyn_incidenttype { });
            var incidentType1Reference = new EntityReference(msdyn_incidenttype.EntityLogicalName, incidentType1);

            var incidentTypeServiceTask1 = orgAdminService.Create(new msdyn_incidenttypeservicetask
            {
                msdyn_TaskType = serviceTaskType1Reference,
                msdyn_IncidentType = incidentType1Reference
            });

            var serviceTaskType2 = orgAdminService.Create(new msdyn_servicetasktype { });
            var serviceTaskType2Reference = new EntityReference(msdyn_servicetasktype.EntityLogicalName, serviceTaskType2);

            var incidentType2 = orgAdminService.Create(new msdyn_incidenttype { });
            var incidentType2Reference = new EntityReference(msdyn_incidenttype.EntityLogicalName, incidentType2);

            var incidentTypeServiceTask2 = orgAdminService.Create(new msdyn_incidenttypeservicetask
            {
                msdyn_TaskType = serviceTaskType2Reference,
                msdyn_IncidentType = incidentType2Reference
            });

            var workOrderId = orgAdminService.Create(new msdyn_workorder
            {
                msdyn_PrimaryIncidentType = incidentType1Reference
            });

            var wost1 = orgAdminService.Create(new msdyn_workorderservicetask
            {
                msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId),
                msdyn_TaskType = serviceTaskType1Reference,
                statuscode = msdyn_workorderservicetask_statuscode.New
            });

            var wost2 = orgAdminService.Create(new msdyn_workorderservicetask
            {
                msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId),
                msdyn_TaskType = serviceTaskType1Reference,
                statuscode = msdyn_workorderservicetask_statuscode.InProgress
            });

            //ACT
            orgAdminService.Update(new msdyn_workorder { Id = workOrderId, msdyn_PrimaryIncidentType = incidentType2Reference });

            //ASSERT
            var query = new QueryExpression(msdyn_workorderservicetask.EntityLogicalName)
            {
                ColumnSet = new ColumnSet("msdyn_tasktype")
            };
            var wosts = orgAdminUIService.RetrieveMultiple(query).Entities.Cast<msdyn_workorderservicetask>().ToList();

            //Expect the Task Type of the second Work Order Service Task (In-Progress) to remain the first Task Type reference
            Assert.Equal(serviceTaskType1Reference.Id, wosts[1].msdyn_TaskType.Id);

            //Expect the Task Type of the new Work Order Service Task (New) to be changed
            Assert.Equal(serviceTaskType2Reference.Id, wosts[2].msdyn_TaskType.Id);
        }

        [Fact]
        public void When_activity_type_of_work_order_updated_to_new_activity_type_expect_old_related_work_order_service_tasks_with_status_new_to_be_inactive()
        {
            //ARRANGE
            var serviceTaskType1 = orgAdminService.Create(new msdyn_servicetasktype { });
            var serviceTaskType1Reference = new EntityReference(msdyn_servicetasktype.EntityLogicalName, serviceTaskType1);

            var incidentType1 = orgAdminService.Create(new msdyn_incidenttype { });
            var incidentType1Reference = new EntityReference(msdyn_incidenttype.EntityLogicalName, incidentType1);

            var incidentTypeServiceTask1 = orgAdminService.Create(new msdyn_incidenttypeservicetask
            {
                msdyn_TaskType = serviceTaskType1Reference,
                msdyn_IncidentType = incidentType1Reference
            });

            var serviceTaskType2 = orgAdminService.Create(new msdyn_servicetasktype { });
            var serviceTaskType2Reference = new EntityReference(msdyn_servicetasktype.EntityLogicalName, serviceTaskType2);

            var incidentType2 = orgAdminService.Create(new msdyn_incidenttype { });
            var incidentType2Reference = new EntityReference(msdyn_incidenttype.EntityLogicalName, incidentType2);

            var incidentTypeServiceTask2 = orgAdminService.Create(new msdyn_incidenttypeservicetask
            {
                msdyn_TaskType = serviceTaskType2Reference,
                msdyn_IncidentType = incidentType2Reference
            });

            var workOrderId = orgAdminService.Create(new msdyn_workorder
            {
                msdyn_PrimaryIncidentType = incidentType1Reference
            });

            var wost = orgAdminService.Create(new msdyn_workorderservicetask
            {
                msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId),
                msdyn_TaskType = serviceTaskType1Reference,
                statuscode = msdyn_workorderservicetask_statuscode.New
            });

            //ACT
            orgAdminService.Update(new msdyn_workorder { Id = workOrderId, msdyn_PrimaryIncidentType = incidentType2Reference });

            //ASSERT
            var query = new QueryExpression(msdyn_workorderservicetask.EntityLogicalName)
            {
                ColumnSet = new ColumnSet("statuscode")
            };
            var wosts = orgAdminUIService.RetrieveMultiple(query).Entities.Cast<msdyn_workorderservicetask>().ToList();

            //Expect the first WOST to now be inactive
            Assert.Equal(msdyn_workorderservicetask_statuscode.Inactive.ToString(), wosts[0].statuscode.ToString());

            //Expect the Task Type of the new Work Order Service Task (New) to be changed
            Assert.NotEqual(wost, wosts[1].Id);
        }

        [Fact]
        public void When_case_of_work_order_is_set_to_null_expect_case_of_all_findings_to_also_be_set_to_null()
        {
            //ARRANGE
            var case1Id = orgAdminUIService.Create(new Incident { });
            var case2Id = orgAdminUIService.Create(new Incident { });

            var workOrderId = orgAdminService.Create(new msdyn_workorder { msdyn_ServiceRequest = new EntityReference(Incident.EntityLogicalName, case1Id) });

            var finding1 = orgAdminService.Create(new ovs_Finding { ovs_CaseId = new EntityReference(Incident.EntityLogicalName, case1Id), ts_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId) });
            var finding2 = orgAdminService.Create(new ovs_Finding { ovs_CaseId = new EntityReference(Incident.EntityLogicalName, case1Id), ts_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId) });
            var finding3 = orgAdminService.Create(new ovs_Finding { ovs_CaseId = new EntityReference(Incident.EntityLogicalName, case1Id), ts_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId) });

            //ACT
            orgAdminService.Update(new msdyn_workorder { Id = workOrderId, msdyn_ServiceRequest = null });

            //ASSERT
            var query = new QueryExpression(ovs_Finding.EntityLogicalName)
            {
                ColumnSet = new ColumnSet("ovs_caseid")
            };
            var findings = orgAdminUIService.RetrieveMultiple(query).Entities.Cast<ovs_Finding>().ToList();

            //Expect all findings to have no case
            Assert.Null(findings[0].ovs_CaseId);
            Assert.Null(findings[1].ovs_CaseId);
            Assert.Null(findings[2].ovs_CaseId);
        }

        [Fact]
        public void When_case_of_work_order_is_changed_expect_all_findings_not_associated_to_work_order_to_remain_unchanged()
        {
            //ARRANGE
            var case1Id = orgAdminUIService.Create(new Incident { });
            var case2Id = orgAdminUIService.Create(new Incident { });
            var case3Id = orgAdminUIService.Create(new Incident { });

            var workOrder1Id = orgAdminService.Create(new msdyn_workorder { msdyn_ServiceRequest = new EntityReference(Incident.EntityLogicalName, case1Id) });
            var workOrder2Id = orgAdminService.Create(new msdyn_workorder { msdyn_ServiceRequest = new EntityReference(Incident.EntityLogicalName, case1Id) });

            //Make 3 findings related to case1 and workOrder1
            var finding1 = orgAdminService.Create(new ovs_Finding { ovs_CaseId = new EntityReference(Incident.EntityLogicalName, case1Id), ts_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrder1Id) });
            var finding2 = orgAdminService.Create(new ovs_Finding { ovs_CaseId = new EntityReference(Incident.EntityLogicalName, case1Id), ts_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrder1Id) });
            var finding3 = orgAdminService.Create(new ovs_Finding { ovs_CaseId = new EntityReference(Incident.EntityLogicalName, case1Id), ts_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrder1Id) });

            //Make 3 findings related to case2 and workOrder2
            var finding4 = orgAdminService.Create(new ovs_Finding { ovs_CaseId = new EntityReference(Incident.EntityLogicalName, case2Id), ts_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrder2Id) });
            var finding5 = orgAdminService.Create(new ovs_Finding { ovs_CaseId = new EntityReference(Incident.EntityLogicalName, case2Id), ts_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrder2Id) });
            var finding6 = orgAdminService.Create(new ovs_Finding { ovs_CaseId = new EntityReference(Incident.EntityLogicalName, case2Id), ts_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrder2Id) });

            //ACT
            //Update workOrder2 to be related to case3
            //All findings related to workOrder2 should then be updated to be related to case3
            orgAdminService.Update(new msdyn_workorder { Id = workOrder2Id, msdyn_ServiceRequest = new EntityReference(Incident.EntityLogicalName, case3Id) });

            //ASSERT
            var query = new QueryExpression(ovs_Finding.EntityLogicalName)
            {
                ColumnSet = new ColumnSet("ovs_caseid")
            };
            var case1findings = orgAdminUIService.RetrieveMultiple(query).Entities.Cast<ovs_Finding>().Where(f => f.ovs_CaseId.Id == case1Id).ToList();
            var case2findings = orgAdminUIService.RetrieveMultiple(query).Entities.Cast<ovs_Finding>().Where(f => f.ovs_CaseId.Id == case2Id).ToList();
            var case3findings = orgAdminUIService.RetrieveMultiple(query).Entities.Cast<ovs_Finding>().Where(f => f.ovs_CaseId.Id == case3Id).ToList();

            //Expect 3 findings to be related to case 1
            Assert.Equal(3, case1findings.Count);

            //Expect 0 findings to be related to case 2
            Assert.Empty(case2findings);

            //Expect 3 findings to be related to case 3
            Assert.Equal(3, case3findings.Count);
        }

        [Fact]
        public void When_case_of_work_order_is_changed_expect_all_work_order_service_tasks_related_to_work_order_to_be_related_to_new_case()
        {
            //ARRANGE
            var case1Id = orgAdminUIService.Create(new Incident { });
            var case2Id = orgAdminUIService.Create(new Incident { });

            var workOrderId = orgAdminService.Create(new msdyn_workorder { msdyn_ServiceRequest = new EntityReference(Incident.EntityLogicalName, case1Id) });

            var wost1 = orgAdminService.Create(new msdyn_workorderservicetask { ovs_CaseId = new EntityReference(Incident.EntityLogicalName, case1Id), msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId) });
            var wost2 = orgAdminService.Create(new msdyn_workorderservicetask { ovs_CaseId = new EntityReference(Incident.EntityLogicalName, case1Id), msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId) });
            var wost3 = orgAdminService.Create(new msdyn_workorderservicetask { ovs_CaseId = new EntityReference(Incident.EntityLogicalName, case1Id), msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId) });

            //ACT
            orgAdminService.Update(new msdyn_workorder { Id = workOrderId, msdyn_ServiceRequest = new EntityReference(Incident.EntityLogicalName, case2Id) });

            //ASSERT
            var query = new QueryExpression(msdyn_workorderservicetask.EntityLogicalName)
            {
                ColumnSet = new ColumnSet("ovs_caseid")
            };
            var wosts = orgAdminUIService.RetrieveMultiple(query).Entities.Cast<msdyn_workorderservicetask>().ToList();

            //Expect all wosts to be associated to Case2
            Assert.Equal(case2Id, wosts[0].ovs_CaseId.Id);
            Assert.Equal(case2Id, wosts[1].ovs_CaseId.Id);
            Assert.Equal(case2Id, wosts[2].ovs_CaseId.Id);
        }

        [Fact]
        public void When_case_of_work_order_is_changed_to_null_expect_all_work_order_service_tasks_related_to_work_order_to_have_no_case()
        {
            //ARRANGE
            var case1Id = orgAdminUIService.Create(new Incident { });
            var case2Id = orgAdminUIService.Create(new Incident { });

            var workOrderId = orgAdminService.Create(new msdyn_workorder { msdyn_ServiceRequest = new EntityReference(Incident.EntityLogicalName, case1Id) });

            var wost1 = orgAdminService.Create(new msdyn_workorderservicetask { ovs_CaseId = new EntityReference(Incident.EntityLogicalName, case1Id), msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId) });
            var wost2 = orgAdminService.Create(new msdyn_workorderservicetask { ovs_CaseId = new EntityReference(Incident.EntityLogicalName, case1Id), msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId) });
            var wost3 = orgAdminService.Create(new msdyn_workorderservicetask { ovs_CaseId = new EntityReference(Incident.EntityLogicalName, case1Id), msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId) });

            //ACT
            orgAdminService.Update(new msdyn_workorder { Id = workOrderId, msdyn_ServiceRequest = null });

            //ASSERT
            var query = new QueryExpression(msdyn_workorderservicetask.EntityLogicalName)
            {
                ColumnSet = new ColumnSet("ovs_caseid")
            };
            var wosts = orgAdminUIService.RetrieveMultiple(query).Entities.Cast<msdyn_workorderservicetask>().ToList();

            //Expect all wosts to have no case
            Assert.Null(wosts[0].ovs_CaseId);
            Assert.Null(wosts[1].ovs_CaseId);
            Assert.Null(wosts[2].ovs_CaseId);
        }

        [Fact]
        public void When_case_of_work_order_is_changed_expect_all_associated_files_to_be_related_to_new_case()
        {
            //ARRANGE
            var file = new ts_File();

            var caseId = orgAdminService.Create(new Incident()
            {
                Title = "My test case"
            });

            var case2Id = orgAdminService.Create(new Incident()
            {
                Title = "My second test case"
            });

            var workOrderId = orgAdminService.Create(new msdyn_workorder()
            {
                msdyn_name = "1234",
                msdyn_ServiceRequest = new EntityReference(Incident.EntityLogicalName, caseId)
            });

            //ACT
            file.Id = orgAdminService.Create(new ts_File()
            {
                ts_File_1 = "File test.txt",
                ts_FileContext = ts_filecontext.TC2020PROMOTINGCOMPLIANCE,
                ts_formintegrationid = "WO 1234",
                ts_VisibletoOtherPrograms = false,
            });

            orgAdminService.Update(new msdyn_workorder { Id = workOrderId, msdyn_ServiceRequest = new EntityReference(Incident.EntityLogicalName, case2Id) });

            //ASSERT
            var result = orgAdminService.Retrieve(ts_File.EntityLogicalName, file.Id, new ColumnSet(true)).ToEntity<ts_File>();

            //Expect the file to be associated to My second test case
            Assert.Equal("My second test case", result.ts_Incident.Name);
        }

    }
}
