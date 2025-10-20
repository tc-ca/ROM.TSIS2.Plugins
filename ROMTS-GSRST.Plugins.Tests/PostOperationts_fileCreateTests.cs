// Commented out since it no longer works for .Net Framework 4.7.1
//using System;
//using System.Linq;
//using Microsoft.Xrm.Sdk;
//using Microsoft.Xrm.Sdk.Query;
//using Xunit;

//namespace ROMTS_GSRST.Plugins.Tests
//{
//    public class PostOperationts_fileCreateTests : UnitTestBase
//    {

//        public PostOperationts_fileCreateTests(XrmMockupFixture fixture) : base(fixture)
//        {
//        }

//        [Fact]
//        public void Create_test_file()
//        {
//            using (var context = new Xrm(orgAdminUIService))
//            {
//                // ARRANGE
//                var file = new ts_File();
//                orgAdminService.Create(new Entity("team", new Guid("e2e3910d-a41f-ec11-b6e6-0022483cb5c7")));
//                orgAdminService.Create(new Entity("team", new Guid("dc3f2b10-28f5-eb11-94ef-000d3af36036")));

//                // ACT
//                file.Id = orgAdminService.Create(new ts_File()
//                {
//                    ts_File_1 = "File test.txt",
//                    ts_FileContext = ts_filecontext.TC2020PROMOTINGCOMPLIANCE,
//                    ts_ProgramAccessTeamNameID = @"e2e3910d-a41f-ec11-b6e6-0022483cb5c7,
//                                                    dc3f2b10-28f5-eb11-94ef-000d3af36036,",
//                    ts_VisibletoOtherPrograms = true,
//                });

//                var result = orgAdminUIService.Retrieve(ts_File.EntityLogicalName, file.Id, new ColumnSet("ts_file")).ToEntity<ts_File>();

//                // ASSERT
//                // Expect the file to be created with 'File test.txt' as the attachment
//                Assert.Equal("File test.txt", result.ts_File_1);
//            }
//        }

//        [Fact]
//        public void Create_test_file_with_no_work_order_in_work_order_service_task()
//        {
//            using (var context = new Xrm(orgAdminService))
//            {
//                // ARRANGE
//                var file = new ts_File();

//                var myCaseId = orgAdminService.Create(new Incident()
//                {
//                    Title = "My test case"
//                });

//                var workOrderId = orgAdminService.Create(new msdyn_workorder()
//                {
//                    msdyn_name = "1234",
//                    msdyn_ServiceRequest = new EntityReference(Incident.EntityLogicalName, myCaseId)
//                });

//                var workOrderServiceTaskId = orgAdminService.Create(new msdyn_workorderservicetask()
//                {
//                    msdyn_name = "1234",
//                    msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId)
//                });

//                // ACT
//                file.Id = orgAdminService.Create(new ts_File()
//                {
//                    ts_File_1 = "File test.txt",
//                    ts_FileContext = ts_filecontext.TC2020PROMOTINGCOMPLIANCE,
//                    ts_formintegrationid = null,
//                    ts_VisibletoOtherPrograms = false,
//                });

//                // ASSERT
//                var result = orgAdminService.Retrieve(ts_File.EntityLogicalName, file.Id, new ColumnSet(true)).ToEntity<ts_File>();

//                // Expect the file to be created with no Work Order
//                Assert.Null(result.ts_msdyn_workorder);
//            }
//        }

//        [Fact]
//        public void Create_test_file_with_case_and_work_order_from_work_order_service_task()
//        {
//            using (var context = new Xrm(orgAdminService))
//            {
//                // ARRANGE
//                var file = new ts_File();
//                var myCaseId = orgAdminService.Create(new Incident()
//                {
//                    Title = "My test case"
//                });

//                var workOrderId = orgAdminService.Create(new msdyn_workorder()
//                {
//                    msdyn_name = "1234",
//                    msdyn_ServiceRequest = new EntityReference(Incident.EntityLogicalName, myCaseId)
//                });

//                var workOrderServiceTaskId = orgAdminService.Create(new msdyn_workorderservicetask()
//                {
//                    msdyn_name = "1234",
//                    msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId)
//                });

//                // ACT
//                file.Id = orgAdminService.Create(new ts_File()
//                {
//                    ts_File_1 = "File test.txt",
//                    ts_FileContext = ts_filecontext.TC2020PROMOTINGCOMPLIANCE,
//                    ts_formintegrationid = "WOST 1234-1",
//                    ts_VisibletoOtherPrograms = false,
//                });

//                // ASSERT
//                var result = orgAdminService.Retrieve(ts_File.EntityLogicalName, file.Id, new ColumnSet(true)).ToEntity<ts_File>();

//                // Expect the file to be associated with Work Order 1234
//                Assert.Equal("1234", result.ts_msdyn_workorder.Name);

//                // Expect the file to be associated with Case 'My test case'
//                Assert.Equal("My test case", result.ts_Incident.Name);

//                // Expect the document type to be 'Work Order Service Task'
//                Assert.Equal(ts_documenttype.WorkOrderServiceTask, result.ts_DocumentType);
//            }
//        }

//        [Fact]
//        public void Create_test_file_with_no_case_and_work_order_from_work_order_service_task()
//        {
//            using (var context = new Xrm(orgAdminService))
//            {
//                // ARRANGE
//                var file = new ts_File();

//                var workOrderId = orgAdminService.Create(new msdyn_workorder()
//                {
//                    msdyn_name = "1234"
//                });

//                var workOrderServiceTaskId = orgAdminService.Create(new msdyn_workorderservicetask()
//                {
//                    msdyn_name = "1234",
//                    msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId)
//                });

//                // ACT
//                file.Id = orgAdminService.Create(new ts_File()
//                {
//                    ts_File_1 = "File test.txt",
//                    ts_FileContext = ts_filecontext.TC2020PROMOTINGCOMPLIANCE,
//                    ts_formintegrationid = "WOST 1234-1",
//                    ts_VisibletoOtherPrograms = false,
//                });

//                // ASSERT
//                var result = orgAdminService.Retrieve(ts_File.EntityLogicalName, file.Id, new ColumnSet(true)).ToEntity<ts_File>();

//                // Expect the file to be associated with Work Order 1234
//                Assert.Equal("1234", result.ts_msdyn_workorder.Name);

//                // Expect the file to not be associated with a Case
//                Assert.Null(result.ts_Incident);

//                // Expect the document type to be 'Work Order Service Task'
//                Assert.Equal(ts_documenttype.WorkOrderServiceTask, result.ts_DocumentType);
//            }
//        }

//        [Fact]
//        public void Create_test_file_with_case_from_work_order()
//        {
//            using (var context = new Xrm(orgAdminService))
//            {
//                // ARRANGE
//                var file = new ts_File();
//                var myCaseId = orgAdminService.Create(new Incident()
//                {
//                    Title = "My test case"
//                });

//                var workOrderId = orgAdminService.Create(new msdyn_workorder()
//                {
//                    msdyn_name = "1234",
//                    msdyn_ServiceRequest = new EntityReference(Incident.EntityLogicalName, myCaseId)
//                });

//                // ACT
//                file.Id = orgAdminService.Create(new ts_File()
//                {
//                    ts_File_1 = "File test.txt",
//                    ts_FileContext = ts_filecontext.TC2020PROMOTINGCOMPLIANCE,
//                    ts_formintegrationid = "WO 1234",
//                    ts_VisibletoOtherPrograms = false,
//                });

//                // ASSERT
//                var result = orgAdminService.Retrieve(ts_File.EntityLogicalName, file.Id, new ColumnSet(true)).ToEntity<ts_File>();

//                // Expect the file to not be associated with a Work Order
//                // Note: Remember files coming in from Document Centre - Work Order do not need to have the ts_msdyn_workorder populated,
//                // only files coming in from Document Centre - Work Order Service Task use that.
//                Assert.Null(result.ts_msdyn_workorder);

//                // Expect the file to be associated with Case 'My test case'
//                Assert.Equal("My test case", result.ts_Incident.Name);

//                // Expect the document type to be 'Work Order'
//                Assert.Equal(ts_documenttype.WorkOrder, result.ts_DocumentType);
//            }
//        }

//        [Fact]
//        public void Create_test_file_with_no_case_from_work_order()
//        {
//            using (var context = new Xrm(orgAdminService))
//            {
//                // ARRANGE
//                var file = new ts_File();

//                var workOrderId = orgAdminService.Create(new msdyn_workorder()
//                {
//                    msdyn_name = "1234"
//                });

//                // ACT
//                file.Id = orgAdminService.Create(new ts_File()
//                {
//                    ts_File_1 = "File test.txt",
//                    ts_FileContext = ts_filecontext.TC2020PROMOTINGCOMPLIANCE,
//                    ts_formintegrationid = "WO 1234",
//                    ts_VisibletoOtherPrograms = false,
//                });

//                // ASSERT
//                var result = orgAdminService.Retrieve(ts_File.EntityLogicalName, file.Id, new ColumnSet(true)).ToEntity<ts_File>();

//                // Expect the file to not be associated with a Work Order
//                // Note: Remember files coming in from Document Centre - Work Order do not need to have the ts_msdyn_workorder populated,
//                // only files coming in from Document Centre - Work Order Service Task use that.
//                Assert.Null(result.ts_msdyn_workorder);

//                // Expect the file to not be associated with a Case
//                Assert.Null(result.ts_Incident);

//                // Expect the document type to be 'Work Order'
//                Assert.Equal(ts_documenttype.WorkOrder, result.ts_DocumentType);
//            }
//        }
//    }
//}
