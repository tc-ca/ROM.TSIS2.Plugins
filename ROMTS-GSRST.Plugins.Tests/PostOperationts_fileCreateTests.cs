using System;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Xunit;

namespace ROMTS_GSRST.Plugins.Tests
{
    public class PostOperationts_fileCreateTests : UnitTestBase
    {

        public PostOperationts_fileCreateTests(XrmMockupFixture fixture) : base(fixture)
        {
        }

        [Fact]
        public void Create_test_file()
        {
            using (var context = new Xrm(orgAdminUIService))
            {
                // ARRANGE
                var file = new ts_File();
                orgAdminService.Create(new Entity("team", new Guid("e2e3910d-a41f-ec11-b6e6-0022483cb5c7")));
                orgAdminService.Create(new Entity("team", new Guid("dc3f2b10-28f5-eb11-94ef-000d3af36036")));

                // ACT
                file.Id = orgAdminService.Create(new ts_File()
                {
                    ts_File_1 = "File test.txt",
                    ts_FileContext = ts_filecontext.TC2020PROMOTINGCOMPLIANCE,
                    ts_ProgramAccessTeamNameID = @"e2e3910d-a41f-ec11-b6e6-0022483cb5c7,
                                                    dc3f2b10-28f5-eb11-94ef-000d3af36036,",
                    ts_VisibletoOtherPrograms = true,
                });

                var result = orgAdminUIService.Retrieve(ts_File.EntityLogicalName, file.Id, new ColumnSet("ts_file")).ToEntity<ts_File>();

                // ASSERT
                // Expect the file to be created with 'File test.txt' as the attachment
                Assert.Equal("File test.txt", result.ts_File_1);
            }
        }

        [Fact]
        public void Create_test_file_with_no_work_order_service_task()
        {
            using (var context = new Xrm(orgAdminService))
            {
                // ARRANGE
                var file = new ts_File();

                var myCaseId = orgAdminService.Create(new Incident()
                {
                    Title = "My test case"
                });

                var workOrderId = orgAdminService.Create(new msdyn_workorder()
                {
                    msdyn_name = "1234",
                    msdyn_ServiceRequest = new EntityReference(Incident.EntityLogicalName, myCaseId)
                });

                var workOrderServiceTaskId = orgAdminService.Create(new msdyn_workorderservicetask()
                {
                    msdyn_name = "1234",
                    msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId)
                });

                // ACT
                file.Id = orgAdminService.Create(new ts_File()
                {
                    ts_File_1 = "File test.txt",
                    ts_FileContext = ts_filecontext.TC2020PROMOTINGCOMPLIANCE,
                    ts_formintegrationid = null,
                    ts_VisibletoOtherPrograms = false,
                });

                // ASSERT
                var result = orgAdminService.Retrieve(ts_File.EntityLogicalName, file.Id, new ColumnSet(true)).ToEntity<ts_File>();

                // Expect the file to be created with no Work Order
                Assert.Null(result.ts_msdyn_workorder);
            }
        }

        [Fact]
        public void Create_test_file_with_work_order_service_task()
        {
            using (var context = new Xrm(orgAdminService))
            {
                // ARRANGE
                var file = new ts_File();
                var myCaseId = orgAdminService.Create(new Incident()
                {
                    Title = "My test case"
                });

                var workOrderId = orgAdminService.Create(new msdyn_workorder()
                {
                    msdyn_name = "1234",
                    msdyn_ServiceRequest = new EntityReference(Incident.EntityLogicalName, myCaseId)
                });

                var workOrderServiceTaskId = orgAdminService.Create(new msdyn_workorderservicetask()
                {
                    msdyn_name = "1234",
                    msdyn_WorkOrder = new EntityReference(msdyn_workorder.EntityLogicalName, workOrderId)
                });

                // ACT
                file.Id = orgAdminService.Create(new ts_File()
                {
                    ts_File_1 = "File test.txt",
                    ts_FileContext = ts_filecontext.TC2020PROMOTINGCOMPLIANCE,
                    ts_formintegrationid = "WOST 1234-1",
                    ts_VisibletoOtherPrograms = false,
                });

                // ASSERT
                var result = orgAdminService.Retrieve(ts_File.EntityLogicalName, file.Id, new ColumnSet(true)).ToEntity<ts_File>();

                // Expect the file to be associated with Work Order 1234
                Assert.Equal("1234", result.ts_msdyn_workorder.Name);

                // Expect the file to be associated with Case 'My test case'
                Assert.Equal("My test case", result.ts_Incident.Name);
            }
        }
    }
}
