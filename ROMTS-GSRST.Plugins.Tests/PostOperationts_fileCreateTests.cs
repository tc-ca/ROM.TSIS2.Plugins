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
                    ts_ProgramAccessTeamNameID= @"e2e3910d-a41f-ec11-b6e6-0022483cb5c7,
                                                    dc3f2b10-28f5-eb11-94ef-000d3af36036,",
                    ts_VisibletoOtherPrograms = true,
                });

                var result = orgAdminUIService.Retrieve(ts_File.EntityLogicalName, file.Id, new ColumnSet("ts_file")).ToEntity<ts_File>();

                // ASSERT
                // Expect the file to be created with 'File test.txt' as the attachment
                Assert.Equal("File test.txt", result.ts_File_1);
            }
        }
    }
}
