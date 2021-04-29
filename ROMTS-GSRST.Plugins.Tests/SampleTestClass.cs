using System;
using DG.XrmContext;
using System.ServiceModel;
using Xunit;
using Xunit.Sdk;

namespace ROMTS_GSRST.Plugins.Tests
{
    public class SampleTestClass : UnitTestBase
    {
        [Fact]
        public void TestPrimaryContactIsCreated()
        {
            using (var context = new Xrm(orgAdminUIService))
            {

            }
        }
    }
}
