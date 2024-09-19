using System;
using Microsoft.Xrm.Sdk;
using DG.Tools.XrmMockup;
using Xunit;
using System.Collections.Generic;
using Microsoft.Crm.Sdk.Messages;
using System.Linq;

namespace ROMTS_GSRST.Plugins.Tests
{
    [Collection("Xrm Collection")]
    public class UnitTestBase : IClassFixture<XrmMockupFixture>
    {
        private static DateTime _startTime { get; set; }

        protected IOrganizationService orgAdminUIService;
        protected IOrganizationService orgAdminService;
        protected static XrmMockup365 crm;

        public UnitTestBase(XrmMockupFixture fixture)
        {
            crm = fixture.crm;
            crm.ResetEnvironment();
            orgAdminUIService = crm.GetAdminService(new MockupServiceSettings(true, false, MockupServiceSettings.Role.UI));
            orgAdminService = crm.GetAdminService();
        }

        public void Dispose()
        {
            crm.ResetEnvironment();
        }
    }
}
