using DG.Tools.XrmMockup;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TSIS2.Plugins;

namespace ROMTS_GSRST.Plugins.Tests
{
    public class XrmMockupFixture : IDisposable
    {
        public XrmMockup365 crm;

        public XrmMockupFixture()
        {
            var settings = new XrmMockupSettings
            {
                BasePluginTypes = new Type[] { typeof(PluginBase) },
                EnableProxyTypes = true,
            };

            crm = XrmMockup365.GetInstance(settings);
        }

        public void Dispose()
        {

        }
    }
}
