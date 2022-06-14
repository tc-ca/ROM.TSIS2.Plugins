using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace TSIS2.Plugins
{
    public class LocalizationHelper
    {
        public static string GetMessage(ITracingService tracingService, IOrganizationService service, string ResourceFile, string ResourceId)
        {
            XmlDocument messages = RetrieveXmlWebResourceByName(service, tracingService, ResourceFile);
            return RetrieveLocalizedStringFromWebResource(tracingService, messages, ResourceId);
        }

        public static int RetrieveUserUILanguageCode(IOrganizationService service, Guid userId)
        {
            QueryExpression userSettingsQuery = new QueryExpression("usersettings");
            userSettingsQuery.ColumnSet.AddColumns("uilanguageid", "systemuserid");
            userSettingsQuery.Criteria.AddCondition("systemuserid", ConditionOperator.Equal, userId);
            EntityCollection userSettings = service.RetrieveMultiple(userSettingsQuery);
            if (userSettings.Entities.Count > 0)
            {
                return (int)userSettings.Entities[0]["uilanguageid"];
            }
            return 0;
        }

        public static XmlDocument RetrieveXmlWebResourceByName(IOrganizationService service, ITracingService tracingService, string webresourceSchemaName)
        {
            tracingService.Trace("Begin:RetrieveXmlWebResourceByName, webresourceSchemaName={0}", webresourceSchemaName);
            QueryExpression webresourceQuery = new QueryExpression("webresource");
            webresourceQuery.ColumnSet.AddColumn("content");
            webresourceQuery.Criteria.AddCondition("name", ConditionOperator.Equal, webresourceSchemaName);
            EntityCollection webresources = service.RetrieveMultiple(webresourceQuery);
            tracingService.Trace("Webresources Returned from server. Count={0}", webresources.Entities.Count);
            if (webresources.Entities.Count > 0)
            {
                byte[] bytes = Convert.FromBase64String((string)webresources.Entities[0]["content"]);
                // The bytes would contain the ByteOrderMask. Encoding.UTF8.GetString() does not remove the BOM.  
                // Stream Reader auto detects the BOM and removes it on the text  
                XmlDocument document = new XmlDocument();
                document.XmlResolver = null;
                using (MemoryStream ms = new MemoryStream(bytes))
                {
                    using (StreamReader sr = new StreamReader(ms))
                    {
                        document.Load(sr);
                    }
                }
                tracingService.Trace("End:RetrieveXmlWebResourceByName , webresourceSchemaName={0}", webresourceSchemaName);
                return document;
            }
            else
            {
                tracingService.Trace("{0} Webresource missing. Reinstall the solution", webresourceSchemaName);
                throw new InvalidPluginExecutionException(String.Format("Unable to locate the web resource {0}.", webresourceSchemaName));
            }
        }
        public static string RetrieveLocalizedStringFromWebResource(ITracingService tracingService, XmlDocument resource, string resourceId)
        {
            XmlNode valueNode = resource.SelectSingleNode(string.Format(CultureInfo.InvariantCulture, "./root/data[@name='{0}']/value", resourceId));
            if (valueNode != null)
            {
                return valueNode.InnerText;
            }
            else
            {
                tracingService.Trace("No Node Found for {0} ", resourceId);
                throw new InvalidPluginExecutionException(String.Format("ResourceID {0} was not found.", resourceId));
            }
        }
    }
}
