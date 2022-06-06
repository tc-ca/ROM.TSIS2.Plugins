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

    [CrmPluginRegistration(
    MessageNameEnum.Update,
    "incident",
    StageEnum.PreOperation,
    ExecutionModeEnum.Synchronous,
    "",
    "TSIS2.Plugins.PreOperationincidentUpdate Plugin",
    1,
    IsolationModeEnum.Sandbox,
    Description = "On Closing Case, validate if there is any Finding is not set to complete status")]
    /// <summary>
    /// PreOperationincidentUpdate Plugin.
    /// </summary>    
    public class PreOperationincidentUpdate : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            // Obtain the tracing service
            ITracingService tracingService =
            (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)
                serviceProvider.GetService(typeof(IPluginExecutionContext));

            // The InputParameters collection contains all the data passed in the message request.
            if (context.InputParameters.Contains("Target") &&
                context.InputParameters["Target"] is Entity)
            {
                // Obtain the target entity from the input parameters.
                Entity target = (Entity)context.InputParameters["Target"];

                // Obtain the organization service reference which you will need for
                // web service calls.
                IOrganizationServiceFactory serviceFactory =
                    (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

                try
                {
                    if (target.LogicalName.Equals(Incident.EntityLogicalName))
                    {
                        if (target.Attributes.Contains("statecode") && (int) target.Attributes["statecode"] == 1)
                        {
                            using (var servicecontext = new Xrm(service))
                            {
                                int UserLanguage = RetrieveUserUILanguageCode(service, context.InitiatingUserId);
                                var findingEntities = servicecontext.ovs_FindingSet.Where(f => f.ovs_CaseId.Id == target.Id && f.ts_findingtype == ts_findingtype.Noncompliance && f.statuscode != ovs_Finding_statuscode.Complete).ToList();
                                if (findingEntities != null && findingEntities.Count>0)
                                {
                                    string ErrorMessage = "Can not close case.";
                                    if (UserLanguage == 1036) //French
                                    {
                                        ErrorMessage = "fr-Can not close case.";
                                    }

                                    throw new InvalidPluginExecutionException(ErrorMessage);
                                }
                                else
                                {
                                    throw new InvalidPluginExecutionException("Can not close case.");
                                }
                            }
                        }
        
                    }
                }
                catch (Exception e)
                {
                    throw new InvalidPluginExecutionException(e.Message);
                }

            }
        }

        protected static int RetrieveUserUILanguageCode(IOrganizationService service, Guid userId)
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

        protected static XmlDocument RetrieveXmlWebResourceByName(IOrganizationService service, ITracingService tracingService, string webresourceSchemaName)
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
        protected static string RetrieveLocalizedStringFromWebResource(ITracingService tracingService, XmlDocument resource, string resourceId)
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