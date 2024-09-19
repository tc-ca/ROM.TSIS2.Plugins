using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Runtime.Remoting.Contexts;
using System.Text;
using System.Threading.Tasks;
using System.Web.Services.Description;

namespace TSIS2.Plugins
{
    public class TSGlobalAction : IPlugin
    {
        public static string messageName = "ts_TSGlobalAction";
        public void Execute(IServiceProvider serviceProvider)
        {
            // Obtain the tracing service
            ITracingService tracingService =
            (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)
                serviceProvider.GetService(typeof(IPluginExecutionContext));

            IOrganizationServiceFactory serviceFactory =
                    (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));

            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            try
            {
                if (context.MessageName.ToLower() == messageName.ToLower())
                {
                    string action = context.InputParameters["actionName"] as string;
                    string recordIdString = context.InputParameters["recordId"] as string;
                    string param1 = recordIdString.Split('|')[0];
                    string lang = recordIdString.Split('|')[1];
                    //string param2 = context.InputParameters["Param2"] as string;


                    tracingService.Trace(string.Format("ts_TSGlobalAction: {1}     param1-{0} ", param1, action));

                    string retMessage = "";
                    string retMessage2 = "";
                    if (!string.IsNullOrEmpty(action) && action == "GetOperationActivityWO")
                    {
                        var retTuple = retrieveSearchHtmlTableLogicDataTableHelper.searchMatchingRecords( param1, service, tracingService);

                        tracingService.Trace(" -----------  WO count: " + retTuple.Entities.Count);
                        var retObj = retrieveSearchHtmlTableLogicDataTableHelper.ConvertEntityCollectionToHtmlDataTable(retTuple, false, service, lang);
                        retMessage = retObj.Item1;
                        retMessage2 = retObj.Item2;

                    }


                    tracingService.Trace(" -----------  RetMsg: " + retMessage);
                    tracingService.Trace(" -----------  RetMsg2: " + retMessage2);
                    context.OutputParameters["result"] = retMessage;
                    context.OutputParameters["result2"] = retMessage2;
                    //context.OutputParameters["RetId"] = "No Guid";
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static string BuildHTMLTable(string recordId)
        {

            return "";
        }


    }
}
