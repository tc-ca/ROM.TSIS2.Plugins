using System;

namespace TSIS2.Plugins
{
    [CrmPluginRegistration(
        "ts_TSGlobalAction",
        "none",
        StageEnum.PostOperation,
        ExecutionModeEnum.Synchronous,
        "",
        "TSIS2.Plugins.TSGlobalAction Plugin",
        1,
        IsolationModeEnum.Sandbox,
        Description = "Custom action for global operations")]
    public class TSGlobalAction : PluginBase
    {
        public static string messageName = "ts_TSGlobalAction";

        public TSGlobalAction() : base(typeof(TSGlobalAction))
        {
        }

        public TSGlobalAction(string unsecure, string secure) : base(typeof(TSGlobalAction))
        {
        }

        protected override void ExecuteCrmPlugin(LocalPluginContext localContext)
        {
            var tracingService = localContext.TracingService;
            var context = localContext.PluginExecutionContext;
            var service = localContext.OrganizationService;

            try
            {
                if (context.MessageName.ToLower() == messageName.ToLower())
                {
                    string action = context.InputParameters["actionName"] as string;
                    string recordIdString = context.InputParameters["recordId"] as string;
                    string param1 = recordIdString.Split('|')[0];
                    string lang = recordIdString.Split('|')[1];
                    //string param2 = context.InputParameters["Param2"] as string;


                    localContext.Trace(string.Format("ts_TSGlobalAction: {1}     param1-{0} ", param1, action));

                    string retMessage = "";
                    string retMessage2 = "";
                    if (!string.IsNullOrEmpty(action) && action == "GetOperationActivityWO")
                    {
                        var retTuple = retrieveSearchHtmlTableLogicDataTableHelper.searchMatchingRecords( param1, service, tracingService);

                        localContext.Trace(" -----------  WO count: " + retTuple.Entities.Count);
                        var retObj = retrieveSearchHtmlTableLogicDataTableHelper.ConvertEntityCollectionToHtmlDataTable(retTuple, false, service, lang);
                        retMessage = retObj.Item1;
                        retMessage2 = retObj.Item2;

                    }


                    localContext.Trace(" -----------  RetMsg: " + retMessage);
                    localContext.Trace(" -----------  RetMsg2: " + retMessage2);
                    context.OutputParameters["result"] = retMessage;
                    context.OutputParameters["result2"] = retMessage2;
                    //context.OutputParameters["RetId"] = "No Guid";
                }
            }
            catch (Exception ex)
            {
                localContext.TraceWithContext("TSGlobalAction Plugin: {0}", ex);
                throw new Exception(ex.Message);
            }
        }

        public static string BuildHTMLTable(string recordId)
        {

            return "";
        }


    }
}
