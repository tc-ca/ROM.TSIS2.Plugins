using System;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace TSIS2.Plugins
{

    [CrmPluginRegistration(
    MessageNameEnum.Create,
    "msdyn_workorderservicetask",
    StageEnum.PreOperation,
    ExecutionModeEnum.Synchronous,
    "",
    "TSIS2.Plugins.PreOperationmsdyn_workorderservicetaskCreate Plugin",
    1,
    IsolationModeEnum.Sandbox,
    Description = "Copies the JSON Questionnaire Definition from a Service Task to a Work Order Service Task.")]
    /// <summary>
    /// PostOperationmsdyn_workorderservicetaskCreate Plugin.
    /// </summary>
    public class PreOperationmsdyn_workorderservicetaskCreate : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            // Obtain the tracing service
            ITracingService tracingService =
            (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            tracingService.Trace("Tracking Service Started.");

            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)
                serviceProvider.GetService(typeof(IPluginExecutionContext));

            tracingService.Trace("The InputParameters collection contains all the data passed in the message request.");
            if (context.InputParameters.Contains("Target") &&
                context.InputParameters["Target"] is Entity)
            {
                tracingService.Trace("Obtain the target entity from the input parameters.");
                Entity target = (Entity)context.InputParameters["Target"];

                tracingService.Trace("Obtain the organization service reference which you will need for web service calls.");
                IOrganizationServiceFactory serviceFactory =
                    (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

                try
                {
                    // Log the system username at the start
                    var systemUser = service.Retrieve("systemuser", context.InitiatingUserId, new ColumnSet("fullname"));
                    tracingService.Trace("Plugin executed by user: {0}", systemUser.GetAttributeValue<string>("fullname"));

                    if (target.LogicalName.Equals(msdyn_workorderservicetask.EntityLogicalName))
                    {
                        tracingService.Trace("Cast the target to the expected entity.");
                        msdyn_workorderservicetask workOrderServiceTask = target.ToEntity<msdyn_workorderservicetask>();

                        if (workOrderServiceTask.msdyn_name != null)
                        {
                            tracingService.Trace("Set the new Work Order service task to be prefixed with the parent Work Order name.");
                            using (var serviceContext = new Xrm(service))
                            {
                                msdyn_workorder workOrder = serviceContext.msdyn_workorderSet.FirstOrDefault(wo => wo.Id == workOrderServiceTask.msdyn_WorkOrder.Id);
                                if (workOrder != null)
                                {
                                    tracingService.Trace("Successfully retrieved Work Order with name: {0}", workOrder.msdyn_name);
                                }
                                else
                                {
                                    tracingService.Trace("Work Order not found for Work Order Service Task ID: {0}", workOrderServiceTask.msdyn_WorkOrder.Id);
                                }

                                tracingService.Trace("Relationships are lazy loaded. Need to explicitly call load property to get the related Work Order service tasks.");
                                serviceContext.LoadProperty(workOrder, "msdyn_msdyn_workorder_msdyn_workorderservicetask_WorkOrder");

                                tracingService.Trace("Suffix based off previous number of Work Order service tasks.");
                                var workOrderServiceTasks = workOrder.msdyn_msdyn_workorder_msdyn_workorderservicetask_WorkOrder;

                                tracingService.Trace("Set the prefix to be at the 200 level for Work Order service tasks.");
                                var prefix = workOrder.msdyn_name.Replace("300-", "200-");

                                tracingService.Trace("If there are previous Work Order service tasks, suffix = count + 1 else 1.");
                                var suffix = (workOrderServiceTasks != null) ? workOrderServiceTasks.Count() + 1 : 1;
                                workOrderServiceTask.msdyn_name = string.Format("{0}-{1}", prefix, suffix);

                            }
                        }
                        // Log the Work Order Service Task Name
                        tracingService.Trace("Work Order Service Task Name: {0}", workOrderServiceTask.msdyn_name);

                        tracingService.Trace("Check Mandatory field from Task Type.");
                        if (target.Contains("msdyn_tasktype") && target["msdyn_tasktype"] != null)
                        {
                            var theType = service.Retrieve("msdyn_servicetasktype", target.GetAttributeValue<EntityReference>("msdyn_tasktype").Id, new ColumnSet("ts_mandatory"));

                            if(theType.Contains("ts_mandatory"))
                            {
                                target["ts_mandatory"] = theType["ts_mandatory"];
                            }
                        }


                        if (target.Attributes.Contains("msdyn_tasktype") && target.Attributes["msdyn_tasktype"] != null
                            && (!target.Attributes.Contains("ovs_questionnaire") || target.Attributes["ovs_questionnaire"] != null))
                        {
                            EntityReference tasktype = (EntityReference)target.Attributes["msdyn_tasktype"];
                            using (var servicecontext = new Xrm(service))
                            {
                                var servicetasktype = (from tt in servicecontext.msdyn_servicetasktypeSet
                                                       where tt.Id == tasktype.Id
                                                       select new
                                                       {
                                                           tt.ovs_Questionnaire
                                                       }).FirstOrDefault();
                                if (servicetasktype != null)
                                {
                                    target.Attributes["ovs_questionnaire"] = servicetasktype.ovs_Questionnaire;
                                    if (servicetasktype.ovs_Questionnaire != null && servicetasktype.ovs_Questionnaire.Id != null)
                                    {

                                        tracingService.Trace("Retrieve Questionnaire Versions. Initiate QueryExpression query.");
                                        var questionnaireVersionsQuery = new QueryExpression("ts_questionnaireversion");

                                        tracingService.Trace("Add columns to query.ColumnSet.");
                                        questionnaireVersionsQuery.ColumnSet.AddColumns("ts_questionnairedefinition");
                                        questionnaireVersionsQuery.AddOrder("modifiedon", OrderType.Descending);

                                        tracingService.Trace("Define filter query.Criteria.");
                                        questionnaireVersionsQuery.Criteria.AddCondition("ts_ovs_questionnaire", ConditionOperator.Equal, servicetasktype.ovs_Questionnaire.Id);

                                        var questionnaireVersions = service.RetrieveMultiple(questionnaireVersionsQuery);

                                        if (questionnaireVersions[0] != null)
                                        {
                                            tracingService.Trace("For now, use the most recently modified version. Will replace with Effective Date logic later.");
                                            var latestQuestionnaireVersion = questionnaireVersions[0];

                                            tracingService.Trace("Set questionnaire definition to latest questionnaire versions' definition.");
                                            target.Attributes["ovs_questionnairedefinition"] = latestQuestionnaireVersion.Attributes["ts_questionnairedefinition"].ToString();
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception e)
                {
                    tracingService.Trace("Error occurred in the plugin execution: {0}", e.Message);
                    throw new InvalidPluginExecutionException(e.Message);
                }
            }
        }
    }
}