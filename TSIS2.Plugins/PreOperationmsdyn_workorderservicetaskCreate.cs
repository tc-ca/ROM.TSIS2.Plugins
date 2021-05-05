using System;
using System.Linq;
using Microsoft.Xrm.Sdk;

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
                    if (target.LogicalName.Equals(msdyn_workorderservicetask.EntityLogicalName))
                    {
                        // Cast the target to the expected entity
                        msdyn_workorderservicetask workOrderServiceTask = target.ToEntity<msdyn_workorderservicetask>();

                        if (workOrderServiceTask.msdyn_name != null)
                        {
                            // Set the new work order service task to be prefixed with the parent work order name
                            using (var serviceContext = new Xrm(service))
                            {
                                msdyn_workorder workOrder = serviceContext.msdyn_workorderSet.FirstOrDefault(wo => wo.Id == workOrderServiceTask.msdyn_WorkOrder.Id);

                                // Relationships are lazy loaded. Need to explicitly call load property to get the related work order service tasks.
                                serviceContext.LoadProperty(workOrder, "msdyn_msdyn_workorder_msdyn_workorderservicetask_WorkOrder");

                                // Suffix based off previous number of work order service tasks.
                                var workOrderServiceTasks = workOrder.msdyn_msdyn_workorder_msdyn_workorderservicetask_WorkOrder;

                                // Set the prefix to be at the 200 level for work order service tasks
                                var prefix = workOrder.msdyn_name.Replace("300-", "200-");

                                // If there are previous work order service tasks, suffix = count + 1 else 1
                                var suffix = (workOrderServiceTasks != null) ? workOrderServiceTasks.Count() + 1 : 1;
                                workOrderServiceTask.msdyn_name = string.Format("{0}-{1}", prefix, suffix);
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
                                        Entity q = service.Retrieve("ovs_questionnaire", servicetasktype.ovs_Questionnaire.Id,
                                            new Microsoft.Xrm.Sdk.Query.ColumnSet("ovs_questionnairedefinition"));
                                        target.Attributes["ovs_questionnairedefinition"] = q.Attributes["ovs_questionnairedefinition"].ToString();
                                    }
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
    }
}