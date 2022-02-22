using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TSIS2.Plugins
{
    [CrmPluginRegistration(
    MessageNameEnum.Delete,
    "msdyn_workorderservicetask",
    StageEnum.PreOperation,
    ExecutionModeEnum.Synchronous,
    "",
    "TSIS2.Plugins.PreOperationmsdyn_workorderservicetaskDelete Plugin",
    1,
    IsolationModeEnum.Sandbox,
        Image1Name = "PreImage", Image1Type = ImageTypeEnum.PreImage, Image1Attributes = "msdyn_name",
    Description = "On Work Order Service Task Delete, delete any associated files.")]
    public class PreOperationmsdyn_workorderservicetaskDelete : IPlugin
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
                context.InputParameters["Target"] is EntityReference)
            {
                // Obtain the target entity from the input parameters.
                EntityReference target = (EntityReference)context.InputParameters["Target"];

                // Obtain the preimage entity - use this to get the Work Order Service Task ID
                Entity preImageEntity = (context.PreEntityImages != null && context.PreEntityImages.Contains("PreImage")) ? context.PreEntityImages["PreImage"] : null;

                // Obtain the organization service reference which you will need for
                // web service calls.
                IOrganizationServiceFactory serviceFactory =
                    (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

                try
                {
                    if (target.LogicalName.Equals(msdyn_workorderservicetask.EntityLogicalName))
                    {
                        // If the Work Order Service Task has any associated files, delete them
                        {
                            using (var serviceContext = new Xrm(service))
                            {
                                // Retrieve all the files that are associated with the Work Order Service Task
                                var allFiles = serviceContext.ts_FileSet.ToList();
                                var workOrderServiceTaskFiles = allFiles.Where(f => f.ts_formintegrationid != null && f.ts_formintegrationid.Replace("WOST ", "").Trim() == preImageEntity.Attributes["msdyn_name"].ToString()).ToList();

                                if (workOrderServiceTaskFiles != null)
                                {
                                    foreach (var file in workOrderServiceTaskFiles)
                                    {
                                        service.Delete(file.LogicalName, file.Id);
                                    }
                                }
                            }
                        }
                    }
                }

                catch (Exception ex)
                {
                    tracingService.Trace("PreOperationmsdyn_workorderservicetaskDelete Plugin: {0}", ex.ToString());
                    throw;
                }
            }
        }
    }
}
