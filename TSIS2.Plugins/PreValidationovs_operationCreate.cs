using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TSIS2.Plugins
{
    [CrmPluginRegistration(
    MessageNameEnum.Create,
    "ovs_operation",
    StageEnum.PreValidation,
    ExecutionModeEnum.Synchronous,
    "",
    "TSIS2.Plugins.PreValidationovs_operationCreate Plugin",
    1,
    IsolationModeEnum.Sandbox,
    Description = "Sets the Operation Properties")]
    /// <summary>
    /// PreValidationovs_operationCreate Plugin.
    /// </summary>
    public class PreValidationovs_operationCreate : IPlugin
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
                    if (target.LogicalName.Equals(ovs_operation.EntityLogicalName))
                    {
                        // Cast the target to the expected entity
                        ovs_operation operation = target.ToEntity<ovs_operation>();

                        using (var serviceContext = new Xrm(service))
                        {
                            if (operation.OwnerId != null && operation.OwnerId.LogicalName != null && operation.OwnerId.LogicalName == "systemuser")
                            {
                                var owningUser = serviceContext.SystemUserSet.FirstOrDefault(user => user.Id == operation.OwnerId.Id);
                                if (owningUser != null && owningUser.BusinessUnitId != null)
                                {
                                    var businessUnit = serviceContext.BusinessUnitSet.FirstOrDefault(bu => bu.BusinessUnitId == owningUser.BusinessUnitId.Id);
                                    if (businessUnit != null)
                                    {
                                        if (businessUnit.Name == "Aviation Security Directorate")
                                        {
                                            target["ownerid"] = new EntityReference(Team.EntityLogicalName, new Guid("6db920a0-baa3-eb11-b1ac-000d3ae8b98c")); //Aviation Security Directorate Team
                                        }
                                        else if (businessUnit.Name == "Aviation Security Directorate - Domestic")
                                        {
                                            target["ownerid"] = new EntityReference(Team.EntityLogicalName, new Guid("8544831b-bead-eb11-8236-000d3ae8b866")); //Aviation Security Directorate - Domestic Team
                                        }
                                        else if (businessUnit.Name == "Intermodal Surface Security Oversight (ISSO)")
                                        {
                                            target["ownerid"] = new EntityReference(Team.EntityLogicalName, new Guid("8544831b-bead-eb11-8236-000d3ae8b866")); //Intermodal Surface Security Oversight (ISSO) Team
                                        }
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

