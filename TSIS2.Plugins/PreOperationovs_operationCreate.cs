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
    StageEnum.PreOperation,
    ExecutionModeEnum.Synchronous,
    "",
    "TSIS2.Plugins.PreOperationovs_operationCreate Plugin",
    1,
    IsolationModeEnum.Sandbox,
    Description = "Sets the Operation Properties")]
    /// <summary>
    /// PreOperationovs_operationCreate Plugin.
    /// </summary>
    public class PreOperationovs_operationCreate : IPlugin
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

                        if (operation.ts_site == null || operation.ts_stakeholder == null || operation.ovs_OperationTypeId == null)
                        {
                            return;
                        }

                        using (var serviceContext = new Xrm(service))
                        {
                            msdyn_FunctionalLocation site = serviceContext.msdyn_FunctionalLocationSet.FirstOrDefault(fl => fl.Id == operation.ts_site.Id);
                            ovs_operationtype operationType = serviceContext.ovs_operationtypeSet.FirstOrDefault(ot => ot.Id == operation.ovs_OperationTypeId.Id);
                            Account stakeholder = serviceContext.AccountSet.FirstOrDefault(acc => acc.Id == operation.ts_stakeholder.Id);

                            if (site.ts_Country == null || stakeholder.ts_Country == null) return;

                            var operationTypeAirCarrierPassengerGuid = "8b614ef0-c651-eb11-a812-000d3af3ac0d";
                            var countryCanadaGuid = "208ef8a1-8e75-eb11-a812-000d3af3fac7";
                            var countryUSAGuid = "7c01709f-8e75-eb11-a812-000d3af3f6ab";

                            if (operationType.Id.ToString() == operationTypeAirCarrierPassengerGuid && site.ts_Country.Id.ToString() == countryCanadaGuid)
                            {
                                if (stakeholder.ts_Country.Id.ToString() == countryCanadaGuid)
                                {
                                    target["ts_domesticflights"] = true;
                                    target["ts_cateringandstores"] = true;
                                    target["ts_oss"] = true;
                                    target["ts_cargo"] = true;
                                    target["ts_unattendedaircraft"] = true;
                                }
                                else if (stakeholder.ts_Country.Id.ToString() == countryUSAGuid)
                                {
                                    target["ts_transborderflights"] = true;
                                    target["ts_cateringandstores"] = true;
                                    target["ts_oss"] = true;
                                    target["ts_cargo"] = true;
                                }
                                else
                                {
                                    target["ts_internationalflights"] = true;
                                    target["ts_cateringandstores"] = true;
                                    target["ts_oss"] = true;
                                    target["ts_cargo"] = true;
                                }
                            }
                            if (site.ts_Country.Id.ToString() != countryCanadaGuid)
                            {
                                target["ts_internationalprogramsbranchipb"] = true;
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

