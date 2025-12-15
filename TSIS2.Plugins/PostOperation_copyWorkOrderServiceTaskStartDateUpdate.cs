using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;

namespace TSIS2.Plugins
{
    [CrmPluginRegistration(
        MessageNameEnum.Update,
        "ts_workorderservicetaskworkspace",
        StageEnum.PostOperation,
        ExecutionModeEnum.Synchronous,
        "ts_name,ts_tasktype,ts_workorder,statecode,ts_fromoffline,ownerid,ts_location,ts_flightnumber,ts_origin,ts_destination,ts_flightcategory,ts_flighttype,ts_reportdetails,ts_scheduledtime,ts_actualtime,ts_paxonboard,ts_paxboarded,ts_cbonboard,ts_cbloaded,ts_aircraftmark,ts_aircraftmanufacturer,ts_aircraftmodel,ts_aircraftmodelother,ts_brandname,ts_passengerservices,ts_rampservices,ts_cargoservices,ts_cateringservices,ts_groomingservices,ts_securitysearchservices,ts_accesscontrolsecurityservices,ts_othersecurityservices,ts_workorderservicetaskstartdate,ts_questionnaireresponse,ts_questionnairedefinition,ts_mandatory,ts_percentcomplete,ts_aocoperation,ts_aocstakeholder,ts_aocoperationtype,ts_aocsite,ts_accesscontrol,ts_workorderservicetaskenddate,statuscode",
        "PostOperation.ts_workorderservicetaskworkspace.CopyStartDateToWorkOrderServiceTaskOnUpdate",
        1,
        IsolationModeEnum.Sandbox,
        Image1Name = "PreImage", Image1Type = ImageTypeEnum.PreImage, Image1Attributes = "ts_workorderservicetask",
        Description = "Copies changed fields to the related msdyn_workorderservicetask record on update.")]
    public class PostOperation_CopyStartDateToTaskOnUpdate : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            var context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            var serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            var service = serviceFactory.CreateOrganizationService(context.UserId);
            var tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            tracingService.Trace("Plugin execution started: PostOperation_CopyStartDateToTaskOnUpdate");

            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity target)
                {
                    tracingService.Trace("Target entity logical name: {0}", target.LogicalName);

                    if (!context.PreEntityImages.Contains("PreImage") || context.PreEntityImages["PreImage"] == null)
                    {
                        tracingService.Trace("PreImage is not available. Exiting plugin.");
                        return;
                    }

                    Entity preImage = context.PreEntityImages["PreImage"];
                    EntityReference workOrderTaskRef = preImage.GetAttributeValue<EntityReference>("ts_workorderservicetask");

                    if (workOrderTaskRef == null)
                    {
                        tracingService.Trace("ts_workorderservicetask reference is null. Exiting plugin.");
                        return;
                    }

                    tracingService.Trace("Updating msdyn_workorderservicetask Id: {0}", workOrderTaskRef.Id);

                    Entity updateTask = new Entity(workOrderTaskRef.LogicalName, workOrderTaskRef.Id);
                    bool anyFieldChanged = false;

                    // Helper for direct 1-to-1 field copies
                    Action<string, string> copyField = (sourceField, destField) =>
                    {
                        if (target.Contains(sourceField))
                        {
                            updateTask[destField] = target[sourceField];
                            tracingService.Trace($"Copied '{sourceField}' to '{destField}'.");
                            anyFieldChanged = true;
                        }
                    };

                    // --- Core Fields ---
                    copyField("ts_name", "msdyn_name");
                    copyField("ts_tasktype", "msdyn_tasktype");
                    copyField("ts_workorder", "msdyn_workorder");
                    copyField("ts_workorderservicetaskstartdate", "ts_servicetaskstartdate");
                    copyField("ts_workorderservicetaskenddate", "ts_servicetaskenddate");
                    copyField("ts_percentcomplete", "msdyn_percentcomplete");
                    copyField("statecode", "statecode");
                    copyField("ts_mandatory", "ts_mandatory");
                    copyField("ts_fromoffline", "ts_fromoffline");
                    copyField("ownerid", "ownerid");

                    // --- Questionnaire Fields ---
                    copyField("ts_questionnairedefinition", "ovs_questionnairedefinition");
                    copyField("ts_questionnaireresponse", "ovs_questionnaireresponse");
                    copyField("ts_accesscontrol", "ts_accesscontrol");

                    // --- Oversight & Flight Fields ---
                    copyField("ts_location", "ts_location");
                    copyField("ts_flightnumber", "ts_flightnumber");
                    copyField("ts_origin", "ts_origin");
                    copyField("ts_destination", "ts_destination");
                    copyField("ts_flightcategory", "ts_flightcategory");
                    copyField("ts_flighttype", "ts_flighttype");
                    copyField("ts_reportdetails", "ts_reportdetails");
                    copyField("ts_scheduledtime", "ts_scheduledtime");
                    copyField("ts_actualtime", "ts_actualtime");

                    // --- Passenger & Cargo Fields ---
                    copyField("ts_paxonboard", "ts_paxonboard");
                    copyField("ts_paxboarded", "ts_paxboarded");
                    copyField("ts_cbonboard", "ts_cbonboard");
                    copyField("ts_cbloaded", "ts_cbloaded");

                    // --- Aircraft Fields ---
                    copyField("ts_aircraftmark", "ts_aircraftmark");
                    copyField("ts_aircraftmanufacturer", "ts_aircraftmanufacturer");
                    copyField("ts_aircraftmodel", "ts_aircraftmodel");
                    copyField("ts_aircraftmodelother", "ts_aircraftmodelother");
                    copyField("ts_brandname", "ts_brandname");

                    // --- AOC Fields ---
                    copyField("ts_aocoperation", "ts_aocoperation");
                    copyField("ts_aocstakeholder", "ts_aocstakeholder");
                    copyField("ts_aocoperationtype", "ts_aocoperationtype");
                    copyField("ts_aocsite", "ts_aocsite");

                    // --- Service Provider Fields ---
                    copyField("ts_passengerservices", "ts_passengerservices");
                    copyField("ts_rampservices", "ts_rampservices");
                    copyField("ts_cargoservices", "ts_cargoservices");
                    copyField("ts_cateringservices", "ts_cateringservices");
                    copyField("ts_groomingservices", "ts_groomingservices");
                    copyField("ts_securitysearchservices", "ts_securitysearchservices");
                    copyField("ts_accesscontrolsecurityservices", "ts_accesscontrolsecurityservices");
                    copyField("ts_othersecurityservices", "ts_othersecurityservices");

                    // --- Fields with Special Logic ---
                    if (target.Contains("statuscode"))
                    {
                        var statusCode = target.GetAttributeValue<OptionSetValue>("statuscode");
                        int? mappedStatusCode = null;
                        if (statusCode != null)
                        {
                            switch (statusCode.Value)
                            {
                                case 1: mappedStatusCode = 1; break; // Active -> Active
                                case 741130001: mappedStatusCode = 918640002; break; // Complete -> Completed
                                case 741130002: mappedStatusCode = 918640004; break; // In Progress -> In Progress
                                case 741130003: mappedStatusCode = 918640005; break; // New -> New
                                default: mappedStatusCode = 1; break; // Default to Active
                            }
                        }
                        if (mappedStatusCode.HasValue)
                        {
                            updateTask["statuscode"] = new OptionSetValue(mappedStatusCode.Value);
                            tracingService.Trace("statuscode changed. New mapped value: {0}", mappedStatusCode.Value);
                            anyFieldChanged = true;
                        }
                    }

                    if (anyFieldChanged)
                    {
                        service.Update(updateTask);
                        tracingService.Trace("Updated msdyn_workorderservicetask with changed fields.");
                    }
                    else
                    {
                        tracingService.Trace("No relevant fields changed. No update performed.");
                    }
                }
                else
                {
                    tracingService.Trace("No target entity found. Exiting plugin.");
                }
            }
            catch (Exception ex)
            {
                tracingService.Trace("An error occurred: {0}", ex.ToString());
                throw;
            }
        }
    }
}