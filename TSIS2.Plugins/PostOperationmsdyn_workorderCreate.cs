using Microsoft.Xrm.Sdk.Query;
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
    "msdyn_workorder",
    StageEnum.PostOperation,
    ExecutionModeEnum.Synchronous,
    "",
    "TSIS2.Plugins.PostOperationmsdyn_workorderCreate Plugin",
    1,
    IsolationModeEnum.Sandbox,
    Image1Name = "PostImage", Image1Type = ImageTypeEnum.PostImage, Image1Attributes = "",
    Description = "Happens after the Work Order has been created")]

    public class PostOperationmsdyn_workorderCreate : PluginBase
    {
        public PostOperationmsdyn_workorderCreate(string unsecure, string secure)
            : base(typeof(PostOperationmsdyn_workorderCreate))
        {

        }
        /// <summary>
        /// Main entry point for he business logic that the plug-in is to execute.
        /// </summary>
        /// <param name="localContext">The <see cref="LocalPluginContext"/> which contains the
        /// <see cref="IPluginExecutionContext"/>,
        /// <see cref="IOrganizationService"/>
        /// and <see cref="ITracingService"/>
        /// </param>
        /// <remarks>
        /// For improved performance, Microsoft Dynamics 365 caches plug-in instances.
        /// The plug-in's Execute method should be written to be stateless as the constructor
        /// is not called for every invocation of the plug-in. Also, multiple system threads
        /// could execute the plug-in at the same time. All per invocation state information
        /// is stored in the context. This means that you should not use global variables in plug-ins.
        /// </remarks>
        protected override void ExecuteCrmPlugin(LocalPluginContext localContext)
        {
            if (localContext == null)
            {
                throw new InvalidPluginExecutionException("localContext");
            }

            IPluginExecutionContext context = localContext.PluginExecutionContext;
            Entity target = (Entity)context.InputParameters["Target"];
            localContext.Trace("Entering ExecuteCrmPlugin method.");

            try
            {
                if (target.Attributes.Contains("ovs_operationtypeid") || target.Attributes.Contains("ts_region"))
                {
                    string workOrderId = target.Id.ToString();
                    string ownerName = "";

                    using (var serviceContext = new Xrm(localContext.OrganizationService))
                    {
                        localContext.Trace("Determine if the region is set to International.");
                        var selectedRegion = target.Attributes["ts_region"] as EntityReference;

                        if (selectedRegion != null && selectedRegion.Id.Equals( new Guid("3bf0fa88-150f-eb11-a813-000d3af3a7a7")))
                        {
                            localContext.Trace("Setting business owner to International.");
                            target.Attributes["ts_businessowner"] = "AvSec International";

                            localContext.Trace("Perform the update to the Work Order.");
                            IOrganizationService service = localContext.OrganizationService;
                            
                            service.Update(target);

                            return;
                        }
                        else
                        {
                            localContext.Trace("Selected region is not International, checking operation type.");
                            // find out what business owns the Work Order
                            string fetchXML = $@"
                                <fetch xmlns:generator='MarkMpn.SQL4CDS'>
                                    <entity name='msdyn_workorder'>
                                    <link-entity name='ovs_operation' to='ovs_operationid' from='ovs_operationid' alias='ovs_operation' link-type='inner'>
                                    <link-entity name='ovs_operationtype' to='ovs_operationtypeid' from='ovs_operationtypeid' alias='ovs_operationtype' link-type='inner'>
                                    <link-entity name='team' to='owningteam' from='teamid' alias='team' link-type='inner'>
                                    <attribute name='name' alias='OwnerName' />
                                    </link-entity>
                                    </link-entity>
                                    </link-entity>
                                    <filter>
                                    <condition attribute='msdyn_workorderid' operator='eq' value='{workOrderId}' />
                                    </filter>
                                    </entity>
                                </fetch>
                            ";

                            EntityCollection businessNameCollection = localContext.OrganizationService.RetrieveMultiple(new FetchExpression(fetchXML));

                            if (businessNameCollection.Entities.Count == 0)
                            {
                                localContext.Trace("No business owner found for work order ID. Exit out if no results.");
                                return;
                            }

                            foreach (Entity workOrder in businessNameCollection.Entities)
                            {
                                if (workOrder["OwnerName"] is AliasedValue aliasedValue)
                                {
                                    localContext.Trace("Cast the AliasedValue to string (or the appropriate type).");
                                    ownerName = aliasedValue.Value as string;
                                }

                                localContext.Trace("Set the Business Owner Label.");
                                workOrder["ts_businessowner"] = ownerName;

                                localContext.Trace("Perform the update to the Work Order.");
                                IOrganizationService service = localContext.OrganizationService;
                                service.Update(workOrder);
                            }
                        }
                    }
                }
            }
            catch (Exception e)
            {
                localContext.TraceWithContext("Exception: {0}", e.Message);
                throw new InvalidPluginExecutionException("PostOperationmsdyn_workorderCreate failed.", e);
            }
        }
    }
}
