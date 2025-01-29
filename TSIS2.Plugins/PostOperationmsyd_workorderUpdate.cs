using Microsoft.Crm.Sdk.Messages;
using Microsoft.FSharp.Core;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Services.Description;
using System.Xml.Linq;

namespace TSIS2.Plugins
{
    [CrmPluginRegistration(
    MessageNameEnum.Update,
    "msdyn_workorder",
    StageEnum.PostOperation,
    ExecutionModeEnum.Synchronous,
    "",
    "TSIS2.Plugins.PostOperationmsdyn_workorderUpdate Plugin",
    1,
    IsolationModeEnum.Sandbox,
    Image1Name = "PostImage", Image1Type = ImageTypeEnum.PostImage, Image1Attributes = "",
    Image2Name = "PreImage", Image2Type = ImageTypeEnum.PreImage, Image2Attributes = "",
    Description = "Happens after the Work Order has been updated")]
    public class PostOperationmsdyn_workorderUpdate : PluginBase
    {
        private readonly string postImageAlias = "PostImage";
        public PostOperationmsdyn_workorderUpdate(string unsecure, string secure)
            : base(typeof(PostOperationmsdyn_workorderUpdate))
        {

            //if (secure != null && !secure.Equals(string.Empty))
            //{

            //}
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
            ITracingService tracingService = localContext.TracingService;
            Entity target = (Entity)context.InputParameters["Target"];

            tracingService.Trace("Obtain the images for the entity.");
            Entity preImageEntity = (context.PreEntityImages != null && context.PreEntityImages.Contains("PreImage")) ? context.PreEntityImages["PreImage"] : null;
            Entity postImageEntity = (context.PostEntityImages != null && context.PostEntityImages.Contains("PostImage")) ? context.PostEntityImages["PostImage"] : null;

            // Guid for the International Region
            string internationalGuid = "3bf0fa88-150f-eb11-a813-000d3af3a7a7";

            try
            {
                tracingService.Trace("Check if the region has changed.");
                // we don't (target.Attributes.Contains("ts_region")) because when an update happens to a Work Order, it run's through here several times
                {
                    IOrganizationService service = localContext.OrganizationService;

                    // Log the system username and Work Order at the start
                    var systemUser = service.Retrieve("systemuser", context.InitiatingUserId, new ColumnSet("fullname"));
                    var WOST = service.Retrieve("msdyn_workorder", context.PrimaryEntityId, new ColumnSet("msdyn_name"));
                    tracingService.Trace("Plugin executed by user: {0}", systemUser.GetAttributeValue<string>("fullname"));
                    tracingService.Trace("Work Order GUID: {0}", context.PrimaryEntityId);
                    tracingService.Trace("Work Order Name: {0}", WOST.GetAttributeValue<string>("msdyn_name"));

                    if (preImageEntity.Contains("ts_region") && postImageEntity.Contains("ts_region"))
                    {
                        var preRegion = preImageEntity["ts_region"] as EntityReference;
                        var postRegion = postImageEntity["ts_region"] as EntityReference;

                        if (preRegion != null && postRegion != null)
                        {
                            if (preRegion.Id.Equals(new Guid(internationalGuid)) && postRegion.Id.Equals(new Guid(internationalGuid)))
                            {
                                tracingService.Trace("If the Region is already set to International, exit the Plugin. Prevents infinite loop.");
                                return;
                            }
                            else if (postRegion.Id.Equals(new Guid(internationalGuid)))
                            {
                                tracingService.Trace("If the Region is set to International, set the owner label to International.");
                                target.Attributes["ts_businessowner"] = "AvSec International";

                                tracingService.Trace("Perform the update to the Work Order.");
                                service.Update(target);
                                return;
                            }
                        }
                    }

                    if (preImageEntity.Contains("ts_trip") && !postImageEntity.Contains("ts_trip") && !target.Contains("ts_ignoreupdate"))
                    {
                        //if trip got removed PBI-372064, remove WO from Trip Inspection -> ts_tripinspection
                        tracingService.Trace("If trip got removed, then remove WO from Trip Inspection.");
                        var tripId = preImageEntity.GetAttributeValue<EntityReference>("ts_trip").Id;
                        localContext.Trace("Trip removed: " + tripId.ToString());

                        var qETripInspection = new QueryExpression("ts_tripinspection");
                        qETripInspection.ColumnSet.AddColumns("ts_trip");

                        qETripInspection.Criteria.AddCondition("ts_trip", ConditionOperator.Equal, tripId);
                        qETripInspection.Criteria.AddCondition("ts_inspection", ConditionOperator.Equal, preImageEntity.Id);
                        EntityCollection returnCol = service.RetrieveMultiple(qETripInspection);
                        if (returnCol.Entities.Count > 0)
                        {
                            localContext.Trace("Trip removed: return  " + returnCol.Entities.Count + " - " + returnCol.Entities[0].Id.ToString());
                            service.Delete(returnCol.Entities[0].LogicalName, returnCol.Entities[0].Id);
                        }

                    }
                    else if (!preImageEntity.Contains("ts_trip") && postImageEntity.Contains("ts_trip") && !target.Contains("ts_ignoreupdate"))
                    {
                        //if trip got removed PBI-372064, remove WO from Trip Inspection -> ts_tripinspection
                        tracingService.Trace("If trip got removed, then remove WO from ts_tripinspection.");
                        var tripId = postImageEntity.GetAttributeValue<EntityReference>("ts_trip").Id;
                        localContext.Trace("Trip added: " + tripId.ToString());

                        var qETripInspection = new QueryExpression("ts_tripinspection");
                        qETripInspection.ColumnSet.AddColumns("ts_trip");

                        qETripInspection.Criteria.AddCondition("ts_trip", ConditionOperator.Equal, tripId);
                        qETripInspection.Criteria.AddCondition("ts_inspection", ConditionOperator.Equal, preImageEntity.Id);
                        EntityCollection returnCol = service.RetrieveMultiple(qETripInspection);
                        if (returnCol.Entities.Count == 0)
                        {
                            localContext.Trace("Trip add: return  " + returnCol.Entities.Count );
                            Entity newEnt= new Entity("ts_tripinspection");
                            newEnt["ts_trip"] = new EntityReference("ts_trip", tripId);
                            newEnt["ts_inspection"] = new EntityReference("msdyn_workorder", preImageEntity.Id);
                            service.Create(newEnt);
                        }

                    }
                }

                tracingService.Trace("Check if an operation type was updated.");
                // we don't (target.Attributes.Contains("ovs_operationtype")) because when an update happens to a Work Order, it run's through here several times
                {
                    string workOrderId = target.Id.ToString();
                    string ownerName = "";

                    using (var serviceContext = new Xrm(localContext.OrganizationService))
                    {
                        tracingService.Trace("Determine what business owns the Work Order.");
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
                            tracingService.Trace("Exit out if there are no results.");
                            return;
                        }

                        foreach (Entity workOrder in businessNameCollection.Entities)
                        {
                            if (workOrder["OwnerName"] is AliasedValue aliasedValue)
                            {
                                tracingService.Trace("Cast the AliasedValue to string (or the appropriate type).");
                                ownerName = aliasedValue.Value as string;

                                if (preImageEntity.Contains("ts_businessowner"))
                                {
                                    tracingService.Trace("If ts_businessowner is already set to the value of ownerName, exit the Plugin.");
                                    // we do this to prevent an infinite loop from happening
                                    var preBusinessOwner = preImageEntity.Attributes["ts_businessowner"];

                                    if (preBusinessOwner != null && preBusinessOwner.ToString() == ownerName)
                                    {
                                        return;
                                    }
                                }

                            }
                            tracingService.Trace("Set the Business Owner Label.");
                            workOrder["ts_businessowner"] = ownerName;

                            tracingService.Trace("Perform the update to the Work Order.");
                            IOrganizationService service = localContext.OrganizationService;
                            service.Update(workOrder);
                            return;
                        }
                    }
                }
            }
            catch (Exception e)
            {
                tracingService.Trace("Exception occurred: {0}");
                throw new InvalidPluginExecutionException(e.Message);
            }
        }
    }
}