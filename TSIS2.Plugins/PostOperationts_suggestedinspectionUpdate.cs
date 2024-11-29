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
    "ts_suggestedinspection",
    StageEnum.PostOperation,
    ExecutionModeEnum.Synchronous,
    "",
    "TSIS2.Plugins.PostOperationts_suggestedinspectionUpdate Plugin",
    1,
    IsolationModeEnum.Sandbox,
    Image1Name = "PostImage", Image1Type = ImageTypeEnum.PostImage, Image1Attributes = "",
    Image2Name = "PreImage", Image2Type = ImageTypeEnum.PreImage, Image2Attributes = "",
    Description = "Happens after the Suggested Inspection has been updated")]
    public class PostOperationts_suggestedinspectionUpdate : PluginBase
    {
        private readonly string postImageAlias = "PostImage";
        public PostOperationts_suggestedinspectionUpdate(string unsecure, string secure)
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
            Entity target = (Entity)context.InputParameters["Target"];

            // Obtain the images for the entity
            Entity preImageEntity = (context.PreEntityImages != null && context.PreEntityImages.Contains("PreImage")) ? context.PreEntityImages["PreImage"] : null;
            Entity postImageEntity = (context.PostEntityImages != null && context.PostEntityImages.Contains("PostImage")) ? context.PostEntityImages["PostImage"] : null;

            try
            {
                {
                    IOrganizationService service = localContext.OrganizationService;
                    //if trip added
                    if (!preImageEntity.Contains("ts_trip") && postImageEntity.Contains("ts_trip"))
                    {
                        var theTrip = postImageEntity["ts_trip"] as EntityReference;
                        var tripEnt = service.Retrieve("ts_trip", theTrip.Id, new ColumnSet("ts_estimatedcost", "ts_estimatedtraveltime", "ts_plannedfiscalquarter"));

                        localContext.Trace("Trip added   ");
                        bool needUpdate = false;
                        Entity updEnt = new Entity("ts_suggestedinspection", postImageEntity.Id);
                        if (tripEnt.Contains("ts_estimatedcost") && (!postImageEntity.Contains("ts_estimatedcost") || postImageEntity["ts_estimatedcost"] != tripEnt["ts_estimatedcost"]))
                        {
                            updEnt["ts_estimatedcost"] = tripEnt["ts_estimatedcost"];
                            needUpdate = true;
                        }

                        if (tripEnt.Contains("ts_estimatedtraveltime") && (!postImageEntity.Contains("ts_estimatedtraveltime") || postImageEntity["ts_estimatedtraveltime"] != tripEnt["ts_estimatedtraveltime"]))
                        {
                            updEnt["ts_estimatedtraveltime"] = tripEnt["ts_estimatedtraveltime"];
                            needUpdate = true;
                        }
                        else if (!tripEnt.Contains("ts_estimatedtraveltime"))
                        {
                            //wait for business team decide logic if removed the value from Suggested Inspection
                        }

                        if (tripEnt.Contains("ts_plannedfiscalquarter"))
                        {
                            var labelQuarter = tripEnt.GetAttributeValue<EntityReference>("ts_plannedfiscalquarter").Name.ToLower();

                            var quarterArray = new string[] { "q1", "q2", "q3", "q4" };
                            foreach ( var quarter in quarterArray )
                            {
                                var fieldName = "ts_" + quarter;
                                if (labelQuarter == quarter)
                                {
                                    updEnt[fieldName] = 1;
                                }
                                else {
                                    updEnt[fieldName] = 0;
                                }
                            }
                            needUpdate = true;
                        }
                        
                        if (needUpdate)
                        {
                            localContext.Trace("Update SuggestedInspection..");
                            service.Update(updEnt);
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