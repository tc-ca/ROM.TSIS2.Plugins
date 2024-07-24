using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace TSIS2.Plugins
{
    [CrmPluginRegistration(
    MessageNameEnum.Update,
    "msdyn_workorder",
    StageEnum.PostOperation,
    ExecutionModeEnum.Synchronous,
    "",
    "TSIS2.Plugins.PostOperationmsdyn_workoderRetrieveBusinessUnit Plugin",
    1,
    IsolationModeEnum.Sandbox,
    Description = "Filter the Active Work Orders view ")]
    public class PostOperationmsdyn_workoderRetrieveBusinessUnit : PluginBase
    {
        public PostOperationmsdyn_workoderRetrieveBusinessUnit(string unsecure, string secure)
            : base(typeof(PostOperationmsdyn_workoderRetrieveBusinessUnit))
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

            if (true)
            {
                IPluginExecutionContext context = localContext.PluginExecutionContext;
                Entity target = (Entity)context.InputParameters["Target"];



                // Check if the entity has already been processed
 

   
                try
                {
                                               
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
                        <condition attribute='msdyn_workorderid' operator='eq' value='909b920b-22c7-ec11-a7b6-000d3a0c7991' />
                        </filter>
                        </entity>
                    </fetch>
                    ";
                    EntityCollection businessNameCollection = localContext.OrganizationService.RetrieveMultiple(new FetchExpression(fetchXML));
                    string ownerName = "";

                    if (businessNameCollection.Entities.Count > 0)
                    {
                        // Get the first business entity
                        Entity businessEntity = businessNameCollection.Entities[0];

                        // Retrieve the OwnerName value, which is an AliasedValue
                        if (businessEntity.Contains("OwnerName") && businessEntity["OwnerName"] is AliasedValue aliasedValue)
                        {
                            // Cast the AliasedValue to string (or the appropriate type)
                            ownerName = aliasedValue.Value as string;

                            // Set the target attribute to the retrieved owner name
                            if (ownerName != null)
                            {
                                target.Attributes["ts_BusinessOwner"] = ownerName;
                            }
                        }
                    }
                    

                    //Entity updateEntity = new Entity(target.LogicalName, target.Id);
                    //updateEntity["ts_businessowner"] = target.Attributes["ts_businessowner"];


                    //localContext.OrganizationService.Update(new ts_File
                    //{
                      //  Id = file.Id,
                        //ts_msdyn_workorder = selectedWorkOrder.ToEntityReference(),
                        //ts_Incident = selectedWorkOrder.msdyn_ServiceRequest
                    //});


                    // Update the target entity to reflect changes in CRM
                    //localContext.OrganizationService.Update(new msdyn_workorder { Id = target.Id, ts_BusinessOwner = ownerName }

                        //);

                }
                catch (Exception e)
                {
                    throw new InvalidPluginExecutionException(e.Message);
                }

            }


        }
    }
}
