using System;
using System.Linq;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;

namespace TSIS2.Plugins
{

    [CrmPluginRegistration(
        MessageNameEnum.Create,
        "ts_entityrisk",
        StageEnum.PostOperation,
        ExecutionModeEnum.Synchronous,
        "",
        "TSIS2.Plugins.PostOperationts_entityriskCreate Plugin",
        1,
        IsolationModeEnum.Sandbox,
        Description = "After an Entity Risk is created, setup the many-to-many relationship")]
    public class PostOperationts_entityriskCreate : PluginBase
    {
        public PostOperationts_entityriskCreate() : base(typeof(PostOperationts_entityriskCreate))
        {
        }

        protected override void ExecuteCrmPlugin(LocalPluginContext localContext)
        {
            var context = localContext.PluginExecutionContext;
            var service = localContext.OrganizationService;

            // The InputParameters collection contains all the data passed in the message request.
            if (context.InputParameters.Contains("Target") &&
                context.InputParameters["Target"] is Entity)
            {
                // Obtain the target entity from the input parameters.
                Entity target = (Entity)context.InputParameters["Target"];
                try
                {
                    if (target.LogicalName.Equals(ts_EntityRisk.EntityLogicalName))
                    {
                        ts_EntityRisk myEntityRisk = target.ToEntity<ts_EntityRisk>();

                        localContext.Trace("Entering the PostOperationts_entityriskCreate method.");
                        localContext.Trace("Check if ts_entityname and ts_entityid are populated.");

                        /*  
                         *  Check if ts_entityname and ts_entityid are populated.  If they are then create the many-to-many relationship record
                        **/
                        {
                            if (!String.IsNullOrWhiteSpace(myEntityRisk.ts_EntityID) &&
                                myEntityRisk.ts_EntityName.HasValue)
                            {
                                using (var serviceContext = new Xrm(service))
                                {
                                    var selectedTable = "";
                                    var relationshipName = "";
                                    localContext.Trace("Determine what table is selected");

                                    switch (myEntityRisk.ts_EntityName)
                                    {
                                        case ts_EntityRisk_ts_EntityName.ActivityType:
                                            selectedTable = "msdyn_incidenttypeid";
                                            relationshipName = "ts_EntityRisk_msdyn_incidenttype_msdyn_incidenttype";
                                            break;
                                        case ts_EntityRisk_ts_EntityName.Operation:
                                            selectedTable = "ovs_operationid";
                                            relationshipName = "ts_EntityRisk_ovs_operation_ovs_operation";
                                            break;
                                        case ts_EntityRisk_ts_EntityName.OperationType:
                                            selectedTable = "ovs_operationid";
                                            relationshipName = "ts_EntityRisk_ovs_operationtype_ovs_operationtype";
                                            break;
                                        case ts_EntityRisk_ts_EntityName.ProgramArea:
                                            selectedTable = "ts_programareaid";
                                            relationshipName = "ts_EntityRisk_ts_programarea_ts_programarea";
                                            break;
                                        case ts_EntityRisk_ts_EntityName.Site:
                                            selectedTable = "msdyn_functionallocationid";
                                            relationshipName = "ts_EntityRisk_msdyn_FunctionalLocation_msdyn_FunctionalLocation";
                                            break;
                                        case ts_EntityRisk_ts_EntityName.Stakeholder:
                                            selectedTable = "accountid";
                                            relationshipName = "ts_EntityRisk_Account_Account";
                                            break;
                                        default:
                                            break;
                                    }

                                    if (String.IsNullOrWhiteSpace(selectedTable))
                                    {
                                        localContext.Trace("selectedTable does not have a value");
                                        throw new InvalidPluginExecutionException(
                                        $"The selectedTable value is not set because the ts_EntityName '{myEntityRisk.ts_EntityName}' does not match any known cases.");
                                    }

                                    localContext.Trace("Create the Associated Request");
                                    AssociateRequest associateRequest = new AssociateRequest
                                    {
                                        Target = myEntityRisk.ToEntityReference(),
                                        RelatedEntities = new EntityReferenceCollection
                                        {
                                            new EntityReference(selectedTable, new Guid(myEntityRisk.ts_EntityID))
                                        },
                                        Relationship = new Relationship(relationshipName)
                                    };

                                    localContext.Trace("Create the many-to-many record");
                                    service.Execute(associateRequest);

                                    localContext.Trace("Exiting the PostOperationts_entityriskCreate method.");
                                }
                            }
                        }

                    }
                }
                catch (Exception e)
                {
                    localContext.Trace("PostOperationts_entityriskCreate Plugin: {0}", e);
                    throw new InvalidPluginExecutionException("PostOperationts_entityriskCreate failed.", e);
                }
            }
        }
    }
}