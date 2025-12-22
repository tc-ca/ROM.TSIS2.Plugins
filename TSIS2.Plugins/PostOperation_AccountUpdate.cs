using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Activities.Statements;
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
    "account",
    StageEnum.PostOperation,
    ExecutionModeEnum.Synchronous,
    "",
    "TSIS2.Plugins.PostOperation_AccountUpdate Plugin",
    1,
    IsolationModeEnum.Sandbox,
    Image1Name = "PostImage", Image1Type = ImageTypeEnum.PostImage, Image1Attributes = "",
    Description = "If a stockholder is updated, the related operation names are being updated")]
    /// <summary>
    /// PostOperationmsdyn_workorderservicetaskUpdate Plugin.
    /// </summary>  
    public class PostOperation_AccountUpdate : PluginBase
    {
        private readonly string postImageAlias = "PostImage";

        /// <summary>
        /// Initializes a new instance of the <see cref="PostOperationmsdyn_workorderservicetaskUpdate"/> class.
        /// </summary>
        /// <param name="unsecure">Contains public (unsecured) configuration information.</param>
        /// <param name="secure">Contains non-public (secured) configuration information. 
        /// When using Microsoft Dynamics 365 for Outlook with Offline Access, 
        /// the secure string is not passed to a plug-in that executes while the client is offline.</param>
        public PostOperation_AccountUpdate(string unsecure, string secure)
            : base(typeof(PostOperation_AccountUpdate))
        {

            //if (secure != null &&!secure.Equals(string.Empty))
            //{

            //}
        }

        protected override void ExecuteCrmPlugin(LocalPluginContext localContext)
        {
            if (localContext == null)
            {
                throw new InvalidPluginExecutionException("localContext");
            }

            // Obtain the tracing service
            ITracingService tracingService = localContext.TracingService;

            tracingService.Trace("Tracking Service Started.");

            IPluginExecutionContext context = localContext.PluginExecutionContext;
            Entity target = (Entity)context.InputParameters["Target"];
            Entity postImageEntity = (context.PostEntityImages != null && context.PostEntityImages.Contains(this.postImageAlias)) ? context.PostEntityImages[this.postImageAlias] : null;


            try
            {
                if (target.LogicalName.Equals(Account.EntityLogicalName))
                {
                    tracingService.Trace("Target is Account entity.");
                    // run this code only if these fields are updated
                    if (target.Attributes.Contains("name") || 
                        target.Attributes.Contains("ovs_accountnameenglish") || 
                        target.Attributes.Contains("ovs_accountnamefrench"))
                    {
                        tracingService.Trace("Account name fields have been updated.");
                        
                        Guid accountId = target.GetAttributeValue<Guid>("accountid");
                        String NewName = target.GetAttributeValue<String>("ovs_accountnameenglish");
                        String NewNameFrench = target.GetAttributeValue<String>("ovs_accountnamefrench");

                        tracingService.Trace("Account ID: {0}", accountId);
                        tracingService.Trace("New English Name: {0}", NewName);
                        tracingService.Trace("New French Name: {0}", NewNameFrench);

                        using (var serviceContext = new Xrm(localContext.OrganizationService))
                        {
                            tracingService.Trace("Retrieve Operations that belong to the Account.");
                            // get the Operations that belong to the Account - retrieve them by the Account ID
                            string fetchXml = $@"
                                <fetch>
                                <entity name='ovs_operation'>
                                  <attribute name='ts_operationnameenglish' />
                                  <attribute name='ovs_name' />
                                  <attribute name='ts_operationnamefrench' />
                                  <attribute name='ovs_operationid' />
                                  <attribute name='ovs_operationtypeid' />
                                  <attribute name='ts_site' />
                                  <attribute name='ownerid' />
                                  <filter>
                                    <condition attribute='ts_stakeholder' operator='eq' value='{accountId.ToString()}' />
                                  </filter>
                                 </entity>
                                 </fetch>";

                            EntityCollection operations = localContext.OrganizationService.RetrieveMultiple(new FetchExpression(fetchXml));
                            tracingService.Trace("Found {0} operations associated with Account.", operations.Entities.Count);
                            if (operations.Entities.Count == 0)
                            {
                                tracingService.Trace("No operations found, exiting.");
                                return;
                            }
                            Entity firstOperation = operations.Entities[0];
                            string firstOperationName = firstOperation.GetAttributeValue<string>("ts_operationnameenglish") ?? firstOperation.GetAttributeValue<string>("ovs_name") ?? "N/A";

                            // Get the owner team reference from the operation
                            EntityReference ownerTeamRef = firstOperation.GetAttributeValue<EntityReference>("ownerid");
                            if (ownerTeamRef == null)
                            {
                                tracingService.Trace("Operation ID: {0}, Name: {1} has no owner, exiting.", firstOperation.Id, firstOperationName);
                                return;
                            }

                            tracingService.Trace("Checking if operation ID: {0}, Name: {1} owner belongs to ISSO.", firstOperation.Id, firstOperationName);
                            ////Check if the record belongs to ISSO - if not don't run the code
                            if (!OrganizationConfig.IsOwnedByISSO(localContext.OrganizationService, ownerTeamRef, tracingService))
                            {
                                tracingService.Trace("Operation ID: {0}, Name: {1} owner does not belong to ISSO, exiting.", firstOperation.Id, firstOperationName);
                                return;
                            }

                            tracingService.Trace("Operation ID: {0}, Name: {1} owner belongs to ISSO, proceeding with updates.", firstOperation.Id, firstOperationName);


                            // Loop over the retrieved Operations
                            // Update the Operation name

                            tracingService.Trace("Processing {0} operations for name updates.", operations.Entities.Count);
                            // Go through each related Operation
                            foreach (Entity operation in operations.Entities)
                            {
                                // Get the english name of the operation
                                string originalOperationName = operation.GetAttributeValue<string>("ts_operationnameenglish");
                                if(originalOperationName == null)
                                {
                                    originalOperationName = operation.GetAttributeValue<string>("ovs_name");
                                }
                                tracingService.Trace("Processing operation ID: {0}, Name: {1}", operation.Id, originalOperationName ?? "N/A");
                                string updatedOperationName = "";
                                string updatedOperationNameFrench = "";
                                
                                if (!(originalOperationName.StartsWith("OP-")))
                                {
                                    tracingService.Trace("Operation name does not start with 'OP-', updating name parts.");
                                    string[] parts = originalOperationName.Split('|');
                                    for (int i = 0; i < parts.Length; i++)
                                    {
                                        parts[i] = parts[i].Trim();
                                    }
                                    parts[0] = NewName;
                                    updatedOperationName = string.Join(" | ", parts);

                                    // Get the french name of the operation
                                    string originalFrenchOperationName = operation.GetAttributeValue<string>("ts_operationnamefrench");

                                    if (originalFrenchOperationName == null)
                                    {
                                        originalFrenchOperationName = operation.GetAttributeValue<string>("ovs_name");
                                    }

                                    string[] parts_French = originalFrenchOperationName.Split('|');
                                    for (int i = 0; i < parts_French.Length; i++)
                                    {
                                        parts_French[i] = parts_French[i].Trim();
                                    }
                                    parts_French[0] = NewNameFrench;
                                    updatedOperationNameFrench = string.Join(" | ", parts_French);
                                    // Update the Operation Name
                                    operation["ovs_name"] = updatedOperationName;
                                    operation["ts_operationnameenglish"] = updatedOperationName;
                                    operation["ts_operationnamefrench"] = updatedOperationNameFrench;





                                    tracingService.Trace("Updated operation ID: {0} name (English): {1}, (French): {2}", operation.Id, updatedOperationName, updatedOperationNameFrench);

                                    // Perform the update to the Operation
                                    IOrganizationService service = localContext.OrganizationService;
                                    tracingService.Trace("Updating operation ID: {0}", operation.Id);
                                    service.Update(operation);
                                    tracingService.Trace("Operation ID: {0} updated successfully.", operation.Id);

                                }
                                else
                                {
                                    tracingService.Trace("Operation ID: {0}, Name: {1} starts with 'OP-', skipping update.", operation.Id, originalOperationName ?? "N/A");
                                }
                                // Logic to update Operation Name goes here
                                // Note: Set the updated Operation Name in 'updatedOperationName'




                            }
                            tracingService.Trace("Finished processing all operations.");
                        }
                    }
                    else
                    {
                        tracingService.Trace("Account name fields were not updated, skipping processing.");
                    }
                }
                else
                {
                    tracingService.Trace("Target entity is not Account, skipping processing.");
                }
            }
            catch (Exception e)
            {
                tracingService.Trace("Error occurred: {0}", e.Message);
                throw new InvalidPluginExecutionException(e.Message);
            }
        }
    }
}
