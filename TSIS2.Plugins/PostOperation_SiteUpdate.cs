using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace TSIS2.Plugins
{
    [CrmPluginRegistration(
    MessageNameEnum.Update,
    "msdyn_functionallocation",
    StageEnum.PostOperation,
    ExecutionModeEnum.Synchronous,
    "",
    "TSIS2.Plugins.PostOperation_SiteUpdate Plugin",
    1,
    IsolationModeEnum.Sandbox,
    Image1Name = "PostImage", Image1Type = ImageTypeEnum.PostImage, Image1Attributes = "",
    Description = "If a Site name is updated, the related operation names are being updated")]
    /// <summary>
    /// PostOperationmsdyn_workorderservicetaskUpdate Plugin.
    /// </summary>  
    public class PostOperation_SiteUpdate : PluginBase
    {
        private readonly string postImageAlias = "PostImage";

        /// <summary>
        /// Initializes a new instance of the <see cref="PostOperationmsdyn_workorderservicetaskUpdate"/> class.
        /// </summary>
        /// <param name="unsecure">Contains public (unsecured) configuration information.</param>
        /// <param name="secure">Contains non-public (secured) configuration information. 
        /// When using Microsoft Dynamics 365 for Outlook with Offline Access, 
        /// the secure string is not passed to a plug-in that executes while the client is offline.</param>
        public PostOperation_SiteUpdate(string unsecure, string secure)
            : base(typeof(PostOperation_SiteUpdate))
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

            localContext.Trace("Tracking Service Started.");

            IPluginExecutionContext context = localContext.PluginExecutionContext;
            Entity target = (Entity)context.InputParameters["Target"];
            Entity postImageEntity = (context.PostEntityImages != null && context.PostEntityImages.Contains(this.postImageAlias)) ? context.PostEntityImages[this.postImageAlias] : null;


            try
            {
                // run this code only if these fields are updated
                if (target.Attributes.Contains("msdyn_name") ||
                    target.Attributes.Contains("ts_functionallocationnameenglish") ||
                    target.Attributes.Contains("ts_functionallocationnamefrench"))
                {
                    localContext.Trace("Site name fields have been updated.");

                    Guid SiteId = target.GetAttributeValue<Guid>("msdyn_functionallocationid");
                    String NewName = target.GetAttributeValue<String>("ts_functionallocationnameenglish");
                    String NewNameFrench = target.GetAttributeValue<String>("ts_functionallocationnamefrench");

                    localContext.Trace("Site ID: {0}", SiteId);
                    localContext.Trace("New English Name: {0}", NewName);
                    localContext.Trace("New French Name: {0}", NewNameFrench);

                    using (var serviceContext = new Xrm(localContext.OrganizationService))
                    {
                        localContext.Trace("Retrieve Operations that belong to the Site.");
                        // get the Operations that belong to the Site - retrieve them by the Site ID
                        string fetchXml = $@"
                            <fetch>
                            <entity name='ovs_operation'>
                                <attribute name='ts_operationnameenglish' />
                                <attribute name='ts_operationnamefrench' />
                                <attribute name='ovs_name' />
                                <attribute name='ovs_operationid' />
                                <attribute name='ovs_operationtypeid' />
                                <attribute name='ts_site' />
                                <attribute name='ownerid' />
                                <filter>
                                  <condition attribute='ts_site' operator='eq' value='{SiteId.ToString()}' />
                                </filter>
                               </entity>
                             </fetch>";

                        EntityCollection operations = localContext.OrganizationService.RetrieveMultiple(new FetchExpression(fetchXml));
                        localContext.Trace("Found {0} operations associated with Site.", operations.Entities.Count);
                        if (operations.Entities.Count == 0)
                        {
                            localContext.Trace("No operations found, exiting.");
                            return;
                        }
                        Entity firstOperation = operations.Entities[0];
                        string firstOperationName = firstOperation.GetAttributeValue<string>("ts_operationnameenglish") ?? firstOperation.GetAttributeValue<string>("ovs_name") ?? "N/A";

                        // Get the owner team reference from the operation
                        EntityReference ownerTeamRef = firstOperation.GetAttributeValue<EntityReference>("ownerid");
                        if (ownerTeamRef == null)
                        {
                            localContext.Trace("Operation ID: {0}, Name: {1} has no owner, exiting.", firstOperation.Id, firstOperationName);
                            return;
                        }

                        localContext.Trace("Checking if operation ID: {0}, Name: {1} owner belongs to ISSO.", firstOperation.Id, firstOperationName);
                        ////Check if the record belongs to ISSO - if not don't run the code
                        if (!OrganizationConfig.IsOwnedByISSO(localContext.OrganizationService, ownerTeamRef, tracingService))
                        {
                            localContext.Trace("Operation ID: {0}, Name: {1} owner does not belong to ISSO, exiting.", firstOperation.Id, firstOperationName);
                            return;
                        }

                        localContext.Trace("Operation ID: {0}, Name: {1} owner belongs to ISSO, proceeding with updates.", firstOperation.Id, firstOperationName);


                        // Loop over the retrieved Operations
                        // Update the Operation name

                        localContext.Trace("Processing {0} operations for name updates.", operations.Entities.Count);
                        // Go through each related Operation
                        foreach (Entity operation in operations.Entities)
                        {
                            // Get the english name of the operation
                            string originalOperationName = operation.GetAttributeValue<string>("ts_operationnameenglish");
                            if (originalOperationName == null)
                            {
                                originalOperationName = operation.GetAttributeValue<string>("ovs_name");
                            }
                            localContext.Trace("Processing operation ID: {0}, Name: {1}", operation.Id, originalOperationName ?? "N/A");
                            string updatedOperationName = "";
                            string updatedOperationNameFrench = "";
                            
                            if (!(originalOperationName.StartsWith("OP-")))
                            {
                                localContext.Trace("Operation name does not start with 'OP-', updating name parts.");
                                string[] parts = originalOperationName.Split('|');

                                // Logic to update Operation Name goes here
                                // Note: Set the updated Operation Name in 'updatedOperationName'

                                for (int i = 0; i < parts.Length; i++)
                                {
                                    parts[i] = parts[i].Trim();
                                }
                                parts[2] = NewName;
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
                                parts_French[2] = NewNameFrench;
                                updatedOperationNameFrench = string.Join(" | ", parts_French);
                                // Update the Operation Name
                                operation["ovs_name"] = updatedOperationName;
                                operation["ts_operationnameenglish"] = updatedOperationName;
                                operation["ts_operationnamefrench"] = updatedOperationNameFrench;

                                localContext.Trace("Updated operation ID: {0} name (English): {1}, (French): {2}", operation.Id, updatedOperationName, updatedOperationNameFrench);

                                // Perform the update to the Operation
                                IOrganizationService service = localContext.OrganizationService;
                                localContext.Trace("Updating operation ID: {0}", operation.Id);
                                service.Update(operation);
                                localContext.Trace("Operation ID: {0} updated successfully.", operation.Id);

                            }
                            else
                            {
                                localContext.Trace("Operation ID: {0}, Name: {1} starts with 'OP-', skipping update.", operation.Id, originalOperationName ?? "N/A");
                            }



                        }
                        localContext.Trace("Finished processing all operations.");
                    }
                }
                else
                {
                    localContext.Trace("Site name fields were not updated, skipping processing.");
                }
            }
            catch (Exception e)
            {
                localContext.TraceWithContext("Exception: {0}", e.Message);
                throw new InvalidPluginExecutionException("PostOperation_SiteUpdate failed.", e);
            }
            
        }
            
    }
}
