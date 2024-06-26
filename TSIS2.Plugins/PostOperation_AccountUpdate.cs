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

            IPluginExecutionContext context = localContext.PluginExecutionContext;
            Entity target = (Entity)context.InputParameters["Target"];
            Entity postImageEntity = (context.PostEntityImages != null && context.PostEntityImages.Contains(this.postImageAlias)) ? context.PostEntityImages[this.postImageAlias] : null;


            try
            {
                if (target.LogicalName.Equals(Account.EntityLogicalName))
                {
                    // run this code only if these fields are updated
                    if (target.Attributes.Contains("name") || 
                        target.Attributes.Contains("ovs_accountnameenglish") || 
                        target.Attributes.Contains("ovs_accountnamefrench"))
                    {
                        
                        Guid accountId = target.GetAttributeValue<Guid>("accountid");
                        String NewName = target.GetAttributeValue<String>("ovs_accountnameenglish");
                        String NewNameFrench = target.GetAttributeValue<String>("ovs_accountnamefrench");

                        using (var serviceContext = new Xrm(localContext.OrganizationService))
                        {
                            // get the Operations that belong to the Account - retrieve them by the Account ID
                            string fetchXml = $@"
                                <fetch>
                                <entity name='ovs_operation'>
                                  <attribute name='ts_operationnameenglish' />
                                  <attribute name='ts_operationnamefrench' />
                                  <attribute name='ovs_operationid' />
                                  <attribute name='ovs_operationtypeid' />
                                  <attribute name='ts_site' />
                                  <link-entity name='team' to='ownerid' from='teamid' alias='team' link-type='inner'>
		                              <attribute name='name' alias='owner' />
		                          </link-entity>
                                  <filter>
                                    <condition attribute='ts_stakeholder' operator='eq' value='{accountId.ToString()}' />
                                  </filter>
                                 </entity>
                                 </fetch>";

                            EntityCollection operations = localContext.OrganizationService.RetrieveMultiple(new FetchExpression(fetchXml));
                            if (operations.Entities.Count == 0)
                            {
                                return;
                            }
                            Entity firstOperation = operations.Entities[0];

                            // Get the AliasedValue for the attribute "owner"
                            AliasedValue ownerAliasedValue = firstOperation.GetAttributeValue<AliasedValue>("owner");

                            // Extract the actual value from the AliasedValue
                            string owner = ownerAliasedValue.Value as string;

                            ////Check if the record belongs to ISSO - if not don't run the code
                            if (!(owner.StartsWith("Intermodal")))
                            {
                                return;
                            }


                            // Loop over the retrieved Operations
                            // Update the Operation name

                            // Go through each related Operation
                            foreach (Entity operation in operations.Entities)
                            {
                                // Get the english name of the operation
                                string originalOperationName = operation.GetAttributeValue<string>("ts_operationnameenglish");
                                string updatedOperationName = "";
                                string updatedOperationNameFrench = "";
                                
                                if (originalOperationName != null)
                                {
                                    string[] parts = originalOperationName.Split('|');
                                    for (int i = 0; i < parts.Length; i++)
                                    {
                                        parts[i] = parts[i].Trim();
                                    }
                                    parts[0] = NewName;
                                    updatedOperationName = string.Join("|", parts);

                                    // Get the french name of the operation
                                    string originalFrenchOperationName = operation.GetAttributeValue<string>("ts_operationnamefrench");
                                    string[] parts_French = originalFrenchOperationName.Split('|');
                                    for (int i = 0; i < parts.Length; i++)
                                    {
                                        parts_French[i] = parts_French[i].Trim();
                                    }
                                    parts_French[0] = NewNameFrench;
                                    updatedOperationNameFrench = string.Join("|", parts_French);
                                    // Update the Operation Name
                                    operation["ovs_name"] = updatedOperationName;
                                    operation["ts_operationnameenglish"] = updatedOperationName;
                                    operation["ts_operationnamefrench"] = updatedOperationNameFrench;





                                    // Perform the update to the Operation
                                    IOrganizationService service = localContext.OrganizationService;
                                    service.Update(operation);

                                }
                                // Logic to update Operation Name goes here
                                // Note: Set the updated Operation Name in 'updatedOperationName'




                            }
                        }
                    }
                    
            }
        }
        catch (Exception e) { throw new InvalidPluginExecutionException(e.Message); }
        }
    }
}
