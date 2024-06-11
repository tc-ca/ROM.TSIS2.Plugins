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
                    if (target.Attributes.Contains("accountid"))
                    {
                        Guid accountId = target.GetAttributeValue<Guid>("accountid");
                        String NewName = target.GetAttributeValue<String>("ovs_accountnameenglish");

                        using (var serviceContext = new Xrm(localContext.OrganizationService))
                        {
                            string fetchXml = $@"
                                <fetch>
                                <entity name='ovs_operation'>
                                  <attribute name='ovs_name' />
                                  <attribute name='ovs_operationid' />
                                  <attribute name='ovs_operationtypeid' />
                                  <attribute name='ts_site' />
                                  <filter>
                                    <condition attribute='ts_stakeholder' operator='eq' value='{accountId.ToString()}' />
                                  </filter>
                                 </entity>
                                 </fetch>";

                             EntityCollection operations = localContext.OrganizationService.RetrieveMultiple(new FetchExpression(fetchXml));



                            foreach (Entity operation in operations.Entities)
                            {

                                //ts_SharePointFile myOperationServiceTaskSharePointFile = null;

                                string ovsName = operation.GetAttributeValue<string>("ovs_name");

                                EntityReference ovsOperationTypeIdRef = operation.GetAttributeValue<EntityReference>("ovs_operationtypeid");
                                EntityReference tsSiteIdRef = operation.GetAttributeValue<EntityReference>("ts_site");


                                Guid ovsOperationId = operation.Id;
                                Guid ovsSiteId = operation.Id;
                                 

                                string operationTypeName = ovsOperationTypeIdRef.Name;
                                string siteName = tsSiteIdRef.Name;
                                operation["ts_stakeholdername"] = "test";

                                //string newOvsName = $"{NewName} | {operationTypeName} | {siteName}";
                                //operation["ovs_name"] = newOvsName;
                                IOrganizationService service = localContext.OrganizationService;
                                service.Update(operation);




                            }


                        }

                    }
                    // get the Operations that belong to the Account - retrieve them by the Account ID

                        // Loop over the retrieved Operations
                        // Update the Operation name
            }
        }
        catch (Exception e) { throw new InvalidPluginExecutionException(e.Message); }
        }
    }
}
