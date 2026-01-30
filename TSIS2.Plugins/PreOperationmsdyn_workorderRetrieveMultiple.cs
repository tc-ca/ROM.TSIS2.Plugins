using System;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System.Xml.Linq;

namespace TSIS2.Plugins
{
    [CrmPluginRegistration(
    MessageNameEnum.RetrieveMultiple,
    "msdyn_workorder",
    StageEnum.PreOperation,
    ExecutionModeEnum.Synchronous,
    "",
    "TSIS2.Plugins.PreOperationworkorderRetrieveMultiple Plugin",
    1,
    IsolationModeEnum.Sandbox,
    Description = "Filter the Active Work Orders view ")]
    public class PreOperationmsdyn_workorderRetrieveMultiple : PluginBase
    {
        public PreOperationmsdyn_workorderRetrieveMultiple(string unsecure, string secure)
            : base(typeof(PreOperationmsdyn_workorderRetrieveMultiple))
        {
        }

        protected override void ExecuteCrmPlugin(LocalPluginContext localContext)
        {
            if (localContext == null)
            {
                throw new InvalidPluginExecutionException("localContext");
            }

            IPluginExecutionContext context = localContext.PluginExecutionContext;
            IOrganizationService service = localContext.OrganizationService;

            if (context.InputParameters.Contains("Query") &&
                context.InputParameters["Query"] is FetchExpression)
            {
                try
                {
                    FetchExpression objFetchExpression = (FetchExpression)context.InputParameters["Query"];
                    XDocument fetchXmlDoc = XDocument.Parse(objFetchExpression.Query);
                    var entityElement = fetchXmlDoc.Descendants("entity").FirstOrDefault();
                    var entityName = entityElement.Attributes("name").FirstOrDefault().Value;

                    var filterElements = entityElement.Descendants("filter");
                    var condition = from c in filterElements.Descendants("condition") //bogus filter in "Active Work Order" to identify it
                                    where c.Attribute("attribute").Value.Equals("msdyn_workordersummary")
                                    select c;

                    Guid userId = context.InitiatingUserId;

                    if (condition.Count() > 0 && !IsUserSystemAdministrator(context.UserId, service))
                    {
                        userId = context.InitiatingUserId;
                        Boolean dualInspector = false;

                        //Get User BU
                        Guid userBusinessUnitId = ((EntityReference)(service.Retrieve("systemuser", userId, new ColumnSet("businessunitid"))).Attributes["businessunitid"]).Id;
                        string userBusinessUnitName = (string)service.Retrieve("businessunit", userBusinessUnitId, new ColumnSet("name")).Attributes["name"];

                        if (userBusinessUnitName != "Transport Canada")
                        {
                            String teamFetchXML = @"<fetch distinct='false' mapping='logical' returntotalrecordcount='true' no-lock='false'>
                                              <entity name='team'>
                                                <attribute name='name' />
                                                <filter type='and'>
                                                  <condition attribute='teamtype' operator='ne' value='1' />
                                                </filter>
                                                <order attribute='name' descending='false' />
                                                <link-entity name='teammembership' intersect='true' visible='false' to='teamid' from='teamid'>
                                                  <link-entity name='systemuser' from='systemuserid' to='systemuserid' alias='bb'>
                                                    <filter type='and'>
                                                      <condition attribute='systemuserid' operator='eq' uitype='systemuser' value='" + userId + @"' />
                                                    </filter>
                                                  </link-entity>
                                                </link-entity>
                                              </entity>
                                            </fetch>";

                            EntityCollection retrievedTeams = service.RetrieveMultiple(new FetchExpression(teamFetchXML));

                            bool inISSOInspectorTeam = false;
                            bool inAvSecInspectorTeam = false;

                            //Check user teams
                            foreach (var team in retrievedTeams.Entities)
                            {
                                if ((string)team.Attributes["name"] == "ISSO Inspectors")
                                {
                                    inISSOInspectorTeam = true;
                                }
                                if ((string)team.Attributes["name"] == "Aviation Security - International - Inspectors" || (string)team.Attributes["name"] == "Aviation Security - Domestic - Inspectors")
                                {
                                    inAvSecInspectorTeam = true;
                                }
                            }

                            dualInspector = inISSOInspectorTeam && inAvSecInspectorTeam;

                            if (!dualInspector)
                            {
                                //Add condition to only show WO with activity type BU equals to user BU
                                entityElement.Add(
                                        new XElement("link-entity",
                                            new XAttribute("name", "msdyn_incidenttype"),
                                            new XAttribute("from", "msdyn_incidenttypeid"),
                                            new XAttribute("to", "msdyn_primaryincidenttype"),
                                            new XAttribute("link-type", "inner"),
                                                 new XElement("filter",
                                                    new XElement("condition",
                                                        new XAttribute("attribute", "owningbusinessunit"),
                                                        new XAttribute("operator", "eq"),
                                                        new XAttribute("value", userBusinessUnitId.ToString())
                                                    )
                                                )
                                            )
                                        );
                                objFetchExpression.Query = fetchXmlDoc.ToString();
                            }
                        }
                    }
                }
                catch (Exception e)
                {
                    localContext.TraceWithContext("Exception: {0}", e.Message);     
                    throw new InvalidPluginExecutionException("PreOperationmsdyn_workorderRetrieveMultiple failed.", e);
                }
            }
        }

        private bool IsUserSystemAdministrator(Guid userId, IOrganizationService service)
        {
            string fetchXml = $@"<fetch>
                                    <entity name='systemuser'>
                                        <attribute name='systemuserid' />
                                        <filter>
                                            <condition attribute='systemuserid' operator='eq' value='{userId}' />
                                        </filter>
                                        <link-entity name='systemuserroles' from='systemuserid' to='systemuserid'>
                                            <link-entity name='role' from='roleid' to='roleid'>
                                                <filter>
                                                    <condition attribute='name' operator='eq' value='System Administrator' />
                                                </filter>
                                            </link-entity>
                                        </link-entity>
                                    </entity>
                                </fetch>";

            EntityCollection result = service.RetrieveMultiple(new FetchExpression(fetchXml));
            return result.Entities.Count > 0;
        }
    }
}