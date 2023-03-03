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
    /// <summary>
    /// PreOperationincidentUpdate Plugin.
    /// </summary>    
    public class PreOperationmsdyn_workorderRetrieveMultiple : IPlugin
    {
        private IOrganizationService service;
        public void Execute(IServiceProvider serviceProvider)
        {
            // Obtain the tracing service
            ITracingService tracingService =
            (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)
                serviceProvider.GetService(typeof(IPluginExecutionContext));

            FetchExpression objFetchExpression = (FetchExpression)context.InputParameters["Query"];
            XDocument fetchXmlDoc = XDocument.Parse(objFetchExpression.Query);
            var entityElement = fetchXmlDoc.Descendants("entity").FirstOrDefault();
            var entityName = entityElement.Attributes("name").FirstOrDefault().Value;

            // The InputParameters collection contains all the data passed in the message request.
            if (context.InputParameters.Contains("Query") &&
                context.InputParameters["Query"] is FetchExpression)
            {
                // Obtain the organization service reference which you will need for
                // web service calls.
                IOrganizationServiceFactory serviceFactory =
                    (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                service = serviceFactory.CreateOrganizationService(context.UserId);

                try
                {

                    var filterElements = entityElement.Descendants("filter");
                    var condition = from c in filterElements.Descendants("condition") //bogus filter in "Active Work Order" to identify it
                                    where c.Attribute("attribute").Value.Equals("msdyn_workordersummary")
                                    select c;

                    if (condition.Count() > 0)
                    {
                        Guid userId = context.InitiatingUserId;
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
                    throw new InvalidPluginExecutionException(e.Message);
                }

            }
        }
    }
}