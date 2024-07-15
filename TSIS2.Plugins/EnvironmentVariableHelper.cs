using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace TSIS2.Plugins
{
    public class EnvironmentVariableHelper
    {
        public static string GetEnvironmentVariableValue(IOrganizationService service, string schemaName)
        {
            // Create a query to retrieve the environment variable definition
            var query = new QueryExpression("environmentvariabledefinition")
            {
                //The value we want to receive
                ColumnSet = new ColumnSet("defaultvalue")
            };
            //filters the result by the chema name
            query.Criteria.AddCondition("schemaname", ConditionOperator.Equal, schemaName);

            // get the environment variable definition
            var definition = service.RetrieveMultiple(query).Entities.FirstOrDefault();

            if (definition != null)
            {
                var query2 = new QueryExpression("environmentvariablevalue")
                {
                    ColumnSet = new ColumnSet("value")
                };
                //filter by environment variable id
                query2.Criteria.AddCondition("environmentvariabledefinitionid", ConditionOperator.Equal, definition.Id);

                //get the environment variable value
                var value = service.RetrieveMultiple(query2).Entities.FirstOrDefault();

                if (value != null)
                {
                    return value.GetAttributeValue<string>("value");
                }

                return definition.GetAttributeValue<string>("defaultvalue");
            }
            return null;
        }
    }
}
