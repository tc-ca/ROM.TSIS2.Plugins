using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace TSIS2.Plugins
{
    public static class retrieveSearchHtmlTableLogicDataTableHelper
    {
        public static int resultRowsLimit = 50;
        public static EntityCollection searchMatchingRecords(string recordId, IOrganizationService service, ITracingService tracingService)
        {

            var opActivityId = new Guid(recordId);
            var opEnt = service.Retrieve("ts_operationactivity", opActivityId, new ColumnSet("ts_operation", "ts_activity"));

            var opId = Guid.Empty;
            var actTypeId = Guid.Empty;
            if (opEnt.Contains("ts_operation"))
            {
                opId = opEnt.GetAttributeValue<EntityReference>("ts_operation").Id;
            }
            if (opEnt.Contains("ts_activity"))
            {
                actTypeId = opEnt.GetAttributeValue<EntityReference>("ts_activity").Id;
            }

            tracingService.Trace("opId : " + opId.ToString());
            QueryExpression qe = new QueryExpression();
            qe.EntityName = "msdyn_workorder";
            qe.ColumnSet = new ColumnSet(true);

            FilterExpression feWO = new FilterExpression();
            feWO.FilterOperator = LogicalOperator.And;

            ConditionExpression ce = new ConditionExpression();
            ce.AttributeName = "ovs_operationid";
            ce.Operator = ConditionOperator.Equal;
            ce.Values.Add(opId);
            feWO.AddCondition(ce);
            
            ConditionExpression ce2 = new ConditionExpression();
            ce2.AttributeName = "msdyn_primaryincidenttype";
            ce2.Operator = ConditionOperator.Equal;
            ce2.Values.Add(actTypeId);
            feWO.AddCondition(ce2);

            qe.Criteria.AddFilter(feWO);

            qe.Orders.Add(new Microsoft.Xrm.Sdk.Query.OrderExpression("msdyn_name", Microsoft.Xrm.Sdk.Query.OrderType.Ascending));

            return service.RetrieveMultiple(qe); ;
        }


        public static Tuple<string, string> ConvertNotesCollectionToHtmlDataTable(EntityCollection ec, IOrganizationService service)
        {
            string html = "";
            string info = "";

            if (ec.Entities.Count > 0)
            {
                Dictionary<string, retrieveSearchHtmlTableDataMappingPreferences> dataTableMappings = setDataTableMapping();
                DataTable dt = createDataTable(ec, dataTableMappings, service);

                html = convertDataTableToHTml(dt, dataTableMappings);

                info = string.Format("<div> Found {0} notes.<br></div>", dt.Rows.Count);

            }
            else
            {
                html = "No related WorkOrder were found with current Operation Activity.";
            }

            return new Tuple<string, string>(html, info);
        }

        public static Tuple<string, string> ConvertEntityCollectionToHtmlDataTable(EntityCollection ec, bool hasSearchCondition, IOrganizationService service)
        {
            string html = "";
            string info = "";

            if (ec.Entities.Count > 0)
            {
                Dictionary<string, retrieveSearchHtmlTableDataMappingPreferences> dataTableMappings = setDataTableMapping();
                DataTable dt = createDataTable(ec, dataTableMappings, service);

                html = convertDataTableToHTml(dt, dataTableMappings);


                if (dt.Rows.Count >= resultRowsLimit)
                {
                    info = string.Format("<div> More than {0} Work Orders found. (Only show the first {1} WO).<br></div>", resultRowsLimit, dt.Rows.Count);
                }
                else
                {
                    info = string.Format("<div> Found {0} related Work Orders <br></div>", dt.Rows.Count);
                }

                //info = string.Format("<div> Found {0} trips with matching criteria.<br></div>", dt.Rows.Count);

            }
            else
            {
                html = "No Work Orders were found with current Operation Activity.";
            }


            return new Tuple<string, string>(html, info);
        }

        private static string convertDataTableToHTml(DataTable dt, Dictionary<string, retrieveSearchHtmlTableDataMappingPreferences> dataTableMappings)
        {
            StringBuilder strHTMLBuilder = new StringBuilder();

            strHTMLBuilder.Append("<table id='findTripUITable'>");

            strHTMLBuilder.Append("<thead><tr>");
            foreach (DataColumn column in dt.Columns)
            {
                if (dataTableMappings[column.ColumnName].isVisibleInGrid)
                {
                    strHTMLBuilder.Append("<th>");


                    strHTMLBuilder.Append("<div class='grid-header-text' style='color:#045999; opacity: 1;'>" + WebUtility.HtmlEncode(column.ColumnName) + "</div><div class='columnSeparator'>&nbsp;</div>");

                    strHTMLBuilder.Append("</th>");
                }
            }
            strHTMLBuilder.Append("</tr></thead>");

            strHTMLBuilder.Append("<tbody>");
            foreach (DataRow row in dt.Rows)
            {

                strHTMLBuilder.Append("<tr>");
                foreach (DataColumn column in dt.Columns)
                {
                    if (dataTableMappings[column.ColumnName].isVisibleInGrid)
                    {
                        strHTMLBuilder.Append("<td>");
                        if (dataTableMappings[column.ColumnName].isToBeHtmlEncoded)
                        {
                            strHTMLBuilder.Append(WebUtility.HtmlEncode(row[column.ColumnName].ToString()));
                        }
                        else
                        {
                            strHTMLBuilder.Append(row[column.ColumnName].ToString());
                        }
                        strHTMLBuilder.Append("</td>");
                    }

                }
                strHTMLBuilder.Append("</tr>");
            }
            strHTMLBuilder.Append("</tbody>");
            strHTMLBuilder.Append("</table>");

            string Htmltext = strHTMLBuilder.ToString();

            return Htmltext;
        }

        private static DataTable createDataTable(EntityCollection ec, Dictionary<string, retrieveSearchHtmlTableDataMappingPreferences> dataTableMappings, IOrganizationService service)
        {
            DataTable dt = new DataTable();
            createDataColumns(dt, dataTableMappings);

            int rowNumber = 1;

            foreach (Entity e in ec.Entities)
            {
                if (rowNumber > resultRowsLimit)
                {
                    break;
                }

                else if (e.Attributes.Contains("msdyn_name"))
                {

                    addRow(e, dt, dataTableMappings, rowNumber);
                    rowNumber++;
                }
            }

            return dt;
        }



        private static void updateDataRow(Entity e, DataRow dr, string specialColumnName, Dictionary<string, retrieveSearchHtmlTableDataMappingPreferences> dataTableMappings)
        {
            dr[specialColumnName] = getFormattedValue(e, dataTableMappings[specialColumnName].attributeName);
        }

        private static void addRow(Entity e, DataTable dt, Dictionary<string, retrieveSearchHtmlTableDataMappingPreferences> dataTableMappings, int rowNumber)
        {
            DataRow newRow = dt.NewRow();

            foreach (KeyValuePair<string, retrieveSearchHtmlTableDataMappingPreferences> kvp in dataTableMappings)
            {
                if (!kvp.Value.isSelectionColumn && !kvp.Value.isSpecialColumn)
                {
                    if (kvp.Value.hyperlinkHelper == null)
                    {
                        if (e.Attributes.Contains(kvp.Value.attributeName))
                        {
                            newRow[kvp.Key] = getFormattedValue(e, kvp.Value.attributeName);
                        }
                        else
                        {
                            newRow[kvp.Key] = "";
                        }
                    }
                    else
                    {
                        if (e.Attributes.Contains(kvp.Value.attributeName))
                        {
                            //string url = string.Format("|OrgUrl|/main.aspx?etn={0}&pagetype=entityrecord&id=%7B{1}%7d", kvp.Value.hyperlinkHelper.EntityName, getFormattedValue(e,kvp.Value.hyperlinkHelper.idAttribute));
                            string url = string.Format("/main.aspx?etn={0}&pagetype=entityrecord&id=%7B{1}%7d", kvp.Value.hyperlinkHelper.EntityName, getFormattedValue(e, kvp.Value.hyperlinkHelper.idAttribute));

                            string hyperLink = "<a href = '" + url + "' target='_blank'> " + getFormattedValue(e, kvp.Value.attributeName) + " </a>";
                            newRow[kvp.Key] = hyperLink;
                        }
                        else
                        {
                            newRow[kvp.Key] = "";
                        }
                    }
                }
                else if (kvp.Value.isSelectionColumn)
                {
                    if (kvp.Value.attributeName == "editcasenote")
                    {
                        var curId = e.Id.ToString();
                        newRow[kvp.Key] = "<button type='button' id='editId" + rowNumber + "' name='editName" + rowNumber + "' data-uid='" + curId + "' onclick='javascript:editNote(" + rowNumber + ")'>Edit</button>";
                    }
                    else
                    {
                        var curId = getFormattedValue(e, dataTableMappings["Trip ID"].attributeName);

                        newRow[kvp.Key] = "<input type='checkbox' id='selectId" + rowNumber + "' name='selectName" + rowNumber + "' value='" + curId + "' onchange='rowSelected(this)'>";

                    }
                }

            }

            dt.Rows.Add(newRow);
        }

        private static void createDataColumns(DataTable dt, Dictionary<string, retrieveSearchHtmlTableDataMappingPreferences> dataTableMappings)
        {
            foreach (KeyValuePair<string, retrieveSearchHtmlTableDataMappingPreferences> kvp in dataTableMappings)
            {
                dt.Columns.Add(kvp.Key);
            }
        }
        private static Dictionary<string, retrieveSearchHtmlTableDataMappingPreferences> setDataTableMapping()
        {
            Dictionary<string, retrieveSearchHtmlTableDataMappingPreferences> dataTableMapping = new Dictionary<string, retrieveSearchHtmlTableDataMappingPreferences>();

            dataTableMapping.Add("WO No.", new retrieveSearchHtmlTableDataMappingPreferences("msdyn_name", false, true, false, false, new retrieveSearchHtmlTableDataMappingPreferences.HyperlinkHelper("msdyn_workorder", "msdyn_workorderid")));

            dataTableMapping.Add("Stakeholder", new retrieveSearchHtmlTableDataMappingPreferences("msdyn_serviceaccount", true, true, false, false, null));
            dataTableMapping.Add("Category", new retrieveSearchHtmlTableDataMappingPreferences("ovs_rational", true, true, false, false, null));
            dataTableMapping.Add("Site", new retrieveSearchHtmlTableDataMappingPreferences("ts_site", true, true, false, false, null));
            dataTableMapping.Add("Activity Type", new retrieveSearchHtmlTableDataMappingPreferences("msdyn_primaryincidenttype", true, true, false, false, null));
            dataTableMapping.Add("Owner", new retrieveSearchHtmlTableDataMappingPreferences("ownerid", true, true, false, false, null));
            dataTableMapping.Add("Region", new retrieveSearchHtmlTableDataMappingPreferences("ts_region", true, true, false, false, null));

            dataTableMapping.Add("# of Findings", new retrieveSearchHtmlTableDataMappingPreferences("ts_numberoffindings", true, true, false, false, null));
            dataTableMapping.Add("Closed Date", new retrieveSearchHtmlTableDataMappingPreferences("msdyn_timeclosed", true, true, false, false, null));
            dataTableMapping.Add("Created On", new retrieveSearchHtmlTableDataMappingPreferences("createdon", true, true, false, false, null));
            return dataTableMapping;
        }

        private static string getFormattedValue(Entity entity, string attributeName)
        {
            KeyValuePair<string, object> kvp = new KeyValuePair<string, object>();
            KeyValuePair<string, string> formattedValue = new KeyValuePair<string, string>();
            string value = "";
            if (entity.FormattedValues.Keys.Contains(attributeName))
            {
                formattedValue = entity.FormattedValues.First(k => k.Key == attributeName);
                value = formattedValue.Value;
            }
            else if (entity.Attributes.Keys.Contains(attributeName))
            {
                kvp = entity.Attributes.First(k => k.Key == attributeName);

                Type t = kvp.Value.GetType();
                if (t.Name == "EntityReference")
                {
                    EntityReference er = (EntityReference)kvp.Value;
                    if (er.Name != null)
                        value = er.Name;
                }
                else if (t.Name == "AliasedValue")
                {
                    AliasedValue av = (AliasedValue)kvp.Value;
                    Type t2 = av.Value.GetType();
                    if (t2.Name == "EntityReference")
                    {
                        EntityReference er2 = (EntityReference)av.Value;
                        if (er2.Name != null)
                            value = er2.Name;
                    }
                    else
                        value = av.Value.ToString();
                }
                else if (t.Name == "OptionSetValue")
                {
                    OptionSetValue osv = (OptionSetValue)kvp.Value;
                    value = osv.Value.ToString();
                }
                else if (attributeName == "entityimage" && t.Name == "Byte[]")
                {
                    byte[] binaryData = (byte[])kvp.Value;
                    value = System.Convert.ToBase64String(binaryData, 0, binaryData.Length);
                }
                else
                {
                    value = kvp.Value.ToString();
                }
            }
            return processExtraValue(entity, attributeName, value);
        }

        //inctrk_accountmanagerteamdirectorlastname
        private static string processExtraValue(Entity entity, string attributeName, string curValue)
        {
            var retV = curValue;
            if (attributeName == "inctrk_accountmanagerteamdirectorlastname")
            {
                retV = entity.GetAttributeValue<string>("inctrk_accountmanagerteamdirectorfirstname") + ", " + curValue;
            }
            else if (attributeName == "inctrk_accountmanagerlastname")
            {
                retV = entity.GetAttributeValue<string>("inctrk_accountmanagerfirstname") + ", " + curValue;
            }
            else if (attributeName == "inctrk_programleaderlastname_ontour")
            {
                retV = entity.GetAttributeValue<string>("inctrk_programleaderfirstname_ontour") + ", " + curValue;
            }


            return retV;
        }
    }
}
