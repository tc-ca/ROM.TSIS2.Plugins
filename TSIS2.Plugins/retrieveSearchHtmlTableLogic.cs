using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace TSIS2.Plugins
{
    public static class retrieveSearchHtmlTableLogic
    {
        public static bool testMode = false;
        

        public static Tuple<string, string> intiateSearchLogicExecute(Entity targetCase, IOrganizationService service)
        {
            Dictionary<string, retrieveSearchHtmlTablePreferenceMapping> tripPreferenceDic = setTripPreferenceDictionary();
            var retTuple = searchMatchingTrips(tripPreferenceDic, targetCase, service);

            return retrieveSearchHtmlTableLogicDataTableHelper.ConvertEntityCollectionToHtmlDataTable(retTuple.Item1, retTuple.Item2, service);
        }

        public static Tuple<string, string> intiateSearchLogicExecute(Guid PrimaryEntityId, string PrimaryEntityName, IOrganizationService service)
        {
            Dictionary<string, retrieveSearchHtmlTablePreferenceMapping> tripPreferenceDic = setTripPreferenceDictionary();



            Entity targetCase = service.Retrieve(PrimaryEntityName, PrimaryEntityId, new ColumnSet(tripPreferenceDic.Keys.ToArray()));
            if (testMode)
            {
                //"Test Cancel" case under "My Cases" list
                targetCase["inctrk_tripid"] = "1";
                targetCase["inctrk_pllastnametrip"] = "";
                targetCase["inctrk_travelername"] = "z"; //z or e
            }

            var retTuple = searchMatchingTrips(tripPreferenceDic, targetCase, service);

            //EntityCollection filteredTrips = retrieveFilteredTrips(tripCollection, tripPreferenceDic, targetCase, service);
            
            return retrieveSearchHtmlTableLogicDataTableHelper.ConvertEntityCollectionToHtmlDataTable(retTuple.Item1, retTuple.Item2, service);
        }


        private static Tuple<EntityCollection, bool> searchMatchingTrips(Dictionary<string, retrieveSearchHtmlTablePreferenceMapping> unitPreferenceDic, Entity targetLead, IOrganizationService service)
        {
            var retTuple = buildSearch(unitPreferenceDic, targetLead);

            EntityCollection tripCollection = new EntityCollection();
            if (retTuple.Item2)
            {
                tripCollection = service.RetrieveMultiple(retTuple.Item1);
            }

            return new Tuple<EntityCollection, bool>(tripCollection, retTuple.Item2);
        }

        private static Tuple<QueryExpression, bool> buildSearch(Dictionary<string, retrieveSearchHtmlTablePreferenceMapping> tripPreferenceDic, Entity targetCase)
        {
            QueryExpression qe = new QueryExpression();
            qe.EntityName = "inctrk_trip";
            qe.ColumnSet = new ColumnSet(true);

            bool hasSearchCondition = false;

            FilterExpression feTrip = new FilterExpression();
            feTrip.FilterOperator = LogicalOperator.And;

            LinkEntity travelerLe = new LinkEntity();
            travelerLe.LinkToEntityName = "inctrk_traveler";
            travelerLe.LinkFromEntityName = "inctrk_trip";
            travelerLe.LinkToAttributeName = "inctrk_tripid";
            travelerLe.LinkFromAttributeName = "inctrk_tripid";
            travelerLe.JoinOperator = JoinOperator.Inner;
            travelerLe.EntityAlias = "traveler";

            bool needSearchTraveler = false;
            foreach (string key in tripPreferenceDic.Keys)
            {
                if (targetCase.Attributes.Contains(key))
                {
                    hasSearchCondition = true;
                    Object value = convertValueDataType(targetCase[key], tripPreferenceDic[key].dataConversionType);

                    ConditionExpression ce = new ConditionExpression();
                    ce.AttributeName = tripPreferenceDic[key].mappedField;
                    ce.Operator = tripPreferenceDic[key].Operator;

                    if (ce.Operator == ConditionOperator.Like)
                    {
                        ce.Values.Add("%" + value + "%");
                    }
                    else
                    {
                        ce.Values.Add(value);
                    }

                    if (tripPreferenceDic[key].mappedEntity == "inctrk_trip")
                    {
                        feTrip.AddCondition(ce);
                    }
                    else if (tripPreferenceDic[key].mappedEntity == "inctrk_traveler")
                    {
                        //inctrk_firstname, inctrk_lastname, inctrk_middlename
                        /*
                        FilterExpression feTraveler = new FilterExpression();
                        feTraveler.FilterOperator = LogicalOperator.Or;
                        feTraveler.AddCondition(ce);

                        ConditionExpression ce2 = new ConditionExpression();
                        ce2.AttributeName = "inctrk_lastname";
                        ce2.Operator = tripPreferenceDic[key].Operator;
                        ce2.Values.Add("%" + value + "%");
                        feTraveler.AddCondition(ce2);

                        ConditionExpression ce3 = new ConditionExpression();
                        ce3.AttributeName = "inctrk_middlename";
                        ce3.Operator = tripPreferenceDic[key].Operator;
                        ce3.Values.Add("%" + value + "%");
                        feTraveler.AddCondition(ce3);

                        travelerLe.LinkCriteria.AddFilter(feTraveler);
                        */
                        travelerLe.LinkCriteria.AddCondition(ce);
                        needSearchTraveler = true;
                    }
                }
            }

            if (needSearchTraveler)
            {
                qe.LinkEntities.Add(travelerLe);
            }

            if (feTrip.Conditions.Count > 0)
            {
                //add date condition
                DateTime last30Days = DateTime.Today.AddDays(-30);
                ConditionExpression ce = new ConditionExpression();
                ce.AttributeName = "inctrk_tripdepartdate";
                ce.Operator = ConditionOperator.GreaterThan;
                ce.Values.Add(last30Days);
                feTrip.AddCondition(ce);



                qe.Criteria.AddFilter(feTrip);

                qe.Orders.Add(new Microsoft.Xrm.Sdk.Query.OrderExpression("inctrk_sourcetripid", Microsoft.Xrm.Sdk.Query.OrderType.Ascending));
                //qe.Orders.Add(new Microsoft.Xrm.Sdk.Query.OrderExpression("lastname", Microsoft.Xrm.Sdk.Query.OrderType.Ascending));
            }

            return new Tuple<QueryExpression, bool>(qe, hasSearchCondition);
        }

        private static object convertValueDataType(object value, retrieveSearchHtmlTablePreferenceMapping.dataConversionTypeCode dataConversionType)
        {
            switch (dataConversionType)
            {
                case retrieveSearchHtmlTablePreferenceMapping.dataConversionTypeCode.MoneyToDecimal:
                    return ((Money)value).Value;
                case retrieveSearchHtmlTablePreferenceMapping.dataConversionTypeCode.LeaseTerm:
                    //todo:  find some way a little more elegant to handle this, it is a challenge because lease term option set text values do not align with lease terms inside of the unit rate entity.
                    if (((OptionSetValue)value).Value == 272970000)//12 MO
                    {
                        return "12";
                    }
                    else if (((OptionSetValue)value).Value == 272970001)//24 MO
                    {
                        return "24";
                    }
                    else
                    {
                        throw new InvalidPluginExecutionException("Unhandled Term:  Invalid Term on lead record please see your system administrator");
                    }
                case retrieveSearchHtmlTablePreferenceMapping.dataConversionTypeCode.Lookup:
                    return ((EntityReference)value).Id;
                default:
                    return value;
            }

        }
        private static Dictionary<string, retrieveSearchHtmlTablePreferenceMapping> setCaseNoteDictionary()
        {
            Dictionary<string, retrieveSearchHtmlTablePreferenceMapping> theDic = new Dictionary<string, retrieveSearchHtmlTablePreferenceMapping>();

            theDic.Add("inctrk_casenote", new retrieveSearchHtmlTablePreferenceMapping(ConditionOperator.Like, "inctrk_casenote", "inctrk_casenote"));

            return theDic;
        }
        private static Dictionary<string, retrieveSearchHtmlTablePreferenceMapping> setTripPreferenceDictionary()
        {
            Dictionary<string, retrieveSearchHtmlTablePreferenceMapping> tripPreferenceDic = new Dictionary<string, retrieveSearchHtmlTablePreferenceMapping>();

            tripPreferenceDic.Add("inctrk_tripid", new retrieveSearchHtmlTablePreferenceMapping(ConditionOperator.Like, "inctrk_sourcetripid", "inctrk_trip"));

            tripPreferenceDic.Add("inctrk_productioncode", new retrieveSearchHtmlTablePreferenceMapping(ConditionOperator.Like, "inctrk_primarytripno", "inctrk_trip"));

            tripPreferenceDic.Add("inctrk_groupno", new retrieveSearchHtmlTablePreferenceMapping(ConditionOperator.Like, "inctrk_groupnumber", "inctrk_trip"));

            tripPreferenceDic.Add("inctrk_groupnametrip", new retrieveSearchHtmlTablePreferenceMapping(ConditionOperator.Like, "inctrk_group_schoolname", "inctrk_trip"));

            tripPreferenceDic.Add("inctrk_pllastnametrip", new retrieveSearchHtmlTablePreferenceMapping(ConditionOperator.Like, "inctrk_programleaderlastname_ontour", "inctrk_trip"));

            tripPreferenceDic.Add("inctrk_travelername", new retrieveSearchHtmlTablePreferenceMapping(ConditionOperator.Like, "inctrk_lastname", "inctrk_traveler"));

            return tripPreferenceDic;
        }
    }
}
