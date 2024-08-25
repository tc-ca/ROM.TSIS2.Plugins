using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TSIS2.Plugins
{
    class retrieveSearchHtmlTablePreferenceMapping
    {
        public enum dataConversionTypeCode { MoneyToDecimal, LeaseTerm, Lookup, None };
        public ConditionOperator Operator { get; private set; }
        public string mappedField { get; private set; }
        public string mappedEntity { get; private set; }
        public dataConversionTypeCode dataConversionType { get; private set; }


        public retrieveSearchHtmlTablePreferenceMapping(ConditionOperator Operator, string mappedField, string mappedEntity, dataConversionTypeCode dataConversionType = dataConversionTypeCode.None)
        {
            this.Operator = Operator;
            this.mappedField = mappedField;
            this.mappedEntity = mappedEntity;
            this.dataConversionType = dataConversionType;
        }
    }
}
