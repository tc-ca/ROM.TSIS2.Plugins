using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TSIS2.Plugins
{
    class retrieveSearchHtmlTableDataMappingPreferences
    {
        public string attributeName { get; private set; }
        public bool isToBeHtmlEncoded { get; private set; }
        public bool isVisibleInGrid { get; private set; }
        public bool isSelectionColumn { get; private set; }
        public bool isSpecialColumn { get; private set; }
        public HyperlinkHelper hyperlinkHelper { get; private set; }

        public string frName { get; private set; }
        public retrieveSearchHtmlTableDataMappingPreferences(string attributeName, bool isToBeHtmlEncoded, bool isVisibleInGrid, bool isSelectionColumn, bool isTermColumn, HyperlinkHelper hyperlinkHelper, string frName = "")
        {
            this.attributeName = attributeName;
            this.isToBeHtmlEncoded = isToBeHtmlEncoded;
            this.isVisibleInGrid = isVisibleInGrid;
            this.isSelectionColumn = isSelectionColumn;
            this.isSpecialColumn = isTermColumn;
            this.hyperlinkHelper = hyperlinkHelper;
            this.frName = frName;
        }

        public class HyperlinkHelper
        {
            public string EntityName { get; private set; }
            public string idAttribute { get; private set; }

            public HyperlinkHelper(string EntityName, string idAttribute)
            {
                this.EntityName = EntityName;
                this.idAttribute = idAttribute;
            }
        }
    }
}
