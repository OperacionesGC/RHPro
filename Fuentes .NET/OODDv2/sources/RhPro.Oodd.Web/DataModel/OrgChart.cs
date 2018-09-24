using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Xml.Linq;

namespace RhPro.Oodd.Web.DataModel
{
    public partial class OrgChart
    {
        public int orgChartCode { get; set; }

        public int returnCode { get; set; }

        public string errorMessage { get; set; }

        public XElement tree { get; set; }
    }
}
