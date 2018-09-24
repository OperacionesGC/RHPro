using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.ComponentModel.DataAnnotations;

namespace RhPro.Oodd.Web.DataModel
{
    [MetadataType(typeof(OrgChart.OrgChartMetadata))]
    public partial class OrgChart
    {
        internal sealed class OrgChartMetadata
        {
            private OrgChartMetadata()
            {
            }
            
            [Key]
            public int orgChartCode;

        }
    }
}
