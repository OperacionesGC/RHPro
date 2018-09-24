using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace RhPro.Oodd.Web.DataModel
{
    public partial class OrgNode
    {
        public int code { get; set; }
        public int oldParentcode { get; set; }
        public int newParentcode { get; set; }
    }
}
