using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.ComponentModel.DataAnnotations;

namespace RhPro.Oodd.Web.DataModel
{
    [MetadataType(typeof(OrgNode.OrgNodeMetadata))]
    public partial class OrgNode
    {
        internal sealed class OrgNodeMetadata
        {
            private OrgNodeMetadata()
            {
            }
            
            [Key]
            public int code;

        }
    }
}
