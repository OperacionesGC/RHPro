using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Entities
{
    [Serializable]
    public class Search
    {
        public string Module { get; set; }
        public string MenuDescription { get; set; }
        public string Action { get; set; }
        public string Description { get; set; }

    }
}
