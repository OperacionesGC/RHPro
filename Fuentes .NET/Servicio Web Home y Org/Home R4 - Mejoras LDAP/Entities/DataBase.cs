using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Common;

namespace Entities
{
    [Serializable]
    public class DataBase
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public Utils.IsDefaultConstants IsDefault { get; set; }
        public Utils.IntegrateSecurityConstants IntegrateSecurity { get; set; }  
    }
    
}
