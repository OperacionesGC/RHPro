using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Entities
{
    [Serializable]
    public class PopUpSearchData
    {
        public string UserName { get; set; }
        public string DataBase { get; set; }
        public string WordToFind { get; set; }
    }
}
