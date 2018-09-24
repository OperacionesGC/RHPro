using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Entities
{
    [Serializable]
    public class PopUpChangePassData
    {
        public Login Login { get; set; }
        public DataBase DataBase { get; set; }
        public string UserName { get; set; }
    }
}
