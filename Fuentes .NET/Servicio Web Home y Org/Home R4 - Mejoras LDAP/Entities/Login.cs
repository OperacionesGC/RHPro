using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Entities
{
    [Serializable]
    public class Login
    {
        public bool IsValid { get; set; }
        public string Messege { get; set; }
        public bool RequiredChangePassword { get; set; }
        public string Lenguaje { get; set; }
        public string MaxEmpl { get; set; }
    }
}
