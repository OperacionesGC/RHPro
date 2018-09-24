using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ServicesProxy
{
    public static class UtilsProxy
    {
        public static  void ChangeWS(string root)
        {
            //Properties.Settings.Default["ServicesProxy_ar_com_rhpro_prueba_Consultas"] = root;
            Properties.Settings.Default["ServicesProxy_rhdesa_Consultas"] = root;            
        }

        public static void ChangeWS_MetaHome(string root)
        {          
            
           Properties.Settings.Default["ServicesProxy_MetaHome_MH_Externo"] = root;    
            
        }
    }
}
