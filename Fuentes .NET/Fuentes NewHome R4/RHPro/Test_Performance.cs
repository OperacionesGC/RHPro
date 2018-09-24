using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Configuration;

namespace RHPro
{
    public partial class PopUpPolitics : System.Web.UI.Page
    { 
        public RHPro.Lenguaje ObjLenguaje; 
        protected void Page_Load(object sender, EventArgs e)
        {              
            ObjLenguaje = new RHPro.Lenguaje();	
        }

        public void Imprimir_Politica()
        {
            String Cliente = Convert.ToString(ConfigurationManager.AppSettings["NombreCliente"]);
            String salida = (ObjLenguaje.Label_Home("PopUpPoliticasNewHome")).Replace("@@TXT@@",Cliente);
            Response.Write(salida);           
        }
    }
}
