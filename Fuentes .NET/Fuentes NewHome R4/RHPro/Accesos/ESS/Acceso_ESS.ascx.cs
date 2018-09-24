using System;
using System.Collections;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Xml.Linq;

namespace Accesos
{
	
    public partial class Acceso_ESS : System.Web.UI.UserControl
    {
    	public RHPro.Lenguaje ObjLenguaje;
        protected void Page_Load(object sender, EventArgs e)
        {   
			ObjLenguaje = new RHPro.Lenguaje();
        }
    }
}