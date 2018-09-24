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

namespace HomeMensajes
{
	
    public partial class Gadget_Mensajes : System.Web.UI.UserControl
    {
    	public RHPro.Lenguaje ObjLenguaje;
        protected void Page_Load(object sender, EventArgs e)
        {
		   //Traigo desde la dll Consultas.dll la conexion para realizar la consulta	 
           ServicesProxy.rhdesa.Consultas cc = new ServicesProxy.rhdesa.Consultas();
		   /*
		   string NroBase = Common.Utils.SessionBaseID;           
		   string CnStr = cc.constr(NroBase).Replace("Provider=SQLOLEDB.1;","");   		   
	       SqlDataSource1.ConnectionString=CnStr;
	       SqlDataSource1.SelectCommand="SELECT  * FROM home_mensaje WHERE rhpro=-1  ORDER BY hmsjfecalta DESC";        
		   */
		   
		   
		   string CnStr = cc.constr("2").Replace("Provider=SQLOLEDB.1;","");   		   
	       SqlDataSource1.ConnectionString=CnStr;
	       SqlDataSource1.SelectCommand="SELECT  * FROM user_per";        
		   
		   
		    	
		   Repeater1.Visible = true;
		   
        }
    }
}