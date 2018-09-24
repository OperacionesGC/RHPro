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
using ServicesProxy.rhdesa;

namespace RHPro.Controls
{
    public partial class ConfigGadgets : System.Web.UI.UserControl
    {
        public Lenguaje Obj_Lenguaje;
        public int ContadorGadget;
        protected void Page_Load(object sender, EventArgs e)
        { 
            Obj_Lenguaje = new Lenguaje();

            string BaseId = Common.Utils.SessionBaseID;
            string UserName = Common.Utils.SessionUserName;

           // string sql = "SELECT ROW_NUMBER() OVER(ORDER BY gadtitulo ASC)  'pos' ,* FROM Gadgets WHERE gadactivo=0 AND gaduser='" + UserName + "' ORDER BY gadtitulo ASC";                       
            string sql = "SELECT  * FROM Gadgets WHERE gadactivo=0 AND gaduser='" + UserName + "' ORDER BY gadtitulo ASC";                       
         
            Consultas cc = new Consultas();
            DataSet ds = cc.get_DataSet(sql, BaseId);

            Repeater1.DataSource = ds;
            Repeater1.DataBind();
          
        } 
        
        protected void Page_PreRender(object sender, EventArgs e)
        {
           
        }
    }


    
}