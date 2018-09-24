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

using Common;
using System.Threading;
using ServicesProxy;

namespace HomeLinkInteres
{
	
    public partial class Gadget_LinkInteres: System.Web.UI.UserControl
    {
    	public RHPro.Lenguaje ObjLenguaje;
        protected void Page_Load(object sender, EventArgs e)
        {	 
		   Repeater11.DataSource =  LinkServiceProxy.Find(Utils.SessionUserName, Utils.SessionBaseID, Utils.Lenguaje);
		   Repeater11.Visible = true;
		   Repeater11.DataBind();
        }
    }
}