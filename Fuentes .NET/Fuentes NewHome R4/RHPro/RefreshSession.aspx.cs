using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace RHPro
{
    public partial class RefreshSession : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
              System.Web.HttpContext.Current.Session["RHPRO_FefreshSession"] = "-1";
              
                
            /*
            if (Session["ChangeLanguage"] == null)
            {
                Response.Write("NULL");
            }
            else
            {
                Response.Write(Session["ChangeLanguage"].ToString());
            }

            Session["ChangeLanguage"] = Session["ChangeLanguage"];
              */
        }
    }
}
