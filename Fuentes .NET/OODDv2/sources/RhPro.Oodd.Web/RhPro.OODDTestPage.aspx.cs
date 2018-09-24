using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using RhPro.Oodd.Web.OrgDao;

namespace RhPro.Oodd.Web
{
    public partial class OODDTestPage : Page
    {

        protected void Page_Load(object sender, EventArgs e)
        {
            //Obtengo del Request el id de la Base con el cual trabajar
            string pbase = this.Page.Request.QueryString.Get("base");

            if (!string.IsNullOrEmpty(pbase)) {
                OrgDaoHandler.BASE = int.Parse(pbase);
            } else {
                OrgDaoHandler.BASE = 2;
            }

            if (Page.IsPostBack)
            {
                if (this.txtPNGBytes.Value.Length > 0)
                {
                    Server.Transfer("PrintPreview.aspx");
                }
            }
        }

        public string setInitialParams()
        {
            string[] keys = System.Web.Configuration.WebConfigurationManager.AppSettings.AllKeys;
            string initParameters = string.Empty;

            if (keys.Count() > 0)
            {
                foreach (string s in keys)
                {
                    if (!string.IsNullOrEmpty(initParameters))
                        initParameters += ",";

                    initParameters += s + "=" + System.Configuration
                        .ConfigurationManager.AppSettings[s];
                }
            }
            System.Diagnostics.Debug.WriteLine(initParameters);
            return initParameters;
        }
    }
}