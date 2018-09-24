using System;
using System.Collections.Generic;
using System.Configuration;
using System.Text;
using System.Threading;
using System.Web.UI;
using Common;
using ServicesProxy;


namespace RHPro.Controls
{
    public partial class Banner : System.Web.UI.UserControl
    {
        private List<Entities.Banner> Banners
        {
            get { return ViewState["Banners"] as List<Entities.Banner>; }
            set { ViewState["Banners"] = value; }
        }

        protected void Page_PreRender(object sender, EventArgs e)
        {
            Banners = BannerServiceProxy.Find(Utils.SessionBaseID, Thread.CurrentThread.CurrentCulture.Name);
            
            StringBuilder imagesArrayDeclaration = new StringBuilder();

            imagesArrayDeclaration.Append("images = new Array(");

            for (int i = 0; i < Banners.Count; i++)
            {
                imagesArrayDeclaration.AppendFormat("'{0}'{1}", Banners[i].ImageUrl, i < Banners.Count - 1 ? "," : string.Empty);
            }

            imagesArrayDeclaration.Append(");");

            Page.ClientScript.RegisterClientScriptBlock(GetType(), "imagesArrayDeclaration", imagesArrayDeclaration.ToString(), true);           

        }
    }
}