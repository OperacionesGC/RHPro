using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Web;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using Common;
using ServicesProxy;

namespace RHPro.Controls
{
    public partial class FooterPage : System.Web.UI.UserControl
    {
        #region Properties
        public string srcFrame
        {
            get { return "./TempPagPie.html"; }
        }
        #endregion
        protected void Page_PreRender(object sender, EventArgs e)
        {
            LoadFrame();
        }
        public  void LoadFrame()
        {
            Entities.FooterPage footer = FooterPageServiceProxy.Find(Utils.SessionUserName, Utils.SessionBaseID, Thread.CurrentThread.CurrentCulture.Name);
            
            if(footer != null && !string.IsNullOrEmpty(footer.Title))
            {
                iframeExterno.Attributes["src"] = "./../" + footer.Title;
                iframeExterno.Attributes["scrolling"] = "auto";
            }else
            {
                iframeExterno.Attributes["src"] = srcFrame;
                iframeExterno.Attributes["scrolling"] = "no";
            }
        }
    }
}