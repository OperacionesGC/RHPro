using System;
using System.Threading;
using Common;
using Entities;
using ServicesProxy;

namespace RHPro.Controls
{
    public partial class Link : System.Web.UI.UserControl
    {


        
        protected void Page_PreRender(object sender, EventArgs e)
        {
            LoadLinks();
        }
        /// <summary>
        ///  Busca y carga los links disponibles
        /// </summary>
        internal void LoadLinks()
        {
            linkRepeater.DataSource = LinkServiceProxy.Find(Utils.SessionUserName, Utils.SessionBaseID, Thread.CurrentThread.CurrentCulture.Name);
            linkRepeater.DataBind();
        }
    }
}