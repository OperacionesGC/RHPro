using System;
using System.Threading;
using Common;
using Entities;
using ServicesProxy;

namespace RHPro.Controls
{
    public partial class Footer : System.Web.UI.UserControl
    {
        #region Page Handles

        protected void Page_PreRender(object sender, EventArgs e)
        {
                LoadFooter();
            
        }

        #endregion


        #region Methods

        /// <summary>
        /// Busca y carga la informacion del footer
        /// </summary>
        private void LoadFooter()
        {
            version.InnerText = VersionServiceProxy.Find(Utils.SessionBaseID, Thread.CurrentThread.CurrentCulture.Name);//carga la version
            patch.InnerText = PatcheServiceProxy.Find(Utils.SessionBaseID, Thread.CurrentThread.CurrentCulture.Name);//carga el patch
        }
        #endregion
    }
}