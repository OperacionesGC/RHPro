using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Common;
using Entities;
using ServicesProxy;

namespace RHPro.Controls
{
    public partial class Message : System.Web.UI.UserControl
    {
        protected void Page_PreRender(object sender, EventArgs e)
        {
                LoadMessage();
        }

        /// <summary>
        /// Busca y carga los mensajes disponibles
        /// </summary>
        internal void LoadMessage()
        {
            messageRepeater.DataSource = MenssageServiceProxy.Find(Utils.SessionBaseID, Thread.CurrentThread.CurrentCulture.Name);
            messageRepeater.DataBind();
        }
    }
}