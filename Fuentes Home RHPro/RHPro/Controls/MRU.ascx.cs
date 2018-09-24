using System;
using System.Configuration;
using System.Threading;
using System.Web.UI;
using Common;
using ServicesProxy;

namespace RHPro.Controls
{
    public partial class MRU : UserControl
    {
        private  int MRUsAvailables = int.Parse(ConfigurationManager.AppSettings["CantidadMRUsVisibles"]);

        protected void Page_PreRender(object sender, EventArgs e)
        {
                LoadMRU();
        }   
        /// <summary>
        /// Busca y carga los MRU
        /// </summary>
        internal void LoadMRU()
        {
            if (string.IsNullOrEmpty(Utils.SessionUserName)){
                mruImage.Visible = true;
                mruCompleto.Visible = false;
            }
            else
            {
                mruImage.Visible = false;
                mruCompleto.Visible = true;

                MRURepeater.DataSource = MRUServiceProxy.Find(Utils.SessionUserName, MRUsAvailables, Utils.SessionBaseID, Thread.CurrentThread.CurrentCulture.Name);
                MRURepeater.DataBind();
            }
            

        }
    }
}