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

namespace WebApplication11
{
    public partial class WebUserControl1 : System.Web.UI.UserControl
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }


        protected void Page_PreRender(object sender, EventArgs e)
        {
            LoadMRU();

        }


        /// <summary>
        /// Busca y carga los MRU
        /// </summary>
        internal void LoadMRU()
        {

             
                //  mruImage.Visible = false;
                mruCompleto.Visible = true;
                MRURepeater.DataSource = SqlDataSource1;
           



        }

    }
}