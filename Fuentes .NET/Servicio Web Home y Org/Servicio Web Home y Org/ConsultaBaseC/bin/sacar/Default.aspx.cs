using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Common;
using Entities;
using ServicesProxy;

namespace RHPro
{

    public partial class Default : System.Web.UI.Page
    {        
        #region Page Handles


        protected void Page_Load(object sender, EventArgs e)
        {
            cLogin.UserLogin += cLogin_UserLogin;
            cLogin.UserLogout += cLogin_UserLogout;


            flagSpain.Visible = bool.Parse(ConfigurationManager.AppSettings["EnableES"]);
            flagBrazil.Visible = bool.Parse(ConfigurationManager.AppSettings["EnablePT"]);
            flagEngland.Visible = bool.Parse(ConfigurationManager.AppSettings["EnableEN"]);

            if (Thread.CurrentThread.CurrentCulture.ToString() != ConfigurationManager.AppSettings["Idioma"] &&
                (Session["ChangeLanguage"] == null || Session["ChangeLanguage"].ToString() != "1"))
            {
                ChangeLanguage(ConfigurationManager.AppSettings["Idioma"]);
            }

        }

        protected void Page_SaveStateComplete(object sender, EventArgs e)
        {
            if (!Utils.IsUserLogin && Utils.SesionIniciada)
            {
                Utils.LogoutUser();
            }
        }

        #endregion

        #region Controls Handles

        protected void cLogin_UserLogin(object sender, EventArgs e)
        {
            LoadPageInfo();
        }

        protected void cLogin_UserLogout(object sender, EventArgs e)
        {
            LoadPageInfo();
        }     

        #endregion

        #region Methods

        /// <summary>
        /// Carga la informacion de la pagina
        /// </summary>
        public void LoadPageInfo()
        {   
            //mruMain.LoadMRU();
            mlsMain.LoadModule();
            //linksMain.LoadLinks();
            //messageMain.LoadMessage();
            //cFooterPage.LoadFrame();
        }   

        #endregion      

    
        protected void flagEngland_Click(object sender, ImageClickEventArgs e)
        {
            Session["ChangeLanguage"] = "1";
            ChangeLanguage("en-US");
        }

        protected void flagSpain_Click(object sender, ImageClickEventArgs e)
        {
            Session["ChangeLanguage"] = "1";
            ChangeLanguage("es-AR");
        }

        protected void flagBrazil_Click(object sender, ImageClickEventArgs e)
        {
            Session["ChangeLanguage"] = "1";
            ChangeLanguage("pt-br");
        }

        private void ChangeLanguage(string selectedLanguage)
        {
            Response.Cookies["Language"].Value = selectedLanguage;
            Response.Redirect(Request.Url.PathAndQuery);
        }
    
    
    }



     
}

