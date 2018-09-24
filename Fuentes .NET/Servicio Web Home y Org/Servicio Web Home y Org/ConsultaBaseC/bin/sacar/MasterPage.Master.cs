using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Configuration;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Common;
using Entities;
using ServicesProxy;
using System.Resources;
using System.Reflection;
 

namespace RHPro
{
    public partial class MasterPage : System.Web.UI.MasterPage
    {
        #region Constants

        /// <summary>
        /// Direccion de la url del popup
        /// </summary>
        private const string UrlPopup = "~/PopUpSearch.aspx";
        
        #endregion
        
        protected void Page_Load(object sender, EventArgs e)
        {
           /*
             flagSpain.Visible = bool.Parse(ConfigurationManager.AppSettings["EnableES"]);
             flagBrazil.Visible = bool.Parse(ConfigurationManager.AppSettings["EnablePT"]);
             flagEngland.Visible = bool.Parse(ConfigurationManager.AppSettings["EnableEN"]);
             
            if (Thread.CurrentThread.CurrentCulture.ToString() != ConfigurationManager.AppSettings["Idioma"] &&
                (Session["ChangeLanguage"] == null || Session["ChangeLanguage"].ToString() != "1"))
            {
                ChangeLanguage(ConfigurationManager.AppSettings["Idioma"]);
            }
            */
            //ResourceManager m_ResourceManager = new ResourceManager("MasterPage.Master", Assembly.GetExecutingAssembly(), null);

            string errorVar = string.Format("var errorMessages = '{0}';", GetLocalResourceObject("msgError").ToString());
            Page.ClientScript.RegisterClientScriptBlock(GetType(), "sessionExpired", errorVar, true);

            //El siguiente script es para controlar el cierre de sesión
            //cuando se cierra el explorador

            string sessionClosedVar;

            //if(Utils.SesionIniciada)
                sessionClosedVar = string.Format("var sessionClosedVar = '{0}';", GetLocalResourceObject("SessionClosed").ToString());
            //else
            //    sessionClosedVar = string.Format("var sessionClosedVar = '{0}';", "");

            Page.ClientScript.RegisterClientScriptBlock(GetType(), "sessionClosed", sessionClosedVar, true); 
        }

        

        protected void scriptManager_AsyncPostBackError(object sender, AsyncPostBackErrorEventArgs e)
        {
            scriptManager.AsyncPostBackErrorMessage = e.Exception.Message;
        }

        

/*
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
        */
        //protected void btnSearchButton_Click(object sender, EventArgs e)
        //{
        //    if (!string.IsNullOrEmpty(Utils.SessionUserName))
        //    {
        //        if (!string.IsNullOrEmpty(txtSearch.Text))
        //        {
        //            ShowPopUpSearchData();
        //        }
        //        else
        //        {
        //            ScriptManager.RegisterStartupScript(Page, GetType(), "Mensaje", string.Format("javascript:alert('{0}');", GetLocalResourceObject("NotText.Text")), true);
        //        }                
        //    }
        //    else
        //    {
        //        ScriptManager.RegisterStartupScript(Page, GetType(), "Mensaje", string.Format("javascript:alert('{0}');", GetLocalResourceObject("NotLog.Text")), true);       
        //    }
        //}

        /// <summary>
        /// Muestra el PopUp de resultados de la búsqueda
        /// </summary>
     /*
        private void ShowPopUpSearchData()
        {
            PopUpSearchData popUpSearchData = new PopUpSearchData
            {
                DataBase = Utils.SessionBaseID,
                UserName = Utils.SessionUserName,
                WordToFind = txtSearch.Text
            };

            Session["PopUpSearchData"] = popUpSearchData;
            ScriptManager.RegisterStartupScript(Page, GetType(), "AbrirPopup", String.Format("javascript:window.open('{0}','urlPopup','height=550,width=500,status=yes,toolbar=no,menubar=no,location=no,resizable=yes,scrollbars=yes,left=5,top=5');", this.ResolveUrl(UrlPopup.ToString())), true);
        }
      */
    }
}
