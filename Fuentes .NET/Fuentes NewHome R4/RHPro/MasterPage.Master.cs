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

                //string errorVar = string.Format("var errorMessages = '{0}';", GetLocalResourceObject("msgError").ToString());
                //Page.ClientScript.RegisterClientScriptBlock(GetType(), "sessionExpired", errorVar, true);

                //El siguiente script es para controlar el cierre de sesión
                //cuando se cierra el explorador

               // string sessionClosedVar;

                //if(Utils.SesionIniciada)
               // sessionClosedVar = string.Format("var sessionClosedVar = '{0}';", GetLocalResourceObject("SessionClosed").ToString());
                //else
                //    sessionClosedVar = string.Format("var sessionClosedVar = '{0}';", "");

               // Page.ClientScript.RegisterClientScriptBlock(GetType(), "sessionClosed", sessionClosedVar, true);

                
            
        }

   
        protected void scriptManager_AsyncPostBackError(object sender, AsyncPostBackErrorEventArgs e)
        {
            scriptManager.AsyncPostBackErrorMessage = e.Exception.Message;
        }

  
    }
}
