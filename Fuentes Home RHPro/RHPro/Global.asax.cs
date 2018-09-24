using System;
using System.Collections.Generic;
using System.Configuration;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Web;
using System.Web.Security;
using System.Web.SessionState;
using Common;
using Entities;
using ServicesProxy;

namespace RHPro
{
    public class Global : System.Web.HttpApplication
    {
        public string lenguageValue { get; set; }
        protected void Application_Start(object sender, EventArgs e)
        {
          
        }

        protected void Session_Start(object sender, EventArgs e)
        {
            // Se crean las variables de sesion requeridas con valor default
            Utils.SetDefaultSessionValues();
            
        }

        protected void Application_BeginRequest(object sender, EventArgs e)
        {
            if (Request.Cookies["Language"] != null && Request.Cookies["Language"].Value != "en")
            {
                string selectedLanguage = Request.Cookies["Language"].Value;
                Thread.CurrentThread.CurrentCulture = new CultureInfo(selectedLanguage);
                Thread.CurrentThread.CurrentUICulture = new CultureInfo(selectedLanguage);
            }
        }

        protected void Application_AuthenticateRequest(object sender, EventArgs e)
        {

        }

        protected void Application_Error(object sender, EventArgs e)
        {
            
        }

        protected void Session_End(object sender, EventArgs e)
        {
            Utils.SetDefaultSessionValues();
            Utils.CopyAspNetSessionToAspSession();
        }

        protected void Application_End(object sender, EventArgs e)
        {
          
        }
    }
}