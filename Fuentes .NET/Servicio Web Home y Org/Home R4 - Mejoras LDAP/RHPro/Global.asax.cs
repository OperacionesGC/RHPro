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
using ServicesProxy.rhdesa;
using System.Data;

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
            //jpb
             //SetBaseIdDefault();
            if (ConfigurationManager.AppSettings["Idioma"] != null)
                Common.Utils.Lenguaje = (String)ConfigurationManager.AppSettings["Idioma"];
            else
                Common.Utils.Lenguaje = "es-AR";
        }

        protected void Application_BeginRequest(object sender, EventArgs e)
        {
           
           /*
            if (Request.Cookies["Language"] != null && Request.Cookies["Language"].Value != "en")
            {
               // string selectedLanguage = "pt-BR";//Request.Cookies["Language"].Value;

            
              //  Common.Utils.Lenguaje = selectedLanguage;
                //System.Web.HttpContext.Current.Session["Lenguaje"] = selectedLanguage;
               
                  string selectedLanguage =  (String) ConfigurationManager.AppSettings["Idioma"]; // Request.Cookies["Language"].Value;
                //Common.Utils.Lenguaje = selectedLanguage;
             
               
                Thread.CurrentThread.CurrentCulture = new CultureInfo(selectedLanguage);
                Thread.CurrentThread.CurrentUICulture = new CultureInfo(selectedLanguage);
                                  
            }
           */             
          
             
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
            //jpb
            SetBaseIdDefault();
            
             
        }

        protected void Application_End(object sender, EventArgs e)
        {
          
        }



        /// <summary>
        /// jpb: Carga la base default
        /// </summary>
        protected internal void SetBaseIdDefault()
        {
            string dsm = ConfigurationManager.AppSettings["DatabaseSelectionMethod"].ToLower();
            System.Collections.Generic.List<DataBase> DataBases = DataBaseServiceProxy.Find(dsm);

            //Busca la base por default del webconfig
            for (int i = 0; i < DataBases.Count; i++)
            {
                if (DataBases[i].IsDefault.ToString() == "TrueValue")//verifica si la base es default
                {
                    Utils.SessionBaseID = DataBases[i].Id;
                }
            }            
           
        }



    }
}