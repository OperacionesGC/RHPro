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
using System.Data.OleDb;
using System.Web.UI;
using System.Text;
using System.Security.Cryptography;

namespace RHPro
{
    public class Global : System.Web.HttpApplication
    {
        public string lenguageValue { get; set; }
        
        
        /*-****************************************************/
        /******************************************************/

        private string GenerateHashKey()
        {
            StringBuilder myStr = new StringBuilder();
            myStr.Append(Request.Browser.Browser);
            myStr.Append(Request.Browser.Platform);
            myStr.Append(Request.Browser.MajorVersion);
            myStr.Append(Request.Browser.MinorVersion);
            myStr.Append(Request.LogonUserIdentity.User.Value);
            SHA1 sha = new SHA1CryptoServiceProvider();
            byte[] hashdata = sha.ComputeHash(Encoding.UTF8.GetBytes(myStr.ToString()));
            return Convert.ToBase64String(hashdata);
        }

        protected void Application_EndRequest(object sender, EventArgs e)
        {
            ////Pass the custom Session ID to the browser.
            //if (Response.Cookies["ASP.NET_SessionId"] != null)
            //{
            //   Response.Cookies["ASP.NET_SessionId"].Value = Request.Cookies["ASP.NET_SessionId"].Value + GenerateHashKey();              
            //}

        }

        protected void Application_AcquireRequestState(object sender, EventArgs e)
        {

            if ((Request.Cookies["ASP.NET_SessionId"] != null) && (Request.Cookies["ASP.NET_SessionId"].Value != null)
                && (!String.IsNullOrEmpty(Request.Cookies["ASP.NET_SessionId"].Value)))
            {

                string newSessionID = Request.Cookies["ASP.NET_SessionID"].Value;
                //Compruebo la longitud del id de sesion. La cookie ASP.NET_SessionID debe tener 24 caracteres
                if (newSessionID.Length <= 24)
                {
                    //*******ATAQUE EXTERNO******//
                    Response.Cookies["TriedTohack"].Value = "True";
                    //****Controlar IP de HACKEO*****
                    throw new HttpException("SE INTENTA UN ATAQUE");
                }

                //Genera un hash key para el usuario,Browser y maquina valida con NewSessionID
                string hashkey = GenerateHashKey();
                if (hashkey != newSessionID.Substring(24))
                {
                    //Log the attack details here
                    Response.Cookies["TriedTohack"].Value = "True";
                    throw new HttpException("ID:" + newSessionID.Substring(24) + " DIFERENTE DE " + hashkey);
                }

                //Use the default one so application will work as usual//ASP.NET_SessionId
                Request.Cookies["ASP.NET_SessionId"].Value = Request.Cookies["ASP.NET_SessionId"].Value.Substring(0, 24);
            }
            
            
        } 
        /******************************************************/
        /*******************************************************/
        
         
        
        
        /* Codigo que se ejectura cuando inicia la aplicacion */ 
        protected void Application_Start(object sender, EventArgs e)
        {
             

        }

        

        /* Codigo que se ejectura cuando se inicia la sesion. */
        protected void Session_Start(object sender, EventArgs e)
        {
            try
            {
               
               // Request.Cookies["ASP.NET_SessionID"].Secure = Request.IsSecureConnection;

                // Se crean las variables de sesion requeridas con valor default
                Utils.SetDefaultSessionValues();
                //jpb
                //SetBaseIdDefault();
                if (ConfigurationManager.AppSettings["Idioma"] != null)
                    Common.Utils.Lenguaje = (String)ConfigurationManager.AppSettings["Idioma"];
                else
                    Common.Utils.Lenguaje = "es-AR";                          
              
            }
            catch (Exception ex) { }
        }

        //protected void Application_BeginRequest(object sender, EventArgs e)
        //{          
          
             
        //}

        protected void Application_AuthenticateRequest(object sender, EventArgs e)
        {

        }


        /* Codigo que se ejectura cuando ocurre algun error en la aplicacion */
        protected void Application_Error(object sender, EventArgs e)
        {
            
        }

        /* Codigo que se ejectura cuando finaliza la sesion.*/
        public void Session_End(object sender, EventArgs e)
        {
            try
            {
                                
                Utils.SetDefaultSessionValues();
                Utils.CopyAspNetSessionToAspSession();
                
                //jpb
                SetBaseIdDefault();              
            
            }
            catch (Exception ex) { }

           
        }


        /* Codigo que se ejectura cuando finaliza la aplicacion */
        protected void Application_End(object sender, EventArgs e)
        {
             
        }



        /// <summary>
        /// jpb: Carga la base default
        /// </summary>
        protected internal void SetBaseIdDefault()
        {
            try{
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
            catch (Exception ex) { }
           
        }



    }
}