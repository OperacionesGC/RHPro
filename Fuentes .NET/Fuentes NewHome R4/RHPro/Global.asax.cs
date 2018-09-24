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

namespace RHPro
{
    public class Global : System.Web.HttpApplication
    {
        public string lenguageValue { get; set; }

        /* Codigo que se ejectura cuando inicia la aplicacion */
        protected void Application_Start(object sender, EventArgs e)
        {

        }


        /* Codigo que se ejectura cuando se inicia la sesion. */
        protected void Session_Start(object sender, EventArgs e)
        {
            try
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
            catch (Exception ex) { }
        }

        protected void Application_BeginRequest(object sender, EventArgs e)
        {


        }

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
            try
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
            catch (Exception ex) { }

        }



    }
}