using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using ServicesProxy.rhdesa;
using System.Data;
using Common;
using System.Text;

namespace RHPro.Controls
{
    public partial class EstilosNewHome : System.Web.UI.UserControl
    {
        public int codestilo;
        public string fondoCabecera = "#efefef";
        public string fuenteCabecera = "#ffffff";        
        public string fondoPiso = "#f00";
        public string fuentePiso = "#333333";
        public string fondoFecha = "#efefef";
        public string fuenteFecha = "#333333";
        public string fondoModulos = "#efefef";
        public string fuenteModulos = "#333333";
        public string coloricono = "#333333";
        public string coloriconomenutop = "#333333";
        public string fondocontppal = "#eef2f5";

        public string FuenteCabeceraGadget_Color = "#5c6578";
        public string FuenteCabeceraGadget_Font = "Arial";
        public string FuenteCabeceraGadget_Size = "9pt";
        public string BackgroundCabeceraGadget = "#f6f8fb";

        public String AnchoMenuLinks = "300px";
        

        public string AnchoPagina = "";
        public string RadioGadget = "";
        public int defecto;
        public int activo;


        public string getColor()
        {
            return fondoPiso;
        }

        protected void Page_Load(object sender, EventArgs e)
        {
            //Usuarios CUsuarios = new Usuarios();
            //CUsuarios.Inicializar_Estilos();
            fondoCabecera = Convert.ToString(System.Web.HttpContext.Current.Session["EstiloR4_fondoCabecera"]);
            fondoPiso = Convert.ToString(System.Web.HttpContext.Current.Session["EstiloR4_fondoPiso"]);
            fuentePiso = Convert.ToString(System.Web.HttpContext.Current.Session["EstiloR4_fuentePiso"]);
            fondoFecha = Convert.ToString(System.Web.HttpContext.Current.Session["EstiloR4_fondoFecha"]);
            fuenteFecha = Convert.ToString(System.Web.HttpContext.Current.Session["EstiloR4_fuenteFecha"]);
            fondoModulos = Convert.ToString(System.Web.HttpContext.Current.Session["EstiloR4_fondoModulos"]);
            fuenteModulos = Convert.ToString(System.Web.HttpContext.Current.Session["EstiloR4_fuenteModulos"]);
            coloricono = Convert.ToString(System.Web.HttpContext.Current.Session["EstiloR4_coloricono"]);
            coloriconomenutop = Convert.ToString(System.Web.HttpContext.Current.Session["EstiloR4_coloriconomenutop"]);
            fondocontppal = Convert.ToString(System.Web.HttpContext.Current.Session["EstiloR4_fondocontppal"]);
            BackgroundCabeceraGadget = Convert.ToString(System.Web.HttpContext.Current.Session["EstiloR4_fondoGadget"]);
            FuenteCabeceraGadget_Color = Convert.ToString(System.Web.HttpContext.Current.Session["EstiloR4_fuenteGadget"]);
            /*
            StringBuilder sb = new StringBuilder();
            string stringSeparator = string.Empty;
            foreach (string key in System.Web.HttpContext.Current.Session.Keys)
            {
                if ( (key.Contains("EstiloR4_")) || (key=="CarpetaEstilo"))
                {
                    sb.AppendFormat("{0}{1}", stringSeparator, Encryptor.Encrypt("56238", string.Concat(key, "@", System.Web.HttpContext.Current.Session[key])));
                    stringSeparator = "_";
                }
            }

            ifrmEst.Attributes.Add("location", string.Format("~/CopyAspNetSessionToAspSession.asp?params={0}&returnURL={1}", HttpContext.Current.Server.UrlEncode(sb.ToString()), HttpContext.Current.Server.UrlEncode(HttpContext.Current.Request.Url.AbsolutePath)));
          */
        }



        //public static void CopySessionEstilos()
        //{//Solamente paso la variable del lenguaje
        //    StringBuilder sb = new StringBuilder();
        //    string stringSeparator = string.Empty;
        //    foreach (string key in System.Web.HttpContext.Current.Session.Keys)
        //    {
        //        if (key.Contains("EstiloR4_"))
        //        {
        //            sb.AppendFormat("{0}{1}", stringSeparator, Encryptor.Encrypt("56238", string.Concat(key, "@", System.Web.HttpContext.Current.Session[key])));
        //            stringSeparator = "_";
        //        }
        //    }

        //    ifrmEst.Attributes.Add("location", string.Format("~/../CopyAspNetSessionToAspSession.asp?params={0}&returnURL={1}", HttpContext.Current.Server.UrlEncode(sb.ToString()), HttpContext.Current.Server.UrlEncode(HttpContext.Current.Request.Url.AbsolutePath)));

        //    //HttpContext.Current.Response.Redirect(string.Format("~/../CopyAspNetSessionToAspSession.asp?params={0}&returnURL={1}", HttpContext.Current.Server.UrlEncode(sb.ToString()), HttpContext.Current.Server.UrlEncode(HttpContext.Current.Request.Url.AbsolutePath)));
        //}
    }
}