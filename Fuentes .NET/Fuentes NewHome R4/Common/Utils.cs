using System;
using System.Configuration;
using System.Text;
using System.Web.SessionState;
using System.Web.UI;
using System.Web;
using System.Collections.Generic;
using System.Text.RegularExpressions;
 

namespace Common
{
    public static class Utils
    {
        private static bool sesionIniciada = false;

        public enum IntegrateSecurityConstants
        {
            TrueValue = -1,
            FalseValue = 0
        }

        public enum IsDefaultConstants
        {
            TrueValue = -1,
            FalseValue = 0
        }

        /// <summary>
        /// Verifica si el navegador esta habilitado para usar imagenes SVG
        /// </summary>
        /// <returns></returns>
        public static bool Navegador_Habilitado_SVG()
        {
           

            Boolean Usa_PNG = false;
            String userAgent = HttpContext.Current.Request.UserAgent.ToLower();
            String[] Motores_JS_Controlados = Convert.ToString(ConfigurationManager.AppSettings["Motores_JS_Controlados"]).Split(';');
            foreach (String motorJS in Motores_JS_Controlados)
            {
                if (userAgent.Contains(motorJS.ToLower()))
                {
                    Usa_PNG = true;
                    break;
                }
            }

            return  !Usa_PNG;
        }

        /// <summary>
        /// Arma los iconos determinando la version del navegador
        /// </summary>
        /// <param name="URL_Icono"></param>
        /// <param name="Title"></param>
        /// <param name="ID"></param>
        /// <param name="evento"></param>
        /// <param name="clase"></param>
        /// <returns></returns>         

        public static String Armar_Icono(String URL_Icono, String clase, String Title, String parametros, String ID)
        {
            String Salida = "";
            bool SVG_Habilitado;
            try
            {
                SVG_Habilitado = Navegador_Habilitado_SVG();

                if (SVG_Habilitado)
                {
                    Salida += "<embed src='" + URL_Icono + "' type='image/svg+xml'   ";
                    if (clase != "")
                        Salida += " class='" + clase + "'";

                    Salida += " pluginspage='http://www.adobe.com/svg/viewer/install/'  ";

                }
                else
                {
                    URL_Icono = URL_Icono.ToUpper().Replace("/SVG/", "/PNG/");
                    URL_Icono = URL_Icono.ToUpper().Replace(".SVG", ".PNG");

                    Salida = "<img src='" + URL_Icono + "'   ";
                    if (clase != "")
                        Salida += "class='" + clase + "_PNG'";
                }


                if (Title != "")
                    Salida += " title='" + Title + "'";

                if (ID == "")
                    ID = DateTime.Now.ToString();

                Salida += " id='" + ID + "'";

                if (parametros!="")
                    Salida += parametros;


                if (SVG_Habilitado)
                    Salida += " ></embed>";
                else
                    Salida += " ></img>";


            }
            catch (Exception ex) { }

            return Salida;

        }

         
        public static String Armar_Icono(String URL_Icono, String clase, String Title, String parametros, String ID, string onclick)
        {
            String Salida = "";
            bool SVG_Habilitado;
            try
            {
                SVG_Habilitado = Navegador_Habilitado_SVG();

                if (SVG_Habilitado)
                {
                    Salida += "<embed src='" + URL_Icono + "' type='image/svg+xml'   ";
                    if (clase != "")
                        Salida += " class='" + clase + "'";

                    Salida += " pluginspage='http://www.adobe.com/svg/viewer/install/'  ";

                    if (onclick!="")
                        Salida += " onload=\"this.getSVGDocument().onclick = function(event){ " + onclick + " };\"";
                }
                else
                {
                    URL_Icono = URL_Icono.ToUpper().Replace("/SVG/", "/PNG/");
                    URL_Icono = URL_Icono.ToUpper().Replace(".SVG", ".PNG");

                    Salida = "<img src='" + URL_Icono + "'  ";
                    if (clase != "")
                        Salida += "class='" + clase + "_PNG'";

                    Salida += " onclick=\"" + onclick + "\" ";
                }


                if (Title != "")
                    Salida += " title='" + Title + "'";

                if (ID == "")
                    ID = DateTime.Now.ToString();

                Salida += " id='" + ID + "'";

                if (parametros != "")
                    Salida += parametros;

                if (SVG_Habilitado)
                    Salida += " ></embed>";
                else
                    Salida += " ></img>";

            }
            catch (Exception ex) { }

            return Salida;

        }
         

        public static string file_get_contents(string fileName)
        {

            string sContents = string.Empty;
            if (fileName.ToLower().IndexOf("http:") > -1)
            {
                // URL 
                System.Net.WebClient wc = new System.Net.WebClient();
                byte[] response = wc.DownloadData(fileName);
                sContents = System.Text.Encoding.ASCII.GetString(response);
            }
            else
            {
                // Regular Filename 
                System.IO.StreamReader sr = new System.IO.StreamReader(fileName);
                sContents = sr.ReadToEnd();
                sr.Close();
            }
            return sContents;
        }

        /*public static String Armar_Icono(String URL_Icono, String clase, String Title, String parametros, String ID)
        {
            String Salida = "";
            try
            {
                if (Navegador_Habilitado_SVG())
                {
                    Salida = "<img src='" + URL_Icono + "'   ";
                    if (clase != "")
                        Salida += "class='" + clase + "'";
                }
                else
                {
                    URL_Icono = URL_Icono.ToUpper().Replace("/SVG/", "/PNG/");
                    URL_Icono = URL_Icono.ToUpper().Replace(".SVG", ".PNG");

                    Salida = "<img src='" + URL_Icono + "'   ";
                    if (clase != "")
                        Salida += "class='" + clase + "_PNG'";
                }


                if (Title != "")
                    Salida += " title='" + Title + "'";

                if (ID != "")
                    Salida += " id='" + ID + "'";

                if (parametros != "")
                    Salida += parametros;


                Salida += " ></img>";
            }
            catch (Exception ex) { }

            return Salida;
            //<img src="img/Modulos/SVG/<%#Eval("MenuName") %>.svg"  class="IconoModulo" title="<%# ObjLenguaje.Label_Home((String)Eval("MenuTitle"))%>   "></img>

        }
        */


       

        //public static String Armar_Icono(String URL_Icono, String clase, String Title, String parametros, String ID)
        //{
        //    String Salida = "";
        //    bool AceptaSVG = false;
        //    try
        //    {
        //        if (Navegador_Habilitado_SVG())
        //        {
        //            Salida = "<object  type='image/svg+xml' data='" + URL_Icono + "'   ";
        //            if (clase != "")
        //                Salida += "class='" + clase + "'";

        //            AceptaSVG = true;
        //        }
        //        else
        //        {
        //            URL_Icono = URL_Icono.ToUpper().Replace("/SVG/", "/PNG/");
        //            URL_Icono = URL_Icono.ToUpper().Replace(".SVG", ".PNG");

        //            Salida = "<img src='" + URL_Icono + "'   ";
        //            if (clase != "")
        //                Salida += "class='" + clase + "_PNG'";
        //        }


        //        if (Title != "")
        //            Salida += " title='" + Title + "'";

        //        if (ID != "")
        //            Salida += " id='" + ID + "'";

        //        if (parametros != "")
        //            Salida += parametros;


        //       if (AceptaSVG)
        //           Salida += " ></object>";
        //        else
        //           Salida += " ></img>";
        //    }
        //    catch (Exception ex) { }

        //    return Salida;
        //    //<img src="img/Modulos/SVG/<%#Eval("MenuName") %>.svg"  class="IconoModulo" title="<%# ObjLenguaje.Label_Home((String)Eval("MenuTitle"))%>   "></img>

        //}

        //<object  data="img/modulos/SVG/CONTROLMODULOS.svg"   type="image/svg+xml"  ></object>

        //public static string getMenuDesc(String NombreModulo)//***DESHABILITAR PARA CAMBIO MENU****///
        //{
        //    String modulo = NombreModulo;

        //    switch (NombreModulo.ToUpper())
        //    {
        //        case "ADMPER": modulo = "adp"; break;
        //        case "ALERTAS": modulo = "ale"; break;
        //        case "ANALISIS": modulo = "analisis"; break;
        //        case "BIENES": modulo = "bdc"; break;
        //        case "CAPACITACION": modulo = "Capacita"; break;
        //        case "EMPLEOS": modulo = "empleos"; break;
        //        case "GTI": modulo = "asistencia"; break;
        //        case "EVALUACION": modulo = "evaluacion"; break;
        //        case "LIQUIDACION": modulo = "liq"; break;
        //        case "PLAN": modulo = "Carreras"; break;
        //        case "POLITICAS": modulo = "pol"; break;
        //        case "SALUD": modulo = "so"; break;
        //        case "SUPERVISOR": modulo = "sup"; break;
        //        case "DIS": modulo = "DIS"; break;
        //        case "BIENESTAR": modulo = "bie"; break;
        //        case "PLANTA": modulo = "PP"; break;
        //        case "EMBARGOS": modulo = "emb"; break;
        //        case "SIM": modulo = "sim"; break;
        //        case "COMPETENCIAS": modulo = "gdc"; break;
        //        case "INFOGER": modulo = "mig"; break;
        //    }
        //    return modulo;
          
        //}


        
 
        public static string getMenuDir(String NombreModulo)
        {
             return NombreModulo;//***HABILITAR PARA CAMBIO MENU****///

/*
            String modulo = NombreModulo;
            switch (NombreModulo)
            {
                case "ADMPER": modulo = "ADP"; break;
                case "ALERTAS": modulo = "ALE"; break;
                case "ANALISIS": modulo = "ANR"; break;
                case "BIENES": modulo = "BDC"; break;
                case "CAPACITACION": modulo = "CAP"; break;
                case "EMPLEOS": modulo = "POST"; break;
                case "GTI": modulo = "GTI"; break;
                case "EVALUACION": modulo = "EVAL"; break;
                case "LIQUIDACION": modulo = "LIQ"; break;
                case "PLAN": modulo = "PDD"; break;
                case "POLITICAS": modulo = "POL"; break;
                case "SALUD": modulo = "SO"; break;
                case "SUPERVISOR": modulo = "SUP"; break;
                case "DIS": modulo = "DIS"; break;
                case "BIENESTAR": modulo = "BIE"; break;
                case "PLANTA": modulo = "PP"; break;
                case "EMBARGOS": modulo = "EMB"; break;
                case "SIM": modulo = "SIM"; break;
                case "COMPETENCIAS": modulo = "GDC"; break;
                case "INFOGER": modulo = "MIG"; break;
            }
            return modulo;
  */
        }


        //public static string GetIP()
        //{
        //    string strHostName = "";

        //    try
        //    {
        //        strHostName = System.Net.Dns.GetHostName();
        //        System.Net.IPHostEntry ipEntry = System.Net.Dns.GetHostEntry(strHostName);
        //        System.Net.IPAddress[] addr = ipEntry.AddressList;
        //        return addr[addr.Length - 1].ToString();
                
               
        //    }
        //    catch { return ""; }

        //}
 

        public static  List<String> Modulos_Habilitados 
        {
            get
            {
                return (List<String>)System.Web.HttpContext.Current.Session["RHPRO_Modulos_Habilitados"];

            }
            set
            {
                System.Web.HttpContext.Current.Session["RHPRO_Modulos_Habilitados"] = value;

            }
        }
 


        public static bool Habilitado(List<String> ListPerfUsr, String ListAccesos)
        {
            
            String[] Misplit = ListAccesos.Split(',');
            foreach (String PerfUsr in ListPerfUsr)
            {
                foreach (String Acceso in Misplit)
                {
                    if ((PerfUsr == Acceso) || (Acceso == "*"))
                    {

                        return true;
                    }
                }
            }
            return false;
        }

        private static readonly HttpSessionState Session = System.Web.HttpContext.Current.Session;


        public static bool SesionIniciada
        {
            get
            {
                try
                {
                    return (bool)System.Web.HttpContext.Current.Session["sesionIniciada"];
                }
                catch
                {
                    System.Web.HttpContext.Current.Session["sesionIniciada"] = false;
                    return (bool)System.Web.HttpContext.Current.Session["sesionIniciada"];
                }
            }
            set
            {
                System.Web.HttpContext.Current.Session["sesionIniciada"] = value;
            }
        }        
        

               

          /// <summary>
        ///  
        /// </summary>
        public static string MSGE_ERROR(Exception e){
           // return "<span   onclick=\"this.style.visibility = 'hidden'\" style='float:left;cursor:pointer; border:font-family:Arial; font-size:9pt; color:#333;border:4px #333333 solid; position:relative; left:30px; top:30px; padding:6px; background-color:#FC9'><img src='img/error.png' align='absmiddle'> ERROR: " + e.Message + "</span>";
            //return "<script>MostrarError('EEEE');</script>";
            return "";
        }
         


        /// <summary>
        /// Contiene el lenguaje
        /// </summary>
        public static string Lenguaje
        {
            get
            {
                //return (string)Session["Lenguaje"];
                 return (string)System.Web.HttpContext.Current.Session["Lenguaje"];                                
              
            }
            set
            {
                //Session["Lenguaje"] = value;
                 System.Web.HttpContext.Current.Session["Lenguaje"] = value;
                 
            }
        }

        /// <summary>
        /// Contiene el máximo de empleados 
        /// </summary>
        public static string MaxEmpleados
        {
            get
            {
                //return (string)Session["MaxEmpleados"];
                return (string)System.Web.HttpContext.Current.Session["MaxEmpleados"];
            }
            set
            {
                //Session["MaxEmpleados"] = value;
                System.Web.HttpContext.Current.Session["MaxEmpleados"] = value;
            }
        }

        /// <summary>
        /// Contiene el nombre del usuario 
        /// </summary>
        public static string SessionUserName
        {
            get 
            {
                //return (string)Session["UserName"];             
                return (string)System.Web.HttpContext.Current.Session["UserName"];
            }
            set 
            { 
                //Session["UserName"] = value;
                System.Web.HttpContext.Current.Session["UserName"] = value;
            }
        }

        /// <summary>
        /// Contiene el password del usuario 
        /// </summary>
        public static string SessionPassword
        {
            get 
            {
                //return (string)Session["Password"];
                return (string)System.Web.HttpContext.Current.Session["Password"];
            }
            set 
            { 
                //Session["Password"] = value;
                System.Web.HttpContext.Current.Session["Password"] = value;
            }
        }

        /// <summary>
        /// Contiene el ID de la base seleccionada
        /// </summary>
        public static string SessionBaseID
        {
            get 
            {
                //return (string)Session["base"];
                return (string)System.Web.HttpContext.Current.Session["base"];
            }
            set 
            { 
                //Session["base"] = value;
                System.Web.HttpContext.Current.Session["base"] = value;
            }
        }
 
         
        /// <summary>
        /// Contiene la fecha de session
        /// </summary>
        public static DateTime SessionTime
        {
            get 
            { 
                //return (DateTime)Session["Time"]; 
                return (DateTime)System.Web.HttpContext.Current.Session["Time"];
            }
            set 
            { 
                //Session["Time"] = value;
                System.Web.HttpContext.Current.Session["Time"] = value;
            }
        }
 


        /// <summary>
        /// Contiene el numero del identificador generado por el metahome
        /// </summary>
        public static string SessionNroTempLogin
        {
            get
            {
                return (string)System.Web.HttpContext.Current.Session["NroTempLogin"];
            }
            set
            {                
                System.Web.HttpContext.Current.Session["NroTempLogin"] = value;
            }
        }


        /// <summary>
        /// Contiene el LCID de la session
        /// </summary>
        public static int SessionLCID
        {
            set 
            { 
                Session.LCID = value;
                System.Web.HttpContext.Current.Session.LCID = value;
            }
            get 
            { 
                //return Session.LCID; 
                return System.Web.HttpContext.Current.Session.LCID;
            }
        }



        /// <summary>
        /// Contiene el valor correspondiente a la seguridad integrada de la base (-1 = True, 0 = False)
        /// </summary>
        public static IntegrateSecurityConstants SessionSeg_NT
        {
            set 
            { 
                //Session["seg_NT"] = (int)value;
                System.Web.HttpContext.Current.Session["seg_NT"] = (int)value;
            }
            get 
            {
                return (IntegrateSecurityConstants)System.Web.HttpContext.Current.Session["seg_NT"];
                //return (IntegrateSecurityConstants)Session["seg_NT"]; 
            } 
        }

        public static String Session_ModuloActivo
        {
            set
            {                
                System.Web.HttpContext.Current.Session["Session_ModuloActivo"] = (String)value;
            }
            get
            {
                return (String)System.Web.HttpContext.Current.Session["Session_ModuloActivo"];
      
            }
        }

        /// <summary>
        /// Si un usuario esta logeado o no.
        /// </summary>
        public static bool IsUserLogin
        {
            get 
            { 
                return !string.IsNullOrEmpty(SessionUserName); 
            }
        }

          /*
        public static string MetaHome_connString
        {
            get
            {
                string Meta_UserId = (String)ConfigurationManager.AppSettings["Meta_UserId"];
                string Meta_Password = (String)ConfigurationManager.AppSettings["Meta_Password"];
                string Meta_ConnString = (String)ConfigurationManager.AppSettings["Meta_ConnString"];
                string EncryptionKey = (String)ConfigurationManager.AppSettings["EncryptionKey"];
                string connStr = Meta_ConnString + "Password=" + Encryptor.Decrypt(EncryptionKey, Meta_Password) + ";User ID=" + Encryptor.Decrypt(EncryptionKey, Meta_UserId) + ";";

                return connStr;
            }
        }
        */

      
        /*
        public static bool MetaHome_Activo()
        {
            return Convert.ToBoolean(ConfigurationManager.AppSettings["Meta_RegistraLogin"]); 
        }

       */




       

        /// <summary>
        ///  Completa las variables de session correpondientes al logeo de un usuario
        /// </summary>
        /// <param name="userName"></param>
        /// <param name="password"></param>
        /// <param name="encriptUserData"></param>
        /// <param name="encryptionKey"></param>
        public static void LoginUser(string userName, 
                                        string password, 
                                        bool encriptUserData, 
                                        string encryptionKey,
                                        string lenguaje,
                                        string maxempl)
        { 
            

            sesionIniciada = true;
            Session["RHPRO_RecienLogueado"] = "-1";
             

            if (encriptUserData)
            {
                Utils.SessionUserName = Encryptor.Encrypt(encryptionKey, userName);
                Utils.SessionPassword = Encryptor.Encrypt(encryptionKey, password);                
            }
            else
            {
                Utils.SessionUserName = userName;
                Utils.SessionPassword = password;
            }

            //Acá se debería obtener el valor de MaxEmpleados
            //de un WS y asignarlo a la propiedad Utils.MaxEmpleados

            Utils.Lenguaje = lenguaje;
            Utils.MaxEmpleados = maxempl;

            Utils.SesionIniciada = true;
            //CopyAspNetSessionToAspSession();            
        }


        public static String Session_MenumsNro_Modulo
        {
            set
            {
                System.Web.HttpContext.Current.Session["RHPRO_MenumsNro_Modulo"] = (String)value;
            }
            get
            {
                return (String)System.Web.HttpContext.Current.Session["RHPRO_MenumsNro_Modulo"];

            }
        }

        public static int getCharOcur(String cadena, string patron)
        {
            return (cadena.Split(new String[] { patron }, StringSplitOptions.None).Length - 1);
        }

        public static string Valida_Javascript(String cadena, bool OrigenMRU)
        {
            string salida = "";

            string parentesis = getCharOcur(cadena, "(") == getCharOcur(cadena, ")") ? "" : "Paréntesis impares,";
            string llaves = getCharOcur(cadena, "{") == getCharOcur(cadena, "}") ? "" : "LLaves impares,";
            string corchetes = getCharOcur(cadena, "[") == getCharOcur(cadena, "]") ? "" : "Corchetes impares,";
            string comentario = getCharOcur(cadena, "/*") == getCharOcur(cadena, "*/") ? "" : "Falta abrir o cerrar comentario,";
            string comillas_Simples_Pares = getCharOcur(cadena, "'") % 2 == 0 ? "" : "Comillas Simples Impares,";
            string comillas_Dobles_Pares = getCharOcur(cadena, "\"") % 2 == 0 ? "" : "Comillas Dobles Impares,";
            string abrirVentana_Inexistente = "";
            string Numeral_Existente ="";
            if (OrigenMRU)
            {
                 abrirVentana_Inexistente = getCharOcur(cadena, "abrirVentana") != 0 ? "" : "Evento incorrecto,";
                 Numeral_Existente = getCharOcur(cadena, "#") == 0 ? "" : "Evento incorrecto,";
            }
            salida = parentesis + llaves + corchetes + comentario + comillas_Simples_Pares
                + comillas_Dobles_Pares + abrirVentana_Inexistente + Numeral_Existente;

            return salida;
        }

        
//        public static String ArmarAction(String action, String modulo, String menumsnro, String menuraiz, String menunro )
        public static String ArmarAction(String action, String modulo, String menumsnro, String menuraiz, String menunro, String MenumsNroModulo)
        {
            String Salida = action;

            if (action.Contains("http"))//Si viene una direccion externa directamente dejo el action como esta.
                return action;

            if (!action.Contains("../"))            
                Salida = action.Replace("('", "('../" + modulo + "/");

            if (String.IsNullOrEmpty(MenumsNroModulo))
                MenumsNroModulo = Session_MenumsNro_Modulo;
                        
            if (Salida.Contains("?"))
                Salida = Salida.Replace("?", "?menumsnro=" + menumsnro + "&menunro=" + menunro + "&MenumsNro_Modulo=" + MenumsNroModulo + "&");
            else
                Salida = Salida.Replace(".asp", ".asp?menumsnro=" + menumsnro + "&menunro=" + menunro + "&MenumsNro_Modulo=" + MenumsNroModulo);

            if (action != "" && action != "#")
                Salida += "; ifrm.location =  '../shared/asp/mru_00.asp?menumsnro=" + menumsnro + "&menuraiz=" + menuraiz + "'";
                     
            return Salida;
        }


    

        /// <summary>
        /// Limpia las variables de session correpondientes al logeo de un usuario
        /// </summary>
        public static void LogoutUser()
        {
            sesionIniciada = false;
            Utils.SessionUserName = string.Empty;
            Utils.SessionPassword = string.Empty;
            Utils.SesionIniciada = false;
            //Elimina todas las variables de session            
            System.Web.HttpContext.Current.Session.Abandon();

            //Session["lstIndex"] = null;
            //System.Web.HttpContext.Current.Session["lstIndex"] = null; 
            System.Web.HttpContext.Current.Session["lstIndex"] = 0; 
            CopyAspNetSessionToAspSession();            
        }

        /// <summary>
        /// Redirect con in a new window
        /// </summary>
        /// <param name="url"></param>
        /// <param name="target"></param>
        /// <param name="windowFeatures"></param>
        public static void Redirect(string url, string target, string windowFeatures)
        {
            HttpContext context = HttpContext.Current;

            if ((String.IsNullOrEmpty(target) ||
                target.Equals("_self", StringComparison.OrdinalIgnoreCase)) &&
                String.IsNullOrEmpty(windowFeatures))
            {
                context.Response.Redirect(url);
            }
            else
            {
                Page page = (Page)context.Handler;
                if (page == null)
                {
                    throw new InvalidOperationException(
                        "Cannot redirect to new window outside Page context.");
                }
                url = page.ResolveClientUrl(url);

                string script;
                if (!String.IsNullOrEmpty(windowFeatures))
                {
                    script = @"window.open(""{0}"", ""{1}"", ""{2}"");";
                }
                else
                {
                    script = @"window.open(""{0}"", ""{1}"");";
                }
                script = String.Format(script, url, target, windowFeatures);
                ScriptManager.RegisterStartupScript(page,
                    typeof(Page),
                    "Redirect",
                    script,
                    true);
            }
        }


        
        

        public static void CopyAspNetSessionToAspSession()
        {
            StringBuilder sb = new StringBuilder();
            string stringSeparator = string.Empty;
            List<String> SessionesRestringidas = new List<String>();

            //Preparo las variables de session restringidas
            //Nota:  La variable de session de la forma RHPRO_Home_MenuPrincipal_<modulo> se controla directamente en el if
            SessionesRestringidas.Add("yaentro");
            SessionesRestringidas.Add("Session_ModuloActivo");
            SessionesRestringidas.Add("NroTempLogin");
            SessionesRestringidas.Add("ActualizaAcceso");
            SessionesRestringidas.Add("ActualizaModulo");            
            SessionesRestringidas.Add("ArgTitulo");
            SessionesRestringidas.Add("ArgUrlImagen");
            SessionesRestringidas.Add("primerIdioma");
            SessionesRestringidas.Add("ViendeDeCambiarPassword");
            SessionesRestringidas.Add("VisualizaModulos");
            SessionesRestringidas.Add("PopUpSearchData");
            SessionesRestringidas.Add("ChangeLanguage");
            SessionesRestringidas.Add("login");
            SessionesRestringidas.Add("RHPRO_ListaPerfUsr");
            SessionesRestringidas.Add("RHPRO_LenguajeActivo");
            SessionesRestringidas.Add("RHPRO_LenguajeSeleccionado");
            SessionesRestringidas.Add("RHPRO_Cambio_Estilo");
            SessionesRestringidas.Add("RHPRO_RecienLogueado");
            SessionesRestringidas.Add("RHPRO_NombreModulo");
            SessionesRestringidas.Add("RHPro_PreLoguin");
            SessionesRestringidas.Add("RHPRO_EtiqTraducidasHome");
            SessionesRestringidas.Add("RHPRO_HayTraducciones");
            SessionesRestringidas.Add("RHPRO_AccesosHomeLogin");
            SessionesRestringidas.Add("RHPRO_AccesosHome");
            SessionesRestringidas.Add("RHPRO_Modulos_Habilitados");
            SessionesRestringidas.Add("RHPRO_FefreshSession");
            SessionesRestringidas.Add("RHPRO_Gadgets_Habilitados");

           

            foreach (string key in System.Web.HttpContext.Current.Session.Keys)
            {
                if ((!SessionesRestringidas.Contains(key)) && (!key.Contains("RHPRO_Home_MenuPrincipal_")))//Si no es una variable restringida
                {                    
                    sb.AppendFormat("{0}{1}", stringSeparator, Encryptor.Encrypt("56238", string.Concat(key, "@", System.Web.HttpContext.Current.Session[key])));                    
                    stringSeparator = "_";
                }
            }

                if (ConfigurationManager.AppSettings["NetToAsp"].ToLower() == "true")
                {
                    try
                    {

//                        HttpContext.Current.Response.Redirect(string.Format("~/../CopyAspNetSessionToAspSession.asp?params={0}&returnURL={1}",
                        HttpContext.Current.Response.Redirect(string.Format("~/../CopyAspNetSessionToAspSession.asp?params={0}&returnURL={1}",
                              HttpContext.Current.Server.UrlEncode(sb.ToString()),
                              HttpContext.Current.Server.UrlEncode(HttpContext.Current.Request.Url.AbsolutePath)),false);
                    }
                    catch(Exception ex)
                    {                     
                        ////Cuando se ejecuta en modo debug
                         HttpContext.Current.Response.Redirect(string.Format("Default.aspx"));
                    }  
                }
                      
        
        }

        
   

        public static void CopyAspNetSessionToAspLenguaje()
        {//Solamente paso la variable del lenguaje
            StringBuilder sb = new StringBuilder();
            string stringSeparator = string.Empty;
 
            sb.AppendFormat("{0}{1}", stringSeparator, Encryptor.Encrypt("56238", string.Concat("Lenguaje", "@", System.Web.HttpContext.Current.Session["Lenguaje"])));        
            stringSeparator = "_";

            if (ConfigurationManager.AppSettings["NetToAsp"].ToLower() == "true")
                HttpContext.Current.Response.Redirect(string.Format("~/../CopyAspNetSessionToAspSession.asp?params={0}&returnURL={1}", 
                    HttpContext.Current.Server.UrlEncode(sb.ToString()), 
                    HttpContext.Current.Server.UrlEncode(HttpContext.Current.Request.Url.AbsolutePath)),false);


        }

        public static void CopyAspNetSessionToAspEstilos()
        {//Solamente paso la variable del lenguaje
            StringBuilder sb = new StringBuilder();
            string stringSeparator = string.Empty;            
            foreach (string key in System.Web.HttpContext.Current.Session.Keys)
            {
                if ((key.Contains("EstiloR4_")) || (key == "CarpetaEstilo"))
                {
                    sb.AppendFormat("{0}{1}", stringSeparator, Encryptor.Encrypt("56238", string.Concat(key, "@", System.Web.HttpContext.Current.Session[key])));
                    stringSeparator = "_";
                }
            }

            if (ConfigurationManager.AppSettings["NetToAsp"].ToLower() == "true")
            {
                try
                {
                    HttpContext.Current.Response.Redirect(string.Format("~/../CopyAspNetSessionToAspSession.asp?params={0}&returnURL={1}",
                        HttpContext.Current.Server.UrlEncode(sb.ToString()),
                        HttpContext.Current.Server.UrlEncode(HttpContext.Current.Request.Url.AbsolutePath)), false);
                } 
                catch(Exception ex)
                {
                   
                    ////Cuando se ejecuta en modo debug
                    HttpContext.Current.Response.Redirect(string.Format("Default.aspx"));
                }  

            }
        }

        public static String Tipo_Ordenamiento_Modulos()
        {
            return Convert.ToString(ConfigurationManager.AppSettings["Tipo_Ordenamiento_Modulos"]);
        }

 
         
        public static void SetDefaultSessionValues()
        {
            Session["BaseModule"] = "";
            Utils.SessionUserName = "";

            if (bool.Parse(ConfigurationManager.AppSettings["EnableIntegrateSecurity"]))
            {
                Utils.SessionSeg_NT = Utils.IntegrateSecurityConstants.TrueValue;
            }
            else
            {
                Utils.SessionSeg_NT = Utils.IntegrateSecurityConstants.FalseValue;  
            }
            
            Utils.SessionPassword = "";
            Utils.SessionBaseID = null;
            Utils.SessionTime = DateTime.Now;
            Utils.MaxEmpleados = "";            
        }


        public static void Escribir_Log(string NombreArchivo, string Texto)
        {
            try
            {
                string fileName = System.Web.HttpContext.Current.Server.MapPath("Log/" + NombreArchivo);
                using (System.IO.StreamWriter file = new System.IO.StreamWriter(fileName, true))
                {
                    file.WriteLine(Texto);
                }

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

         


       

 
  
    }
}