using System;
using System.Configuration;
using System.Text;
using System.Web.SessionState;
using System.Web.UI;
using System.Web;


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
 

            CopyAspNetSessionToAspSession();            
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

 
            //foreach (string key in Session.Keys)
            foreach (string key in System.Web.HttpContext.Current.Session.Keys)
            {
                sb.AppendFormat("{0}{1}", stringSeparator, Encryptor.Encrypt("56238", string.Concat(key, "@", System.Web.HttpContext.Current.Session[key])));
                //sb.AppendFormat("{0}{1}", stringSeparator, Encryptor.Encrypt("56238", string.Concat(key, "@", Session[key])));
                //sb.AppendFormat("{0}{1}", stringSeparator, string.Concat(key, "@", Session[key]));
                stringSeparator = "_";    
            }


            if (ConfigurationManager.AppSettings["NetToAsp"].ToLower() == "true")
                HttpContext.Current.Response.Redirect(string.Format("~/../CopyAspNetSessionToAspSession.asp?params={0}&returnURL={1}", HttpContext.Current.Server.UrlEncode(sb.ToString()), HttpContext.Current.Server.UrlEncode(HttpContext.Current.Request.Url.AbsolutePath)));
             
        
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
       
        
    }
}