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
using System.Threading;
using ServicesProxy.rhdesa;
using System.Data;
using System.Text.RegularExpressions;
using System.Web.Security;
 
 
 
namespace RHPro
{
 
        
   
    public partial class Default : System.Web.UI.Page
    {        
        #region Page Handles

        public Lenguaje Obj_Lenguaje;
                //Se define el objeto conexión
        public System.Data.SqlClient.SqlConnection conn;
        public System.Data.SqlClient.SqlDataReader reader;
        public System.Data.SqlClient.SqlCommand sql;
       
        public RHPro.Usuarios CUsuarios;
        public static  bool Control_NET_SessionId = false;
        MetaHome MH_Ext;


        private static readonly string EncryptionKey = ConfigurationManager.AppSettings["EncryptionKey"];
        private static readonly bool EncriptUserData = bool.Parse(ConfigurationManager.AppSettings["EncriptUserData"]);

        private string URL_Meta_Login = ConfigurationManager.AppSettings["URL_Meta_Login"];

        //private void CreateServiceInstance()
        //{
             
        //    //Using localhost requires the page to run on the CRM server box. 
        //    string urlSrv = ConfigurationSettings.AppSettings["urlServidor"];
        //    string crmUsr = ConfigurationSettings.AppSettings["uidAdmin"];
        //    string crmPwd = ConfigurationSettings.AppSettings["pwdAdmin"];
        //    string crmDom = ConfigurationSettings.AppSettings["dominio"];
        //    // Entra con credenciales de administrador para poder crear la obra
        //    System.Net.NetworkCredential credenciales =  new System.Net.NetworkCredential(crmUsr, crmPwd, crmDom);
        //    service.Url = "http://" + urlSrv + "/mscrmservices/2006/crmservice.asmx";
        //    service.Credentials = credenciales;
        //    service.Credentials = System.Net.CredentialCache.DefaultCredentials;

         
        //}

        /// <summary>
        ///  JPB - Actualiza la cookie donde mantiene el id de la aplicacion
        /// </summary>
        public void Actualizar_ASPNET_SessionId()
        { 
            if (Request.Cookies["ASP.NET_SessionId"] != null)
            {               
                Request.Cookies.Remove("ASP.NET_SessionId");
                Request.Cookies.Remove("AuthToken");              
            }
        }

        /// <summary>
        /// Construye un hash con datos especificos desde del cliente
        /// </summary>
        /// <returns></returns>
        private string GenerateHashKey()
        {
            string IP = Request.ServerVariables.Get("REMOTE_ADDR");
            string userName = "";//System.Net.Dns.GetHostEntry(IP).HostName;//Request.ServerVariables.Get("AUTH_USER");           
            String Browser = Request.ServerVariables.Get("HTTP_USER_AGENT");

            System.Text.StringBuilder myStr = new System.Text.StringBuilder();
            //myStr.Append(Browser);
            //myStr.Append(Request.Browser.Platform);
            //myStr.Append(Request.Browser.MajorVersion);
            //myStr.Append(Request.Browser.MinorVersion);
            //myStr.Append(Request.LogonUserIdentity.User.Value);
            myStr.Append(userName);
            myStr.Append(IP);
            System.Security.Cryptography.SHA1 sha = new System.Security.Cryptography.SHA1CryptoServiceProvider();
            byte[] hashdata = sha.ComputeHash(System.Text.Encoding.UTF8.GetBytes(myStr.ToString()));
            return Convert.ToBase64String(hashdata);
        }

        
        public void RedirigirHack()
        {
            Actualizar_ASPNET_SessionId();
            Response.Redirect("InfoHack.aspx");
        }       
         

        private void CONTROLAR_HACKEO()
        {
            if  (Session["AuthToken"] != null)
            {
                string hash = GenerateHashKey();
                
                if (!hash.Equals(Convert.ToString(Session["AuthToken"]))) 
                    RedirigirHack();                            }
            else
            {
                RedirigirHack();
                //Actualizar_ASPNET_SessionId();
                //Response.Redirect("Default.aspx");
            }            
        }


        //META_LOGIN
        /// <summary>
        /// Se conecta con el usuario ess y busca el password del usuario segun el token emitido por el ws_login.
        /// Luego arma el string de conexion con el usuario+pass+base deducidos segun el token de acceso
        /// </summary>
        /// <param name="token"></param>
        private void MetaLogin_Controlar_Token(String token)
        {
            String Usuario = "";
            String Password = "";
            String Base = "";

            MH_Ext = new MetaHome();
            MH_Ext.Iniciar_Ws_Ext();
                        
            String[] DatosAcceso = MH_Ext.get_Data_TokenAcceso(token);
            
            if (DatosAcceso.Length > 0)
            {
                 Usuario = DatosAcceso[0];
                 Base = DatosAcceso[1];
            }
             
            if (Usuario != "" && Base != "")
            {

                string sql = "";
                sql += " select H.husrpass,H.iduser,L.lencod from hist_pass_usr H ";
                sql += " inner join user_per U on U.iduser = H.iduser ";
                sql += " inner join lenguaje L on L.lennro = U.lennro";
                sql += "  where Upper(H.iduser)=Upper('" + Usuario + "') and ((H.hpassfecfin is null)  or ( H.hpassfecfin > GETDATE())) ";
                 

                try
                {
                    Consultas cc = new Consultas();
                    DataTable infoUsr = cc.get_DataTable(sql, Base);

                    if (infoUsr.Rows.Count > 0)
                    {
                        if (!String.IsNullOrEmpty(Convert.ToString(infoUsr.Rows[0]["husrpass"])))
                        {
                            Password = Common.Encryptor.Decrypt(EncryptionKey, Convert.ToString(infoUsr.Rows[0]["husrpass"]));

                            Entities.Login LoginToReturn = new Entities.Login
                            {
                                IsValid = true,
                                Messege = "",
                                RequiredChangePassword = false,
                                Lenguaje = Convert.ToString(infoUsr.Rows[0]["lencod"]),
                                MaxEmpl = "1",
                            };

                            Conexion con = new Conexion();
                            string guid = GenerateHashKey();

                            con.Iniciar_Sesion(LoginToReturn, Base, Usuario, Password, guid, get_NombreBase(Base), Base);
                        }
                        else
                        {
                            Response.Write("<script>alert('" + Obj_Lenguaje.Label_Home("El usuario " + Usuario + " no tiene password configurado en el cliente") + "');</script>");
                            Response.Write("<script>this.location = '" + URL_Meta_Login + "';</script>");
                            //Response.Redirect(URL_Meta_Login, true);  
                        }
                    }
                    else {
                        Response.Write("<script>alert('" + Obj_Lenguaje.Label_Home("El usuario " + Usuario + " no tiene password configurado en el cliente") + "');</script>");
                        Response.Write("<script>this.location = '" + URL_Meta_Login + "';</script>");
                    }
                }
                catch (Exception ex)
                {
                    Response.Write("<script>alert('" + ex.Message.Replace("'", "")+ "');</script>");
                    Response.Write("<script>this.location = '" + URL_Meta_Login + "';</script>");
                    //Response.Redirect(URL_Meta_Login, true);  
                }
            }
            else
            {
                Response.Write("<script>alert('" + Obj_Lenguaje.Label_Home("Credenciales de acceso vencidas") + "');</script>");
                Response.Write("<script>this.location = '" + URL_Meta_Login + "';</script>"); 
            }
            
        }







        public void Iniciar_Sesion(Entities.Login login, String SelectedDatabaseId, String User, String Pass)
        {
            /**********************************************************/
            /* Se genera el token de seguridad  de hackeo             */
            string guid = GenerateHashKey();
            Session["AuthToken"] = guid;
            //Response.Cookies.Add(new HttpCookie("AuthToken", guid));
            /**********************************************************/


            //EtiqLenguaje = (String)System.Web.HttpContext.Current.Session["Lenguaje"];
            //jpb: Cambio el idioma del home por el configurado para el usuario logeado

            //Recupero en objeto login
            Session["login"] = login;

            String EtiqLenguaje = login.Lenguaje;
            Consultas cc = new Consultas();
            if ((Utils.SessionNroTempLogin == null) || ((String)Utils.SessionNroTempLogin == ""))
            {
                Common.Utils.SessionBaseID = SelectedDatabaseId;

                //jpb -Actualizo el nombre del InitialCatalog de la conexion
                //Consultas cc1 = new Consultas();
                //Paso las credenciales al web service
                cc.Credentials = System.Net.CredentialCache.DefaultCredentials;

                Session["InitialCatalog"] = cc.Initial_Catalog(SelectedDatabaseId);

            }
            else
            {
                //jpb -Actualizo el nombre del InitialCatalog de la conexion
                //Consultas cc2 = new Consultas();
                //Paso las credenciales al web service
                cc.Credentials = System.Net.CredentialCache.DefaultCredentials;
                //-----------------------------------------------------------
                Session["InitialCatalog"] = cc.Initial_Catalog(Common.Utils.SessionBaseID);

            }

            //Variables de session que utiliza el metahome
            HttpContext context = HttpContext.Current;

            string url = HttpContext.Current.Request.Url.AbsoluteUri;
            Session["RHPRO_URL_Login"] = url;
            Session["RHPRO_IP_Login"] = Server.HtmlEncode(Request.UserHostAddress);
             
            Session["RHPRO_Browser_Type"] = Request.UserAgent;
           
               
            Utils.LoginUser(User, Pass, EncriptUserData, EncryptionKey, EtiqLenguaje, login.MaxEmpl);

            /*********************************/
            //Una vez logueado controlo si tengo que asociar gadget al usuario
            if (Convert.ToBoolean(ConfigurationManager.AppSettings["Controlar_Primer_Acceso"]))
            {
                cc.Controlar_Gadget_EnLoguin(User, Pass, "false", SelectedDatabaseId);
            }
            /*********************************/

            //Guardo el webservices utilizado

            String[] Ws_Actual = Regex.Split(Convert.ToString(ConfigurationManager.AppSettings["RootWS"]), "//");
            Ws_Actual = Regex.Split(Ws_Actual[1], "/");
            Session["RHPRO_WS"] = Ws_Actual[Ws_Actual.Length - 2];
             
            /*********************************/
            //Usuarios CUsuarios = new Usuarios();
            //CUsuarios.Inicializar_Estilos(true);
            Inicializar_Estilos(true);
            
            Session["RHPRO_NombreModulo"] = "RHPROX2";
            /**********************************/


            //Vuelvo RHPRO_HayTraducciones en vacio para que resuelva las traducciones con el lenguaje del usuario
            System.Web.HttpContext.Current.Session["RHPRO_EtiqTraducidasHome"] = "";
            System.Web.HttpContext.Current.Session["RHPRO_HayTraducciones"] = "";


            /**********************************/
            //Guarda el string de conexion para pasarselo al conn_db          
            /**********************************/
            Utils.CopyAspNetSessionToAspSession();
            /********************************/
            //Cambio el menu Login
         

            //if (Utils.IsUserLogin)
            //{

            //    Session["Username"] = User;
            //    Session["Password"] = Pass;
            //}

             

            Session["RHPro_PreLoguin"] = "0";

        }


        //public void Tiempo_Ejec(String mje)
        //{
        //    try
        //    {
        //        ///* -sacar ---------------------------------*/
        //        ///
        //        mje = mje.Replace("<br>", " ############ ");
        //        mje = mje.Replace("\\n", " ############ ");

        //        //Consultas cc = new Consultas();
        //        //DataTable dt = cc.get_DataTable("select cnstring from conexion order by cnnro ASC ", Utils.SessionBaseID);

        //        System.Data.OleDb.OleDbConnection cn3 = new System.Data.OleDb.OleDbConnection();
        //        //cn3.ConnectionString = Convert.ToString(dt.Rows[0]["cnstring"]); // "Provider=SQLOLEDB.1;Password=ess;User ID=ess;Data Source=RHDESA;Initial Catalog=BASE_0_R3_ARG;";
        //        cn3.ConnectionString = "Provider=SQLOLEDB.1;Password=6852593102166269536E;User ID=raetlatam;Persist Security Info=False;Data Source=SD-P-RHPAPP01;Initial Catalog=HEID_AR_BA0_T1";
        //        cn3.Open();
        //        string sqlSS3 = "insert into sacar (sql) values ('" + mje.Replace("'", "##") + "')";
        //        System.Data.OleDb.OleDbCommand cmd = new System.Data.OleDb.OleDbCommand(sqlSS3, cn3);
        //        cmd.ExecuteNonQuery();
        //        ///*-----------------------------------*/
        //    }
        //    catch (Exception e) { }
        //}
 
        protected void Page_Load(object sender, EventArgs e)
        {


            //--SACAR----------------------------------------------------------------------------------
            //System.Diagnostics.Stopwatch tiempoprueba = System.Diagnostics.Stopwatch.StartNew();
            //string MjeTiempo = "";
            //--SACAR----------------------------------------------------------------------------------
                        
            //if (Response.Cookies.Count > 0)
            //{
            //    foreach (string s in Response.Cookies.AllKeys)
            //    {
            //        Response.Cookies[s].Secure = true;
            //    }
            //}           
             
            
            Obj_Lenguaje = new Lenguaje();           

            if (Utils.IsUserLogin)
            {               
                CONTROLAR_HACKEO();                
            }

           

            //META_LOGIN-------------------
            if (!Utils.IsUserLogin)
            {                
                MetaHome MH = new MetaHome();
                
                if (MH.MetaHome_Activo() && (MH.MetaHome_TipoFiltroLogin() == "3"))//Si proviene del MetaLogin y esta activo SaaS                   
                {                        
                    Btn_Login_MenuTop.Visible = false;

                    if (!String.IsNullOrEmpty(Request["token"]))                                                    
                        MetaLogin_Controlar_Token(Request["token"]);                         
                    else                             
                        Response.Redirect(URL_Meta_Login,true);                                                                                                          
                }                
            }
            //----------------------------
                         
          
            if (System.Web.HttpContext.Current.Session["yaentro"] == null)
                System.Web.HttpContext.Current.Session["yaentro"] = false;
            if (System.Web.HttpContext.Current.Session["primerIdioma"] == null)
                System.Web.HttpContext.Current.Session["primerIdioma"] = false;
                        
            //Inicializo el lenguaje dejault           
            if (!Utils.IsUserLogin)
            {         
        
                //Cargo el string de conexion por defecto      
                SetBaseIdDefault();

                
                //LoadConexionDefault(); //JPB: Se elimina
                if (!(bool)System.Web.HttpContext.Current.Session["primerIdioma"])
                {
                    System.Web.HttpContext.Current.Session["primerIdioma"] = true;
                    string[] LengDefault = Obj_Lenguaje.Lenguaje_Default().Split(',');
                    string Idioma;
                    //string bandera = LengDefault[0];

                    //string IdiomaDefault = LengDefault[0].Substring(0, 2) + "-" + LengDefault[0].Substring(2, 2);
                    if (LengDefault.Length > 0)
                        Idioma = LengDefault[1];
                    else
                        Idioma = "es-AR";

                    //Common.Utils.Lenguaje = IdiomaDefault;
                    Common.Utils.Lenguaje = Obj_Lenguaje.Etiq_Leng_Default();
                    string bandera = Obj_Lenguaje.Etiq_Leng_Default().Replace("-", "");
                    System.Web.HttpContext.Current.Session["ArgTitulo"] = Idioma;
                    System.Web.HttpContext.Current.Session["ArgUrlImagen"] = "~/img/Flags/flag_" + bandera + ".png";
                }
                            

            }
            else
            {
                //TitBtnFavoritos.InnerHtml = Obj_Lenguaje.Label_Home("Favoritos");
                //TitBtnGadget.InnerText = Obj_Lenguaje.Label_Home("Gadgets"); 
                 
                if (!(bool)System.Web.HttpContext.Current.Session["yaentro"])
                {

                    System.Web.HttpContext.Current.Session["yaentro"] = true;
                    string UsuarioLogueado = Common.Utils.SessionUserName;
                    string[] LengDefault = Obj_Lenguaje.Lenguaje_Usuario(UsuarioLogueado).Split(',');

                    string bandera = LengDefault[0];
                    string IdiomaDefault = LengDefault[0].Substring(0, 2) + "-" + LengDefault[0].Substring(2, 2);
                    string Idioma = LengDefault[1];

                    Common.Utils.Lenguaje = IdiomaDefault;

                    System.Web.HttpContext.Current.Session["ArgTitulo"] = Idioma;
                    System.Web.HttpContext.Current.Session["ArgUrlImagen"] = "~/img/Flags/flag_" + bandera + ".png";



                }
                
                //------------------------------------------------
                //Controlo dias vencidos del password
                Consultas consultas = new Consultas();
                //Paso las credenciales al web service
                //consultas.Credentials = System.Net.CredentialCache.DefaultCredentials;
                //-----------------------------------------------------------
               
 
                String DiasVence = consultas.Info_Dias_VencimientoPass(Common.Utils.SessionUserName, Utils.SessionBaseID);               
                
                if (!String.IsNullOrEmpty(DiasVence))
                {
                    String Info = "<img src='Images/atencion.png' align='absmiddle' height='23'> " + Obj_Lenguaje.Label_Home("Su contraseña expira en @@NUM1@@ días");
                    Info = Info.Replace("@@NUM1@@", DiasVence);
                    InfoCambioPass.Controls.Add(new LiteralControl(Info));
                    LinkButton LB = new LinkButton();
                    LB.ID = "CambioDePassword";
                    LB.Text = Obj_Lenguaje.Label_Home("Cambiar Contraseña");
                    LB.CssClass = "Boton";
                    LB.Click += new EventHandler(this.PopUp_Cambiar_Password);
                    LB.DataBind();
                    InfoCambioPass.Controls.Add(LB);
                    InfoCambioPass.DataBind();
                }

                bool Viene = false;
                if (!String.IsNullOrEmpty((String)Session["ViendeDeCambiarPassword"]))
                    Viene = true;

                if (Viene)
                {
                    Session["ViendeDeCambiarPassword"] = null;
                    Utils.LogoutUser();
                }
                


               
            }

            cLogin.UserLogin += cLogin_UserLogin;
            cLogin.UserLogout += cLogin_UserLogout;
            //                Modulos.AsignarContPpal(ContenedorPrincipal);


            Cargar_Modulos();           
            Cargar_Datos_De_Usuario();   
            Cargar_Datos_De_Lenguaje();
            Cargar_Datos_De_Politicas_Top();
             
            if (Utils.IsUserLogin)
            {                 
                Cargar_Datos_De_Favoritos();                 
                Cargar_Datos_De_Complementos();                
                Cargar_Datos_De_Estilos();                
            }


         
            if (bool.Parse(ConfigurationManager.AppSettings["VisualizarFooter"]))
                Cargar_Info_Piso();             

             Cargar_Estilo_De_Pagina();
             

            if (!Utils.IsUserLogin)
            {
                 Btn_Login_MenuTop.Attributes.Add("onclick", "setTimeout(\"document.aspnetForm.ctl00$content$cLogin$txtUserName.focus();\", 30);");
                //Btn_Login_MenuTop.Attributes.Add("onmouseover", "setTimeout(\"document.aspnetForm.ctl00$content$cLogin$txtUserName.focus();\", 30);");               
            }
            else
            {
                Btn_Login_MenuTop.Attributes.Add("onclick", "setTimeout(\"document.getElementById('ctl00_content_cLogin_CerrarSesion').focus();\", 30);");
                //Btn_Login_MenuTop.Attributes.Add("onmouseover", "setTimeout(\"document.getElementById('ctl00_content_cLogin_CerrarSesion').focus();\", 30);");
            }

            
            //Le paso una referencia de la clase Default a ciertos controles
            ContenedorPrincipal.InicializarPadre(this);
            cIdiomas.InicializarPadre(this);
            ConfigGadgets.InicializarPadre(this);
            
          
        }

        //public void Actualizar_Loguin(){
        //    cLogin.Redibujar_Campos_y_Botones(false,true);
        //}

        /// <summary>
        /// Verifica si un boton esta habilitado en un fuente del home
        /// </summary>
        /// <param name="boton"></param>
        /// <param name="fuente"></param>
        /// <returns></returns>
         public bool Habilitado(String boton,string fuente)
         {
             ConfiguracionesHome ch = new ConfiguracionesHome();
             return ch.Habilitado(boton, fuente);
         }

 
         /// <summary>
         /// Verifica si los perfiles de un gadget habilitan a un usuario dado
         /// </summary>
         /// <param name="User"></param>
         /// <param name="Gadnro"></param>
         /// <returns></returns>
         public bool  Gadget_Habilitado(String User, String Gadnro)
         {
             Consultas cc = new Consultas();
             DataTable dtUser = cc.get_DataTable("SELECT listperfnro FROM user_perfil WHERE Upper(iduser)=Upper('" + User + "')", Utils.SessionBaseID);             
             String lista = "";
             String gadperfil = "";
             if (dtUser.Rows.Count > 0)
             {
                 lista = Convert.ToString(dtUser.Rows[0]["listperfnro"]);
                 if (lista == "*")//El usuario ve todo
                     return true;
                 else
                 {
                     DataTable dtGadget = cc.get_DataTable("SELECT gadperfil FROM Gadgets_Perfil where gadnro=" + Gadnro + "   ", Utils.SessionBaseID);

                     foreach (DataRow dr in dtGadget.Rows)
                     {
                         gadperfil = Convert.ToString(dr["gadperfil"]);
                         String[] Misplit = lista.Split(',');
                         foreach (String PerfUsr in Misplit)
                         {
                             if ((PerfUsr == gadperfil) || (PerfUsr == "*"))
                                 {
                                     return true;
                                 }                            
                         }
                     }
                 }
             }

             return false;

         }

         /// <summary>
         /// Verifica si los perfiles de un gadget habilitan a un usuario dado
         /// </summary>
         /// <param name="User"></param>
         /// <param name="Gadnro"></param>
         /// <returns></returns>
         public string Lista_User_Perfil(String User)
         {
             Consultas cc = new Consultas();
             DataTable dtUser = cc.get_DataTable("SELECT listperfnro FROM user_perfil WHERE Upper(iduser)=Upper('" + User + "')", Utils.SessionBaseID);
             String lista = "";   
             if (dtUser.Rows.Count > 0)
             {
                 lista = Convert.ToString(dtUser.Rows[0]["listperfnro"]);                 
             }

             return lista;

         }

         public string Lista_Gadget_Permitidos(String User)
         {
             Consultas cc = new Consultas();             
             String listperfnro = Lista_User_Perfil(User);
             String lista = "";
             String[] Misplit = listperfnro.Split(',');
             DataTable dtGadget = cc.get_DataTable("SELECT DISTINCT gadnro,gadperfil FROM Gadgets_Perfil ", Utils.SessionBaseID);
            
             foreach (DataRow dr in dtGadget.Rows)
             {


                 if (Misplit.Contains(Convert.ToString(dr["gadperfil"])) || (Misplit.Contains("*")))
                 {
                     if (lista == "")
                         lista = Convert.ToString(dr["gadnro"]);
                     else
                         lista += "," + Convert.ToString(dr["gadnro"]);
                 }
             }

              

             return lista;

         }

         public void Pasar_Estilos_ASP()
         {
             System.Text.StringBuilder sb = new System.Text.StringBuilder();
             string stringSeparator = string.Empty;
             foreach (string key in System.Web.HttpContext.Current.Session.Keys)
             {
                 if ((key.Contains("EstiloR4_")) || (key == "CarpetaEstilo"))
                 {
                     sb.AppendFormat("{0}{1}", stringSeparator, Encryptor.Encrypt("56238", string.Concat(key, "@", System.Web.HttpContext.Current.Session[key])));
                     stringSeparator = "_";
                 }
             }

             ifrmEst.Attributes.Add("location", string.Format("~/../CopyAspNetSessionToAspSession.asp?params={0}&returnURL={1}", HttpContext.Current.Server.UrlEncode(sb.ToString()), HttpContext.Current.Server.UrlEncode(HttpContext.Current.Request.Url.AbsolutePath)));
         }


         public void Inicializar_Estilos(Boolean Logueado)
         {

             String sql = " SELECT * FROM estilo_homex2 X2 ";
             sql += " inner join estilos_home H On H.idestilo = X2.idcarpetaestilo  ";

             if (Logueado)
             {
                 sql += " WHERE codestilo =(select U2.estiloactivo from estilos_home_user U2 where Upper(U2.iduser)=Upper('" + Utils.SessionUserName + "')  )  ";
             }
             else
             {
                 sql += " WHERE defecto = -1 ";
             }


             Consultas cc = new Consultas();
             //Paso las credenciales al web service
             //cc.Credentials = System.Net.CredentialCache.DefaultCredentials;

             DataTable dt = cc.get_DataTable(sql, Utils.SessionBaseID);

             if (dt.Rows.Count > 0)
             { 
                 Session["EstiloR4_coloricono"] = Convert.ToString(dt.Rows[0]["coloricono"]);
                 Session["EstiloR4_fondoCabecera"] = Convert.ToString(dt.Rows[0]["fondoCabecera"]);
                 Session["EstiloR4_fuentePiso"] = Convert.ToString(dt.Rows[0]["fuentePiso"]);
                 Session["EstiloR4_fondoFecha"] = Convert.ToString(dt.Rows[0]["fondoFecha"]);
                 Session["EstiloR4_fuenteFecha"] = Convert.ToString(dt.Rows[0]["fuenteFecha"]);
                 Session["EstiloR4_fondoModulos"] = Convert.ToString(dt.Rows[0]["fondoModulos"]);
                 Session["EstiloR4_fuenteModulos"] = Convert.ToString(dt.Rows[0]["fuenteModulos"]);
                 Session["EstiloR4_fondoPiso"] = Convert.ToString(dt.Rows[0]["fondoPiso"]);
                 Session["EstiloR4_coloriconomenutop"] = Convert.ToString(dt.Rows[0]["coloriconomenutop"]);
                 Session["EstiloR4_fondocontppal"] = Convert.ToString(dt.Rows[0]["fondocontppal"]);

                 Session["EstiloR4_fondoGadget"] = Convert.ToString(dt.Rows[0]["fondogadget"]);
                 Session["EstiloR4_fuenteGadget"] = Convert.ToString(dt.Rows[0]["fuentegadget"]);


                 if (!dt.Rows[0]["logoEstilo"].Equals(System.DBNull.Value))
                     Session["EstiloR4_logoEstilo"] = dt.Rows[0]["logoEstilo"];
                 else
                     Session["EstiloR4_logoEstilo"] = ConfigurationManager.AppSettings["urlLogo"];

                 Session["CarpetaEstilo"] = dt.Rows[0]["estilocarpeta"];

             }
             else
             {
                 sql = " SELECT * FROM estilo_homex2 X2 ";
                 sql += " inner join estilos_home H On H.idestilo = X2.idcarpetaestilo  ";
                 sql += " WHERE defecto = -1 ";
                 dt = cc.get_DataTable(sql, Utils.SessionBaseID);
                 if (dt.Rows.Count > 0)
                 {
                     Session["EstiloR4_coloricono"] = Convert.ToString(dt.Rows[0]["coloricono"]);
                     Session["EstiloR4_fondoCabecera"] = Convert.ToString(dt.Rows[0]["fondoCabecera"]);
                     Session["EstiloR4_fuentePiso"] = Convert.ToString(dt.Rows[0]["fuentePiso"]);
                     Session["EstiloR4_fondoFecha"] = Convert.ToString(dt.Rows[0]["fondoFecha"]);
                     Session["EstiloR4_fuenteFecha"] = Convert.ToString(dt.Rows[0]["fuenteFecha"]);
                     Session["EstiloR4_fuenteModulos"] = Convert.ToString(dt.Rows[0]["fuenteModulos"]);
                     Session["EstiloR4_fondoPiso"] = Convert.ToString(dt.Rows[0]["fondoPiso"]);
                     Session["EstiloR4_coloriconomenutop"] = Convert.ToString(dt.Rows[0]["coloriconomenutop"]);
                     Session["EstiloR4_fondocontppal"] = Convert.ToString(dt.Rows[0]["fondocontppal"]);
                     Session["EstiloR4_fondoGadget"] = Convert.ToString(dt.Rows[0]["fondogadget"]);
                     Session["EstiloR4_fuenteGadget"] = Convert.ToString(dt.Rows[0]["fuentegadget"]);

                     if (!dt.Rows[0]["logoEstilo"].Equals(System.DBNull.Value))
                         Session["EstiloR4_logoEstilo"] = dt.Rows[0]["logoEstilo"];
                     else
                         Session["EstiloR4_logoEstilo"] = ConfigurationManager.AppSettings["urlLogo"];

                     Session["CarpetaEstilo"] = dt.Rows[0]["estilocarpeta"];
                 }
             }
         } 


         public void Cargar_Estilo_De_Pagina()
         {
             if (Utils.SesionIniciada)
             {
                // if ((Convert.ToString(Session["RHPRO_Cambio_Estilo"]) == "-1") || (Convert.ToString(Session["RHPRO_RecienLogueado"]) == "-1"))                
                 if  ( Convert.ToString(Session["RHPRO_Cambio_Estilo"]) == "-1" )             
                 {                    
                     Session["RHPRO_Cambio_Estilo"] = "0";
                     Session["RHPRO_RecienLogueado"] = "0";
                     //CUsuarios = new Usuarios();
                     //CUsuarios.Inicializar_Estilos(true);
                     Inicializar_Estilos(true);
                      
                     Utils.CopyAspNetSessionToAspEstilos();
                    // Pasar_Estilos_ASP();
                  }                
                 
             }
             else
             {
              //   if (Convert.ToString(System.Web.HttpContext.Current.Session["EstiloR4_coloricono"]) == "")
                 //if (Convert.ToString(Session["EstiloR4_coloricono"]) == "")
                 //{

                 //CUsuarios = new Usuarios();
                 //CUsuarios.Inicializar_Estilos(false);
                 Inicializar_Estilos(false);                                          
                    // Pasar_Estilos_ASP();
                 //}
             }
             
             Logo_Empresa.Attributes.Add("src", (String)System.Web.HttpContext.Current.Session["EstiloR4_logoEstilo"]);
             
         }

         

         public void Cargar_Redes_Sociales()
         {
             //RedesSociales
             String icono = "";
             String NombreIcono = "";

             Consultas cc = new Consultas();
             ////Paso las credenciales al web service
             //cc.Credentials = System.Net.CredentialCache.DefaultCredentials;
             String sql = " SELECT  hredtitulo, hredpagina, paisnro, rhpro, ess, icono FROM home_redes WHERE rhpro=-1";
             DataTable dt = cc.get_DataTable(sql, Utils.SessionBaseID);
             String accion = "";
             foreach (DataRow dr in dt.Rows)
             {
                // accion = "abrirVentana('" + (String)dr["hredpagina"] + "','Red','800','600')";
                 accion = "window.open('" + (String)dr["hredpagina"] + "','_blank','location=yes, toolbar=yes, scrollbars=yes, resizable=yes, width=800, height=600', top=50)";
             
                 NombreIcono = (String)dr["icono"];
                 //icono += "<img  src='img/REDES/" + NombreIcono + "' border='0' style='cursor: pointer;' title='" + (String)dr["hredtitulo"] + "' class='IconoREDES' onclick=\"" + accion + "\">";
                 icono += Utils.Armar_Icono("img/REDES/" + NombreIcono , "IconoREDES", (String)dr["hredtitulo"], "' border='0' style='cursor: pointer;'" , "",accion);
                 
             }
            
             RedesSociales.Controls.Add(new LiteralControl(icono));
         }
		  
          public void Cargar_Info_Piso()
         {
             NombreDeEmpresa.InnerText = (String)ConfigurationManager.AppSettings["NombreDeEmpresa"];//"Heidt & Asociados S.A.";
             DireccionEmpresa.InnerText = (String)ConfigurationManager.AppSettings["DireccionEmpresa"];//"Suipacha 72 - 4º A CP C1008AAB - Buenos Aires Argentina.";
             TelefonoMailEmpresa.InnerText = (String)ConfigurationManager.AppSettings["TelefonoMailEmpresa"];//"Tel./Fax: +54 11 5252 7300 Email: ventas@rhpro.com ";
             Slogan.InnerText = (String)ConfigurationManager.AppSettings["SloganCliente"];
             //versionMI.InnerText = VersionServiceProxy.Find(Utils.SessionBaseID, Thread.CurrentThread.CurrentCulture.Name);//carga la version
             //patchMI.InnerText = PatcheServiceProxy.Find(Utils.SessionBaseID, Thread.CurrentThread.CurrentCulture.Name);//carga el patch
             versionMI.InnerText = Obj_Lenguaje.Label_Home("Versión") + ": " + VersionServiceProxy.Find(Utils.SessionBaseID,Utils.Lenguaje);//carga la version
             patchMI.InnerText = Obj_Lenguaje.Label_Home("Patch") + ": " + PatcheServiceProxy.Find(Utils.SessionBaseID, Utils.Lenguaje);//carga el patch
             Cargar_Redes_Sociales();

             String accion = "window.open('http://www.rhpro.com/','_blank','location=yes, toolbar=yes, scrollbars=yes, resizable=yes, width=1000, height=600', top=50)";

             URL_Logo.Controls.Add(new LiteralControl("<img onclick=\"" + accion + "\" src='img/Logo_RHPro.png' width='110' style='cursor:pointer;margin-bottom:4px;' >"));
              
             
         }
          

         public void Cargar_Datos_De_Lenguaje()
         {
             
             Info_Lenguaje.Controls.Clear();
             Info_Lenguaje.Controls.Add(new LiteralControl(Utils.Armar_Icono("img/Modulos/SVG/IDIOMA.svg", "IconosBarraTop", Obj_Lenguaje.Label_Home("Idioma"), " border='0'", "") + System.Web.HttpContext.Current.Session["ArgTitulo"]));             
         }

         public void Cargar_Datos_De_Politicas_Top()
         {
             Info_Politicas.Controls.Clear();
             //Info_Politicas.Controls.Add(new LiteralControl("<img src='img/Modulos/SVG/POLITICAS_TOP.svg' border='0' class='IconosBarraTop' > " +  Obj_Lenguaje.Label_Home("Politicas")));
             Info_Politicas.Controls.Add(new LiteralControl(Utils.Armar_Icono("img/Modulos/SVG/POLITICAS_TOP.svg", "IconosBarraTop", Obj_Lenguaje.Label_Home("Politicas"), " border='0'", "") + Obj_Lenguaje.Label_Home("Politicas")));
             
         }

         public void Cargar_Modulos()
         {
             Controls.Modulos NuevoControl = (Controls.Modulos)Page.LoadControl("Controls/Modulos.ascx");
             CModulos.Controls.Add(NuevoControl);
             ((Controls.Modulos)NuevoControl).AsignarContPpal(ContenedorPrincipal);
            
         }

         public void Cargar_Datos_De_Favoritos()
         {
             Controls.Contenedor_Favoritos NuevoControl = (Controls.Contenedor_Favoritos)Page.LoadControl("Controls/Contenedor_Favoritos.ascx");
             //NuevoControl.Refrescar(this, new EventArgs());
             //NuevoControl.DataBind(); 
             //try
             //{
             //    NuevoControl.Imprimir_Favoritos();
             //}catch(Exception ex){                
             //}                         

             CFavoritos2.Controls.Add(NuevoControl);
 
             Info_Favoritos.Controls.Clear();
             //Info_Favoritos.Controls.Add(new LiteralControl("<img src='img/Modulos/SVG/FAVORITO_TOP.svg' border='0' class='IconosBarraTop' > " + Obj_Lenguaje.Label_Home("Favoritos")));
             Info_Favoritos.Controls.Add(new LiteralControl(Utils.Armar_Icono("img/Modulos/SVG/FAVORITO_TOP.svg", "IconosBarraTop", "", " border='0'", "") + Obj_Lenguaje.Label_Home("Favoritos")));
             
         }

         public void Cargar_Datos_De_Complementos()
         {

             Info_Complementos.Controls.Clear();
//             Info_Complementos.Controls.Add(new LiteralControl("<img src='img/Modulos/SVG/GADGET.svg' border='0' class='IconosBarraTop'  > " + Obj_Lenguaje.Label_Home("Gadgets")));
             Info_Complementos.Controls.Add(new LiteralControl(Utils.Armar_Icono("img/Modulos/SVG/GADGET.svg", "IconosBarraTop", "", " border='0'", "") + Obj_Lenguaje.Label_Home("Gadgets")));
              
         }

         public void Cargar_Datos_De_Estilos()
         {
             Info_Estilos.Controls.Clear(); 
//             Info_Estilos.Controls.Add(new LiteralControl("<img src='img/Modulos/SVG/ESTILOS.svg' border='0' class='IconosBarraTop'  > " + Obj_Lenguaje.Label_Home("Estilos")));
             Info_Estilos.Controls.Add(new LiteralControl(Utils.Armar_Icono("img/Modulos/SVG/ESTILOS.svg", "IconosBarraTop", "", " border='0'", "") + Obj_Lenguaje.Label_Home("Estilos")));
             
         }


        
         public void Cargar_Datos_De_Usuario()
         {
             Info_Usuario.Controls.Clear();
            
               String Salida = "";                 
               String imagen = Utils.Armar_Icono("img/Modulos/SVG/USER.svg", "IconosBarraTop", "", " border='0'", "");             

                 if (Common.Utils.IsUserLogin)
                 {
                     String sql = "";
                     Consultas cc = new Consultas();                      
                     
                     sql = "SELECT E.terape, E.terape2, E.ternom, E.ternom2,tipimdire, tipimanchodef, tipimaltodef, terimnombre, ter_imag.terimfecha FROM ter_imag  ";
                     sql += " LEFT JOIN tipoimag ON tipoimag.tipimnro = ter_imag.tipimnro   ";
                     sql += " INNER JOIN empleado E ON E.ternro = ter_imag.ternro ";
                     sql += " WHERE ter_imag.ternro = (select ternro from user_ter where iduser='" + Utils.SessionUserName + "'	)	 AND ter_imag.tipimnro = 3  ";
                     sql += " ORDER BY ter_imag.terimfecha DESC ";

                     DataTable dt = cc.get_DataTable(sql, Utils.SessionBaseID);
                     if (dt.Rows.Count > 0)
                     {
                         if (dt.Rows[0]["tipimdire"] != null)
                             imagen = "<img src='" + dt.Rows[0]["tipimdire"] + dt.Rows[0]["terimnombre"] + "' border='0' class='IconoBarraTopUser' align='absmiddle'    > ";

                         Salida = "<table  border='0' cellspacing='0' cellpadding='0' class='Info_Log_Usr' height='10'>";
                         Salida += "  <tr>";
                         Salida += "    <td valign='middle' align='center'   rowspan='2' >" + imagen + "</td>";
                         Salida += "    <td nowrap valign='middle' align='left' ><span class='Log_Nombre'> " + dt.Rows[0]["ternom"] + " " + dt.Rows[0]["terape"] + "</span></td>";
                         Salida += "  </tr>";
                         Salida += "  <tr>";
                         Salida += "    <td nowrap  valign='middle' align='left'  > <span class='Log_User'> " + Common.Utils.SessionUserName + "</span></td>";
                         Salida += "  </tr>";
                         Salida += "</table> ";

                     }
                     else
                         Salida = imagen + " " + Common.Utils.SessionUserName;

                 }
                 else
                     Salida = imagen + " " + Obj_Lenguaje.Label_Home("Ingresar");


                  Info_Usuario.Controls.Add(new LiteralControl(Salida));
             
         }

         public void Abrir_Ventana_Generica(object sender, CommandEventArgs e)
        {
            String urlControl = (String)e.CommandArgument;
            //PopUp_Generico.Attributes.Add("style", "display:inline");
            //Contenedor_Ventana_Generica.Attributes.Add("style", "display:inline");
            //urlControl = "~/Controls/" + urlControl;                         
            //Control GadgetControl = (Control)Page.LoadControl(urlControl);
            ////Panel_Generico.Controls.Add(new LiteralControl("<DIV ID='PopUp_Generico'  Class='PopUp_FondoTransparente'></DIV>"));
            //Panel_Generico.Controls.Add(GadgetControl);
            //Panel_Generico.DataBind();
            
        }

         public void Cerrar_Ventana_Generica(object sender, CommandEventArgs e)
         {
           
             //PopUp_Generico.Attributes.Add("style", "display:none");
             //Contenedor_Ventana_Generica.Attributes.Add("style", "display:none");                         
             Panel_Generico.Controls.Clear();
             Panel_Generico.DataBind();

         }

         public void Visualizar_Gadgets(object sender, EventArgs e)
         {
             Session["RHPRO_NombreModulo"] = "RHPROX2";
             Utils.Session_ModuloActivo = "RHPROX2";
             //Response.Redirect("Default.aspx");
             ContenedorPrincipal.Update_Gadget(1);
             //Habilito el armado de los submenues
            // ScriptManager.RegisterStartupScript(this, typeof(Page), "Logo_InicializaMenu", "$(function() {  $('#main-menu').smartmenus({  subMenusSubOffsetX: 0, subMenusSubOffsetY: -1  }); });  ", true);
            // ScriptManager.RegisterStartupScript(this, typeof(Page), "Logo_InicializaMenuTop", "$(function() {  $('#main-menuTop').smartmenus({  subMenusSubOffsetX: 0, subMenusSubOffsetY: -1, mainMenuSubOffsetX:0,mainMenuSubOffsetY:0,subMenusMinWidth:'60px', subMenusMaxWidth:'900px'  }); });  ", true);
             //Habilito el armado de los submenues
             ScriptManager.RegisterStartupScript(this, typeof(Page), "InicializaMenu", "$(function() {  $('#main-menu').smartmenus({  subMenusSubOffsetX: 0, subMenusSubOffsetY: -1 , hideOnClick: false  }); });  ", true);
             ScriptManager.RegisterStartupScript(this, typeof(Page), "InicializaMenuTop", "$(function() {  $('#main-menuTopLoguin').smartmenus({  subMenusSubOffsetX: 0, subMenusSubOffsetY: -1, mainMenuSubOffsetX:0,mainMenuSubOffsetY:0,subMenusMinWidth:'60px', subMenusMaxWidth:'1060px', hideOnClick: true   }); });  ", true);
         }


         public void PopUp_Cambiar_Password(object sender, EventArgs e)
         {
           

             // Disparas popup para que cambie el pass con el mensaje  y carga en el session los datos del logueo
             DataBase BD = new DataBase();
             BD.Id = Utils.SessionBaseID;
             BD.Name = Utils.SessionBaseID;
             PopUpChangePassData popUpChangePassData = new PopUpChangePassData
             {
                 Login =  (Entities.Login) Session["login"],
                 UserName = Common.Utils.SessionUserName,
                 DataBase = BD// SelectedDatabase
             };        
             ShowPopUpChangePassword(popUpChangePassData);
         }
 

         private List<DataBase> DataBases
         {
             get
             {
                 return ViewState["DataBases"] as List<DataBase>;
             }
             set
             {
                 ViewState["DataBases"] = value;
             }
         }



         private void ShowPopUpChangePassword(PopUpChangePassData popUpChangePassData)
         {
             string jscript;

             jscript = "javascript:";
             jscript = jscript + "scrW = (document.body.clientWidth/2)-(530/2); ";
             jscript = jscript + "scrH = (document.body.clientHeight/2)-(280/2)-100; ";
             jscript = jscript + "window.open('{0}','urlPopup','height=290,width=540,status=no,toolbar=no,menubar=no,location=no,resizable=yes,scrollbars=no,left='+scrW+',top='+scrH);";

             Session["PopUpChangePassData"] = popUpChangePassData;
             Session["PopUpCambioPassActivo"] = "-1";

             ScriptManager.RegisterStartupScript(Page, GetType(), "AbrirPopup", String.Format(jscript, this.ResolveUrl("PopUpChangePassword.aspx")), true);
             


         }

         //private void ShowPopUpChangePassword(PopUpChangePassData popUpChangePassData)
         //{
         //    string jscript;

         //    jscript = "javascript:";
         //    jscript = jscript + "scrW = (document.body.clientWidth/2)-(530/2); ";
         //    jscript = jscript + "scrH = (document.body.clientHeight/2)-(280/2)-100; ";
         //    jscript = jscript + "window.open('{0}','urlPopup','height=380,width=560,status=yes,toolbar=no,menubar=no,location=no,resizable=no,scrollbars=no,left='+scrW+',top='+scrH);";

         //    Session["PopUpChangePassData"] = popUpChangePassData;           
         //    Page.ClientScript.RegisterStartupScript(GetType(), "AbrirPopup", String.Format(jscript, this.ResolveUrl("PopUpChangePassword.aspx")), true);
         //    //Deslogueo al usuario para poder realizar el cambio de password
           
         //}



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

        public string get_NombreBase(string key)
        {
            string dsm = ConfigurationManager.AppSettings["DatabaseSelectionMethod"].ToLower();
            System.Collections.Generic.List<DataBase> DataBases = DataBaseServiceProxy.Find(dsm);
            
            for (int i = 0; i < DataBases.Count; i++)
            {
                try
                {
                    if (Convert.ToString(DataBases[i].Id) == key)//verifica si la base es default
                    {
                        return DataBases[i].Name;
                    }
                }
                catch { }
            }

            return "";
        }


        /// <summary>
        /// Devuelve la fecha actual traducida, segun el lenguaje activo
        /// </summary>
        public string Traducir_Fecha() {
             
            string fecTraducida;
            string IdiomaSel = "es-ES";
            string LenguajeActivo = Obj_Lenguaje.Idioma();

            if (LenguajeActivo.Length == 4)
            {//Viene de la forma "enUS"
                IdiomaSel = LenguajeActivo.Substring(0, 2) + "-" + LenguajeActivo.Substring(2, 2);
            }
            else {
                if (LenguajeActivo.Length == 5)
                    IdiomaSel = LenguajeActivo;
            }

            DateTime dt = DateTime.Now;
            //Modifico el idioma de la fecha al idioma seleccionado
            try
            {
                Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo(IdiomaSel);
            }
            catch {
                Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("es-ES");
            }
            
            //convierto la fecha al formato largo.Ej: miercoles, 5 de Abril de 2012
            fecTraducida = dt.ToLongDateString();            
            
            return fecTraducida;        
             
        }


        /// <summary>
        /// Carga en la session "ConnString" el string de conexion a la base por default. Se encuentra en el web.config del newhome
        /// </summary>
        protected internal void LoadConexionDefault()
        {
           string dsm = ConfigurationManager.AppSettings["DatabaseSelectionMethod"].ToLower();

            System.Collections.Generic.List<DataBase> DataBases = DataBaseServiceProxy.Find(dsm);
                //Busca la base por default del webconfig
                for (int i = 0; i < DataBases.Count; i++)
                {                             
                    if (  DataBases[i].IsDefault.ToString() == "TrueValue" )//verifica si la base es default
                    {                      
                        String Cs = ConfigurationManager.ConnectionStrings[DataBases[i].Id.ToString()].ConnectionString;                      
                        System.Data.SqlClient.SqlConnection conex = new System.Data.SqlClient.SqlConnection(Cs);

                        if (System.Web.HttpContext.Current.Session["ConnString"] == null)
                        {//Copia en la variable de session "ConnString" el string de conexion                           
                          System.Web.HttpContext.Current.Session["ConnString"] = Cs;                                                   
                        }
                       
                        
                    }                 
                }
       }
    
 
        public void ActualizaGadgets(object sender, EventArgs e) {
            
            ContenedorPrincipal.ActualizaGadgets(sender, e);
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
           
        }

       


        #endregion      

        
   
    
    }




     
}

