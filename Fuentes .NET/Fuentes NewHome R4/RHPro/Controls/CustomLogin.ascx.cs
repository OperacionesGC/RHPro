using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Threading;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Common;
using Entities;
using ServicesProxy;
using Login=Entities.Login;


using ServicesProxy.rhdesa;
using ServicesProxy.MetaHome;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Data;
using System.Text.RegularExpressions;




namespace RHPro.Controls
{
    public partial class CustomLogin : UserControl
    {
        #region Events

        protected internal delegate void UserLoginHandle(object sender, EventArgs e);
        protected internal delegate void UserLogoutHandle(object sender, EventArgs e);

        protected internal event UserLoginHandle UserLogin;
        protected internal event UserLogoutHandle UserLogout;
        protected Default Padre;


        #endregion

        #region Constants

        /// <summary>
        /// Direccion de la url del popup para cambiar el passs
        /// </summary>
        private const string UrlPopup = "../PopUpChangePassword.aspx";

        /// <summary>
        /// Direccion de la url del popup de politicas
        /// </summary>
        private const string UrlPolitic = "../PopUpPolitics.aspx";

        /// <summary>
        /// 
        /// </summary>
        private static readonly string EncryptionKey = ConfigurationManager.AppSettings["EncryptionKey"];
        /// <summary>
        /// 
        /// </summary>
        private static readonly bool EncriptUserData = bool.Parse(ConfigurationManager.AppSettings["EncriptUserData"]);
        private static readonly Int32 DiasBloqueo = Convert.ToInt32(ConfigurationManager.AppSettings["DiasBloqueo"]);

        public String ScriptSmartMenu = " $(function() {  $('#main-menuTop').smartmenus({  subMenusSubOffsetX: 0, subMenusSubOffsetY: -1, mainMenuSubOffsetX:0,mainMenuSubOffsetY:0,subMenusMinWidth:'60px', subMenusMaxWidth:'900px'  }); });  ";
        public String ScriptReloadPag = " this.location=this.location;  ";

        #endregion


        public static RHPro.Lenguaje ObjLenguaje;
        public MetaHome MH;

        #region Properties

        /// <summary>
        /// 
        /// Base de datos seleccionada
        /// </summary>
        private DataBase SelectedDatabase
        {
            get
            {
                return DataBases.Find(db => db.Id == SelectedDatabaseId);
            }
        }

        public void InicializaControl(Default p)
        {
            Padre = p;
        }
         
        /// <summary>
        /// Id de la base de datos seleccionada
        /// </summary>
        private string SelectedDatabaseId
        {
            get
            {
                string dsm = ConfigurationManager.AppSettings["DatabaseSelectionMethod"].ToLower();

                
                    for (int i = 0; i < DataBases.Count; i++)
                    {
                        if (DataBases[i].Name == cmbDatabase.SelectedItem.Text)
                            return DataBases[i].Id;
                    }
                

                return "";
            }
        }

        /// <summary>
        /// Bases de datos disponibles
        /// </summary>
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

        #endregion

        #region Page Handles

        

        protected void Page_Init(object sender, EventArgs e)
        {
            MH = new MetaHome();
            Page.PreLoad += new EventHandler(Page_PreLoad);
        
        }

      
        protected void Page_PreLoad(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                //levanto la ruta del WS
                UtilsProxy.ChangeWS(ConfigurationManager.AppSettings["RootWS"]);                               

                if (MH.MetaHome_Activo())//Si esta en modo SaaS me conecto al webservice ws_ext
                {                    
                    MH.Iniciar_Ws_Ext();
                }

                if (!Utils.IsUserLogin)
                {                    
                    LoadDatabases();
                    ViewState.Add("lstIndex", -1);
                    
                }
            }            
        }

    


        public void cmbDatabase_SelectedIndexChanged(object sender, EventArgs e)
        {             
            //En el caso que se haya logeado carga la base seleccionada
            if (!Utils.IsUserLogin)            
            {
                Utils.SessionBaseID = SelectedDatabaseId;               
            }             
        }
 
        

        protected void Page_Load(object sender, EventArgs e)
        {
            Boolean KeyPressJS = (Request.Form["__EVENTARGUMENT"] == "loginJS");

          

            if (MH.MetaHome_Activo())//Si esta en modo SaaS me conecto al webservice ws_ext
            {                
                MH.Iniciar_Ws_Ext();
            }

            //RegistrarScript(ScriptSmartMenu, KeyPressJS);                
            ScriptManager.RegisterStartupScript(this, typeof(Page), "InicializaMenuTop_1", "$(function() {  $('#main-menuTop').smartmenus({  subMenusSubOffsetX: 0, subMenusSubOffsetY: -1, mainMenuSubOffsetX:0,mainMenuSubOffsetY:0,subMenusMinWidth:'60px', subMenusMaxWidth:'900px'  }); });  ", true);

            if (KeyPressJS)
            {                
                if (!Utils.IsUserLogin)
                {
                    if (MH.MetaHome_Activo() && (MH.MetaHome_TipoFiltroLogin() == "2"))//Si filtra por usuario y esta activo SaaS, no muestro las bases
                    {
                        if (cmbDatabase.Visible == true)
                            do_login(true);
                        else
                            do_control(true);
                    }
                    else
                    {
                        
                        do_login(true);
                    }
                }
            }
            

            //Si estoy en el modo Pre Loguin, actualizo los campos del loguin para no perderlos
            if ((!Utils.IsUserLogin) && (Convert.ToString(Session["RHPro_PreLoguin"]) == "-1"))               
                Redibujar_Campos_y_Botones(false, true);
                
                  

           
            ObjLenguaje = new RHPro.Lenguaje();
           
            LabelUsr.InnerText = ObjLenguaje.Label_Home("Usuario");
            LabelPass.InnerText = ObjLenguaje.Label_Home("Contraseña");
            TituloSelBase.Text = ObjLenguaje.Label_Home("Base de Datos");
            PopUp_Limpiar.Text = ObjLenguaje.Label_Home("Cancelar");

            //Bienvenido.Text = ObjLenguaje.Label_Home("Bienvenido");
             
            /******************************************************** */
            /* Aqui verifica si el home es abierto desde el Meta Home */
            if ((Request.QueryString["id"] != "") && (Request.QueryString["id"] != null))
            {
               // if (MH.MetaHome_Activo() && MH.MetaHome_RegistraLoguin())
                if (  MH.MetaHome_RegistraLoguin() )
                {
                    Session["BaseDesabr"] = Request.QueryString["BaseDesabr"];
                    Login_Desde_MetaHome(Request.QueryString["id"]);
                }
                else
                    Response.Redirect("Default.aspx");

            }            
            /*********************************************************/             

            try {                
        
                if (!IsPostBack)
                {
                  
                    ShowUserPanel(Utils.IsUserLogin);
                  
                }
               

                if (Utils.IsUserLogin)
                {                    
                    PopUp_BotonControlar.Attributes.CssStyle.Add("display", "none !important");
                    PopUp_BotonLogin.Attributes.CssStyle.Add("display", "none !important");
                    //PopUp_Politicas.Attributes.CssStyle.Add("display", "none !important");
                              

                    CerrarSesion.Attributes.CssStyle.Add("display", "");
                    CerrarSesion.Text = ObjLenguaje.Label_Home("Cerrar Sesion");

                    Info_Base_Seleccionada.Attributes.CssStyle.Add("display", "");
                    Combo_Bases_Formulario.Attributes.CssStyle.Add("display", "none !important");
                    Campos_Formulario.Attributes.CssStyle.Add("display", "none !important");

                    //Info_Base_Seleccionada.Controls.Add(new LiteralControl("<DIV class='info_Login'> <img  src='img/Modulos/SVG/LOGINBASE.SVG' align='absmiddle' class='IconoLogin'/>  " + ObjLenguaje.Label_Home("Base") + ": " + Convert.ToString(Session["NombreBaseSeleccionada"]) + "</DIV>"));
                    //Info_User_Logueado.Controls.Add(new LiteralControl("<DIV class='info_Login'> <img  src='img/Modulos/SVG/LOGINUSER.SVG' align='absmiddle' class='IconoLogin'/>     "+ObjLenguaje.Label_Home("Usuario")+":" + Utils.SessionUserName + "</DIV>"));
                    Info_Base_Seleccionada.Controls.Add(new LiteralControl("<DIV class='info_Login'> " + Utils.Armar_Icono("img/Modulos/SVG/LOGINBASE.svg", "IconoLogin", "", "align='absmiddle'", "") + ObjLenguaje.Label_Home("Base") + ": " + Convert.ToString(Session["NombreBaseSeleccionada"]) + "</DIV>"));
                    Info_User_Logueado.Controls.Add(new LiteralControl("<DIV class='info_Login'>  " + Utils.Armar_Icono("img/Modulos/SVG/LOGINUSER.svg", "IconoLogin", "", "align='absmiddle'", "") + ObjLenguaje.Label_Home("Usuario") + ":" + Utils.SessionUserName + "</DIV>"));

                    PopUp_Limpiar.Visible = false;
  
                }
                else
                {
                    PopUp_BotonLogin.Text = ObjLenguaje.Label_Home("Acceder");
                    PopUp_BotonControlar.Text = ObjLenguaje.Label_Home("Acceder");
                    CerrarSesion.Text = ObjLenguaje.Label_Home("Cerrar Sesion");
                    //PopUp_Politicas.Text = ObjLenguaje.Label_Home("Políticas");
                    //btnLogOut.Text = ObjLenguaje.Label_Home("Cerrar Sesion");

                    

                    /*NUEVA*/                    
                    //PopUp_FondoTransparente.Attributes.CssStyle.Add("display", "");
                    //PopUp_NewHome.Attributes.CssStyle.Add("display", "");
                    //PopUp_Cabecera.Controls.Add(new LiteralControl("Acceso"));


                    if (MH.MetaHome_Activo() && (MH.MetaHome_TipoFiltroLogin() == "2"))//Si filtra por usuario y esta activo SaaS, no muestro las bases
                    {
                        PopUp_BotonLogin.Attributes.CssStyle.Add("display", "none !important");
                        //PopUp_Politicas.Attributes.CssStyle.Add("display", "none !important");
                        PopUp_BotonControlar.Attributes.CssStyle.Add("display", "");
                    }
                    else
                    {
                        PopUp_BotonControlar.Attributes.CssStyle.Add("display", "none !important");
                        PopUp_BotonLogin.Attributes.CssStyle.Add("display", "");
                        //PopUp_Politicas.Attributes.CssStyle.Add("display", "");
                    }

                    CerrarSesion.Attributes.CssStyle.Add("display", "none !important");
                    Info_Base_Seleccionada.Attributes.CssStyle.Add("display", "none !important");
                    Combo_Bases_Formulario.Attributes.CssStyle.Add("display", "");
                    Campos_Formulario.Attributes.CssStyle.Add("display", "");
                    //Info_Estilo_Usuario.Attributes.CssStyle.Add("display", "none");
                    Info_User_Logueado.Attributes.CssStyle.Add("display", "none !important");
                    
                    //Page.ClientScript.RegisterStartupScript(GetType(), "AbrirGlobo", String.Format("javascript:Abrir_Globo('Globo_Loguin');", this.ResolveUrl(UrlPolitic.ToString())), true);
                }
     

                if (bool.Parse(ConfigurationManager.AppSettings["EnableIntegrateSecurity"]) || bool.Parse(ConfigurationManager.AppSettings["LDAP_UseAuthentication"]))
                {
                    txtUserName.Disabled = true;
                    /* string userName = Request.ServerVariables["AUTH_USER"];               
                     if (userName.Contains(@"\"))
                         userName = userName.Substring(userName.IndexOf(@"\") + 1);
                     */
                    string userName = Request.ServerVariables["LOGON_USER"];

                    if (userName.Contains(@"\"))
                        userName = userName.Substring(userName.IndexOf(@"\") + 1);
                    txtUserName.Value = userName;

                    // txtUserName.Value = Utils.SessionUserName;
                }

                if (bool.Parse(ConfigurationManager.AppSettings["EnableIntegrateSecurity"]) && bool.Parse(ConfigurationManager.AppSettings["LDAP_UseAuthentication"]) == false)
                    txtPassword.Disabled = true;

             }
             catch (Exception exe) { Response.Write("ERROR: " + exe.Message); }

 
        }

        /*Este metodo recupera el usuario y password ingresados desde el Meta Home*/
        public void Login_Desde_MetaHome(string idTemp)
        {
            try
            {

                List<String> datosMH = MH.MetaHome_fromLogin(idTemp);
                
                txtUserName.Value = datosMH[0];
                txtPassword.Value = datosMH[1];
                Utils.SessionBaseID = datosMH[2];
                Utils.SessionNroTempLogin = idTemp;

                ListItem LI = new ListItem("ControlMeta", Utils.SessionBaseID);
                LI.Selected = true;
                cmbDatabase.Items.Add(LI);

                //Comienza el login
                do_login(false);

            }
            catch (Exception e)
            {
                Err_MH.Attributes.Add("class", "Err_MH_Visible");
                Err_MH.Controls.Add(new LiteralControl(ObjLenguaje.Label_Home("Error") + ": RootWS_MetaHome: " + e.Message + " - " +  Utils.SessionNroTempLogin));
            }
        }

        #endregion

        #region Controls Handles

        protected void doPolitic_click(object sender, EventArgs e)
        {
            Page.ClientScript.RegisterStartupScript(GetType(), "AbrirPopup", String.Format("javascript:window.open('{0}','urlPopup','height=350,width=450,status=yes,toolbar=no,menubar=no,location=no,resizable=yes,scrollbars=yes,left=5,top=5');", this.ResolveUrl(UrlPolitic.ToString())), true);
        }

          

        

        /*Este metodo elimina la entrada creada en la tabla Temp_Login del Meta Home*/
         public void Logout_Desde_MetaHome()
         {
             try
             { 
                 MH.MetaHome_Logout();
             }
             catch (Exception ex) { throw ex; }             
         }

        
        
        
         /// <summary>
         ///  JPB - Actualiza la cookie donde mantiene el id de la aplicacion
         /// </summary>
         public void Actualizar_ASPNET_SessionId()
          {
              if (Request.Cookies["ASP.NET_SessionId"] != null)
              {
                  //Session.Abandon();
                  Response.Cookies["ASP.NET_SessionId"].Expires = DateTime.Now.AddYears(-30);
                  Response.Cookies["ASP.NET_SessionId"].Value = "";
              }

              //if (Request.Cookies["AuthToken"] != null)
              //{
              //    Response.Cookies["AuthToken"].Value = string.Empty;
              //    Response.Cookies["AuthToken"].Expires = DateTime.Now.AddMonths(-20);
              //}
         }




        

        protected void btnLogOut_Click(object sender, EventArgs e)
        {   //Actualiza el lenguaje al default
            System.Web.HttpContext.Current.Session["primerIdioma"] = false;
            System.Web.HttpContext.Current.Session["ConnString"] = null;

            Common.Utils.SessionBaseID = DatabaseIdDefault();

           /*Deslogea en el MetaHome*/
            if ((Utils.SessionNroTempLogin != null) && ((String)Utils.SessionNroTempLogin != ""))
            {
                //if (MH.MetaHome_Activo() && MH.MetaHome_RegistraLoguin())
                if (MH.MetaHome_RegistraLoguin())
                  Logout_Desde_MetaHome();
            }

            //Fuerza la actualizacion de la cookie donde se mantiene la informacion de la aplicacion
            /* ********* NOTA: ESTA LIMPIEZA DEL SessionId rompe cuando se utiliza el BigIP ******** */
            Actualizar_ASPNET_SessionId();
            
            Utils.LogoutUser();
            
            ShowUserPanel(Utils.IsUserLogin);                                 
            
            if (UserLogout != null)
            {
                UserLogout(this, new EventArgs());
            }             

            //Corro el Garbage Collector para que limpie los objetos sin uso
            System.GC.Collect();

        }

        /// <summary>
        /// Busca el identificador de la base por defecto
        /// </summary>
        public string DatabaseIdDefault()
        {
            string dsm = ConfigurationManager.AppSettings["DatabaseSelectionMethod"].ToLower();
            string def = "";

            
            System.Collections.Generic.List<DataBase> DataBases = DataBaseServiceProxy.Find(dsm);
            

            //Busca el id de la base por default del webconfig
            for (int i = 0; i < DataBases.Count; i++)
            {
                if (DataBases[i].IsDefault.ToString() == "TrueValue") //verifica si la base es default
                {                    
                    return DataBases[i].Id.ToString();
                }
            }
            return def;
        }


        /// <summary>
        /// Carga en la session "ConnString" el string de conexion a la base por default
        /// </summary>
        protected internal void LoadConexionDefault()
        {
            string dsm = ConfigurationManager.AppSettings["DatabaseSelectionMethod"].ToLower();

           
            System.Collections.Generic.List<DataBase> DataBases = DataBaseServiceProxy.Find(dsm);
           
            //Busca la base por default del webconfig
            for (int i = 0; i < DataBases.Count; i++)
            {
                if (DataBases[i].IsDefault.ToString() == "TrueValue") //verifica si la base es default
                {
                    String Cs = ConfigurationManager.ConnectionStrings[DataBases[i].Id.ToString()].ConnectionString;
                    System.Data.SqlClient.SqlConnection conex = new System.Data.SqlClient.SqlConnection(Cs);
 
                    System.Web.HttpContext.Current.Session["ConnString"] = Cs;                    
                }
            }
        }

       
        public void doLogin_Click(object sender, EventArgs e)
        {            
            do_login(false);            
        }

        public void doLogin_Control(object sender, EventArgs e)
        {
            do_control(false);
        }

        public void btnLogOut_Limpiar(object sender, EventArgs e)
        {
            Response.Redirect("Default.aspx");
        }


        public void do_control(bool DeKeyPress) {
            

            if ((txtUserName.Value != "") && (txtPassword.Value != ""))
            {
                cmbDatabase.Items.Clear();
                cmbDatabase.Visible = true;
                cmbDatabase.Enabled = true;

                List<int> ListaBasesPermitidas = new List<int>();
                //Recupero la url del home con la que me deseo loguear
                string URL = HttpContext.Current.Request.Url.AbsoluteUri.Trim();

                try
                {
                    try
                    {
                        //REMOTO
                        ListaBasesPermitidas = MH.MetaHome_getBases(URL, txtUserName.Value, txtPassword.Value);
                    }
                    catch (Exception exz) { throw exz; }

                   // if (!DBNull.Value.Equals(ListaBasesPermitidas))
                    if (ListaBasesPermitidas!=null)
                    {                        
                        string dsm = ConfigurationManager.AppSettings["DatabaseSelectionMethod"].ToLower();
                        DataBases = DataBaseServiceProxy.Find(dsm);
                        bool primera = true;
                        //Por cada uno de los string de conexion en ws/web.config verifico cuales estan dentro de la lista de bases obtenidas por ws_ext
                        for (int i = 0; i < DataBases.Count; i++)
                        {
                            if (ListaBasesPermitidas.Contains(Convert.ToInt32(DataBases[i].Id)))
                            {
                                ListItem li = new ListItem(DataBases[i].Name, i.ToString());
                                if (primera)
                                {
                                    primera = false;
                                    li.Selected = true;
                                }
                                
                                cmbDatabase.Items.Add(li);
                            }
                        }

                         
                        if (ListaBasesPermitidas.Count > 0)
                        {
                            //Informo que estoy en estado de Pre-Loguin
                            Session["RHPro_PreLoguin"] = "-1";
                            //Habilito armado del submenu del tope y actualizo valores de campos                            
                            RegistrarScript(ScriptSmartMenu, false);                
                            //ScriptManager.RegisterStartupScript(this, typeof(Page), "InicializaMenuTop_3", "$(function() {  $('#main-menuTop').smartmenus({  subMenusSubOffsetX: 0, subMenusSubOffsetY: -1, mainMenuSubOffsetX:0,mainMenuSubOffsetY:0,subMenusMinWidth:'60px', subMenusMaxWidth:'900px'  }); });  ", true);

                            if (ListaBasesPermitidas.Count == 1)//Si tiene una sola base se loguea automaticamente
                            {
                                cmbDatabase.Items[0].Selected = true;
                                do_login(false);
                            }
                            else
                            {                                 
                                //Refresco los botones y campos del loguin
                                Redibujar_Campos_y_Botones(DeKeyPress, true);
                            }
                        }
                        else
                        {
                            //ObjLenguaje = new RHPro.Lenguaje();
                            cmbDatabase.Visible = false;
                            ScriptManager.RegisterStartupScript(Page, GetType(), "MensajeErr1", string.Format("alert('" + ObjLenguaje.Label_Home("Error") + ":" + ObjLenguaje.Label_Home("Consulte con el administrador") + "');this.location=this.location;"), true);
                        }
                        
                    }
                    else
                    {
                        //ObjLenguaje = new RHPro.Lenguaje();
                        ScriptManager.RegisterStartupScript(Page, GetType(), "MensajeErr2", string.Format("alert('" + ObjLenguaje.Label_Home("Error") + ":" +  ObjLenguaje.Label_Home("Consulte con el administrador") + "');this.location=this.location;"), true);                         
                        
                    }


                }
                catch (Exception ex)
                {
                    //  throw ex;
                    Informes_Error.Controls.Add(new LiteralControl("ERROR:" + ex.Message));

                }
            }
        }

        /// <summary>
        /// Redibuja los campos y botones del loguin
        /// </summary>
        /// <param name="DeKeyPress">Se indica que viene de un evento KeyPress</param>
        /// <param name="ExpandirLogin">Especifica si tiene que expandir el loguin</param>
        public void Redibujar_Campos_y_Botones(bool DeKeyPress, bool ExpandirLogin)
        {
            Session["RHPro_PreLoguin"] = "";
        
            PopUp_ImagenUsr.Visible = false;
            PopUp_BotonControlar.Visible = false;
            PopUp_BotonLogin.Visible = true;
            //PopUp_Politicas.Visible = true;
            TituloSelBase.Visible = true;
           // IconoBases.Visible = true;

            txtPassword.Disabled = true;
            txtUserName.Disabled = true;

            //Refresco los botones y campos del loguin
            if (DeKeyPress)
            {
                String WOnLoad = "window.onload = function() { ";

                if (ExpandirLogin)
                {
                    WOnLoad += " document.getElementById('ctl00_content_Btn_Login_MenuTop').click();";
                }

                WOnLoad += " document.getElementById('" + PopUp_BotonLogin.ClientID + "').style.display='';";
                WOnLoad += " document.getElementById('IconoBases').style.display='';";
                 
                //WOnLoad += " document.getElementById('" + PopUp_Politicas.ClientID + "').style.display='';";
                WOnLoad += " document.getElementById('" + PopUp_BotonLogin.ClientID + "').focus();";
                WOnLoad += " document.getElementById('" + txtUserName.ClientID + "').className='InputDeshab';";
                WOnLoad += " document.getElementById('" + txtPassword.ClientID + "').className='InputDeshab';";
                WOnLoad += " } ";
                Response.Write("<script defer='true' async='true'>" + WOnLoad + "  </script>");

            }
            else
            {
                String SCManager = "";

                if (ExpandirLogin)
                {
                    ScriptManager.RegisterStartupScript(this, typeof(Page), "InicializaMenu", "$(function() {  $('#main-menu').smartmenus({  subMenusSubOffsetX: 0, subMenusSubOffsetY: -1 , hideOnClick: false  }); });  ", true);
                    ScriptManager.RegisterStartupScript(this, typeof(Page), "InicializaMenuTop", "$(function() {  $('#main-menuTopLoguin').smartmenus({  subMenusSubOffsetX: 0, subMenusSubOffsetY: -1, mainMenuSubOffsetX:0,mainMenuSubOffsetY:0,subMenusMinWidth:'60px', subMenusMaxWidth:'1060px', hideOnClick: true   }); });  ", true);
                    SCManager += " document.getElementById('ctl00_content_Btn_Login_MenuTop').click(); ";
                }

                SCManager += "  document.getElementById('" + PopUp_BotonLogin.ClientID + "').style.display='';";
                SCManager += " document.getElementById('IconoBases').style.display='';";
                SCManager += "  document.getElementById('" + txtUserName.ClientID + "').className='InputDeshab';";
                SCManager += "  document.getElementById('" + txtPassword.ClientID + "').className='InputDeshab';";

                ScriptManager.RegisterStartupScript(this, typeof(Page), "AperturaLogin_2", SCManager, true);
                PopUp_BotonControlar.Attributes.CssStyle.Add("display", "none !important");
                PopUp_BotonLogin.Attributes.CssStyle.Add("display", "");                 
                //Hago foco en el boton loguin                    
                ScriptManager.RegisterStartupScript(this, typeof(Page), "AperturaLogin_3", "  document.getElementById('" + PopUp_BotonLogin.ClientID + "').focus();", true);
                
            }                               

            ScriptManager.RegisterStartupScript(this, typeof(Page), "RefrescarPass", "document.getElementById('" + txtPassword.ClientID + "').value='" + txtPassword.Value + "';  ", true);
               
        }

 

        /// <summary>
        /// Inicializa las variables de sesion que mantiene los estilos del home y del usuario
        /// </summary>
        public void Crear_Variables_Estilo()
        {
             String sql="";
            
             Consultas cc = new Consultas();
             
             //Paso las credenciales al web service
             //cc.Credentials = System.Net.CredentialCache.DefaultCredentials;
            ////-----------------------------------------------------------
 

            sql = "select  H.estilocarpeta from estilos_home_user U ";
            sql += " inner join estilo_homex2 X2 on X2.idcarpetaestilo = U.estiloactivo ";
            sql += " inner join estilos_home H On H.idestilo = X2.idcarpetaestilo   ";
            sql += " where Upper(U.iduser)=Upper('" + Utils.SessionUserName + "') ";

            DataTable dt = cc.get_DataTable(sql, Utils.SessionBaseID);
            if (dt.Rows.Count>0)
              Session["CarpetaEstilo"] = dt.Rows[0]["estilocarpeta"];
            else
            {
              sql = " SELECT estilocarpeta FROM estilos_home WHERE idestilo = 1 ";
              dt = cc.get_DataTable(sql, Utils.SessionBaseID);
              if (dt.Rows.Count>0)
                  Session["CarpetaEstilo"] = dt.Rows[0]["estilocarpeta"];
              else
                  Session["CarpetaEstilo"] = "CSS_Neutro";
            }

        }

        /// <summary>
        /// Verifica SQL Injection
        /// </summary>
        /// <param name="valor"></param>
        /// <returns></returns>
        public bool Validar_SQL_INJECTION(String valor)
        {            
           
            string sPattern = "['\"]";
            string sPattern2 = "select|delete|drop|update|<|>|script|alter";

            return ((System.Text.RegularExpressions.Regex.Match(valor.ToLower(), sPattern).Success) || (System.Text.RegularExpressions.Regex.Match(valor.ToLower(), sPattern2).Success));
        }

        public bool Validar_CamposLoguin()
        {
            bool salida = true;

            //EVITAR COMILLAS SIMPLES, DOBLES y otros caracteres especificos que generan INYECCION DE SQL//
            if (Validar_SQL_INJECTION(txtUserName.Value) || Validar_SQL_INJECTION(txtPassword.Value))
            {
                RHPro.Lenguaje ObjLenguaje2= new Lenguaje();
                //ScriptManager.RegisterStartupScript(Page, GetType(), "MjeDatosInvalidos", "alert('" + ObjLenguaje2.Label_Home("Datos invalidos") + "');", true);
                //Response.Write("<script>alert('" + ObjLenguaje2.Label_Home("Datos invalidos") + "');</script>");                
                return false;
            }

  
            if (Utils.SessionSeg_NT == 0)//Si no esta habilitada la seguridad integrada controlo los campos de loguin
            {
                 
                    if ((txtUserName.Value == "") || (txtPassword.Value == "") || (!(cmbDatabase.SelectedIndex >= 0)))
                    {
                        salida = false;
                        if (txtUserName.Value == "")
                        {
                            ScriptManager.RegisterStartupScript(Page, GetType(), "MjeCamposVacios", "alert('" + ObjLenguaje.Label_Home("Falta Usuario") + "')", true);
                        }
                        else
                        {
                            if (txtPassword.Value == "")
                                ScriptManager.RegisterStartupScript(Page, GetType(), "MjeCamposVacios", "alert('" + ObjLenguaje.Label_Home("Falta Contraseña") + "'); document.getElementById('" + txtPassword.ClientID + "').focus();", true);
                            else
                                ScriptManager.RegisterStartupScript(Page, GetType(), "MjeCamposVacios", "alert('" + ObjLenguaje.Label_Home("Seleccione una base") + "'); document.getElementById('" + txtPassword.ClientID + "').focus();", true);

                        }

                    }
                
            }
                         
            return salida;
        }


         

        private string GenerateHashKey()
        {
            
            string IP = Request.ServerVariables.Get("REMOTE_ADDR");
            string userName = "";// System.Net.Dns.GetHostEntry(IP).HostName;//Request.ServerVariables.Get("AUTH_USER");
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

        public void RegistrarScript(String Script, Boolean ConJS)
        {

            if (ConJS)
                Response.Write("<script defer async>"+Script+"</script>");
            else
                ScriptManager.RegisterStartupScript(Page, GetType(), "BloqueScript",Script, true);
        }

        

        public void do_login(Boolean ConJS) {
            
            Boolean PuedeContinuar = true; 
            Consultas cc = new Consultas();
            String Jscript = "";

            string IP =  Request.ServerVariables.Get("REMOTE_ADDR");
            string userName = "";// System.Net.Dns.GetHostEntry(IP).HostName;//Request.ServerVariables.Get("AUTH_USER");
            String BaseSelec = Utils.SessionBaseID;
            
            /////JPB: SE LLEVO A ConsultaBaseC.dll 
            bool camposvalidos = Validar_CamposLoguin(); 
            //if (cc.Control_Bloqueo_IP(userName, IP, BaseSelec, DiasBloqueo)) //Controla que la ip no este bloqueada
            //{
            //    Jscript = "alert('" + ObjLenguaje.Label_Home("Cliente Bloqueado") + "'); ";
            //    Jscript += ScriptSmartMenu + ScriptReloadPag;
            //    RegistrarScript(Jscript,ConJS);                
            //    //Si no se pudo logear volver a la base por default
            //    Common.Utils.SessionBaseID = DatabaseIdDefault();
            //    PuedeContinuar = false; 
            //}

            //if ((PuedeContinuar) &&(!Validar_CamposLoguin()) )//Si no valida los campos de loguin no continuo el loguin
            //{
              
            //    //Actualizo intentos fallidos por IP
            //    cc.Actualizar_Bloqueos_IP(BaseSelec, userName, IP);

            //    Jscript = "alert('" + ObjLenguaje.Label_Home("Datos invalidos") + "'); ";                
            //    Jscript += ScriptSmartMenu + ScriptReloadPag;
            //    RegistrarScript(Jscript, ConJS);                
            //    //Si no se pudo logear volver a la base por default
            //    Common.Utils.SessionBaseID = DatabaseIdDefault();
            //    PuedeContinuar = false;

            //}

                         

            if (PuedeContinuar)
            {
                //Inicio el proceso de logueo validando el usuario y pass
                /*En el caso que el home verique que el password no es correcto debo borrar todo dato de inicio de sesion en el Meta Home*/
                /*Deslogea desde MetaHome*/
                Login login;

                if ((Utils.SessionNroTempLogin != null) && ((String)Utils.SessionNroTempLogin != ""))
                //En este caso deshabilito la seguridad integrada en el caso que este activa. Tiene prioridad el MetaHome
                //login = LoginServiceProxy.Find(txtUserName.Value, txtPassword.Value, EncryptionKey, Utils.IntegrateSecurityConstants.FalseValue, Utils.SessionBaseID, EncriptUserData, Thread.CurrentThread.CurrentCulture.Name);
                {
                    
                    //login = LoginServiceProxy.Find(txtUserName.Value, txtPassword.Value, EncryptionKey, Utils.IntegrateSecurityConstants.FalseValue, BaseSelec, EncriptUserData, Utils.Lenguaje, userName, IP);
                    login = LoginServiceProxy.Find(txtUserName.Value, txtPassword.Value, EncryptionKey, Utils.IntegrateSecurityConstants.FalseValue, BaseSelec, EncriptUserData, Utils.Lenguaje, userName, IP
                        ,camposvalidos,DiasBloqueo);
                    
                }
                else
                {
                    
                    
                    //login = LoginServiceProxy.Find(txtUserName.Value, txtPassword.Value, EncryptionKey, SelectedDatabase.IntegrateSecurity, SelectedDatabaseId, EncriptUserData, Thread.CurrentThread.CurrentCulture.Name);
                    //login = LoginServiceProxy.Find(txtUserName.Value, txtPassword.Value, EncryptionKey, SelectedDatabase.IntegrateSecurity, SelectedDatabaseId, EncriptUserData, Utils.Lenguaje, userName, IP);
                    login = LoginServiceProxy.Find(txtUserName.Value, txtPassword.Value, EncryptionKey, SelectedDatabase.IntegrateSecurity, SelectedDatabaseId, EncriptUserData, Utils.Lenguaje, userName, IP
                        , camposvalidos, DiasBloqueo);
                }


                if (login.IsValid)
                {
                   
                    Conexion con = new Conexion();
                    string guid = GenerateHashKey();
                    con.Iniciar_Sesion(login, SelectedDatabaseId, txtUserName.Value, txtPassword.Value, guid, cmbDatabase.SelectedItem.Text, Convert.ToString(cmbDatabase.SelectedIndex));
                }
                else
                {
                    System.Web.HttpContext.Current.Session["lstIndex"] = cmbDatabase.SelectedIndex;
                    System.Web.HttpContext.Current.Session["NombreBaseSeleccionada"] = cmbDatabase.SelectedItem.Text;

                    /*En el caso que el home verique que el password no es correcto debo borrar todo dato de inicio de sesion en el Meta Home*/
                    /*Deslogea desde MetaHome*/
                    if ((Utils.SessionNroTempLogin != null) && ((String)Utils.SessionNroTempLogin != ""))
                    {
                        //if (MH.MetaHome_Activo() && MH.MetaHome_RegistraLoguin())
                        if (MH.MetaHome_RegistraLoguin())
                            Logout_Desde_MetaHome();
                    }


                    if (login.RequiredChangePassword)
                    {
                        // Disparas popup para que cambie el pass con el mensaje  y carga en el session los datos del logueo
                        PopUpChangePassData popUpChangePassData = new PopUpChangePassData
                        {
                            Login = login,
                            UserName = txtUserName.Value,
                            DataBase = SelectedDatabase
                        };
                        /*Cambio de base a la seleccionada*/
                        //Common.Utils.SessionBaseID = SelectedDatabaseId;
                        ShowPopUpChangePassword(popUpChangePassData);
                    }
                    else
                    {
                        //ShowLoginInvalidMessage(true, login.Messege);
                        PopUp_BotonControlar.Attributes.CssStyle.Add("display", "none !important");
                        PopUp_ImagenUsr.Visible = false;
                        PopUp_BotonControlar.Visible = false;
                        PopUp_BotonLogin.Attributes.CssStyle.Add("display", "none !important");
                        PopUp_BotonLogin.Visible = true;
                        ShowLoginInvalidMessage(login.Messege);
                        /*jpb: Si no se pudo logear volver a la base por default*/
                        Common.Utils.SessionBaseID = DatabaseIdDefault();
                    }
                }
                //Control_ASPNET_SessionId();         
            }
 
        }


        //public void Iniciar_Sesion(Login login)
        //{
        //    /**********************************************************/
        //    /* Se genera el token de seguridad  de hackeo             */
        //    string guid = GenerateHashKey();
        //    Session["AuthToken"] = guid;
        //    //Response.Cookies.Add(new HttpCookie("AuthToken", guid));
        //    /**********************************************************/


        //    //EtiqLenguaje = (String)System.Web.HttpContext.Current.Session["Lenguaje"];
        //    //jpb: Cambio el idioma del home por el configurado para el usuario logeado

        //    //Recupero en objeto login
        //    Session["login"] = login;

        //    String EtiqLenguaje = login.Lenguaje;
        //    Consultas cc = new Consultas();
        //    if ((Utils.SessionNroTempLogin == null) || ((String)Utils.SessionNroTempLogin == ""))
        //    {
        //        Common.Utils.SessionBaseID = SelectedDatabaseId;

        //        //jpb -Actualizo el nombre del InitialCatalog de la conexion
        //        //Consultas cc1 = new Consultas();
        //        //Paso las credenciales al web service
        //        cc.Credentials = System.Net.CredentialCache.DefaultCredentials;

        //        Session["InitialCatalog"] = cc.Initial_Catalog(SelectedDatabaseId);

        //    }
        //    else
        //    {
        //        //jpb -Actualizo el nombre del InitialCatalog de la conexion
        //        //Consultas cc2 = new Consultas();
        //        //Paso las credenciales al web service
        //        cc.Credentials = System.Net.CredentialCache.DefaultCredentials;
        //        //-----------------------------------------------------------
        //        Session["InitialCatalog"] = cc.Initial_Catalog(Common.Utils.SessionBaseID);

        //    }

        //    //Variables de session que utiliza el metahome
        //    HttpContext context = HttpContext.Current;

        //    string url = HttpContext.Current.Request.Url.AbsoluteUri;
        //    Session["RHPRO_URL_Login"] = url;
        //    Session["RHPRO_IP_Login"] = Server.HtmlEncode(Request.UserHostAddress);
        //    //HttpBrowserCapabilities brObject = Request.Browser;             
        //    Session["RHPRO_Browser_Type"] = Request.UserAgent;
        //    //Session["RHPRO_Browser_Version"] = brObject.Version;

        //    Utils.LoginUser(txtUserName.Value, txtPassword.Value, EncriptUserData, EncryptionKey, EtiqLenguaje, login.MaxEmpl);

        //    /*********************************/
        //    //Una vez logueado controlo si tengo que asociar gadget al usuario
        //    if (Convert.ToBoolean(ConfigurationManager.AppSettings["Controlar_Primer_Acceso"]))
        //    {
        //        cc.Controlar_Gadget_EnLoguin(txtUserName.Value, txtPassword.Value, (SelectedDatabase.IntegrateSecurity).ToString(), Utils.SessionBaseID);
        //    }
        //    /*********************************/

        //    //Guardo el webservices utilizado
        //    String[] Ws_Actual = Regex.Split(Convert.ToString(ConfigurationManager.AppSettings["RootWS"]), "//");
        //    Ws_Actual = Regex.Split(Ws_Actual[1], "/");
        //    Session["RHPRO_WS"] = Ws_Actual[Ws_Actual.Length - 2];

        //    /*********************************/
        //    Usuarios CUsuarios = new Usuarios();
        //    CUsuarios.Inicializar_Estilos(true);
        //    Session["RHPRO_NombreModulo"] = "RHPROX2";
        //    /**********************************/


        //    //Vuelvo RHPRO_HayTraducciones en vacio para que resuelva las traducciones con el lenguaje del usuario
        //    System.Web.HttpContext.Current.Session["RHPRO_EtiqTraducidasHome"] = "";
        //    System.Web.HttpContext.Current.Session["RHPRO_HayTraducciones"] = "";


        //    /**********************************/
        //    //Guarda el string de conexion para pasarselo al conn_db
        //    //Session["RHPRO_constrUsu"] = cc.constr(Utils.SessionBaseID);
        //    /**********************************/
        //    Utils.CopyAspNetSessionToAspSession();
        //    /********************************/
        //    //Cambio el menu Login
        //    //ShowLoginInvalidMessage(false, string.Empty);

        //    if (Utils.IsUserLogin)
        //    {

        //        Session["Username"] = txtUserName.Value;
        //        Session["Password"] = txtPassword.Value;
        //    }

        //    ShowUserPanel(Utils.IsUserLogin);


        //    if (UserLogin != null)
        //    {
        //        UserLogin(this, new EventArgs());
        //    }

        //    Session["RHPro_PreLoguin"] = "0";
 
        //}

 
        public void doChangeDB_Click(object sender, EventArgs e)
        {
             
            if (!Utils.IsUserLogin)            
            {
                Utils.SessionBaseID = SelectedDatabaseId;
                
            }
              
        }

        #endregion

        #region Methods

        public string GetCurrentPageName()
        {
            string sPath = System.Web.HttpContext.Current.Request.Url.AbsolutePath;
            System.IO.FileInfo oInfo = new System.IO.FileInfo(sPath);
            string sRet = oInfo.Name;
            return sRet;
        } 

         
        /// <summary>
        /// Busca y carga las bases disponibles 
        /// </summary>
        public   void LoadDatabases()
        {         
            if (MH.MetaHome_Activo() && (MH.MetaHome_TipoFiltroLogin() == "2"))//Si filtra por usuario y esta activo SaaS, no muestro las bases
            {                
                cmbDatabase.Enabled = false;
                cmbDatabase.Visible = false;
                TituloSelBase.Visible = false;
                //IconoBases.Visible = false;
                PopUp_ImagenUsr.Visible = true;
                ScriptManager.RegisterStartupScript(this, typeof(Page), "Icbase_2", "document.getElementById('IconoBases').style.display='none';", true);

            }
            else
            {
                PopUp_ImagenUsr.Visible = false;
                TituloSelBase.Visible = true;
                //IconoBases.Visible = true;
                ScriptManager.RegisterStartupScript(this, typeof(Page), "Icbase_2", "document.getElementById('IconoBases').style.display='';", true);
                string dsm = ConfigurationManager.AppSettings["DatabaseSelectionMethod"].ToLower();

                DataBases = DataBaseServiceProxy.Find(dsm);

                if (Session["lstIndex"] == null)
                    Session["lstIndex"] = -1;
                                
                    cmbDatabase.Visible = true;
                    cmbDatabase.DataValueField = "Id";
                    cmbDatabase.DataTextField = "Name";
                    cmbDatabase.Items.Clear();
 
                    //Si esta activo SaaS, verifico cuales puedo mostrar para dicho Entorno
                    if (MH.MetaHome_Activo() && (MH.MetaHome_TipoFiltroLogin() == "1"))//Si esta activo SaaS y el tipo de filtro es por URL
                    {
                        List<int> ListaBasesPermitidas = new List<int>();
                        String Protocolo = Request.Url.Scheme;
                        Uri MyUrl = Request.UrlReferrer;
                        string URL = HttpContext.Current.Request.Url.AbsoluteUri.Trim();
                                                                         
                        ListaBasesPermitidas = MH.MetaHome_getBases(URL, "", "");                         

                        for (int i = 0; i < DataBases.Count; i++)
                        { 
                            if (ListaBasesPermitidas.Contains(Convert.ToInt32(DataBases[i].Id)))
                            {                        
                                ListItem li = new ListItem(DataBases[i].Name, i.ToString());                                
                                cmbDatabase.Items.Add(li);                                
                            }
                        }                       
                    }
                    else
                    {
                        for (int i = 0; i < DataBases.Count; i++)
                        {
                            ListItem li = new ListItem(DataBases[i].Name, i.ToString());
                            cmbDatabase.Items.Add(li);                           
                        }
                    }

                    if (string.IsNullOrEmpty(Utils.SessionBaseID))
                    {
                        cmbDatabase.SelectedIndex = DataBases.IndexOf(DataBases.Find(db => db.IsDefault.Equals(Utils.IsDefaultConstants.TrueValue)));
                        Utils.SessionBaseID = DataBases[DataBases.IndexOf(DataBases.Find(db => db.IsDefault.Equals(Utils.IsDefaultConstants.TrueValue)))].Id;
                        Session["lstIndex"] = cmbDatabase.SelectedIndex;
                     }
                    else
                    {
                        cmbDatabase.SelectedIndex = (int)Session["lstIndex"];                         
                    }
                 
            }
           
              

        }

        private void ShowUserPanel(bool visible)
        {
            if (visible) //Si inició sesión...
            {
                //lblUser.InnerText = (string)Session["UserName"];            
              
              
                //LoginON.Style.Add(HtmlTextWriterStyle.Display, "none");
                //LoginOFF.Style.Add(HtmlTextWriterStyle.Display, "block");
              
                //if ((Utils.SessionNroTempLogin != null) && ((String)Utils.SessionNroTempLogin != ""))
                //    LabelBaseSeleccionada.Text = (String) Session["BaseDesabr"]; 

                //else
                //{

                //    if (cmbDatabase.Visible)
                //    {                      
                //        LabelBaseSeleccionada.Text = cmbDatabase.SelectedItem.Text;

                //    }
                    
                //}
                


            }
            else //Si cerró sesión...
            {
                //lblUser.InnerText = string.Empty;
                //vacio la sesion
                System.Web.HttpContext.Current.Session["yaentro"] = null;

                //LoginON.Style.Add(HtmlTextWriterStyle.Display, "block");
                //LoginOFF.Style.Add(HtmlTextWriterStyle.Display, "none");
                //LabelBaseSeleccionada.Text = "";

                if (cmbDatabase.Visible)
                {                    
                    cmbDatabase.SelectedIndex = (int)System.Web.HttpContext.Current.Session["lstIndex"];                    
                }
                
            }
        }

 

        private void ShowLoginInvalidMessage(string mensaje)
        {
            //ScriptManager.RegisterStartupScript(Page, GetType(), "Mensaje", string.Concat("$(document).ready(function() { ", string.Format("alert('{0}');", mensaje), "});"), true);
            //ScriptManager.RegisterStartupScript(Page, GetType(), "Mensaje", string.Format("alert('{0}');", mensaje), true);
            ScriptManager.RegisterStartupScript(Page, GetType(), "Mensaje", string.Format("alert('{0}');this.location=this.location;", mensaje), true);
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
            //Page.ClientScript.RegisterStartupScript(GetType(), "AbrirPopup", String.Format("javascript:window.open('{0}','urlPopup','height=260,width=530,status=yes,toolbar=no,menubar=no,location=no,resizable=no,scrollbars=no,left=document.body.clientWidth / 2,top= document.body.clientHeight / 2');", this.ResolveUrl(UrlPopup.ToString())), true);
            //Page.ClientScript.RegisterStartupScript(GetType(), "AbrirPopup", String.Format(jscript, this.ResolveUrl(UrlPopup.ToString())), true);

            ScriptManager.RegisterStartupScript(Page, GetType(), "AbrirPopup", String.Format(jscript, this.ResolveUrl(UrlPopup.ToString())), true);
            ScriptManager.RegisterStartupScript(Page, GetType(), "RecargarDefault", "this.location = this.location;", true);          

            
        }



        #endregion

                 
     
        
    }
}