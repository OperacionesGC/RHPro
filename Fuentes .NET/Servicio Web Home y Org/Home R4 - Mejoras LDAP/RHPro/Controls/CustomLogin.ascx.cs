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
using System.Data.OleDb;



namespace RHPro.Controls
{
    public partial class CustomLogin : UserControl
    {
        #region Events

        protected internal delegate void UserLoginHandle(object sender, EventArgs e);
        protected internal delegate void UserLogoutHandle(object sender, EventArgs e);

        protected internal event UserLoginHandle UserLogin;
        protected internal event UserLogoutHandle UserLogout;


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

        #endregion


        public RHPro.Lenguaje ObjLenguaje;         
        

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

        /// <summary>
        /// Id de la base de datos seleccionada
        /// </summary>
        private string SelectedDatabaseId
        {
            get
            {
                string dsm = ConfigurationManager.AppSettings["DatabaseSelectionMethod"].ToLower();

                if (dsm == "c")
                //return cmbDatabase.Text;
                {
                    for (int i = 0; i < DataBases.Count; i++)
                    {
                        if (DataBases[i].Name == cmbDatabase.SelectedItem.Text)
                            return DataBases[i].Id;
                    }
                }
                else
                {
                    for (int i = 0; i < DataBases.Count; i++)
                    {
                        if (DataBases[i].Name == lstDatabase.SelectedItem.Text)
                            return DataBases[i].Id;
                    }
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
            Page.PreLoad += new EventHandler(Page_PreLoad);                         
        }
 

        protected void Page_PreLoad(object sender, EventArgs e)
        { 
            if (!IsPostBack)
            {
                //levanto la ruta del WS
                UtilsProxy.ChangeWS(ConfigurationManager.AppSettings["RootWS"]);
                LoadDatabases();    
                ViewState.Add("lstIndex", -1);
            }            
        }

        public void cmbDatabase_SelectedIndexChanged(object sender, EventArgs e)
        {             
            //En el caso que se haya logeado carga la base seleccionada
            if (!Utils.IsUserLogin)       
            
            {
                Utils.SessionBaseID = SelectedDatabaseId;
                txtUserName.Focus();
            }
             
        }
       

        protected void Page_Load(object sender, EventArgs e)
        {        
             try {
          ObjLenguaje = new RHPro.Lenguaje();
        
            if (!IsPostBack)
            {
                ShowUserPanel(Utils.IsUserLogin);
                //btnClean.OnClientClick = String.Format("ClearValue('{0}');ClearValue('{1}');return false;", txtUserName.UniqueID.Replace("$", "_"), txtPassword.UniqueID.Replace("$", "_"));
            }


            if (Utils.IsUserLogin)
                 Candado.ImageUrl = "~/img/logout.png";
             

            if (bool.Parse(ConfigurationManager.AppSettings["EnableIntegrateSecurity"]) || bool.Parse(ConfigurationManager.AppSettings["LDAP_UseAuthentication"]))
            {
                txtUserName.Disabled = true;
                
               /* string userName = Request.ServerVariables["AUTH_USER"];               
                if (userName.Contains(@"\"))
                    userName = userName.Substring(userName.IndexOf(@"\") + 1);
                */
                string userName =  Request.ServerVariables["LOGON_USER"];
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

        #endregion

        #region Controls Handles

        protected void doPolitic_click(object sender, EventArgs e)
        {
            Page.ClientScript.RegisterStartupScript(GetType(), "AbrirPopup", String.Format("javascript:window.open('{0}','urlPopup','height=350,width=450,status=yes,toolbar=no,menubar=no,location=no,resizable=yes,scrollbars=yes,left=5,top=5');", this.ResolveUrl(UrlPolitic.ToString())), true);
        }

        protected void btnLogOut_Click(object sender, EventArgs e)
        {   //Actualiza el lenguaje al default
            System.Web.HttpContext.Current.Session["primerIdioma"] = false;
            System.Web.HttpContext.Current.Session["ConnString"] = null;

            Common.Utils.SessionBaseID = DatabaseIdDefault();

            //LoadConexionDefault(); //JPB: Se elimina
            Utils.LogoutUser();
           
            ShowUserPanel(Utils.IsUserLogin);
            Ingresar.Text = string.Empty;                       
            
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
             
            String EtiqLenguaje;
                                         
            if (lstDatabase.Visible)
                Session["lstIndex"] = lstDatabase.SelectedIndex;
            else
                Session["lstIndex"] = cmbDatabase.SelectedIndex;
       
            //Inicio el proceso de logueo validando el usuario y pass

            Login login = LoginServiceProxy.Find(txtUserName.Value, txtPassword.Value, EncryptionKey, SelectedDatabase.IntegrateSecurity, SelectedDatabaseId, EncriptUserData, Thread.CurrentThread.CurrentCulture.Name);
             
            if (login.IsValid)           
                {                
                   //EtiqLenguaje = (String)System.Web.HttpContext.Current.Session["Lenguaje"];
                  
                   //jpb: Cambio el idioma del home por el configurado para el usuario logeado
                   EtiqLenguaje = login.Lenguaje;

                    /*Cambio de base a la seleccionada*/
                    Common.Utils.SessionBaseID = SelectedDatabaseId;

                    Utils.LoginUser(txtUserName.Value, txtPassword.Value, EncriptUserData, EncryptionKey, EtiqLenguaje, login.MaxEmpl);
                   
                    //Cambio el menu Login
                    //ShowLoginInvalidMessage(false, string.Empty);

                    if (Utils.IsUserLogin)
                    {
                        Session["Username"] = txtUserName.Value;
                        Session["Password"] = txtPassword.Value;
                    
                    }

                    ShowUserPanel(Utils.IsUserLogin);
                    //ShowUserPanel(true);

                    if (UserLogin != null)
                    {
                        UserLogin(this, new EventArgs());
                    }       
 

                }
                else
                {

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
                        ShowLoginInvalidMessage(login.Messege);
                        /*jpb: Si no se pudo logear volver a la base por default*/
                        Common.Utils.SessionBaseID = DatabaseIdDefault();
                    }
                }
               
        }



        public void doChangeDB_Click(object sender, EventArgs e)
        {
             
            if (!Utils.IsUserLogin)            
            {
                Utils.SessionBaseID = SelectedDatabaseId;
                txtUserName.Focus();
            }
              
        }

        #endregion

        #region Methods

 

        /// <summary>
        /// Busca y carga las bases disponibles 
        /// </summary>
        protected internal void LoadDatabases()
        {
            string dsm = ConfigurationManager.AppSettings["DatabaseSelectionMethod"].ToLower();

            DataBases = DataBaseServiceProxy.Find(dsm);

            if (Session["lstIndex"] == null)
                Session["lstIndex"] = -1;

            if (dsm == "c")
            {
                cmbDatabase.Visible = true;
                lstDatabase.Visible = false;
                PanellstDatabase.Visible = false;

                cmbDatabase.DataValueField = "Id";
                cmbDatabase.DataTextField = "Name";

                //cmbDatabase.DataSource = DataBases;
                //cmbDatabase.DataBind();

                cmbDatabase.Items.Clear();

                for (int i = 0; i < DataBases.Count; i++)
                {
                    ListItem li = new ListItem(DataBases[i].Name, i.ToString());
                    cmbDatabase.Items.Add(li);
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
            else
            {
                if (dsm == "l")
                {
                    cmbDatabase.Visible = false;
                    lstDatabase.Visible = true;
                    PanellstDatabase.Visible = true;

                    lstDatabase.DataValueField = "Id";
                    lstDatabase.DataTextField = "Name";

                    //lstDatabase.DataSource = DataBases;
                    //lstDatabase.DataBind();

                    lstDatabase.Items.Clear();

                    for (int i = 0; i < DataBases.Count; i++)
                    {
                        ListItem li = new ListItem(DataBases[i].Name, i.ToString());
                        lstDatabase.Items.Add(li);
                    }

                    if (string.IsNullOrEmpty(Utils.SessionBaseID))
                    {
                        lstDatabase.SelectedIndex = DataBases.IndexOf(DataBases.Find(db => db.IsDefault.Equals(Utils.IsDefaultConstants.TrueValue)));
                        Utils.SessionBaseID = DataBases[DataBases.IndexOf(DataBases.Find(db => db.IsDefault.Equals(Utils.IsDefaultConstants.TrueValue)))].Id;
                        Session["lstIndex"] = lstDatabase.SelectedIndex;
                    }
                    else
                    {
                        lstDatabase.SelectedIndex = (int)Session["lstIndex"];
                    }
                }
            }
           
              

        }

        private void ShowUserPanel(bool visible)
        {
            if (visible) //Si inició sesión...
            {
                lblUser.InnerText = (string)Session["UserName"];
                Ingresar.Text = (string)Session["UserName"];
              
              
                LoginON.Style.Add(HtmlTextWriterStyle.Display, "none");
                LoginOFF.Style.Add(HtmlTextWriterStyle.Display, "block");

                if (cmbDatabase.Visible)
                {
                    //cmbDatabase.Text = Utils.SessionBaseID;
                    LabelBaseSeleccionada.Text = cmbDatabase.SelectedItem.Text;
                }
                else
                    if (lstDatabase.Visible)
                    {
                        LabelBaseSeleccionada.Text = lstDatabase.SelectedItem.Text;
                    }


            }
            else //Si cerró sesión...
            {
                lblUser.InnerText = string.Empty;
                //vacio la sesion
                System.Web.HttpContext.Current.Session["yaentro"] = null;

                LoginON.Style.Add(HtmlTextWriterStyle.Display, "block");
                LoginOFF.Style.Add(HtmlTextWriterStyle.Display, "none");
                LabelBaseSeleccionada.Text = "";

                if (cmbDatabase.Visible)
                {
                    //cmbDatabase.Text = Utils.SessionBaseID;
                    cmbDatabase.SelectedIndex = (int)System.Web.HttpContext.Current.Session["lstIndex"];
                    //cmbDatabase.SelectedIndex = (int)Session["lstIndex"];
                }
                else
                    if (lstDatabase.Visible)
                    {
                        //lstDatabase.SelectedIndex = (int)Session["lstIndex"];
                        lstDatabase.SelectedIndex = (int)System.Web.HttpContext.Current.Session["lstIndex"];
                        lstDatabase.Focus();
                    }
            }
        }

        //private void ShowLoginInvalidMessage(bool visible, string mensaje)
        //{
        //    ErrorMessege.CssClass = visible ? "ErrorMessegeON" : "ErrorMessegeOFF";
        //    if (!string.IsNullOrEmpty(mensaje))
        //        ErrorMessege.Text = mensaje;
        //    //ajuste de estilos para IE
        //    if (Request.Browser.Browser == "IE")
        //    {
        //        btnLogin.Style.Add(HtmlTextWriterStyle.MarginLeft, "0px");
        //    }
        //    ScriptManager.RegisterStartupScript(Page, GetType(), "Mensaje", string.Concat("$(document).ready(function() { ", string.Format("alert('{0}');",mensaje) , "});"), true);
        //}

        private void ShowLoginInvalidMessage(string mensaje)
        {
            //ScriptManager.RegisterStartupScript(Page, GetType(), "Mensaje", string.Concat("$(document).ready(function() { ", string.Format("alert('{0}');", mensaje), "});"), true);
            ScriptManager.RegisterStartupScript(Page, GetType(), "Mensaje", string.Format("alert('{0}');", mensaje), true);
        }

        private void ShowPopUpChangePassword(PopUpChangePassData popUpChangePassData)
        {
            string jscript;

            jscript = "javascript:";
            jscript = jscript + "scrW = (document.body.clientWidth/2)-(530/2); ";
            jscript = jscript + "scrH = (document.body.clientHeight/2)-(280/2)-100; ";
            jscript = jscript + "window.open('{0}','urlPopup','height=380,width=560,status=yes,toolbar=no,menubar=no,location=no,resizable=no,scrollbars=no,left='+scrW+',top='+scrH);";

            Session["PopUpChangePassData"] = popUpChangePassData;
            //Page.ClientScript.RegisterStartupScript(GetType(), "AbrirPopup", String.Format("javascript:window.open('{0}','urlPopup','height=260,width=530,status=yes,toolbar=no,menubar=no,location=no,resizable=no,scrollbars=no,left=document.body.clientWidth / 2,top= document.body.clientHeight / 2');", this.ResolveUrl(UrlPopup.ToString())), true);
            Page.ClientScript.RegisterStartupScript(GetType(), "AbrirPopup", String.Format(jscript, this.ResolveUrl(UrlPopup.ToString())), true);
        }



        #endregion

        protected void lstDatabase_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

         
     
        
    }
}