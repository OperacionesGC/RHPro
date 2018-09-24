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
        public RHPro.Lenguaje ObjLenguaje;
        public RHPro.Usuarios CUsuarios;

 
        
         protected void Page_Load(object sender, EventArgs e)
        {
 
            
                if (System.Web.HttpContext.Current.Session["yaentro"] == null)
                    System.Web.HttpContext.Current.Session["yaentro"] = false;
                if (System.Web.HttpContext.Current.Session["primerIdioma"] == null)
                    System.Web.HttpContext.Current.Session["primerIdioma"] = false;

                Obj_Lenguaje = new Lenguaje();

                

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
                    consultas.Credentials = System.Net.CredentialCache.DefaultCredentials;
                    //-----------------------------------------------------------
                    String DiasVence = consultas.Info_Dias_VencimientoPass(Common.Utils.SessionUserName, Utils.SessionBaseID);
                    if (!String.IsNullOrEmpty(DiasVence))
                    {   
                            String Info = "<img src='Images/atencion.png' align='absmiddle' height='23'> " + ObjLenguaje.Label_Home("Su contraseña expira en @@NUM1@@ días");
                            Info = Info.Replace("@@NUM1@@", DiasVence);
                            InfoCambioPass.Controls.Add(new LiteralControl(Info));
                            LinkButton LB = new LinkButton();
                            LB.ID = "CambioDePassword";
                            LB.Text = ObjLenguaje.Label_Home("Cambiar Contraseña");
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
                Modulos.AsignarContPpal(ContenedorPrincipal);
              

                Cargar_Datos_De_Usuario();
                Cargar_Datos_De_Lenguaje();
                Cargar_Datos_De_Favoritos();
                Cargar_Datos_De_Complementos();
                Cargar_Datos_De_Estilos();

                if (bool.Parse(ConfigurationManager.AppSettings["VisualizarFooter"]))
                  Cargar_Info_Piso();


                 Cargar_Estilo_De_Pagina();

                 ScriptManager.RegisterStartupScript(this, typeof(Page), "InicializaMenuTop", " if (document.getElementById('main-menuTop')){document.getElementById('main-menuTop').style.zIndex=400; if (document.getElementById('main-menu')){document.getElementById('main-menu').style.zIndex=300;}}", true);                 
                  
                 if (!Utils.IsUserLogin)
                 {
                     Btn_Login_MenuTop.Attributes.Add("onclick", "setTimeout(\"document.aspnetForm.ctl00$content$cLogin$txtUserName.focus();\", 30);");
                     Btn_Login_MenuTop.Attributes.Add("onmouseover", "setTimeout(\"document.aspnetForm.ctl00$content$cLogin$txtUserName.focus();\", 30);");               
                 }
                 else 
                 {
                     Btn_Login_MenuTop.Attributes.Add("onclick", "setTimeout(\"document.getElementById('ctl00_content_cLogin_CerrarSesion').focus();\", 30);");
                     Btn_Login_MenuTop.Attributes.Add("onmouseover", "setTimeout(\"document.getElementById('ctl00_content_cLogin_CerrarSesion').focus();\", 30);");
                 }
               
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

         public void Cargar_Estilo_De_Pagina()
         {
             if (Utils.SesionIniciada)
             {
                // if ((Convert.ToString(Session["RHPRO_Cambio_Estilo"]) == "-1") || (Convert.ToString(Session["RHPRO_RecienLogueado"]) == "-1"))                
                 if  ( Convert.ToString(Session["RHPRO_Cambio_Estilo"]) == "-1" )             
                 {
                     Session["RHPRO_Cambio_Estilo"] = "0";
                     Session["RHPRO_RecienLogueado"] = "0";                      
                     CUsuarios = new Usuarios();
                     CUsuarios.Inicializar_Estilos(true);
                     Utils.CopyAspNetSessionToAspEstilos();
                    // Pasar_Estilos_ASP();
                  }                

             }
             else
             {
              //   if (Convert.ToString(System.Web.HttpContext.Current.Session["EstiloR4_coloricono"]) == "")
                 //if (Convert.ToString(Session["EstiloR4_coloricono"]) == "")
                 //{
                     CUsuarios = new Usuarios();
                     CUsuarios.Inicializar_Estilos(false);
                     //Pasar_Estilos_ASP();
                 //}
             }

             Logo_Empresa.Attributes.Add("src", (String)System.Web.HttpContext.Current.Session["EstiloR4_logoEstilo"]);
             
         }

         

         public void Cargar_Redes_Sociales()
         {
             //RedesSociales
             String icono = "";

             Consultas cc = new Consultas();
             ////Paso las credenciales al web service
             cc.Credentials = System.Net.CredentialCache.DefaultCredentials;
             String sql = " SELECT  hredtitulo, hredpagina, paisnro, rhpro, ess, icono FROM home_redes WHERE rhpro=-1";
             DataTable dt = cc.get_DataTable(sql, Utils.SessionBaseID);
             String accion = "";
             foreach (DataRow dr in dt.Rows)
             {
                 accion = "abrirVentana('" + (String)dr["hredpagina"] + "','Red','800','600')";
                 //datos += "   <span onclick=\"abrirVentana('" + URL + "','VN','" + (String)dr["altoventana"] + "','" + (String)dr["anchoventana"] + "' )\">";
                 icono += "<img  src='img/REDES/" + (String)dr["icono"] + "' border='0' style='cursor: pointer;' title='" + (String)dr["hredtitulo"] + "' class='IconoREDES' onclick=\""+accion+"\">";
             }
            
             RedesSociales.Controls.Add(new LiteralControl(icono));
         }
		  
          public void Cargar_Info_Piso()
         {
             NombreDeEmpresa.InnerText = (String)ConfigurationManager.AppSettings["NombreDeEmpresa"];//"Heidt & Asociados S.A.";
             DireccionEmpresa.InnerText = (String)ConfigurationManager.AppSettings["DireccionEmpresa"];//"Suipacha 72 - 4º A CP C1008AAB - Buenos Aires Argentina.";
             TelefonoMailEmpresa.InnerText = (String)ConfigurationManager.AppSettings["TelefonoMailEmpresa"];//"Tel./Fax: +54 11 5252 7300 Email: ventas@rhpro.com ";
             Slogan.InnerText = (String)ConfigurationManager.AppSettings["SloganCliente"];
             versionMI.InnerText = VersionServiceProxy.Find(Utils.SessionBaseID, Thread.CurrentThread.CurrentCulture.Name);//carga la version
             patchMI.InnerText = PatcheServiceProxy.Find(Utils.SessionBaseID, Thread.CurrentThread.CurrentCulture.Name);//carga el patch
             Cargar_Redes_Sociales();

              URL_Logo.Controls.Add( new LiteralControl("<img  src='img/Logo_RHPro.png' width='110' style='margin-bottom:4px;' >"));
              
             
         }
          

         public void Cargar_Datos_De_Lenguaje()
         {
             
             Info_Lenguaje.Controls.Add(new LiteralControl("<img src='img/Modulos/SVG/IDIOMA.svg' border='0' class='IconosBarraTop' title='" + ObjLenguaje.Label_Home("Idioma") + "'> " + System.Web.HttpContext.Current.Session["ArgTitulo"]));             
         }

         public void Cargar_Datos_De_Favoritos()
         {
             Info_Favoritos.Controls.Add(new LiteralControl("<img src='img/Modulos/SVG/FAVORITO_TOP.svg' border='0' class='IconosBarraTop' > " + Obj_Lenguaje.Label_Home("Favoritos")));
         }

         public void Cargar_Datos_De_Complementos()
         {
             Info_Complementos.Controls.Add(new LiteralControl("<img src='img/Modulos/SVG/GADGET.svg' border='0' class='IconosBarraTop'  > " + Obj_Lenguaje.Label_Home("Gadgets")));
         }

         public void Cargar_Datos_De_Estilos()
         {
            // TituloEstilos.InnerText = ObjLenguaje.Label_Home("Selector de Estilos");
             Info_Estilos.Controls.Add(new LiteralControl("<img src='img/Modulos/SVG/ESTILOS.svg' border='0' class='IconosBarraTop'  > " + Obj_Lenguaje.Label_Home("Estilos")));
         }

         public void Cargar_Datos_De_Usuario()         
         {
             String Salida = "";
             String imagen = "<img src='img/Modulos/SVG/USER.svg' border='0' class='IconosBarraTop'    > ";

             if (Common.Utils.IsUserLogin)
             {
                 String sql = "";
                 Consultas cc = new Consultas();
                 //Paso las credenciales al web service
                 cc.Credentials = System.Net.CredentialCache.DefaultCredentials;
                 //-----------------------------------------------------------
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

                     ////Salida = imagen + " <span class='Log_User'> " + Common.Utils.SessionUserName + "</span> <span class='Log_Nombre'> " + dt.Rows[0]["ternom"] + "</span>";

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
                 Salida = imagen + " " + ObjLenguaje.Label_Home("Ingresar");


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
             jscript = jscript + "window.open('{0}','urlPopup','height=380,width=560,status=yes,toolbar=no,menubar=no,location=no,resizable=no,scrollbars=no,left='+scrW+',top='+scrH);";

             Session["PopUpChangePassData"] = popUpChangePassData;           
             Page.ClientScript.RegisterStartupScript(GetType(), "AbrirPopup", String.Format(jscript, this.ResolveUrl("PopUpChangePassword.aspx")), true);
             //Deslogueo al usuario para poder realizar el cambio de password
           
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
            
            //convierto la fecha al formato largo.Ej: mieroles, 5 de Abril de 2012
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

