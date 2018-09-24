using System;
using System.Collections;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Xml.Linq;

using Common;
using ServicesProxy;
using System.Xml;
using System.Threading;
using System.Collections.Generic;
using Entities;
using ServicesProxy.rhdesa;
 
 
 

namespace RHPro.Controls
{
    public partial class Modulos : System.Web.UI.UserControl
    {
        public RHPro.Lenguaje ObjLenguaje;
        public RHPro.Controls.ContenedorPrincipal ContPpal;
        public int posmenu = 0;
        //Se define el objeto conexión
        public System.Data.SqlClient.SqlConnection conn;
        public System.Data.SqlClient.SqlDataReader reader;
        public System.Data.SqlClient.SqlCommand sql;
        private bool VerModulosInactivos = bool.Parse(ConfigurationManager.AppSettings["VisualizarModulosInhabilitados"]);


        protected Default Padre;
        /// <summary>
        /// Modulos disponibles
        /// </summary>
        protected List<Module> AvailableModules
        {
            get { return ViewState["AvailableModules"] as List<Module>; }
            set { ViewState["AvailableModules"] = value; }
        }

        public void InicializaControl(Default p)
        {
            Padre = p;
        }
 
        protected void Page_Load(object sender, EventArgs e)
        {
            //ScriptManager.RegisterStartupScript(this, typeof(Page), "Controla", "document.getElementById('ControlaIDMod').innerHTML = 'OTRA:" + Convert.ToString(Request.Cookies["ASP.NET_SessionId"].Value) + "'; ", true);
            //Response.Write(" OTRA:" + Convert.ToString(Request.Cookies["ASP.NET_SessionId"].Value));
           
 
            if ((String)Session["Expandido"] == "")
                Session["Expandido"] = "1";

            if ( (Hashtable)Session["VisualizaModulos"] ==null )
            {
                Session["VisualizaModulos"] = new Hashtable();
                Utils.Modulos_Habilitados = new List<string>();
            }
            
            ObjLenguaje = new RHPro.Lenguaje();
            txtInactivos.InnerText = ObjLenguaje.Label_Home("Modulos Inactivos");
                  
            Seccion_Accesos.InnerText =  ObjLenguaje.Label_Home("Accesos");
            Cargar_Accesos();

            LoadModule();   
        }

        //protected void Page_PreRender(object sender, EventArgs e)
        //{           
        //    LoadModule();           
        //}




        public void ExpandirColapsar(object sender, EventArgs  e)
        {
            if ((String)Session["Expandido"] == "1")
            {
              Session["Expandido"] = "0";
                ScriptManager.RegisterStartupScript(Page, GetType(), "ControlIndicadoresModulo", "DesplazarBarraMenu_AJAX('1');", true);
               
            }
            else
            {
                Session["Expandido"] = "1";
                ScriptManager.RegisterStartupScript(Page, GetType(), "ControlIndicadoresModulo", "DesplazarBarraMenu_AJAX('0');", true);
                
            }

        }


        public string Agregar_Modulo_Visibilidad(String menudir, Boolean condicion)
        {
            //if (Session["VisualizaModulos"] == null)
            //{

            //UpdatePanel ccm = (UpdatePanel)Parent.FindControl("Update_Modulos");
            //AsyncPostBackTrigger trigger = new AsyncPostBackTrigger();

            //trigger.ControlID = "LinkButton_" + menudir;
            //trigger.EventName = "Command";
            //ccm.Triggers.Add(trigger);
             
            
                if (!((Hashtable)Session["VisualizaModulos"]).Contains(menudir))
                {
                    ((Hashtable)Session["VisualizaModulos"]).Add(menudir, condicion);
                     //Utils.Modulos_Habilitados.Add(Utils.getMenuDir(menudir));                    
                }

            if (condicion==true)
                Utils.Modulos_Habilitados.Add(Utils.getMenuDir(menudir)); 
            //}
            return "";
        }

      

        
        public void LoadModule()
        {
            String IdiomaActivo = Common.Utils.Lenguaje;
            //String IdiomaActivo = Thread.CurrentThread.CurrentCulture.Name;
            ServicesProxy.rhdesa.Consultas cc = new ServicesProxy.rhdesa.Consultas();          
            //Recupero los modulos disponibles para el usuario Anonimo
            AvailableModules = ModuleServiceProxy.Find(Utils.SessionUserName, Utils.SessionBaseID, IdiomaActivo);
            //Cargo el repeater con los modulos disponibles
            Repeater1.DataSource = AvailableModules;
            Repeater1.DataBind();      
             

            if (Utils.IsUserLogin)
            {
                //Recupero los modulos habilitados para el usuario y base seleccionados. 
                AvailableModules = ModuleServiceProxy.Find_Modulos(Utils.SessionUserName, Utils.SessionBaseID, IdiomaActivo, true);
                //Cargo el repeater con los modulos             
                Repeater2.DataSource = AvailableModules;
                Repeater2.DataBind();

                //Recupero los modulos inhabilitados para el usuario y base seleccionados
                AvailableModules = ModuleServiceProxy.Find_Modulos(Utils.SessionUserName, Utils.SessionBaseID, IdiomaActivo, false);
                //Cargo el repeater con los modulos
                if (AvailableModules.Count > 0)
                {
                    RepeaterModulosInactivos.DataSource = AvailableModules;
                    RepeaterModulosInactivos.DataBind();
                }
                else
                    Menu_Links_Inact.Visible = false;
                   // Menu_Links_Inact.Controls.Clear();
                
            }
             
        }

        public string TituloDeMenu(string titulo) {
            return ObjLenguaje.Label_Home(titulo);
        }

        public string Imprimir_Action(string accion, string MenuName,string menumsnro, string menuraiz)
        {
            string TAG_IMG = "";

            //Verifico si la accion tiene contenido
            if ((accion != "#") && (accion != ""))
            {
                HttpBrowserCapabilities brObject = Request.Browser;
                //TAG_IMG = " <img src='img/Modulos/SVG/APERTURA_MODULO.svg' border='0' class='IconoAperturaModulo' onclick=\"AbrirModulo(" + accion + ",'" + MenuName + "');AbrirMRU('" + menumsnro + "','" + menuraiz + "')\"  >  ";             
                TAG_IMG = Utils.Armar_Icono("img/Modulos/SVG/APERTURA_MODULO.svg", "IconoModuloAcceso", "", " onclick=\"AbrirModulo(" + accion + ",'" + MenuName + "');AbrirMRU('" + menumsnro + "','" + menuraiz + "')\" ", "");             
                
            }
           

            return TAG_IMG;
        }


        public System.Data.SqlClient.SqlConnection ConnexionDef()
        {
            System.Data.SqlClient.SqlConnection conex;
            String conexString = "";
            conexString = (String)System.Web.HttpContext.Current.Session["ConnString"];
            conex = new System.Data.SqlClient.SqlConnection(conexString);
            return conex;
        }
  
        public String Leer_XML() {
            string TR = "";
            string Nombre = "";
            string URL = "";
            string isLogin = "";
            string Activo = "";
            string ArchivoDescripcion = "";
            string idUser = "";

            //Busco el nombre del archivo de configuracion de Accesos
            String URL_XML = (String)ConfigurationManager.AppSettings["AccesosHomeXML"];
            URL_XML = "../" + URL_XML;
            DataSet ds = new DataSet();
            //ds.ReadXml(MapPath(URL_XML));
            //ds.ReadXml(MapPath("../Accesos_Home.xml"));
             
            foreach (DataRow row in ds.Tables["Acceso"].Rows)
            {                        
                Nombre = (String)row["Nombre"]; 
                URL = (String)row["URL"];  
                isLogin = (String)row["isLogin"]; 
                Activo = (String)row["Activo"];
                ArchivoDescripcion = (String)row["ArchivoDescripcion"];
                idUser = (String)row["idUser"];
                TR += Construir_Acceso(posmenu, Activo, Nombre, URL, isLogin, ArchivoDescripcion, idUser);                 
                posmenu++;            
            }

            return TR;
        }

        //------------------------------------------------------------------------------

        public DataSet DS_Datos_XML()
        {
            //Busco el nombre del archivo de configuracion de Accesos
            String URL_XML = (String) ConfigurationManager.AppSettings["AccesosHomeXML"];
            URL_XML = "../" + URL_XML;
            DataSet ds = new DataSet();
            ds.ReadXml(MapPath(URL_XML));
            return ds;
        }

       /*Este metodo carga todos los accesos leyendolos desde un archivo XML.*/      
        protected void Cargar_Accesos()
        {
            RepAccesos.Controls.Clear();

            
             
                string BaseId = Common.Utils.SessionBaseID;
                string UserName = Common.Utils.SessionUserName;

                /* ****** Preparo el filtro de accesos que deseo visualizar desde la base****** */

                //string sql = "SELECT * FROM Home_Accesos ";
                //sql += " WHERE Activo = -1 AND (";
                //if (Common.Utils.IsUserLogin) //Busco todos los accesos configurados para el usuario + los anonimos
                //    sql += "  idUser = '" + Common.Utils.SessionUserName + "' OR idUser IS NULL";
                //else
                //    sql += "  idUser IS NULL "; //Busco todos los accesos anonimos
                //sql += " ) ";
                //sql += " ORDER BY Nombre ASC ";
                string sql = "SELECT * FROM Home_Accesos ";
                sql += " WHERE Activo = -1 ";                
                sql += " ORDER BY Nombre ASC ";

                if (Common.Utils.IsUserLogin)
                {
                    if (Session["RHPRO_AccesosHomeLogin"] == null)
                    {
                        Consultas cc = new Consultas();
                        DataSet ds = cc.get_DataSet(sql, BaseId);
                        Session["RHPRO_AccesosHomeLogin"] = ds;                       
                    }
                    RepAccesos.DataSource = (DataSet)Session["RHPRO_AccesosHomeLogin"];
                }
                else
                {
                    if (Session["RHPRO_AccesosHome"] == null)
                    {
                        Consultas cc = new Consultas();
                        DataSet ds = cc.get_DataSet(sql, BaseId);
                        Session["RHPRO_AccesosHome"] = ds;                        
                    }
                    RepAccesos.DataSource = (DataSet)Session["RHPRO_AccesosHome"];
                }
                //RepAccesos.DataSource = ds;
             
            
           
            RepAccesos.DataBind();
            



            /* ****** Preparo el filtro de accesos que deseo visualizar desde el XML****** */
            /*
             string filtro = "Activo=true ";
             DataView dv = DS_Datos_XML().Tables[0].DefaultView;
             if (Common.Utils.IsUserLogin) //Busco todos los accesos configurados para el usuario + los anonimos
                 //filtro += " AND (idUser = '" + UserName + "' OR idUser = '') ";               
                 filtro += " AND idUser = '" + UserName + "'";               
             else
                 filtro += " AND idUser=''";  //Busco todos los accesos anonimos
                
             dv.RowFilter = filtro;
             dv.Sort = "Nombre ASC";
            
           
           
             RepAccesos.DataSource = dv;   
             //RepAccesos.DataSource = DS_Datos_XML();              
             RepAccesos.DataBind();            
              }
             
             */
            /* *************************************************************** */

        }


        public String ImprimirLink(bool isLogin, string URL){
            
         string link = "";
         if (isLogin == true)
         {
             if (Common.Utils.IsUserLogin)
             {
                 //link = "<img src='img/Modulos/SVG/APERTURA_MODULO.svg' class='IconoAperturaModulo' border='0' onclick=\"AbrirModulo('" + URL + "','ESS')\"   >";           
                 link = Common.Utils.Armar_Icono("img/Modulos/SVG/CONTROLMODULOS.svg", "IconoAperturaModulo", "", " border='0' onclick=\"AbrirModulo('" + URL + "','ESS')\" ", "");           
                  
                 
             }
         }
         else {
             //link = "<img src='img/Modulos/SVG/APERTURA_MODULO.svg' class='IconoAperturaModulo'  border='0'  onclick=\"AbrirModulo('" + URL + "','ESS')\"   >";           
             link = Common.Utils.Armar_Icono("img/Modulos/SVG/APERTURA_MODULO.svg", "IconoAperturaModulo", "", " border='0' onclick=\"AbrirModulo('" + URL + "','ESS')\" ", "");           
             
         }

         return link;

        }

 

        //public string Accesos(int pos, bool activo, string nombre, string URL, bool isLogin, string ArchDesc)
        //{
        //    string TR = "";
             
        //        if (activo == true)
        //        {
        //            TR = "<tr id='Link" + pos + "' onclick=\"Seleccionar('Link" + pos + "','')\"   onmouseover='Sobre(this)' onmouseout='Sale(this)'  > ";
        //            TR += "<td nowrap='nowrap'><img src='img/link.png' border='0' align='absmiddle'  style='margin-left: 4px;'/></td>";
        //            TR += "<td><span style='margin-left:3px;'>";
        //            TR += ObjLenguaje.Label_Home(nombre);
        //            TR += "</span></td>";
        //            TR += "<td align='right'>";

        //            if (isLogin == true)
        //            {
        //                if (Utils.IsUserLogin)
        //                {
        //                    /*
        //                    TR += "<img src='img/plusG.png' border='0' align='absmiddle' style='margin-right:9px;' onmouseover=\"this.src = 'img/plus_hover.png'\"";
        //                    TR += "onmouseout=\"this.src = 'img/plusG.png'\" onclick=\"AbrirModulo('" + URL + "','ESS')\"> ";
        //                      */
        //                    //TR += "<span class='AperturaModulo' onclick=\"AbrirModulo('" + URL + "','ESS')\" > » </span>";
        //                    //TR += "<img src='img/Modulos/SVG/APERTURA_MODULO.svg' class='IconoAperturaModulo'   onclick=\"AbrirModulo('" + URL + "','ESS')\"    >";
        //                    TR += Common.Utils.Armar_Icono("img/Modulos/SVG/APERTURA_MODULO.svg", "IconoAperturaModulo", "", " border='0' onclick=\"AbrirModulo('" + URL + "','ESS')\" ", "");
                            
        //                }
        //            }
        //            else
        //            {/*
        //                TR += "<img src='img/plusG.png' border='0' align='absmiddle' style='margin-right:9px;' onmouseover=\"this.src = 'img/plus_hover.png'\"";
        //                TR += "onmouseout=\"this.src = 'img/plusG.png'\" onclick=\"AbrirModulo('" + URL + "','ESS')\"> ";
        //              */
                     
        //             //   TR += "<img src='img/Modulos/SVG/APERTURA_MODULO.svg' class='IconoAperturaModulo'  onclick=\"AbrirModulo('" + URL + "','ESS')\"    >";
        //                  TR += Common.Utils.Armar_Icono("img/Modulos/SVG/APERTURA_MODULO.svg", "IconoAperturaModulo", "", " border='0' onclick=\"AbrirModulo('" + URL + "','ESS')\" ", "");
                        
        //            }
        //            TR += " </td>";
        //            TR += "</tr>   ";
        //        }
            

        //    return TR;
        //}
        //-------------------------------------------------------


        public string Construir_Acceso(int pos, string activo, string nombre, string URL, string isLogin, string ArchDesc, string idUser)
        {
            string TR = "";

            if (activo == "true")
            {
                TR = "<tr id='Link" + pos + "' onclick=\"Seleccionar('Link" + pos + "','')\"   onmouseover='Sobre(this)' onmouseout='Sale(this)'  > ";
                TR += "<td nowrap='nowrap'><img src='img/link.png' border='0' align='absmiddle'  style='margin-left: 4px;'/></td>";
                TR += "<td><span style='margin-left:3px;'>";
                TR += ObjLenguaje.Label_Home(nombre);
                TR += "</span></td>";
                TR += "<td align='right'>";

                //if (isLogin == "true")
                if (Utils.SessionUserName==idUser)
                {
                    if (Utils.IsUserLogin)
                    {/*
                        TR += "<img src='img/plus.png' border='0' align='absmiddle' style='margin-right:9px;' onmouseover=\"this.src = 'img/plus_hover.png'\"";
                        TR += "onmouseout=\"this.src = 'img/plus.png'\" onclick=\"AbrirModulo('" + URL + "','ESS')\"    width='16' height='16' > ";
                       */
                        
                        //TR += "<img src='img/Modulos/SVG/APERTURA_MODULO.svg' class='IconoAperturaModulo'  onclick=\"AbrirModulo('" + URL + "','ESS')\"    >";
                        TR += Common.Utils.Armar_Icono("img/Modulos/SVG/APERTURA_MODULO.svg", "IconoAperturaModulo", "", " border='0' onclick=\"AbrirModulo('" + URL + "','ESS')\" ", "");
                        
                    }
                }
                else
                {/*
                    TR += "<img src='img/plus.png' border='0' align='absmiddle' style='margin-right:9px;' onmouseover=\"this.src = 'img/plus_hover.png'\"";
                    TR += "onmouseout=\"this.src = 'img/plus.png'\" onclick=\"AbrirModulo('" + URL + "','ESS')\"   width='16' height='16' > ";
                  */
                    //TR += "<span class='AperturaModulo' onclick=\"AbrirModulo('" + URL + "','ESS')\" > » </span>";
                    //TR += "<img src='img/Modulos/SVG/APERTURA_MODULO.svg' class='IconoAperturaModulo'  onclick=\"AbrirModulo('" + URL + "','ESS')\"    >";
                    TR += Common.Utils.Armar_Icono("img/Modulos/SVG/APERTURA_MODULO.svg", "IconoAperturaModulo", "", " border='0' onclick=\"AbrirModulo('" + URL + "','ESS')\" ", "");
                }
                TR += " </td>";
                TR += "</tr>   ";
            }
           
            return TR;
        }




        public int IncrementaPosmenu()
        {
            posmenu++;
            return posmenu;
        }
 
         public void ActualizarContenedor(object sender, CommandEventArgs e)
        {
          
            String arg = (String)e.CommandArgument;
            ContPpal.EnviarInfoContenedorPrincipal(arg);
            //Habilito el armado de los submenues
            //ScriptManager.RegisterStartupScript(this, typeof(Page), "Logo_InicializaMenu", "$(function() {  $('#main-menu').smartmenus({  subMenusSubOffsetX: 0, subMenusSubOffsetY: -1  }); });  ", true);
            //ScriptManager.RegisterStartupScript(this, typeof(Page), "Logo_InicializaMenuTop", "$(function() {  $('#main-menuTop').smartmenus({  subMenusSubOffsetX: 0, subMenusSubOffsetY: -1, mainMenuSubOffsetX:0,mainMenuSubOffsetY:0,subMenusMinWidth:'60px', subMenusMaxWidth:'900px'  }); });  ", true);
            //Habilito el armado de los submenues
            ScriptManager.RegisterStartupScript(this, typeof(Page), "InicializaMenu", "$(function() {  $('#main-menu').smartmenus({  subMenusSubOffsetX: 0, subMenusSubOffsetY: -1 , hideOnClick: false  }); });  ", true);
            ScriptManager.RegisterStartupScript(this, typeof(Page), "InicializaMenuTop", "$(function() {  $('#main-menuTopLoguin').smartmenus({  subMenusSubOffsetX: 0, subMenusSubOffsetY: -1, mainMenuSubOffsetX:0,mainMenuSubOffsetY:0,subMenusMinWidth:'60px', subMenusMaxWidth:'1060px', hideOnClick: true   }); });  ", true);
            
            

        }

        public void ActualizaGadgets(object sender, EventArgs e)
        {
            ContPpal.ActualizaGadgets(sender, e);
        }

        public void AsignarContPpal(RHPro.Controls.ContenedorPrincipal CP) {
            ContPpal = CP;
        }

        public void Actualizar_Accesos_XML(object sender, CommandEventArgs e)
        {         
            //int nroAcceso = int.Parse((String)e.CommandArgument);
            string nroAcceso = (String) e.CommandArgument;
            ContPpal.Actualizar_Accesos_XML(nroAcceso);
        }

        public void Actualizar_Accesos(object sender, CommandEventArgs e)
        {         
            //int nroAcceso = int.Parse((String)e.CommandArgument);
            string nroAcceso = (String) e.CommandArgument;
            //ContPpal.Actualizar_Accesos(nroAcceso);
        }

        protected string Visibilidad(bool condicion)
        {
            if (condicion)
                return "visible";
            else
                return "hidden";
        }
       

    }
}