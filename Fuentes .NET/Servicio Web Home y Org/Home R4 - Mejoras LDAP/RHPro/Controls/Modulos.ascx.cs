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

        /// <summary>
        /// Modulos disponibles
        /// </summary>
        protected List<Module> AvailableModules
        {
            get { return ViewState["AvailableModules"] as List<Module>; }
            set { ViewState["AvailableModules"] = value; }
        }

        protected void Page_Load(object sender, EventArgs e)
        {        


            ObjLenguaje = new RHPro.Lenguaje();
            Cargar_Accesos();
        }


        protected void Page_PreRender(object sender, EventArgs e)
        {
            if (bool.Parse(ConfigurationManager.AppSettings["VisualizarComplementos"]))
            { LinkButton1.Text = ObjLenguaje.Label_Home("Gadgets"); }
         
            LoadModule();
           
        }


        public void LoadModule()
        {
            ServicesProxy.rhdesa.Consultas cc = new ServicesProxy.rhdesa.Consultas();
            //Response.Write(cc.constr(Utils.SessionBaseID));
            
         
            /*
            //Recupero los modulos disponibles para el usuario y base seleccionados
            AvailableModules = ModuleServiceProxy.Find(Utils.SessionUserName, Utils.SessionBaseID, Thread.CurrentThread.CurrentCulture.Name);
            //Cargo el repeater con los modulos disponibles
            Repeater1.DataSource = AvailableModules;
            Repeater1.DataBind();
            */             
            //Recupero los modulos disponibles para el usuario Anonimo
            AvailableModules = ModuleServiceProxy.Find(Utils.SessionUserName, Utils.SessionBaseID, Thread.CurrentThread.CurrentCulture.Name);
            //Cargo el repeater con los modulos disponibles
            Repeater1.DataSource = AvailableModules;
            Repeater1.DataBind();

            //Recupero los modulos habilitados para el usuario y base seleccionados. 
            AvailableModules = ModuleServiceProxy.Find_Modulos(Utils.SessionUserName, Utils.SessionBaseID, Thread.CurrentThread.CurrentCulture.Name, true);
            //Cargo el repeater con los modulos
            Repeater2.DataSource = AvailableModules;
            Repeater2.DataBind();

            //Recupero los modulos inhabilitados para el usuario y base seleccionados
            AvailableModules = ModuleServiceProxy.Find_Modulos(Utils.SessionUserName, Utils.SessionBaseID, Thread.CurrentThread.CurrentCulture.Name, false);
            //Cargo el repeater con los modulos
            Repeater3.DataSource = AvailableModules;
            Repeater3.DataBind();

 
            
             
        }

        public string TituloDeMenu(string titulo) {
            return ObjLenguaje.Label_Home(titulo);
        }

        public string Imprimir_Action(string accion, string MenuName)
        {
            string TAG_IMG = "";

            //Verifico si la accion tiene contenido
            if ((accion != "#") && (accion != ""))
            {
                TAG_IMG = "<img src='img/plusG.png' border='0' width='17' height='17' align='absmiddle' onmouseover=\"this.src = 'img/plusG.png'\"  onmouseout=\"this.src = 'img/plusG.png'\" onclick=\"AbrirModulo(" + accion + ",'" + MenuName + "')\">  ";
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

            string BaseId = Common.Utils.SessionBaseID;
            string UserName = Common.Utils.SessionUserName;

            /* ****** Preparo el filtro de accesos que deseo visualizar desde la base****** */
           /*             
            string sql = "SELECT * FROM Home_Accesos ";
            sql += " WHERE Activo = 1 ";
            if (Common.Utils.IsUserLogin) //Busco todos los accesos configurados para el usuario + los anonimos
                sql += " AND idUser = '" + UserName + "' OR idUser IS NULL";
            else
                sql += " AND idUser IS NULL "; //Busco todos los accesos anonimos
            sql += " ORDER BY Nombre ASC ";

            Consultas cc = new Consultas();
            DataSet ds = cc.get_DataSet(sql, BaseId);

            RepAccesos.DataSource = ds;           
            RepAccesos.DataBind();
            */
            /* ****** Preparo el filtro de accesos que deseo visualizar desde el XML****** */
            string filtro = "Activo=true ";
            DataView dv = DS_Datos_XML().Tables[0].DefaultView;
            if (Common.Utils.IsUserLogin) //Busco todos los accesos configurados para el usuario + los anonimos
                //filtro += " AND (idUser = '" + UserName + "' OR idUser = '') ";               
                filtro += " AND idUser = '" + UserName + "'";               
            else
                filtro += " AND idUser=''";  //Busco todos los accesos anonimos
                
            dv.RowFilter = filtro;
            dv.Sort = "Nombre ASC";
            /* *************************************************************** */
           
            RepAccesos.DataSource = dv;   
            //RepAccesos.DataSource = DS_Datos_XML();              
            RepAccesos.DataBind();            
            
        }


        public String ImprimirLink(bool isLogin, string URL){
         string link = "";
         if (isLogin == true)
         {
             if (Common.Utils.IsUserLogin)
             {
                 link = "<img src='img/plus.png' border='0' align='absmiddle' style='margin-right:9px;' onmouseover=\"this.src = 'img/plus.png'\"";
                 link += "  width='16' height='16'  onmouseout=\"this.src = 'img/plus.png'\" onclick=\"AbrirModulo('" + URL + "','ESS')\">";
             }
         }
         else {
             link = "<img src='img/plus.png' border='0' align='absmiddle' style='margin-right:9px;' onmouseover=\"this.src = 'img/plus.png'\"";
             link += "  width='16' height='16'   onmouseout=\"this.src = 'img/plus.png'\" onclick=\"AbrirModulo('" + URL + "','ESS')\">";
         }

         return link;

        }


        //public string Accesos_Home()
        //{
        //    string Misql = "SELECT * FROM Home_Accesos WHERE Activo = 1 ORDER BY Nombre ASC ";
        //    Consultas cc = new Consultas();
        //    DataTable dt = cc.get_DataTable(Misql, Utils.SessionBaseID);
        //    string Salida = "";
        //    foreach (System.Data.DataRow dr in dt.Rows)
        //     {
        //         if (!dr["Nombre"].Equals(System.DBNull.Value))
        //         {
        //             Salida += Accesos(posmenu, (bool)dr["Activo"], (string)dr["Nombre"], (string)dr["URL"], (bool)dr["isLogin"], (string)dr["archivoDescripcion"]);
        //             posmenu++;
        //         }
        //     }
        //    return Salida;            
        // }


        public string Accesos(int pos, bool activo, string nombre, string URL, bool isLogin, string ArchDesc)
        {
            string TR = "";
             
                if (activo == true)
                {
                    TR = "<tr id='Link" + pos + "' onclick=\"Seleccionar('Link" + pos + "','')\"   onmouseover='Sobre(this)' onmouseout='Sale(this)'  > ";
                    TR += "<td nowrap='nowrap'><img src='img/link.png' border='0' align='absmiddle'  style='margin-left: 4px;'/></td>";
                    TR += "<td><span style='margin-left:3px;'>";
                    TR += ObjLenguaje.Label_Home(nombre);
                    TR += "</span></td>";
                    TR += "<td align='right'>";

                    if (isLogin == true)
                    {
                        if (Utils.IsUserLogin)
                        {
                            TR += "<img src='img/plusG.png' border='0' align='absmiddle' style='margin-right:9px;' onmouseover=\"this.src = 'img/plus_hover.png'\"";
                            TR += "onmouseout=\"this.src = 'img/plusG.png'\" onclick=\"AbrirModulo('" + URL + "','ESS')\"> ";
                        }
                    }
                    else
                    {
                        TR += "<img src='img/plusG.png' border='0' align='absmiddle' style='margin-right:9px;' onmouseover=\"this.src = 'img/plus_hover.png'\"";
                        TR += "onmouseout=\"this.src = 'img/plusG.png'\" onclick=\"AbrirModulo('" + URL + "','ESS')\"> ";
                    }
                    TR += " </td>";
                    TR += "</tr>   ";
                }
            

            return TR;
        }
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
                    {
                        TR += "<img src='img/plus.png' border='0' align='absmiddle' style='margin-right:9px;' onmouseover=\"this.src = 'img/plus_hover.png'\"";
                        TR += "onmouseout=\"this.src = 'img/plus.png'\" onclick=\"AbrirModulo('" + URL + "','ESS')\"    width='16' height='16' > ";
                    }
                }
                else
                {
                    TR += "<img src='img/plus.png' border='0' align='absmiddle' style='margin-right:9px;' onmouseover=\"this.src = 'img/plus_hover.png'\"";
                    TR += "onmouseout=\"this.src = 'img/plus.png'\" onclick=\"AbrirModulo('" + URL + "','ESS')\"   width='16' height='16' > ";
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
            ContPpal.ActualizaContenido(arg);
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
            ContPpal.Actualizar_Accesos(nroAcceso);
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