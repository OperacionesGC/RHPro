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
using ServicesProxy.rhdesa;
using ServicesProxy;
using System.Threading;
using System.Collections.Generic;


namespace RHPro.Controls
{
    public partial class ContenedorPrincipal : System.Web.UI.UserControl
    {

        public RHPro.Lenguaje ObjLenguaje;
        //public Hashtable Hash_RecorridoMenu;
        //public Hashtable Hash_DatosMenu;
        //public Menu NavigationMenu;
        public Hashtable Hash_Etiquetas;
        //public List<String> Hash_ModulosActivos;
        public String Max_AnchoGadget = "99%";
        public String Min_AnchoGadget = "49%";
       
        public static Default Padre = new Default();

        
        
    

        protected void Page_Load(object sender, EventArgs e)
        {
            ObjLenguaje = new Lenguaje();
           // Loader_Gadgets.InnerText = ObjLenguaje.Label_Home("Cargando");
            if (String.IsNullOrEmpty(Utils.Session_ModuloActivo))
                Utils.Session_ModuloActivo ="RHPROX2";
                                  
        }

        public void InicializarPadre(Default P)
        {
            Padre = P;

        }


        public void ActualizaGadgets(object sender, EventArgs e)
        {
            //Update_Gadget(2);
//            ActualizaContenido((String)System.Web.HttpContext.Current.Session["ActualizaModulo"]);
        }


        protected void Page_PreRender(object sender, EventArgs e)
        {
             
            if (  ((String)System.Web.HttpContext.Current.Session["ActualizaAcceso"] != "-1") 
                 && ((String)System.Web.HttpContext.Current.Session["ActualizaAcceso"]!= "")
                 && ((String)System.Web.HttpContext.Current.Session["ActualizaAcceso"] != null) 
                 )
            {               
                Actualizar_Accesos_XML((String)System.Web.HttpContext.Current.Session["ActualizaAcceso"]); 
            } 
            else {




                if (((String)System.Web.HttpContext.Current.Session["ActualizaModulo"] == "-1")
                     || ((String)System.Web.HttpContext.Current.Session["ActualizaModulo"] == "")
                     || ((String)System.Web.HttpContext.Current.Session["ActualizaModulo"] == null)
                     )
                {
                    Update_Gadget(1);
                   
                }
                else
                    ActualizaContenido((String)System.Web.HttpContext.Current.Session["ActualizaModulo"]);

                   }
            
        }

        /// <summary>
        /// Verfica si el modulo tiene algun menu actualizado
        /// </summary>
        /// <returns></returns>
        private Boolean MenuModificadoEnModulo(string modulo)
        {
            bool salida = false;

            try
            {
                Consultas cc = new Consultas();
                String sql = "SELECT menufecmodif FROM menuraiz WHERE menudesc = '" + modulo + "' AND NOT menufecmodif is null";
                DataTable dt = cc.get_DataTable(sql, Utils.SessionBaseID);
                if (dt!=null)
                    if (dt.Rows.Count>0)
                    {
                        if (!DBNull.Value.Equals(dt.Rows[0]["menufecmodif"]))
                        {                            
                            if (Convert.ToString(Session["RHPRO_menufecmodif_" + modulo]) != Convert.ToString(dt.Rows[0]["menufecmodif"]))
                            {
                                Session["RHPRO_menufecmodif_" + modulo] = Convert.ToString(dt.Rows[0]["menufecmodif"]);
                                return true;
                            }
                            else return false;
                             
                        }
                        else return false;
                    }                
            }
            catch { }

            return salida;
        }
        

        /**********************************************************************/
        private void ArmoElMenu(String NombreModulo)
        {
            if (Convert.ToString(Session["RHPRO_ListaPerfUsr"]) == "")
            {
                Usuarios Usr = new Usuarios();
                Session["RHPRO_ListaPerfUsr"] = Usr.getPerfilesUsuario(Utils.SessionUserName);              
            }

            String modulo = "";            
            modulo = Utils.getMenuDir(NombreModulo);
            ConsultaDatos C_datos = new ConsultaDatos();
            Utils.Session_ModuloActivo = modulo;

            if (Convert.ToString(Session["RHPRO_LenguajeActivo"]) == "")
                Session["RHPRO_LenguajeActivo"] = Convert.ToString(System.Web.HttpContext.Current.Session["Lenguaje"]);

            
            //if (Convert.ToString(Session["RHPRO_Home_MenuPrincipal_" + modulo]) == "")//Si aun no ingrese al modulo, armo su menu
            //Validar si es la primera vez que ingreso al modulo o se ha modificado alguna entrada de menu del modulo
            if ( (Convert.ToString(Session["RHPRO_Home_MenuPrincipal_" + modulo]) == "") 
                || MenuModificadoEnModulo(modulo))
            {
                DataTable menuData = GetMenuData(modulo);
                DataTable gruporest = C_datos.Grupo_Restricciones();
              
                AddTopMenuItems(menuData, modulo, NombreModulo, gruporest);

            }
            else//Si ya ingrese al modulo en algun momento, entonces tomo el menu ya armado
            {
                if (Convert.ToString(Session["RHPRO_LenguajeSeleccionado"]) != "")
                {
                    //if (Convert.ToString(Session["RHPRO_LenguajeActivo"]) != Convert.ToString(Session["RHPRO_LenguajeSeleccionado"]))
                    if (Convert.ToString(Session["RHPRO_Home_MenuPrincipal_Idioma_" + modulo])!=Convert.ToString(Session["RHPRO_LenguajeSeleccionado"]))
                    {
                        DataTable menuData = GetMenuData(modulo);
                        DataTable gruporest = C_datos.Grupo_Restricciones();
                      
                        AddTopMenuItems(menuData, modulo, NombreModulo, gruporest);
                        Session["RHPRO_LenguajeActivo"] = Convert.ToString(Session["RHPRO_LenguajeSeleccionado"]);
                    }
                    else
                    {
                        MenuPrincipalModulo.Controls.Clear();
                        MenuPrincipalModulo.Controls.Add(new LiteralControl(Convert.ToString(Session["RHPRO_Home_MenuPrincipal_" + modulo])));
                        MenuPrincipalModulo.DataBind();
                    }
                }
                else//En el caso que no haya cambiado de idioma armo el menu con lo que tengo en la variable de sesion
                {
                    MenuPrincipalModulo.Controls.Clear();
                    MenuPrincipalModulo.Controls.Add(new LiteralControl(Convert.ToString(Session["RHPRO_Home_MenuPrincipal_" + modulo])));
                    MenuPrincipalModulo.DataBind();
                }
            }

            

            //Habilito el armado de los submenues
            ScriptManager.RegisterStartupScript(this, typeof(Page), "InicializaMenu", "$(function() {  $('#main-menu').smartmenus({  subMenusSubOffsetX: 0, subMenusSubOffsetY: -1 , hideOnClick: false  }); });  ", true);
            ScriptManager.RegisterStartupScript(this, typeof(Page), "InicializaMenuTop", "$(function() {  $('#main-menuTopLoguin').smartmenus({  subMenusSubOffsetX: 0, subMenusSubOffsetY: -1, mainMenuSubOffsetX:0,mainMenuSubOffsetY:0,subMenusMinWidth:'60px', subMenusMaxWidth:'1060px', hideOnClick: true   }); });  ", true);
        }

        private DataTable GetMenuData(String modulo)
        {            
            Consultas cc = new Consultas();
            String TipoBD = cc.get_TipoBase(Utils.SessionBaseID);
            string sql = " SELECT MenuName, MenuOrder, Parent ParentId, tipo, action, menuaccess, menuimg, '0', menumsnro, menuraiz menunro ";
            if (TipoBD == "MSSQL")
            {
                sql += " , (select top(1)  LE." + System.Web.HttpContext.Current.Session["Lenguaje"] + " from lenguaje_etiqueta LE where LE.etiqueta=menumstr.menuname COLLATE Modern_Spanish_CS_AS  and ( LE.modulo='" + modulo + "' Or LE.modulo is null) ORDER BY LE.modulo DESC ) TraduccionEtiqueta   ";
                sql += " , (select top(1)  LE.esAR from lenguaje_etiqueta LE where LE.etiqueta=menumstr.menuname COLLATE Modern_Spanish_CS_AS  and ( LE.modulo='" + modulo + "' Or LE.modulo is null) ORDER BY LE.modulo DESC ) TraduccionEtiquetaESAR   ";
            }
            else
            {
                sql += " , (select  LE." + System.Web.HttpContext.Current.Session["Lenguaje"] + " from lenguaje_etiqueta LE where rownum=1 AND LE.etiqueta=menumstr.menuname   and ( LE.modulo='" + modulo + "' Or LE.modulo is null)  ) TraduccionEtiqueta   ";
                sql += " , (select  LE.esAR from lenguaje_etiqueta LE where rownum=1 AND LE.etiqueta=menumstr.menuname    and ( LE.modulo='" + modulo + "' Or LE.modulo is null) ) TraduccionEtiquetaESAR   ";
            }

            sql += " , MR.menunro MenuRaiz ";
            sql += " , MR.menudesc ";
            sql += " , MR.menudir ";            
            sql += " , MR.* ";
            sql += " FROM menumstr ";
            sql += " INNER JOIN menuraiz MR ON menumstr.menuraiz = MR.menunro";                       
            sql += " WHERE Upper(MR.menudesc) = Upper('" + modulo + "')"; 
            sql += " ORDER BY parent desc, menuorder ";
            
            DataTable dt = cc.get_DataTable(sql, Utils.SessionBaseID);
            return dt;
                       
        }
 

        private void AddTopMenuItems(DataTable menuData, String modulo, String NombreModulo, DataTable GrupoRestriccion)
        {
            String Salida = "";
            String Cabecera ="";

            try
            {
                if (Convert.ToString(Session["RHPRO_menufecmodif_" + modulo]) == "")
                    Session["RHPRO_menufecmodif_" + modulo] = menuData.Rows[0]["menufecmodif"];
            }
            catch { }
            
   
            DataView view = new DataView(menuData);                 
            //view.RowFilter = " ParentId = '" + Utils.getMenuDesc(NombreModulo) + "' ";
            view.RowFilter = " ParentId = '" + Convert.ToString(menuData.Rows[0]["menudesc"]) + "' "; //***HABILITAR PARA CAMBIO MENU****///           
            Cabecera += "<DIV class='ContenedorBarraNavegacion' >";
            Cabecera += "<UL id='main-menu' class='sm sm-blue smblue_ppal' onclick=\"this.style.zIndex=400; if (document.getElementById('main-menuTopLoguin')){document.getElementById('main-menuTopLoguin').style.zIndex=300; $('#main-menuTopLoguin').smartmenus('menuHideAll');}  \"  >";              
            //Cabecera += "<UL id='main-menu' class='sm sm-blue smblue_ppal' onclick=\"this.style.zIndex = 300; if (document.getElementById('main-menu')) { document.getElementById('main-menuTop').style.zIndex = 400; }\" onmouseover=\"this.style.zIndex=300; if (document.getElementById('main-menuTop')){document.getElementById('main-menuTop').style.zIndex=400;}  \">";              
            // Cabecera += "<UL id='main-menu' class='sm sm-blue smblue_ppal'>";              
            
            //Salida += "<UL id='main-menu' class='sm sm-blue smblue_ppal'  style='padding-left:5px;' >";              
            
            String EtiqTraucida = "";
            //Recupero la lista de perfiles del usuario
            List<String> ListaPerfUsr = (List<String>)Session["RHPRO_ListaPerfUsr"];  

            //Agrego el modulo a la lista de modulos habilitados
            ((List<String>)Utils.Modulos_Habilitados).Add("RHPROX2");           
            
            String SalidaAyuda = "";

            String MenuRaiz = Convert.ToString(menuData.Rows[0]["menudir"]);
            
            ConsultaDatos C_datos = new ConsultaDatos();
            
            foreach (DataRowView row in view)
            { 
              //  if (Utils.Habilitado(ListaPerfUsr, Convert.ToString(row["menuaccess"])))
                bool Habilitado = C_datos.Menu_Habilitado(Convert.ToString(row["menuaccess"]), Convert.ToInt32(row["menumsnro"]));
                                               
                if (Habilitado)
                {
                    
                    //Agrego el modulo a la lista de modulos habilitados
                    ((List<String>)Utils.Modulos_Habilitados).Add(modulo);


                   // if (!DBNull.Value.Equals(row["TraduccionEtiqueta"]))
                   if ((!DBNull.Value.Equals(row["TraduccionEtiqueta"])) && (!String.IsNullOrEmpty(Convert.ToString(row["TraduccionEtiqueta"]))))
                        EtiqTraucida = Convert.ToString(row["TraduccionEtiqueta"]);
                    else
                    {
                        //if (!DBNull.Value.Equals(row["TraduccionEtiquetaESAR"]))
                        if ((!DBNull.Value.Equals(row["TraduccionEtiquetaESAR"])) && (!String.IsNullOrEmpty(Convert.ToString(row["TraduccionEtiquetaESAR"]))))

                            EtiqTraucida = Convert.ToString(row["TraduccionEtiquetaESAR"]);
                        else
                            EtiqTraucida = Convert.ToString(row["MenuName"]);
                    }

                    if (Convert.ToString(row["MenuName"]).ToUpper() == "AYUDA")
                    {
                        SalidaAyuda += " <LI style=''><a   class='BtnTransparenteAyuda'> ";                        
                        SalidaAyuda += " </a>";
                      
                        SalidaAyuda += Utils.Armar_Icono("img/Modulos/SVG/" + NombreModulo + ".svg", "IconosBarraAyuda", "", "align='absmiddle'", "", "$('#main-menuTopLoguin').smartmenus('menuHideAll'); ");

                        SalidaAyuda += AddChildMenuItems(menuData, Convert.ToString(row["MenuOrder"]), MenuRaiz, NombreModulo, ListaPerfUsr, GrupoRestriccion);
                        SalidaAyuda += " </LI> ";

                    }
                    else
                    { 
                         Salida += " <LI><a> ";

                        if (Convert.ToString(row["menuimg"]) != "")
                        { Salida += " <img src='../" + Convert.ToString(row["menuimg"]) + "' border='0' class='IconoAcceso'    > "; }

                        Salida += EtiqTraucida;
                         

                        Salida += " </a>";
                        //Salida += AddChildMenuItems(menuData, Convert.ToString(row["MenuOrder"]), modulo, NombreModulo, ListaPerfUsr);
                        Salida += AddChildMenuItems(menuData, Convert.ToString(row["MenuOrder"]), MenuRaiz, NombreModulo, ListaPerfUsr, GrupoRestriccion);
                        Salida += " </LI> ";
                    }


                }
            }
            
            Salida = SalidaAyuda + Salida;
            Salida = Cabecera + Salida;
            /*****************************************/
            
            Salida += "</UL>";                 
            Salida += "</DIV>";
 

            Session["RHPRO_Home_MenuPrincipal_" + modulo] = Salida;
            Session["RHPRO_Home_MenuPrincipal_Idioma_" + modulo] = Convert.ToString(Session["RHPRO_LenguajeSeleccionado"]);
           // ScriptManager.RegisterStartupScript(this, typeof(Page), "InicializaMenu", "$(function() {  $('#main-menu').smartmenus({  subMenusSubOffsetX: 0, subMenusSubOffsetY: -1  }); });  ", true);

            MenuPrincipalModulo.Controls.Clear();            
            MenuPrincipalModulo.Controls.Add(new LiteralControl(Salida));
            MenuPrincipalModulo.DataBind();
        }




        private string AddChildMenuItems(DataTable menuData, String MenuOrder, String modulo, String NombreModulo, List<String> ListaPerfUsr, DataTable GrupoRestriccion)
        {
            DataView view = new DataView(menuData);
            ConsultaDatos C_datos = new ConsultaDatos();

            //view.RowFilter = "ParentId = '" + MenuOrder + modulo + "'";  
            //view.RowFilter = " ParentId = '" + MenuOrder + Utils.getMenuDesc(NombreModulo) + "' ";
            view.RowFilter = " ParentId = '" + MenuOrder + Convert.ToString(menuData.Rows[0]["menudesc"]) + "' ";//***HABILITAR PARA CAMBIO MENU****///
            
            String Salida = "";
            String EtiqTraucida = "";
            String LogoMenu = "";
            bool paso = false;
            bool Habilitado;
            string action = "";
            string ErrorAction = "";
            String[] Arr_Err;
            String ErrTraducidos = "";
          
            foreach (DataRowView row in view)
            {
                ErrorAction = "";
                ErrTraducidos = "";
               //Verifico si el menu esta habilitado y el usuario no esta inhabilitado por grupo de restricciones
                Habilitado = C_datos.Menu_Habilitado(Convert.ToString(row["menuaccess"]), Convert.ToInt32(row["menumsnro"]));
                //Habilitado = (Utils.Habilitado(ListaPerfUsr, Convert.ToString(row["menuaccess"])) && (C_datos.Habilitado_Por_GrupoRestricciones(Convert.ToInt32(row["menumsnro"]), ListaPerfUsr, GrupoRestriccion)));

                if (Habilitado)       
                {
                    if (!String.IsNullOrEmpty(Convert.ToString(row["MenuName"])))
                    {
                      /* if (!DBNull.Value.Equals(row["TraduccionEtiqueta"]))                        
                            EtiqTraucida = Convert.ToString(row["TraduccionEtiqueta"]);
                        else
                            EtiqTraucida = Convert.ToString(row["MenuName"]);
                       */
                       
                        if ((!DBNull.Value.Equals(row["TraduccionEtiqueta"])) && (!String.IsNullOrEmpty(Convert.ToString(row["TraduccionEtiqueta"]))))
                            EtiqTraucida = Convert.ToString(row["TraduccionEtiqueta"]);
                        else
                        {
                            //if (!DBNull.Value.Equals(row["TraduccionEtiquetaESAR"]))
                            if ((!DBNull.Value.Equals(row["TraduccionEtiquetaESAR"])) && (!String.IsNullOrEmpty(Convert.ToString(row["TraduccionEtiquetaESAR"]))))

                                EtiqTraucida = Convert.ToString(row["TraduccionEtiquetaESAR"]);
                            else
                                EtiqTraucida = Convert.ToString(row["MenuName"]);
                        }
                   

                        if (paso == false)
                            Salida = "<UL>";                        
                  
                        Salida += " <LI ";
                        action = Convert.ToString(row["action"]);
                        ErrorAction = Utils.Valida_Javascript(action, false);
                        if (ErrorAction != "")
                        {
                            Arr_Err = System.Text.RegularExpressions.Regex.Split(ErrorAction, ",");
                            ErrTraducidos = "";
                            foreach (string s in Arr_Err)
                            {
                                if (ErrTraducidos == "")
                                    ErrTraducidos = ObjLenguaje.Label_Home(s.Trim());
                                else ErrTraducidos += "," + ObjLenguaje.Label_Home(s.Trim());
                            }

                            ErrorAction = ErrTraducidos;
                        }
                    
                        if (action != "#")
                        {
                            if (ErrorAction == "")
                              Salida += " onclick =\"" + Utils.ArmarAction(Convert.ToString(row["action"]), modulo, Convert.ToString(row["menumsnro"]), Convert.ToString(row["MenuRaiz"]), Convert.ToString(row["menunro"]), "") + "\"  ";
                            else
                                Salida += " onclick=\" alert('" + ObjLenguaje.Label_Home("Error") + ":\\n" + ObjLenguaje.Label_Home("Menu mal armado") + ".\\n" + ObjLenguaje.Label_Home("Consulte con el administrador") + ".') \" ";
                        }

                        Salida += "><a ";

                        if (ErrorAction != "")
                            Salida += " style='color:#CCCCCC !important' title='" + ObjLenguaje.Label_Home("Error") + ": \n" + ObjLenguaje.Label_Home("Menu mal armado") + ".\n" + ObjLenguaje.Label_Home("Consulte con el administrador")  + ".' >   ";

                        else Salida += "> ";


                        if (Convert.ToString(row["menuimg"]) != "")
                        {
                            LogoMenu = Convert.ToString(row["menuimg"]);
                            if (LogoMenu.Contains("../"))
                                Salida += " <img src='../" + LogoMenu + "' border='0' class='IconoAcceso'    > ";
                            else
                                Salida += " <img src='../shared/images/" + LogoMenu + "' border='0' class='IconoAcceso'    > ";
                        }
                        else
                            Salida += "<span class='SeparadorIconoMenu' > </span> ";


                        Salida += EtiqTraucida;

                        if ((Convert.ToString(row["action"]) == "#") || (Convert.ToString(row["action"]) == ""))
                            Salida += "<span class='sub-arrow'>+</span>";

                        Salida += " </a>";

                        Salida += AddChildMenuItems(menuData, Convert.ToString(row["MenuOrder"]), modulo, NombreModulo, ListaPerfUsr, GrupoRestriccion);
                        Salida += " </LI> " + System.Environment.NewLine;
                        paso = true;
                    }
                }
            }

            if (paso == true)
                Salida += "</UL>" + System.Environment.NewLine;

            return Salida;
        }

        /********************************************************************/
        
        
        /**********************************************************************/


        public void EnviarInfoContenedorPrincipal(String DatosDelModulo)
        {
             
           // String[] Datos = DatosDelModulo.Split(new Char[] { '@' });
            String[] Datos = DatosDelModulo.Split('@');

            Session["RHPRO_NombreModulo"] = Utils.getMenuDir(Datos[0]);
            Utils.Session_MenumsNro_Modulo = Datos[1];
            //Session["RHPRO_MenumsNro_Modulo"] = Datos[1];
            ActualizaContenido(Datos[0]);

            //Mantiene la barra de los modulos Expandida o Colapsada
            ScriptManager.RegisterStartupScript(Page, GetType(), "ControlExpandido", "setTimeout('if (ControlExpand==1){ControlExpand=0;} else {ControlExpand=1;};DesplazarBarraMenu()', 30);", true);


        }


        public void ActualizaContenido(String modulo) {
                        
            //Session["RHPRO_menunro"] = getMenuDir(modulo);
            
           //ARMA EL HASH DEL MENU-----           
            
            if (Common.Utils.IsUserLogin)
            {
                if (((Hashtable)Session["VisualizaModulos"]).Contains(modulo))
                {
                    if ((Boolean)((Hashtable)Session["VisualizaModulos"])[modulo])
                    {
                        try
                        {
                            ArmoElMenu(modulo);
                        }
                        catch (Exception ex)
                        {
                            
                            //ScriptManager.RegisterStartupScript(this, typeof(Page), "ArmadoMenuErr", " alert('Directirio incorrecto');", true);
                        }
                      
                        Update_Gadget(2);
                    }
                 
                }

                System.Web.HttpContext.Current.Session["ActualizaAcceso"] = "-1";
                System.Web.HttpContext.Current.Session["ActualizaModulo"] = modulo;

                   
            }
            else 
                Publicar_Desc_Modulo(modulo);

        }

       


        public void Publicar_Desc_Modulo(String modulo)
        {
            System.Web.HttpContext.Current.Session["ActualizaAcceso"] = "-1";
            System.Web.HttpContext.Current.Session["ActualizaModulo"] = modulo;

            String accion = "";
            String desabr = "";
            String linkmanual = "";
            String linkdvd = "";
            String menudetalle = "";
            string Misql = "";
            bool puede;


            //Vacio el panel de modulos
            MiPanel.Controls.Clear();
            MiPanel.Visible = false;

            Cuerpo.Visible = true;


            //Misql = "SELECT menudetalle,menudesabr,action,linkmanual,linkdvd,menuname FROM menumstr WHERE (menudetalle IS NOT NULL) AND  menuname = '" + modulo + "'";
            Misql = "SELECT menudetalle,menudesabr,action,linkmanual,linkdvd,menuname FROM menumstr WHERE   menuname = '" + modulo + "'";
             
            Consultas cc = new Consultas();          
           
            DataTable dt = cc.get_DataTable(Misql, Utils.SessionBaseID);

            foreach (System.Data.DataRow dr in dt.Rows)
            {
                menudetalle = "";
              
                if (!dr["menudetalle"].Equals(System.DBNull.Value))
                {
                    menudetalle = ObjLenguaje.Traducir_Modulo((String)dr["menudetalle"], (String)dr["menuname"]);
                }

                if (!dr["menudesabr"].Equals(System.DBNull.Value))
                    desabr = (String)dr["menudesabr"];
                if (!dr["action"].Equals(System.DBNull.Value))
                    accion = (String)dr["action"];
                if (!dr["linkmanual"].Equals(System.DBNull.Value))
                    linkmanual = (String)dr["linkmanual"];
                if (!dr["linkdvd"].Equals(System.DBNull.Value))
                    linkdvd = (String)dr["linkdvd"];
                //Antes de imprimir el encabezado verifica si puede acceder al modulo. Si no puede directamente no muestra el acceso
                puede = ModuleServiceProxy.Puede_Acceder(Utils.SessionUserName, Utils.SessionBaseID, Utils.Lenguaje, dr["menuname"].ToString());
                //puede = ModuleServiceProxy.Puede_Acceder(Utils.SessionUserName, Utils.SessionBaseID, Thread.CurrentThread.CurrentCulture.Name, dr["menuname"].ToString());

//              menudetalle =  ObjLenguaje.Traducir_Modulo((String)dr["menudetalle"], (String)dr["menuname"]);
                Cuerpo.InnerHtml = TopeInfoModulos(desabr, modulo, accion, linkmanual, linkdvd, puede, menudetalle.Replace(".", ".<BR><BR>"));

                
            }
        }
 

        public void Update_Gadget(int IdContenedor) {         
         
             ConfiguracionesHome ch = new ConfiguracionesHome();
             if (ch.Gadgets_Habilitados())                     
            {
                System.Web.HttpContext.Current.Session["ActualizaModulo"] = "-1";
                System.Web.HttpContext.Current.Session["ActualizaAcceso"] = "-1";
                Control GadgetControl;
                String urlControl;
              
                Cuerpo.Visible = false;
                //Vacio el panel de modulos
                MiPanel.Controls.Clear();
                MiPanel.Visible = true;
                Gadgets_Del_Modulo.Controls.Clear();
                Gadgets_Del_Modulo.Visible = true;

                string BaseId = Common.Utils.SessionBaseID;
                string sql;
                Consultas cc;
                sql = "";
                String TipoBD = "MSSQL";

                if (Utils.IsUserLogin)
                {
                    string GadgetPermitidos = Padre.Lista_Gadget_Permitidos(Utils.SessionUserName);
                    cc = new Consultas();
                    TipoBD = cc.get_TipoBase(Utils.SessionBaseID);

                    if (GadgetPermitidos != "")
                    {
                        string UserName = Common.Utils.SessionUserName;
                      
                        if (Utils.Session_ModuloActivo != "RHPROX2")//Si estoy logueado y seleccione un modulo
                        {
                            if (TipoBD == "MSSQL")
                            {
                                sql = " SELECT  GU.gadusrnro, GU.gadusractivo,GU.gadusrposicion,GU.iduser, GU.gadnro  ";
                                sql += "  ,(select GT.gadactivo from Gadgets_Tipo GT where GT.gadnro=GU.gadnro) gadactivo ";
                                sql += "  ,(select GT.gadURL from Gadgets_Tipo GT where GT.gadnro=GU.gadnro) gadURL ";
                                sql += "  ,(select GT.gaddesabr from Gadgets_Tipo GT where GT.gadnro=GU.gadnro) gaddesabr ";
                                sql += "  ,(select GT.gadtitulo from Gadgets_Tipo GT where GT.gadnro=GU.gadnro) gadtitulo ";
                                sql += "  ,(select GM.gadusraltofull from Gadgets_User_Modulo GM where GM.menumsnro=" + Utils.Session_MenumsNro_Modulo + " and GM.gadusrnro=GU.gadusrnro) gadusraltofull ";
                                sql += "  ,(select GM.gadusranchofull from Gadgets_User_Modulo GM where GM.menumsnro=" + Utils.Session_MenumsNro_Modulo + " and GM.gadusrnro=GU.gadusrnro) gadusranchofull ";
                                sql += "  ,(select 1) gadtipo ";
                                sql += " FROM Gadgets_User GU  ";
                                sql += " WHERE (select GT.gadactivo from Gadgets_Tipo GT where GT.gadnro=GU.gadnro) = -1 ";
                                sql += "       AND (select GT.gadtipo from Gadgets_Tipo GT where GT.gadnro=GU.gadnro) = 1 ";//1: Es el tipo modulo
                                sql += "       AND GU.iduser='" + UserName + "' ";
                                sql += "       AND GU.gadusractivo=-1  ";
                                sql += "       AND GU.gadestado=-1    ";
                                sql += "       AND GU.gadusrnro IN ( select GM.gadusrnro from Gadgets_User_Modulo GM where GM.menumsnro=" + Utils.Session_MenumsNro_Modulo + " )  ";
                                if (GadgetPermitidos != "")
                                    sql += "  AND   GU.gadnro in (" + GadgetPermitidos + ") ";

                                //Muestra el gadget en el modulo si no tiene asignado ninguno al mismo, o si tiene asignado el gadget al modulo especifico                        
                                //sql += "       AND  ( GU.gadusrnro IN ( select GM.gadusrnro from Gadgets_User_Modulo GM where GM.menumsnro=" + Utils.Session_MenumsNro_Modulo + " )  ";
                                //sql += "              OR NOT GU.gadusrnro IN ( select GM.gadusrnro from Gadgets_User_Modulo GM where GM.gadusrnro=GU.gadusrnro)";                        
                                //sql += "            ) ";
                                sql += "  ORDER BY gadusrposicion ASC ";

                            }
                            else//ORA
                            {

                                  sql = " SELECT GU.gadusrnro, GU.gadusractivo,GU.gadusrposicion,GU.iduser, GU.gadnro  "; 
                                  sql += "   ,(select GT.gadactivo from Gadgets_Tipo GT where GT.gadnro=GU.gadnro) gadactivo "; 
                                  sql += "   ,(select GT.gadURL from Gadgets_Tipo GT where GT.gadnro=GU.gadnro) gadURL "; 
                                  sql += "   ,(select GT.gaddesabr from Gadgets_Tipo GT where GT.gadnro=GU.gadnro) gaddesabr "; 
                                  sql += "   ,(select GT.gadtitulo from Gadgets_Tipo GT where GT.gadnro=GU.gadnro) gadtitulo ";
                                  sql += "   ,(select GM.gadusraltofull from Gadgets_User_Modulo GM where GM.menumsnro=" + Utils.Session_MenumsNro_Modulo + " and GM.gadusrnro=GU.gadusrnro) gadusraltofull ";
                                  sql += "  ,(select GM.gadusranchofull from Gadgets_User_Modulo GM where GM.menumsnro=" + Utils.Session_MenumsNro_Modulo + " and GM.gadusrnro=GU.gadusrnro) gadusranchofull "; 
                                  sql += "   ,1 gadtipo "; 
                                   sql += " FROM Gadgets_User GU ";  
                                  sql += "  WHERE (select GT.gadactivo from Gadgets_Tipo GT where GT.gadnro=GU.gadnro) = -1 "; 
                                  sql += "        AND (select GT.gadtipo from Gadgets_Tipo GT where GT.gadnro=GU.gadnro) = 1  ";
                                  sql += "        AND GU.iduser='" + UserName + "'";  
                                  sql += "        AND GU.gadusractivo=-1 ";  
                                  sql += "        AND GU.gadestado=-1  ";
                                  sql += "        AND GU.gadusrnro IN ( select GM.gadusrnro from Gadgets_User_Modulo GM where GM.menumsnro=" + Utils.Session_MenumsNro_Modulo + " )  ";
                                  if (GadgetPermitidos != "")
                                          sql += "AND   GU.gadnro in (" + GadgetPermitidos + ")  ";                            
                                   sql += "  ORDER BY GU.gadusrposicion ASC "; 

                            }                            
                             
                        }   
                        else //Si estoy logueado y estoy en la portada principal
                        {                       
                            sql = " SELECT  GU.* ";
                            sql += "  ,(select GT.gadactivo from Gadgets_Tipo GT where GT.gadnro=GU.gadnro) gadactivo ";
                            sql += "  ,(select GT.gadURL from Gadgets_Tipo GT where GT.gadnro=GU.gadnro) gadURL ";
                            sql += "  ,(select GT.gaddesabr from Gadgets_Tipo GT where GT.gadnro=GU.gadnro) gaddesabr ";
                            sql += "  ,(select GT.gadtitulo from Gadgets_Tipo GT where GT.gadnro=GU.gadnro) gadtitulo ";
                            if (TipoBD == "MSSQL")
                                sql += "  ,(select 0) gadtipo ";
                            else
                                sql += "  ,0 gadtipo ";

                            sql += " FROM Gadgets_User GU  ";
                            sql += " WHERE (select GT.gadactivo from Gadgets_Tipo GT where GT.gadnro=GU.gadnro) = -1 ";
                            sql += "       AND (select GT.gadtipo from Gadgets_Tipo GT where GT.gadnro=GU.gadnro) = 0 ";//0: Es el tipo home
                            sql += "       AND GU.gadestado=-1    ";                         
                            sql += "       AND iduser='" + UserName + "' ";
                            sql += "       AND GU.gadusractivo = -1 ";

                            if (GadgetPermitidos != "")
                                sql += "  AND   GU.gadnro in (" + GadgetPermitidos + ") ";

                            sql += "  ORDER BY gadusrposicion ASC ";                      

                        }

                  }
                }
                else
                {
                    //En este caso recupero los gadgets habilitados del home sin estar en estado logueado (gadgets publicos)
                    sql = " SELECT  GU.* ";
                    sql += "  ,(select GT.gadactivo from Gadgets_Tipo GT where GT.gadnro=GU.gadnro) gadactivo ";
                    sql += "  ,(select GT.gadURL from Gadgets_Tipo GT where GT.gadnro=GU.gadnro) gadURL ";
                    sql += "  ,(select GT.gaddesabr from Gadgets_Tipo GT where GT.gadnro=GU.gadnro) gaddesabr ";
                    sql += "  ,(select GT.gadtitulo from Gadgets_Tipo GT where GT.gadnro=GU.gadnro) gadtitulo ";
                    sql += "  ,(select 0) gadtipo ";
                    sql += " FROM Gadgets_User GU  ";
                    sql += " WHERE (select GT.gadactivo from Gadgets_Tipo GT where GT.gadnro=GU.gadnro) = -1 ";
                    sql += "       AND (select GT.gadtipo from Gadgets_Tipo GT where GT.gadnro=GU.gadnro) = 0 ";//0: Es el tipo home
                    sql += "       AND GU.gadusractivo=-1  ";
                    sql += "       AND iduser IS NULL ";                
                    sql += "  ORDER BY gadusrposicion ASC ";
                }

                if (sql != "")
                {


                    cc = new Consultas();
                    DataTable dt = cc.get_DataTable(sql, Utils.SessionBaseID);
                    String IdDrag = "";
                    int pos = 0;
                    String AnchoGadget = Min_AnchoGadget;
                    int gadusraltofull = 0;
                    int gadusranchofull = 0;
                    Boolean QuitarScroll;

                    foreach (System.Data.DataRow dr in dt.Rows)
                    {
                        if (Convert.ToString(dr["gadURL"]) != "")
                        {
                            try
                            {
                                urlControl = "~/Gadgets/" + (Convert.ToString(dr["gadURL"])).Substring(0, (Convert.ToString(dr["gadURL"])).Length);
                                GadgetControl = (Control)Page.LoadControl(urlControl);
                                
                            }
                            catch {
                                GadgetControl = null;
                            }
                            
                            QuitarScroll = false;

                            if (GadgetControl != null)
                            {
                                if (!DBNull.Value.Equals(dr["gadusranchofull"]))
                                {
                                    if (Convert.ToInt32(dr["gadusranchofull"]) == -1)
                                        AnchoGadget = Max_AnchoGadget;
                                    else
                                        AnchoGadget = Min_AnchoGadget;

                                    gadusranchofull = Convert.ToInt32(dr["gadusranchofull"]);
                                }

                                if (!DBNull.Value.Equals(dr["gadusraltofull"]))
                                {
                                    gadusraltofull = Convert.ToInt32(dr["gadusraltofull"]);
                                }

                                if (IdContenedor == 1)
                                {
                                    IdDrag = Convert.ToString(dr["gadusrnro"]);

                                    if (Convert.ToString(dr["gadURL"]).Contains("HomeIndicadores/HomeIndicadores.ascx"))
                                    {
                                        IdDrag = "indicadores";
                                        idIdentidicador.ID = "idIdentidicador";
                                        idIdentidicador.Value = "gadnro_" + Convert.ToInt32(dr["gadusrnro"]);
                                    }


                                    if (Convert.ToString(dr["gadURL"]).Contains("HomeImagenCorporativa/Gadget_Corporativa.ascx"))
                                    {
                                        QuitarScroll = true;
                                    }

                                    if (Convert.ToInt32(dr["gadusraltofull"]) == -1)
                                        MiPanel.Controls.Add(new LiteralControl("<DIV class='GadgetFlotante' style='width:" + AnchoGadget + " !important'  id='gadnro_" + Convert.ToInt32(dr["gadusrnro"]) + "'  onmouseup='Soltar(this)'  onmousemove='Mover()' onmouseover='color(this)' onmouseout='saleTD(this)'>"));
                                    else
                                        MiPanel.Controls.Add(new LiteralControl("<DIV class='GadgetFlotante ContenedorGadget_Alto' style='width:" + AnchoGadget + " !important'  id='gadnro_" + Convert.ToInt32(dr["gadusrnro"]) + "'  onmouseup='Soltar(this)'  onmousemove='Mover()' onmouseover='color(this)' onmouseout='saleTD(this)'>"));

                                    MiPanel.Controls.Add(new LiteralControl(TopeModulo(QuitarScroll, ObjLenguaje.Label_Home(Convert.ToString(dr["gadtitulo"])), Convert.ToInt32(dr["gadusrnro"]), Convert.ToString(dr["gaddesabr"]), Convert.ToInt32(gadusranchofull), Convert.ToInt32(gadusraltofull), Convert.ToInt32(dr["gadtipo"]))));
                                    MiPanel.Controls.Add(GadgetControl);
                                    MiPanel.Controls.Add(new LiteralControl(PisoModulo()));
                                    MiPanel.Controls.Add(new LiteralControl("</DIV>"));

                                }
                                else
                                {
                                    IdDrag = Convert.ToString(dr["gadusrnro"]);

                                    switch (Convert.ToString(dr["gadURL"]))
                                    {
                                        case "HomeIndicadores/HomeIndicadores.ascx":
                                            {
                                                Gadgets_Del_Modulo.Controls.Add(new LiteralControl("<script>GADNRO_INDICADORES = 'gadnro_" + Convert.ToInt32(dr["gadusrnro"]) + "';</script>"));
                                                ScriptManager.RegisterStartupScript(Page, GetType(), "ControlIndicadoresModulo", "GADNRO_INDICADORES = 'gadnro_" + Convert.ToInt32(dr["gadusrnro"]) + "';", true);
                                            }
                                            break;
                                        case "HomeMRU_Modulo/MRU_Modulo.ascx":
                                            {
                                                Gadgets_Del_Modulo.Controls.Add(new LiteralControl("<script>GADNRO_MRU_MODULO = 'gadnro_" + Convert.ToInt32(dr["gadusrnro"]) + "';</script>"));
                                                ScriptManager.RegisterStartupScript(Page, GetType(), "ControlMRUModulo", "GADNRO_MRU_MODULO = 'gadnro_" + Convert.ToInt32(dr["gadusrnro"]) + "';", true);
                                            }
                                            break;
                                    }

                                    //Gadgets_Del_Modulo.Controls.Add(new LiteralControl("<DIV class='GadgetFlotante' style='width:" + AnchoGadget + " !important' name='gadnro_" + IdDrag + "' id='gadnro_" + Convert.ToInt32(dr["gadnro"]) + "'  onmouseup='Soltar(this)'  onmousemove='Mover()' onmouseover='color(this)' onmouseout='saleTD(this)'>"));
                                    Gadgets_Del_Modulo.Controls.Add(new LiteralControl("<DIV class='GadgetFlotante' style='width:" + AnchoGadget + " !important'  id='gadnro_" + Convert.ToInt32(dr["gadusrnro"]) + "'  onmouseup='Soltar(this)'  onmousemove='Mover()' onmouseover='color(this)' onmouseout='saleTD(this)'>"));

                                    Gadgets_Del_Modulo.Controls.Add(new LiteralControl(TopeModulo(QuitarScroll, ObjLenguaje.Label_Home(Convert.ToString(dr["gadtitulo"])), Convert.ToInt32(dr["gadusrnro"]), Convert.ToString(dr["gaddesabr"]), Convert.ToInt32(gadusranchofull), Convert.ToInt32(gadusraltofull), Convert.ToInt32(dr["gadtipo"]))));
                                    Gadgets_Del_Modulo.Controls.Add(GadgetControl);
                                    Gadgets_Del_Modulo.Controls.Add(new LiteralControl(PisoModulo()));
                                    Gadgets_Del_Modulo.Controls.Add(new LiteralControl("</DIV>"));
                                }

                            }
                        }


                    }
                }

                //Arma nuevamente el slider de los banners
                ScriptManager.RegisterStartupScript(this, typeof(Page), "SliderBanners_", " if (typeof ArmarGaleria == 'function')  {ArmarGaleria();};  ", true);
                ScriptManager.RegisterStartupScript(this, typeof(Page), "InicializaZindex", "if (document.getElementById('main-menu')){document.getElementById('main-menu').style.zIndex=400;} if (document.getElementById('main-menuTopLoguin')){document.getElementById('main-menuTopLoguin').style.zIndex=300;}   ", true);
                
             }            
                         
            
        }


        public string TopeInfoModulos(String desabr, String icono,String accion,String linkmanual,String linkdvd, bool puede, String DescripcionModulo) {
            string TopeInfo;            
            //TopeInfo = " <table width='623' height='44' border='0' cellspacing='0' cellpadding='0' align='center' class='TopeInfoModulos'> ";
            TopeInfo = " <table width='100%' height='44' border='0' cellspacing='0' cellpadding='0' align='center' class='ContenedorModulo' > ";

            TopeInfo += "  <tr class='ContenedorModulo_Cab'> ";
            TopeInfo += "<td width='401' align='left'  valign='middle' nowrap='nowrap'> ";
            //TopeInfo += " <object wmode='transparent'  type='image/svg+xml' data='img/Modulos/SVG/"+icono+".svg' class='IconoModulo' >";
            //TopeInfo += "   <img src='img/Modulos/PNG/"+icono+".png'>";
            //TopeInfo += " </object>";

            //TopeInfo += " <img src=' img/Modulos/SVG/" + icono + ".svg' align='absmiddle'  class='IconoInfoModulo'  >";
            TopeInfo += Utils.Armar_Icono("img/Modulos/SVG/" + icono + ".svg", "IconoInfoModulo", "", "align='absmiddle'", "");
             

            
            TopeInfo +=   ObjLenguaje.Label_Home(desabr)  ;
            TopeInfo += "</td>";           
            TopeInfo += " <td width='80' align='right'  valign='middle' nowrap='nowrap'> ";

            

            if (Utils.IsUserLogin)
            {
                if (puede)
                {
                   // TopeInfo += " <span  onclick='AbrirLink(\"../" + accion + "\",\"" + icono + "\")' style = 'cursor:pointer'> " + ObjLenguaje.Label_Home("Acceder") + "</span>   ";
                   // TopeInfo += " <img src='img/Modulos/SVG/APERTURA_MODULO.svg' align='absmiddle'  class='IconoAperturaModulo'  onclick='AbrirLink(\"../" + accion + "\")' style = 'cursor:pointer'> ";
                }
            }
            else
            {
                if (icono == "ESS")
                {
                    TopeInfo += " <span  onclick='AbrirLink(\"../" + accion + "\")' style = 'cursor:pointer'> " + ObjLenguaje.Label_Home("Acceder") + "</span> ";
                    //TopeInfo += " <img src='img/Modulos/SVG/APERTURA_MODULO.svg' align='absmiddle' class='IconoAperturaModulo'  onclick='AbrirLink(\"../" + accion + "\")' style = 'cursor:pointer'> ";
                    //TopeInfo += " <object wmode='transparent'  type='image/svg+xml' data='img/Modulos/SVG/APERTURA_MODULO.svg' class='IconoAperturaModulo'   ";
                    //TopeInfo += "    onclick='AbrirLink(\"../" + accion + "\")'  style = 'cursor:pointer'>  ";
                    //TopeInfo += "   <img src='img/Modulos/PNG/APERTURA_MODULO.png'>";
                    //TopeInfo += " </object>";
                }
            }

            TopeInfo +=  "</td>";       
            
            
            
            TopeInfo += "</tr>";

            TopeInfo += "<tr class='ContenedorModulo_Info'>";
            TopeInfo += "<td colspan='2'> ";
            TopeInfo += DescripcionModulo;
            TopeInfo += "</td>";
            TopeInfo += "</tr>";
            TopeInfo += "</table> ";

            return TopeInfo;
        }





        public string TopeModulo(Boolean QuitarScroll,string Titulo,int gnro, string detalle,int AnchoFull,int AltoFull, int gadtipo)
        {
            String Tope;
            Tope = " <table style='width:100%'  border='0' cellspacing='0' cellpadding='0' align='center' class='BordeGris'";

           // if (Utils.IsUserLogin) 
           //     Tope += "  onmouseout=\"CerrarTooltipHelp('Identificador" + gnro + "')\" ";

            Tope += "  id='drag_" + gnro + "' >";           
            Tope += "        <tr> ";
            Tope += "     <td valign='middle' align='left' class='CabeceraDrag'  ";
             
            Tope += " > ";          
            
            Tope += "    <table width='100%' border='0' cellspacing='0' cellpadding='0' align='center'  >";
            Tope += "               <tr class='PisoGris' >";
            Tope += "                 <td valign='middle' align='center'    ";

            //if (Utils.IsUserLogin)
            //    Tope += " onmousedown='Tomar(document.getElementById(\"drag_" + gnro + "\"))' onmousemove='Mover()' style='cursor:move;'  ";
            Tope += " > ";
            
            Tope += " <div class='TituloDetModulo' ";
            if (Utils.IsUserLogin)
                Tope += " onmousedown='Tomar(document.getElementById(\"drag_" + gnro + "\"))' onmousemove='Mover()' style='cursor:move;'  title='" + ObjLenguaje.Label_Home("Mover Gadget") + "' ";

            Tope +=" >" + Titulo + "</div>";
           // Tope += " </td>";
           // Tope += "  <td style='vertical-align:middle !important; text-align:right; padding-right:3px; '  nowrap>";
            
            if (Utils.IsUserLogin)
            {
                Tope += "<span style='float:right !important'  >";

                //Tope += "<span class='BotonCabeceraGadget'>" + Utils.Armar_Icono("~/../img/Modulos/SVG/EXPANDEALTO.svg", "IconoModuloGadget", ObjLenguaje.Label_Home("Ajustar Alto"), " border='0' ", "","ExpandirAltura(" + gnro + "," + AltoFull + "," + gadtipo + ")") + "</span> ";
                //Tope += "<span class='BotonCabeceraGadget'>" + Utils.Armar_Icono("~/../img/Modulos/SVG/EXPANDEANCHO.svg", "IconoModuloGadget", ObjLenguaje.Label_Home("Ajustar Ancho"), " border='0' ", "","ExpandirAncho(" + gnro + "," + AnchoFull + "," + gadtipo + ")") + " </span>  ";
                //Tope += "<span class='BotonCabeceraGadget'>" + Utils.Armar_Icono("~/../img/Modulos/SVG/UP.svg", "IconoModuloGadget", ObjLenguaje.Label_Home("Subir"), " border='0'  ", "","Subir(" + gnro + ")") + " </span> ";
                //Tope += "<span class='BotonCabeceraGadget'>" + Utils.Armar_Icono("~/../img/Modulos/SVG/DOWN.svg", "IconoModuloGadget", ObjLenguaje.Label_Home("Bajar"), " border='0'   ", "","Bajar(" + gnro + ")") + " </span> ";
                //Tope += "<span class='BotonCabeceraGadget'>" + Utils.Armar_Icono("~/../img/Modulos/SVG/APAGAR.svg", "IconoModuloGadget", ObjLenguaje.Label_Home("Desactivar"), " border='0'   ", "","Desactivar(" + gnro + ",'" + ObjLenguaje.Label_Home("Deséa desactivar el control?") + "')") + "</span>    ";

                
                //Tope += Utils.file_get_contents("~/../img/Modulos/SVG/EXPANDEALTO.svg");
                Tope += "<span class='BotonCabeceraGadget'> <a class='BtnTransparenteOcultar' onclick=\"ExpandirAltura(" + gnro + "," + AltoFull + "," + gadtipo + ")\"> </a>  " + Utils.Armar_Icono("~/../img/Modulos/SVG/EXPANDEALTO.svg", "IconoModuloGadget", ObjLenguaje.Label_Home("Ajustar Alto"), " border='0' ", "", "") + "</span>";
                Tope += "<span class='BotonCabeceraGadget'> <a class='BtnTransparenteOcultar' onclick=\"ExpandirAncho(" + gnro + "," + AnchoFull + "," + gadtipo + ")\"> </a>  " + Utils.Armar_Icono("~/../img/Modulos/SVG/EXPANDEANCHO.svg", "IconoModuloGadget", ObjLenguaje.Label_Home("Ajustar Ancho"), " border='0' ", "", "") + "</span>";
                Tope += "<span class='BotonCabeceraGadget'> <a class='BtnTransparenteOcultar' onclick=\"Subir(" + gnro + ")\"> </a>  " + Utils.Armar_Icono("~/../img/Modulos/SVG/UP.svg", " IconoModuloGadget", ObjLenguaje.Label_Home("Subir"), " border='0'  ", "", "") + "</span>";
                Tope += "<span class='BotonCabeceraGadget'> <a class='BtnTransparenteOcultar' onclick=\"Bajar(" + gnro + ")\"> </a>  " + Utils.Armar_Icono("~/../img/Modulos/SVG/DOWN.svg", " IconoModuloGadget", ObjLenguaje.Label_Home("Bajar"), " border='0'   ", "", "") + "</span>";
                Tope += "<span class='BotonCabeceraGadget'> <a class='BtnTransparenteOcultar' onclick=\"Desactivar(" + gnro + ",'" + ObjLenguaje.Label_Home("Deséa desactivar el control?") + "')\"> </a>  " + Utils.Armar_Icono("~/../img/Modulos/SVG/APAGAR.svg", " IconoModuloGadget", ObjLenguaje.Label_Home("Desactivar"), " border='0'   ", "", "") + "</span>";
                

                //Tope += "<img src='~/../img/Modulos/SVG/MORE.svg' border='0' class='IconoModuloGadget'  onclick=\"AbrirTooltipHelp('Identificador" + gnro + "')\"  title='" + ObjLenguaje.Label_Home("Detalle") + "' >     " + Configurador(gnro, detalle)  ;                

                Tope += "</span>";
                Tope += "</td>";
                 
            }
              

            Tope += "                </tr>";
            Tope += "              </table></td>";
            Tope += "           </tr>";
            Tope += "           <tr>";
            Tope += "             <td  valign='top' align='center' style='background-color:#FFFFFF;width:100%' >";
            
            
            if (AltoFull==0)
                Tope += "  <div  class='ContenedorGadget' style='width:100% ";
            else
                Tope += "  <div  class='ContenedorGadgetAltoFull' style='width:100%  ";


            if (QuitarScroll)
                Tope += " ;overflow:hidden !important;'  ";

            Tope += "  '> ";

            return Tope;
        }

        public String PisoModulo()
        {
            String Piso = "</div>";
            Piso += "    </td>";
            Piso += " </tr>";
            Piso += "<tr>";
            Piso += "   <td colspan='2' class='TopeGris'></td>";
            Piso += "</tr>";
            Piso += "</table>";
            return Piso;
        }

       // public string Configurador(int gadnro)
        public string Configurador(int gadnro,string detalle)
        {
            string conf = " <table width='130' border='0' cellspacing='0' cellpadding='0' class='tooltiphelp'  id='Identificador" + gadnro + "'   ";
            conf += "          onmouseover=\"AbrirTooltipHelp('Identificador" + gadnro + "')\"   onmouseout=\"CerrarTooltipHelp('Identificador" + gadnro + "')\">";
 
            conf += "<tr> ";
            conf += " <td align='right' valign='bottom'  style='background-color:transparent'><img src='~/../img/Help2/top_izq.png'   /></td> ";
            conf += " <td width='100%' style='background:url(~/../img/Help2/top.png) repeat-x bottom' valign='bottom' align='center' >";
            conf += "<img src='~/../img/Help/punta.png' style='margin-bottom:8px;//margin-bottom:5px'   /> ";
            conf += "    </td> ";
            conf += " <td align='left' valign='bottom' style='background-color:transparent'><img src='~/../img/Help2/top_der.png'   /></td> ";
            conf += "</tr>  ";
            conf += "<tr>  ";
            conf += "<td style='background:url(~/../img/Help2/izq.png) repeat-y top' >&nbsp;</td> ";
            conf += "<td style='white-space: normal;background:url(~/../img/Help2/centro.png) repeat top'  class='contenidoTooltip' width='100%' >";
            conf += "<div style='width:130px; text-align:left; overflow-x: visible; '  >   " + ObjLenguaje.Label_Home(detalle) + " </div>";
            conf += "</td> ";
            conf += "    <td style='background:url(~/../img/Help2/der.png) repeat-y top' >&nbsp;</td> ";
            conf += "  </tr>  ";
            conf += "  <tr> ";
            conf += "    <td align='right' valign='top' style='background-color:transparent'><img src='~/../img/Help2/bottom_izq.png'   /></td> ";
            conf += "    <td style='background:url(~/../img/Help2/bottom.png) repeat-x top' > </td> ";
            conf += "    <td align='left' valign='top' style='background-color:transparent'><img src='~/../img/Help2/bottom_der.png'   /></td> ";
            conf += "</tr>  ";
            conf += "</table> ";

            return conf;
        }
       
        
        public string Configurador2(int gadnro)
        {
            string conf = " <table width='100' border='0' cellspacing='0' cellpadding='0' class='tooltiphelp'  id='Identificador" + gadnro + "'  ";
            conf += "          onmouseover=\"AbrirTooltipHelp('Identificador" + gadnro + "')\"   onmouseout=\"CerrarTooltipHelp('Identificador" + gadnro + "')\">";
            
            conf += "<tr> ";   
            conf += " <td width='100%'  valign='bottom' align='center' >";
            conf += " <img src='~/../img/Help/punta2.png' style='margin-bottom:0px'   /> ";
            conf += "</td> "; 
            conf += "</tr>  ";


            conf += "<tr> ";
            conf += "<td  nowrap class='tool contenidoTooltip' >";
               conf += "<div  ><a href='' onclick=\"Desactivar(" + gadnro + ",'" + ObjLenguaje.Label_Home("Deséa desactivar el control?") + "')\"> &raquo; " + ObjLenguaje.Label_Home("Desactiva") + "</a></div>";
               conf += "<div  ><a href=''> &raquo; " + ObjLenguaje.Label_Home("Modificar") + "</a></div>";
               conf += "<div  ><a href=''> &raquo; " + ObjLenguaje.Label_Home("Eliminar") + "</a></div>";
             conf += "</td> ";
            conf += "</tr>  ";
            conf += "</table> ";
                    
            return conf;
         }
 

        //Este metodo carga el detalle de un determinado acceso en el contenedor principal.
        public void Actualizar_Accesos_XML(string nroAcceso)
        {
            Control AccesoControl;
            String urlAcceso;
            try
            {  
                //-------------------
                string BaseId = Common.Utils.SessionBaseID;
                string UserName = Common.Utils.SessionUserName;

                /* ****** Preparo el filtro de accesos que deseo visualizar desde la base****** */
 

                string sql = "SELECT * FROM Home_Accesos ";
                sql += " WHERE Activo = -1 ";                
                sql += "  AND nroAcceso=" + nroAcceso;
                sql += " ORDER BY Nombre ASC ";

                Consultas cc = new Consultas();
                

                DataSet ds = cc.get_DataSet(sql, BaseId);  
 
                if (ds.Tables.Count > 0)
                {
                    if (ds.Tables[0].Rows.Count > 0)
                    {                       
                        if (!String.IsNullOrEmpty(Convert.ToString(ds.Tables[0].Rows[0]["ArchivoDescripcion"])))
                        {
                            Cuerpo.Visible = false;
                            //Vacio el panel de modulos          
                            MiPanel.Controls.Clear();
                            MiPanel.Visible = true;
                            String ArchivoDescripcion = Convert.ToString(ds.Tables[0].Rows[0]["ArchivoDescripcion"]);
                            String Nombre = (String)ds.Tables[0].Rows[0]["Nombre"];
                            String URL = (String)ds.Tables[0].Rows[0]["URL"];
                            bool isLogin = (bool)ds.Tables[0].Rows[0]["isLogin"];
                            urlAcceso = "~/Accesos/" + (ArchivoDescripcion).Substring(0, (ArchivoDescripcion).Length);
                            AccesoControl = (Control)Page.LoadControl(urlAcceso);
                            if (AccesoControl != null)
                            {
                                System.Web.HttpContext.Current.Session["ActualizaAcceso"] = nroAcceso;
                                
                                MiPanel.Controls.Add(new LiteralControl(CabeceraAccesos(Nombre, URL, isLogin)));
                                MiPanel.Controls.Add(new LiteralControl("<TR class='ContenedorModulo_Info'><TD colspan='3'>"));
                                MiPanel.Controls.Add(AccesoControl);
                                MiPanel.Controls.Add(new LiteralControl("</TD></TR>"));
                                MiPanel.Controls.Add(new LiteralControl("</TABLE>"));
                                
                                MiPanel.DataBind();
                            }
                        }
                    }
                }

              
            }
            catch (Exception ex) {
                throw ex;                
            }

            //Mantiene la barra de los modulos Expandida o Colapsada
            ScriptManager.RegisterStartupScript(Page, GetType(), "ControlExpandidoAcc", "setTimeout('if (ControlExpand==1){ControlExpand=0;} else {ControlExpand=1;};DesplazarBarraMenu()', 30);", true);
            ScriptManager.RegisterStartupScript(this, typeof(Page), "InicializaMenu", "$(function() {  $('#main-menu').smartmenus({  subMenusSubOffsetX: 0, subMenusSubOffsetY: -1 , hideOnClick: false  }); });  ", true);
            ScriptManager.RegisterStartupScript(this, typeof(Page), "InicializaMenuTop", "$(function() {  $('#main-menuTopLoguin').smartmenus({  subMenusSubOffsetX: 0, subMenusSubOffsetY: -1, mainMenuSubOffsetX:0,mainMenuSubOffsetY:0,subMenusMinWidth:'60px', subMenusMaxWidth:'1060px', hideOnClick: true   }); });  ", true);
            //Arma nuevamente el slider de los banners y el menutop
            ScriptManager.RegisterStartupScript(this, typeof(Page), "SliderBanners_Acc", " if (typeof ArmarGaleria == 'function')  {ArmarGaleria();};  ", true);
            ScriptManager.RegisterStartupScript(this, typeof(Page), "InicializaZindex_acc", "if (document.getElementById('main-menu')){document.getElementById('main-menu').style.zIndex=400;} if (document.getElementById('main-menuTopLoguin')){document.getElementById('main-menuTopLoguin').style.zIndex=300;}   ", true);



        }


        public string CabeceraAccesos(String desabr, String accion, bool isLogin)
        {
            string TopeInfo;
            TopeInfo = " <table  border='0' cellspacing='0' cellpadding='0' align='center' class='ContenedorModulo' style='margin-top:12px'  > ";
            TopeInfo += "  <tr class='ContenedorModulo_Cab'> ";
            TopeInfo += "<td width='401' align='left'  valign='middle' nowrap='nowrap'>";
            //TopeInfo += "<b><span style='margin-left:5px'> <img src='img/modulos/SVG/LINK.svg' align='absmiddle' class='IconoModuloGadget' >  " + ObjLenguaje.Label_Home(desabr) + "</span></b>";
            TopeInfo += "<b><span style='margin-left:5px'> " + Utils.Armar_Icono("img/modulos/SVG/LINK.svg", "IconoModuloGadget", "", "align='absmiddle'", "") + ObjLenguaje.Label_Home(desabr) + "</span></b>";
            
            TopeInfo += "</td>";
            TopeInfo += " <td width='80' align='left'  valign='middle' nowrap='nowrap'>&nbsp; ";
            TopeInfo += "</td>";
            TopeInfo += " <td width='142' align='right'  valign='middle' nowrap='nowrap'> ";
            if (isLogin == true)
            {
                if (Common.Utils.IsUserLogin)
                {
                    TopeInfo += " <span  onclick='AbrirModulo(\"" + accion + "\",\"ESS\")' style = 'cursor:pointer; margin-right:10px;//margin-right:8px'> <b>" + ObjLenguaje.Label_Home("Acceder") + "  </b> ";
                    //TopeInfo += " <img src='img/modulos/SVG/APERTURA_MODULO.svg' align='absmiddle' class='IconoModuloGadget'  style = 'cursor:pointer; '> </span>";
                    TopeInfo += Utils.Armar_Icono("img/modulos/SVG/APERTURA_MODULO.svg", "IconoModuloAcceso", "", "align='absmiddle' style = 'cursor:pointer; '", "") + " </span>";
                     
                }
            }
            else {
                TopeInfo += " <span  onclick='AbrirModulo(\"" + accion + "\",\"ESS\")' style = 'cursor:pointer; margin-right:10px;//margin-right:8px'> <b>" + ObjLenguaje.Label_Home("Acceder") + "  </b> ";
                //TopeInfo += " <img src='img/modulos/SVG/APERTURA_MODULO.svg' align='absmiddle' class='IconoModuloGadget'   style = 'cursor:pointer; '> </span>";
                TopeInfo += Utils.Armar_Icono("img/modulos/SVG/APERTURA_MODULO.svg", "IconoModuloAcceso", "", "align='absmiddle' style = 'cursor:pointer; '", "") + " </span>";
                
            }
            
            TopeInfo += "</td>";
  
            TopeInfo += "</tr>";
            //TopeInfo += "</table>";

            return TopeInfo;
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