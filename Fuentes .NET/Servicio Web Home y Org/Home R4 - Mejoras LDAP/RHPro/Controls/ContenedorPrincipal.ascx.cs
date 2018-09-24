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


namespace RHPro.Controls
{
    public partial class ContenedorPrincipal : System.Web.UI.UserControl
    {

        public RHPro.Lenguaje ObjLenguaje;
        
        protected void Page_Load(object sender, EventArgs e)
        {
            ObjLenguaje = new Lenguaje();
            
        }

        public void ActualizaGadgets(object sender, EventArgs e)
        {
            Update_Gadget();
        }


        protected void Page_PreRender(object sender, EventArgs e)
        {
            
            if (  ((String)System.Web.HttpContext.Current.Session["ActualizaAcceso"] != "-1") 
                 && ((String)System.Web.HttpContext.Current.Session["ActualizaAcceso"]!= "")
                 && ((String)System.Web.HttpContext.Current.Session["ActualizaAcceso"] != null) 
                 )
            {
                //Actualizar_Accesos((String)System.Web.HttpContext.Current.Session["ActualizaAcceso"]); 
                Actualizar_Accesos_XML((String)System.Web.HttpContext.Current.Session["ActualizaAcceso"]); 
            } else {
                    if (  ((String)System.Web.HttpContext.Current.Session["ActualizaModulo"] == "-1") 
                         || ((String)System.Web.HttpContext.Current.Session["ActualizaModulo"] == "")
                         || ((String)System.Web.HttpContext.Current.Session["ActualizaModulo"] == null) 
                         )
                         Update_Gadget();
                     else
                         ActualizaContenido((String)System.Web.HttpContext.Current.Session["ActualizaModulo"]);
                   }
        }

        public void ActualizaContenido(String modulo) {


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

            
            Misql = "SELECT menudetalle,menudesabr,action,linkmanual,linkdvd,menuname FROM menumstr WHERE (menudetalle IS NOT NULL) AND  menuname = '" + modulo + "'";
            Consultas cc = new Consultas();
            DataTable dt = cc.get_DataTable(Misql, Utils.SessionBaseID);
            foreach (System.Data.DataRow dr in dt.Rows) {

                if (dr["menudetalle"] != null)
                 {
                    if (!dr["menudesabr"].Equals(System.DBNull.Value))
                        desabr = (String)dr["menudesabr"];
                    if (!dr["action"].Equals(System.DBNull.Value))
                        accion = (String)dr["action"];
                    if (!dr["linkmanual"].Equals(System.DBNull.Value))
                        linkmanual = (String)dr["linkmanual"];
                    if (!dr["linkdvd"].Equals(System.DBNull.Value))
                        linkdvd = (String)dr["linkdvd"];
                    //Antes de imprimir el encabezado verifica si puede acceder al modulo. Si no puede directamente no muestra el acceso
                    puede =  ModuleServiceProxy.Puede_Acceder(Utils.SessionUserName, Utils.SessionBaseID, Thread.CurrentThread.CurrentCulture.Name, dr["menuname"].ToString());
                    
                    Cuerpo.InnerHtml = TopeInfoModulos(desabr, modulo, accion, linkmanual, linkdvd,puede);
                    Cuerpo.InnerHtml += "<DIV class='InfoModulos'>";
                    menudetalle = ObjLenguaje.Traducir_Modulo((String)dr["menudetalle"], (String)dr["menuname"]);
                    Cuerpo.InnerHtml += menudetalle.Replace(".", ".<BR><BR>");
                    Cuerpo.InnerHtml += "</DIV>";
                 }
             }
            
  

        }
 
 

        public void Update_Gadget() {
            if (bool.Parse(ConfigurationManager.AppSettings["VisualizarComplementos"]))
            {
                System.Web.HttpContext.Current.Session["ActualizaModulo"] = "-1";
                System.Web.HttpContext.Current.Session["ActualizaAcceso"] = "-1";
                Control GadgetControl;
                String urlControl;
                String Ancho;
                Cuerpo.Visible = false;
                //Vacio el panel de modulos
                MiPanel.Controls.Clear();
                MiPanel.Visible = true;

                //gadnro,gadposicion,gadURL,gadtitulo,gadfull,gaddesabr 
                string BaseId = Common.Utils.SessionBaseID;
                string sql;

                if (Utils.IsUserLogin)
                {
                    string UserName = Common.Utils.SessionUserName;
                    //sql = "SELECT ROW_NUMBER() OVER(ORDER BY gadtitulo ASC)  'pos' ,* FROM Gadgets WHERE gadactivo=-1 AND gaduser='" + UserName + "' ORDER BY gadposicion ASC";
                    sql = "SELECT * FROM Gadgets WHERE gadactivo=-1 AND gaduser='" + UserName + "' ORDER BY gadposicion ASC";
                }
                else
                    //sql = "SELECT ROW_NUMBER() OVER(ORDER BY gadtitulo ASC)  'pos' ,* FROM Gadgets WHERE gadactivo=-1 AND gaduser is NULL ORDER BY gadposicion ASC";
                    sql = "SELECT  * FROM Gadgets WHERE gadactivo=-1 AND gaduser is NULL ORDER BY gadposicion ASC";

                Consultas cc = new Consultas();
                DataTable dt = cc.get_DataTable(sql, Utils.SessionBaseID);

                MiPanel.Controls.Add(new LiteralControl("<TABLE width='97%' id='contenedor' cellpadding='0' cellspacing='0' border='0' align='center' >"));
                int pos = 0;
                foreach (System.Data.DataRow dr in dt.Rows)
                { 
                    //-----------------------------------------
                    urlControl = "~/Gadgets/" + ((String)dr["gadURL"]).Substring(0, ((String)dr["gadURL"]).Length);
                    GadgetControl = (Control)Page.LoadControl(urlControl);
                    if (GadgetControl != null)
                    {
                        //En el caso que el gadget sea full, el ancho va a ser del 97%, sino del 48%                       
                        if ( Convert.ToInt32( dr["gadfull"] ) == -1)
                        {
                            Ancho = "100%";
                            if (pos % 2 == 0)
                                MiPanel.Controls.Add(new LiteralControl("<TR><TD colspan='2' width='100%' id='gadnro_" + Convert.ToInt32(dr["gadnro"]) + "' align='center' valign='middle'  onmouseup='Soltar(this)'  onmousemove='Mover()' onmouseover='color(this)' onmouseout='saleTD(this)'>"));
                            else
                            {
                                MiPanel.Controls.Add(new LiteralControl(" <TD width='50%'   ></TD></TR> <TR><TD colspan='2' width='100%' id='gadnro_" + Convert.ToInt32(dr["gadnro"]) + "' align='center' valign='middle'>"));//venia de uno chico                                
                                pos = pos + 1;
                            }

                            MiPanel.Controls.Add(new LiteralControl(TopeModulo(ObjLenguaje.Label_Home(Convert.ToString(dr["gadtitulo"])), Ancho, Convert.ToInt32(dr["gadnro"]), Convert.ToString(dr["gaddesabr"]))));
                            MiPanel.Controls.Add(GadgetControl);
                            MiPanel.Controls.Add(new LiteralControl(PisoModulo()));

                            MiPanel.Controls.Add(new LiteralControl("</TD></TR>"));

                            pos = pos + 2;
                        }
                        else
                        {
                            Ancho = "99%";
                            if (pos % 2 == 0)
                                MiPanel.Controls.Add(new LiteralControl("<TR><TD width='50%'  id='gadnro_" + Convert.ToInt32(dr["gadnro"]) + "' align='left' valign='middle'>"));
                            else
                                MiPanel.Controls.Add(new LiteralControl("<TD width='50%'  id='gadnro_" + Convert.ToInt32(dr["gadnro"]) + "' align='right' valign='middle'>"));

                            MiPanel.Controls.Add(new LiteralControl(TopeModulo(ObjLenguaje.Label_Home(Convert.ToString(dr["gadtitulo"])), Ancho, Convert.ToInt32(dr["gadnro"]), Convert.ToString(dr["gaddesabr"]))));
                            MiPanel.Controls.Add(GadgetControl);
                            MiPanel.Controls.Add(new LiteralControl(PisoModulo()));

                            if (pos % 2 == 0)
                                MiPanel.Controls.Add(new LiteralControl("</TD>"));
                            else
                                MiPanel.Controls.Add(new LiteralControl("</TD></TR>"));

                            pos = pos + 1;
                        }

                    }
                    //-----------------------------------------
                     
                }
                if (pos % 2 != 0)
                    MiPanel.Controls.Add(new LiteralControl("<TD width='50%'></TD></TR>"));

                MiPanel.Controls.Add(new LiteralControl("</TABLE>"));
            }

            
        }


        public string TopeInfoModulos(String desabr, String icono,String accion,String linkmanual,String linkdvd, bool puede) {
            string TopeInfo;            
            TopeInfo = " <table width='623' height='44' border='0' cellspacing='0' cellpadding='0' align='center' class='TopeInfoModulos'> ";
            TopeInfo += "  <tr> ";
            TopeInfo += "<td width='401' align='left'  valign='middle' nowrap='nowrap'><img src='img/Modulos/originales/"+icono+".png' align='absmiddle' style='margin-left:5px;'>";
            TopeInfo += "<b> " + ObjLenguaje.Label_Home(desabr) + "</b>";
            TopeInfo += "</td>";
            TopeInfo += " <td width='142' align='left'  valign='middle' nowrap='nowrap'>&nbsp; ";

            if (Utils.IsUserLogin)
            {
                if (puede)
                {
                    TopeInfo += "<b> <span  onclick='AbrirLink(\"../" + accion + "\",\"" + icono + "\")' style = 'cursor:pointer'> " + ObjLenguaje.Label_Home("Acceder") + "</span> </b> ";
                    TopeInfo += " <img src='img/plusG.png' align='absmiddle' onclick='AbrirLink(\"../" + accion + "\")' style = 'cursor:pointer'> ";
                }
            }
            else {
                if (icono == "ESS") {
                    TopeInfo += "<b> <span  onclick='AbrirLink(\"../" + accion + "\")' style = 'cursor:pointer'> " + ObjLenguaje.Label_Home("Acceder") + "</span> </b> ";
                    TopeInfo += " <img src='img/plusG.png' align='absmiddle' onclick='AbrirLink(\"../" + accion + "\")' style = 'cursor:pointer'> ";
                }
            }
            TopeInfo += "</td>";
            TopeInfo += " <td width='80' align='left'  valign='middle' nowrap='nowrap'> ";

            if (linkdvd != "")
            {
                TopeInfo += " <img src='img/Modulos/DVD.png' align='absmiddle' ";
                TopeInfo += " onclick='AbrirLink(\"Controls/PopVideo/popVideo/index.html?path=./../../../../" + linkdvd + "&title=" + ObjLenguaje.Label_Home(desabr) + "\")' ";
                TopeInfo += " style = 'cursor:pointer'> ";
            }

            if (linkmanual != "")
            {
                TopeInfo += " <img src='img/Modulos/pdf.png' align='absmiddle' onclick='AbrirLink(\"../" + linkmanual + "\")' style = 'cursor:pointer'>  ";
            }
            TopeInfo +=  "</td>";       
            
            
            
            TopeInfo += "</tr>";
            TopeInfo += "</table>";

            return TopeInfo;
        }





        public string TopeModulo(string Titulo, string width, int gnro, string detalle)
        {
            String Tope;
            Tope = " <table style='width:" + width + "'  border='0' cellspacing='0' cellpadding='0' align='center' class='BordeGris'";

           // if (Utils.IsUserLogin) 
           //     Tope += "  onmouseout=\"CerrarTooltipHelp('Identificador" + gnro + "')\" ";

            Tope += "  id='drag_" + gnro + "' >";           
            Tope += "        <tr> ";
            Tope += "     <td valign='middle' align='left' class='CabeceraDrag'  ";
             
            Tope += " > ";          
            
            Tope += "    <table width='100%' border='0' cellspacing='0' cellpadding='0' align='center'  >";
            Tope += "               <tr>";
            Tope += "                 <td  class='PisoGris' valign='middle' align='center'    ";

            if (Utils.IsUserLogin)
                Tope += " onmousedown='Tomar(document.getElementById(\"drag_" + gnro + "\"))' onmousemove='Mover()' style='cursor:move;'  ";
            Tope += " > ";

            Tope += " <span style='margin-left:10px;'>" + Titulo + "</span>";
            Tope += " </td>";
            Tope += "  <td valign='middle' align='right' class='PisoGris' nowrap>";

            if (Utils.IsUserLogin)
            {
               // Tope += "    <img src='~/../img/Gizq.png'  onmouseover=\"this.src='~/../img/Gizq-hover.png'\" onmouseout=\"this.src='~/../img/Gizq.png'\" style='margin-right:5px;cursor:pointer' align='absmiddle' onclick='Subir(" + gnro + ")' title='" + ObjLenguaje.Label_Home("Izquierda") + "' /> ";
               // Tope += "    <img src='~/../img/Gder.png' onmouseover=\"this.src='~/../img/Gder-hover.png'\" onmouseout=\"this.src='~/../img/Gder.png'\"  style='margin-right:10px;cursor:pointer' align='absmiddle' onclick='Bajar(" + gnro + ")' title='" + ObjLenguaje.Label_Home("Derecha") + "'/> ";
                Tope += " <img src='~/../img/up.png' onmouseover=\"this.src='~/../img/up_hover.png'\" onmouseout=\"this.src='~/../img/up.png'\"  style='margin-right:6px;cursor:pointer' align='absmiddle' onclick='Subir(" + gnro + ")' title='" + ObjLenguaje.Label_Home("Subir") + "'/> ";
                Tope += " <img src='~/../img/down.png' onmouseover=\"this.src='~/../img/down_hover.png'\" onmouseout=\"this.src='~/../img/down.png'\" style='margin-right:8px;cursor:pointer' align='absmiddle' onclick='Bajar(" + gnro + ")' title='" + ObjLenguaje.Label_Home("Bajar") + "'/> ";
               // Tope += " <img src='~/../img/desactivar.png' style='margin-right:9px;cursor:pointer' onmouseover=\"this.src='~/../img/desactivar-hover.png'\" onmouseout=\"this.src='~/../img/desactivar.png'\" align='absmiddle' onclick=\"AbrirTooltipHelp('Identificador" + gnro + "')\"  title='" + ObjLenguaje.Label_Home("Desactivar") + "' /> " + Configurador(gnro) + " </td>";
                Tope += " <img src='~/../img/desactivar.png' style='margin-right:9px;cursor:pointer' onmouseover=\"this.src='~/../img/desactivar-hover.png'\" onmouseout=\"this.src='~/../img/desactivar.png'\" align='absmiddle' onclick=\"Desactivar(" + gnro + ",'" + ObjLenguaje.Label_Home("Deséa desactivar el control?") + "')\" title='" + ObjLenguaje.Label_Home("Desactivar") + "' /> "  ;
                Tope += " <img src='~/../img/detalle.png' style='margin-right:9px;cursor:pointer' onmouseover=\"this.src='~/../img/detalle-hover.png'\" onmouseout=\"this.src='~/../img/detalle.png'\" align='absmiddle' onclick=\"AbrirTooltipHelp('Identificador" + gnro + "')\"  title='" + ObjLenguaje.Label_Home("Detalle") + "' />  " + Configurador(gnro, detalle)+" </td>";
            }
              

            Tope += "                </tr>";
            Tope += "              </table></td>";
            Tope += "           </tr>";
            Tope += "           <tr>";
            Tope += "             <td  valign='top' align='center' style='background-color:#FFFFFF;width:100%' >";
            Tope += "  <div  class='ContenedorGadget' style='width:100%'> ";
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
           // conf += "<div  ><a href='' onclick=\"Desactivar(" + gadnro + ",'" + ObjLenguaje.Label_Home("Deséa desactivar el control?") + "')\"> &raquo; " + ObjLenguaje.Label_Home("Desactiva") + "</a></div>";
            //conf += "<div  ><a href=''> &raquo; " + ObjLenguaje.Label_Home("Modificar") + "</a></div>";
            //conf += "<div  ><a href=''> &raquo; " + ObjLenguaje.Label_Home("Eliminar") + "</a></div>";
            //conf += "<div  ><a href=''> &raquo; " + ObjLenguaje.Label_Home("Detalle") + "</a></div>";
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
        public void Actualizar_Accesos(string nroAcceso)
        {           
            Control AccesoControl;
            String urlAcceso;       
            Cuerpo.Visible = false;
            //Vacio el panel de modulos          
            MiPanel.Controls.Clear();            
            MiPanel.Visible = true;
          
            string BaseId = Common.Utils.SessionBaseID;
            string sql;

            string UserName = Common.Utils.SessionUserName;
            sql = "SELECT * FROM Home_Accesos WHERE Activo = 1 AND nroAcceso = " + nroAcceso;
            
            Consultas cc = new Consultas();
            DataTable dt = cc.get_DataTable(sql, Utils.SessionBaseID);
           
            foreach (System.Data.DataRow dr in dt.Rows)
            {
                urlAcceso = "~/Accesos/" + ((String)dr["ArchivoDescripcion"]).Substring(0, ((String)dr["ArchivoDescripcion"]).Length);
              //  urlAcceso = "~/Accesos/" + ((String)dr["ArchivoDescripcion"]);
                AccesoControl = (Control)Page.LoadControl(urlAcceso);
                if (AccesoControl != null)
                {
                    System.Web.HttpContext.Current.Session["ActualizaAcceso"] = nroAcceso;
                    MiPanel.Controls.Add(new LiteralControl(CabeceraAccesos((String)dr["Nombre"], (String)dr["URL"], (bool)dr["isLogin"])));    
                    MiPanel.Controls.Add(AccesoControl);
                }               
            }
        }

        //Este metodo carga el detalle de un determinado acceso en el contenedor principal.
        public void Actualizar_Accesos_XML(string nroAcceso)
        {
            Control AccesoControl;
            String urlAcceso;
            try
            {

                Cuerpo.Visible = false;
                //Vacio el panel de modulos          
                MiPanel.Controls.Clear();
                MiPanel.Visible = true;

                //Busco el nombre del archivo de configuracion de Accesos
                String URL_XML = (String)ConfigurationManager.AppSettings["AccesosHomeXML"];
                URL_XML = "../" + URL_XML;
                DataSet ds = new DataSet();
                ds.ReadXml(MapPath(URL_XML));

                int pos = int.Parse(nroAcceso) - 1;

                String ArchivoDescripcion = (String)ds.Tables["Acceso"].Rows[pos]["ArchivoDescripcion"];
                String Nombre = (String)ds.Tables["Acceso"].Rows[pos]["Nombre"];
                String URL = (String)ds.Tables["Acceso"].Rows[pos]["URL"];
                bool isLogin = bool.Parse((String)ds.Tables["Acceso"].Rows[pos]["isLogin"]);
                urlAcceso = "~/Accesos/" + (ArchivoDescripcion).Substring(0, (ArchivoDescripcion).Length);

                AccesoControl = (Control)Page.LoadControl(urlAcceso);
                if (AccesoControl != null)
                {
                    System.Web.HttpContext.Current.Session["ActualizaAcceso"] = nroAcceso;
                    MiPanel.Controls.Add(new LiteralControl(CabeceraAccesos(Nombre, URL, isLogin)));
                    MiPanel.Controls.Add(AccesoControl);
                }
            }
            catch (Exception ex) {
               
               // Response.Write(Utils.MSGE_ERROR(ex)); 
             //   Response.Write( "<span   onclick=\"this.style.visibility = 'hidden'\" style='float:left;cursor:pointer; border:font-family:Arial; font-size:9pt; color:#333;border:4px #333333 solid; position:relative; left:30px; top:30px; padding:6px; background-color:#FC9'><img src='img/error.png' align='absmiddle'> ERROR: " + ex.Message + "</span>");
                
            }

        }


        public string CabeceraAccesos(String desabr, String accion, bool isLogin)
        {
            string TopeInfo;
            TopeInfo = " <table width='623' height='44' border='0' cellspacing='0' cellpadding='0' align='center' class='TopeInfoModulos'> ";
            TopeInfo += "  <tr> ";
            TopeInfo += "<td width='401' align='left'  valign='middle' nowrap='nowrap'>";
            TopeInfo += "<b><span style='margin-left:5px'> <img src='images/LinkAccesos.png' align='absmiddle'>  " + ObjLenguaje.Label_Home(desabr) + "</span></b>";
            TopeInfo += "</td>";
            TopeInfo += " <td width='80' align='left'  valign='middle' nowrap='nowrap'>&nbsp; ";
            TopeInfo += "</td>";
            TopeInfo += " <td width='142' align='right'  valign='middle' nowrap='nowrap'> ";
            if (isLogin == true)
            {
                if (Common.Utils.IsUserLogin)
                {
                    TopeInfo += " <span  onclick='AbrirModulo(\"" + accion + "\",\"ESS\")' style = 'cursor:pointer; margin-right:10px;//margin-right:8px'> <b>" + ObjLenguaje.Label_Home("Acceder") + "  </b> ";
                    TopeInfo += " <img src='img/plusG.png' align='absmiddle'  style = 'cursor:pointer; '> </span>";
                }
            }
            else {
                TopeInfo += " <span  onclick='AbrirModulo(\"" + accion + "\",\"ESS\")' style = 'cursor:pointer; margin-right:10px;//margin-right:8px'> <b>" + ObjLenguaje.Label_Home("Acceder") + "  </b> ";
                TopeInfo += " <img src='img/plusG.png' align='absmiddle'  style = 'cursor:pointer; '> </span>";
            }
            
            TopeInfo += "</td>";
  
            TopeInfo += "</tr>";
            TopeInfo += "</table>";

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