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

        public string Home_ESS;

        protected void Page_Load(object sender, EventArgs e)
        { 
            ObjLenguaje = new RHPro.Lenguaje();
            Home_ESS = ((String)ConfigurationManager.AppSettings["Home_ESS"]);
            Repeater1.DataSourceID = "SqlDataSource1";
            DataSet ds = new DataSet();
            ds.ReadXml(MapPath("../Accesos_Home.xml"));
            rpMyRepeater.DataSource = ds;
            rpMyRepeater.DataBind();
            
            //Repeater1.DataSourceID = "SqlDataSource1";

 
           // if (Utils.IsUserLogin)
           //    JPB.Text = "LOGEADO..";          
        }
        protected void Page_PreRender(object sender, EventArgs e) {
           
             
        }


        public bool AccesoActivo(string condicion) {
            if (condicion == "true")
                return true;
            else
                return false;
        }

        public string Construir_Acceso(string activo, string nombre, string URL, string isLogin) {
            string TR = "";
          
            if (activo=="true")
            {
                TR = "<tr id='Link" + posmenu + "' onclick=\"Seleccionar('Link" + posmenu + "','adp')\"   onmouseover='Sobre(this)' onmouseout='Sale(this)'  > ";
                TR += "<td nowrap='nowrap'><img src='img/link.png' border='0' align='absmiddle'  style='margin-left: 4px;'/></td>";
                TR += "<td><span style='margin-left:3px;'>";
                TR += posmenu + ObjLenguaje.Label_Home(nombre);
                TR += "</span></td>";
                TR += "<td align='right'>";

                if (isLogin=="true")
                {
                    if (Utils.IsUserLogin)
                    {
                        TR += "<img src='img/plus.png' border='0' align='absmiddle' style='margin-right:4px;' onmouseover=\"this.src = 'img/plus_hover.png'\"";
                        TR += "onmouseout=\"this.src = 'img/plus.png'\" onclick=\"AbrirModulo('" + URL + "','ESS')\"> ";
                    }
                }
                else {
                    TR += "<img src='img/plus.png' border='0' align='absmiddle' style='margin-right:4px;' onmouseover=\"this.src = 'img/plus_hover.png'\"";
                    TR += "onmouseout=\"this.src = 'img/plus.png'\" onclick=\"AbrirModulo('" + URL + "','ESS')\"> ";
                }
                TR += " </td>";
                TR += "</tr>   ";
            }
            IncrementaPosmenu(); 
            return TR;
        }
        

        public int IncrementaPosmenu() {
            posmenu++;
            return posmenu;
        }

        public int IncrementaPosmenu(int valor)
        {
            return valor;
        }
        public void Traducir(string frase)  
        {
            // ((LinkButton)sender).Text = "cambio";
            //LB.Text = texto + "CCCCC";
       
             Response.Write(ObjLenguaje.Label_Home(frase));
        }

        public void Traduce(object sender, EventArgs e)
        {
          
            // ((LinkButton)sender).Text = "cambio";
            //LB.Text = texto + "CCCCC";
           //((LinkButton)sender).Text = ObjLenguaje.Label_Home(((LinkButton)sender).Text);
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

        protected string Visibilidad(bool condicion)
        {
            if (condicion)
                return "visible";
            else
                return "hidden";
        }
       

    }
}