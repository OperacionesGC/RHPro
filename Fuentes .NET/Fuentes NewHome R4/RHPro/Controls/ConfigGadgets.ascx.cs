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
using ServicesProxy.rhdesa;
using Common;

namespace RHPro.Controls
{
    public partial class ConfigGadgets : System.Web.UI.UserControl
    {
        public Lenguaje Obj_Lenguaje;
        public int ContadorGadget;
        public static Default Padre = new Default();
        public ConsultaDatos c_datos;
        protected void Page_Load(object sender, EventArgs e)
        {   
            if (Utils.IsUserLogin)
            {   c_datos = new ConsultaDatos();
                Obj_Lenguaje = new Lenguaje();
                Consultas cc = new Consultas();
                string sql = "";
                string BaseId = Common.Utils.SessionBaseID;
                string UserName = Common.Utils.SessionUserName;
                string GadgetPermitidos = Padre.Lista_Gadget_Permitidos(Utils.SessionUserName);
                //Response.Write("ZZ_"+GadgetPermitidos);
                //Busca todos los gadget habilitados y que tengan estado -1 en la relacion gadget-usuario
                if (GadgetPermitidos != "")
                {
                    sql = " SELECT  DISTINCT GU.* ";
                    sql += "  ,(select GT.gadactivo from Gadgets_Tipo GT where GT.gadnro=GU.gadnro) gadactivo ";
                    sql += "  ,(select GT.gadURL from Gadgets_Tipo GT where GT.gadnro=GU.gadnro) gadURL ";
                    sql += "  ,(select GT.gaddesabr from Gadgets_Tipo GT where GT.gadnro=GU.gadnro) gaddesabr ";
                    sql += "  ,(select GT.gadtitulo from Gadgets_Tipo GT where GT.gadnro=GU.gadnro) gadtitulo ";
                    sql += " FROM Gadgets_User GU  ";
                    sql += "  WHERE GU.gadestado=-1    ";
                    sql += "  AND   GU.gadnro in (" + GadgetPermitidos + ") ";
                    sql += "  AND (select GT.gadactivo from Gadgets_Tipo GT where GT.gadnro=GU.gadnro) = -1 ";
                    sql += "  AND   iduser='" + UserName + "'";
                    sql += "  ORDER BY gaddesabr ASC";

                    DataSet ds = cc.get_DataSet(sql, BaseId);
                    Repeater1.DataSource = ds;
                    Repeater1.DataBind();
                }
            }
        }

 

        public string Imprimir_Slider(int gadnro, int gadactivo)
        {
            String salida = "";

            if (gadactivo == 0)
            {
                salida = "<table border='0' cellspacing='0' cellpadding='0' onclick='ActivarGadget(" + gadnro + ")' class='SLIDER' title='" + Obj_Lenguaje.Label_Home("Activar") + "'>";
                salida += "  <tr>";
                salida += " <td  class='SliderClaroOFF'> ";
                salida += Obj_Lenguaje.Label_Home("OFF");
                salida += "  </td>";
                salida += "<td class='SliderOscuroOFF'>&nbsp;";
                salida += "</td>";
                salida += "</tr>";
            }
            else
            {
                salida = "<table border='0' cellspacing='0' cellpadding='0' onclick='DesactivarGadget(" + gadnro + ")' class='SLIDER' title='" + Obj_Lenguaje.Label_Home("Desactivar") + "'>";
                salida += "  <tr>";
                salida += "<td class='SliderOscuroON'>&nbsp;";
                salida += "</td>";
                salida += " <td  class='SliderClaroON'> ";
                salida += Obj_Lenguaje.Label_Home("ON");
                salida += " </td>";

                salida += "</tr>";
            }

            salida += "  </table>";
            
          
            return salida;
        }

        public string Imprimir_Led(Int32 gadusractivo)
        {
          
            String salida = "";             
            return salida;
        }

        public void InicializarPadre(Default P)
        {
            Padre = P;

        }




    }


    
}