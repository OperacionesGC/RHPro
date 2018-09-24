using System;
using System.Data;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Xml.Linq;

namespace RHPro
{
    public class Gadget
    {
        /// <summary>
        /// Constructor de la clase
        /// </summary>
        public Gadget()
        {

        }


        public string GloboConfiguracion()
        {
          
            String Globo;
            Globo = "<table width='200' border='0' cellspacing='0' cellpadding='0'>";
            Globo += "  <tr>";
            Globo += "    <td width='49%' class='lineaPiso'>&nbsp;</td>";
            Globo += "    <td width='2%' align='center' valign='bottom'><img src='PuntaGlobo.png' align='bottom'></td>";
            Globo += "    <td width='49%' class='lineaPiso'>&nbsp;</td>";
            Globo += "  </tr>";
            Globo += "  <tr>";
            Globo += "    <td colspan='3' align='left' valign='top' class='lineasContenido'>";
            Globo += "    datos";
            Globo += "    </td>";
            Globo += "  </tr>";
            Globo += "</table>";

            return Globo;
        }

        public string TopeModulo(string Titulo, string width)
        {
            String Tope;

            Tope = " <table style='width:" + width + "' id='div1' border='0' cellspacing='0' cellpadding='0' align='center' class='BordeGris'  >";
            Tope += "        <tr> ";
            Tope += "             <td style='width:100%;'  valign='middle' align='left'><table width='100%' border='0' cellspacing='0' cellpadding='0' align='center'  >";
            Tope += "               <tr>";
            Tope += "                 <td style='width:100%;' class='PisoGris' valign='middle' align='center'  >";
            Tope += " <span style='margin-left:10px;'>" + Titulo + "</span>";
            Tope += " </td>";
            Tope += "  <td valign='middle' align='right' class='PisoGris' nowrap>";

            //Tope += "X<asp:ImageButton ID='ImageButton1' OnClick='G.IntercambiarPosicion(5,Siguiente_Gadget(5))' runat='server' ImageUrl='../img/Gizq.png' />X";
          
            //Tope += "<a onclick='Subir(5)'>aca</a>";
            Tope += "    <img src='~/../img/Gizq.png' style='cursor:pointer' align='absmiddle' onclick='Subir(5)' /> ";
            Tope += "    <img src='~/../img/Gder.png' style='cursor:pointer' align='absmiddle' onclick='Bajar(5)'/> ";
            Tope += "    <img src='~/../img/Configurar.png' style='cursor:pointer' align='absmiddle' /> </td>";
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

 




      

    }

}
