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


        ///// <summary>
        ///// Intercambia las posiciones de dos gadgets
        ///// </summary>
        //public void IntercambiarPosicion(int gadnro1, int gadnro2)
        //{
           

        //    try
        //    {
        //        ///*Actualizo la posicion del gadnro1 con la posicion del gadnro2*/
        //        //string updateSql1 = "UPDATE Gadgets set gadposicion = (select gadposicion from Gadgets where gadnro=" + gadnro2 + ") where gadnro=" + gadnro1;
        //        //System.Data.SqlClient.SqlCommand UpdateCmd1 = new System.Data.SqlClient.SqlCommand(updateSql1, conn);
        //        //UpdateCmd1.ExecuteNonQuery();

        //        ///*Actualizo la posicion del gadnro1 con la posicion del gadnro2*/
        //        //string updateSql2 = "UPDATE Gadgets set gadposicion = (select gadposicion from Gadgets where gadnro=" + gadposicion1 + ") where gadnro=" + gadnro2;
        //        //System.Data.SqlClient.SqlCommand UpdateCmd2 = new System.Data.SqlClient.SqlCommand(updateSql2, conn);
        //        //UpdateCmd2.ExecuteNonQuery();

        //    int gadposicion1;
        //    //Obtengo la posicion del primer gadtet antes de modificarlo
        //    gadposicion1 = get_Posicion(gadnro1);

        //    System.Data.SqlClient.SqlConnection conn;
        //    conn = new System.Data.SqlClient.SqlConnection();
        //    conn.ConnectionString = "Password=ess;Persist Security Info=True;User ID=ess;Initial Catalog=Base_0_R3_ARG;Data Source=RHDESA";

        //        /*Actualizo la posicion del gadnro1 con la posicion del gadnro2*/
        //        string updateSql1 = "UPDATE Gadgets set gadposicion = (select gadposicion from Gadgets where gadnro=@gadnro2) where gadnro=@gadnro1";
        //        using (System.Data.SqlClient.SqlConnection con = new System.Data.SqlClient.SqlConnection(conn.ConnectionString))
        //        {
        //            System.Data.SqlClient.SqlCommand cmd = new System.Data.SqlClient.SqlCommand(updateSql1, con);
        //            cmd.Parameters.AddWithValue("@gadnro1", gadnro1);
        //            cmd.Parameters.AddWithValue("@gadnro2", gadnro2);
                 
        //            con.Open();
        //            int t = cmd.ExecuteNonQuery();
        //            con.Close();
 
        //        }

        //        /*Actualizo la posicion del gadnro2 con la posicion del gadnro1*/
        //        string updateSql2 = "UPDATE Gadgets SET gadposicion = @gadposicion1 WHERE gadnro=@gadnro2";
        //        using (System.Data.SqlClient.SqlConnection con2 = new System.Data.SqlClient.SqlConnection(conn.ConnectionString))
        //        {
        //            System.Data.SqlClient.SqlCommand cmd2 = new System.Data.SqlClient.SqlCommand(updateSql2, con2);
        //            cmd2.Parameters.AddWithValue("@gadnro2", gadnro2);                   
        //            cmd2.Parameters.AddWithValue("@gadposicion1", gadposicion1);

        //            con2.Open();
        //            int t = cmd2.ExecuteNonQuery();
        //            con2.Close();

        //        }
        //    }
        //    catch (Exception EX) { }
        //}


        //public int get_Posicion(int gadnro) {
        //    int gadposicion = -1;
        //    //Se define el objeto conexión            
        //    System.Data.SqlClient.SqlConnection conn;
        //    System.Data.SqlClient.SqlDataReader reader;
        //    System.Data.SqlClient.SqlCommand sql;
        //    conn = new System.Data.SqlClient.SqlConnection();
        //    conn.ConnectionString = "Password=ess;Persist Security Info=True;User ID=ess;Initial Catalog=Base_0_R3_ARG;Data Source=RHDESA";
        //    conn.Open();
        //    /*busco la posicion del gadget gadnro1 */
        //    sql = new System.Data.SqlClient.SqlCommand();
        //    sql.CommandText = "SELECT gadnro,gadposicion FROM Gadgets WHERE gadnro = " + gadnro;
        //    sql.Connection = conn;             
        //    reader = sql.ExecuteReader();
        //    while (reader.Read())
        //    {
        //        gadposicion = (int)reader.GetValue(1);
        //    }
        //    return gadposicion;
        //}


        // /// <summary>
        ///// Retorna el gandnro del siguiente gadget segun la posicion
        ///// </summary>
        //public int Siguiente_Gadget(int pos)
        //{
        //    int gadnroSiguiente;
        //    //Se define el objeto conexión            
        //    System.Data.SqlClient.SqlConnection conn;
        //    System.Data.SqlClient.SqlDataReader reader;
        //    System.Data.SqlClient.SqlCommand sql;
        //    conn = new System.Data.SqlClient.SqlConnection();
        //    conn.ConnectionString = "Password=ess;Persist Security Info=True;User ID=ess;Initial Catalog=Base_0_R3_ARG;Data Source=RHDESA";
        //    conn.Open();
        //    /*busco la posicion del gadget gadnro1 */
        //    sql = new System.Data.SqlClient.SqlCommand();
        //    sql.CommandText = "select top(1) gadnro gadnroSig  from Gadgets where gadposicion > " + pos + " order by gadposicion asc ";
        //    sql.Connection = conn;

        //    reader = sql.ExecuteReader();
        //    gadnroSiguiente = -1;
        //    while (reader.Read())
        //    {
        //        gadnroSiguiente = (int)reader.GetValue(0);
        //    }

        //    return gadnroSiguiente;
        //}

        ///// <summary>
        ///// Retorna el gandnro del gadget anterior segun la posicion
        ///// </summary>
        //public int Anterior_Gadget(int pos)
        //{
        //    int gadnroAnt;
        //    //Se define el objeto conexión            
        //    System.Data.SqlClient.SqlConnection conn;
        //    System.Data.SqlClient.SqlDataReader reader;
        //    System.Data.SqlClient.SqlCommand sql;
        //    conn = new System.Data.SqlClient.SqlConnection();
        //    conn.ConnectionString = "Password=ess;Persist Security Info=True;User ID=ess;Initial Catalog=Base_0_R3_ARG;Data Source=RHDESA";
        //    conn.Open();
        //    /*busco la posicion del gadget gadnro1 */
        //    sql = new System.Data.SqlClient.SqlCommand();
        //    sql.CommandText = "select top(1) gadnro gadnroSig from Gadgets where gadposicion < " + pos + " order by gadposicion desc ";
        //    sql.Connection = conn;

        //    reader = sql.ExecuteReader();
        //    gadnroAnt = -1;
        //    while (reader.Read())
        //    {
        //        gadnroAnt = (int)reader.GetValue(0);
        //    }

        //    return gadnroAnt;
        //}






      

    }

}
