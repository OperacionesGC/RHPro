using System;
using System.Data;
using System.Web;
using System.Data.OleDb;

namespace ConsultaBaseC
{
    public class EtiquetasMI{

        /*******************************************************************/
        //DEVUELVE EL MENSAJE DE ERROR EN EL IDIOMA CORRESPONDIENTE
        /*******************************************************************/
 

        public static string EtiquetaErr(string etiqueta, string lenguaje, string Base)
        {
            //Defino un objeto tipo DAL
            //DAL MiDAL = new DAL(); 

            string cn = DAL.constr(Base);
            string sql;
            DataSet ds = new DataSet();
            string Salida = etiqueta;
            OleDbDataAdapter daPass;
            string lang = lenguaje.Replace("-", "");


            if ((etiqueta != null) && (etiqueta != "") && (lenguaje != null) && (lenguaje != ""))
            {

                sql = "SELECT esAR," + lang;
                sql = sql + " FROM lenguaje_etiqueta ";
                sql = sql + " WHERE etiqueta = '" + etiqueta + "'";
                sql = sql + " AND modulo = 'HOME'";

                daPass = new OleDbDataAdapter(sql, cn);

                try
                {
                    daPass.Fill(ds);
                    if (ds.Tables[0].Rows.Count > 0)
                    {
                        if ((ds.Tables[0].Rows[0].ItemArray[1] != DBNull.Value) && ((String)ds.Tables[0].Rows[0].ItemArray[1] != ""))
                        {   //Devuelve traducción del idioma seleccionado.
                            Salida = Convert.ToString(ds.Tables[0].Rows[0].ItemArray[1]);
                        }
                        else
                        {
                            //Devuelve esAR cuando no encontro la traducción.
                            Salida = Convert.ToString(ds.Tables[0].Rows[0].ItemArray[0]);
                        }

                    }

                }
                catch (Exception ex)
                {
                    //throw ex;
                }
            }



            return Salida;
        }

        /**********************************/
    }
        
}