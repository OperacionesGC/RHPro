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

namespace RHPro
{
    public partial class Config_Gadget : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }


        /// <summary>
        /// Intercambia las posiciones de dos gadgets
        /// </summary>
        public bool IntercambiarPosicion(int gadnro1, int gadnro2)
        { 
            string sql = "";
            Consultas cc = new Consultas();
            string ConnStr = cc.constr(Common.Utils.SessionBaseID);
            System.Data.OleDb.OleDbConnection cn = new System.Data.OleDb.OleDbConnection();
            cn.ConnectionString = ConnStr;
            Int32 gadposicion1;
            //Obtengo la posicion del primer gadtet antes de modificarlo           
            gadposicion1 = get_Posicion(gadnro1);            

            try
            {
                cn.Open();

                sql = "UPDATE Gadgets set gadposicion = (select gadposicion from Gadgets where gadnro=" + gadnro2 + ") where gadnro=" + gadnro1;
                System.Data.OleDb.OleDbCommand cmd = new System.Data.OleDb.OleDbCommand(sql, cn);
                cmd.ExecuteNonQuery();

                sql = "UPDATE Gadgets SET gadposicion = " + gadposicion1 + " WHERE gadnro = " + gadnro2;
                System.Data.OleDb.OleDbCommand cmd2 = new System.Data.OleDb.OleDbCommand(sql, cn);
                cmd2.ExecuteNonQuery();                
            }
            catch
            {
                return false;
            }
            finally
            {
                if (cn.State == System.Data.ConnectionState.Open) cn.Close();
            }

            return true;

        }

        /// <summary>
        /// Dado el numero de gadget devuelve la posicion del gadget
        /// </summary>
        public int get_Posicion(int gadnro)
        {
            int gadposicion = -1;
            try
            {
                string sql = "SELECT gadposicion FROM Gadgets WHERE gadnro = " + gadnro;
                Consultas cc = new Consultas();
                System.Data.DataTable dt = cc.get_DataTable(sql, Common.Utils.SessionBaseID);
                foreach (System.Data.DataRow dr in dt.Rows)
                {                     
                   //gadposicion = (int)dr["gadposicion"];                   
                   gadposicion = Convert.ToInt32(dr["gadposicion"]);                
                }
            }
            catch (Exception ex)
            {
             
                throw ex;
            }

            return gadposicion;
        }


        /// <summary>
        /// Retorna el gandnro del siguiente gadget segun la posicion
        /// </summary>
        public int Siguiente_Gadget(int pos)
        {
            int gadnroSiguiente = -1;

            try
            {
                string sql = "SELECT top(1) gadnro gadnroSig  FROM Gadgets WHERE gadposicion < " + pos + "  AND gadactivo=-1  AND gaduser='" + Common.Utils.SessionUserName + "' order by gadposicion DESC ";
                Consultas cc = new Consultas();
                System.Data.DataTable dt = cc.get_DataTable(sql, Common.Utils.SessionBaseID);
                foreach (System.Data.DataRow dr in dt.Rows)
                {

                    gadnroSiguiente = (int)dr["gadnroSig"];
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return gadnroSiguiente;
        }

        /// <summary>
        /// Retorna el gandnro del gadget anterior segun la posicion
        /// </summary>
        public int Anterior_Gadget(int pos)
        {
            int gadnroAnt = -1;
            try
            {
                string sql = "SELECT  top(1) gadnro gadnroAnt FROM Gadgets WHERE gadposicion > " + pos + " AND gadactivo=-1 AND gaduser='" + Common.Utils.SessionUserName + "' ORDER BY gadposicion ASC ";
                Consultas cc = new Consultas();
                System.Data.DataTable dt = cc.get_DataTable(sql, Common.Utils.SessionBaseID);
                foreach (System.Data.DataRow dr in dt.Rows)
                {
                    gadnroAnt = (int)dr["gadnroAnt"];
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return gadnroAnt;
        }


        /// <summary>
        /// Rotorna la posicion maxima de los gadgets
        /// </summary>
        public int get_Max_Posicion()
        {
            int maxposicion = -1;
            try
            {
                string sql = "SELECT  MAX(gadposicion) maxpos FROM Gadgets SELECT  MIN(gadposicion) minpos FROM Gadgets WHERE  gadactivo=-1 AND gaduser='" + Common.Utils.SessionUserName + "'";
                Consultas cc = new Consultas();
                System.Data.DataTable dt = cc.get_DataTable(sql, Common.Utils.SessionBaseID);
                foreach (System.Data.DataRow dr in dt.Rows)
                {

                    maxposicion = (int)dr["maxpos"];
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }
            return maxposicion;
        }

        /// <summary>
        /// Rotorna la posicion minima de los gadgets
        /// </summary>
        public int get_Min_Posicion()
        {
            int minposicion = -1;
            try
            {
                string sql = "SELECT  MIN(gadposicion) minpos FROM Gadgets WHERE  gadactivo=-1 AND gaduser='" + Common.Utils.SessionUserName + "'";
                Consultas cc = new Consultas();
                System.Data.DataTable dt = cc.get_DataTable(sql, Common.Utils.SessionBaseID);
                foreach (System.Data.DataRow dr in dt.Rows)
                {

                    minposicion = (int)dr["minpos"];
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return minposicion;
        }

        /// <summary>
        /// Desactiva el gadget
        /// </summary>
        public bool Desactivar(int gadnro)
        {
            string sql = "";
            Consultas cc = new Consultas();
            string ConnStr = cc.constr(Common.Utils.SessionBaseID);
            System.Data.OleDb.OleDbConnection cn = new System.Data.OleDb.OleDbConnection();
            cn.ConnectionString = ConnStr;
            try
            {
                cn.Open();
                sql = "UPDATE Gadgets SET gadactivo = 0 WHERE gadnro=" + gadnro;
                System.Data.OleDb.OleDbCommand cmd = new System.Data.OleDb.OleDbCommand(sql, cn);
                cmd.ExecuteNonQuery();
            }
            catch
            {
                return false;
            }
            finally
            {
                if (cn.State == System.Data.ConnectionState.Open) cn.Close();
            }

            return true;


        }

        /// <summary>
        /// Activa el gadget
        /// </summary>
        public bool Activar(int gadnro)
        {
            string sql = "";
            Consultas cc = new Consultas();
            string ConnStr = cc.constr(Common.Utils.SessionBaseID);
            System.Data.OleDb.OleDbConnection cn = new System.Data.OleDb.OleDbConnection();
            cn.ConnectionString = ConnStr;
            try
            {
                cn.Open();
                sql = "UPDATE Gadgets SET gadactivo = -1 WHERE gadnro=" + gadnro;
                System.Data.OleDb.OleDbCommand cmd = new System.Data.OleDb.OleDbCommand(sql, cn);
                cmd.ExecuteNonQuery();
            }
            catch
            {
                return false;
            }
            finally
            {
                if (cn.State == System.Data.ConnectionState.Open) cn.Close();
            }

            return true;

        }
 


    }
}
