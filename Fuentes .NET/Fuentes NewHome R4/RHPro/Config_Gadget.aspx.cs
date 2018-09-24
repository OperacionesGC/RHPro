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
using System.Data.SqlClient;
using System.Data.OleDb;
using ServicesProxy;

 


namespace RHPro
{
    public partial class Config_Gadget : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }


        public bool AjustarAlto(int gadnro, int valor, int gadtipo)
        {
            Consultas cc = new Consultas();
            string Base = Common.Utils.SessionBaseID;
            return cc.Update_Dimension_Gadget(gadnro, valor, 1, Base, gadtipo, Convert.ToInt32(Common.Utils.Session_MenumsNro_Modulo));
        }

        public bool AjustarAncho(int gadnro, int valor, int gadtipo)
        {
            Consultas cc = new Consultas();
            string Base = Common.Utils.SessionBaseID;
            return cc.Update_Dimension_Gadget(gadnro, valor, 2, Base, gadtipo,Convert.ToInt32(Common.Utils.Session_MenumsNro_Modulo));           

        }

        /// <summary>
        /// Intercambia las posiciones de dos gadgets
        /// </summary>
        public bool IntercambiarPosicion(int gadnro1, int gadnro2)
        { 
            //string sql = "";
            Consultas cc = new Consultas();
            string Base = Common.Utils.SessionBaseID;
            //string ConnStr = cc.constr(Base);
            //System.Data.OleDb.OleDbConnection cn = new System.Data.OleDb.OleDbConnection();
            //cn.ConnectionString = ConnStr;
            Int32 gadposicion1,gadposicion2;
            //Obtengo la posicion del primer gadtet antes de modificarlo           
            gadposicion1 = get_Posicion(gadnro1);            
            //Obtengo la posicion del segundo gadtet antes de modificarlo           
            gadposicion2 = get_Posicion(gadnro2);            

            try
            {
                cc.Update_Pos_Gadget(gadnro1, gadposicion2, Base);
                cc.Update_Pos_Gadget(gadnro2, gadposicion1, Base);

                               
            }
            catch
            {
                return false;
            }
            finally
            {
                //if (cn.State == System.Data.ConnectionState.Open) cn.Close();
            }

            return true;

        }

        /// <summary>
        /// Dado el numero de gadget devuelve la posicion del gadget
        /// </summary>
        public Int32 get_Posicion(int gadnro)
        {
            Int32 gadposicion = -1;
            try
            {
               // string sql = "SELECT gadposicion FROM Gadgets WHERE gadnro = " + gadnro;
                string sql = "SELECT gadusrposicion FROM Gadgets_User WHERE gadusrnro = " + gadnro;
                Consultas cc = new Consultas();
                //Paso las credenciales al web service
                //cc.Credentials = System.Net.CredentialCache.DefaultCredentials;
                //-----------------------------------------------------------
                System.Data.DataTable dt = cc.get_DataTable(sql, Common.Utils.SessionBaseID);
                foreach (System.Data.DataRow dr in dt.Rows)
                {                     
                   //gadposicion = (int)dr["gadposicion"];                   
                   gadposicion = Convert.ToInt32(dr["gadusrposicion"]);                
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
        public Int32 Siguiente_Gadget(int pos)
        {
            Int32 gadnroSiguiente = -1;

            try
            {  
                Consultas cc = new Consultas();              
                //string sql = "SELECT gadnro gadnroSig  FROM Gadgets WHERE gadposicion < " + pos + "  AND gadactivo=-1  AND gaduser='" + Common.Utils.SessionUserName + "' order by gadposicion DESC ";
                String TipoBD = cc.get_TipoBase(Common.Utils.SessionBaseID);

                string sql = "SELECT ";
                
                if (TipoBD == "MSSQL")
                    sql += " top(1) ";
                
                sql +=" gadusrnro gadnroSig  FROM Gadgets_User WHERE gadusrposicion < " + pos + "  AND gadusractivo=-1  AND iduser ='" + Common.Utils.SessionUserName + "' ";
                
                if (TipoBD == "ORA")
                    sql += " and rownum = 1";
                
                sql += " order by gadusrposicion DESC ";
               
                //Paso las credenciales al web service
                //cc.Credentials = System.Net.CredentialCache.DefaultCredentials;
                //-----------------------------------------------------------
                System.Data.DataTable dt = cc.get_DataTable(sql, Common.Utils.SessionBaseID);
                foreach (System.Data.DataRow dr in dt.Rows)
                {

                    gadnroSiguiente = Convert.ToInt32(dr["gadnroSig"]);
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
        public Int32 Anterior_Gadget(int pos)
        {
            Int32 gadnroAnt = -1;
            try
            {
                Consultas cc = new Consultas();
                String TipoBD = cc.get_TipoBase(Common.Utils.SessionBaseID);
               //string sql = "SELECT gadnro gadnroAnt FROM Gadgets WHERE gadposicion > " + pos + " AND gadactivo=-1 AND gaduser='" + Common.Utils.SessionUserName + "' ORDER BY gadposicion ASC ";
                string sql = "SELECT ";
                
                if (TipoBD == "MSSQL")
                sql += " top(1) ";
                
                sql += " gadusrnro gadnroAnt  FROM Gadgets_User WHERE gadusrposicion > " + pos + "  AND gadusractivo=-1  AND iduser ='" + Common.Utils.SessionUserName + "'";

                if (TipoBD == "ORA")
                    sql += " and rownum = 1";

                sql +=" order by gadusrposicion ASC ";
                
                //Paso las credenciales al web service
                //cc.Credentials = System.Net.CredentialCache.DefaultCredentials;
                //-----------------------------------------------------------
                System.Data.DataTable dt = cc.get_DataTable(sql, Common.Utils.SessionBaseID);
                foreach (System.Data.DataRow dr in dt.Rows)
                {
                    gadnroAnt = Convert.ToInt32(dr["gadnroAnt"]);
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
        public Int32 get_Max_Posicion()
        {
            Int32 maxposicion = -1;
            try
            {                
                //string sql = "SELECT  cast(MAX(gadposicion)  AS int) maxpos FROM Gadgets  WHERE  gadactivo=-1 AND gaduser='" + Common.Utils.SessionUserName + "'";
                string sql = "SELECT cast(MAX(gadusrposicion)  AS int) maxpos  FROM Gadgets_User WHERE iduser='" + Common.Utils.SessionUserName + "'";
                 
                Consultas cc = new Consultas();
                //Paso las credenciales al web service
                //cc.Credentials = System.Net.CredentialCache.DefaultCredentials;
                //-----------------------------------------------------------
                System.Data.DataTable dt = cc.get_DataTable(sql, Common.Utils.SessionBaseID);
                foreach (System.Data.DataRow dr in dt.Rows)
                {

                    maxposicion = Convert.ToInt32(dr["maxpos"]);
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
        public Int32 get_Min_Posicion()
        {
            Int32 minposicion = -1;
            try
            {               
//                string sql = "SELECT  CAST(MIN(gadposicion) AS int) minpos FROM Gadgets WHERE  gadactivo=-1 AND gaduser='" + Common.Utils.SessionUserName + "'";                
                string sql = "SELECT  CAST(MIN(gadusrposicion) AS int) minpos FROM Gadgets_User WHERE  iduser = '" + Common.Utils.SessionUserName + "'";                
                Consultas cc = new Consultas();
                //Paso las credenciales al web service
                //cc.Credentials = System.Net.CredentialCache.DefaultCredentials;
                //-----------------------------------------------------------
                System.Data.DataTable dt = cc.get_DataTable(sql, Common.Utils.SessionBaseID);
                foreach (System.Data.DataRow dr in dt.Rows)
                {

                    minposicion = Convert.ToInt32(dr["minpos"]);
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

            Consultas cc = new Consultas();
            //cc.Credentials = System.Net.CredentialCache.DefaultCredentials;
            try
            {              
                return cc.Act_Desact_Gadget(gadnro, 0, Common.Utils.SessionBaseID);
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        /// <summary>
        /// Activa el gadget
        /// </summary>
        public bool Activar(int gadnro)
        {
            Consultas cc = new Consultas();
            //cc.Credentials = System.Net.CredentialCache.DefaultCredentials;
            try
            {                
                return cc.Act_Desact_Gadget(gadnro, -1, Common.Utils.SessionBaseID);
            }
            catch (Exception ex)
            {
                return false;
            }
             

        }
 


    }
}
