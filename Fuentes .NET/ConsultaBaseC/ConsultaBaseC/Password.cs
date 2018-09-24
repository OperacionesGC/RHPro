using System;
using System.Data;
using System.Web;
using System.Data.OleDb;
using System.Text.RegularExpressions;
using System.Diagnostics;

namespace ConsultaBaseC
{
    public class Password
    {
        

        public static string valorUserPolCuenta(string User, string Campo, string Base)
        {
            //Defino un objeto tipo DAL
            //DAL MiDAL = new DAL();   

            string cn = DAL.constr(Base);
            string sql;
            DataSet ds = new DataSet();
            string Salida = "";
            OleDbDataAdapter daPass;

            sql = "SELECT usr_pol_cuenta." + Campo + " ";
            sql = sql + "FROM usr_pol_cuenta ";
            sql = sql + "WHERE upper(usr_pol_cuenta.iduser) = '" + User.ToUpper() + "' ";
            sql = sql + "AND upcfecfin IS NULL ";

            daPass = new OleDbDataAdapter(sql, cn);

            try
            {
                daPass.Fill(ds);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            if (ds.Tables[0].Rows.Count > 0)
                Salida = Convert.ToString(ds.Tables[0].Rows[0].ItemArray[0]);

            return Salida;
        }

        public static string valorPolCuenta(long PolNro, string Campo, string Base)
        {
            //Defino un objeto tipo DAL
            //DAL MiDAL = new DAL();   

            string cn = DAL.constr(Base);
            string sql;
            DataSet ds = new DataSet();
            string Salida = "";
            OleDbDataAdapter daPass;

            sql = "SELECT pol_cuenta." + Campo + " ";
            sql = sql + "FROM pol_cuenta ";
            sql = sql + "WHERE pol_cuenta.pol_nro = " + PolNro + " ";

            daPass = new OleDbDataAdapter(sql, cn);

            try
            {
                daPass.Fill(ds);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            if (ds.Tables[0].Rows.Count > 0)
                Salida = Convert.ToString(ds.Tables[0].Rows[0].ItemArray[0]);

            return Salida;
        }

        //LED - Funcion que chequea si el password contiene caracteres especiales
        public static int contieneCaractEspeciales(string pass)
        {
            int contiene, pos, i;

            string caractEspeciales = "{}.<>;:?/|`~!@#$%^&*()_-+=";

            contiene = 0;
            i = 0;
            while (i < caractEspeciales.Length && contiene == 0)
            {
                pos = pass.IndexOf(caractEspeciales[i], 0);
                if (pos > 0)
                    contiene = -1;
                i++;
            }
            return contiene;
        }

        public static int contieneMinusculas(string pass)
        {
            int contiene, pos;
            Regex isnumber = new Regex(@"[a-z]");

            contiene = 0;

            if (isnumber.IsMatch(pass))
                contiene = -1;

            return contiene;

        }

        public static int contieneMayusculas(string pass)
        {
            int contiene, pos;
            Regex isnumber = new Regex(@"[A-Z]");

            contiene = 0;

            if (isnumber.IsMatch(pass))
                contiene = -1;                           

            return contiene;
        }

        public static int contieneNumeros(string pass)
        {
            int contiene, pos;
            Regex isnumber = new Regex(@"[0-9]");

            contiene = 0;

            if (isnumber.IsMatch(pass))
                contiene = -1;

            return contiene;

        }

         
        

        public static string valorUserPer(string User, string Campo, string Base)
        {
            //Defino un objeto tipo DAL
            //DAL MiDAL = new DAL();   

            string cn = DAL.constr(Base);
            string sql;
            DataSet ds = new DataSet();
            string Salida = "";
            OleDbDataAdapter daPass;

            sql = "SELECT user_per." + Campo + " ";
            sql = sql + "FROM user_per ";
            sql = sql + "WHERE upper(user_per.iduser) = '" + User.ToUpper() + "' ";
            
            daPass = new OleDbDataAdapter(sql, cn);

            try
            {
                daPass.Fill(ds);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            if (ds.Tables[0].Rows.Count > 0)
                Salida = Convert.ToString(ds.Tables[0].Rows[0].ItemArray[0]);

            return Salida;
        }

        public static bool usuarioValido(string User, string Base)
        {
            //Defino un objeto tipo DAL
            //DAL MiDAL = new DAL();   

            string cn = DAL.constr(Base);
            string sql;
            DataSet ds = new DataSet();
            bool Salida = false;
            OleDbDataAdapter daSal;            

            sql = "SELECT user_per.iduser ";
             
            sql = sql + "FROM user_per ";
            sql = sql + "WHERE upper(user_per.iduser) = '" + User.ToUpper() + "' ";
 
            daSal = new OleDbDataAdapter(sql, cn); 

            try
            {                 
               daSal.Fill(ds);                
               
            }
            catch (Exception ex)
            {
                throw ex;
            }

            if (ds.Tables[0].Rows.Count > 0)
                Salida = true;

            return Salida;
        }

        public static bool ctaBloqueada(string User, string Base)
        {
            //Defino un objeto tipo DAL
            //DAL MiDAL = new DAL();   

            string cn = DAL.constr(Base);
            string sql;
            DataSet ds = new DataSet();
            bool Salida = false;
            OleDbDataAdapter daSal;

            sql = "SELECT user_per.ctabloqueada ";
            sql = sql + "FROM user_per ";
            sql = sql + "WHERE upper(user_per.iduser) = '" + User.ToUpper() + "' ";

            daSal = new OleDbDataAdapter(sql, cn);

            try
            {
                daSal.Fill(ds);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            if (ds.Tables[0].Rows.Count > 0)
            {
                Salida = Convert.ToString(ds.Tables[0].Rows[0].ItemArray[0]) == "-1";
            }

            return Salida;
        }

        public static string valorHistPass(string User, string Campo, string Base)
        {
            //Defino un objeto tipo DAL
            //DAL MiDAL = new DAL();   

            string cn = DAL.constr(Base);
            string sql;
            DataSet ds = new DataSet();
            string Salida = "";
            OleDbDataAdapter daPass;

            sql = "SELECT hist_pass_usr." + Campo + " ";
            sql = sql + "FROM hist_pass_usr ";
            sql = sql + "WHERE upper(hist_pass_usr.iduser) = '" + User.ToUpper() + "' ";
            sql = sql + "AND hpassfecfin IS NULL ";

            daPass = new OleDbDataAdapter(sql, cn);

            try
            {
                daPass.Fill(ds);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            if (ds.Tables[0].Rows.Count > 0)
                Salida = Convert.ToString(ds.Tables[0].Rows[0].ItemArray[0]);

            return Salida;
        }

        public static long logueosFallidos(string User, string Base)
        {
            //Defino un objeto tipo DAL
            //DAL MiDAL = new DAL();   

            string cn = DAL.constr(Base);
            string sql;
            DataSet ds = new DataSet();
            long Salida = 0;
            OleDbDataAdapter daPass;

            sql = "SELECT hlogintfallidos ";
            sql = sql + "FROM hist_log_usr ";
            sql = sql + "WHERE upper(hist_log_usr.iduser) = '" + User.ToUpper() + "' ";
            
            daPass = new OleDbDataAdapter(sql, cn);

            try
            {
                daPass.Fill(ds);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            if (ds.Tables[0].Rows.Count > 0)
                if(Convert.ToString(ds.Tables[0].Rows[0].ItemArray[0]).Length > 0)
                    Salida = Convert.ToInt64(ds.Tables[0].Rows[0].ItemArray[0]);

            return Salida;
        }

        public static void bloquearCuenta(string User, string Valor, string Base)
        {
            //Defino un objeto tipo DAL
            //DAL MiDAL = new DAL();   

            OleDbConnection cn = new OleDbConnection();
            cn.ConnectionString = DAL.constr(Base);

            try
            {
                cn.Open();

                string sql = "UPDATE user_per SET ctabloqueada = " + Valor + " WHERE upper(iduser) = '" + User.ToUpper() + "' ";
                
                OleDbCommand cmd = new OleDbCommand(sql, cn);

                cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            { 
                if (cn.State == ConnectionState.Open) cn.Close(); 
            }
        }

        public static void bajarCuenta(string User, string Base)
        {
            //Defino un objeto tipo DAL
            //DAL MiDAL = new DAL();   

            OleDbConnection cn = new OleDbConnection();
            cn.ConnectionString = DAL.constr(Base);

            string FechaAct = DateTime.Today.ToString("dd/MM/yyyy");
            string HoraAct = DateTime.Now.ToString("hh:mm:ss");

            try
            {
                cn.Open();


                string sql = "UPDATE hist_pass_usr SET hpassfecfin = " + Fecha.cambiaFecha(FechaAct, DAL.TipoBase(Base)) + " ";
                sql = sql + ", hpasshorafin = '" + HoraAct + "' ";
                sql = sql + "WHERE upper(iduser) = '" + User.ToUpper() + "' ";
                sql = sql + "AND hpassfecfin IS NULL ";

                OleDbCommand cmd = new OleDbCommand(sql, cn);

                cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (cn.State == ConnectionState.Open) cn.Close();
            }
        }

        public static void actLogFallidos(string User, long Cant, string Base)
        {
            //Defino un objeto tipo DAL
            //DAL MiDAL = new DAL();   

            OleDbConnection cn = new OleDbConnection();
            cn.ConnectionString = DAL.constr(Base);

            try
            {
                cn.Open();

                string sql = "UPDATE hist_log_usr SET hlogintfallidos = " + Cant.ToString() + " ";
                sql = sql + "WHERE upper(iduser) = '" + User.ToUpper() + "' ";

                OleDbCommand cmd = new OleDbCommand(sql, cn);

                cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (cn.State == ConnectionState.Open) cn.Close();
            }
        }

        public static void ActLogFallidos_NTUser_IP(string Base, Boolean EsNuevo, String AUTH_USER, String REMOTE_ADDR, Boolean LimpiarIP)
        {            
            OleDbConnection cn = new OleDbConnection();
            cn.ConnectionString = DAL.constr(Base);

            try
            {
                cn.Open();

                string FechaAct = DateTime.Today.ToString("dd/MM/yyyy");
                string HoraAct = DateTime.Now.ToString("hh:mm:ss");


                string sql = "";
                if (LimpiarIP)
                {
                    sql = " DELETE rhpro_seg_login ";
                    sql += " WHERE rhseglogip = '" + REMOTE_ADDR + "' AND rhseglogpc ='" + AUTH_USER + "' ";
                }
                else
                {

                    if (EsNuevo)
                    {
                        sql = " INSERT INTO rhpro_seg_login (rhseglogip, rhseglogpc, rhseglogHost, appnro, rhseglogfec, rhsegloghora, rhseglogcant) ";
                        sql += "  VALUES ('" + REMOTE_ADDR + "','" + AUTH_USER + "','', 1 , " + Fecha.cambiaFecha(FechaAct, DAL.TipoBase(Base));
                        sql += " ,'" + HoraAct + "',1)  ";
                    }
                    else
                    {
                        sql = "UPDATE rhpro_seg_login SET rhseglogcant = rhseglogcant + 1, rhseglogfec= " + Fecha.cambiaFecha(FechaAct, DAL.TipoBase(Base));
                        sql += "WHERE rhseglogip = '" + REMOTE_ADDR + "' AND rhseglogpc ='" + AUTH_USER + "' ";
                    }
                }
 


                OleDbCommand cmd = new OleDbCommand(sql, cn);

                cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                
                throw ex;
            }
            finally
            {
                if (cn.State == ConnectionState.Open) cn.Close();
            }
        }



        public static string valorHistLog(string User, string Campo, string Base)
        {

            //Defino un objeto tipo DAL
            //DAL MiDAL = new DAL();   

            string cn = DAL.constr(Base);
            string sql;
            DataSet ds = new DataSet();
            string Salida = "";
            OleDbDataAdapter daPass;

            sql = "SELECT hist_log_usr." + Campo + " ";
            sql = sql + "FROM hist_log_usr ";
            sql = sql + "WHERE upper(hist_log_usr.iduser) = '" + User.ToUpper() + "' ";

            daPass = new OleDbDataAdapter(sql, cn);

            try
            {
                daPass.Fill(ds);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            if (ds.Tables[0].Rows.Count > 0)
                Salida = Convert.ToString(ds.Tables[0].Rows[0].ItemArray[0]);

            return Salida;
        }
        //----
        //----
        public static void ValidaHis_log_usr(string User, string Base)
        {
            //Defino un objeto tipo DAL
            //DAL MiDAL = new DAL(); 

            string TipoDB = DAL.TipoBase(Base);
            OleDbConnection cn = new OleDbConnection();
            cn.ConnectionString = DAL.constr(Base);

            DataSet ds = new DataSet();

            OleDbDataAdapter daLg;

            string FechaAct = DateTime.Today.ToString("dd/MM/yyyy");
            string HoraAct = DateTime.Now.ToString("hh:mm:ss");

            //cn.Open();

            string sql = "SELECT hlognro FROM hist_log_usr ";
            sql = sql + "WHERE upper(iduser) = '" + User.ToUpper() + "' ";

            daLg = new OleDbDataAdapter(sql, cn);
            try
            {
                daLg.Fill(ds);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            if (ds.Tables[0].Rows.Count > 0)
            {
            }
            else
            {
                sql = "INSERT INTO hist_log_usr (iduser, hlogfecini, hloghoraini, hlogintfallidos) ";
                sql = sql + "VALUES ( ";
                sql = sql + "'" + User.ToUpper() + "' ";
                sql = sql + "," + Fecha.cambiaFecha(FechaAct, TipoDB) + " ";
                sql = sql + ",'" + HoraAct + "' ";
                sql = sql + ",0) ";

                cn.Open();
                OleDbCommand cmd = new OleDbCommand(sql, cn);
                cmd.ExecuteNonQuery();

            }

            if (cn.State == ConnectionState.Open) cn.Close();
        }
        //----
        //----

        public static void ingresarLogueo(string User, string Base)
        {
            //Defino un objeto tipo DAL
            //DAL MiDAL = new DAL(); 

            string TipoDB = DAL.TipoBase(Base);
            OleDbConnection cn = new OleDbConnection();
            cn.ConnectionString = DAL.constr(Base);

            DataSet ds = new DataSet();

            OleDbDataAdapter daLg;

            string FechaAct = DateTime.Today.ToString("dd/MM/yyyy");
            string HoraAct = DateTime.Now.ToString("hh:mm:ss");

            //cn.Open();

            string sql = "SELECT hlognro FROM hist_log_usr ";
            sql = sql + "WHERE upper(iduser) = '" + User.ToUpper() + "' ";

            daLg = new OleDbDataAdapter(sql, cn);
            try
            {
                daLg.Fill(ds);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            if (ds.Tables[0].Rows.Count > 0)
            {
                sql = "UPDATE hist_log_usr SET hlogfecini = " + Fecha.cambiaFecha(FechaAct, TipoDB) + " ";
                sql = sql + ",hloghoraini = '" + HoraAct + "' ";
                sql = sql + ",hlogintfallidos = 0 ";
                sql = sql + "WHERE upper(iduser) = '" + User.ToUpper() + "' ";

                cn.Open();
                OleDbCommand cmd = new OleDbCommand(sql, cn);
                cmd.ExecuteNonQuery();


            }
            else
            {
                sql = "INSERT INTO hist_log_usr (iduser, hlogfecini, hloghoraini, hlogintfallidos) ";
                sql = sql + "VALUES ( ";
              //  sql = sql + "'" + User.ToUpper() + "' ";
                sql = sql + "'" + User.ToLower() + "' ";
                sql = sql + "," + Fecha.cambiaFecha(FechaAct, TipoDB) + " ";
                sql = sql + ",'" + HoraAct + "' ";
                sql = sql + ",0) ";

                cn.Open();
                OleDbCommand cmd = new OleDbCommand(sql, cn);
                cmd.ExecuteNonQuery();                
            }

            if (cn.State == ConnectionState.Open) cn.Close();
        }

        public static bool passRepetido(string User, string Pass, long Cant, string Base)
        {
            //Defino un objeto tipo DAL
            //DAL MiDAL = new DAL(); 

            string cn = DAL.constr(Base);
            string sql;
            DataSet ds = new DataSet();
            OleDbDataAdapter daSal;

            sql = "SELECT hist_pass_usr.husrpass ";
            sql = sql + "FROM hist_pass_usr ";
            sql = sql + "WHERE upper(hist_pass_usr.iduser) = '" + User.ToUpper() + "' ";
            sql = sql + "ORDER BY hpassfecini DESC, hpasshoraini DESC ";

            daSal = new OleDbDataAdapter(sql, cn);

            try
            {
                daSal.Fill(ds);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            long Ind = 1;
            bool Termino = (Ind > Cant);
            bool Encontro = false;

            while ((Ind <= ds.Tables[0].Rows.Count) && !Termino)
            {
                Encontro = Convert.ToString(ds.Tables[0].Rows[Convert.ToInt16(Ind - 1)].ItemArray[0]) == Pass;
                Ind++;
                Termino = (Ind > Cant) || Encontro;
            }

            return Encontro;
        }

        public static void eliminarHistPass(string User, long Cant, string Base)
        {
            //Defino un objeto tipo DAL
            //DAL MiDAL = new DAL(); 

            OleDbConnection cn = new OleDbConnection();
            cn.ConnectionString = DAL.constr(Base);
            
            DataSet ds = new DataSet();

            OleDbDataAdapter daSal;

            string sql;

            cn.Open();

            sql = "SELECT hist_pass_usr.hpassnro ";
            sql = sql + "FROM hist_pass_usr ";
            sql = sql + "WHERE upper(hist_pass_usr.iduser) = '" + User.ToUpper() + "' ";
            sql = sql + "ORDER BY hpassfecini DESC, hpasshoraini DESC ";

            daSal = new OleDbDataAdapter(sql, cn);

            try
            {
                daSal.Fill(ds);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            long Ind = 1;
            

            while (Ind <= ds.Tables[0].Rows.Count)
            {
                if (Ind > Cant)
                {
                    //Borra las demas
                    sql = "DELETE FROM hist_pass_usr WHERE hpassnro = " + Convert.ToString(ds.Tables[0].Rows[Convert.ToInt16(Ind - 1)].ItemArray[0]) + " ";
                    
                    OleDbCommand cmd = new OleDbCommand(sql, cn);
                    
                    cmd.ExecuteNonQuery();
                }

                Ind++;
            }
            if (cn.State == ConnectionState.Open) cn.Close();
        }

        public static void ingresarPass(string User, string Pass, string Base)
        {
            //Defino un objeto tipo DAL
            //DAL MiDAL = new DAL(); 

            OleDbConnection cn = new OleDbConnection();
            cn.ConnectionString = DAL.constr(Base);

            string FechaAct = DateTime.Today.ToString("dd/MM/yyyy");
            string HoraAct = DateTime.Now.ToString("hh:mm:ss");
            string sql;         
            

            /*jpb: Recupera el iduser con el mismo Ucase*/
            string UserAux = User;
            string cn2 = DAL.constr(Base);            
            DataSet ds = new DataSet();

            sql = "SELECT iduser FROM user_per WHERE upper(iduser) = '" + User.ToUpper() + "'";

            OleDbDataAdapter da = new OleDbDataAdapter(sql, cn2);

            try
            {
                da.Fill(ds);
            }
            catch (Exception ex)
            {
              
            }

            if (ds.Tables[0].Rows.Count > 0)
            {
                if (ds.Tables[0].Rows[0].ItemArray[0] != DBNull.Value)
                {
                    UserAux =  ds.Tables[0].Rows[0].ItemArray[0].ToString();
                }
            }
            /**********/



            try
            {         
                cn.Open();        
                sql= "INSERT INTO hist_pass_usr (iduser, husrpass, hpassfecini, hpasshoraini) ";
               // sql = sql + " VALUES ('" + User.ToUpper() + "','" + Pass + "'," + Fecha.cambiaFecha(FechaAct, DAL.TipoBase(Base)) + " ";               
                sql = sql + " VALUES ('" + UserAux + "','" + Pass + "'," + Fecha.cambiaFecha(FechaAct, DAL.TipoBase(Base)) + " ";               
                sql = sql + " ,'" + HoraAct + "') ";
      
                OleDbCommand cmd = new OleDbCommand(sql, cn);                
                cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (cn.State == ConnectionState.Open) cn.Close();
            }
        }


        public static void CambiarPassUser(string User, string Valor, string Base)
        {
            //Defino un objeto tipo DAL
            //DAL MiDAL = new DAL(); 

            OleDbConnection cn = new OleDbConnection();
            cn.ConnectionString = DAL.constr(Base);
                try
                {
                    DAL.AddLogEvent("Se cambiará el password en RH Pro", EventLogEntryType.Information, 504);
                    cn.Open();
                    string sql = "UPDATE user_per SET usrpasscambiar = " + Valor + " WHERE upper(iduser) = '" + User.ToUpper() + "' ";
                    OleDbCommand cmd = new OleDbCommand(sql, cn);
                    cmd.ExecuteNonQuery();
                    DAL.AddLogEvent("El password en RH Pro se actualizó correctamente", EventLogEntryType.Information, 505);

                }
                catch (Exception ex)
                {
                    DAL.AddLogEvent("Error al cambiar el password en RH Pro: " + ex.Message + "\n\n" + ex.StackTrace, EventLogEntryType.Error, 506);
                    throw ex;
                }
                finally
                {
                    if (cn.State == ConnectionState.Open) cn.Close();
                }

        }

        public static void CambiarPassBase(string User, string PassNew, string PassOld, string Base,  bool SUP)
        {
            //Defino un objeto tipo DAL
            //DAL MiDAL = new DAL(); 

            string InitialCatalog_Actualizada = "";

            OleDbConnection cn = new OleDbConnection();
            //cn.ConnectionString = DAL.constr(Base);
            //28/10/2013 - CDM - A partir de ahora el cambio de clave lo hace USUSUP y no el ESS
            cn.ConnectionString = DAL.constrSUP(Base);

            DAL.AddLogEvent("Se actualizará la clave del usuario '" + User + "'. Detalle de la conexión: " + cn.ConnectionString, EventLogEntryType.Information, 500);

            string TipoDB = DAL.TipoBase(Base.ToString());

            DAL.AddLogEvent("Se procede a actualizarla en el motor.", EventLogEntryType.Information, 501);



            if (TipoDB.ToUpper() != "ORA")
            {
                //SQL-----------------------------------------------------------
                try
                {
                    cn.Open();

                    OleDbCommand cmd = new OleDbCommand();

                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Connection = cn;
                    cmd.CommandText = "sp_password";

                    OleDbParameter pOld = new OleDbParameter("old", OleDbType.VarChar);
                    OleDbParameter pNew = new OleDbParameter("new", OleDbType.VarChar);
                    OleDbParameter pLoginame = new OleDbParameter("loginame", OleDbType.VarChar);

                    pOld.Value = PassOld;
                    pNew.Value = PassNew;
                    pLoginame.Value = User;

                    cmd.Parameters.Add(pOld);
                    cmd.Parameters.Add(pNew);
                    cmd.Parameters.Add(pLoginame);

                    cmd.ExecuteNonQuery();
                    DAL.AddLogEvent("Se cambió el password del usuario " + User + " en MS-SQL", EventLogEntryType.Information, 504);



                }
                catch (Exception ex)
                {
                    DAL.AddLogEvent("Error al cambiar el password: " + ex.Message + "\n\n" + ex.StackTrace, EventLogEntryType.Error, 502);
                    throw ex;
                }
                finally
                {
                    if (cn.State == ConnectionState.Open) cn.Close();
                }
            }
            else
            {
                //ORA-----------------------------------------------------------
                try
                {
                    cn.Open();
                    //string sql = @"ALTER USER " + User.ToUpper() + " IDENTIFIED BY " + Convert.ToChar(34) + PassNew.ToUpper() + Convert.ToChar(34);
                    string sql = @"ALTER USER " + Convert.ToChar(34) + User.ToUpper() + Convert.ToChar(34) + " IDENTIFIED BY " + Convert.ToChar(34) + PassNew.ToUpper() + Convert.ToChar(34);
                    OleDbCommand cmd = new OleDbCommand(sql, cn);
                    cmd.ExecuteNonQuery();
                    DAL.AddLogEvent("Se cambió el password del usuario " + User + " en Oracle", EventLogEntryType.Information, 505);

                }
                catch (Exception ex)
                {
                    DAL.AddLogEvent("Error al cambiar el password: " + ex.Message + "\n\n" + ex.StackTrace, EventLogEntryType.Error, 503);
                    throw ex;
                }
                finally
                {
                    if (cn.State == ConnectionState.Open) cn.Close();
                }
            }
        }
         
    }
}