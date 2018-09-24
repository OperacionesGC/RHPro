using System;
using System.Data;
using System.Web;
using System.Data.SqlClient;

namespace ConsultaBaseC
{
    public class Password
    {

        
        public static string valorUserPolCuenta(string User, string Campo, string Base)
        {
            string cn = DAL.constr(Base);
            string sql;
            DataSet ds = new DataSet();
            string Salida = "";
            SqlDataAdapter daPass;

            sql = "SELECT usr_pol_cuenta." + Campo;
            sql = sql + " FROM usr_pol_cuenta";
            sql = sql + " WHERE upper(usr_pol_cuenta.iduser) = '" + User.ToUpper() + "'";
            sql = sql + " AND upcfecfin IS NULL";
            daPass = new SqlDataAdapter(sql, cn);
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
            string cn = DAL.constr(Base);
            string sql;
            DataSet ds = new DataSet();
            string Salida = "";
            SqlDataAdapter daPass;

            sql = "SELECT pol_cuenta." + Campo;
            sql = sql + " FROM pol_cuenta";
            sql = sql + " WHERE pol_cuenta.pol_nro = " + PolNro;

            daPass = new SqlDataAdapter(sql, cn);
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

        public static string valorUserPer(string User, string Campo, string Base)
        {
            string cn = DAL.constr(Base);
            string sql;
            DataSet ds = new DataSet();
            string Salida = "";
            SqlDataAdapter daPass;

            sql = "SELECT user_per." + Campo;
            sql = sql + " FROM user_per";
            sql = sql + " WHERE upper(user_per.iduser) = '" + User.ToUpper() + "'";
            daPass = new SqlDataAdapter(sql, cn);
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
            string cn = DAL.constr(Base);
            string sql;
            DataSet ds = new DataSet();
            bool Salida = false;
            SqlDataAdapter daSal;

            sql = "SELECT user_per.iduser";
            sql = sql + " FROM user_per";
            sql = sql + " WHERE upper(user_per.iduser) = '" + User.ToUpper() + "'";
            daSal = new SqlDataAdapter(sql, cn);
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
            string cn = DAL.constr(Base);
            string sql;
            DataSet ds = new DataSet();
            bool Salida = false;
            SqlDataAdapter daSal;

            sql = "SELECT user_per.ctabloqueada";
            sql = sql + " FROM user_per";
            sql = sql + " WHERE upper(user_per.iduser) = '" + User.ToUpper() + "'";
            daSal = new SqlDataAdapter(sql, cn);
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
            string cn = DAL.constr(Base);
            string sql;
            DataSet ds = new DataSet();
            string Salida = "";
            SqlDataAdapter daPass;

            sql = "SELECT hist_pass_usr." + Campo;
            sql = sql + " FROM hist_pass_usr";
            sql = sql + " WHERE upper(hist_pass_usr.iduser) = '" + User.ToUpper() + "'";
            sql = sql + " AND hpassfecfin IS NULL";
            daPass = new SqlDataAdapter(sql, cn);
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
            string cn = DAL.constr(Base);
            string sql;
            DataSet ds = new DataSet();
            long Salida = 0;
            SqlDataAdapter daPass;

            sql = "SELECT hlogintfallidos";
            sql = sql + " FROM hist_log_usr";
            sql = sql + " WHERE upper(hist_log_usr.iduser) = '" + User.ToUpper() + "'";
            daPass = new SqlDataAdapter(sql, cn);
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
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = DAL.constr(Base);

            try
            {
                cn.Open();
                string sql = "UPDATE user_per SET ctabloqueada = " + Valor + " WHERE upper(iduser) = '" + User.ToUpper() + "'";
                SqlCommand cmd = new SqlCommand(sql, cn);
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
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = DAL.constr(Base);

            string FechaAct = DateTime.Today.ToString("dd/MM/yyyy");
            string HoraAct = DateTime.Now.ToString("hh:mm:ss");

            try
            {
                cn.Open();
                string sql = "UPDATE hist_pass_usr SET hpassfecfin = " + Fecha.cambiaFecha(FechaAct, Base);
                sql = sql + " , hpasshorafin = '" + HoraAct + "'";
                sql = sql + " WHERE upper(iduser) = '" + User.ToUpper() + "'";
                sql = sql + " AND hpassfecfin IS NULL";
                SqlCommand cmd = new SqlCommand(sql, cn);
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
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = DAL.constr(Base);

            try
            {
                cn.Open();
                string sql = "UPDATE hist_log_usr SET hlogintfallidos = " + Cant.ToString();
                sql = sql + " WHERE upper(iduser) = '" + User.ToUpper() + "'";
                SqlCommand cmd = new SqlCommand(sql, cn);
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
            string cn = DAL.constr(Base);
            string sql;
            DataSet ds = new DataSet();
            string Salida = "";
            SqlDataAdapter daPass;

            sql = "SELECT hist_log_usr." + Campo;
            sql = sql + " FROM hist_log_usr";
            sql = sql + " WHERE upper(hist_log_usr.iduser) = '" + User.ToUpper() + "'";
            daPass = new SqlDataAdapter(sql, cn);
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


        public static void ingresarLogueo(string User, string Base)
        {
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = DAL.constr(Base);
            DataSet ds = new DataSet();
            SqlDataAdapter daLg;
            string FechaAct = DateTime.Today.ToString("dd/MM/yyyy");
            string HoraAct = DateTime.Now.ToString("hh:mm:ss");

            cn.Open();

            string sql = "SELECT hlognro FROM hist_log_usr";
            sql = sql + " WHERE upper(iduser) = '" + User.ToUpper() + "'";
            daLg = new SqlDataAdapter(sql, cn);
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
                sql = "UPDATE hist_log_usr SET hlogfecini = " + Fecha.cambiaFecha(FechaAct, Base);
                sql = sql + " ,hloghoraini = '" + HoraAct + "'";
                sql = sql + " ,hlogintfallidos = 0";
                sql = sql + " WHERE upper(iduser) = '" + User.ToUpper() + "'";
                SqlCommand cmd = new SqlCommand(sql, cn);
                cmd.ExecuteNonQuery();
            }
            else
            {
                sql = "INSERT INTO hist_log_usr (iduser, hlogfecini, hloghoraini, hlogintfallidos)";
                sql = sql + " VALUES(";
                sql = sql + " '" + User + "'";
                sql = sql + " ," + Fecha.cambiaFecha(FechaAct, Base);
                sql = sql + " ,'" + HoraAct + "'";
                sql = sql + " ,0)";
                SqlCommand cmd = new SqlCommand(sql, cn);
                cmd.ExecuteNonQuery();
            }

            if (cn.State == ConnectionState.Open) cn.Close();

        }


        public static bool passRepetido(string User, string Pass, long Cant, string Base)
        {
            string cn = DAL.constr(Base);
            string sql;
            DataSet ds = new DataSet();
            SqlDataAdapter daSal;

            sql = "SELECT hist_pass_usr.husrpass";
            sql = sql + " FROM hist_pass_usr";
            sql = sql + " WHERE upper(hist_pass_usr.iduser) = '" + User.ToUpper() + "'";
            sql = sql + " ORDER BY hpassfecini DESC, hpasshoraini DESC";
            daSal = new SqlDataAdapter(sql, cn);
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

            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = DAL.constr(Base);
            DataSet ds = new DataSet();
            SqlDataAdapter daSal;
            string sql;

            cn.Open();

            sql = "SELECT hist_pass_usr.hpassnro";
            sql = sql + " FROM hist_pass_usr";
            sql = sql + " WHERE upper(hist_pass_usr.iduser) = '" + User.ToUpper() + "'";
            sql = sql + " ORDER BY hpassfecini DESC, hpasshoraini DESC";
            daSal = new SqlDataAdapter(sql, cn);
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
                    sql = "DELETE FROM hist_pass_usr WHERE hpassnro = " + Convert.ToString(ds.Tables[0].Rows[Convert.ToInt16(Ind - 1)].ItemArray[0]);
                    SqlCommand cmd = new SqlCommand(sql, cn);
                    cmd.ExecuteNonQuery();
                }
                Ind++;
            }
            if (cn.State == ConnectionState.Open) cn.Close();
        }


        public static void ingresarPass(string User, string Pass, string Base)
        {
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = DAL.constr(Base);

            string FechaAct = DateTime.Today.ToString("dd/MM/yyyy");
            string HoraAct = DateTime.Now.ToString("hh:mm:ss");

            try
            {
                cn.Open();
                string sql = "INSERT INTO hist_pass_usr (iduser, husrpass, hpassfecini, hpasshoraini)";
                sql = sql + " VALUES ('" + User + "','" + Pass + "'," + Fecha.cambiaFecha(FechaAct, Base);
                sql = sql + " ,'" + HoraAct + "')";
                SqlCommand cmd = new SqlCommand(sql, cn);
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
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = DAL.constr(Base);

            try
            {
                cn.Open();
                string sql = "UPDATE user_per SET usrpasscambiar = " + Valor + " WHERE upper(iduser) = '" + User.ToUpper() + "'";
                SqlCommand cmd = new SqlCommand(sql, cn);
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

        public static void CambiarPassBase(string User, string PassNew, string PassOld, string Base)
        {
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = DAL.constr(Base);

            try
            {
                cn.Open();

                string sql = "exec sp_password '" + PassOld + "','" + PassNew + "','" + User + "'" ;
                SqlCommand cmd = new SqlCommand(sql, cn);
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


    }
}
