using System.Configuration;
using System.Data;
using System.Collections;
using System.Collections.Specialized;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Web;
using System.Web.Configuration;
using System.IO;
using System;
using System.Linq;
using System.Xml;

namespace RHProX2
{
    public class DAL
    {
        public static string Error(int Codigo, string Idioma)
        {
            string salida = BuscarError(Convert.ToString(Codigo)+Idioma);

            if (salida == null)
            {
                return "Codigo de error " + Codigo + " para el idioma " + Idioma + " no encontrado.";
            }
            else
            {
                return salida;
            }
        }
      
        public static string constr(string NroBase)
        {
            //string cnnAux = Encriptar.Decrypt(DAL.EncrKy(), ConfigurationManager.ConnectionStrings[NroBase].ConnectionString);
            string cnnAux = ConfigurationManager.ConnectionStrings[NroBase].ConnectionString;
            cnnAux = cnnAux + " User Id=" + Encriptar.Decrypt(DAL.EncrKy(),UsuESS()) + ";";
            cnnAux = cnnAux + " Password=" + Encriptar.Decrypt(DAL.EncrKy(),PassESS()) + ";";

            if (TipoBase(NroBase).ToUpper()=="ORA")
            {
                Esquema(cnnAux, NroBase);
            }

            return cnnAux;

        }

        public static void Esquema(string conexion, string NroBase)
        {
            OleDbConnection cn2 = new OleDbConnection();
            cn2.ConnectionString = conexion;
            cn2.Open();
            string sqlSS = "ALTER SESSION SET NLS_SORT = BINARY";
            OleDbCommand cmd = new OleDbCommand(sqlSS, cn2);
            cmd.ExecuteNonQuery();
            sqlSS = "ALTER SESSION SET CURRENT_SCHEMA = " + BuscarEsquema(NroBase);
            cmd = new OleDbCommand(sqlSS, cn2);
            cmd.ExecuteNonQuery();
            if (cn2.State == ConnectionState.Open) cn2.Close();
        }

        public static string constrUsu(string User, string Pass, string segNT, string NroBase)
        {
            string cnnAux = ConfigurationManager.ConnectionStrings[NroBase].ConnectionString;
            
            if (segNT == "TrueValue")
            {
                cnnAux = cnnAux + " Integrated Security=SSPI;";
            }
            else
            {
                cnnAux = cnnAux + " User Id=" + User + ";";
                cnnAux = cnnAux + " Password=" + Pass + ";";
            }

            return cnnAux;
        }


        ///<summary>
        ///Lee la configuración de la aplicación desde el web.config.
        ///</summary>
        ///<return>
        ///Devuelveel string de conexion desde el web.config, concatenandole el isr y pass.
        ///</return>
        ///<param name="User">Usuario.</param>
        ///<param name="Pass">Contraseña</param>
        ///<param name="segNT">Tipo de seguridad(1=Integrada)</param>
        ///<param name="NroBase">Nombre del connectionString del web.config</param>
        ///
        public static string constrWsUsu(string User, string Pass, string segNT, string NroBase){
            
            /* get the connectionstring from web.config by [base][UserName][Password] */
            /*
            string dir = HttpContext.Current.Request.ApplicationPath;
            dir = "/ws/";
            string physicalPath = HttpContext.Current.Request.MapPath(dir);
            string smp = Path.GetDirectoryName(physicalPath);
            //VirtualDirectoryMapping vdm = new VirtualDirectoryMapping(@"ws",false);
            VirtualDirectoryMapping vdm = new VirtualDirectoryMapping(smp, false);
            WebConfigurationFileMap wcfm = new WebConfigurationFileMap();
            wcfm.VirtualDirectories.Add("/", vdm);
            // Get the connectionString
            Configuration config = WebConfigurationManager.OpenMappedWebConfiguration(wcfm, "/ws/");
            string connection = config.ConnectionStrings.ConnectionStrings[NroBase].ToString();
            string cnnAux = connection;// ConfigurationManager.ConnectionStrings[NroBase].ConnectionString;
            */
            /*
            if (segNT == "1"){
                cnnAux = cnnAux + " Integrated Security=SSPI;";
            }else{
                cnnAux = cnnAux + " User Id=" + User + ";";
                cnnAux = cnnAux + " Password=" + Pass + ";";
            }
            */
            
            System.Configuration.Configuration configws =WebConfigurationManager.OpenWebConfiguration("/ws") as System.Configuration.Configuration; 
            //Configuration configws = WebConfigurationManager.OpenWebConfiguration("~");

            // Get the connectionStrings section.
            ConnectionStringsSection section = configws.GetSection("connectionStrings") as ConnectionStringsSection;
            string cnnAux = section.ConnectionStrings[NroBase].ToString();
            cnnAux = constrGenerate(cnnAux, User, Pass, segNT); 
            //cnnAux = section.ConnectionStrings[NroBase].ToString();
            return cnnAux;
        }

        ///<summary>
        ///Genera un string de conexion para sql.net en base al strig del web.config (conn_db)
        ///</summary>
        ///<return>
        ///Devuelve el string de conexion concatenandole el usr y pass.
        ///</return>
        ///<param name="User">Usuario.</param>
        ///<param name="Pass">Contraseña</param>
        ///<param name="segNT">Tipo de seguridad(1=Integrada)</param>
        ///<param name="CnnString">String de conexion del connectionString del web.config (conn_db)</param>
        public static string constrGenerate(string CnnString, string User, string Pass, string segNT)
        {
            string[] arrCn = CnnString.Split(';');
            string strcn = "";

            for (int a=0; a<arrCn.Count()-1; a++){
                string[] arrcn2 = arrCn[a].Split('=');
                if (arrcn2[0].ToString().ToUpper() != "Provider".ToUpper()){
                    strcn += arrcn2[0].ToString() + "=" + arrcn2[1].ToString() + ";";
                }
            }
            
            if (segNT == "1"){
                strcn += " Integrated Security=SSPI;";
            }else{
                strcn += "User id=" + User + ";Password=" + Pass;
            }

            return strcn;
        }


        public static DataTable Bases() 
        {
            //Creo la tabla de salida
            DataTable tablaSalida = new DataTable("table");
            DataColumn Columna = new DataColumn();
            Columna.DataType = System.Type.GetType("System.String");
            Columna.ColumnName = "combo";
            Columna.AutoIncrement = false;
            Columna.Unique = false;
            tablaSalida.Columns.Add(Columna);

            NameValueCollection appSettings = ConfigurationManager.AppSettings;
            IEnumerator appSettingsEnum = appSettings.Keys.GetEnumerator();

            int i = 0;

            while (appSettingsEnum.MoveNext())
            {
                string key = appSettings.Keys[i];
                if (isNumeric(key))
                {
                    DataRow fila = tablaSalida.NewRow();
                    fila["combo"] = appSettings[key].ToString();
                    tablaSalida.Rows.Add(fila);
                }
                i += 1;
            }

            return tablaSalida;
        }

        public static string TipoBase(string NroBase)
        {
            string Salida = "SQL";

            NameValueCollection appSettings = ConfigurationManager.AppSettings;
            IEnumerator appSettingsEnum = appSettings.Keys.GetEnumerator();

            bool Ciclar = true;
            bool Encontro = false;
            int i = 0;
            string Fila;
            string[] ArrFila;

            while (Ciclar)
            {
                if (isNumeric(appSettings.Keys[i]))
                {
                    Fila = appSettings[appSettings.Keys[i]].ToString();
                    ArrFila = Fila.Split(new char[] { ',' });
                    if (ArrFila[1] == NroBase)
                    {
                        Encontro = true;
                        if (ArrFila.Length >= 5)
                            Salida = ArrFila[4];
                    }
                }
                else
                    Ciclar = false;

                i += 1;

                if ((i > appSettings.Count) || (Encontro))
                    Ciclar = false;
            }

            return Salida;
        }

        public static string BuscarError(string Clave)
        {
            return System.Configuration.ConfigurationManager.AppSettings[Clave];
        }

        public static string UsuESS()
        {
            return System.Configuration.ConfigurationManager.AppSettings["UsuESS"];
        }

        public static string PassESS()
        {
            return System.Configuration.ConfigurationManager.AppSettings["PassESS"];
        }

        public static string EncrKy()
        {
            return System.Configuration.ConfigurationManager.AppSettings["EncrKy"];
        }

        public static bool isNumeric(object value)
        {
            try
            {
                double d = System.Double.Parse(value.ToString(), System.Globalization.NumberStyles.Any);
                return true;
            }
            catch (System.FormatException)
            {
                return false;
            }
        }

        
        public static string DescEstr(int Est)
        {
            if (System.Configuration.ConfigurationManager.AppSettings["TEDESC" + Est.ToString()] == null)
                return "";
            else
                return System.Configuration.ConfigurationManager.AppSettings["TEDESC" + Est.ToString()];
        }
        
        public static string BuscarEsquema(string NroBase)
        {
            
            if (System.Configuration.ConfigurationManager.AppSettings["SCHEMA" + NroBase] == null)
                return "";
            else
                return System.Configuration.ConfigurationManager.AppSettings["SCHEMA" + NroBase];
        }
        
        public static string NroEstr(int Est)
        {
            if (System.Configuration.ConfigurationManager.AppSettings["TENRO" + Est.ToString()] == null)
                return "0";
            else
                return System.Configuration.ConfigurationManager.AppSettings["TENRO" + Est.ToString()];
        }
         
    }
}