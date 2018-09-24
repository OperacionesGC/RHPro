using System.Configuration;
using System.Data;
using System.Collections;
using System;
using System.Collections.Specialized;
using System.Data.OleDb;
using System.Diagnostics;
using System.Web.Configuration;
using System.Text.RegularExpressions;
using System.Globalization;



namespace ConsultaBaseC
{
    public static class DAL
    {

        public static string Error(int Codigo, string Idioma, String Base)
        {
            string salida = "Error";
            string EtiquetaTraducida = "";
            string EtiquetaDefault = BuscarError(Convert.ToString(Codigo)+ "es-AR");

            if ((EtiquetaDefault != null) && (EtiquetaDefault != ""))
            {
                 EtiquetaTraducida = EtiquetasMI.EtiquetaErr(EtiquetaDefault, Idioma, Base);
                 salida = EtiquetaTraducida;
            }

            return salida;
            
            /*
             if (salida == null){
                //return "Codigo de error " + Codigo + " para el idioma " + Idioma + " no encontrado.";
                return salida;
            }
            else
            {
                return salida2;
            }
             */
        }


        public static string constr(string NroBase)
        {
            //string cnnAux = Encriptar.Decrypt(DAL.EncrKy(), ConfigurationManager.ConnectionStrings[NroBase].ConnectionString);
            string cnnAux = ConfigurationManager.ConnectionStrings[NroBase].ConnectionString;
             
               // cnnAux = cnnAux + " User Id=" + Encriptar.Decrypt(DAL.EncrKy(), UsuESS()) + ";";
              // cnnAux = cnnAux + " Password=" + Encriptar.Decrypt(DAL.EncrKy(), PassESS()) + ";"
            
            cnnAux = cnnAux + " User Id=" + Encriptar.Decrypt(EncrKy(), UsuESS_X(NroBase)) + ";";
            cnnAux = cnnAux + " Password=" + Encriptar.Decrypt(EncrKy(), PassESS_X(NroBase)) + ";";
                        
            if (TipoBase(NroBase).ToUpper() == "ORA")                                    
            { 
                Esquema(cnnAux, NroBase);
            }           
           
            return cnnAux;
        }


        public static string constrSUP(string NroBase)
        {
            //string cnnAux = Encriptar.Decrypt(DAL.EncrKy(), ConfigurationManager.ConnectionStrings[NroBase].ConnectionString);
            string cnnAux = ConfigurationManager.ConnectionStrings[NroBase].ConnectionString;

            // cnnAux = cnnAux + " User Id=" + Encriptar.Decrypt(DAL.EncrKy(), UsuESS()) + ";";
            // cnnAux = cnnAux + " Password=" + Encriptar.Decrypt(DAL.EncrKy(), PassESS()) + ";"

            /*string VirtualPath = System.Web.VirtualPathUtility.GetDirectory(
                    System.Web.HttpContext.Current.Request.Url.AbsolutePath);

            string VirtualPath2 = System.Web.HttpContext.Current.Server.MapPath("");*/

            //Configuration config = WebConfigurationManager.OpenWebConfiguration("/rhprox2/ws");
           
            cnnAux = cnnAux + " User Id=" + Encriptar.Decrypt(EncrKy(), ConfigurationSettings.AppSettings["UsuSUP"]) + ";";
           // cnnAux = cnnAux + " Password=" + Encriptar.Decrypt(DAL.EncrKy(), config.AppSettings.Settings["PassSUP"].Value) + ";";
            cnnAux = cnnAux + " Password=" + Encriptar.Decrypt(DAL.EncrKy(), ConfigurationSettings.AppSettings["PassSUP"]) + ";";

            if (TipoBase(NroBase).ToUpper() == "ORA")
            {
                Esquema(cnnAux, NroBase);
            }

            return cnnAux;
        }

        public static string GetVirtualPath(string url)
        {
            if (System.Web.HttpContext.Current.Request.ApplicationPath == "/")
            {
                return "~" + url;
            }

            return Regex.Replace(url, "^" +
                           System.Web.HttpContext.Current.Request.ApplicationPath + "(.+)$", "~$1");
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
            string[] ArrFila;
            string Salida = "MSSQL";
            string strCon = ConfigurationSettings.AppSettings["Params_"+NroBase];

            if (strCon != null)
            {
                ArrFila = strCon.Split(new char[] { ',' });
                if (ArrFila.Length > 0)
                    Salida = ArrFila[0];                 
            }          
            
            return Salida;

            /*
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
             */ 
        }
        
        //******************************//
        public static string BuscarError(string Clave)
        {
            return ConfigurationSettings.AppSettings[Clave];
        }

        public static string BuscarErrorOLD(string Clave)
        {
            return ConfigurationSettings.AppSettings[Clave];
        }
        //******************************//
        public static string UsuESS()
        {
            return ConfigurationSettings.AppSettings["UsuESS"];
        }

        public static string PassESS()
        {
            return ConfigurationSettings.AppSettings["PassESS"];
        }
         


        public static string EncrKy()
        {
            return ConfigurationSettings.AppSettings["EncrKy"];
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
            if (ConfigurationSettings.AppSettings["TEDESC" + Est.ToString()] == null)
                return "";
            else
                return ConfigurationSettings.AppSettings["TEDESC" + Est.ToString()];
        }

        public static string BuscarEsquema(string NroBase)
        {
            /*  
             <!-- value="TIPOBD,VERSION,ROLE,SCHEMA,TableSpace,TempTableSpace" -->
             <add key="Params_4" value="ORA,11g,RRHHX2,SOSR3,TEMPORARY_DATA,RHPROX2" /> 
             <add key="Params_2" value="MSSQL,2008r2,RRHHX2,TEMPORARY_DATA,RHPROX2" />
             */

            string[] Arr;
            string schema = "";
            if (ConfigurationSettings.AppSettings["Params_" + NroBase] != null)
            {
                string strPar = ConfigurationSettings.AppSettings["Params_" + NroBase];
                Arr = strPar.Split(new char[] { ',' });
                if (Arr.Length > 0)
                    schema = Arr[3];              
            }

            return schema;
        }

        public static string NroEstr(int Est)
        {
            if (ConfigurationSettings.AppSettings["TENRO" + Est.ToString()] == null)
                return "0";
            else
                return ConfigurationSettings.AppSettings["TENRO" + Est.ToString()];
        }



        public static string UsuESS_X(String NroBase)
        {

            if (ConfigurationSettings.AppSettings["UsuESS_" + NroBase] != null)
            {                 
                return ConfigurationSettings.AppSettings["UsuESS_" + NroBase];
            }
            else
            {
                return ConfigurationSettings.AppSettings["UsuESS"];
            }
        }

        public static string PassESS_X(String NroBase)
        {
            if (ConfigurationSettings.AppSettings["PassESS_" + NroBase] != null)
                return ConfigurationSettings.AppSettings["PassESS_" + NroBase];
            else
                return ConfigurationSettings.AppSettings["PassESS"];
        }


        public static void CheckSupPass(string Base)
        {
            try
            {
                AddLogEvent("Inicio verificación de clave de supervisor", EventLogEntryType.Information, 100);
                int interval = int.Parse(ConfigurationSettings.AppSettings["ResetPassSUPInterval"]);
                if (interval < 0)
                {
                    AddLogEvent("El cambio de clave del supervisor esta desactivado. Acción abortada!", EventLogEntryType.Information, 107);
                    return;
                }

                DateTime lastReset;
                try
                {
                    lastReset = DateTime.ParseExact(Encriptar.Decrypt(DAL.EncrKy(), ConfigurationSettings.AppSettings["LastResetPassSUP"]), "dd-MM-yyyy", CultureInfo.InvariantCulture);
                }
                catch
                {
                    lastReset = DateTime.ParseExact("01-01-1900", "dd-MM-yyyy", CultureInfo.InvariantCulture);
                }

                if ((DateTime.Now - lastReset).TotalDays >= interval)
                {
                    AddLogEvent("Se actualizará el cambio de clave de USUSUP, ultimo cambio: " + lastReset.ToString("dd/MM/yyyy"), EventLogEntryType.Information, 101);
                    Configuration config = WebConfigurationManager.OpenWebConfiguration("/rhprox2/ws");
                    string usuSup = Encriptar.Decrypt(DAL.EncrKy(), config.AppSettings.Settings["UsuSUP"].Value);
                    string passSup = Encriptar.Decrypt(DAL.EncrKy(), config.AppSettings.Settings["PassSUP"].Value);
                    string newPassSup = generaClave();
                    AddLogEvent("Nueva clave generada.", EventLogEntryType.Information, 102);

                    Password.CambiarPassBase(usuSup, newPassSup, passSup, Base, true);
                    AddLogEvent("Clave actualizada en el motor. Se procede a guardarla en el archivo de configuración", EventLogEntryType.Information, 103);

                    config.AppSettings.Settings["PassSUP"].Value = Encriptar.Encrypt(DAL.EncrKy(), newPassSup);
                    config.AppSettings.Settings["LastResetPassSUP"].Value = Encriptar.Encrypt(DAL.EncrKy(), DateTime.Now.ToString("dd-MM-yyyy"));
                    config.Save(ConfigurationSaveMode.Modified);

                    AddLogEvent("Clave correctamente guardada en el archivo de configuración", EventLogEntryType.Information, 104);
                }
                else
                {
                    AddLogEvent("No es necesario actualizar la clave de USUSUP, ultimo cambio: " + lastReset.ToString("dd/MM/yyyy"), EventLogEntryType.Information, 105);
                }
            }
            catch (Exception ex)
            {
                AddLogEvent("Error inesperado: " + ex.Message + "\n\n" + ex.StackTrace, EventLogEntryType.Error, 106);

                throw ex;
            }
        }

        private static string generaClave()
        {
            ///Configuración del usuario supervisor (administrador de base y usuarios a nivel de motor)
            ///		UsuSUP: Nombre del usuario administrador (encriptado)		
            ///		PassSUP: Password del usuario (encriptado)
            ///		PassSUPExpr: Expresión regular para la generación de contraseña
            ///				Valores de la expresion regular:
            ///					A: Reemplaza el caracter de la expresión por una letra mayúscula
            ///					a: Reemplaza el caracter de la expresión por una letra minúscula
            ///					#: Reemplaza el caracter de la expresión por un número
            ///					$: Reemplaza el caracter de la expresión por un símbolo ( $ @ - _ )
            ///     			@: Reemplaza el caracter de la expresión por uno al azar
            ///     			!: Si se especifica este caracter al inicio de la expresión se verifica que el resultado contenga al menos una letra mayúscula
            ///     ResetPassSUPInterval: Intervalo en cantidad de dias en el que se debe refrescar el password del administrador (si es menor que 0 no se ejecuta la actualización de password)
            ///     LastResetPassSUP: Fecha de la última actualización de password (encriptado)

            var chExpr = "Aa#$";
            var CHARS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            var chars = "abcdefghijklmnopqrstuvwxyz";
            var numbers = "0123456789";
            var simbols = "$@-_";
            var random = new Random();
            bool usoMay = false; // indica si ya se uso una mayúscula

            string salida = "";
            //Expresión regular de la contraseña
            string regexp = ConfigurationSettings.AppSettings["PassSUPExpr"];

            //Reemplazo los caracteres random
            if (regexp.IndexOf('@') > -1)
            {
                string newRegExp = "";
                foreach (char c in regexp)
                {
                    newRegExp += (c == '@') ? chExpr[random.Next(chExpr.Length)] : c;
                }
                regexp = newRegExp;
            }

            if (regexp.StartsWith("!"))
            {
                regexp = regexp.Substring(1);
                if (regexp.IndexOf('A') < 0)
                {
                    char[] reemplazo = regexp.ToCharArray();
                    reemplazo[random.Next(regexp.Length)] = 'A';
                    regexp = new string(reemplazo);
                }
            }

            foreach (char c in regexp)
            {
                switch (c)
                {
                    case 'A':
                        salida += CHARS[random.Next(CHARS.Length)];
                        usoMay = true;
                        break;
                    case 'a':
                        salida += chars[random.Next(chars.Length)];
                        break;
                    case '#':
                        salida += numbers[random.Next(numbers.Length)];
                        break;
                    case '$':
                        salida += simbols[random.Next(simbols.Length)];
                        break;
                }
            }

            return salida;
        }


        public static void AddLogEvent(string evento, EventLogEntryType tipo, int id)
        {
             
            /*
             * Tipos de notificaciones
             * ReportEvents: None|Erros|All|id (id de los eventos a mostrar separados por comas)
             */
            try
            { 
                string confEvent = ConfigurationSettings.AppSettings["ReportEvents"].ToLower(); ;

                if ((confEvent != "none") && (confEvent != ""))
                {
                    if ((confEvent == "all") || ((confEvent == "error") && (tipo == EventLogEntryType.Error)) ||
                        (("," + confEvent + ",").IndexOf("," + id.ToString() + ",") >= 0))
                    {
                        string sSource;
                        string sLog;
                        string sEvent;

                        sSource = "RH Pro X2";
                        sLog = "Application";

                        if (!EventLog.SourceExists(sSource))
                            EventLog.CreateEventSource(sSource, sLog);
                        
                        EventLog.WriteEntry(sSource, evento, tipo, id);
                    }
                }
            }
            catch (Exception ex)
            {
                //throw ex;
            }
        }

    }
}