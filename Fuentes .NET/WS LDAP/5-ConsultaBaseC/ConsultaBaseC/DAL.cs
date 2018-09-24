using System.Configuration;
using System.Data;
using System.Collections;
using System;
using System.Collections.Specialized;
using System.Data.OleDb;



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



    }
}