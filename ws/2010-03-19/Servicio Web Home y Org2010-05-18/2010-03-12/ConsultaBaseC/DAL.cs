using System.Configuration;
using System.Data;
using System.Collections;
using System;
using System.Collections.Specialized;


namespace ConsultaBaseC
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
        return cnnAux;
    }


    public static string constrUsu(string User, string Pass, string segNT, string NroBase)
    {
        //string cnnAux = Encriptar.Decrypt(DAL.EncrKy(), ConfigurationManager.ConnectionStrings[NroBase].ConnectionString);
        string cnnAux = ConfigurationManager.ConnectionStrings[NroBase].ConnectionString;
        if (segNT == "-1")
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

    public static string BuscarError(string Clave)
    {
        return ConfigurationSettings.AppSettings[Clave];
    }

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

    public static string NroEstr(int Est)
    {
        if (ConfigurationSettings.AppSettings["TENRO" + Est.ToString()] == null)
            return "0";
        else
            return ConfigurationSettings.AppSettings["TENRO" + Est.ToString()];
    }

  }
}