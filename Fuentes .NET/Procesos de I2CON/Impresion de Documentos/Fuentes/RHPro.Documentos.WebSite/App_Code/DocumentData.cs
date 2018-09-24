using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using i2Con.Data.Connection;
using i2Con.Web.UI;
using RHPro.ReportesAFD.Clases;

public class DocumentData
{
    
    #region
    private static string _sql;
    #endregion

    
    public static DataTable Persons(int id)
    {
        try
        {
    
            //return Execute.ExecuteSQLDataset(Document(id)["query"].ToString() + Document(id)["query2"].ToString() + Document(id)["query3"].ToString()).Tables[0];
            _sql = Document(id)["query"].ToString() + Document(id)["query2"].ToString() + Document(id)["query3"].ToString();
            return I2Database.CreateDataSet(AppSession.RHProDBConnection, _sql).Tables[0];
        }
        catch
        {
            return null;
        }
    }
    public static DataRow Document(int id)
    {
        DataTable objData = new DataTable();
        string sqlCommand = "";
        sqlCommand += " select ";
        sqlCommand += " id, ";
        sqlCommand += " xml, ";
        sqlCommand += " query, ";
        sqlCommand += " query2, ";
        sqlCommand += " query3, ";
        sqlCommand += " individual, ";
        sqlCommand += " iduser ";
        sqlCommand += " from " + I2AppSettings.KeyValue("TableName");
        sqlCommand += " where id = " + id.ToString();
        //objData = Execute.ExecuteSQLDataset(sqlCommand).Tables[0];
        objData = I2Database.CreateDataSet(AppSession.RHProDBConnection, sqlCommand).Tables[0];
        
        if (objData.Rows.Count == 1)
            return objData.Rows[0];
        else
            return null;
    }
    public static string TemplateDocumentsDirectory()
    {
        DataTable objData = new DataTable();
        string sqlCommand = "SELECT sisdirsegleg AS TemplateDocumentsDirectory FROM sistema";
        //objData = Execute.ExecuteSQLDataset(sqlCommand).Tables[0];
        objData = I2Database.CreateDataSet(AppSession.RHProDBConnection, sqlCommand).Tables[0];
        if (objData.Rows.Count > 0)
            return objData.Rows[0][0].ToString();
        else
            return "";
    }
}
