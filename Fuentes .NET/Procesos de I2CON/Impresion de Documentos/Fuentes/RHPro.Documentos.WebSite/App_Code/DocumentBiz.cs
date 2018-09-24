using System;
using System.Collections.Generic;
using System.Data;
using System.Web;
using System.Text;
using System.IO;
using i2Con.Web.UI;
using i2Con.Data.Connection;
using RHPro.ReportesAFD.Clases;

public class DocumentBiz
{
    private int     _id;
    private DataRow _rowData;
    private string  _errorMessage = "";

    public DocumentBiz(int id)
    {
        string   documentUrl = "";
        string[] filesInOutputDirectory;
        int      index;
        FileInfo fileInfo;

        _id = id;
        _rowData = DocumentData.Document(id);
        if (_rowData == null)
        {
            _errorMessage = "Error en la configuración del Empleado.";
        }
        else
        {
            documentUrl            = HttpContext.Current.Server.MapPath("~\\" + I2AppSettings.KeyValue("OutputDocuments"));
            filesInOutputDirectory = Directory.GetFiles(documentUrl, "*.xml");
            for (index = 0; index < filesInOutputDirectory.Length; index++)
            {
                fileInfo = new FileInfo(filesInOutputDirectory[index]);
                if (fileInfo.LastWriteTime.Date <= DateTime.Now.AddDays(-1).Date)
                {
                    try
                    {
                        fileInfo.Delete();
                    }
                    catch
                    {
                    }
                }
            }
        }
    }
    public string CreateDocuments()
    {
        int       index;
        int       columnIndex;
        int       initIndex;
        int       endIndex;
        string    sourceFileName;
        string    documentUrl = "";
        string    documentName;
        string    sourceText;
        string    targetTextBody = "";
        DataTable objData = DocumentData.Persons(_id);

        if (objData == null)
        {
            _errorMessage = "Error en la configuración de los Documentos.";
            return "";
        }

        sourceFileName = DocumentData.TemplateDocumentsDirectory();
        if (sourceFileName == "")
        {
            _errorMessage = "No se especificó el directorio donde recuperar los templates.";
            return "";
        }
        sourceFileName += "\\";
        sourceFileName += _rowData["xml"].ToString();
        documentUrl += "~\\";
        documentUrl += I2AppSettings.KeyValue("OutputDocuments");
        documentUrl += "\\";
        documentUrl += Guid.NewGuid().ToString();
        documentUrl += (sourceFileName.LastIndexOf(".") >= 0
                          ? sourceFileName.Substring(sourceFileName.LastIndexOf("."))
                          : ""
                       );
        documentName = HttpContext.Current.Server.MapPath(documentUrl);
        // ABRO EL DOCUMENTO BASE Y CREO EL DOCUMENTO DE SALIDA.
        StreamReader sourceReader = new StreamReader(sourceFileName);
        StreamWriter targetReader = new StreamWriter(documentName);



        //FGZ - le agregué el try - catch
        // _errorMessage = "Ruta " + sourceFileName + documentName;
        // return "";
        // try
        // {
        //    sourceText = sourceReader.ReadToEnd();
        // }
        // catch
        // {
        //    _errorMessage = "Ruta incorrecta." + sourceFileName ;
        //    return "";
        // }
        //FGZ - le agregué el try - catch

        // LEO EL CONTENIDO TOTAL DEL DOCUMENTO BASE. SEPARO EL HEADER, BODY Y FOOTER.
        sourceText = sourceReader.ReadToEnd();
        initIndex  = sourceText.IndexOf("<w:body>");
        endIndex   = sourceText.IndexOf("</w:body>");

        int i, index2, saltoPagina;
        string block1 = "", block2 = "", block3 = "", block4 = "";
        string sqlquery2;
        string sqlqueryIteracion;
        DataTable objData2 = new DataTable();
        DataTable objData3 = new DataTable();      
        
        for (index = 0; index < objData.Rows.Count; index++)
        {
            targetTextBody += sourceText.Substring(initIndex, endIndex - initIndex + 9);
            for (columnIndex = 0; columnIndex < objData.Columns.Count; columnIndex++) {
                targetTextBody = targetTextBody.Replace(objData.Columns[columnIndex].ColumnName, objData.Rows[index][columnIndex].ToString());
            }
            //---------------------------------------------------------------------------------------------------------------------
            sqlquery2 = "SELECT secquery, tag FROM imp_docu_seccion where id = " + _id;
            //objData2 = Execute.ExecuteSQLDataset(sqlquery2).Tables[0];
            objData2 = I2Database.CreateDataSet(AppSession.RHProDBConnection, sqlquery2).Tables[0];

            saltoPagina = 0;

            for (i = 0; i < objData2.Rows.Count; i++)
            {                
                   
                if (targetTextBody.IndexOf(objData2.Rows[i][1].ToString()) > -1)
                {
                    saltoPagina = -1;
                    block1 = targetTextBody.Substring(1, targetTextBody.IndexOf(objData2.Rows[i][1].ToString()) - 1);

                    block2 = targetTextBody.Substring(targetTextBody.IndexOf(objData2.Rows[i][1].ToString()), (targetTextBody.LastIndexOf(objData2.Rows[i][1].ToString()) + objData2.Rows[i][1].ToString().Length) - targetTextBody.IndexOf(objData2.Rows[i][1].ToString()));
                    block3 = "";
                    block4 = targetTextBody.Substring(targetTextBody.LastIndexOf(objData2.Rows[i][1].ToString()) + objData2.Rows[i][1].ToString().Length, (targetTextBody.Length) - (targetTextBody.LastIndexOf(objData2.Rows[i][1].ToString()) + objData2.Rows[i][1].ToString().Length));

                    sqlqueryIteracion = objData2.Rows[i][0].ToString();
                    if (sqlqueryIteracion.IndexOf("WHERE") > 0)
                    {
                        if (sqlqueryIteracion.IndexOf("ORDER") > 0)
                        {
                            //sqlqueryIteracion = sqlqueryIteracion.Substring(0, sqlqueryIteracion.IndexOf("WHERE") + 5) + " empleado.ternro = " + objData.Rows[index][0] + " " + sqlqueryIteracion.Substring(1, sqlqueryIteracion.IndexOf("WHERE") + 5);
                            sqlqueryIteracion = sqlqueryIteracion.Substring(0, sqlqueryIteracion.IndexOf("WHERE") + 5) + " empleado.ternro = " + objData.Rows[index][0] + " AND " + sqlqueryIteracion.Substring(sqlqueryIteracion.IndexOf("WHERE") + 5, sqlqueryIteracion.Length - (sqlqueryIteracion.IndexOf("WHERE") + 5));
                        }
                        else
                        {
                            //sqlqueryIteracion = sqlqueryIteracion.Substring(0, sqlqueryIteracion.IndexOf("WHERE") + 5) + " empleado.ternro = " + objData.Rows[index][0];
                            sqlqueryIteracion = sqlqueryIteracion + " AND empleado.ternro = " + objData.Rows[index][0];
                        }
                    }
                    else
                    {
                        if (sqlqueryIteracion.IndexOf("ORDER") > 0)
                        {
                            sqlqueryIteracion = sqlqueryIteracion.Substring(0, sqlqueryIteracion.IndexOf("ORDER") - 1) + " WHERE empleado.ternro = " + objData.Rows[index][0] + " " + sqlqueryIteracion.Substring(sqlqueryIteracion.IndexOf("ORDER"), sqlqueryIteracion.Length - sqlqueryIteracion.IndexOf("ORDER"));
                        }
                        else
                        {
                            sqlqueryIteracion = sqlqueryIteracion + " WHERE empleado.ternro = " + objData.Rows[index][0];
                        }
                    }

                    objData3 = Execute.ExecuteSQLDataset(sqlqueryIteracion).Tables[0];
                    for (index2 = 0; index2 < objData3.Rows.Count; index2++)
                    {
                        if (block3 == "")
                        {
                            block3 = block2;
                        }
                        else
                        {
                            block3 += block2;
                        }
                        if (objData3.Rows.Count > 0)
                        {
                            for (columnIndex = 0; columnIndex < objData3.Columns.Count; columnIndex++)
                            {
                                block3 = block3.Replace(objData3.Columns[columnIndex].ColumnName, objData3.Rows[index2][columnIndex].ToString());
                            }
                        }
                        else
                        {
                            //Si la consulta de las iteraciones da vacia, borro los tags
                            for (columnIndex = 0; columnIndex < objData3.Columns.Count; columnIndex++)
                            {
                                block3 = block3.Replace(objData3.Columns[columnIndex].ColumnName, "");
                            }

                        }
                    }

                    //quito los tags
                    //if (index2 > 0)
                    //{
                        block3 = block3.Replace(objData2.Rows[i][1].ToString(), "");
                        targetTextBody = targetTextBody.Replace(block2, block3);
                        //if (index < objData.Rows.Count - 1) {
                        //    targetTextBody += "<w:p><w:r><w:br w:type='page' /></w:r></w:p>";
                        //}
                    //}
                }
            }
            //---------------------------------------------------------------------------------------------------------------------
            //if (saltoPagina == -1)
            //{
                if (index < objData.Rows.Count - 1)
                {
                    targetTextBody += "<w:p><w:r><w:br w:type='page' /></w:r></w:p>";
                }

            //}
            
        }
        
        // GUARDO EL TEXTO CON LOS REEMPLAZOS EN LA SALIDA, AGREGÁNDOLE EL HEADER Y FOOTER DEL DOCUMENTO ORIGINAL.
        targetTextBody = sourceText.Substring(0, initIndex) + targetTextBody  + sourceText.Substring(endIndex + 9);
        targetReader.Write(targetTextBody);
        // CIERRO EL ARCHIVO DE ENTRADA Y SALIDA.
        sourceReader.Close();
        targetReader.Close();
        return documentUrl;
    }
    public string ErrorMessage { get { return _errorMessage.Trim(); } }
}
