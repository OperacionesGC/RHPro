using System;
using System.Data;
using System.Configuration;
using System.Collections;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using RHPro.ReportesAFD.Clases;
using RHPro.ReportesAFD.BussinesLayer.Biz;
using i2Con.Web.UI;

public partial class CreateReportAFD : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        I2MessageTran.Visible = true;
        try
        {
            int NroReporte;
            int tipoAFD;
            #region Lectura de Parámetros de llamada del proceso
            if (Request.QueryString["NroReporte"] == null) throw new Exception("Falta el parámetro de Nro. de Reporte.");
            try
            {
                NroReporte = int.Parse(Request.QueryString["NroReporte"].ToString());
                tipoAFD = int.Parse(Request.QueryString["tipoAFD"].ToString());
            }
            catch
            {
                throw new Exception("Formato del parámetro incorrecto.");
            }
            #endregion

            string ArchivoEntradaXML = DocumentData.TemplateDocumentsDirectory() + RHProBiz.ObtenerNombrePlantillaXML(NroReporte) + ".xml";
            string ArchivoSalidaDOC = AppDomain.CurrentDomain.BaseDirectory + I2AppSettings.KeyValue("OutputDocuments") + "\\Reporte de Salida.doc";

            #region Búsqueda de los Query + Tags del Reporte
            AfdReporteBiz reporte = new AfdReporteBiz(NroReporte, null, null, null, null);
            if (reporte.Lista.Count == 0)
            {
                throw new Exception("No se han encontrado los parámetros de armado del reporte en la DB. Proceso Cancelado.");
            }
            #endregion

            bool flagPrimerReepmpazo = true;
            int cantReg = 0;
            
            foreach (AfdReporteSBiz r in reporte.Lista)
            {
                #region Búsqueda de los Campos del Tag
                AfdReporteCampoBiz campos = new AfdReporteCampoBiz(r.Nrotag, null, null, null);
                if (campos.Lista.Count != 0)
                {
                    if (tipoAFD == -1) //se ejecuta el reporte completo
                    {
                        cantReg = 0;
                    }

                    if (tipoAFD == 0) //se ejecuta el reporte borrador con la cantidad de registros cargados en la tabla
                    {
                        cantReg = r.Cantreg;
                    }


                    DataTable dt = RHProBiz.ObtenerRegistros(r.Tablequery, cantReg, campos.ObtenerCamposAlias());
                    int cantRegTotal = RHProBiz.ObtenerCantRegistros(r.Tablequery);

                    WordImport.InsertarTabla(r.Nomtag, ArchivoEntradaXML, ArchivoSalidaDOC, dt, cantRegTotal, flagPrimerReepmpazo);
                    flagPrimerReepmpazo = false;
                }
                else
                {
                    throw new Exception("No se han encontrado campos para el Tag: " + r.Nomtag + " en la DB.");
                }
                #endregion
            }
            Response.Redirect(I2AppSettings.KeyValue("OutputDocuments") + "\\Reporte de Salida.doc", false);
        }
        catch (Exception ex)
        {
            ShowMessage(ex.Message, I2Message.I2MessageType.Error);
        }
    }
    private void ShowMessage(string errorMessage, I2Message.I2MessageType messageType)
    {
        I2MessageTran.Show(messageType,
                           "Se ha producido un error durante la creación del documento solicitado.<br />" +
                           errorMessage + "<br />" +
                           "Por favor, pongase en contacto con el Administrador del Sistema.");
    }
}
