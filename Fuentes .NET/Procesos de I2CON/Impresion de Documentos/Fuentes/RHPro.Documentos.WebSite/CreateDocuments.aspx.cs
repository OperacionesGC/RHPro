using System;
using System.Data;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using RHPro.ReportesAFD.Clases;
using i2Con.Web.UI;

public partial class CreateDocuments : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {

        I2MessageTran.Visible = true;

        try
        {
            int id;
            string pageName;
            string idConexion;
            
            if ( Request.QueryString["Id"] != null && Request.QueryString["ix"] != null )
            {

                ///Recupero Id de la base en la que fue ejecutada.
                idConexion = Request.QueryString["ix"].ToString(); 
                AppSession.Base(idConexion);

                if (int.TryParse(Request.QueryString["Id"].ToString(), out id))
                {

                    DocumentBiz documentBiz = new DocumentBiz(id);
                    if (documentBiz.ErrorMessage != "")
                        ShowMessage(documentBiz.ErrorMessage, I2Message.I2MessageType.Warning);
                    else
                    {
                        pageName = documentBiz.CreateDocuments();
                        if (documentBiz.ErrorMessage != "")
                            ShowMessage(documentBiz.ErrorMessage, I2Message.I2MessageType.Warning);
                        else
                            //ShowMessage(pageName, I2Message.I2MessageType.Warning);
                            Response.Redirect(pageName, false);
                    }
                }
                else
                    ShowMessage("Formato del parámetro [" + Request.QueryString["Id"].ToString() + "] incorrecto.", I2Message.I2MessageType.Warning);
                                
            }
        }
        catch (Exception ex)
        {
            ShowMessage(ex.Message, I2Message.I2MessageType.Error);
        }
    }

    

    

    private void ShowMessage(string errorMessage, I2Message.I2MessageType messageType)
    {
        string mensaje1;
        string mensaje2;
        mensaje1 = "Se ha producido un error durante la creación del documento solicitado.";
        mensaje2 = "Por favor, pongase en contacto con el Administrador del Sistema.";
        I2MessageTran.Show(messageType, mensaje1 + "<br/>" + errorMessage + "<br/>" + mensaje2);
    }
    
}

