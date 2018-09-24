using System;
using System.Threading;
using System.Web.UI;
using Entities;
using ServicesProxy;

namespace RHPro
{
    public partial class PopUp : Page
    {
        #region Properties

        public PopUpChangePassData PopUpChangePassData
        {
            get 
            { 
                return ViewState["PopUpChangePassData"] as PopUpChangePassData; 
            }
            set
            { 
                ViewState["PopUpChangePassData"] = value; 
            }
        }

        #endregion

        #region Page Handles

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {   
                PopUpChangePassData = Session["PopUpChangePassData"] as PopUpChangePassData;
                Session.Remove("PopUpChangePassData");

                if (PopUpChangePassData == null)
                {
                    Page.ClientScript.RegisterStartupScript(GetType(), "Mensaje", string.Format("javascript:alert('{0}');", GetLocalResourceObject("formaIncorrecta.Text").ToString()), true);
                    Page.ClientScript.RegisterStartupScript(GetType(), "cerrarPopUp", "javascript:window.close()", true);
                }
                else
                {
                    if (string.IsNullOrEmpty(PopUpChangePassData.Login.Messege))
                    {
                        title.Text = GetLocalResourceObject("titulo.Text").ToString();
                    }
                    else
                    {
                        title.Text = PopUpChangePassData.Login.Messege;
                    }
                }

                txtOldPassword.Focus();
            }
        }

        #endregion

        #region Controls Handles

        protected void btnConfirmar_Click(object sender, EventArgs e)
        {
            //Llamar al proxy de actualizar password

            string changePassword = ChangePasswordServiceProxy.ChangePassword(PopUpChangePassData.UserName, txtOldPassword.Value,
                                                      txtNewPassword.Value, txtVerifyPassword.Value,
                                                      PopUpChangePassData.DataBase.Id, Thread.CurrentThread.CurrentCulture.Name);

            if (changePassword!="")
            {
                errorMess.Text = changePassword;
                errorMess.Visible = true;
                txtOldPassword.Focus();
            }
            else
            {
                errorMess.Text = "";
                errorMess.Visible = false;

                Page.ClientScript.RegisterStartupScript(GetType(), "Mensaje", String.Format("javascript:alert('{0}');",GetLocalResourceObject("Relog.Text")), true);
                ClientScript.RegisterStartupScript(GetType(), "CerrarPopup", "javascript:window.close();",true);
            }
        }

        #endregion
    }
}