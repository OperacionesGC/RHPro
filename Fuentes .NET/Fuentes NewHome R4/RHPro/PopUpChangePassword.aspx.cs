using System;
using System.Threading;
using System.Web.UI;
using Entities;
using ServicesProxy;
using ServicesProxy.rhdesa;
using System.Text;
using System.Web;
using System.IO;
using System.Net;
using Common;
using System.Configuration;

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
        
        public RHPro.Lenguaje ObjLenguaje;

        protected void Page_Load(object sender, EventArgs e)
        {
            ObjLenguaje = new RHPro.Lenguaje();

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
                        title.Text = ObjLenguaje.Label_Home("Cambiar Contraseña");// GetLocalResourceObject("titulo.Text").ToString();
                    }
                    else
                    {
                        title.Text = ObjLenguaje.Label_Home("Debe Cambiar su contraseña");//PopUpChangePassData.Login.Messege;
                    }
                }

                txtOldPassword.Focus();

              
            }
        }

        #endregion

        #region Controls Handles


        protected void btnConfirmar_Click(object sender, EventArgs e)
        {
            string IndicadorErr = "ERROR_METAHOME";
            string SeparadorErr = "DET_ERR_META"; 
            
            String ResultadoErr = "";
            //Si la modalidad SaaS esta activa, debe replicar el cambio en todas las bases en las que el usuario este asociado 
            MetaHome MH = new MetaHome();

            if (MH.MetaHome_Activo())//Si esta activo el modo SaaS actualiza el password externamente desde ws_ext del metahome
            {
                Consultas cc = new Consultas();
                string Catalog = cc.Initial_Catalog(PopUpChangePassData.DataBase.Id);
                string DataSource = cc.Data_Source(PopUpChangePassData.DataBase.Id);

                string URL = HttpContext.Current.Request.Url.AbsoluteUri.Trim().ToLower().Replace("popupchangepassword", "default");
                string keyConnStr = PopUpChangePassData.DataBase.Id; 

                //Conecto con el ws externo para actualizar en todas las bases a la que pertenezca el usuario. Si hay error viene un string!=""
                string ResultadoMH = Actualizar_Multiples_Bases(PopUpChangePassData.UserName, txtNewPassword.Value, Catalog, DataSource, keyConnStr, URL);
                                
                if (ResultadoMH.Contains(IndicadorErr))
                {
                    
                    //Page.ClientScript.RegisterStartupScript(GetType(), "Err", String.Format("javascript:alert('{0}');", ObjLenguaje.Label_Home("Error") + ": " + ObjLenguaje.Label_Home("Consulte con el administrador")), true);
                    
                    //Divide los errores segun el separador.
                    string[] ArrErrores = System.Text.RegularExpressions.Regex.Split(ResultadoMH, IndicadorErr);
                    string[] detalleErrores = System.Text.RegularExpressions.Regex.Split(ArrErrores[1], SeparadorErr);
                    //Arma la salida con todos los errores deducidos.
                    foreach (String E in detalleErrores)
                    {                      
                        ResultadoErr += "\\n " + E;
                    }

                    Page.ClientScript.RegisterStartupScript(GetType(), "Err", String.Format("javascript:alert('{0}');",  ResultadoErr), true);
                }
                else
                {
                    errorMess.Text = "";
                    errorMess.Visible = false;
                    Session["ViendeDeCambiarPassword"] = "SI";
                    Page.ClientScript.RegisterStartupScript(GetType(), "Mensaje", String.Format("javascript:alert('{0}');", GetLocalResourceObject("Relog.Text")), true);
                }
            }
            else//Si no esta activo el modo SaaS actualiza el password localmente
            {
                string changePassword = ChangePasswordServiceProxy.ChangePassword(PopUpChangePassData.UserName, txtOldPassword.Value,
                                                 txtNewPassword.Value, txtVerifyPassword.Value,
                                                 PopUpChangePassData.DataBase.Id, Utils.Lenguaje);

                if (changePassword != "")
                {
                    errorMess.Text = changePassword;
                    errorMess.Visible = true;
                    txtOldPassword.Focus();
                }
                else
                {
                    errorMess.Text = "";
                    errorMess.Visible = false;
                    Session["ViendeDeCambiarPassword"] = "SI";
                }

                Page.ClientScript.RegisterStartupScript(GetType(), "Mensaje", String.Format("javascript:alert('{0}');", GetLocalResourceObject("Relog.Text")), true);
            }

            ClientScript.RegisterStartupScript(GetType(), "CerrarPopup", "javascript:window.opener.location='Default.aspx';window.close();", true);
            
        }

        
        /// <summary>
        /// Metodo que actualiza en password de un usuario en todas las bases a las que pertenece 
        /// </summary>
        /// <param name="usuario"></param>
        /// <param name="password"></param>
        /// <param name="Catalog"></param>
        /// <param name="DataSource"></param>
        /// <returns></returns>
 
        //public bool Actualizar_Multiples_Bases(string usuario, string password, string Catalog, string DataSource)
        public string Actualizar_Multiples_Bases(string usuario, string password, string Catalog, string DataSource
            , String keyConnStr, String URL)
        {
            string EncryptionKey = (String)ConfigurationManager.AppSettings["EncryptionKey"];
            string PassEncript = Encryptor.Encrypt(EncryptionKey, password);

            MetaHome MH = new MetaHome();
            MH.Iniciar_Ws_Ext();

            return MH.MetaHome_ActualizaMultiplesBases(usuario, PassEncript, Catalog, DataSource,  keyConnStr, URL);

        }
    
       

        //protected void btnConfirmar_Click(object sender, EventArgs e)
        //{             

        //    //string changePassword = ChangePasswordServiceProxy.ChangePassword(PopUpChangePassData.UserName, txtOldPassword.Value,
        //    //                                          txtNewPassword.Value, txtVerifyPassword.Value,
        //    //                                          PopUpChangePassData.DataBase.Id, Thread.CurrentThread.CurrentCulture.Name);

        //    string changePassword = ChangePasswordServiceProxy.ChangePassword(PopUpChangePassData.UserName, txtOldPassword.Value,
        //                                             txtNewPassword.Value, txtVerifyPassword.Value,
        //                                             PopUpChangePassData.DataBase.Id, Utils.Lenguaje);

        //    if (changePassword!="")
        //    {
        //        errorMess.Text = changePassword;
        //        errorMess.Visible = true;
        //        txtOldPassword.Focus();
        //    }
        //    else
        //    {
        //        errorMess.Text = "";
        //        errorMess.Visible = false;
        //        Session["ViendeDeCambiarPassword"] = "SI";
        //        //Si la modalidad SaaS esta activa, debe replicar el cambio en todas las bases en las que el usuario este asociado 
        //        MetaHome MH = new MetaHome();
        //        if (MH.MetaHome_Activo())
        //        {                    
        //            Consultas cc = new Consultas();                    
        //            string Catalog = cc.Initial_Catalog(PopUpChangePassData.DataBase.Id);
        //            string DataSource = cc.Data_Source(PopUpChangePassData.DataBase.Id);
        //            if (!Actualizar_Multiples_Bases(PopUpChangePassData.UserName, txtNewPassword.Value, Catalog, DataSource))
        //            {
        //                //Page.ClientScript.RegisterStartupScript(GetType(), "Err", "javascript:alert('" + ObjLenguaje.Label_Home("Error") + ":" + ObjLenguaje.Label_Home("Consulte con el administrador") + "') ", true);
        //                Page.ClientScript.RegisterStartupScript(GetType(), "Err", String.Format("javascript:alert('{0}');", ObjLenguaje.Label_Home("Error") + ":" + ObjLenguaje.Label_Home("Consulte con el administrador")), true);
        //            }
        //            else Page.ClientScript.RegisterStartupScript(GetType(), "Mensaje", String.Format("javascript:alert('{0}');", GetLocalResourceObject("Relog.Text")), true);
        //        }               
        //        else 
        //            Page.ClientScript.RegisterStartupScript(GetType(), "Mensaje", String.Format("javascript:alert('{0}');",GetLocalResourceObject("Relog.Text")), true);
                
        //        ClientScript.RegisterStartupScript(GetType(), "CerrarPopup", "javascript:window.opener.location='Default.aspx';window.close();", true);
        //        //ClientScript.RegisterStartupScript(GetType(), "CerrarPopup", "javascript:window.opener.location='Default.aspx';", true);
                
        //    }
        //}


       
    
 



        #endregion
    }
}