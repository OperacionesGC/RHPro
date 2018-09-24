using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Common;
using System.Configuration;
using System.Data.OleDb;
using System.Data;
using ServicesProxy.MetaHome;

namespace RHPro
{
    public partial class EndSession : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
           
            MetaHome MH = new MetaHome();
            if ((Utils.SessionNroTempLogin != null) && ((String)Utils.SessionNroTempLogin != ""))
            {
               // if (MH.MetaHome_Activo() && MH.MetaHome_RegistraLoguin())
                if (MH.MetaHome_RegistraLoguin())
                {
                   // string EncryptionKey = (String)ConfigurationManager.AppSettings["EncryptionKey"];
                    try
                    {
                     /*   MH_Externo MHome = new MH_Externo();
                        string idTemp = (String)Utils.SessionNroTempLogin;
                        string nroTemp = Encryptor.Decrypt(EncryptionKey, idTemp);
                        MHome.logout_TempLogin(nroTemp); 
                        Utils.SessionNroTempLogin = null;
                      */
                        MH.MetaHome_Logout();
                       
                       
                    }
                    catch (Exception ex) { throw ex; }

                }
                /*
                try
                {
                    OleDbConnection cn = new OleDbConnection();
                    cn.ConnectionString = connStr;
                    cn.Open();
                    DataSet ds = new DataSet();

                    sql = "DELETE Temp_Login WHERE nroTemp = " + nroTemp;
                    OleDbCommand cmd = new OleDbCommand(sql, cn);
                    cmd.ExecuteNonQuery();
                    cn.Close();

                    Utils.SessionNroTempLogin = null;
                }
                catch (Exception exec)
                {
                    throw exec;
                }
                */
                  /*Seteo todas las variables de sesion en null y se las paso a las variables asp*/
                Utils.SetDefaultSessionValues();
                Utils.CopyAspNetSessionToAspSession();
                /**/
               
            }

            RHPro.Lenguaje ObjLenguaje = new RHPro.Lenguaje();
            string LenguajeDefecto = ObjLenguaje.Etiq_Leng_Default(); //(String)ConfigurationManager.AppSettings["Idioma"];
            System.Web.HttpContext.Current.Session["Lenguaje"] = LenguajeDefecto;
            //Response.Write("<script>alert('" + ObjLenguaje.Label_Home("Expiro el tiempo de sesión") + "');</script>");

  
           // Session.Abandon();
            Session.Abandon();
            Session.RemoveAll();

            Session["RHPRO_SesionFinalizada"] = "-1";

             
            
        }
    }
}
