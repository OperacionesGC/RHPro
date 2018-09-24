using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Configuration;
using ServicesProxy.rhdesa;
using Common;
using System.Text.RegularExpressions;
using Entities;

namespace RHPro
{
    public class Conexion : System.Web.UI.Page
    {


        //public void Tiempo_Ejec(String mje)
        //{
        //    try
        //    {
        //        mje = mje.Replace("<br>", " ############ ");
        //        mje = mje.Replace("\\n", " ############ ");
        //        ///* -sacar ---------------------------------*/
        //        //Consultas cc = new Consultas();
        //        //System.Data.DataTable dt = cc.get_DataTable("select cnstring from conexion order by cnnro ASC ", Utils.SessionBaseID);

        //        System.Data.OleDb.OleDbConnection cn3 = new System.Data.OleDb.OleDbConnection();
        //       // cn3.ConnectionString = Convert.ToString(dt.Rows[0]["cnstring"]); // "Provider=SQLOLEDB.1;Password=ess;User ID=ess;Data Source=RHDESA;Initial Catalog=BASE_0_R3_ARG;";
        //        cn3.ConnectionString = "Provider=SQLOLEDB.1;Password=6852593102166269536E;User ID=raetlatam;Persist Security Info=False;Data Source=SD-P-RHPAPP01;Initial Catalog=HEID_AR_BA0_T1";
        //        cn3.Open();
        //        string sqlSS3 = "insert into sacar (sql) values ('" + mje.Replace("'", "##") + "')";
        //        System.Data.OleDb.OleDbCommand cmd = new System.Data.OleDb.OleDbCommand(sqlSS3, cn3);
        //        cmd.ExecuteNonQuery();
        //        ///*-----------------------------------*/

        //        //  Response.Write("<script> alert('" + mje.Replace("'", "") + "'); </script>");

        //    }
        //    catch (Exception e) { }            
        //}
        //--SACAR----------------------------------------------------------------------------------
        //public System.Diagnostics.Stopwatch tiempoprueba = System.Diagnostics.Stopwatch.StartNew();
        //string MjeTiempo = "";
        //--SACAR----------------------------------------------------------------------------------

        
        public void Iniciar_Sesion(Entities.Login login, String SelectedDatabaseId, String User, String Pass, String guid, string NombreBase, string NroBase)
        {

            //--SACAR----------------------------------------------------------------------------------                            
           // MjeTiempo += "<br> Inicio_Session: " + tiempoprueba.Elapsed.Milliseconds.ToString();
            //--SACAR----------------------------------------------------------------------------------

            /**********************************************************/
            /* Se genera el token de seguridad  de hackeo             */
            string EncryptionKey = ConfigurationManager.AppSettings["EncryptionKey"];
            bool EncriptUserData = bool.Parse(ConfigurationManager.AppSettings["EncriptUserData"]);
            Session["AuthToken"] = guid;
            //Response.Cookies.Add(new HttpCookie("AuthToken", guid));
            /**********************************************************/
            
            System.Web.HttpContext.Current.Session["lstIndex"] = NroBase;
            System.Web.HttpContext.Current.Session["NombreBaseSeleccionada"] = NombreBase;

            
            Session["login"] = login;

            String EtiqLenguaje = login.Lenguaje;
            Consultas cc = new Consultas();
            if ((Utils.SessionNroTempLogin == null) || ((String)Utils.SessionNroTempLogin == ""))
            {
                Common.Utils.SessionBaseID = SelectedDatabaseId;

                //jpb -Actualizo el nombre del InitialCatalog de la conexion
                //Consultas cc1 = new Consultas();
                //Paso las credenciales al web service
                cc.Credentials = System.Net.CredentialCache.DefaultCredentials;
                Session["InitialCatalog"] = cc.Initial_Catalog(SelectedDatabaseId);
            }
            else
            {
                //jpb -Actualizo el nombre del InitialCatalog de la conexion
                //Consultas cc2 = new Consultas();
                //Paso las credenciales al web service
                cc.Credentials = System.Net.CredentialCache.DefaultCredentials;
                //-----------------------------------------------------------
                Session["InitialCatalog"] = cc.Initial_Catalog(Common.Utils.SessionBaseID);

            }



            //Variables de session que utiliza el metahome
            HttpContext context = HttpContext.Current;
             

            Utils.LoginUser(User, Pass, EncriptUserData, EncryptionKey, EtiqLenguaje, login.MaxEmpl);

           
            /*********************************/
            //Una vez logueado controlo si tengo que asociar gadget al usuario
            if (Convert.ToBoolean(ConfigurationManager.AppSettings["Controlar_Primer_Acceso"]))
            {               
               cc.Controlar_Gadget_EnLoguin(User, Pass, "false", SelectedDatabaseId);               
            }
            /*********************************/

            //Guardo el webservices utilizado

            String[] Ws_Actual = Regex.Split(Convert.ToString(ConfigurationManager.AppSettings["RootWS"]), "//");
            Ws_Actual = Regex.Split(Ws_Actual[1], "/");
            Session["RHPRO_WS"] = Ws_Actual[Ws_Actual.Length - 2];

            /*********************************/
            Usuarios CUsuarios = new Usuarios();
            CUsuarios.Inicializar_Estilos(true);
            Session["RHPRO_NombreModulo"] = "RHPROX2";
            /**********************************/
            

            //Vuelvo RHPRO_HayTraducciones en vacio para que resuelva las traducciones con el lenguaje del usuario
            System.Web.HttpContext.Current.Session["RHPRO_EtiqTraducidasHome"] = "";
            System.Web.HttpContext.Current.Session["RHPRO_HayTraducciones"] = "";


            /**********************************/
            //Guarda el string de conexion para pasarselo al conn_db          
            /**********************************/
            Utils.CopyAspNetSessionToAspSession();
            /********************************/
            //Cambio el menu Login


            //if (Utils.IsUserLogin)
            //{

            //    Session["Username"] = User;
            //    Session["Password"] = Pass;
            //}



            Session["RHPro_PreLoguin"] = "0";

        }




    }
}
