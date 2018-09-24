using System;
using System.Collections;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Xml.Linq;

using System.Threading;
using ServicesProxy;
using ServicesProxy.rhdesa;

namespace RHPro.Controls
{
  
    public partial class Idiomas : System.Web.UI.UserControl
    {

        //Se define el objeto conexión
        public System.Data.SqlClient.SqlConnection conn;
        public System.Data.SqlClient.SqlDataReader reader;
        public System.Data.SqlClient.SqlCommand sql;
        protected System.Web.UI.WebControls.Panel panelFlags;
    
 /*
        public CustomLogin Def_CustomLogin;

        public void AsignaCL( RHPro.Controls.CustomLogin Clegal) {
            Def_CustomLogin = Clegal;          
        }
  */

        protected void Page_Load(object sender, EventArgs e)
        {
            Cargar_Banderas();
            //Preparo las variables de sesion relacionadas al idioma
            if ((System.Web.HttpContext.Current.Session["Lenguaje"] == null)
                 && (System.Web.HttpContext.Current.Session["ArgTitulo"] == null)
                 && (System.Web.HttpContext.Current.Session["ArgUrlImagen"] == null))
            {
                System.Web.HttpContext.Current.Session["Lenguaje"] = "es-AR";
                System.Web.HttpContext.Current.Session["ArgTitulo"] = "Español - Latam";
                System.Web.HttpContext.Current.Session["ArgUrlImagen"] = "~/img/Flags/flag_esAR.png";
            }

            Common.Utils.Lenguaje = ((String)System.Web.HttpContext.Current.Session["Lenguaje"]).Replace("-","");          
            Idioma.Text = (String)System.Web.HttpContext.Current.Session["ArgTitulo"];
            Bandera.ImageUrl = (String)System.Web.HttpContext.Current.Session["ArgUrlImagen"];           
                       
        }

        //Carga las banderas segun la base de datos activa
        protected void Cargar_Banderas(){
         
            string BaseId = Common.Utils.SessionBaseID;
            string UserName = Common.Utils.SessionUserName;

            string sql = "SELECT * FROM lenguaje WHERE lenactivo=-1 ORDER BY lendesabr ASC";  
            
            Consultas cc = new Consultas();          
            DataSet ds = cc.get_DataSet(sql, BaseId);
            Repeater1.DataSource = ds;
            Repeater1.DataBind();
        }

      
        //Evento que actuliza el idioma del sitio segun la bandera seleccionada
        protected void Idioma_Click(object sender, EventArgs e)
        {   //Recupero los argumentos que vienen de la forma ArgIdioma@ArgTitulo@ArgUrlImagen
            String Leng =(String)((LinkButton)sender).CommandArgument;
            Leng = Leng.Substring(0, Leng.Length);          
            string[] Args = System.Text.RegularExpressions.Regex.Split(Leng, "@");   
            String ArgIdioma = Args[0];
            String ArgTitulo = Args[1]; 
            String ArgUrlImagen = Args[2];            
            RefrescarComboIdioma(ArgIdioma,ArgTitulo,ArgUrlImagen);
        }

        //Refresca las variables de sesion relacionadas al idioma
        public void RefrescarComboIdioma(string ArgIdioma,string ArgTitulo, string ArgUrlImagen)
        {
            Session["ChangeLanguage"] = "1";
            
            System.Web.HttpContext.Current.Session["Lenguaje"] = ArgIdioma;
            System.Web.HttpContext.Current.Session["ArgTitulo"] = ArgTitulo;
            System.Web.HttpContext.Current.Session["ArgUrlImagen"] = ArgUrlImagen;
            Common.Utils.Lenguaje = ArgIdioma;
            Idioma.Text = ArgTitulo;
            Bandera.ImageUrl = ArgUrlImagen;           

            //Por cada variable de sesion se lo paso de .net a asp. Esto sirve para actualizar el idioma desde .net a asp            
            if (Common.Utils.SesionIniciada)
            {
                Common.Utils.CopyAspNetSessionToAspSession();               
            }

        }   


    }
}