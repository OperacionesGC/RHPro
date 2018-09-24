using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Common;
using Entities;
using ServicesProxy;
using System.Threading;
using ServicesProxy.rhdesa;


 
namespace RHPro
{
    
   
    public partial class Default : System.Web.UI.Page
    {        
        #region Page Handles

        public Lenguaje Obj_Lenguaje;
                //Se define el objeto conexión
        public System.Data.SqlClient.SqlConnection conn;
        public System.Data.SqlClient.SqlDataReader reader;
        public System.Data.SqlClient.SqlCommand sql;
      
         protected void Page_Load(object sender, EventArgs e)
        {
 
                if (System.Web.HttpContext.Current.Session["yaentro"] == null)
                    System.Web.HttpContext.Current.Session["yaentro"] = false;
                if (System.Web.HttpContext.Current.Session["primerIdioma"] == null)
                    System.Web.HttpContext.Current.Session["primerIdioma"] = false;

                Obj_Lenguaje = new Lenguaje();

                //Inicializo el lenguaje dejault           
                if (!Utils.IsUserLogin)
                {  //Cargo el string de conexion por defecto      

                    SetBaseIdDefault();
                    //LoadConexionDefault(); //JPB: Se elimina
                    if (!(bool)System.Web.HttpContext.Current.Session["primerIdioma"])
                    {
                        System.Web.HttpContext.Current.Session["primerIdioma"] = true;
                        string[] LengDefault = Obj_Lenguaje.Lenguaje_Default().Split(',');
                        string Idioma;
                        //string bandera = LengDefault[0];
                       
                        //string IdiomaDefault = LengDefault[0].Substring(0, 2) + "-" + LengDefault[0].Substring(2, 2);
                        if (LengDefault.Length > 0)
                            Idioma = LengDefault[1];
                        else
                            Idioma = "es-AR";
                       
                        //Common.Utils.Lenguaje = IdiomaDefault;
                        Common.Utils.Lenguaje = Obj_Lenguaje.Etiq_Leng_Default();
                        string bandera = Obj_Lenguaje.Etiq_Leng_Default().Replace("-", "");
                        System.Web.HttpContext.Current.Session["ArgTitulo"] = Idioma;
                        System.Web.HttpContext.Current.Session["ArgUrlImagen"] = "~/img/Flags/flag_" + bandera + ".png";
                    }
                }
                else
                {
                    if (!(bool)System.Web.HttpContext.Current.Session["yaentro"])
                    {

                        System.Web.HttpContext.Current.Session["yaentro"] = true;
                        string UsuarioLogueado = Common.Utils.SessionUserName;
                        string[] LengDefault = Obj_Lenguaje.Lenguaje_Usuario(UsuarioLogueado).Split(',');

                        string bandera = LengDefault[0];
                        string IdiomaDefault = LengDefault[0].Substring(0, 2) + "-" + LengDefault[0].Substring(2, 2);                    
                        string Idioma = LengDefault[1];

                        Common.Utils.Lenguaje = IdiomaDefault;

                        System.Web.HttpContext.Current.Session["ArgTitulo"] = Idioma;
                        System.Web.HttpContext.Current.Session["ArgUrlImagen"] = "~/img/Flags/flag_" + bandera + ".png";
                    }

                }

                cLogin.UserLogin += cLogin_UserLogin;
                cLogin.UserLogout += cLogin_UserLogout;
                Modulos.AsignarContPpal(ContenedorPrincipal);
                if (bool.Parse(ConfigurationManager.AppSettings["VisualizarFooter"]))
                {
                    versionMI.InnerText = VersionServiceProxy.Find(Utils.SessionBaseID, Thread.CurrentThread.CurrentCulture.Name);//carga la version
                    patchMI.InnerText = PatcheServiceProxy.Find(Utils.SessionBaseID, Thread.CurrentThread.CurrentCulture.Name);//carga el patch
                
                }

 
        }

        /// <summary>
        /// jpb: Carga la base default
        /// </summary>
        protected internal void SetBaseIdDefault()
        {
            string dsm = ConfigurationManager.AppSettings["DatabaseSelectionMethod"].ToLower();
            System.Collections.Generic.List<DataBase> DataBases = DataBaseServiceProxy.Find(dsm);

            //Busca la base por default del webconfig
            for (int i = 0; i < DataBases.Count; i++)
            {
                if (DataBases[i].IsDefault.ToString() == "TrueValue")//verifica si la base es default
                {
                    Utils.SessionBaseID = DataBases[i].Id;
                }
            }
        }


        /// <summary>
        /// Devuelve la fecha actual traducida, segun el lenguaje activo
        /// </summary>
        public string Traducir_Fecha() {
             
            string fecTraducida;
            string IdiomaSel = "es-ES";
            string LenguajeActivo = Obj_Lenguaje.Idioma();

            if (LenguajeActivo.Length == 4)
            {//Viene de la forma "enUS"
                IdiomaSel = LenguajeActivo.Substring(0, 2) + "-" + LenguajeActivo.Substring(2, 2);
            }
            else {
                if (LenguajeActivo.Length == 5)
                    IdiomaSel = LenguajeActivo;
            }

            DateTime dt = DateTime.Now;
            //Modifico el idioma de la fecha al idioma seleccionado
            try
            {
                Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo(IdiomaSel);
            }
            catch {
                Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("es-ES");
            }
            
            //convierto la fecha al formato largo.Ej: mieroles, 5 de Abril de 2012
            fecTraducida = dt.ToLongDateString();            
            
            return fecTraducida;        
             
        }


        /// <summary>
        /// Carga en la session "ConnString" el string de conexion a la base por default. Se encuentra en el web.config del newhome
        /// </summary>
        protected internal void LoadConexionDefault()
        {
           string dsm = ConfigurationManager.AppSettings["DatabaseSelectionMethod"].ToLower();

            System.Collections.Generic.List<DataBase> DataBases = DataBaseServiceProxy.Find(dsm);
                //Busca la base por default del webconfig
                for (int i = 0; i < DataBases.Count; i++)
                {                             
                    if (  DataBases[i].IsDefault.ToString() == "TrueValue" )//verifica si la base es default
                    {                      
                        String Cs = ConfigurationManager.ConnectionStrings[DataBases[i].Id.ToString()].ConnectionString;                      
                        System.Data.SqlClient.SqlConnection conex = new System.Data.SqlClient.SqlConnection(Cs);

                        if (System.Web.HttpContext.Current.Session["ConnString"] == null)
                        {//Copia en la variable de session "ConnString" el string de conexion                           
                          System.Web.HttpContext.Current.Session["ConnString"] = Cs;                                                   
                        }
                       
                        
                    }                 
                }
       }
    
 
        public void ActualizaGadgets(object sender, EventArgs e) {
            ContenedorPrincipal.ActualizaGadgets(sender, e);
        }

        protected void Page_SaveStateComplete(object sender, EventArgs e)
        {
            if (!Utils.IsUserLogin && Utils.SesionIniciada)
            {
                Utils.LogoutUser();
            }
        }

        #endregion

        #region Controls Handles

        protected void cLogin_UserLogin(object sender, EventArgs e)
        {
            LoadPageInfo();
        }

        protected void cLogin_UserLogout(object sender, EventArgs e)
        {
            LoadPageInfo();
        }     

        #endregion

        #region Methods

        /// <summary>
        /// Carga la informacion de la pagina
        /// </summary>
        public void LoadPageInfo()
        {   
           
        }

       


        #endregion      

        
   
    
    }




     
}

