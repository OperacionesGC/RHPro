using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using ServicesProxy.rhdesa;
using Common;
using System.Data;

namespace RHPro.Controls
{
    public partial class Contenedor_Estilos : System.Web.UI.UserControl
    {
        protected void Page_Load(object sender, EventArgs e)
        {
        
            Levantar_Estilos();
        }

        public void Levantar_Estilos()
        {
            String sql = "";
            Consultas cc = new Consultas();
            //Paso las credenciales al web service
            //cc.Credentials = System.Net.CredentialCache.DefaultCredentials;
            //Recupero todos los estilos                        
            sql = "SELECT X2.*, H.estilocarpeta,H.estiloRGB from estilo_homex2 X2 ";
            sql +=" inner join estilos_home H On H.idestilo = X2.idcarpetaestilo ";
            sql += " WHERE estdesabr<>'RHPROX2' AND activo=-1 ";
            sql += " ORDER BY estdesabr ASC ";
       
            DataSet ds = cc.get_DataSet(sql, Utils.SessionBaseID);
            
            IteradorEstilos.DataSource = ds;
            IteradorEstilos.DataBind();
        }



        public void Estilo_Click(object sender, EventArgs e)
        {
            
           
            String ArrDatos = (String)((LinkButton)sender).CommandArgument;
            ArrDatos = ArrDatos.Substring(0, ArrDatos.Length);
            string[] Args = System.Text.RegularExpressions.Regex.Split(ArrDatos, "@@");
            if (Args.Length > 0)
            {
                if ((Args[0] != "") && (Args[1] != ""))
                {
                    String idcarpetaestilo = Args[0];
                    String estilocarpeta = Args[1];
                    String codestilo = Args[2];
                    Session["CarpetaEstilo"] = estilocarpeta;
 
          
                    Consultas cc = new Consultas();
                    ////Paso las credenciales al web service
                    //cc.Credentials = System.Net.CredentialCache.DefaultCredentials;
                    
                   
                    //Paso las credenciales al web service
                    //cc.Credentials = System.Net.CredentialCache.DefaultCredentials;
                    cc.Cambiar_Estilo(Utils.SessionUserName, Utils.SessionBaseID, idcarpetaestilo, codestilo);
                    Session["RHPRO_Cambio_Estilo"] = "-1";
                    //Por cada variable de sesion se lo paso de .net a asp. Esto sirve para actualizar el idioma desde .net a asp            
                    if (Common.Utils.SesionIniciada)
                    {
                        Common.Utils.CopyAspNetSessionToAspEstilos();
                    }
                }

            }

 
        }
    }
}