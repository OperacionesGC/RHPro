using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using ServicesProxy.rhdesa;
using System.Data;
using Common;
using System.Collections;

namespace RHPro
{
    public class ConfiguracionesHome
    {
        public static Hashtable DiccionarioConfnro = new Hashtable();
                  
        /// <summary>
        /// Verifica si un boton esta habilitado en un fuente del home
        /// </summary>
        /// <param name="boton"></param>
        /// <param name="fuente"></param>
        /// <returns></returns>
        public bool Habilitado(String boton, string fuente)
        {
            DiccionarioConfnro.Clear();
            DiccionarioConfnro.Add("FAVORITOS",17);
            DiccionarioConfnro.Add("GADGETS", 18);
            DiccionarioConfnro.Add("ESTILOS", 19);
            
            bool Salida = true;
            Consultas cc = new Consultas();
            ////Paso las credenciales al web service
            //cc.Credentials = System.Net.CredentialCache.DefaultCredentials;
            String sql = "";
            DataTable dt;
            //Si es alguno de estos botones verifico que haya alguna configuracion en el confper
            if ((boton.ToUpper() == "FAVORITOS") || (boton.ToUpper() == "GADGETS") || (boton.ToUpper() == "ESTILOS"))
            {                
                sql = " select confint,confactivo from confper where confnro =  " + DiccionarioConfnro[boton.ToUpper()];                
                dt = cc.get_DataTable(sql, Utils.SessionBaseID);
                if (dt.Rows.Count > 0)
                {
                    if ((Convert.ToInt32(dt.Rows[0]["confactivo"]) != -1) || (Convert.ToInt32(dt.Rows[0]["confint"]) != -1))
                    {
                        return false;
                    }
                }
            }
            
            sql = " SELECT btnaccess FROM menubtn  ";
            sql += " WHERE btnpagina = '" + fuente + "' and upper(btnnombre) = upper('" + boton + "')";
            dt = cc.get_DataTable(sql, Utils.SessionBaseID);

            if (dt.Rows.Count > 0)
            {
                Usuarios Usr = new Usuarios();
                List<String> ListaPerfUsr = Usr.getPerfilesUsuario(Utils.SessionUserName);
                Salida = (Convert.ToString(dt.Rows[0]["btnaccess"]) == "*") || Utils.Habilitado(ListaPerfUsr, Convert.ToString(dt.Rows[0]["btnaccess"]));
            }

            return Salida;

        }


        /// <summary>
        /// Verifica si esta habilitado la visibilidad de los gadgets en la configuracion de Empresa
        /// </summary>
        /// <returns></returns>
        public bool Gadgets_Habilitados()
        {
            bool Salida = true;

            if (Convert.ToString(System.Web.HttpContext.Current.Session["RHPRO_Gadgets_Habilitados"]) == "")
            {

                Consultas cc = new Consultas();
                ////Paso las credenciales al web service
                //cc.Credentials = System.Net.CredentialCache.DefaultCredentials;
                String sql = "";
                DataTable dt;

                //Si es alguno de estos botones verifico que haya alguna configuracion en el confper  y este habilitada                    
                sql = " select confint,confactivo from confper where confnro =  18 and confactivo = -1";
                dt = cc.get_DataTable(sql, Utils.SessionBaseID);
                if (dt.Rows.Count == 0)
                {
                    System.Web.HttpContext.Current.Session["RHPRO_Gadgets_Habilitados"] = "0";
                    Salida = false;
                    //return false;
                }
                else
                {
                    System.Web.HttpContext.Current.Session["RHPRO_Gadgets_Habilitados"] = "-1";
                    Salida = true;
                }

                 
            }
            else
            {

                Salida = ((String)System.Web.HttpContext.Current.Session["RHPRO_Gadgets_Habilitados"] == "-1");
            }
           
            return Salida;

        }

        

    }




}
