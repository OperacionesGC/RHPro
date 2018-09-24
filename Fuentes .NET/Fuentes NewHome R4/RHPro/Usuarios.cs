using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using ServicesProxy.rhdesa;
using System.Data;
using Common;
using Entities;
using System.Configuration;
 

namespace RHPro
{
    public class Usuarios
    {
        public List<String> getPerfilesUsuario(String idUser)
        {
            List<String> Lista = new List<string>();
            string[] Misplit1;
            string[] Misplit2;
            Consultas cc = new Consultas();
            
            String fechaActual = Fecha.cambiaFecha(DateTime.Today.ToString("dd/MM/yyyy"), cc.get_TipoBase(Utils.SessionBaseID));
            //Busco el perfil del usuario
            string sql = "SELECT listperfnro ";
            sql = sql + " FROM user_perfil ";
            sql = sql + " WHERE UPPER(user_perfil.iduser) = '" + Utils.SessionUserName + "' ";
            sql = sql + " UNION ALL ";
            sql = sql + " SELECT listperfnro from bk_perfil INNER JOIN bk_cab ON bk_cab.bkcabnro = bk_perfil.bkcabnro ";
            sql = sql + " AND (bk_cab.fdesde <= " + fechaActual + " AND (bk_cab.fhasta IS NULL OR bk_cab.fhasta >= " + fechaActual + " )) ";
            sql = sql + " AND upper(bk_perfil.iduser) = Upper('" + Utils.SessionUserName + "')";
                      
            DataTable dt = cc.get_DataTable(sql, Utils.SessionBaseID);
            if (dt.Rows.Count > 0)
            {
                Misplit1 = (Convert.ToString(dt.Rows[0]["listperfnro"])).Split(',');//.Split(new Char[] {','});
                foreach (string Perfil in Misplit1)
                {
                    Lista.Add(Perfil);
                }
                if (dt.Rows.Count > 1)
                {
                    Misplit2 = (Convert.ToString(dt.Rows[1]["listperfnro"])).Split(',');//.Split(new Char[] {','});
                    foreach (string Perfil2 in Misplit2)
                    {
                        Lista.Add(Perfil2);
                    }
                }
            }


            return Lista;
        }



        public void Inicializar_Estilos(Boolean Logueado)
        {
 
            String sql = " SELECT * FROM estilo_homex2 X2 ";
            sql += " inner join estilos_home H On H.idestilo = X2.idcarpetaestilo  ";

            if (Logueado)
            {
                sql += " WHERE codestilo =(select U2.estiloactivo from estilos_home_user U2 where Upper(U2.iduser)=Upper('" + Utils.SessionUserName + "')  )  "; 
            }
            else
            { 
                sql += " WHERE defecto = -1 "; 
            }        


            Consultas cc = new Consultas();
            //Paso las credenciales al web service
            //cc.Credentials = System.Net.CredentialCache.DefaultCredentials;
            
            DataTable dt = cc.get_DataTable(sql, Utils.SessionBaseID);
 
            if (dt.Rows.Count > 0)
            {
              
                
                System.Web.HttpContext.Current.Session["EstiloR4_coloricono"] = Convert.ToString(dt.Rows[0]["coloricono"]);                
                System.Web.HttpContext.Current.Session["EstiloR4_fondoCabecera"] = Convert.ToString(dt.Rows[0]["fondoCabecera"]);
                System.Web.HttpContext.Current.Session["EstiloR4_fuentePiso"] = Convert.ToString(dt.Rows[0]["fuentePiso"]);
                System.Web.HttpContext.Current.Session["EstiloR4_fondoFecha"] = Convert.ToString(dt.Rows[0]["fondoFecha"]);
                System.Web.HttpContext.Current.Session["EstiloR4_fuenteFecha"] = Convert.ToString(dt.Rows[0]["fuenteFecha"]);
                System.Web.HttpContext.Current.Session["EstiloR4_fondoModulos"] = Convert.ToString(dt.Rows[0]["fondoModulos"]);
                System.Web.HttpContext.Current.Session["EstiloR4_fuenteModulos"] = Convert.ToString(dt.Rows[0]["fuenteModulos"]);
                System.Web.HttpContext.Current.Session["EstiloR4_fondoPiso"] = Convert.ToString(dt.Rows[0]["fondoPiso"]);
                System.Web.HttpContext.Current.Session["EstiloR4_coloriconomenutop"] = Convert.ToString(dt.Rows[0]["coloriconomenutop"]);
                System.Web.HttpContext.Current.Session["EstiloR4_fondocontppal"] = Convert.ToString(dt.Rows[0]["fondocontppal"]);

                System.Web.HttpContext.Current.Session["EstiloR4_fondoGadget"] = Convert.ToString(dt.Rows[0]["fondogadget"]);
                System.Web.HttpContext.Current.Session["EstiloR4_fuenteGadget"] = Convert.ToString(dt.Rows[0]["fuentegadget"]);
                 

                if (!dt.Rows[0]["logoEstilo"].Equals(System.DBNull.Value))
                    System.Web.HttpContext.Current.Session["EstiloR4_logoEstilo"] = dt.Rows[0]["logoEstilo"];
                else
                    System.Web.HttpContext.Current.Session["EstiloR4_logoEstilo"] = ConfigurationManager.AppSettings["urlLogo"];

                System.Web.HttpContext.Current.Session["CarpetaEstilo"] = dt.Rows[0]["estilocarpeta"];
                 
            }
            else
            {
                sql = " SELECT * FROM estilo_homex2 X2 ";
                sql += " inner join estilos_home H On H.idestilo = X2.idcarpetaestilo  ";
                sql += " WHERE defecto = -1 ";
                dt = cc.get_DataTable(sql, Utils.SessionBaseID);
                if (dt.Rows.Count > 0)
                {
                   System.Web.HttpContext.Current.Session["EstiloR4_coloricono"] = Convert.ToString(dt.Rows[0]["coloricono"]);
                    System.Web.HttpContext.Current.Session["EstiloR4_fondoCabecera"] = Convert.ToString(dt.Rows[0]["fondoCabecera"]);
                    System.Web.HttpContext.Current.Session["EstiloR4_fuentePiso"] = Convert.ToString(dt.Rows[0]["fuentePiso"]);
                    System.Web.HttpContext.Current.Session["EstiloR4_fondoFecha"] = Convert.ToString(dt.Rows[0]["fondoFecha"]);
                    System.Web.HttpContext.Current.Session["EstiloR4_fuenteFecha"] = Convert.ToString(dt.Rows[0]["fuenteFecha"]);
                    System.Web.HttpContext.Current.Session["EstiloR4_fuenteModulos"] = Convert.ToString(dt.Rows[0]["fuenteModulos"]);
                    System.Web.HttpContext.Current.Session["EstiloR4_fondoPiso"] = Convert.ToString(dt.Rows[0]["fondoPiso"]);
                    System.Web.HttpContext.Current.Session["EstiloR4_coloriconomenutop"] = Convert.ToString(dt.Rows[0]["coloriconomenutop"]);
                    System.Web.HttpContext.Current.Session["EstiloR4_fondocontppal"] = Convert.ToString(dt.Rows[0]["fondocontppal"]);
                    System.Web.HttpContext.Current.Session["EstiloR4_fondoGadget"] = Convert.ToString(dt.Rows[0]["fondogadget"]);
                    System.Web.HttpContext.Current.Session["EstiloR4_fuenteGadget"] = Convert.ToString(dt.Rows[0]["fuentegadget"]);

                    if (!dt.Rows[0]["logoEstilo"].Equals(System.DBNull.Value))
                        System.Web.HttpContext.Current.Session["EstiloR4_logoEstilo"] = dt.Rows[0]["logoEstilo"];
                    else
                        System.Web.HttpContext.Current.Session["EstiloR4_logoEstilo"] = ConfigurationManager.AppSettings["urlLogo"];

                    System.Web.HttpContext.Current.Session["CarpetaEstilo"] = dt.Rows[0]["estilocarpeta"];
                }
            }
        }


    }
}
