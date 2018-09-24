using System;
using System.Data;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Xml.Linq;

using System.Collections;
using ServicesProxy.rhdesa;

 
using System.Collections.Generic;
 
using Common;
using Entities;
using ServicesProxy;
using System.Threading;

namespace RHPro
{
    public class Lenguaje 
    {


        System.Data.SqlClient.SqlConnection conex;
 
        /// <summary>
        /// Constructor de la clase
        /// </summary>
        public  Lenguaje()
        {             
        }

        public System.Data.SqlClient.SqlConnection ConnexionDef() {
            String conexString = "";
            conexString =   (String)System.Web.HttpContext.Current.Session["ConnString"];
            conex = new System.Data.SqlClient.SqlConnection(conexString);
            return conex;
        }
        /// <summary>
        /// Retorna el lenguaje activo
        /// </summary>
        public string Idioma() {           
            return (String)System.Web.HttpContext.Current.Session["Lenguaje"]; 
        }

        /// <summary>
        /// Retorna un string separado por comas: "[lencod], [lendesabr], [paisdesc]"
        /// </summary>
        public string Lenguaje_Default()
        {
            string sql = "SELECT pais.paisdef defecto,lencod,lendesabr,paisdesc FROM lenguaje  ";
                   sql +=  "INNER JOIN pais ON pais.paisnro = lenguaje.paisnro ";
                   sql += "WHERE lenactivo <> 0 AND pais.paisdef=-1 ";
                   sql +=  "     ORDER BY paisdef,paisdesc ASC";
            
 
            string lengdef = "esAR,Español-Latam, ARGENTINA";
            //Busco el pais default y su lenguaje///
            Consultas cc = new Consultas();
            DataTable dt = cc.get_DataTable(sql, Utils.SessionBaseID);

            foreach (System.Data.DataRow dr in dt.Rows)
            {
                if ((Int16)dr["defecto"] == -1) {
                    lengdef = dr["lencod"] + "," + dr["lendesabr"] + "," + dr["paisdesc"];
                }
            }
           
            return lengdef;
        }

        /// <summary>
        /// Retorna un string separado por comas: "[lencod], [lendesabr]" que describe el lenguaje seleccionado para un determinado usuario
        /// </summary>
        public string Lenguaje_Usuario(string Usuario)
        {
            //string Cs = "Password=ess;Persist Security Info=True;User ID=ess;Initial Catalog=Base_0_R3_ARG;Data Source=RHDESA";
            string sql = "SELECT lenguaje.lencod, lenguaje.lendesabr FROM user_per  ";
            sql += "INNER JOIN lenguaje ON lenguaje.lennro = user_per.lennro ";
            sql += "WHERE iduser = '"+Usuario+"'";

            string lengusr = "esAR,Español-Latam";
            //Busco el pais default y su lenguaje///

             Consultas cc = new Consultas();
             DataTable dt = cc.get_DataTable(sql, Utils.SessionBaseID);
            foreach (System.Data.DataRow dr in dt.Rows)
            {               
                    lengusr = dr["lencod"] + "," + dr["lendesabr"];                
            }
           
            return lengusr;
        }

        /// <summary>
        /// Retorna el label traducido al idioma activo
        /// </summary>
        public string Label_Home(string Etiqueta)
        {

            String sql;
            String LabelTraducido;
            String EtiqLenguaje;            
            
            LabelTraducido = Etiqueta;           
            try
            {
                if (System.Web.HttpContext.Current.Session["Lenguaje"] != null)
                {
                    EtiqLenguaje = (String)System.Web.HttpContext.Current.Session["Lenguaje"];

                    //if ( (EtiqLenguaje != "") && (EtiqLenguaje != null))
                    if (EtiqLenguaje != "")
                    {
                        EtiqLenguaje = EtiqLenguaje.Replace("-", "");
                        //Busco la etiqueta en el lenguaje seleccionado                
                        sql = "SELECT " + EtiqLenguaje + " FROM lenguaje_etiqueta ";
                        sql += " WHERE etiqueta = '" + Etiqueta + "' ";

                        Consultas cc = new Consultas();
                        DataTable dt = cc.get_DataTable(sql, Utils.SessionBaseID); 

                        foreach (System.Data.DataRow dr in dt.Rows)
                        {
                            if ((dr[EtiqLenguaje] == null) || (dr[EtiqLenguaje].ToString() == ""))
                                LabelTraducido = Etiqueta;
                            else
                                LabelTraducido = dr[EtiqLenguaje].ToString();
                            break;
                        }                        
                    }
                }
                
            }
            catch (Exception ex) {  }
            return LabelTraducido;
        }


        /// <summary>
        /// Retorna un texto traducido de la descripción de un determinado modulo, según el idioma activo
        /// </summary>
        public string Traducir_Modulo(string Etiqueta)
        {
 
            String LabelTraducido;
            String EtiqLenguaje;

            LabelTraducido = Etiqueta;
            EtiqLenguaje = (String)System.Web.HttpContext.Current.Session["Lenguaje"];
            if ((EtiqLenguaje.ToUpper() == "ENUS") || (EtiqLenguaje.ToUpper() == "PTBR") || (EtiqLenguaje.ToUpper() == "PTPT") || (EtiqLenguaje.ToUpper() == "ESES"))
            {
                EtiqLenguaje = EtiqLenguaje.Substring(0, 2) + "-" + EtiqLenguaje.Substring(2, 2);
                EtiqLenguaje = "menudetalle" + EtiqLenguaje.ToUpper();

                Consultas cc = new Consultas();
                DataTable dt = cc.get_Traduccion_Modulo(EtiqLenguaje, Etiqueta, Utils.SessionBaseID);

                
                foreach (System.Data.DataRow dr in dt.Rows)
                {
                    if ((dr[EtiqLenguaje] == null) || (dr[EtiqLenguaje].ToString() == ""))
                        LabelTraducido = Etiqueta;
                    else
                        LabelTraducido = dr[EtiqLenguaje].ToString();
                    break;
                }

            }

            return LabelTraducido; 


            //String sql;
            //String LabelTraducido;
            //String EtiqLenguaje;
          
            //LabelTraducido = Etiqueta;
            //EtiqLenguaje = (String)System.Web.HttpContext.Current.Session["Lenguaje"];
            //if ((EtiqLenguaje.ToUpper() == "ENUS") || (EtiqLenguaje.ToUpper() == "PTBR") || (EtiqLenguaje.ToUpper() == "PTPT") || (EtiqLenguaje.ToUpper() == "ESES"))
            //{
            //    EtiqLenguaje = EtiqLenguaje.Substring(0, 2) + "-" + EtiqLenguaje.Substring(2, 2);
            //    EtiqLenguaje = "menudetalle" + EtiqLenguaje.ToUpper();
                
                
            //    sql = "SELECT  " + EtiqLenguaje + " FROM menumstr ";           
            //    sql += " WHERE CAST(menudetalle AS VARCHAR(MAX)) = CAST('" + Etiqueta + "' AS VARCHAR(MAX) ) ";


              
            //    Consultas cc = new Consultas();
            //    DataTable dt = cc.get_DataTable(sql, Utils.SessionBaseID);
            //    foreach (System.Data.DataRow dr in dt.Rows)
            //    {
            //        if ((dr[EtiqLenguaje] == null) || (dr[EtiqLenguaje].ToString() == ""))
            //            LabelTraducido = Etiqueta;
            //        else
            //            LabelTraducido = dr[EtiqLenguaje].ToString();
            //        break;
            //    }
            //}

            //return  LabelTraducido; 
        }
    }
}
