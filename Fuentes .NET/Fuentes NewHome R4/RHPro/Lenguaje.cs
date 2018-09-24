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
        /// Retorna el [lencod] del lenguaje default configurado en el key Idioma de Settings.config
        /// </summary>
        public string Etiq_Leng_Default()
        {
           string EtiqLenguaje = "es-AR";
           string[] Arr;
           if (ConfigurationManager.AppSettings["Idioma"] != null)
           {
               Arr = ((String)ConfigurationManager.AppSettings["Idioma"]).Split(',');
               if (Arr.Length > 0)
                   EtiqLenguaje =  Arr[0];
           }
           
           return EtiqLenguaje;
        }

        public Hashtable Preparar_Arr_Etiquetas()
        {
            Hashtable Dicc_Idioma = new Hashtable();
            String EtiqLenguaje = (String)System.Web.HttpContext.Current.Session["Lenguaje"];
            EtiqLenguaje = EtiqLenguaje.Replace("-", "");


            String sql;

            try
            {

                //Busco la etiqueta en el lenguaje seleccionado                
                sql = "SELECT etiqueta,esAR," + EtiqLenguaje + " EtiqTraducida FROM lenguaje_etiqueta ";

                Consultas cc = new Consultas();
                //Paso las credenciales al web service
                //cc.Credentials = System.Net.CredentialCache.DefaultCredentials;
                //-----------------------------------------------------------
                DataTable dt = cc.get_DataTable(sql, Utils.SessionBaseID);

                foreach (System.Data.DataRow dr in dt.Rows)
                {
                    if (!Dicc_Idioma.Contains(Convert.ToString(dr["etiqueta"])))
                    {
                        if ((dr["EtiqTraducida"] == null) || (dr["EtiqTraducida"].ToString() == ""))
                        {
                            Dicc_Idioma.Add(Convert.ToString(dr["etiqueta"]), Convert.ToString(dr["EtiqTraducida"]));
                        }
                        else
                        {
                            if ((dr["esAR"] == null) || (dr["esAR"].ToString() == ""))
                            {
                                Dicc_Idioma.Add(Convert.ToString(dr["etiqueta"]), Convert.ToString(dr["esAR"]));
                            }
                            else
                            {
                                if ((dr["etiqueta"] == null) || (dr["etiqueta"].ToString() == ""))
                                {
                                    Dicc_Idioma.Add(Convert.ToString(dr["etiqueta"]), Convert.ToString(dr["etiqueta"]));
                                }
                            }
                        }
                    }

                }

            }
            catch (Exception ex) { }

            return Dicc_Idioma;

        }


        /// <summary>
        /// Retorna el lenguaje default configurado en el key Idioma de Settings.config
        /// separado por comas: "[lencod], [lendesabr], [paisdesc]"
        /// </summary>
        public string Lenguaje_Default()
        {
            string EtiqLenguaje = "es-AR,Español Latam,ARGENTINA";
            string Aux;
             if (ConfigurationManager.AppSettings["Idioma"] != null)
               Aux = (String)ConfigurationManager.AppSettings["Idioma"];
             else 
               Aux = "es-AR,Español Latam,ARGENTINA";
 
            if (Aux != "")
                EtiqLenguaje = Aux;

            return EtiqLenguaje;

           /* string sql = "SELECT pais.paisdef defecto,lencod,lendesabr,paisdesc FROM lenguaje  ";
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
            */

        }

        /// <summary>
        /// Retorna un string separado por comas: "[lencod], [lendesabr]" que describe el lenguaje seleccionado para un determinado usuario
        /// </summary>
        public string Lenguaje_Usuario(string Usuario)
        {            
            string sql = "SELECT lenguaje.lencod, lenguaje.lendesabr FROM user_per  ";
            sql += "INNER JOIN lenguaje ON lenguaje.lennro = user_per.lennro ";
            sql += "WHERE Upper(iduser) = Upper('" + Usuario + "')";

            string lengusr = "esAR,Español-Latam";
            //Busco el pais default y su lenguaje///

             Consultas cc = new Consultas();
             //Paso las credenciales al web service
             //cc.Credentials = System.Net.CredentialCache.DefaultCredentials;
             //-----------------------------------------------------------
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

            /* ************************ */
            if ((Convert.ToString(System.Web.HttpContext.Current.Session["RHPRO_HayTraducciones"])!="")
                &&(System.Web.HttpContext.Current.Session["RHPRO_HayTraducciones"]!=null))
            {
                Hashtable Dicc = (Hashtable) System.Web.HttpContext.Current.Session["RHPRO_EtiqTraducidasHome"];
                if (Dicc.Contains(Etiqueta))
                {
                    LabelTraducido = (String)Dicc[Etiqueta];
                    return LabelTraducido;                    
                }                
            }            

            /* ************************** */

            try
            {
                if (System.Web.HttpContext.Current.Session["Lenguaje"] != null)
                {
                    EtiqLenguaje =  (String)System.Web.HttpContext.Current.Session["Lenguaje"];

                    //if ( (EtiqLenguaje != "") && (EtiqLenguaje != null))
                    if (EtiqLenguaje != "")
                    {
                        EtiqLenguaje = EtiqLenguaje.Replace("-", "");
                        //Busco la etiqueta en el lenguaje seleccionado                
                        sql = "SELECT etiqueta,esAR," + EtiqLenguaje + " EtiqTraducida FROM lenguaje_etiqueta ";                     
                        sql += " WHERE etiqueta = '" + Etiqueta + "' ";

                        Consultas cc = new Consultas();
                        //Paso las credenciales al web service
                        //cc.Credentials = System.Net.CredentialCache.DefaultCredentials;
                       
                        DataTable dt = cc.get_DataTable(sql, Utils.SessionBaseID);


                        /* **************************************************************************** */ 
                        if (dt.Rows.Count > 0)
                        {
                            if ((dt.Rows[0]["EtiqTraducida"] != null) && (dt.Rows[0]["EtiqTraducida"].ToString() != ""))
                            {                               
                                LabelTraducido = Convert.ToString(dt.Rows[0]["EtiqTraducida"]);
                            }
                            else
                            {
                                if ((dt.Rows[0]["esAR"] != null) && (dt.Rows[0]["esAR"].ToString() != ""))
                                {
                                    LabelTraducido = Convert.ToString(dt.Rows[0]["esAR"]);
                                }
                                else
                                {
                                    if ((dt.Rows[0]["etiqueta"] != null) && (dt.Rows[0]["etiqueta"].ToString() != ""))
                                    {
                                        LabelTraducido = Convert.ToString(dt.Rows[0]["etiqueta"]);
                                    }
                                }
                            }
                        }

                        Hashtable Dicc;
                        if ((Convert.ToString(System.Web.HttpContext.Current.Session["RHPRO_HayTraducciones"]) != "") && (System.Web.HttpContext.Current.Session["RHPRO_HayTraducciones"] != null))
                             Dicc = (Hashtable)System.Web.HttpContext.Current.Session["RHPRO_EtiqTraducidasHome"];                                                  
                        else                        
                             Dicc = new Hashtable();                                                  

                        if (!Dicc.Contains(LabelTraducido))
                        {
                            Dicc.Add(Etiqueta, LabelTraducido);
                            System.Web.HttpContext.Current.Session["RHPRO_EtiqTraducidasHome"] = Dicc;
                            System.Web.HttpContext.Current.Session["RHPRO_HayTraducciones"] = "1";
                        }
                        /* ************************************ */ 

                        //foreach (System.Data.DataRow dr in dt.Rows)
                        //{
                        //    if ((dr[EtiqLenguaje] == null) || (dr[EtiqLenguaje].ToString() == ""))
                        //        LabelTraducido = Etiqueta;
                        //    else
                        //        LabelTraducido = dr[EtiqLenguaje].ToString();
                        //    break;
                        //}                        
                    }
                }
                
            }
            catch (Exception ex) {  }
            return LabelTraducido;
        }


        /// <summary>
        /// Retorna un texto traducido de la descripción de un determinado modulo, según el idioma activo
        /// </summary>
        public string Traducir_Modulo(string Etiqueta, string MenuName)
        {
           
                       
            String EtiqLenguaje;
 /*
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
                        sql += " WHERE modulo='HOME' AND etiqueta = '" + MenuName + "' ";

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
            catch (Exception ex) { }
            return LabelTraducido;
            */
 
            ////String LabelTraducido;
            //String EtiqLenguaje;

            ////Ahora traigo la traduccion de los modulos desde la libreria de ConsultasBaseC
            Consultas cc = new Consultas();
            EtiqLenguaje = (String)System.Web.HttpContext.Current.Session["Lenguaje"];
            EtiqLenguaje = EtiqLenguaje.Replace("-", "");
            //Paso las credenciales al web service
            //cc.Credentials = System.Net.CredentialCache.DefaultCredentials;
            //-----------------------------------------------------------
            return cc.get_Traduccion_Modulo(EtiqLenguaje, MenuName, Utils.SessionBaseID);
            ////-------------

            /*

            LabelTraducido = Etiqueta;


            if (((String)System.Web.HttpContext.Current.Session["Lenguaje"]).Length == 5)
            {
                EtiqLenguaje = (String)System.Web.HttpContext.Current.Session["Lenguaje"];
                EtiqLenguaje = "menudetalle" + EtiqLenguaje.ToUpper();
            }
            else
            {
                if (((String)System.Web.HttpContext.Current.Session["Lenguaje"]).Length == 4)
                {
                    EtiqLenguaje = (String)System.Web.HttpContext.Current.Session["Lenguaje"];
                    EtiqLenguaje = EtiqLenguaje.Substring(0, 2) + "-" + EtiqLenguaje.Substring(2, 2);
                    EtiqLenguaje = "menudetalle" + EtiqLenguaje.ToUpper();
                }
                else
                {
                    EtiqLenguaje = "menudetalle";
                }
            }

           // if ((EtiqLenguaje.ToUpper() == "ENUS") || (EtiqLenguaje.ToUpper() == "PTBR") || (EtiqLenguaje.ToUpper() == "PTPT") || (EtiqLenguaje.ToUpper() == "ESES"))            


            //ver
            EtiqLenguaje = "menudetalle";
            //
           
            try
            {
                    Consultas cc = new Consultas();                   
                    DataTable dt = cc.get_Traduccion_Modulo(EtiqLenguaje, MenuName, Utils.SessionBaseID);
                    if (dt == null)
                        EtiqLenguaje = "menudetalle";
                 
                    foreach (System.Data.DataRow dr in dt.Rows)
                    {
                        if ((dr[EtiqLenguaje] == null) || (dr[EtiqLenguaje].ToString() == ""))
                            LabelTraducido = Etiqueta;
                        else
                            LabelTraducido = (String) dr[EtiqLenguaje];
                        break;
                    }
               
             
            }
            catch (Exception exmod) {
                throw exmod; 

                //En el caso que no exista el campo referente al detalle del modulo en algun idioma, toma el default
                //Consultas cc = new Consultas();
                //DataTable dt = cc.get_Traduccion_Modulo("menudetalle", MenuName, Utils.SessionBaseID);

                //foreach (System.Data.DataRow dr in dt.Rows)
                //{
                //    if ((dr["menudetalle"] == null) || (dr["menudetalle"].ToString() == ""))
                //        LabelTraducido = Etiqueta;
                //    else
                //        LabelTraducido = dr["menudetalle"].ToString();
                //    break;
                //}
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
        
             */
        }
    }
}
