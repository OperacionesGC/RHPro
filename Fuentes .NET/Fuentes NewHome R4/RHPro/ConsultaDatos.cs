using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using ServicesProxy.rhdesa;
using System.Data;
using Common;

namespace RHPro
{
    public class ConsultaDatos
    {

        public bool Menu_Habilitado(String menuaccess, int menumsnro)
        {
            Consultas cc = new Consultas();
            if (Convert.ToString(System.Web.HttpContext.Current.Session["RHPRO_ListaPerfUsr"]) == "")
            {
              //  System.Web.HttpContext.Current.Session["RHPRO_ListaPerfUsr"] = cc.getPerfilesUsuario(Utils.SessionUserName, Utils.SessionBaseID);
                System.Web.HttpContext.Current.Session["RHPRO_ListaPerfUsr"] = cc.getPerfilesUsuario(Utils.SessionUserName, Utils.SessionBaseID).ToList();
            }
            

            List<String> ListaPerfUsr = (List<String>)System.Web.HttpContext.Current.Session["RHPRO_ListaPerfUsr"];
            
            //String[] ListaPerfUsr = (String[])System.Web.HttpContext.Current.Session["RHPRO_ListaPerfUsr"];

            return cc.Menu_Habilitado(menuaccess, menumsnro, Utils.SessionUserName, Utils.SessionBaseID);

        }


        public DataTable Grupo_Restricciones()
        {

            Consultas cc = new Consultas();
            return cc.Grupo_Restricciones(Utils.SessionBaseID);

        }



        /// <summary>
        /// Retorna una lista de menuname, de los modulos habilitados para el usuario
        /// </summary>
        /// <returns></returns>
        public List<string> get_ModulosHabilitados()
        {
            List<string> Salida = new List<string>();
            Consultas cc = new Consultas();
            string sql = " ";
            sql = " SELECT  menumsnro, menuname,menuaccess ";
            sql += " FROM menumstr ";
            sql += " WHERE  (Upper(parent) = 'RHPRO') AND (menuactivo = - 1) ";

            DataTable dt = cc.get_DataTable(sql, Utils.SessionBaseID);
            Usuarios Usr = new Usuarios();
            //List<String> ListaPerfUsr = Usr.getPerfilesUsuario(Utils.SessionUserName);

            foreach (DataRow dr in dt.Rows)
            {
              //  if (Utils.Habilitado(ListaPerfUsr, Convert.ToString(dr["menuaccess"])))
                if (Menu_Habilitado(Convert.ToString(dr["menuaccess"]), Convert.ToInt32(dr["menumsnro"])))
                { 
                    Salida.Add(Convert.ToString(dr["menuname"]).ToUpper());
                }
            }

            return Salida;
        }


/*************************************************************************************************************/


        //public bool Menu_Habilitado(String menuaccess, int menumsnro)
        //{
        //    if (Convert.ToString(System.Web.HttpContext.Current.Session["RHPRO_ListaPerfUsr"]) == "")
        //    {
        //        Usuarios Usr = new Usuarios();
        //        Consultas cc = new Consultas();
        //        //System.Web.HttpContext.Current.Session["RHPRO_ListaPerfUsr"] = Usr.getPerfilesUsuario(Utils.SessionUserName);
        //        System.Web.HttpContext.Current.Session["RHPRO_ListaPerfUsr"] = cc.getPerfilesUsuario(Utils.SessionUserName,Utils.SessionBaseID).ToList();
        //    }

        //    List<String> ListaPerfUsr = (List<String>)System.Web.HttpContext.Current.Session["RHPRO_ListaPerfUsr"];
        //    DataTable grupo = Grupo_Restricciones();

        //    //Esta habilitado si esta habilitado por armado de menu, o por grupo de acceso            
        //    int Salida = Habilitado_Por_GrupoRestricciones(menumsnro, ListaPerfUsr, grupo);

        //    if (Salida < 1)
        //    {
        //        if (Salida == -1)
        //            return true;
        //        else
        //            return false;
        //    }
        //    else
        //    {
        //        return Utils.Habilitado(ListaPerfUsr, menuaccess);
        //        //return ((Utils.Habilitado(ListaPerfUsr, menuaccess)) || (Habilitado_Por_GrupoRestricciones(menumsnro, ListaPerfUsr, grupo)));
        //    }
        //}

        //public bool MRU_Habilitado(int menumsnro)
        //{
        //    if (Convert.ToString(System.Web.HttpContext.Current.Session["RHPRO_ListaPerfUsr"]) == "")
        //    {
        //        Usuarios Usr = new Usuarios();
        //        System.Web.HttpContext.Current.Session["RHPRO_ListaPerfUsr"] = Usr.getPerfilesUsuario(Utils.SessionUserName);
        //    }

        //    List<String> ListaPerfUsr = (List<String>)System.Web.HttpContext.Current.Session["RHPRO_ListaPerfUsr"];
        //    DataTable grupo = Grupo_Restricciones();

        //    //Esta habilitado si esta habilitado  por grupo de acceso
        //    int Salida = Habilitado_Por_GrupoRestricciones(menumsnro, ListaPerfUsr, grupo);

        //    if ((Salida == -1) || (Salida == 1))
        //        return true;
        //    else
        //        return false;
        //}

    //    public DataTable Grupo_Restricciones()
    //    {

    //        Consultas cc = new Consultas();

    //        String fechaActual = Entities.Fecha.cambiaFecha(DateTime.Today.ToString("dd/MM/yyyy"), cc.get_TipoBase(Utils.SessionBaseID));

    //        string sql = " ";
    //        sql = " SELECT A.alesch_frecrep,A.alesch_fecini,A.alesch_fecfin, A.frectipnro,A.schedhora,A.schedhora2, A.scheddesc ";
    //        sql += "     ,P.listperfnro,M.menumstrnros ,  MG.* ";
    //        sql += " FROM menugrp MG ";
    //        sql += "   inner join ale_sched A ON A.schednro = MG.schednro ";
    //        sql += "   inner join  menugrp_perf P ON P.menugnro = MG.menugnro ";
    //        sql += "   inner join  menugrp_menu M ON M.menugnro = MG.menugnro ";
    //        sql += " WHERE  ( (" + fechaActual + " >= A.alesch_fecini) AND (" + fechaActual + "<= A.alesch_fecfin)) ";

    //        DataTable dt = cc.get_DataTable(sql, Utils.SessionBaseID);

    //        return dt;
    //    }

    //    /// <summary>
    //    /// Verifica si tiene un grupo de acceso. 
    //    /// </summary>
    //    /// <param name="menumsnro"></param>
    //    /// <param name="ListaPerfUsr"></param>
    //    /// <param name="GrupoRestriccion"></param>
    //    /// <returns>-1: Puede Acceder, 0:No Puede Acceder, 1:No hay nada especificado</returns>
    //    public int Habilitado_Por_GrupoRestricciones(int menumsnro, List<String> ListaPerfUsr, DataTable GrupoRestriccion)
    //    {


    //        //Por cada restriccion configurada se verifica si se debe tomar en cuenta para restringir el acceso a ciertos usuasrios
    //        foreach (DataRow row in GrupoRestriccion.Rows)
    //        {
    //            //Verifico si el menumsnro esta asociado a la lista de menu en la restriccion
    //            if (!DBNull.Value.Equals(row["menumstrnros"]))
    //            {
    //                string[] arrMenumsnro = Convert.ToString(row["menumstrnros"]).Split(',');

    //                bool Existe = false;
    //                foreach (string valor in arrMenumsnro)
    //                {
    //                    if (valor != "")
    //                    {
    //                        if (Convert.ToInt32(valor) == menumsnro)
    //                        {
    //                            Existe = true;
    //                            break;
    //                        }
    //                    }
    //                }

    //                if (!Existe)
    //                    return 1;

    //                /*
    //                  if (!arrMenumsnro.Contains("0" + Convert.ToString(menumsnro)))                    
    //                      return true;
    //                  */

    //            }


    //            //Verifico si estoy en el rango adecuado para realizar el control         
    //            if (!GrupoRestringido(Convert.ToInt32(row["frectipnro"]), row))
    //            {
    //                return 1;
    //            }


    //            //Verifico el tipo de restriccion
    //            if (!DBNull.Value.Equals(row["frectipnro"]))
    //            {
    //                //Controlo si los perfiles del usuario estan dentro del grupo y si ademas esta restringido segun el tipo de restriccion                                    
    //                if (Utils.Habilitado(ListaPerfUsr, Convert.ToString(row["listperfnro"])))
    //                    return -1;
    //                else
    //                    return 0;
    //            }
    //        }

    //        //Si no pudo controlar, se asume que no tiene ninguna restriccion para ver el menu
    //        return 1;
    //    }


    //    protected bool GrupoRestringido(int tipo, DataRow row)
    //    {
    //        bool Habilitado = false;
    //        bool HoraHabilitada = ((DateTime.Now >= Convert.ToDateTime(row["schedhora"])) && (DateTime.Now <= Convert.ToDateTime(row["schedhora2"])));

    //        switch (tipo)
    //        {
    //            case 1://Diariamente                     
    //                Habilitado = HoraHabilitada;
    //                break;
    //            case 2://Semanalmente
    //                //1 = Domingo // 2 = Lunes // 3 = Martes // 4 = Miercoles // 5 = Jueves // 6 = Viernes // 7 = Sabado  
    //                int numero_dia = Convert.ToInt32(DateTime.Today.DayOfWeek) + 1;
    //                Habilitado = (Convert.ToInt32(row["alesch_frecrep"]) == numero_dia) && (HoraHabilitada);
    //                break;
    //            case 3://Mensualmente                  
    //                if (!DBNull.Value.Equals(row["diassel"]))
    //                {
    //                    string[] arrDias = Convert.ToString(row["diassel"]).Split(',');
    //                    Habilitado = (arrDias.Contains(Convert.ToString(DateTime.Today.Day))) && (HoraHabilitada);
    //                }
    //                else //Si no hay dias configurados se adopta el mismo control que Diariamente
    //                    Habilitado = HoraHabilitada;

    //                break;
    //            case 4://Temporalmente
    //                Habilitado = Control_TemporalDias(Convert.ToDateTime(row["alesch_fecini"]), Convert.ToInt32(row["alesch_frecrep"])) && (HoraHabilitada);
    //                break;
    //            default:

    //                break;
    //        }

    //        return Habilitado;

    //    }

    //    /// <summary>
    //    /// Retorna verdadero si la fecha actual cae dentro de la configuracion temporal
    //    /// </summary>
    //    /// <param name="FechaInicio"></param>
    //    /// <param name="Incremento_Cada"></param>
    //    /// <returns></returns>
    //    protected bool Control_TemporalDias(DateTime FechaInicio, int Incremento_Cada)
    //    {

    //        bool Salida = true;
    //        String fecha = String.Format("{0:dd/MM/yyyy}", DateTime.Now.Date);
    //        DateTime fechaActual = Convert.ToDateTime(fecha);
    //        //String fechaActual = Entities.Fecha.cambiaFecha(DateTime.Today.ToString("dd/MM/yyyy"), cc.get_TipoBase(Utils.SessionBaseID));
    //        DateTime fechaControl = Convert.ToDateTime(String.Format("{0:dd/MM/yyyy}", FechaInicio));//FechaInicio;            
    //        int comparacion;
    //        bool SeguirControl = true;

    //        while (SeguirControl)//Incremento la fecha de inicio hasta que sea mayor o igual que la fecha actual
    //        {
    //            fechaControl = fechaControl.AddDays(Incremento_Cada);
    //            //comparacion = DateTime.Compare(fechaControl,fechaActual);
    //            comparacion = DateTime.Compare(fechaActual, fechaControl);
    //            if (comparacion == 0)//Si las fechas son iguales
    //            {
    //                SeguirControl = false;
    //                Salida = true;
    //            }
    //            else
    //                if (comparacion < 0)//Si ya se paso la suma quiere decir que el dia actual no cae dentro del espacio temporal
    //                {
    //                    SeguirControl = false;
    //                    Salida = false;
    //                }
    //        }

    //        return Salida;
    //    }


    //    /// <summary>
    //    /// Retorna una lista de menuname, de los modulos habilitados para el usuario
    //    /// </summary>
    //    /// <returns></returns>
    //    public List<string> get_ModulosHabilitados()
    //    //public String get_ModulosHabilitados()
    //    {
    //        List<string> Salida = new List<string>();
    //        //  string Salida = "";
    //        Consultas cc = new Consultas();
    //        string sql = " ";
    //        sql = " SELECT menuname,menuaccess ";
    //        sql += " FROM menumstr ";
    //        sql += " WHERE     (parent = 'rhpro') AND (menuactivo = - 1) ";

    //        DataTable dt = cc.get_DataTable(sql, Utils.SessionBaseID);
    //        Usuarios Usr = new Usuarios();
    //        List<String> ListaPerfUsr = Usr.getPerfilesUsuario(Utils.SessionUserName); //(List<String>)System.Web.HttpContext.Current.Session["RHPRO_ListaPerfUsr"];
    //        foreach (DataRow dr in dt.Rows)
    //        {
    //            if (Utils.Habilitado(ListaPerfUsr, Convert.ToString(dr["menuaccess"])))
    //            {/*
    //                if (Salida == "")
    //                    Salida = Convert.ToString(dr["menuname"]).ToUpper();
    //                else
    //                    Salida += "," + Convert.ToString(dr["menuname"]).ToUpper();
    //                */
    //                Salida.Add(Convert.ToString(dr["menuname"]).ToUpper());
    //            }
    //        }

    //        return Salida;
    //    }

     }
}
