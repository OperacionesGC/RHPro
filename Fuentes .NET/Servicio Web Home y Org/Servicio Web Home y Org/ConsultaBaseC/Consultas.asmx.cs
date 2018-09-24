using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Web;
using System.Web.Services;
using System.Web.Services.Protocols;
using System.Data.OleDb;
using System.Configuration;

namespace ConsultaBaseC
{
    /// <summary>
    /// Descripción breve de Consultas
    /// </summary>
    [WebService(Namespace = "http://rhpro.com.ar/")]
    [WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
    [ToolboxItem(false)]
    // Para permitir que se llame a este servicio web desde un script, usando ASP.NET AJAX, quite la marca de comentario de la línea siguiente. 
    // [System.Web.Script.Services.ScriptService]
    public class Consultas : System.Web.Services.WebService
    {        
        private OleDbDataAdapter da;

        //-------------------------------------------------------------------------------------
        //-------------------------------------------------------------------------------------
        [WebMethod(Description = "Devuelve la version de Servicio Web")]
        public string VersionWS()
        {
            string Salida = "";

            /*
            Salida = "Version 1.01: Se agrego versionado. Se realizo un Merge del WS del organigrama y ";
            Salida += " el home";
            */

            //Salida = "Version 1.02: LDAP para la caja.";

            //Salida = "Version 1.03: Servicio Web Loguin dos nuevos parametros de salida, lenguaje y MaxEmpl";

            //Salida = "Version 1.04: Tipo de Base SQL u ORA para los cambiafecha";

            //Salida = "Version 1.05: Error en metodo Login cuando traia de la politica la cantidad de intentos fallidos.";

            //Salida = "Version 1.06: Se modifico MRU para que aplique la seguridad por usuarios.";

            //Salida = "Version 1.07: Se modifico DAL para que funcione con ORA. Alter de Schema y busca Schema en Web.config";

            //Salida = "Version 1.08: Se agrego metodo EstadoPostulante y TablaPlana para interfaz inteligente";

            //Salida = "Version 1.09: Se adecuo el servicio para cambiar el pass en base ORA";

            //Salida = "Version 1.10: En el metodo TablaPlana se agrego tipodoc";
            Salida = "Version 1.0.1.0: JPB - 06/11/2012 - Se soluciono bug en metodo MRU cuando en la action venia un valor vacio";
            return Salida;
        }

        //-------------------------------------------------------------------------------------
        //-------------------------------------------------------------------------------------
        
        [WebMethod(Description = "(HOME) Devuelve la version del sistema.")]
        public string Version(string Base, string Idioma)
        {
            string cn = DAL.constr(Base);
            string sql;
            DataSet ds = new DataSet();
            string Salida = "";

            sql = "SELECT sisnom from sistema";
            
            da = new OleDbDataAdapter(sql, cn);

            try
            {
                da.Fill(ds);
            }
            catch(Exception ex)
            {
                throw ex;
            }

            if (ds.Tables[0].Rows.Count > 0)
            {
                if (ds.Tables[0].Rows[0].ItemArray[0] != DBNull.Value)
                {
                    Salida = Convert.ToString(ds.Tables[0].Rows[0].ItemArray[0].ToString());
                }
            }

            return Salida;
        }

        //-------------------------------------------------------------------------------------
        //-------------------------------------------------------------------------------------
        
        [WebMethod(Description = "(HOME) Devuelve el Patch del sistema.")]
        public string Patch(string Base, string Idioma)
        {
            string cn = DAL.constr(Base);
            string sql;
            DataSet ds = new DataSet();
            string Salida = "";

            sql = "SELECT patchdesabr FROM patch ORDER BY patchnro DESC ";

            da = new OleDbDataAdapter(sql, cn);

            try
            {
                da.Fill(ds,0,1,"Resultado");
            }
            catch(Exception ex)
            {
                throw ex;
            }

            if (ds.Tables[0].Rows.Count > 0)
            {
                if (ds.Tables[0].Rows[0].ItemArray[0] != DBNull.Value)
                {
                    Salida = Convert.ToString(ds.Tables[0].Rows[0].ItemArray[0].ToString());
                }
            }

            return Salida;
        }

        //-------------------------------------------------------------------------------------
        //-------------------------------------------------------------------------------------
        
        [WebMethod(Description = "(HOME) Devuelve los links de interes para el banner segun el usuario.")]
        public DataTable Link(string Usuario, string Base, string Idioma)
        {
            string cn = DAL.constr(Base);
            string sql;
            DataSet ds = new DataSet();

            if (Usuario.Trim().Length == 0)
            {
                Usuario = "usr_logout";
            }

            sql = "SELECT hlinktitulo, hlinkpagina ";
            sql = sql + "FROM user_link ";
            sql = sql + "INNER JOIN home_link ON home_link.hlinknro = user_link.hlinknro ";
            sql = sql + "WHERE UPPER(iduser) = '" + Usuario.ToUpper() + "' ";
            sql = sql + "ORDER BY home_link.hlinknro ";
            
            da = new OleDbDataAdapter(sql, cn);

            try
            {
                da.Fill(ds);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return ds.Tables[0];
        }

        //-------------------------------------------------------------------------------------
        //-------------------------------------------------------------------------------------
        
        [WebMethod(Description = "(HOME) Devuelve un titulo y una descripción de las noticias a mostrar en el banner de comunidad.")]
        public DataTable Mensaje(string Base, string Idioma)
        {
            string cn = DAL.constr(Base);
            string sql;
            DataSet ds = new DataSet();

            sql = "SELECT hmsjtitulo, hmsjcuerpo ";
            sql = sql + "FROM home_mensaje ";
            sql = sql + "WHERE home_mensaje.hmsjactivo = -1 ";
            sql = sql + "ORDER BY home_mensaje.hmsjnro ";
            
            da = new OleDbDataAdapter(sql, cn);

            try
            {
                da.Fill(ds);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return ds.Tables[0];
        }

        //-------------------------------------------------------------------------------------
        //-------------------------------------------------------------------------------------
        
        [WebMethod(Description = "(HOME) Devuelve un titulo, una imagen (guardada en el directorio shared\\images\\ de rhpro) y una desc extendida a mostrar en el banner.")]
        public DataTable Banner(string Base, string Idioma)
        {
            string cn = DAL.constr(Base);
            string sql;
            DataSet ds = new DataSet();

            sql = "SELECT hbandesc, hbanimage, hbandescext ";
            sql = sql + "FROM home_banner ";
            sql = sql + "WHERE home_banner.hbanactivo = -1 ";
            sql = sql + "ORDER BY home_banner.hbannro ";

            da = new OleDbDataAdapter(sql, cn);

            try
            {
                da.Fill(ds);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return ds.Tables[0];
        }

        //-------------------------------------------------------------------------------------
        //-------------------------------------------------------------------------------------
        
        [WebMethod(Description = "(HOME) Devuelve los modulos segun el usuario.")]
        public DataTable Modulos(string Usuario, string Base, string Idioma)
        {
            string cn = DAL.constr(Base);
            string sql;
            DataSet ds = new DataSet();
            DataSet dsAux = new DataSet();
            DataSet dsAux2 = new DataSet();
            DataTable tablaAux;
            string Access = "";
            string[] arrAccess;
            string[] arrPerfUser;
            string listaPerfUser = "";
            bool Hay = false;
            DataColumn Columna;
            DataRow filaAux;

            if (Usuario.Trim().Length == 0)
            {
                sql = "SELECT menudesabr,menudetalle,'' action,menuobjetivo,menubeneficio,linkmanual,linkdvd ";
                sql = sql + "FROM menumstr ";
                sql = sql + "WHERE menuraiz = 74 ";
                sql = sql + "AND menuactivo = -1 ";
                sql = sql + "ORDER BY menudesabr ";
                
                da = new OleDbDataAdapter(sql, cn);

                try
                {
                    da.Fill(ds);
                }
                catch (Exception ex)
                {
                    throw ex;
                }

                return ds.Tables[0];
            }
            else
            {
                //Busco el perfil del usuario

                sql = "SELECT listperfnro ";
                sql = sql + "FROM user_perfil ";
                sql = sql + "WHERE UPPER(user_perfil.iduser) = '" + Usuario.ToUpper() + "' ";
                
                da = new OleDbDataAdapter(sql, cn);

                try
                {
                    da.Fill(dsAux);
                }
                catch (Exception ex)
                {
                    throw ex;
                }

                if (dsAux.Tables[0].Rows.Count > 0)
                {
                    listaPerfUser = Convert.ToString(dsAux.Tables[0].Rows[0].ItemArray[0]);

                    //Busco todos los menu que tienen al perfil

                    sql = "SELECT menuaccess, ";
                    sql = sql + "menudesabr, ";
                    sql = sql + "menudetalle, ";
                    //sql = sql + "'abrirVentana(' + CHAR(39) + action + CHAR(39) + ','''',670,520)' action, ";
                    sql = sql + "'abrirVentana(' action1, ";
                    sql = sql + "action action2, ";
                    sql = sql + "',670,520)' action3, ";
                    sql = sql + "menuobjetivo, ";
                    sql = sql + "menubeneficio, ";
                    sql = sql + "linkmanual, ";
                    sql = sql + "linkdvd ";
                    sql = sql + "FROM menumstr ";
                    sql = sql + "WHERE menuraiz = 74 ";
                    sql = sql + "AND menuactivo = -1 ";
                    sql = sql + "AND menumstr.action <> '#' ";
                    //sql = sql + "AND menumstr.action <> '' ";
                    sql = sql + "AND menumstr.action IS NOT NULL ";
                    sql = sql + "ORDER BY menudesabr ";

                    da = new OleDbDataAdapter(sql, cn);

                    try
                    {
                        da.Fill(dsAux2);
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }

                    //Creo la tabla de salida

                    DataTable tablaSalida = new DataTable("table");
                    Columna = new DataColumn();
                    Columna.DataType = System.Type.GetType("System.String");
                    Columna.ColumnName = "menudesabr";
                    Columna.AutoIncrement = false;
                    Columna.Unique = false;
                    tablaSalida.Columns.Add(Columna);
                    Columna = new DataColumn();
                    Columna.DataType = System.Type.GetType("System.String");
                    Columna.ColumnName = "menudetalle";
                    Columna.AutoIncrement = false;
                    Columna.Unique = false;
                    tablaSalida.Columns.Add(Columna);
                    Columna = new DataColumn();
                    Columna.DataType = System.Type.GetType("System.String");
                    Columna.ColumnName = "action";
                    Columna.AutoIncrement = false;
                    Columna.Unique = false;
                    tablaSalida.Columns.Add(Columna);
                    Columna = new DataColumn();
                    Columna.DataType = System.Type.GetType("System.String");
                    Columna.ColumnName = "menuobjetivo";
                    Columna.AutoIncrement = false;
                    Columna.Unique = false;
                    tablaSalida.Columns.Add(Columna);
                    Columna = new DataColumn();
                    Columna.DataType = System.Type.GetType("System.String");
                    Columna.ColumnName = "menubeneficio";
                    Columna.AutoIncrement = false;
                    Columna.Unique = false;
                    tablaSalida.Columns.Add(Columna);
                    Columna = new DataColumn();
                    Columna.DataType = System.Type.GetType("System.String");
                    Columna.ColumnName = "linkmanual";
                    Columna.AutoIncrement = false;
                    Columna.Unique = false;
                    tablaSalida.Columns.Add(Columna);
                    Columna = new DataColumn();
                    Columna.DataType = System.Type.GetType("System.String");
                    Columna.ColumnName = "linkdvd";
                    Columna.AutoIncrement = false;
                    Columna.Unique = false;
                    tablaSalida.Columns.Add(Columna);

                    //Ciclo por cada modulo

                    if (dsAux2.Tables[0].Rows.Count > 0)
                    {
                        tablaAux = dsAux2.Tables[0];

                        foreach (DataRow fila in tablaAux.Rows)
                        {
                            //Copio todas las filas menos la de access que depende del perfil
                            
                            filaAux = tablaSalida.NewRow();
                            filaAux["menudesabr"] = fila["menudesabr"].ToString();
                            filaAux["menudetalle"] = fila["menudetalle"].ToString();
                            filaAux["menuobjetivo"] = fila["menuobjetivo"].ToString();
                            filaAux["menubeneficio"] = fila["menubeneficio"].ToString();
                            filaAux["linkmanual"] = fila["linkmanual"].ToString();
                            filaAux["linkdvd"] = fila["linkdvd"].ToString();

                            Access = Convert.ToString(fila["menuaccess"].ToString());
                            Hay = false;

                            //Por cada perfil del usuario

                            arrPerfUser = listaPerfUser.Split(new char[] { ',' });

                            foreach (string PerfUser in arrPerfUser)
                            { 
                                arrAccess = Access.Split(new char[] { ',' });
                                
                                //Por cada perfil del menu
                                foreach (string perfil in arrAccess)
                                {
                                    if ((perfil == "*") || (perfil.ToUpper() == PerfUser.ToUpper()))
                                    {
                                        Hay = true;
                                        //Salgo del ciclo de perfiles asociados al usuario
                                        break;
                                    }
                                }
                                if (Hay)
                                    //Salgo del ciclo de perfiles del usuario
                                    break;
                            }

                            if (Hay)
                                filaAux["action"] = fila["action1"].ToString() + "'" + fila["action2"].ToString() + "',''" + fila["action3"].ToString();
                            else
                                filaAux["action"] = "";

                            //Inserto la fila en la tabla de salida

                            tablaSalida.Rows.Add(filaAux);
                        }

                        return tablaSalida;
                    }
                    else
                    {
                        return tablaSalida;
                    }
                }
                else
                {
                    //El usuario no existe o no tiene perfil

                    sql = "SELECT menudesabr,menudetalle,'' action,menuobjetivo,menubeneficio,linkmanual,linkdvd ";
                    sql = sql + "FROM menumstr where menuraiz = 74 ";
                    sql = sql + "AND menuactivo = -1 ";
                    sql = sql + "ORDER BY menudesabr ";

                    da = new OleDbDataAdapter(sql, cn);

                    try
                    {
                        da.Fill(ds);
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }

                    return ds.Tables[0];
                }
            }
        }

        //-------------------------------------------------------------------------------------
        //-------------------------------------------------------------------------------------
        
        [WebMethod(Description = "(HOME) Devuelve los menu mas utilizados por el usuario.")]
        public DataTable MRU(string Usuario, int Cant, string Base, string Idioma)
        {
            string cn = DAL.constr(Base);
            string sql;
            string menuAccion = "";
            string cadena1 = "";
            string cadena2 = "";
            string[] arrPerfUser;
            string listaPerfUser = "";
            string[] arrAccess;
            string Access = "";
            bool Hay = false;

            DataRow filaAux;
            DataSet ds = new DataSet();
            DataSet dsAux = new DataSet();
            
            //Creo la tabla de salida

            DataTable tablaSalida = new DataTable("table");

            DataColumn Columna = new DataColumn();

            Columna.DataType = System.Type.GetType("System.String");
            Columna.ColumnName = "menuname";
            Columna.AutoIncrement = false;
            Columna.Unique = false;
            tablaSalida.Columns.Add(Columna);

            Columna = new DataColumn();
            Columna.DataType = System.Type.GetType("System.String");
            Columna.ColumnName = "action";
            Columna.AutoIncrement = false;
            Columna.Unique = false;
            tablaSalida.Columns.Add(Columna);

            Columna = new DataColumn();
            Columna.DataType = System.Type.GetType("System.String");
            Columna.ColumnName = "raiz";
            Columna.AutoIncrement = false;
            Columna.Unique = false;
            tablaSalida.Columns.Add(Columna);

            if (Usuario.Trim().Length == 0)
            {
                Usuario = "usr_logout";
            }

            sql = "SELECT menumstr.menuname, menumstr.action, menuraiz.menunombre raiz, menuraiz.menudir, menumstr.menuaccess FROM mru ";
            sql = sql + "INNER JOIN menumstr ON menumstr.menumsnro = mru.menumsnro ";
            sql = sql + "INNER JOIN menuraiz ON menuraiz.menunro = mru.menuraiz ";
            sql = sql + "WHERE UPPER(mru.iduser) = '" + Usuario.ToUpper() + "' ";
            sql = sql + "ORDER BY mrufecha DESC, mruhora DESC ";

            da = new OleDbDataAdapter(sql, cn);

            try
            {
                da.Fill(ds, 0, Cant, "Resultado");
            }
            catch (Exception ex)
            {
                throw ex;
            }

            if (ds.Tables[0].Rows.Count > 0)
            {
                DataTable tablaAux = ds.Tables[0];

                //Busco el perfil del usuario
                sql = "SELECT listperfnro ";
                sql = sql + "FROM user_perfil ";
                sql = sql + "WHERE UPPER(user_perfil.iduser) = '" + Usuario.ToUpper() + "' ";
                
                da = new OleDbDataAdapter(sql, cn);

                try
                {
                    da.Fill(dsAux);
                }
                catch (Exception ex)
                {
                    throw ex;
                }

                if (dsAux.Tables[0].Rows.Count > 0)
                {
                    listaPerfUser = Convert.ToString(dsAux.Tables[0].Rows[0].ItemArray[0]);

                    //por cada menu
                    foreach (DataRow fila in tablaAux.Rows)
                    {

                        arrPerfUser = listaPerfUser.Split(new char[] { ',' });

                        Access = Convert.ToString(fila["menuaccess"].ToString());
                        Hay = false;

                        foreach (string PerfUser in arrPerfUser)
                        {
                            arrAccess = Access.Split(new char[] { ',' });

                            //Por cada perfil del menu
                            foreach (string perfil in arrAccess)
                            {
                                if ((perfil == "*") || (perfil.ToUpper() == PerfUser.ToUpper()))
                                {
                                    Hay = true;
                                    //Salgo del ciclo de perfiles asociados al usuario
                                    break;
                                }
                            }
                            if (Hay)
                                //Salgo del ciclo de perfiles del usuario
                                break;
                        }

                       

                        if (Hay)
                        {
                                menuAccion = fila["action"].ToString();
                                menuAccion = menuAccion.Replace("Javascript:", "");
                                menuAccion = menuAccion.Replace("JavaScript:", "");
                                menuAccion = menuAccion.Replace("javaScript:", "");
                                menuAccion = menuAccion.Replace("javascript:", "");

                                if (menuAccion.IndexOf("('../", StringComparison.CurrentCultureIgnoreCase) != -1)
                                {
                                    cadena1 = menuAccion.Substring(0, menuAccion.IndexOf("../", StringComparison.CurrentCultureIgnoreCase));
                                    cadena2 = menuAccion.Substring(menuAccion.IndexOf("../", StringComparison.CurrentCultureIgnoreCase) + 3, menuAccion.Length - menuAccion.IndexOf("../", StringComparison.CurrentCultureIgnoreCase) - 3);
                                }
                                else
                                {
                                    //cadena1 = menuAccion.Substring(0, menuAccion.IndexOf("('", StringComparison.CurrentCultureIgnoreCase) + 2);
                                    //cadena2 = menuAccion.Substring(menuAccion.IndexOf("('", StringComparison.CurrentCultureIgnoreCase) + 2, menuAccion.Length - menuAccion.IndexOf("('", StringComparison.CurrentCultureIgnoreCase) - 2);
                                    //cadena2 = fila["menudir"].ToString() + "/" + cadena2;

                                    //JPB:CAS-17231 Se controla que la accion venga con algun valor
                                    if (menuAccion.Length > 0)
                                    {
                                        cadena1 = menuAccion.Substring(0, menuAccion.IndexOf("('", StringComparison.CurrentCultureIgnoreCase) + 2);
                                        cadena2 = menuAccion.Substring(menuAccion.IndexOf("('", StringComparison.CurrentCultureIgnoreCase) + 2, menuAccion.Length - menuAccion.IndexOf("('", StringComparison.CurrentCultureIgnoreCase) - 2);
                                        cadena2 = fila["menudir"].ToString() + "/" + cadena2;
                                    }
                                    else
                                    { cadena1 = "";
                                      cadena2 = "";
                                    }
                                    

                                }
                               

                            filaAux = tablaSalida.NewRow();
                            filaAux["menuname"] = fila["menuname"].ToString();
                            filaAux["action"] = cadena1 + cadena2; ;
                            filaAux["raiz"] = fila["raiz"].ToString();
                            tablaSalida.Rows.Add(filaAux);
                        }
                    }
                }
            }

            return tablaSalida;
        }

        //-------------------------------------------------------------------------------------
        //-------------------------------------------------------------------------------------
        
        [WebMethod(Description = "(HOME) Devuelve la pagina a mostrar en la parte inferior de la pantalla.")]
        public DataTable PagPie(string Usuario, string Base, string Idioma)
        {
            string cn = DAL.constr(Base);
            string sql;
            DataSet ds = new DataSet();

            if (Usuario.Trim().Length == 0)
            {
                Usuario = "usr_logout";
            }

            sql = "SELECT hpptitulo, hpppagina ";
            sql = sql + "FROM home_pagpie ";
            sql = sql + "WHERE UPPER(iduser) = '" + Usuario.ToUpper() + "' ";
            sql = sql + "AND hpppactivo = -1 ";

            da = new OleDbDataAdapter(sql, cn);

            try
            {
                da.Fill(ds);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return ds.Tables[0];
        }

        //-------------------------------------------------------------------------------------
        //-------------------------------------------------------------------------------------
        
        [WebMethod(Description = "(HOME) Devuelve las bases para el combo. Formato de salida es un string separado por coma NombreBase,NroConexion,SegIntegrada(-1),ValorDefault(-1)")]
        public DataTable comboBase()
        {
            return DAL.Bases();
        }

        //-------------------------------------------------------------------------------------
        //-------------------------------------------------------------------------------------
        
        [WebMethod(Description = "(HOME) Devuelve una tabla con Nombre de Modulo, Nombre de Menu y Accion de Menu segun la palabra buscada y usuario.")]
        public DataTable Search(string Usuario, string Palabra, string Base, string Idioma)
        {
            string cn = DAL.constr(Base);
            string sql;
            string Access = "";
            string[] arrAccess;
            bool Ingresa = false;
            string menuAccion = "";
            DataSet ds = new DataSet();
            DataSet dsPerfil = new DataSet();
            string[] arrPerfUser;
            string listaPerfUser = "";
            DataRow filaAux;

            //Creo la tabla de salida
            DataTable tablaSalida = new DataTable("table");

            DataColumn Columna = new DataColumn();

            Columna.DataType = System.Type.GetType("System.String");
            Columna.ColumnName = "Modulo";
            Columna.AutoIncrement = false;
            Columna.Unique = false;
            tablaSalida.Columns.Add(Columna);

            Columna = new DataColumn();
            Columna.DataType = System.Type.GetType("System.String");
            Columna.ColumnName = "DescrMenu";
            Columna.AutoIncrement = false;
            Columna.Unique = false;
            tablaSalida.Columns.Add(Columna);

            Columna = new DataColumn();
            Columna.DataType = System.Type.GetType("System.String");
            Columna.ColumnName = "Accion";
            Columna.AutoIncrement = false;
            Columna.Unique = false;
            tablaSalida.Columns.Add(Columna);

            Columna = new DataColumn();
            Columna.DataType = System.Type.GetType("System.String");
            Columna.ColumnName = "DescrExt";
            Columna.AutoIncrement = false;
            Columna.Unique = false;
            tablaSalida.Columns.Add(Columna);

            if (Usuario.Trim().Length != 0)
            {
                //Busco el perfil del usuario

                sql = "SELECT listperfnro ";
                sql = sql + "FROM user_perfil ";
                sql = sql + "WHERE UPPER(user_perfil.iduser) = '" + Usuario.ToUpper() + "' ";
                
                da = new OleDbDataAdapter(sql, cn);

                try
                {
                    da.Fill(dsPerfil);
                }
                catch (Exception ex)
                {
                    throw ex;
                }

                if (dsPerfil.Tables[0].Rows.Count > 0)
                {
                    listaPerfUser = Convert.ToString(dsPerfil.Tables[0].Rows[0].ItemArray[0]);

                    //Busco todos los menu que tienen la palabra buscada

                    sql = "SELECT menunombre, menuname, action, menumstr.menuaccess, menuraiz.menudir, menumstr.menudesabr ";
                    sql = sql + "FROM menumstr ";
                    sql = sql + "INNER JOIN menuraiz ON menuraiz.menunro = menumstr.menuraiz ";
                    sql = sql + "WHERE menuname LIKE '%" + Palabra + "%' ";
                    sql = sql + "AND menumstr.menuraiz <> 74 ";
                    sql = sql + "AND menumstr.menuraiz <> 73 ";
                    sql = sql + "AND menumstr.menuraiz <> 81 ";
                    sql = sql + "AND menumstr.action <> '#' ";
                    sql = sql + "AND menumstr.action <> '' ";
                    sql = sql + "AND menumstr.action IS NOT NULL ";
                    sql = sql + "ORDER BY menunombre, menuname ";

                    da = new OleDbDataAdapter(sql, cn);

                    try
                    {
                        da.Fill(ds);
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }

                    if (ds.Tables[0].Rows.Count > 0)
                    {
                        DataTable tablaAux = ds.Tables[0];
                        //Por cada menu encontrado verifico si el usuario puede verlo
                        
                        foreach (DataRow fila in tablaAux.Rows)
                        {
                            Access = Convert.ToString(fila["menuaccess"].ToString());
                            Ingresa = false;

                            //Por cada perfil del usuario
                            arrPerfUser = listaPerfUser.Split(new char[] { ',' });

                            foreach (string PerfUser in arrPerfUser)
                            {
                                arrAccess = Access.Split(new char[] { ',' });
                                //Por cada perfil del menu
                                foreach (string perfil in arrAccess)
                                {
                                    if ((perfil == "*") || (perfil.ToUpper() == PerfUser.ToUpper()))
                                    {
                                        Ingresa = true;
                                        break;
                                    }
                                }

                                if (Ingresa)
                                    //Salgo del ciclo de perfiles del usuario
                                    break;
                            }

                            if (Ingresa)
                            {
                                bool Crear = false;
                                string cadena1 = "";
                                string cadena2 = "";
                                string cadena3 = "";

                                if (fila["action"] != DBNull.Value)
                                {
                                    menuAccion = fila["action"].ToString();
                                    menuAccion = menuAccion.Replace("Javascript:", "");
                                    menuAccion = menuAccion.Replace("JavaScript:", "");
                                    menuAccion = menuAccion.Replace("javaScript:", "");
                                    menuAccion = menuAccion.Replace("javascript:", "");

                                    if ((menuAccion.IndexOf("../", StringComparison.CurrentCultureIgnoreCase) != -1)
                                        && (menuAccion.LastIndexOf("/", StringComparison.CurrentCultureIgnoreCase) != -1)
                                        && (menuAccion.IndexOf("../", StringComparison.CurrentCultureIgnoreCase) != menuAccion.LastIndexOf("/", StringComparison.CurrentCultureIgnoreCase))
                                        )
                                    {
                                        //Debo acomodar la accion de menu porque accede a otro directorio (saco ../)
                                        cadena1 = menuAccion.Substring(0, menuAccion.IndexOf("../", StringComparison.CurrentCultureIgnoreCase));
                                        cadena2 = "";
                                        cadena3 = menuAccion.Substring(menuAccion.IndexOf("../", StringComparison.CurrentCultureIgnoreCase) + 3, menuAccion.Length - menuAccion.IndexOf("../", StringComparison.CurrentCultureIgnoreCase) - 3);
                                        Crear = true;
                                    }
                                    else
                                    {
                                        if (menuAccion.IndexOf("'", StringComparison.CurrentCultureIgnoreCase) != -1)
                                        {
                                            //Debo acomodar la accion de menu para concatenarle en menu raiz
                                            cadena1 = menuAccion.Substring(0, menuAccion.IndexOf("'", StringComparison.CurrentCultureIgnoreCase) + 1);
                                            cadena2 = fila["menudir"].ToString() + "/";
                                            cadena3 = menuAccion.Substring(menuAccion.IndexOf("'", StringComparison.CurrentCultureIgnoreCase) + 1, menuAccion.Length - menuAccion.IndexOf("'", StringComparison.CurrentCultureIgnoreCase) - 1);
                                            Crear = true;
                                        }
                                    }

                                    //Agrego la fila a la tabla
                    
                                    if (Crear)
                                    {
                                        menuAccion = cadena1 + cadena2 + cadena3;
                                        filaAux = tablaSalida.NewRow();
                                        filaAux["Modulo"] = fila["menunombre"].ToString();
                                        filaAux["DescrMenu"] = fila["menuname"].ToString();
                                        filaAux["Accion"] = menuAccion;
                                        filaAux["DescrExt"] = fila["menudesabr"].ToString();
                                        tablaSalida.Rows.Add(filaAux);
                                    }
                                }
                            }
                        }
                    }
                }
            }

            return tablaSalida;
        }

        //-------------------------------------------------------------------------------------
        //-------------------------------------------------------------------------------------
        
        [WebMethod(Description = "(HOME) Retorna si el usuario ingresa, un mensaje de error y si el mismo debe cambiar la password.")]
        public DataTable Login(string Usuario, string Pass, string SegNt, string Base, string Idioma)
        {
            bool Ingresa = true;
            string Mess = "";
            bool cambiaPass = false;
            long polNro = 0;
	 	    long passExpiraDias = 0;
	 	    long passCambDias = 0;
	 	    long passIntFallidos  = 0;
            long passDiasLog = 0;
            bool usrPassCambiar = false;
            long diffDias = 0;
            DateTime hLogFecini = DateTime.Today;
            DateTime hPassFecini = DateTime.Today;

            //--------------------------------------------------------------------------------
            //Creo la tabla de salida para devolver los datos
            //--------------------------------------------------------------------------------
            
            DataTable tablaSalida = new DataTable("table");
            DataColumn Columna = new DataColumn();
            Columna.DataType = System.Type.GetType("System.Boolean");
            Columna.ColumnName = "Ingresa";
            Columna.AutoIncrement = false;
            Columna.Unique = false;
            tablaSalida.Columns.Add(Columna);
            Columna = new DataColumn();
            Columna.DataType = System.Type.GetType("System.String");
            Columna.ColumnName = "mensaje";
            Columna.AutoIncrement = false;
            Columna.Unique = false;
            tablaSalida.Columns.Add(Columna);
            Columna = new DataColumn();
            Columna.DataType = System.Type.GetType("System.Boolean");
            Columna.ColumnName = "CambiarPass";
            Columna.AutoIncrement = false;
            Columna.Unique = false;
            tablaSalida.Columns.Add(Columna);
            Columna = new DataColumn();
            Columna.DataType = System.Type.GetType("System.String");
            Columna.ColumnName = "lenguaje";
            Columna.AutoIncrement = false;
            Columna.Unique = false;
            tablaSalida.Columns.Add(Columna);
            Columna = new DataColumn();
            Columna.DataType = System.Type.GetType("System.String");
            Columna.ColumnName = "MaxEmpl";
            Columna.AutoIncrement = false;
            Columna.Unique = false;
            tablaSalida.Columns.Add(Columna);

            //--------------------------------------------------------------------------------
            //Si se utiliza validación por servicio de directorio LDAP, verifico la existencia 
            //del usuario en el mismo.
            //
            //Tener en cuenta que No se debe tener habilitada la opción de
            //Seguridad Integrada (SegNt != TrueValue) para que esta modalidad funcione correctamente.
            //--------------------------------------------------------------------------------

            string LDAP_UseAuthentication = ConfigurationManager.AppSettings["LDAP_UseAuthentication"].ToString().ToLower().Trim();

            if (SegNt != "TrueValue" && LDAP_UseAuthentication == "true") //Si se debe validar el usuario por LDAP...
            {
                LDAP ldap = new LDAP();

                if (!ldap.usuarioValido(Usuario, Pass)) //Si el usuario no es válido en LDAP...
                {
                    Mess = DAL.Error(1, Idioma);
                    Ingresa = false;
                }
                else //Si el usuario es válido...
                {
                    //Actualizo la password del usuario (para que coincida con la del servidor LDAP).
                    
                    //Tener en cuenta que no debe haber políticas de password definidas
                    //para que esta modalidad funcione correctamente.

                    Mess = this.CambiarPass(Usuario, Pass, Pass, Pass, Base, Idioma);

                    if(Mess!="")
                        Ingresa=false;
                }
            }

            //--------------------------------------------------------------------------------
            //Pruebo la conexion con los datos del usuario
            //--------------------------------------------------------------------------------

            if (Ingresa)
            {

                OleDbConnection connUsu = new OleDbConnection(DAL.constrUsu(Usuario, Encriptar.Encrypt(DAL.EncrKy(), Pass), SegNt, Base));

                try
                {
                    connUsu.Open();
                }
                catch (Exception ex)
                {
                    //Mess = "Usuario o constraseña incorrecta.";
                    Mess = DAL.Error(1, Idioma);
                    Ingresa = false;
                }

                connUsu.Close();
            }

            //--------------------------------------------------------------------------------
            //Verifico si el usuario es valido
            //--------------------------------------------------------------------------------
            
            if ((Ingresa) && (!Password.usuarioValido(Usuario, Base)))
            {
                //Mess = "Usuario no válido.";
                Mess = DAL.Error(2, Idioma);
                Ingresa = false;
            }
            
            //Seguridad Base de Datos
            
            if (Ingresa && SegNt != "TrueValue" && LDAP_UseAuthentication == "false")
            {
                //--------------------------------------------------------------------------------
                //Control de cuenta bloqueada por usuario no por politica
                //--------------------------------------------------------------------------------
                
                if ((Password.ctaBloqueada(Usuario, Base)))
                {
                    //Mess = "Cuenta Bloqueada. Consulte con el administrador.";
                    Mess = DAL.Error(3, Idioma);
                    Ingresa = false;
                }

                //--------------------------------------------------------------------------------
                //Busca politica de cuenta
                //--------------------------------------------------------------------------------
                
                if (Ingresa)
                {
                    polNro = (Password.valorUserPolCuenta(Usuario, "pol_nro", Base).Length != 0) ? Convert.ToInt64(Password.valorUserPolCuenta(Usuario, "pol_nro", Base)) : 0;
                    
                    if (polNro == 0)
                    {
                        //Mess = "No se encontro la politica de cuenta del usuario.";
                        Mess = DAL.Error(4, Idioma);
                        Ingresa = false;
                    }
                }

                //--------------------------------------------------------------------------------
                //Control de la contraseña ingresada contra la almacenada en la base encriptada
                //--------------------------------------------------------------------------------
                
                if (Password.valorHistPass(Usuario, "husrpass", Base) != Encriptar.Encrypt(DAL.EncrKy(), Pass))
                {
                    Ingresa = false;

                    //Control de politica de intentos fallidos
                    passIntFallidos = (Password.valorPolCuenta(polNro, "pass_int_fallidos", Base).Length != 0) ? Convert.ToInt64(Password.valorPolCuenta(polNro, "pass_int_fallidos", Base)) : 0;

                    //Recupero la cantidad de intentos fallidos
                    long intentosFallidos = Password.logueosFallidos(Usuario, Base) + 1;

                    if ((passIntFallidos != 0) && (intentosFallidos >= passIntFallidos))
                    {
                        //Bloqueo las cuentas
                        Password.bloquearCuenta(Usuario, "-1", Base);
                        Password.bajarCuenta(Usuario, Base);
                        //Mess = "Cuenta Bloqueada por intentos fallidos.";
                        Mess = DAL.Error(5, Idioma);
                    }
                    else
                    {
                        Password.actLogFallidos(Usuario, intentosFallidos, Base);
                        //Mess = "Contraseña incorrecta.";
                        Mess = DAL.Error(6, Idioma);
                    }
                }

                //--------------------------------------------------------------------------------
                //Control de cambio de contraseña no por politica sino por usuario
                //--------------------------------------------------------------------------------
                
                if (Ingresa)
                {
                    //Recupero el dato del usuario de cambio de contraseña
                    
                    usrPassCambiar = (Password.valorUserPer(Usuario, "usrpasscambiar", Base).Length != 0) ? (Convert.ToInt64(Password.valorUserPer(Usuario, "usrpasscambiar", Base)) == -1) : false;
                    
                    if (usrPassCambiar)
                    {
                        Ingresa = false;
                        cambiaPass = true;
                        //Mess = "Debe Cambiar su contraseña.";
                        Mess = DAL.Error(7, Idioma);
                    }
                }

                //--------------------------------------------------------------------------------
                //Control Politica Dias sin loguearse
                //--------------------------------------------------------------------------------
                
                if (Ingresa)
                {
                    passDiasLog = (Password.valorPolCuenta(polNro, "pass_dias_log", Base).Length != 0) ? Convert.ToInt64(Password.valorPolCuenta(polNro, "pass_dias_log", Base)) : 0;
                    
                    if (passDiasLog != 0)
                    {
                        //Recupero el ultimo Login
                        hLogFecini = (Password.valorHistLog(Usuario, "hlogfecini", Base).Length != 0) ? Convert.ToDateTime(Password.valorHistLog(Usuario, "hlogfecini", Base)) : DateTime.Today;

                        //Calculo la diferencia al dia de hoy
                        long diasSinLogin = Fecha.DateDiff(DateInterval.Day, hLogFecini, DateTime.Today);

                        //Control si exedio los dias permitidos
                        if (diasSinLogin >= passDiasLog)
                        {
                            Ingresa = false;

                            //Bloqueo las cuentas
                            Password.bloquearCuenta(Usuario, "-1", Base);
                            Password.bajarCuenta(Usuario, Base);
                            //Mess = "Cuenta Bloqueada. Plazo excedido sin loguearse.";
                            Mess = DAL.Error(8, Idioma);
                        }
                    }
                }

                //--------------------------------------------------------------------------------
                //Control Politica expiracion de cuenta
                //--------------------------------------------------------------------------------
                
                if (Ingresa)
                {
                    passExpiraDias = (Password.valorPolCuenta(polNro, "pass_expira_dias", Base).Length != 0) ? Convert.ToInt64(Password.valorPolCuenta(polNro, "pass_expira_dias", Base)) : 0;
                    
                    if (passExpiraDias != 0)
                    {
                        hPassFecini = (Password.valorHistPass(Usuario, "hpassfecini", Base).Length != 0) ? Convert.ToDateTime(Password.valorHistPass(Usuario, "hpassfecini", Base)) : DateTime.Today;
                        diffDias = Fecha.DateDiff(DateInterval.Day, hPassFecini, DateTime.Today);

                        if ((passExpiraDias - 1) < diffDias)
                        {
                            Ingresa = false;

                            //Bloqueo las cuentas
                            Password.bloquearCuenta(Usuario, "-1", Base);
                            Password.bajarCuenta(Usuario, Base);
                            //Mess = "Cuenta Bloqueada. Expiro su contraseña.";
                            Mess = DAL.Error(9, Idioma);
                        }
                    }
                }

                //--------------------------------------------------------------------------------
                //Control Politica cambio de contraseña 
                //--------------------------------------------------------------------------------
                
                if (Ingresa)
                {
                    passCambDias = (Password.valorPolCuenta(polNro, "pass_camb_dias", Base).Length != 0) ? Convert.ToInt64(Password.valorPolCuenta(polNro, "pass_camb_dias", Base)) : 0;
                    
                    if (passCambDias != 0)
                    {
                        hPassFecini = (Password.valorHistPass(Usuario, "hpassfecini", Base).Length != 0) ? Convert.ToDateTime(Password.valorHistPass(Usuario, "hpassfecini", Base)) : DateTime.Today;
                        diffDias = Fecha.DateDiff(DateInterval.Day, hPassFecini, DateTime.Today);

                        if (passCambDias <= diffDias)
                        {
                            Ingresa = false;
                            cambiaPass = true;
                            //Mess = "Debe Cambiar su contraseña.";
                            Mess = DAL.Error(10, Idioma);
                        }
                    }
                }
            }
            
            //--------------------------------------------------------------------------------
            //Registro el login del usuario
            //--------------------------------------------------------------------------------
            
            if (Ingresa)
            { 
                Password.ingresarLogueo(Usuario,Base);
            }

            //Genero la salida

            DataRow fila = tablaSalida.NewRow();
            fila["Ingresa"] = Ingresa;
            fila["mensaje"] = Mess;
            fila["CambiarPass"] = cambiaPass;
            fila["lenguaje"] = Idioma;
            fila["MaxEmpl"] = "100";

            tablaSalida.Rows.Add(fila);

            return tablaSalida;
        }

        [WebMethod(Description = "(HOME) Cambia la contraseña del Usuario con contraseña anterior PassOld a la nueva PassNew. Devuelve string que si es vacio entonces cambio password ok, sino devuelve el error.")]
        public string CambiarPass(string Usuario, string PassOld, string PassNew, string PassConfirm, string Base, string Idioma)
        {
            bool Termino = false;
            string Mess = "";
            long passHistoria = 0;

            //--------------------------------------------------------------------------------
            //Control de coincidencia con confirmacion
            //--------------------------------------------------------------------------------
            
            if (!Termino && (PassConfirm != PassNew))
            {
                //Mess = "La confirmación de la contraseña no es coincidente.";
                Mess = DAL.Error(11, Idioma);
                Termino = true;
            }

            //--------------------------------------------------------------------------------
            //Control de usuario valido
            //--------------------------------------------------------------------------------
            
            if (!Termino && !Password.usuarioValido(Usuario, Base))
            {
                //Mess = "Usuario no válido.";
                Mess = DAL.Error(12, Idioma);
                Termino = true;
            }

            //--------------------------------------------------------------------------------
            //Control de cuenta bloqueada
            //--------------------------------------------------------------------------------
            
            if (!Termino && (Password.ctaBloqueada(Usuario, Base)))
            {
                //Mess = "Cuenta Bloqueada. Consulte con el administrador.";
                Mess = DAL.Error(13, Idioma);
                Termino = true;
            }

            //--------------------------------------------------------------------------------
            //Control politica de la cuenta del usuario
            //--------------------------------------------------------------------------------
            
            long polNro = (Password.valorUserPolCuenta(Usuario, "pol_nro", Base).Length != 0) ? Convert.ToInt64(Password.valorUserPolCuenta(Usuario, "pol_nro", Base)) : 0;
            
            if (!Termino)
            {
                if (polNro == 0)
                {
                    //Mess = "No se encontro la politica de cuenta del usuario.";
                    Mess = DAL.Error(14, Idioma);
                    Termino = true;
                }
            }

            //--------------------------------------------------------------------------------
            //Control de password anterior
            //--------------------------------------------------------------------------------
            
            if (!Termino && (Password.valorHistPass(Usuario, "husrpass", Base) != Encriptar.Encrypt(DAL.EncrKy(), PassOld)))
            {
                Termino = true;

                //Veo como es la politica de intentos fallidos
                long passIntFallidos = (Password.valorPolCuenta(polNro, "pass_int_fallidos", Base).Length != 0) ? Convert.ToInt64(Password.valorPolCuenta(polNro, "pass_int_fallidos", Base)) : 0;

                //Recupero la cantidad de intentos fallidos
                long intentosFallidos = Password.logueosFallidos(Usuario, Base) + 1;

                //Control de bloqueo de cuenta por intentos fallidos si esta activo (!= 0)
                
                if ((passIntFallidos != 0) && (intentosFallidos >= passIntFallidos))
                {
                    //Bloqueo las cuentas
                    Password.bloquearCuenta(Usuario, "-1", Base);
                    Password.bajarCuenta(Usuario, Base);
                    //Mess = "Cuenta Bloqueada por intentos fallidos.";
                    Mess = DAL.Error(15, Idioma);
                }
                else
                {
                    Password.actLogFallidos(Usuario, intentosFallidos, Base);
                    //Mess = "Contraseña incorrecta.";
                    Mess = DAL.Error(16, Idioma);
                }
            }

            //--------------------------------------------------------------------------------
            //Control de longitud de cuenta
            //--------------------------------------------------------------------------------
            
            if (!Termino)
            {
                //Recupero politica
                
                long passLongitud = (Password.valorPolCuenta(polNro, "pass_longitud", Base).Length != 0) ? Convert.ToInt64(Password.valorPolCuenta(polNro, "pass_longitud", Base)) : 0;
                
                if ((passLongitud > 0) && (PassNew.Length == 0))
                {
                    //Mess = "No se permite contraseña en blanco.";
                    Mess = DAL.Error(17, Idioma);
                    Termino = true;
                }
                else
                {
                    if ((passLongitud > 0) && (PassNew.Length < passLongitud))
                    {
                        //Mess = "La longitud mínima es de " + passLongitud + " caracteres.";
                        Mess = DAL.Error(18, Idioma) + passLongitud + DAL.Error(19, Idioma);
                        Termino = true;
                    }
                }
            }

            //--------------------------------------------------------------------------------
            //Control de historia de la contraseña
            //--------------------------------------------------------------------------------
            
            if (!Termino)
            {
                //Recupero la politica de historia
                passHistoria = (Password.valorPolCuenta(polNro, "pass_historia", Base).Length != 0) ? Convert.ToInt64(Password.valorPolCuenta(polNro, "pass_historia", Base)) : 0;
                if (Password.passRepetido(Usuario, Encriptar.Encrypt(DAL.EncrKy(), PassNew), passHistoria, Base))
                {
                    //Mess = "La Contraseña coincide con una histórica.";
                    DAL.Error(21, Idioma);
                    Termino = true;
                }
            }

            //--------------------------------------------------------------------------------
            //Actualizacion de la contraseña en la base
            //--------------------------------------------------------------------------------
            
            if (!Termino)
            { 
                //Blanqueo los intentos fallidos
                Password.actLogFallidos(Usuario, 0, Base);

                //Bajo el password viejo
                Password.bajarCuenta(Usuario, Base);

                //Elimino la cantidad de pass historicos definida en la politica de cuenta
                Password.eliminarHistPass(Usuario, passHistoria, Base);

                //Ingreso el nuevo password
                Password.ingresarPass(Usuario, Encriptar.Encrypt(DAL.EncrKy(), PassNew), Base);

                //Recupero el dato del usuario de cambio de contraseña
                bool usrPassCambiar = (Password.valorUserPer(Usuario, "usrpasscambiar", Base).Length != 0) ? (Convert.ToInt64(Password.valorUserPer(Usuario, "usrpasscambiar", Base)) == -1) : false;
                if (usrPassCambiar)
                    Password.CambiarPassUser(Usuario, "0", Base);

                //Cambio el password en la base
                Password.CambiarPassBase(Usuario, Encriptar.Encrypt(DAL.EncrKy(), PassNew), Encriptar.Encrypt(DAL.EncrKy(), PassOld), Base);
            }

            return Mess;        
        }

        //-------------------------------------------------------------------------------------
        //-------------------------------------------------------------------------------------
        
        [WebMethod(Description = "(ORG) Dado un codigo de empleado y un codigo de base retorna el codigo de empleado padre.")]
        public long Padre(long CodEmp, int Base)
        {
            string cn = DAL.constr(Base.ToString());
            string sql;
            DataSet ds = new DataSet();
            long Salida = 0;

            sql = "SELECT empreporta, ternro CodEmp ";
            sql = sql + "FROM empleado ";
            sql = sql + "WHERE empleado.ternro = " + CodEmp.ToString() + " ";
            sql = sql + "AND empleado.empest = -1 ";

            da = new OleDbDataAdapter(sql, cn);

            try
            {
                da.Fill(ds);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            if (ds.Tables[0].Rows.Count > 0)
            {
                if (ds.Tables[0].Rows[0].ItemArray[0] != DBNull.Value)
                {
                    Salida = Convert.ToInt64(ds.Tables[0].Rows[0].ItemArray[0].ToString());
                }
            }

            return Salida;
        }

        //-------------------------------------------------------------------------------------
        //-------------------------------------------------------------------------------------
        
        [WebMethod(Description = "(ORG) Dado un codigo de empleado y un codigo de base retorna todos los codigos de empleados hijos.")]
        public DataTable Hijos(long CodEmp, int Base)
        {
            string cn = DAL.constr(Base.ToString());
            string sql;
            DataSet ds = new DataSet();

            sql = "SELECT ternro CodEmp ";
            sql = sql + "FROM empleado ";
            sql = sql + "WHERE empleado.empreporta = " + CodEmp.ToString() + " ";
            sql = sql + "AND empleado.empest = -1 ";

            da = new OleDbDataAdapter(sql, cn);

            try
            {
                da.Fill(ds);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return ds.Tables[0];
        }

        //-------------------------------------------------------------------------------------
        //-------------------------------------------------------------------------------------
        
        [WebMethod(Description = "(ORG) Dado un codigo de empleado y un codigo de base retorna los datos del empleado.")]
        public DataTable DatosEmp(long CodEmp, int Base)
        {
            string cn = DAL.constr(Base.ToString());
            string sql;
            DataSet ds = new DataSet();
            DataSet ds2 = new DataSet();
            DataRow filaAux;

            //Creo la tabla de salida
            DataTable tablaSalida = new DataTable("table");

            DataColumn Columna = new DataColumn();

            Columna.DataType = System.Type.GetType("System.String");
            Columna.ColumnName = "legajo";
            Columna.AutoIncrement = false;
            Columna.Unique = false;
            tablaSalida.Columns.Add(Columna);
            Columna = new DataColumn();
            Columna.DataType = System.Type.GetType("System.String");
            Columna.ColumnName = "apellido";
            Columna.AutoIncrement = false;
            Columna.Unique = false;
            tablaSalida.Columns.Add(Columna);
            Columna = new DataColumn();
            Columna.DataType = System.Type.GetType("System.String");
            Columna.ColumnName = "nombre";
            Columna.AutoIncrement = false;
            Columna.Unique = false;
            tablaSalida.Columns.Add(Columna);
            Columna = new DataColumn();
            Columna.DataType = System.Type.GetType("System.String");
            Columna.ColumnName = "mail";
            Columna.AutoIncrement = false;
            Columna.Unique = false;
            tablaSalida.Columns.Add(Columna);
            Columna = new DataColumn();
            Columna.DataType = System.Type.GetType("System.String");
            Columna.ColumnName = "interno";
            Columna.AutoIncrement = false;
            Columna.Unique = false;
            tablaSalida.Columns.Add(Columna);
            Columna = new DataColumn();
            Columna.DataType = System.Type.GetType("System.String");
            Columna.ColumnName = "Documento";
            Columna.AutoIncrement = false;
            Columna.Unique = false;
            tablaSalida.Columns.Add(Columna);
            Columna = new DataColumn();
            Columna.DataType = System.Type.GetType("System.String");
            Columna.ColumnName = "TipoEst1";
            Columna.AutoIncrement = false;
            Columna.Unique = false;
            tablaSalida.Columns.Add(Columna);
            Columna = new DataColumn();
            Columna.DataType = System.Type.GetType("System.String");
            Columna.ColumnName = "Est1";
            Columna.AutoIncrement = false;
            Columna.Unique = false;
            tablaSalida.Columns.Add(Columna);
            Columna = new DataColumn();
            Columna.DataType = System.Type.GetType("System.String");
            Columna.ColumnName = "TipoEst2";
            Columna.AutoIncrement = false;
            Columna.Unique = false;
            tablaSalida.Columns.Add(Columna);
            Columna = new DataColumn();
            Columna.DataType = System.Type.GetType("System.String");
            Columna.ColumnName = "Est2";
            Columna.AutoIncrement = false;
            Columna.Unique = false;
            tablaSalida.Columns.Add(Columna);
            Columna = new DataColumn();
            Columna.DataType = System.Type.GetType("System.String");
            Columna.ColumnName = "TipoEst3";
            Columna.AutoIncrement = false;
            Columna.Unique = false;
            tablaSalida.Columns.Add(Columna);
            Columna = new DataColumn();
            Columna.DataType = System.Type.GetType("System.String");
            Columna.ColumnName = "Est3";
            Columna.AutoIncrement = false;
            Columna.Unique = false;
            tablaSalida.Columns.Add(Columna);
            Columna = new DataColumn();
            Columna.DataType = System.Type.GetType("System.String");
            Columna.ColumnName = "Imagen";
            Columna.AutoIncrement = false;
            Columna.Unique = false;
            tablaSalida.Columns.Add(Columna);

            sql = "SELECT empleado.empleg legajo, terape apellido,terape2 apellido2, ternom nombre, ternom2 nombre2, empemail mail, empinterno interno, empleado.ternro ";
            sql = sql + "FROM empleado ";
            sql = sql + "WHERE empleado.ternro = " + CodEmp.ToString() + " ";

            da = new OleDbDataAdapter(sql, cn);

            try
            {
                da.Fill(ds);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            if (ds.Tables[0].Rows.Count > 0)
            {
                DataTable tablaAux = ds.Tables[0];

                foreach (DataRow fila in tablaAux.Rows)
                {
                    filaAux = tablaSalida.NewRow();
                    filaAux["legajo"] = fila["legajo"].ToString();
                    filaAux["apellido"] = fila["apellido"].ToString() + " " + fila["apellido2"].ToString();
                    filaAux["nombre"] = fila["nombre"].ToString() + " " + fila["nombre2"].ToString();
                    filaAux["mail"] = fila["mail"].ToString();
                    filaAux["interno"] = fila["interno"].ToString();
                    filaAux["Documento"] = Documento(CodEmp, Base);
                    filaAux["TipoEst1"] = DAL.DescEstr(1);
                    filaAux["Est1"] = Estructura(CodEmp, Convert.ToInt64(DAL.NroEstr(1)), Base);
                    filaAux["TipoEst2"] = DAL.DescEstr(2);
                    filaAux["Est2"] = Estructura(CodEmp, Convert.ToInt64(DAL.NroEstr(2)), Base);
                    filaAux["TipoEst3"] = DAL.DescEstr(3);
                    filaAux["Est3"] = Estructura(CodEmp, Convert.ToInt64(DAL.NroEstr(3)), Base);

                    string img = "nofoto.jpg";

                    sql = "SELECT terimnombre, tipimanchodef, tipimaltodef, tipimdire, ter_imag.terimfecha ";
                    sql = sql + "FROM ter_imag ";
                    sql = sql + "LEFT JOIN tipoimag ON tipoimag.tipimnro = ter_imag.tipimnro ";
                    sql = sql + "WHERE  ter_imag.tipimnro = 3 ";
                    sql = sql + "AND ter_imag.ternro = " + fila["ternro"].ToString() + " ";
                    sql = sql + "ORDER BY ter_imag.terimfecha DESC ";

                    da = new OleDbDataAdapter(sql, cn);

                    try
                    {
                        da.Fill(ds2);
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }

                    if (ds2.Tables[0].Rows.Count > 0)
                    {
                        if (ds2.Tables[0].Rows[0].ItemArray[0] != DBNull.Value)
                        {
                            img = Convert.ToString(ds2.Tables[0].Rows[0].ItemArray[0].ToString());
                        }
                    }

                    filaAux["Imagen"] = img;
                    tablaSalida.Rows.Add(filaAux);
                }
            }

            return tablaSalida;
        }

        //-------------------------------------------------------------------------------------
        //-------------------------------------------------------------------------------------
        
        private string Estructura(long CodEmp, long TeNro, int Base)
        {
            string cn = DAL.constr(Base.ToString());
            string sql;
            DataSet ds = new DataSet();
            string Salida = "";
            string TipoDB = DAL.TipoBase(Base.ToString());
            string FechaAct = DateTime.Today.ToString("dd/MM/yyyy");

            sql = "SELECT estructura.estrdabr ";
            sql = sql + "FROM his_estructura ";
            sql = sql + "INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro ";
            sql = sql + "WHERE his_estructura.ternro = " + CodEmp.ToString() + " ";
            sql = sql + "AND (his_estructura.htetdesde <= " + Fecha.cambiaFecha(FechaAct, TipoDB) + " ";
            sql = sql + "AND (his_estructura.htethasta is null or his_estructura.htethasta >= " + Fecha.cambiaFecha(FechaAct, TipoDB) + ")) ";
            sql = sql + "AND his_estructura.tenro = " + TeNro.ToString() + " ";

            da = new OleDbDataAdapter(sql, cn);

            try
            {
                da.Fill(ds);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            if (ds.Tables[0].Rows.Count > 0)
            {
                if (ds.Tables[0].Rows[0].ItemArray[0] != DBNull.Value)
                {
                    Salida = Convert.ToString(ds.Tables[0].Rows[0].ItemArray[0].ToString());
                }
            }

            return Salida;
        }

        //-------------------------------------------------------------------------------------
        //-------------------------------------------------------------------------------------
        
        private string Documento(long CodEmp, int Base)
        {
            string cn = DAL.constr(Base.ToString());
            string sql;
            DataSet ds = new DataSet();
            string Salida = "";

            sql = "SELECT ter_doc.ternro, tipodocu.tidsigla, ter_doc.nrodoc ";
            sql = sql + "FROM ter_doc ";
            sql = sql + "INNER JOIN tipodocu ON tipodocu.tidnro = ter_doc.tidnro ";
            sql = sql + "WHERE ter_doc.tidnro <= 5 ";
            sql = sql + "AND ter_doc.ternro =  " + CodEmp.ToString() + " ";
            sql = sql + "ORDER BY ter_doc.tidnro ";

            da = new OleDbDataAdapter(sql, cn);

            try
            {
                da.Fill(ds);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            if (ds.Tables[0].Rows.Count > 0)
            {
                Salida = Convert.ToString(ds.Tables[0].Rows[0].ItemArray[1].ToString()) + " " + Convert.ToString(ds.Tables[0].Rows[0].ItemArray[2].ToString());
            }

            return Salida;
        }

        //-------------------------------------------------------------------------------------
        //-------------------------------------------------------------------------------------
        
        [WebMethod(Description = "(ORG) Asigna a CodHijo el padre CodPadre para la base Base")]
        public Boolean AsignarPadre(long CodPadre, long CodHijo, int Base)
        {
            string sql = "";
        
            OleDbConnection cn = new OleDbConnection();
            cn.ConnectionString = DAL.constr(Base.ToString());

            try
            {
                cn.Open();
                
                sql = "UPDATE empleado ";
                sql = sql + "SET empreporta = " + CodPadre.ToString() + " ";
                sql = sql + "WHERE ternro = " + CodHijo.ToString() + " ";

                OleDbCommand cmd = new OleDbCommand(sql, cn);

                cmd.ExecuteNonQuery();
            }
            catch
            {
                return false;
            }
            finally
            {
                if (cn.State == ConnectionState.Open) cn.Close();
            }

            return true;
        }

        //-------------------------------------------------------------------------------------
        //-------------------------------------------------------------------------------------
        
        [WebMethod(Description = "(ORG) Asigna a CodHijo el padre CodPadre para la base Base")]
        public Boolean AsignarHijo(long CodHijo, long CodPadre, int Base)
        {
            string sql = "";
            OleDbConnection cn = new OleDbConnection();
            cn.ConnectionString = DAL.constr(Base.ToString());

            try
            {
                cn.Open();

                sql = "UPDATE empleado ";
                sql = sql + "SET empreporta = " + CodPadre.ToString() + " ";
                sql = sql + "WHERE ternro = " + CodHijo.ToString() + " ";

                OleDbCommand cmd = new OleDbCommand(sql, cn);

                cmd.ExecuteNonQuery();
            }
            catch
            {
                return false;
            }
            finally
            {
                if (cn.State == ConnectionState.Open) cn.Close();
            }

            return true;
        }

        //-------------------------------------------------------------------------------------
        //-------------------------------------------------------------------------------------
        
        [WebMethod(Description = "(ORG) Borra relacion entre padre e hijo para la base Base")]
        public Boolean BorrarHijo(long CodHijo, long CodPadre, int Base)
        {
            string sql = "";
            OleDbConnection cn = new OleDbConnection();
            cn.ConnectionString = DAL.constr(Base.ToString());

            try
            {
                cn.Open();

                sql = "UPDATE empleado ";
                sql = sql + "SET empreporta = NULL ";
                sql = sql + "WHERE ternro = " + CodHijo.ToString() + " ";

                OleDbCommand cmd = new OleDbCommand(sql, cn);

                cmd.ExecuteNonQuery();
            }
            catch
            {
                return false;
            }
            finally
            {
                if (cn.State == ConnectionState.Open) cn.Close();
            }

            return true;
        }

        //-------------------------------------------------------------------------------------
        //-------------------------------------------------------------------------------------
        
        [WebMethod(Description = "(ORG) Borra relacion entre padre e hijo para la base Base")]
        public Boolean BorrarPadre(long CodHijo, long CodPadre, int Base)
        {
            string sql = "";
            OleDbConnection cn = new OleDbConnection();
            cn.ConnectionString = DAL.constr(Base.ToString());

            try
            {
                cn.Open();

                sql = "UPDATE empleado ";
                sql = sql + "SET empreporta = NULL ";
                sql = sql + "WHERE ternro = " + CodHijo.ToString() + " ";

                OleDbCommand cmd = new OleDbCommand(sql, cn);

                cmd.ExecuteNonQuery();
            }
            catch
            {
                return false;
            }
            finally
            {
                if (cn.State == ConnectionState.Open) cn.Close();
            }

            return true;
        }

        //-------------------------------------------------------------------------------------
        //-------------------------------------------------------------------------------------
        
        [WebMethod(Description = "(ORG) Dado un legajo de empleado, un filtro de estados activos (-1), inactivos (0) o ambos (1), y un codigo de base retorna codigo interno y nombre y apellido del empleado. Si Legajo es cero retorna el primer empleado.")]
        public DataTable BuscarEmpleado(long Legajo, long Activo, int Base)
        {
            string cn = DAL.constr(Base.ToString());
            string sql;
            DataSet ds = new DataSet();

            DataTable tablaSalida = new DataTable("table");

            DataColumn Columna1 = new DataColumn();
            Columna1.DataType = System.Type.GetType("System.Int32");
            Columna1.ColumnName = "Legajo";
            Columna1.AutoIncrement = false;
            Columna1.Unique = false;
            tablaSalida.Columns.Add(Columna1);

            DataColumn Columna2 = new DataColumn();
            Columna2.DataType = System.Type.GetType("System.Int32");
            Columna2.ColumnName = "CodEmp";
            Columna2.AutoIncrement = false;
            Columna2.Unique = false;
            tablaSalida.Columns.Add(Columna2);

            DataColumn Columna3 = new DataColumn();
            Columna3.DataType = System.Type.GetType("System.String");
            Columna3.ColumnName = "Apellido";
            Columna3.AutoIncrement = false;
            Columna3.Unique = false;
            tablaSalida.Columns.Add(Columna3);

            DataColumn Columna4 = new DataColumn();
            Columna4.DataType = System.Type.GetType("System.String");
            Columna4.ColumnName = "Nombre";
            Columna4.AutoIncrement = false;
            Columna4.Unique = false;
            tablaSalida.Columns.Add(Columna4);

            sql = "SELECT ";
            sql = sql + "empleg Legajo, ";
            sql = sql + "ternro CodEmp, ";
            sql = sql + "terape, ";
            sql = sql + "terape2, ";
            sql = sql + "ternom, ";
            sql = sql + "ternom2 ";
            sql = sql + "FROM empleado ";

/*          switch (Activo)
            {
                case -1:
*/                    sql = sql + "WHERE empest = -1 ";
/*                    break;
                case 0:
                    sql = sql + "WHERE empest <> -1 ";
                    break;
                case 1:
                    sql = sql + "WHERE (1 = 1) ";
                    break;
            }
*/
            if (Legajo != 0)
                sql = sql + "AND empleg = " + Legajo.ToString() + " ";
            
            sql = sql + "ORDER BY empleg ";

            da = new OleDbDataAdapter(sql, cn);

            try
            {
                if (Legajo != 0)
                    da.Fill(ds);
                else
                    da.Fill(ds, 0, 1, "Resultado");

                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    DataRow r = tablaSalida.NewRow();

                    r["Legajo"] = ds.Tables[0].Rows[i]["Legajo"];
                    r["CodEmp"] = ds.Tables[0].Rows[i]["CodEmp"];
                    r["Apellido"] = ds.Tables[0].Rows[i]["terape"] + " " + ds.Tables[0].Rows[i]["terape2"];
                    r["Nombre"] = ds.Tables[0].Rows[i]["ternom"] + " " + ds.Tables[0].Rows[i]["ternom2"];

                    tablaSalida.Rows.Add(r);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return tablaSalida;
        }

        //-------------------------------------------------------------------------------------
        //-------------------------------------------------------------------------------------
        
        [WebMethod(Description = "(ORG) Dado un legajo de empleado, un filtro de estados activos (-1), inactivos (0) o ambos (1), y un codigo de base retorna el siguiente empleado.")]
        public DataTable SgtEmpl(long Legajo, long Activo, int Base)
        {
            string cn = DAL.constr(Base.ToString());
            string sql;
            DataSet ds = new DataSet();

            DataTable tablaSalida = new DataTable("table");
            
            DataColumn Columna1 = new DataColumn();
            Columna1.DataType = System.Type.GetType("System.Int32");
            Columna1.ColumnName = "Legajo";
            Columna1.AutoIncrement = false;
            Columna1.Unique = false;
            tablaSalida.Columns.Add(Columna1);

            DataColumn Columna2 = new DataColumn();
            Columna2.DataType = System.Type.GetType("System.Int32");
            Columna2.ColumnName = "CodEmp";
            Columna2.AutoIncrement = false;
            Columna2.Unique = false;
            tablaSalida.Columns.Add(Columna2);

            DataColumn Columna3 = new DataColumn();
            Columna3.DataType = System.Type.GetType("System.String");
            Columna3.ColumnName = "Apellido";
            Columna3.AutoIncrement = false;
            Columna3.Unique = false;
            tablaSalida.Columns.Add(Columna3);

            DataColumn Columna4 = new DataColumn();
            Columna4.DataType = System.Type.GetType("System.String");
            Columna4.ColumnName = "Nombre";
            Columna4.AutoIncrement = false;
            Columna4.Unique = false;
            tablaSalida.Columns.Add(Columna4);

            sql = "SELECT ";
            sql = sql + "empleg Legajo, ";
            sql = sql + "ternro CodEmp, ";
            sql = sql + "terape, ";
            sql = sql + "terape2, ";
            sql = sql + "ternom, ";
            sql = sql + "ternom2 ";
            sql = sql + "FROM empleado ";

/*            switch (Activo)
            {
                case -1:
*/                    sql = sql + "WHERE empest = -1 ";
/*                    break;
                case 0:
                    sql = sql + "WHERE empest <> -1 ";
                    break;
                case 1:
                    sql = sql + "WHERE (1 = 1) ";
                    break;
            }
*/
            sql = sql + "AND empleg > " + Legajo.ToString() + " ";
            sql = sql + "ORDER BY empleg ";

            da = new OleDbDataAdapter(sql, cn);

            try
            {
                da.Fill(ds, 0, 1, "Resultado");

                //Cargo la tabla de salida

                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    DataRow r = tablaSalida.NewRow();

                    r["Legajo"] = ds.Tables[0].Rows[i]["Legajo"];
                    r["CodEmp"] = ds.Tables[0].Rows[i]["CodEmp"];
                    r["Apellido"] = ds.Tables[0].Rows[i]["terape"] + " " + ds.Tables[0].Rows[i]["terape2"];
                    r["Nombre"] = ds.Tables[0].Rows[i]["ternom"] + " " + ds.Tables[0].Rows[i]["ternom2"];

                    tablaSalida.Rows.Add(r);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return tablaSalida;
        }

        //-------------------------------------------------------------------------------------
        //-------------------------------------------------------------------------------------
        
        [WebMethod(Description = "(ORG) Dado un legajo de empleado, un filtro de estados activos (-1), inactivos (0) o ambos (1), y un codigo de base retorna el anterior empleado.")]
        public DataTable AntEmpl(long Legajo, long Activo, int Base)
        {
            string cn = DAL.constr(Base.ToString());
            string sql;
            DataSet ds = new DataSet();

            DataTable tablaSalida = new DataTable("table");

            DataColumn Columna1 = new DataColumn();
            Columna1.DataType = System.Type.GetType("System.Int32");
            Columna1.ColumnName = "Legajo";
            Columna1.AutoIncrement = false;
            Columna1.Unique = false;
            tablaSalida.Columns.Add(Columna1);

            DataColumn Columna2 = new DataColumn();
            Columna2.DataType = System.Type.GetType("System.Int32");
            Columna2.ColumnName = "CodEmp";
            Columna2.AutoIncrement = false;
            Columna2.Unique = false;
            tablaSalida.Columns.Add(Columna2);

            DataColumn Columna3 = new DataColumn();
            Columna3.DataType = System.Type.GetType("System.String");
            Columna3.ColumnName = "Apellido";
            Columna3.AutoIncrement = false;
            Columna3.Unique = false;
            tablaSalida.Columns.Add(Columna3);

            DataColumn Columna4 = new DataColumn();
            Columna4.DataType = System.Type.GetType("System.String");
            Columna4.ColumnName = "Nombre";
            Columna4.AutoIncrement = false;
            Columna4.Unique = false;
            tablaSalida.Columns.Add(Columna4);

            sql = "SELECT ";
            sql = sql + "empleg Legajo, ";
            sql = sql + "ternro CodEmp, ";
            sql = sql + "terape, ";
            sql = sql + "terape2, ";
            sql = sql + "ternom, ";
            sql = sql + "ternom2 ";
            sql = sql + "FROM empleado ";

/*            switch (Activo)
            {
                case -1:
*/                    sql = sql + "WHERE empest = -1 ";
/*                    break;
                case 0:
                    sql = sql + "WHERE empest <> -1 ";
                    break;
                case 1:
                    sql = sql + "WHERE (1 = 1) ";
                    break;
            }
*/
            sql = sql + "AND empleg < " + Legajo.ToString() + " ";
            sql = sql + "ORDER BY empleg DESC ";

            da = new OleDbDataAdapter(sql, cn);

            try
            {
                da.Fill(ds, 0, 1, "Resultado");

                //Cargo la tabla de salida

                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    DataRow r = tablaSalida.NewRow();

                    r["Legajo"] = ds.Tables[0].Rows[i]["Legajo"];
                    r["CodEmp"] = ds.Tables[0].Rows[i]["CodEmp"];
                    r["Apellido"] = ds.Tables[0].Rows[i]["terape"] + " " + ds.Tables[0].Rows[i]["terape2"];
                    r["Nombre"] = ds.Tables[0].Rows[i]["ternom"] + " " + ds.Tables[0].Rows[i]["ternom2"];

                    tablaSalida.Rows.Add(r);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return tablaSalida;
        }


        //-------------------------------------------------------------------------------------
        //-------------------------------------------------------------------------------------

        [WebMethod(Description = "(Interfaz Inteligente) Devuelve el estado del postulante: 1) Empleado, 2) ExEmpleado, 3)Postulante Activo, 4)Postulante Inactivo, 5) Lista Negra")]
        public DataTable EstadoPostulante(int tipoDoc, string Doc, int Sexo, string Base)
        {
            string cn = DAL.constr(Base);
            string sql;
            DataSet ds = new DataSet();
            DataTable tablaSalida = new DataTable("table");
            int estado = 0;
            string fecha = "";
            string Desc = "";

            //Creo la estructura de la tabla de salida
            DataColumn Columna1 = new DataColumn();
            Columna1.DataType = System.Type.GetType("System.Int16");
            Columna1.ColumnName = "Estado";
            Columna1.AutoIncrement = false;
            Columna1.Unique = false;
            tablaSalida.Columns.Add(Columna1);
            DataColumn Columna2 = new DataColumn();
            Columna2.DataType = System.Type.GetType("System.String");
            Columna2.ColumnName = "Fecha";
            Columna2.AutoIncrement = false;
            Columna2.Unique = false;
            tablaSalida.Columns.Add(Columna2);
            DataColumn Columna3 = new DataColumn();
            Columna3.DataType = System.Type.GetType("System.String");
            Columna3.ColumnName = "Descr";
            Columna3.AutoIncrement = false;
            Columna3.Unique = false;
            tablaSalida.Columns.Add(Columna3);

            //Busco al tercero en la black list
            sql = "SELECT tercero.ternro FROM b_list";
            sql = sql + " INNER JOIN tercero ON tercero.ternro = b_list.ternro";
            sql = sql + " INNER JOIN ter_doc ON ter_doc.ternro = tercero.ternro";
            sql = sql + " WHERE tercero.tersex = " + Sexo.ToString();
            sql = sql + " AND ter_doc.tidnro = " + tipoDoc.ToString();
            sql = sql + " AND ter_doc.nrodoc = '" + Doc + "'";
            da = new OleDbDataAdapter(sql, cn);
            try
            {
                da.Fill(ds, "b_list");
            }
            catch (Exception ex)
            {
                throw ex;
            }
            if (ds.Tables["b_list"].Rows.Count > 0)
            {
                estado = 5;
                fecha = DateTime.Now.ToString();
            }

            //Busco al tercero como empleado
            if (estado == 0)
            {
                sql = "SELECT tercero.ternro FROM empleado";
                sql = sql + " INNER JOIN tercero ON tercero.ternro = empleado.ternro";
                sql = sql + " INNER JOIN ter_doc ON ter_doc.ternro = tercero.ternro";
                sql = sql + " WHERE tercero.tersex = " + Sexo.ToString();
                sql = sql + " AND empleado.empest = -1";
                sql = sql + " AND ter_doc.tidnro = " + tipoDoc.ToString();
                sql = sql + " AND ter_doc.nrodoc = '" + Doc + "'";
                da = new OleDbDataAdapter(sql, cn);
                try
                {
                    da.Fill(ds, "empleado");
                }
                catch (Exception ex)
                {
                    throw ex;
                }
                if (ds.Tables["empleado"].Rows.Count > 0)
                {
                    estado = 1;
                    fecha = DateTime.Now.ToString();
                }
            }

            //Busco al tercero como ex-empleado
            if (estado == 0)
            {
                sql = "SELECT tercero.ternro FROM empleado";
                sql = sql + " INNER JOIN tercero ON tercero.ternro = empleado.ternro";
                sql = sql + " INNER JOIN ter_doc ON ter_doc.ternro = tercero.ternro";
                sql = sql + " WHERE tercero.tersex = " + Sexo.ToString();
                sql = sql + " AND empleado.empest <> -1";
                sql = sql + " AND ter_doc.tidnro = " + tipoDoc.ToString();
                sql = sql + " AND ter_doc.nrodoc = '" + Doc + "'";
                da = new OleDbDataAdapter(sql, cn);
                try
                {
                    da.Fill(ds, "exempleado");
                }
                catch (Exception ex)
                {
                    throw ex;
                }
                if (ds.Tables["exempleado"].Rows.Count > 0)
                {
                    estado = 1;
                    fecha = DateTime.Now.ToString();
                }
            }

            //Busco al tercero como postulante activo
            if (estado == 0)
            {
                sql = "SELECT TOP 1 (pos_actividad.actdesabr + ' - ' + pos_estado.estdesabr) Descr, pos_seguimiento.segfec Fecha";
                sql = sql + " FROM pos_postulante";
                sql = sql + " LEFT JOIN pos_seguimiento ON pos_seguimiento.ternro =  pos_postulante.ternro";
                sql = sql + " LEFT JOIN pos_actividad ON pos_actividad.actnro = pos_seguimiento.actnro";
                sql = sql + " LEFT JOIN pos_estado ON pos_estado.estnro = pos_seguimiento.estnro";
                sql = sql + " INNER JOIN tercero ON tercero.ternro = pos_postulante.ternro";
                sql = sql + " INNER JOIN ter_doc ON ter_doc.ternro = tercero.ternro";
                sql = sql + " WHERE posest = -1";
                sql = sql + " AND tercero.tersex = " + Sexo.ToString();
                sql = sql + " AND ter_doc.tidnro = " + tipoDoc.ToString();
                sql = sql + " AND ter_doc.nrodoc = '" + Doc + "'";
                sql = sql + " ORDER BY segfec DESC";
                da = new OleDbDataAdapter(sql, cn);
                try
                {
                    da.Fill(ds, "postulante");
                }
                catch (Exception ex)
                {
                    throw ex;
                }
                if (ds.Tables["postulante"].Rows.Count > 0)
                {
                    estado = 3;
                    fecha = ds.Tables["postulante"].Rows[0]["Fecha"].ToString();
                    Desc = ds.Tables["postulante"].Rows[0]["Descr"].ToString();
                }
            }

            //Busco al tercero como postulante inactico
            if (estado == 0)
            {
                sql = "SELECT TOP 1 (pos_actividad.actdesabr + '-' +  pos_estado.estdesabr) Descr, pos_seguimiento.segfec Fecha";
                sql = sql + " FROM pos_postulante";
                sql = sql + " LEFT JOIN pos_seguimiento ON pos_seguimiento.ternro =  pos_postulante.ternro";
                sql = sql + " LEFT JOIN pos_actividad ON pos_actividad.actnro = pos_seguimiento.actnro";
                sql = sql + " LEFT JOIN pos_estado ON pos_estado.estnro = pos_seguimiento.estnro";
                sql = sql + " INNER JOIN tercero ON tercero.ternro = pos_postulante.ternro";
                sql = sql + " INNER JOIN ter_doc ON ter_doc.ternro = tercero.ternro";
                sql = sql + " WHERE posest <> -1";
                sql = sql + " AND tercero.tersex = " + Sexo.ToString();
                sql = sql + " AND ter_doc.tidnro = " + tipoDoc.ToString();
                sql = sql + " AND ter_doc.nrodoc = '" + Doc + "'";
                sql = sql + " ORDER BY segfec DESC";
                da = new OleDbDataAdapter(sql, cn);
                try
                {
                    da.Fill(ds, "postulanteInac");
                }
                catch (Exception ex)
                {
                    throw ex;
                }
                if (ds.Tables["postulante"].Rows.Count > 0)
                {
                    estado = 4;
                    fecha = ds.Tables["postulanteInac"].Rows[0]["Fecha"].ToString();
                    Desc = ds.Tables["postulanteInac"].Rows[0]["Descr"].ToString();
                }
            }

            DataRow fila = tablaSalida.NewRow();
            fila["Estado"] = estado;
            fila["Fecha"] = fecha;
            fila["Descr"] = Desc;
            tablaSalida.Rows.Add(fila);

            return tablaSalida;

        }


        [WebMethod(Description = "(Interfaz Inteligente) Devuelve codigo y descripcion de la tabla plana seleccionada.")]
        public DataTable TablaPlana(string Tabla, string Base)
        {
            string cn = DAL.constr(Base);
            string sql = "";
            DataSet ds = new DataSet();

            switch (Tabla.ToUpper())
            {
                case "PAIS":
                    sql = "SELECT pais.paisnro cod, pais.paisdesc descr FROM pais";
                    break;
                case "PROVINCIA":
                    sql = "SELECT provnro cod, provdesc descr FROM provincia";
                    break;
                case "PARTIDO":
                    sql = "SELECT partnro cod, partnom descr FROM partido";
                    break;
                case "ZONA":
                    sql = "SELECT zonanro cod, zonadesc descr FROM zona";
                    break;
                case "LOCALIDAD":
                    sql = "SELECT locnro cod, locdesc descr FROM localidad";
                    break;
                case "PROCEDENCIA":
                    sql = "SELECT pronro cod, prodesabr descr FROM pos_procedencia";
                    break;
                case "NIVELESTUDIO":
                    sql = "SELECT nivnro cod, nivdesc descr FROM nivest";
                    break;
                case "TITULO":
                    sql = "SELECT titnro cod, titdesabr descr FROM titulo";
                    break;
                case "INSTITUCION":
                    sql = "SELECT instnro cod, instdes descr FROM institucion";
                    break;
                case "CARRERA":
                    sql = "SELECT carredunro cod, carredudesabr descr FROM cap_carr_edu";
                    break;
                case "CARGO":
                    sql = "SELECT carnro cod, cardesabr descr FROM cargo";
                    break;
                case "LISTAEMPRESA":
                    sql = "SELECT lempnro cod, lempdes descr FROM listaemp";
                    break;
                case "CAUSA":
                    sql = "SELECT caunro cod, caudes descr FROM causa";
                    break;
                case "IDIOMA":
                    sql = "SELECT idinro cod, ididesc descr FROM Idioma";
                    break;
                case "IDIOMANIVEL":
                    sql = "SELECT idnivnro cod, idnivdesabr descr FROM idinivel";
                    break;
                case "TIPOCURSO":
                    sql = "SELECT tipcurnro cod, tipcurdesabr descr FROM cap_tipocurso";
                    break;
                case "ESPECIALIZACION":
                    sql = "SELECT espnro cod, espdesabr descr FROM especializacion";
                    break;
                case "ELEMENTOESPEC":
                    sql = "SELECT eltananro cod, eltanadesabr descr FROM eltoana";
                    break;
                case "NIVELESPC":
                    sql = "SELECT espnivnro cod, espnivdesabr descr FROM espnivel";
                    break;
                case "TIPOTELEF":
                    sql = "SELECT titelnro cod, titeldes descr FROM tipotel";
                    break;
                case "TIPODOC":
                    sql = "SELECT tidnro cod, tidsigla descr FROM tipodocu ORDER BY tidnro";
                    break;
                case "AREA":
                    sql = "SELECT arenro cod, aredesabr descr FROM areas ORDER BY arenro";
                    break;
                case "ESTADOCIVIL":
                    sql = "SELECT estcivnro cod, estcivdesabr descr FROM estcivil ORDER BY estcivnro";                    
                    break;
                case "NACIONALIDAD":
                    sql = "SELECT nacionalnro cod, nacionaldes descr FROM nacionalidad ORDER BY nacionalnro";                    
                    break;
                case "INDUSTRIA":
                    sql = "";
                    break;
                default:
                    throw new Exception("Nombre de tabla erronea");
                }

            if (sql.Length != 0)
            {
                da = new OleDbDataAdapter(sql, cn);
                try
                {
                    da.Fill(ds);
                }
                catch (Exception ex)
                {
                    throw ex;
                }
                return ds.Tables[0];
            }
            else
            {
                DataTable tablaSalida = new DataTable("table");

                //Creo la estructura de la tabla de salida
                DataColumn Columna1 = new DataColumn();
                Columna1.DataType = System.Type.GetType("System.String");
                Columna1.ColumnName = "cod";
                Columna1.AutoIncrement = false;
                Columna1.Unique = false;
                tablaSalida.Columns.Add(Columna1);
                DataColumn Columna2 = new DataColumn();
                Columna2.DataType = System.Type.GetType("System.String");
                Columna2.ColumnName = "descr";
                Columna2.AutoIncrement = false;
                Columna2.Unique = false;
                tablaSalida.Columns.Add(Columna2);

                return tablaSalida;

            }
        }

    }
}