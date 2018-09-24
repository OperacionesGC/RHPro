using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Web;
using System.Web.Services;
using System.Web.Services.Protocols;
using System.Data.OleDb;
using System.Configuration;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Reflection;
using System.Diagnostics;
using System.Data.OracleClient;
 


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

        private Boolean multiidioma;
        private String idiomausuario; //idioma de usuario con guion medio
        private String idiomausuario2; //idioma de usuario sin guion medio
        private Dictionary<string, int> DiccionarioMRU = new Dictionary<string, int>();
        private Dictionary<string, int> DiccionarioPonderacion = new Dictionary<string, int>();

       
 
        // NAM - Determina si existe la tabla lenguaje_etiqueta para saber si se aplica o no multiidioma.
        private void determinarMI(String Base)
        {   

            string cn =  DAL.constr(Base);
            
            DataSet idiomaAux = new DataSet();
            String sql;
            multiidioma = false;
 

            if (DAL.TipoBase(Base).ToUpper() == "MSSQL")
            {

                //me fijo si existe la tabla lenguaje etiqueta para aplicar multiidioma a los titulos
                sql = "SELECT * FROM information_schema.tables WHERE table_name = 'lenguaje_etiqueta'";

                da = new OleDbDataAdapter(sql, cn);

                try
                {
                    da.Fill(idiomaAux);
                    multiidioma = (idiomaAux.Tables[0].Rows.Count > 0);
                }
                catch (Exception ex)
                {
                  //  throw ex;
                }

                // si no existe la tabla, entonces no aplico multiidioma
                //if (idiomaAux.Tables[0].Rows.Count == 0)
                //{
                //    multiidioma = false;
                //}
                //else
                //{
                //    multiidioma = true;
                //}

            }
            else
            {
                //me fijo si existe la tabla lenguaje etiqueta para aplicar multiidioma a los titulos
                //sql = "select table_name from user_tables where lower(table_name) like '%lenguaje_etiqueta%'";
                sql = "select etiqueta from lenguaje_etiqueta";

                da = new OleDbDataAdapter(sql, cn);

                try
                {
                    da.Fill(idiomaAux);
                    //multiidioma = (idiomaAux.Tables[0].Rows.Count > 0);
                    multiidioma = true;
                }
                catch (Exception ex)
                {
                    multiidioma = false;
                   // throw ex;
                }

                // si no existe la tabla, entonces no aplico multiidioma
                //if (idiomaAux.Tables[0].Rows.Count == 0)
                //{
                //    multiidioma = false;
                //}
                //else
                //{
                //    multiidioma = true;
                //}

            }
            
        }

        // NAM - Inicializa variables de idioma segun el idioma del usuario configurado por base.
        private void determinarIdioma(String Usuario,String Base)
        {
     /*
 
            string cn = DAL.constr(Base);
            DataSet idiomaAux = new DataSet();
            String sql;

            //Busco el idioma del usuario

            sql = "SELECT lencod FROM user_per INNER JOIN lenguaje ON lenguaje.lennro = user_per.lennro ";
            sql = sql + " WHERE UPPER(iduser) = '" + Usuario.ToUpper() + "'";
             
            
            da = new OleDbDataAdapter(sql, cn);

            try
            {
                da.Fill(idiomaAux);
                idiomausuario2 = idiomaAux.Tables[0].Rows[0].ItemArray[0].ToString().Replace("-", "");
                idiomausuario = idiomausuario2.Substring(0, 2) + "-" + idiomausuario2.Substring(2, 2);
              
            }
            catch (Exception ex)
            {
              
                idiomausuario2 = "esAR";
                idiomausuario = "es-AR";
            }
       */
            idiomausuario2 = getIdiomaUsuario(Usuario,Base);
            idiomausuario = idiomausuario2;
        }

        private string getIdiomaUsuario(String Usuario, String Base)
        {
            string salida = "esAR";
             
            String sql;

            //Busco el idioma del usuario
            sql = "SELECT lencod FROM user_per INNER JOIN lenguaje ON lenguaje.lennro = user_per.lennro ";
            sql = sql + " WHERE UPPER(iduser) = Upper('" + Usuario + "')";

            try
            {
                DataTable dt = get_DataTable(sql,Base);
             
                if (dt.Rows.Count > 0)
                    salida = Convert.ToString(dt.Rows[0]["lencod"]);
                else
                    salida = "esAR";
            }
            catch  { }

            return salida;
        }

        //-------------------------------------------------------------------------------------
        //-------------------------------------------------------------------------------------
        [WebMethod(Description = "Devuelve la version de Servicio Web")]
        public string VersionWS()
        {
            string Salida = "";

            Assembly assem = Assembly.GetExecutingAssembly();
            AssemblyName aName = assem.GetName();

            string version = aName.Version.ToString();


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

            //Salida = "Version 1.11: Se aplico multiidioma.";

            //Salida = "Version 1.12: Gonzalez Nicolás - Se valida si la contraseña es correcta e inserto login fallido en //Pruebo la conexion con los datos del usuario";
            //Salida += " Se agregó mensaje de error 'La Contraseña coincide con una histórica.' en //Control de historia de la contraseña";

            //Salida = "Version 1.13: 09/04/2012 - Gonzalez Nicolás - Se agregó combo con lista de paises + Mensajes de Error en MI ";
            //Se corrigieron errores en consultas de ORACLE
            //Se creo fcn ComboBanderasMI() Devuelve lenguaje y nombre de pais para el combo del Home
            //Se corrigio SQL en función Consultas.Search() para que traduzca el contenido del menu en el idioma correspondiente.

            //Salida = "Version 1.14: 31/10/2012 - Gonzalez Nicolás - Se buscan traducciones de menudetalle con lenguaje_etiqueta";

            //Salida = "Version 1.15: 04/07/2013 - Brzozowski Juan Pablo - ";
            //Salida += " Se agregó el método publico Info_Dias_VencimientoPass, el cual retorna la cantidad de días que faltan para que ";
            //Salida += " a un usuario le caduque la contraseña.";

            //Salida = "Version 1.16: 21/08/2013 - Deluchi Ezequiel - 19483 ";
            //Salida += " Se agregaron metodos para controlar configuraciones de la politica de el empleado en primer logueo. ";
            //Salida += " Controles en el password. Que tengan al menos un caracter especial, una mayuscula, una minuscula y/o un numero";

            //Salida = "Version 1.17: 22/10/2013 - Brzozowski Juan Pablo - ";
            //Salida += " Se creó nuevamente el metodo TablaPlanaErecruiting, ya que este se habia perdido en versiones anteriores. ";
            //Salida += " ";



            ////1.18            
            //Salida = "Version 1.18 -Carlos Masson - ";
            //Salida += "[06/01/2014 - Carlos Masson - 20566 - Se agregan parametros para usuario administrador con posibilidad de actualizar periodicamente la clave del mismo automáticamente.]";
            //Salida += "[Mejoras en la performance con LDAP. (Requiere configurar el usuario del directorio con el que se realiza la búsqueda)]";
            //Salida += "[Se agrega funcionalidad para reportar eventos de sistema] ";


            //1.19           
            //Salida = "Version 1.19 - Mauricio Zwenger - ";
            //Salida += "[CAS-20270 - H&A - Otimización de traducción de etiquetas y Menu General]";
            //Salida += "[Se fuerza el collation 'Modern_Spanish_CI_AS' en la comparacion con el campo etiqueta de la tabla Lenguaje_Etiqueta], en los metodos Modulos, MRU y Search";


            //Salida = "Version 1.20 - Brzozowski Juan Pablo - ";
            //Salida += " [CAS-25445 - Heidt & Asoc. - Bugs SEG] ";
            //Salida += " [Se modifico el contenido de los metodos Mensaje y Banner para que respeten las fechas de creación y vencimiento] ";

            //Salida = "Version 1.21 - Brzozowski Juan Pablo - 12/08/2014 - ";
            //Salida += " [CAS-26719 - 5CA - BUG EN INGRESO A LA APLICACION] ";
            //Salida += " [Se modifico la funcion Modulos. Contempla el caso que el idioma no se contemple en la tabla lenguaje_etiqueta] ";



            //Salida = "Version 1.22: 28/08/2014 - Brzozowski Juan Pablo -  CAS-13764 - H&A - Mejoras MRU ";
            //Salida += " - Se modificó el metodo Mensajes para que filtre por fecha de vencimiento y tipo de mensajes (rhpro). ";
            //Salida += " - Se agregó el metodo Armar_Diccionario_MRU para armar un diccionario que contenga la cantidad mru de cada modulo";
            //Salida += " - Se agregó el metodo Armar_Diccionario_Ponderacion para armar un diccionario que contenga la ponderacion del un determinado modulo.";            
            //Salida += " - Se modificó el metodo Modulos para agregarles las columnas AccesosMRU y menuponderacion, en las cuales se especifican los accesos mru y la ponderacion del modulo respectivamente.";

//            Salida = "Version 1.23: 14/10/2014 - Brzozowski Juan Pablo -   CAS-20903 - Heidt & Asoc. - Ingreso clientes desde MetaHome [Entrega 4]   ";
//            Salida += " - Se agregó el metodo Controlar_Gadget_PrimerAcceso para asignar modulos activos faltantes al loguearse";

            //Salida = "Version 1.24: 22/10/2014 - Brzozowski Juan Pablo -   CAS-20903 - Heidt & Asoc. - Ingreso clientes desde MetaHome [Entrega 5]   ";
            //Salida += " - Se reemplazo el metodo Controlar_Gadget_PrimerAcceso  por Controlar_Gadget_EnLoguin";


            //Salida = "Version 1.25: 23/10/2014 - Brzozowski Juan Pablo -   CAS-26972 - H&A - Bugs detectados en R4 - Error en la configuración de Grupo de restricciones   ";
            //Salida += " - Se agregaron las funciones para el CONTROL DE ACCESO  TEMPORAL Y POR ARMADO DE MENU ";

            //Salida = "Version 1.26: 06/02/2015 - Brzozowski Juan Pablo -   CAS-26028 - H&A - Mejora en el idioma del home por usuario  ";
            //Salida += " - Se agrego la funcion Cambiar_Idioma para modificar el idioma del usuario logueado";
            //Salida += " - En el metodo Modulos, se comprueba que el menu tenga el campo menulicenciado en -1 como condición de habilitado";
            //Salida += " - Se agrega mas control y log de eventos al metodo Controlar_Gadget_EnLoguin";

            //Salida = "Version 1.27: 18/05/2015 - Brzozowski Juan Pablo -  CAS-30306 - H&A - BUG indicadores  ";
            //Salida += " - En el metodo Controlar_Gadget_EnLoguin se quito una subconsulta anidada que daba problemas con ciertos motores SQL ";

            //Salida = "Version 1.28: 21/05/2015 - Brzozowski Juan Pablo -  CAS-19713 - H&A - Estandarización ICI (CAS-15298) [Entrega 5]  ";
            //Salida += " - En el metodo TablaPlana y TablaPlanaErecruiting se agrego una nueva tabla de control ";

            //Salida = "Version 1.29: 31/07/2015 - Brzozowski Juan Pablo -  CAS-24137 - H&A - Calidad - Funcionalidad - Seguridad - PenetrationTest - Hito 3  ";
            //Salida += " - Se agrega bloqueo por IP al realizar varios intentos fallidos de login.  ";
            //Salida += " - El caso CAS-20270 - H&A - NewHome - Oracle se soluciona tambien en esta versión";

            //Salida = "Version 1.30: 15/09/2015 - Brzozowski Juan Pablo -  CAS-33015 - TATA - Error en cambio de contraseña  ";
            //Salida += " - Cuando se usa Oracle como motor de base, daba error al querer hacer un alter sobre un usuario formado con puntos '.' ";
 
            Salida = "Version 1.31: 14/12/2015 - Brzozowski Juan Pablo -  CAS-32645 - Raet Latam - Errores Meta Home  ";
            Salida += " -  Se mejora la performance de conexión en el metodo Login.  ";
 

            

            //CDM - Dejar la siguiente línea como está
            return "Version " + version + ": " + Salida;
                                  
        }

        //-------------------------------------------------------------------------------------
        //-------------------------------------------------------------------------------------
        
        [WebMethod(Description = "(HOME) Devuelve la version del sistema.")]
        public string Version(string Base, string Idioma)
        {
            //Defino un objeto tipo DAL
            //DAL MiDAL = new DAL();   

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
            //Defino un objeto tipo DAL
            //DAL MiDAL = new DAL();   

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
            //Defino un objeto tipo DAL
            //DAL MiDAL = new DAL();   
 
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
            //Defino un objeto tipo DAL
            //DAL MiDAL = new DAL();   

            string cn = DAL.constr(Base);
            string sql;
            DataSet ds = new DataSet();

            //sql = "SELECT hmsjtitulo, hmsjcuerpo ";
            //sql = sql + "FROM home_mensaje ";
            //sql = sql + "WHERE home_mensaje.hmsjactivo = -1 ";
            //sql = sql + "ORDER BY home_mensaje.hmsjnro ";

            //JPB - 14/07/2014 - Se modifica la consulta para que respete la fecha de inicio y vencimiento del mensaje 
            sql = "SELECT hmsjtitulo, hmsjcuerpo ";
            sql = sql + " FROM home_mensaje ";
            sql = sql + " WHERE home_mensaje.hmsjactivo = -1 ";
            if (DAL.TipoBase(Base).ToUpper() == "MSSQL")
                sql = sql + "  AND ( ( GETDATE()>=home_mensaje.hmsjfecalta AND GETDATE()<=hmsjfecvto)   AND rhpro=-1  ) ";
            else
                sql = sql + "  AND ( ( (SELECT SYSDATE  FROM DUAL)>=home_mensaje.hmsjfecalta AND (SELECT SYSDATE  FROM DUAL)<=hmsjfecvto)   AND rhpro=-1  ) ";            


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
            //Defino un objeto tipo DAL
            //DAL MiDAL = new DAL();   

            string cn = DAL.constr(Base);
            string sql;
            DataSet ds = new DataSet();

            //JPB - 14/07/2014 - Se modifica la consulta para que respete la fecha de inicio y vencimiento del banner
            sql = "SELECT hbandesc, hbanimage, hbandescext ";
            sql = sql + "FROM home_banner ";
            sql = sql + "WHERE home_banner.hbanactivo = -1 ";
            if (DAL.TipoBase(Base).ToUpper() == "MSSQL")
                sql = sql + "   AND ( ( GETDATE()>=home_banner.hbanfecalta AND GETDATE()<=home_banner.hbanfecvto)   AND rhpro=-1  )";
            else
                sql = sql + "   AND ( ( (SELECT SYSDATE  FROM DUAL)>=home_banner.hbanfecalta AND (SELECT SYSDATE  FROM DUAL)<=home_banner.hbanfecvto)   AND rhpro=-1  )";
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


        /// <summary>
        /// JPB: 09/06/2014 - Arma un diccionario con la cantidad de accesos de cada menuraiz
        /// </summary>
        /// <param name="Base"></param>
        /// <param name="Usuario"></param>
        private void Armar_Diccionario_MRU(String Base, String Usuario)
        {
            string cn = DAL.constr(Base);
            string sql;
            string Modulo = "";
            string Salida = "0";
            DataSet ds = new DataSet();


            sql = "   SELECT   menuraiz.menudir , (  COUNT(menuraiz.menudir) + menuraiz.menuponderacion) AccesosMRU  ";            
            sql = sql + " FROM mru ";
            sql = sql + " INNER JOIN menumstr ON menumstr.menumsnro = mru.menumsnro ";
            sql = sql + " INNER JOIN menuraiz ON menuraiz.menunro = mru.menuraiz ";
            sql = sql + " WHERE UPPER(mru.iduser) = Upper('" + Usuario + "')  ";      
            sql = sql + " group by  menuraiz.menudir,menuraiz.menuponderacion ";
            

            da = new OleDbDataAdapter(sql, cn);
            try
            {

                da.Fill(ds);
                foreach (DataRow fila in ds.Tables[0].Rows)
                {
                    // switch (Convert.ToString(fila["menudir"]))//***DESHABILITAR PARA CAMBIO MENU****///
                    // {
                    //    case "ADP": Modulo = "ADMPER"; break;
                    //    case "ALE": Modulo = "ALERTAS"; break;
                    //    case "ANR": Modulo = "ANALISIS"; break;
                    //    case "BDC": Modulo = "BIENES"; break;
                    //    case "CAP": Modulo = "CAPACITACION"; break;
                    //    case "POST": Modulo = "EMPLEOS"; break;
                    //    case "GTI": Modulo = "GTI"; break;
                    //    case "EVAL": Modulo = "EVALUACION"; break;
                    //    case "LIQ": Modulo = "LIQUIDACION"; break;
                    //    case "PDD": Modulo = "PLAN"; break;
                    //    case "POL": Modulo = "POLITICAS"; break;
                    //    case "SO": Modulo = "SALUD"; break;
                    //    case "SUP": Modulo = "SUPERVISOR"; break;
                    //    case "DIS": Modulo = "DIS"; break;
                    //    case "BIE": Modulo = "BIENESTAR"; break;
                    //    case "PP": Modulo = "PLANTA"; break;
                    //    case "EMB": Modulo = "EMBARGOS"; break;
                    //    case "SIM": Modulo = "SIM"; break;
                    //    case "GDC": Modulo = "COMPETENCIAS"; break;
                    //    case "MIG": Modulo = "INFOGER"; break;                             
                    //} 

                    //--
                     Modulo = Convert.ToString(fila["menudir"]);//***HABILITAR PARA CAMBIO MENU****///

                    if (!DiccionarioMRU.ContainsKey(Modulo))
                        DiccionarioMRU[Modulo] = Convert.ToInt32(fila["AccesosMRU"]);
                     
                } 

            }
            catch (Exception ex)
            {
                throw ex;
            } 
        }

        /// <summary>
        /// JPB: 09/06/2014 - Arma un diccionario con la ponderacion de cada menuraiz
        /// </summary>
        /// <param name="Base"></param>
        /// <param name="Usuario"></param>
        private void Armar_Diccionario_Ponderacion(String Base, String Usuario)
        {
            string cn = DAL.constr(Base);
            string sql;
            string Modulo = "";            
            DataSet ds = new DataSet();
            
            sql = " SELECT menudir,menuponderacion FROM  menuraiz  ";
            da = new OleDbDataAdapter(sql, cn);
            try
            {
                da.Fill(ds);
                foreach (DataRow fila in ds.Tables[0].Rows)
                {

                    //switch (Convert.ToString(fila["menudir"]))//***DESHABILITAR PARA CAMBIO MENU****///
                     
                    //{
                    //    case "ADP": Modulo = "ADMPER"; break;
                    //    case "ALE": Modulo = "ALERTAS"; break;
                    //    case "ANR": Modulo = "ANALISIS"; break;
                    //    case "BDC": Modulo = "BIENES"; break;
                    //    case "CAP": Modulo = "CAPACITACION"; break;
                    //    case "POST": Modulo = "EMPLEOS"; break;
                    //    case "GTI": Modulo = "GTI"; break;
                    //    case "EVAL": Modulo = "EVALUACION"; break;
                    //    case "LIQ": Modulo = "LIQUIDACION"; break;
                    //    case "PDD": Modulo = "PLAN"; break;
                    //    case "POL": Modulo = "POLITICAS"; break;
                    //    case "SO": Modulo = "SALUD"; break;
                    //    case "SUP": Modulo = "SUPERVISOR"; break;
                    //    case "DIS": Modulo = "DIS"; break;
                    //    case "BIE": Modulo = "BIENESTAR"; break;
                    //    case "PP": Modulo = "PLANTA"; break;
                    //    case "EMB": Modulo = "EMBARGOS"; break;
                    //    case "SIM": Modulo = "SIM"; break;
                    //    case "GDC": Modulo = "COMPETENCIAS"; break;
                    //    case "MIG": Modulo = "INFOGER"; break;    
                    //}

                  
                    //--
                     Modulo = Convert.ToString(fila["menudir"]);//***HABILITAR PARA CAMBIO MENU****///


                    if (!DiccionarioPonderacion.ContainsKey(Modulo))
                        DiccionarioPonderacion[Modulo] = Convert.ToInt32(fila["menuponderacion"]);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

      

        //-------------------------------------------------------------------------------------
        //-------------------------------------------------------------------------------------
        [WebMethod(Description = "(HOME) Devuelve los modulos segun el usuario.")]
        public DataTable Modulos(string Usuario, string Base, string Idioma)
        {

            //Defino un objeto tipo DAL
            //DAL MiDAL = new DAL();   
            string TipoDB = DAL.TipoBase(Base);

            string cn = DAL.constr(Base);
            string sql;
            DataSet ds = new DataSet();
            DataSet dsAux = new DataSet();
            DataSet dsAux2 = new DataSet();
            DataTable tablaAux;
            string Access = "";
            String Collate = "";
            string[] arrAccess;
            string[] arrPerfUser;
            string listaPerfUser = "";
            bool Hay = false;
            DataColumn Columna;
            DataRow filaAux;


            DataSet idiomaAux = new DataSet();

            String idio;

            determinarMI(Base);

            if (Idioma == "")
            {
                //Busco el idioma del usuario
                determinarIdioma(Usuario, Base);
            }
            else
            {
                idiomausuario = Idioma;
                idiomausuario2 = Idioma.Replace("-", "");
            }

            if (idiomausuario.Substring(0, 2) == "es" || idiomausuario.Substring(0, 2) == "AR")
            {
                idio = "";
            }
            else
            {
                idio = idiomausuario;
            }

            //Verifica si trata de acceder a una columna que no existe en lenguaje_etiqueta
            if (TipoDB == "MSSQL")
            {
                //me fijo si existe la columna del idioma seleccionado por el usuario                
                sql = "SELECT * FROM INFORMATION_SCHEMA.COLUMNS AS c1 where c1.column_name = '" + idio.Replace("-", "") + "'";
                sql = sql + " and c1.table_name = 'lenguaje_etiqueta'";

                da = new OleDbDataAdapter(sql, cn);

                try
                {
                    da.Fill(idiomaAux);
                    //JPB: Si el campo de idioma no existe, entonces el idioma por defecto es esAR
                    if (idiomaAux.Tables[0].Rows.Count == 0)
                    {
                        idio = "";
                        idiomausuario2 = "esAR";
                    } 
                }
                catch (Exception ex)
                {                   
                    idiomausuario2 = "esAR";
                }               
            }
            else//ORACLE
            {
                //me fijo si existe la columna del idioma seleccionado por el usuario              
                sql = "SELECT " + idio.Replace("-", "") + " FROM lenguaje_etiqueta ";

                da = new OleDbDataAdapter(sql, cn);

                try
                {
                    da.Fill(idiomaAux);

                    //JPB: Si viene vacio entonces no existe la columna de idioma
                    if (idiomaAux.Tables[0].Rows.Count == 0)
                        idiomausuario2 = "esAR";
                }
                catch (Exception ex)
                {// si no existe la columna... entonces voy a buscar la descripcion en español por defecto.                   
                    idio = "";
                    idiomausuario2 = "esAR";
                } 
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

            Columna = new DataColumn();
            Columna.DataType = System.Type.GetType("System.String");
            Columna.ColumnName = "menuname";
            Columna.AutoIncrement = false;
            Columna.Unique = false;
            tablaSalida.Columns.Add(Columna);
            /*** Nuevos **/
            Columna = new DataColumn();
            Columna.DataType = System.Type.GetType("System.String");
            Columna.ColumnName = "menumsnro";
            Columna.AutoIncrement = false;
            Columna.Unique = false;
            tablaSalida.Columns.Add(Columna);

            Columna = new DataColumn();
            Columna.DataType = System.Type.GetType("System.String");
            Columna.ColumnName = "menuraiz";
            Columna.AutoIncrement = false;
            Columna.Unique = false;
            tablaSalida.Columns.Add(Columna);

            Columna = new DataColumn();
            Columna.DataType = System.Type.GetType("System.Int32");
            Columna.ColumnName = "AccesosMRU";
            Columna.AutoIncrement = false;
            Columna.Unique = false;
            tablaSalida.Columns.Add(Columna);

            Columna = new DataColumn();
            Columna.DataType = System.Type.GetType("System.Int32");
            Columna.ColumnName = "menuponderacion";
            Columna.AutoIncrement = false;
            Columna.Unique = false;
            tablaSalida.Columns.Add(Columna);
            
            //Armo los diccionarios para los Menues Mas Usados
            Armar_Diccionario_Ponderacion(Base, Usuario);
            Armar_Diccionario_MRU(Base, Usuario);
            //

            if (TipoDB == "MSSQL")
                Collate = " COLLATE Modern_Spanish_CS_AS ";


            if (Usuario.Trim().Length == 0)//Este es el caso en el que el usuario no se ha logueado.
            {
                
                sql = "SELECT menudesabr,'' action,menuobjetivo,menubeneficio,linkmanual,linkdvd,menuname,menudetalle,menumsnro,menuraiz";
                       
                if (multiidioma == true)
                {
                    if (idiomausuario2 == "")
                    {
                        idiomausuario2 = "esAR";
                    }
                    sql = sql + ",lenguaje_Etiqueta." + idiomausuario2.Replace("-", "") + " TituloMI";
                    sql = sql + ",le." + idiomausuario2.Replace("-", "") + " menudetalleDescMI";
                    sql = sql + " FROM menumstr";
                    sql = sql + " LEFT JOIN lenguaje_etiqueta ON lenguaje_etiqueta.etiqueta = menumstr.menudesabr " + Collate + " AND (lenguaje_etiqueta.pagina is null OR lenguaje_etiqueta.pagina = '') AND (lenguaje_etiqueta.modulo is null OR lenguaje_etiqueta.modulo = '') ";
                    sql = sql + " LEFT JOIN lenguaje_etiqueta LE ON LE.etiqueta = menumstr.menuname  " + Collate + " AND LE.modulo = 'HOME' ";
                    //sql = sql + " LEFT JOIN lenguaje_etiqueta LE ON LE.etiqueta = menumstr.menuname  AND (LE.pagina is null OR LE.pagina = '') AND (LE.modulo is null OR LE.modulo = 'HOME') ";
                    //sql = sql + "," + idiomausuario2.Replace("-", "") + " tituloMI ";
                    //sql = sql + " FROM menumstr ";
                    //sql = sql + " LEFT JOIN lenguaje_etiqueta ON lenguaje_etiqueta.etiqueta = menumstr.menudesabr AND (lenguaje_etiqueta.pagina is null OR lenguaje_etiqueta.pagina = '') AND (lenguaje_etiqueta.modulo is null OR lenguaje_etiqueta.modulo = '') ";
                }
                else
                {
                    sql = sql + " FROM menumstr ";
                }

                //---------
                sql = sql + " WHERE menuraiz = 74 ";
                sql = sql + " AND menuactivo = -1 ";
                //sql = sql + " AND menumstr.menuname <> 'HOME' ";
                sql = sql + " ORDER BY menudesabr ";

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
                    tablaAux = ds.Tables[0];

                    foreach (DataRow fila in tablaAux.Rows)
                    {
                        filaAux = tablaSalida.NewRow();
                        if (multiidioma == true)
                        {
                            if (fila["tituloMI"].ToString().Length == 0)
                            {
                                filaAux["menudesabr"] = fila["menudesabr"].ToString();
                            }
                            else
                            {
                                filaAux["menudesabr"] = fila["tituloMI"].ToString();
                            }

                            if (fila["menudetalleDescMI"].ToString().Length == 0)
                            {
                                filaAux["menudetalle"] = fila["menudetalle"].ToString();
                            }
                            else
                            {
                                filaAux["menudetalle"] = fila["menudetalleDescMI"].ToString();

                            }



                        }
                        else
                        {
                            filaAux["menudesabr"] = fila["menudesabr"].ToString();
                            filaAux["menudetalle"] = fila["menudetalle"].ToString();

                        }

                        //filaAux["menudetalle"] = fila["menudetalle" + idio.ToUpper()].ToString();

                        //filaAux["menudetalle"] = fila["menuname"].ToString();

                        filaAux["menuobjetivo"] = fila["menuobjetivo"].ToString();
                        filaAux["menubeneficio"] = fila["menubeneficio"].ToString();
                        filaAux["linkmanual"] = fila["linkmanual"].ToString();
                        filaAux["linkdvd"] = fila["linkdvd"].ToString();
                        filaAux["menuname"] = fila["menuname"].ToString();
                        filaAux["action"] = "";
                        filaAux["menumsnro"] = fila["menumsnro"].ToString();
                        filaAux["menuraiz"] = fila["menuraiz"].ToString();                        
                        filaAux["AccesosMRU"] = (DiccionarioMRU.ContainsKey(Convert.ToString(fila["menuname"])))? DiccionarioMRU[Convert.ToString(fila["menuname"])] : 0;
                        filaAux["menuponderacion"] = (DiccionarioPonderacion.ContainsKey(Convert.ToString(fila["menuname"]))) ? DiccionarioPonderacion[Convert.ToString(fila["menuname"])] : 0;

                         
                        //Inserto la fila en la tabla de salida

                        tablaSalida.Rows.Add(filaAux);
                    }

                }
                return tablaSalida;
                //  return ds.Tables[0];
            }
            else
            {

                //Busco el perfil del usuario
                sql = "SELECT listperfnro ";
                sql = sql + " FROM user_perfil ";
                sql = sql + " WHERE UPPER(user_perfil.iduser) = '" + Usuario.ToUpper() + "' ";
                sql = sql + " UNION ALL ";
                sql = sql + " SELECT listperfnro from bk_perfil INNER JOIN bk_cab ON bk_cab.bkcabnro = bk_perfil.bkcabnro ";
                sql = sql + " AND (bk_cab.fdesde <= " + Fecha.cambiaFecha(DateTime.Today.ToString("dd/MM/yyyy"), TipoDB) + " AND (bk_cab.fhasta IS NULL OR bk_cab.fhasta >= " + Fecha.cambiaFecha(DateTime.Today.ToString("dd/MM/yyyy"), TipoDB) + " )) ";
                sql = sql + " AND upper(bk_perfil.iduser) = '" + Usuario.ToUpper() + "'";

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
                    listaPerfUser = "0";
                    foreach (DataRow f in dsAux.Tables[0].Rows)
                    {
                        listaPerfUser = listaPerfUser + "," + Convert.ToString(f.ItemArray[0]);
                    }

                    //Busco todos los menu que tienen al perfil
                    sql = "SELECT menuaccess, menuname, ";
                    sql = sql + "menudesabr , menudetalle detalleMI, ";
                    //if (DAL.TipoBase(Base).ToUpper() == "MSSQL")
                    //{
                    //    //sql = sql + "[menudetalle" + idio.ToUpper() + "], ";
                    //}
                    //else
                    //{
                    //    //ORACLE
                    //    //sql = sql + "\"MENUDETALLE" + idio.ToUpper() + "\", ";
                    //}

                    /////////////sql = sql + "'abrirVentana(' + CHAR(39) + action + CHAR(39) + ','''',670,520)' action, ";
                    sql = sql + "'abrirVentana(' action1, ";
                    sql = sql + "action action2, ";
                    sql = sql + "',670,520)' action3, ";
                    sql = sql + "menuobjetivo, ";
                    sql = sql + "menubeneficio, ";
                    sql = sql + "linkmanual, ";
                    sql = sql + "linkdvd, ";
                    sql = sql + "menumsnro,";
                    sql = sql + "menuraiz "; 
                    if (multiidioma == true)
                    {
                        if (idiomausuario2 == "")
                        {
                            idiomausuario2 = "esAR";
                        }
                        //JPB: 06/03/2015 - Se agrega control del menulicenciado en menuraiz
                        
                        sql = sql + ",lenguaje_Etiqueta." + idiomausuario2.Replace("-", "") + " TituloMI";
                        sql = sql + ",le." + idiomausuario2.Replace("-", "") + " menudetalleDescMI";
                        sql = sql + ",M.menulicenciado ";
                        sql = sql + " FROM menumstr ";
                        sql = sql + " LEFT JOIN lenguaje_etiqueta ON lenguaje_etiqueta.etiqueta = menumstr.menudesabr " + Collate + " AND (lenguaje_etiqueta.pagina is null OR lenguaje_etiqueta.pagina = '') AND (lenguaje_etiqueta.modulo is null OR lenguaje_etiqueta.modulo = '') ";
                        sql = sql + " LEFT JOIN lenguaje_etiqueta LE ON LE.etiqueta = menumstr.menuname " + Collate + " AND LE.modulo = 'HOME' ";
                        sql = sql + " INNER  JOIN menuraiz M on UPPER(M.menudesc) = UPPER(menumstr.menuname) ";
                    }
                    else
                    {
                        sql = sql + "FROM menumstr ";
                    }


                    sql = sql + " WHERE menuraiz = 74 ";
                    //sql = sql + " AND menumssqlactivo = -1 ";
                    sql = sql + " AND menuactivo = -1 ";
                    sql = sql + " AND menumstr.action <> '#' ";
                    //sql = sql + "AND menumstr.action <> '' ";
                    sql = sql + " AND menumstr.action IS NOT NULL ";
                    sql = sql + " ORDER BY menudesabr ";

  

                    da = new OleDbDataAdapter(sql, cn);

                    try
                    {
                        da.Fill(dsAux2);
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }

                    //Ciclo por cada modulo

                    if (dsAux2.Tables[0].Rows.Count > 0)
                    {
                        tablaAux = dsAux2.Tables[0];

                        foreach (DataRow fila in tablaAux.Rows)
                        {
                            //Copio todas las filas menos la de access que depende del perfil

                            filaAux = tablaSalida.NewRow();
                            if (multiidioma == true)
                            {
                                if (fila["tituloMI"].ToString().Length == 0)
                                {
                                    filaAux["menudesabr"] = fila["menudesabr"].ToString();
                                }
                                else
                                {
                                    filaAux["menudesabr"] = fila["tituloMI"].ToString();
                                }

                                if (fila["menudetalleDescMI"].ToString().Length == 0)
                                {
                                    filaAux["menudetalle"] = fila["detalleMI"].ToString();
                                }
                                else
                                {
                                    filaAux["menudetalle"] = fila["menudetalleDescMI"].ToString();

                                }

                            }
                            else
                            {
                                filaAux["menudesabr"] = fila["menudesabr"].ToString();
                                filaAux["menudetalle"] = fila["detalleMI"].ToString();
                            }
 

                            filaAux["menuobjetivo"] = fila["menuobjetivo"].ToString();
                            filaAux["menubeneficio"] = fila["menubeneficio"].ToString();
                            filaAux["linkmanual"] = fila["linkmanual"].ToString();
                            filaAux["linkdvd"] = fila["linkdvd"].ToString();

                            filaAux["menuname"] = fila["menuname"].ToString();

                            Access = Convert.ToString(fila["menuaccess"].ToString());
                            Hay = false;

                            //Por cada perfil del usuario
                            /*
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
                              */

                            /***************************************************************************************************************/
                            /*JPB: Comprueba si el menu esta habilitado por armado de menu o por grupo de acceso*/
                            /*     También comprueba que el menu tenga el campo menulicenciado en -1*/
                             Hay  =  Menu_Habilitado(Convert.ToString(fila["menuaccess"]), Convert.ToInt32(fila["menumsnro"]), Usuario, Base) 
                               && (Convert.ToInt32(fila["menulicenciado"]) == -1);
                            /***************************************************************************************************************/

                            if (Hay)
                                filaAux["action"] = fila["action1"].ToString() + "'" + fila["action2"].ToString() + "',''" + fila["action3"].ToString();
                            else
                                filaAux["action"] = "";

                            filaAux["menumsnro"] = fila["menumsnro"].ToString();
                            filaAux["menuraiz"] = fila["menuraiz"].ToString();
                            //filaAux["AccesosMRU"] = Cantidad_Accesos_Al_Modulo(Base, Usuario, fila["menuname"].ToString());
                            filaAux["AccesosMRU"] = (DiccionarioMRU.ContainsKey(Convert.ToString(fila["menuname"]))) ? DiccionarioMRU[Convert.ToString(fila["menuname"])] : 0;
                            filaAux["menuponderacion"] = (DiccionarioPonderacion.ContainsKey(Convert.ToString(fila["menuname"]))) ? DiccionarioPonderacion[Convert.ToString(fila["menuname"])] : 0;
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

                    sql = "SELECT menudesabr,'' action, menudetalle detalleMI, menuobjetivo,menubeneficio,linkmanual,linkdvd,menuname,menumsnro,menuraiz ";
                    //if (DAL.TipoBase(Base).ToUpper() == "MSSQL")
                    //{
                    //    sql = sql + ",[menudetalle" + idio.ToUpper() + "] ";
                    //}
                    //else
                    //{
                    //    //ORACLE
                    //    sql = sql + ",\"MENUDETALLE" + idio.ToUpper() + "\" ";
                    //}


                    if (multiidioma == true)
                    {
                        if (idiomausuario2 == "")
                        {
                            idiomausuario2 = "esAR";
                        }
                        //sql = sql + "," + idiomausuario2.Replace("-", "") + " tituloMI ";
                        sql = sql + ",lenguaje_Etiqueta." + idiomausuario2.Replace("-", "") + " TituloMI";
                        sql = sql + ",le." + idiomausuario2.Replace("-", "") + " menudetalleDescMI";
                        sql = sql + " FROM menumstr ";
                        sql = sql + " LEFT JOIN lenguaje_etiqueta ON lenguaje_etiqueta.etiqueta = menumstr.menudesabr " + Collate + " AND (lenguaje_etiqueta.pagina is null OR lenguaje_etiqueta.pagina = '') AND (lenguaje_etiqueta.modulo is null OR lenguaje_etiqueta.modulo = '') ";
                        sql = sql + " LEFT JOIN lenguaje_etiqueta LE ON LE.etiqueta = menumstr.menuname " + Collate + " AND LE.modulo = 'HOME' ";
                    }
                    else
                    {
                        sql = sql + " FROM menumstr ";
                    }
                    sql = sql + " WHERE menuraiz = 74 ";
                    sql = sql + " AND menuactivo = -1 ";
                    sql = sql + " ORDER BY menudesabr ";

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
                        tablaAux = ds.Tables[0];

                        foreach (DataRow fila in tablaAux.Rows)
                        {
                            filaAux = tablaSalida.NewRow();
                            if (multiidioma == true)
                            {
                                if (fila["tituloMI"].ToString().Length == 0)
                                {
                                    filaAux["menudesabr"] = fila["menudesabr"].ToString();
                                }
                                else
                                {
                                    filaAux["menudesabr"] = fila["tituloMI"].ToString();
                                }

                                if (fila["menudetalleDescMI"].ToString().Length == 0)
                                {
                                    filaAux["menudetalle"] = fila["detalleMI"].ToString();
                                }
                                else
                                {
                                    filaAux["menudetalle"] = fila["menudetalleDescMI"].ToString();

                                }

                            }
                            else
                            {
                                filaAux["menudesabr"] = fila["menudesabr"].ToString();
                                filaAux["menudetalle"] = fila["detalleMI"].ToString();
                            }

                            //filaAux["menudetalle"] = fila["menudetalle" + idio.ToUpper()].ToString();

                            //if (filaAux["menudetalle"] == "")
                            //{
                            //    filaAux["menudetalle"] = fila["detalleMI"];
                            //}

                            filaAux["menuobjetivo"] = fila["menuobjetivo"].ToString();
                            filaAux["menubeneficio"] = fila["menubeneficio"].ToString();
                            filaAux["linkmanual"] = fila["linkmanual"].ToString();
                            filaAux["linkdvd"] = fila["linkdvd"].ToString();
                            filaAux["action"] = "";

                            filaAux["menuname"] = fila["menuname"].ToString();

                            filaAux["menumsnro"] = fila["menumsnro"].ToString();
                            filaAux["menuraiz"] = fila["menuraiz"].ToString();
                            //filaAux["AccesosMRU"] = Cantidad_Accesos_Al_Modulo(Base, Usuario, fila["menuname"].ToString());
                            filaAux["AccesosMRU"] = (DiccionarioMRU.ContainsKey(Convert.ToString(fila["menuname"]))) ? DiccionarioMRU[Convert.ToString(fila["menuname"])] : 0;
                            filaAux["menuponderacion"] = (DiccionarioPonderacion.ContainsKey(Convert.ToString(fila["menuname"]))) ? DiccionarioPonderacion[Convert.ToString(fila["menuname"])] : 0;

                            //Inserto la fila en la tabla de salida

                            tablaSalida.Rows.Add(filaAux);
                        }

                    }
                    return tablaSalida;
                }
            }
        }

        /// <summary>
        /// Si tiene datos la consulta retorna -1; Si no tiene datos retorna 0; Si la consulta esta mal armada retorna 1
        /// </summary>
        /// <param name="sql"></param>
        /// <param name="Base"></param>
        /// <returns></returns>
        private int TieneDatos(String sql, String Base)
        {
            int salida;

            try{
                DataTable dt = get_DataTable(sql, Base);
                if (dt.Rows.Count > 0)
                    salida = -1;//La consulta tiene datos
                else
                    salida = 0;//Esta bien armada la consulta pero no tiene datos   
            }
            catch 
            {
                salida = 1;//La consulta esta mal armada o no existe alguna tabla o vista dentro de la misma.
            }

            return salida;
        }

        [WebMethod(Description = "(HOME) Cambia el estilo del home.")]
        public void Cambiar_Estilo(string User, string Base, string idcarpetaestilo,  string codestilo)
        {
            
            string TipoDB = DAL.TipoBase(Base);
            OleDbConnection cn = new OleDbConnection();
            cn.ConnectionString = DAL.constr(Base);
            String sql="";

            //sql = " if not exists(select estiloactivo from estilos_home_user where Upper(iduser)=Upper('" + User + "') ) ";
            //sql += " insert into estilos_home_user (iduser, estiloactivo) values ('"+User+"'," + idcarpetaestilo + "); ";
            //sql += " else ";
            //sql += " UPDATE estilos_home_user set estiloactivo = " + idcarpetaestilo + " where iduser='" + User + "'; ";
           // sql += " UPDATE   estilo_homex2 SET activo=0;  ";
           // sql += " UPDATE   estilo_homex2 SET activo=-1 WHERE codestilo= " + codestilo+";";

            //sql = " if not exists(select estiloactivo from estilos_home_user where Upper(iduser)=Upper('" + User + "') ) ";
            //sql += " insert into estilos_home_user (iduser, estiloactivo) values ('" + User + "'," + codestilo + "); ";
            //sql += " else ";
            //sql += " UPDATE estilos_home_user set estiloactivo = " + codestilo + " where iduser='" + User + "'; ";

            //Si no tiene datos lo inserto, sino lo actualizo
            if (TieneDatos("SELECT estiloactivo FROM estilos_home_user WHERE Upper(iduser)=Upper('" + User + "')",Base) == 0 )            
                sql = " INSERT into estilos_home_user (iduser, estiloactivo) values ('" + User + "'," + codestilo + ") ";
            else
                sql = " UPDATE estilos_home_user SET estiloactivo = " + codestilo + " where Upper(iduser)=Upper('" + User + "') ";

            try {

            cn.Open();
            OleDbCommand cmd = new OleDbCommand(sql, cn);
            cmd.ExecuteNonQuery();
                        
            if (cn.State == ConnectionState.Open)
                cn.Close();

            }
            catch (Exception ex)
            {
                //throw ex;

            }
             
        }


        [WebMethod(Description = "(HOME) Cambia el idioma de un usuario.")]
        public void Cambiar_Idioma(string User, string Base, string Idioma)
        {

            string TipoDB = DAL.TipoBase(Base);
            OleDbConnection cn = new OleDbConnection();
            cn.ConnectionString = DAL.constr(Base);
            String sql = "";

            sql = " UPDATE user_per SET lennro = (SELECT lennro FROM lenguaje WHERE lencod='" + Idioma + "' ) ";
            sql +=" WHERE Upper(iduser)='" + User.ToUpper() + "'  ";

            
            try
            {
                cn.Open();
                OleDbCommand cmd = new OleDbCommand(sql, cn);
                cmd.ExecuteNonQuery();

                if (cn.State == ConnectionState.Open)
                    cn.Close();
            }
            catch (Exception ex)
            {
             
                DAL.AddLogEvent("Error al cambiar a idioma "+Idioma+" del usuario '" + User + "'", EventLogEntryType.Information, 730);
                //throw ex;
            }

        }




        //-------------------------------------------------------------------------------------
        //-------------------------------------------------------------------------------------
        
        [WebMethod(Description = "(HOME) Devuelve los menu mas utilizados por el usuario.")]
        public DataTable MRU(string Usuario, int Cant, string Base, string Idioma)
        {
            //Defino un objeto tipo DAL
            //DAL MiDAL = new DAL();   
 
            string cn = DAL.constr(Base);
            string sql;
            string menuAccion = "";
            string cadena1 = "";
            string cadena2 = "";
            string[] arrPerfUser;
            string listaPerfUser = "";
            string[] arrAccess;
            String Collate = "";
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

            Columna = new DataColumn();
            Columna.DataType = System.Type.GetType("System.String");
            Columna.ColumnName = "menuimg";
            Columna.AutoIncrement = false;
            Columna.Unique = false;
            tablaSalida.Columns.Add(Columna);

            Columna = new DataColumn();
            Columna.DataType = System.Type.GetType("System.String");
            Columna.ColumnName = "menumsnro";
            Columna.AutoIncrement = false;
            Columna.Unique = false;
            tablaSalida.Columns.Add(Columna);

            Columna = new DataColumn();
            Columna.DataType = System.Type.GetType("System.String");
            Columna.ColumnName = "menuraiz";
            Columna.AutoIncrement = false;
            Columna.Unique = false;
            tablaSalida.Columns.Add(Columna);

            Columna = new DataColumn();
            Columna.DataType = System.Type.GetType("System.String");
            Columna.ColumnName = "menunro";
            Columna.AutoIncrement = false;
            Columna.Unique = false;
            tablaSalida.Columns.Add(Columna);

            Columna = new DataColumn();
            Columna.DataType = System.Type.GetType("System.String");
            Columna.ColumnName = "mrucant";
            Columna.AutoIncrement = false;
            Columna.Unique = false;
            tablaSalida.Columns.Add(Columna);

            if (DAL.TipoBase(Base).ToUpper() == "MSSQL")
                Collate = " COLLATE Modern_Spanish_CS_AS ";

            if (Usuario.Trim().Length == 0)
            {
                Usuario = "usr_logout";
            }

            determinarMI(Base);

            if (Idioma == "")
            {
                determinarIdioma(Usuario, Base);
            }
            else
            {
                idiomausuario = Idioma;
                idiomausuario2 = Idioma.Replace("-", "");
            }

            sql = "SELECT mru.mrucant, menumstr.menumsnro,menumstr.menuraiz,menuraiz.menunro,menumstr.menuname, menumstr.action, menuraiz.menunombre raiz, menuraiz.menudir, menumstr.menuaccess, menumstr.menuimg ";
            if (multiidioma == true)
            {
                if (idiomausuario2 == "")
                {
                    idiomausuario2 = "esAR";
                }
                sql = sql + ", t." + idiomausuario2 + " tituloMI";
                sql = sql + ", r." + idiomausuario2 + " raizMI";
               
            }
            sql = sql + " FROM mru";  
            sql = sql + " INNER JOIN menumstr ON menumstr.menumsnro = mru.menumsnro ";

            if (multiidioma == true)
            {
                sql = sql + " LEFT JOIN lenguaje_etiqueta t ON t.etiqueta = menumstr.menuname " + Collate + " and (t.pagina is null or t.pagina = '') and (t.modulo is null or t.modulo = '') ";
            }
           
            sql = sql + "INNER JOIN menuraiz ON menuraiz.menunro = mru.menuraiz ";

            if (multiidioma == true)
            {
                sql = sql + " LEFT JOIN lenguaje_etiqueta r ON r.etiqueta = menuraiz.menunombre " + Collate + " and (r.pagina is null or r.pagina = '') and (r.modulo is null or r.modulo = '') ";
            }
            sql = sql + "WHERE UPPER(mru.iduser) = '" + Usuario.ToUpper() + "' ";

            sql = sql + "ORDER BY  mru.mrucant DESC ";
            //sql = sql + "ORDER BY mrufecha DESC, mruhora DESC ";

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
                sql = sql + " UNION ALL ";
                sql = sql + " SELECT listperfnro from bk_perfil INNER JOIN bk_cab ON bk_cab.bkcabnro = bk_perfil.bkcabnro ";
                sql = sql + " AND (bk_cab.fdesde <= " + Fecha.cambiaFecha(DateTime.Today.ToString("dd/MM/yyyy"), DAL.TipoBase(Base.ToString())) + " AND (bk_cab.fhasta IS NULL OR bk_cab.fhasta >= " + Fecha.cambiaFecha(DateTime.Today.ToString("dd/MM/yyyy"), DAL.TipoBase(Base.ToString())) + " )) ";
              //  sql = sql + " AND (bk_cab.fdesde <= " + DateTime.Today.ToString(ConfigurationManager.AppSettings.Get("DateFormat").ToString()) + " AND (bk_cab.fhasta IS NULL OR bk_cab.fhasta >= " + DateTime.Today.ToString(ConfigurationManager.AppSettings.Get("DateFormat").ToString()) + " )) ";
               // sql = sql + " AND (bk_cab.fdesde <= " + Fecha.cambiaFecha(Convert.ToString(DateTime.Today), "SQL") + " AND (bk_cab.fhasta IS NULL OR bk_cab.fhasta >= " + Fecha.cambiaFecha(Convert.ToString(DateTime.Today), "SQL") + " )) ";
                sql = sql + " AND upper(bk_perfil.iduser) = '" + Usuario.ToUpper() + "'";

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
                    listaPerfUser = "0";
                    foreach (DataRow f in dsAux.Tables[0].Rows)
                    {
                        listaPerfUser = listaPerfUser + "," + Convert.ToString(f.ItemArray[0]);
                    }

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

                            if (menuAccion != "" && menuAccion != "#")
                            {

                                if (menuAccion.IndexOf("('../", StringComparison.CurrentCultureIgnoreCase) != -1)
                                {


                                    cadena1 = menuAccion.Substring(0, menuAccion.IndexOf("../", StringComparison.CurrentCultureIgnoreCase));
                                    cadena2 = menuAccion.Substring(menuAccion.IndexOf("../", StringComparison.CurrentCultureIgnoreCase) + 3, menuAccion.Length - menuAccion.IndexOf("../", StringComparison.CurrentCultureIgnoreCase) - 3);
                                }
                                else
                                {
                                    cadena1 = menuAccion.Substring(0, menuAccion.IndexOf("('", StringComparison.CurrentCultureIgnoreCase) + 2);
                                    cadena2 = menuAccion.Substring(menuAccion.IndexOf("('", StringComparison.CurrentCultureIgnoreCase) + 2, menuAccion.Length - menuAccion.IndexOf("('", StringComparison.CurrentCultureIgnoreCase) - 2);
                                    cadena2 = fila["menudir"].ToString() + "/" + cadena2;
                                }
                                //menuAccion.Replace("('../","('")

                            }
                            else
                            {
                                cadena1 = "";
                                cadena2 = "";
                            }

                            filaAux = tablaSalida.NewRow();


                            if (multiidioma == true)
                            {
//                                if (fila["tituloMI"].ToString().Length == 0)
                                if (fila["tituloMI"].ToString() == "")
                                {
                                    filaAux["menuname"] = fila["menuname"].ToString();
                                }
                                else
                                {
                                    filaAux["menuname"] = fila["tituloMI"].ToString();
                                }
                            //    if (fila["raizMI"].ToString().Length == 0)
                                if (fila["raizMI"].ToString() == "")
                                {
                                    filaAux["raiz"] = fila["raiz"].ToString();
                                }
                                else
                                {
                                    filaAux["raiz"] = fila["raizMI"].ToString();
                                }
                            }
                            else
                            {
                                filaAux["menuname"] = fila["menuname"].ToString();
                                filaAux["raiz"] = fila["raiz"].ToString();
                            }
                            filaAux["action"] = cadena1 + cadena2; ;

                            filaAux["menuimg"] = fila["menuimg"];

                            filaAux["menumsnro"] = fila["menumsnro"].ToString();
                            filaAux["menuraiz"] = fila["menuraiz"].ToString();
                            filaAux["menunro"] = fila["menunro"].ToString();
                            filaAux["mrucant"] = fila["mrucant"].ToString();

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
            //Defino un objeto tipo DAL
            //DAL MiDAL = new DAL();   
 
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
        {   //Defino un objeto tipo DAL
            //DAL MiDAL = new DAL();   

            return DAL.Bases();
        }

        //-------------------------------------------------------------------------------------
        //-------------------------------------------------------------------------------------
        
        [WebMethod(Description = "(HOME) Devuelve una tabla con Nombre de Modulo, Nombre de Menu y Accion de Menu segun la palabra buscada y usuario.")]
        public DataTable Search(string Usuario, string Palabra, string Base, string Idioma)
        {
            //Defino un objeto tipo DAL
            //DAL MiDAL = new DAL();   
 
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

            String Collate = "";


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

            if (DAL.TipoBase(Base).ToUpper() == "MSSQL")
                Collate = " COLLATE Modern_Spanish_CS_AS ";

            if (Usuario.Trim().Length != 0)
            {
                //Busco el perfil del usuario

                //Busco el perfil del usuario
                sql = "SELECT listperfnro ";
                sql = sql + "FROM user_perfil ";
                sql = sql + "WHERE UPPER(user_perfil.iduser) = '" + Usuario.ToUpper() + "' ";
                sql = sql + " UNION ALL ";
                sql = sql + " SELECT listperfnro from bk_perfil INNER JOIN bk_cab ON bk_cab.bkcabnro = bk_perfil.bkcabnro ";
                sql = sql + " AND (bk_cab.fdesde <= " + Fecha.cambiaFecha(DateTime.Today.ToString("dd/MM/yyyy"), DAL.TipoBase(Base.ToString())) + " AND (bk_cab.fhasta IS NULL OR bk_cab.fhasta >= " + Fecha.cambiaFecha(DateTime.Today.ToString("dd/MM/yyyy"), DAL.TipoBase(Base.ToString())) + " )) ";
                //sql = sql + " AND (bk_cab.fdesde <= " + DateTime.Today.ToString(ConfigurationManager.AppSettings.Get("DateFormat").ToString()) + " AND (bk_cab.fhasta IS NULL OR bk_cab.fhasta >= " + DateTime.Today.ToString(ConfigurationManager.AppSettings.Get("DateFormat").ToString()) + " )) ";
                //sql = sql + " AND (bk_cab.fdesde <= " + Fecha.cambiaFecha(Convert.ToString(DateTime.Today), "SQL") + " AND (bk_cab.fhasta IS NULL OR bk_cab.fhasta >= " + Fecha.cambiaFecha(Convert.ToString(DateTime.Today), "SQL") + " )) ";
                sql = sql + " AND upper(bk_perfil.iduser) = '" + Usuario.ToUpper() + "'";
                
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
                    listaPerfUser = "0";
                    foreach (DataRow f in dsPerfil.Tables[0].Rows)
                    {
                        listaPerfUser = listaPerfUser + "," + Convert.ToString(f.ItemArray[0]);
                    }

                    //Busco todos los menu que tienen la palabra buscada

                    //sql = "SELECT menuraiz.menunombre, menuname, action, menumstr.menuaccess, menuraiz.menudir, menumstr.menudesabr ";
                    sql = "SELECT menuraiz.menunombre,lenguaje_etiqueta." + Idioma.Replace("-", "") + " menuname, action, menumstr.menuaccess, menuraiz.menudir, menumstr.menudesabr ";
                    sql = sql + " FROM menumstr ";
                    sql = sql + " INNER JOIN menuraiz ON menuraiz.menunro = menumstr.menuraiz ";
                    sql = sql + " INNER JOIN lenguaje_etiqueta ON lenguaje_etiqueta.etiqueta = menumstr.menuname " + Collate + " ";
                    sql = sql + " WHERE lenguaje_etiqueta." + Idioma.Replace("-", "");
                    //WHERE menuname LIKE '%empleado%' 
                    


                    sql = sql + " LIKE '%" + Palabra + "%' ";
                    
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
                                        //filaAux["DescrMenu"] = fila[Idioma.Replace("-", "")].ToString();
                                        
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
        [WebMethod(Description = "(HOME) Retorna la cantidad de días que faltan para que a un usuario le caduque la contraseña.")]
        public String Info_Dias_VencimientoPass(string Usuario, string Base)
       {    
            String Salida = "";
            long polNro = 0;
            long passExpiraDias = 0;
            long diffDias = 0;
            long diasAlerta = 0;
            
            DateTime hPassFecini = DateTime.Today;
            //Busco Política asociada  al usuario
            polNro = (Password.valorUserPolCuenta(Usuario, "pol_nro", Base).Length != 0) ? Convert.ToInt64(Password.valorUserPolCuenta(Usuario, "pol_nro", Base)) : 0;
            //Busco la cantidad de días de expiracion de la politica
            passExpiraDias = (Password.valorPolCuenta(polNro, "pass_expira_dias", Base).Length != 0) ? Convert.ToInt64(Password.valorPolCuenta(polNro, "pass_expira_dias", Base)) : 0;                    

            //Si tiene algun valor configurado, calculo la fecha y cantidad de días que faltan para el vencimiento del password
            if (passExpiraDias != 0)
            {
                //Recupero la fecha de creación del password del usuario
                hPassFecini = (Password.valorHistPass(Usuario, "hpassfecini", Base).Length != 0) ? Convert.ToDateTime(Password.valorHistPass(Usuario, "hpassfecini", Base)) : DateTime.Today;
                //A la fecha de creación del password del usuario le sumo la cantidad de días de expiracion configurada en la politica
                DateTime FechaExpira = Convert.ToDateTime(hPassFecini); 
                FechaExpira = FechaExpira.AddDays(passExpiraDias);     

                //Recupero de la politica, la cantidad de días del alerta previo vencimiento
                diasAlerta = (Password.valorPolCuenta(polNro, "pass_alerta_dias", Base).Length != 0) ? Convert.ToInt64(Password.valorPolCuenta(polNro, "pass_alerta_dias", Base)) : 0;
  
                //Obtengo la diferencia de dias entre la fecha de expiracion y la fecha actual
                diffDias = Fecha.DateDiff(DateInterval.Day, DateTime.Today, FechaExpira);

                //Si la diferencia de dias entre la fecha de expiracion y la fecha actual es menor o igual a la cantidad de días del alerta previo vencimiento
                //informo la cantidad de dias de expiracion
                if (  diffDias <= diasAlerta )
                {
                    Salida = Convert.ToString(diffDias);
                }
            }
            
            return Salida;                          

    }
 


        //-------------------------------------------------------------------------------------
        //-------------------------------------------------------------------------------------
     
        [WebMethod(Description = "(HOME) Retorna si el usuario ingresa, un mensaje de error y si el mismo debe cambiar la password.")]
//        public DataTable Login(string Usuario, string Pass, string SegNt, string Base, string Idioma, String AUTH_USER, String REMOTE_ADDR)
          public DataTable Login(string Usuario, string Pass, string SegNt, string Base, string Idioma, String AUTH_USER, String REMOTE_ADDR,
                                 bool Validar_CamposLoguin, int DiasBloqueo)
            
        {
            //Defino un objeto tipo DAL
            //DAL MiDAL = new DAL();   
            DAL.AddLogEvent("Inicio login para el usuario '" + Usuario + "'", EventLogEntryType.Information, 700);

            bool Ingresa = true;
            string MessSUP = "";
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

            string cn = DAL.constr(Base);
            string sql;
            DataSet idiomaAux = new DataSet();
            string idio;
  
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

            // Verifico si hay que cambiar el password del supervisor
            try
            {
                DAL.CheckSupPass(Base);
            }
            catch
            {
                Ingresa = false;
                MessSUP = DAL.Error(0, Idioma, Base);
            }



            //--------------------------------------------------------------------------------
            //Verifico si la IP esta bloqueada
            //--------------------------------------------------------------------------------
            DAL.AddLogEvent("Validacion de IP Bloqueada", EventLogEntryType.Information, 704);
            int Bloq_IP = IP_Bloqueada(REMOTE_ADDR, AUTH_USER, Base, Validar_CamposLoguin, DiasBloqueo);

            if ((Ingresa) && (Bloq_IP>0))
            {
                if (Bloq_IP == 1) //Mess = "Cliente Bloqueado.";
                    Mess = DAL.Error(25, Idioma, Base);
                else //Mess = "Datos Invalidos.";
                    Mess = DAL.Error(26, Idioma, Base);

                Ingresa = false;
            }
                      
            
            //--------------------------------------------------------------------------------
            //Si se utiliza validación por servicio de directorio LDAP, verifico la existencia 
            //del usuario en el mismo.
            //
            //Tener en cuenta que No se debe tener habilitada la opción de
            //Seguridad Integrada (SegNt != TrueValue) para que esta modalidad funcione correctamente.
            //--------------------------------------------------------------------------------
            string LDAP_UseAuthentication = ConfigurationManager.AppSettings["LDAP_UseAuthentication"].ToString().ToLower().Trim();

            if (Ingresa)
            {
                

                if (SegNt != "TrueValue" && LDAP_UseAuthentication == "true") //Si se debe validar el usuario por LDAP...
                {
                    DAL.AddLogEvent("LDAP Activo", EventLogEntryType.Information, 701);
                    LDAP ldap = new LDAP();



                    if (!ldap.usuarioValido(Usuario, Pass)) //Si el usuario no es válido en LDAP...
                    {
                        Mess = DAL.Error(1, Idioma, Base);
                        Ingresa = false;
                    }
                    else //Si el usuario es válido...
                    {
                        //Actualizo la password del usuario (para que coincida con la del servidor LDAP).                    
                        //Tener en cuenta que no debe haber políticas de password definidas
                        //para que esta modalidad funcione correctamente. 

                        Mess = this.CambiarPass(Usuario, Pass, Pass, Pass, Base, Idioma);

                        if (Mess != "")
                            Ingresa = false;
                    }
                }
            }

            //--------------------------------------------------------------------------------
            //Pruebo la conexion con los datos del usuario
            //--------------------------------------------------------------------------------

            if (Ingresa)
            {
               DAL.AddLogEvent("Validaciones RH Pro", EventLogEntryType.Information, 702);
               OleDbConnection connUsu = new OleDbConnection(DAL.constrUsu(Usuario, Encriptar.Encrypt(DAL.EncrKy(), Pass), SegNt, Base));                
                
               try
               {
                    connUsu.Open();
               }
               catch (Exception ex)
               {
                    DAL.AddLogEvent("Incrementa cantidad de intentos fallidos por IP + NTUSER", EventLogEntryType.Information, 703);
                    //JPB: Incremento los intentos fallidos desde la IP + NT/User                    
                    Password.ActLogFallidos_NTUser_IP(Base, ValidaIPNueva(Base, REMOTE_ADDR, AUTH_USER), AUTH_USER, REMOTE_ADDR, false);
                   
                    DAL.AddLogEvent("Validaciones de políticas RH Pro", EventLogEntryType.Information, 703);

                    //Busco Política asociada
                    polNro = (Password.valorUserPolCuenta(Usuario, "pol_nro", Base).Length != 0) ? Convert.ToInt64(Password.valorUserPolCuenta(Usuario, "pol_nro", Base)) : 0;
                    
                    //Si no existe el historico de logueo lo creo
                    //Password.ValidaHis_log_usr(Usuario, Base);
                   
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
                        Mess = DAL.Error(5, Idioma, Base);
                    }
                    else
                    {                         
                        Password.actLogFallidos(Usuario, intentosFallidos, Base);
                        //Mess = "Contraseña incorrecta.";
                        Mess = DAL.Error(6, Idioma, Base);
                    }
                    //------------------
                    //Mess = "Usuario o constraseña incorrecta.";
                    //Mess = DAL.Error(1, Idioma);
                    Ingresa = false;
                }
                DAL.AddLogEvent("Fin validaciones de políticas RH Pro", EventLogEntryType.Information, 704);

                connUsu.Close();
            }

            //--------------------------------------------------------------------------------
            //Verifico si el usuario es valido
            //--------------------------------------------------------------------------------
            DAL.AddLogEvent("Validacion de usuario RH Pro", EventLogEntryType.Information, 704);

            if ((Ingresa) && (!Password.usuarioValido(Usuario, Base))){                
                //Mess = "Usuario no válido.";
                Mess = DAL.Error(2, Idioma, Base);
                Ingresa = false;
            }


            //--------------------------------------------------------------------------------
            //Control de cuenta bloqueada por usuario no por politica
            //--------------------------------------------------------------------------------
            DAL.AddLogEvent("Validacion de clave RH Pro", EventLogEntryType.Information, 705);

            //if ((Password.ctaBloqueada(Usuario, Base)))
            if (Ingresa && (Password.ctaBloqueada(Usuario, Base)))
            
            {
                //Mess = "Cuenta Bloqueada. Consulte con el administrador.";
                Mess = DAL.Error(3, Idioma, Base);
                Ingresa = false;
            }
            
            //Seguridad Base de Datos
            DAL.AddLogEvent("Validacion de políticas", EventLogEntryType.Information, 706);

            if (Ingresa && SegNt != "TrueValue" && LDAP_UseAuthentication == "false")
            //if (Ingresa && LDAP_UseAuthentication == "false")
            {

                //--------------------------------------------------------------------------------
                //Busca politica de cuenta
                //--------------------------------------------------------------------------------
                
                if (Ingresa)
                {
                    polNro = (Password.valorUserPolCuenta(Usuario, "pol_nro", Base).Length != 0) ? Convert.ToInt64(Password.valorUserPolCuenta(Usuario, "pol_nro", Base)) : 0;
                    
                    if (polNro == 0)
                    {
                        //Mess = "No se encontro la politica de cuenta del usuario.";
                        Mess = DAL.Error(4, Idioma, Base);
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
                        Mess = DAL.Error(5, Idioma, Base);
                    }
                    else
                    {
                        Password.actLogFallidos(Usuario, intentosFallidos, Base);
                        //Mess = "Contraseña incorrecta.";
                        Mess = DAL.Error(6, Idioma, Base);
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
                        Mess = DAL.Error(7, Idioma, Base);
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
                            Mess = DAL.Error(8, Idioma, Base);
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
                            Mess = DAL.Error(9, Idioma, Base);
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
                            Mess = DAL.Error(10, Idioma, Base);
                        }
                    }
                }
            }
            
            //--------------------------------------------------------------------------------
            //Registro el login del usuario
            //--------------------------------------------------------------------------------
            DAL.AddLogEvent("Registro login del usuario", EventLogEntryType.Information, 707);

            if (Ingresa)
            {
                //Limpio los intentos fallidos por IP
                Password.ActLogFallidos_NTUser_IP(Base, ValidaIPNueva(Base, REMOTE_ADDR, AUTH_USER), AUTH_USER, REMOTE_ADDR, true);
                Password.ingresarLogueo(Usuario,Base);
            }

            //Genero la salida

            DataRow fila = tablaSalida.NewRow();
            fila["Ingresa"] = Ingresa;           
            fila["mensaje"] = (MessSUP == "") ? Mess : MessSUP;
            fila["CambiarPass"] = cambiaPass;

 

            if (Ingresa)
            {      //Recupero el lenguaje configurado para dicho usuario
                    DAL.AddLogEvent("Recupero el lenguaje del usuario", EventLogEntryType.Information, 708);
                     
                    //sql = "SELECT lencod FROM user_per INNER JOIN lenguaje ON lenguaje.lennro = user_per.lennro ";
                    //sql = sql + " WHERE iduser = '" + Usuario.ToUpper() + "'";
     
                    //da = new OleDbDataAdapter(sql, cn);

                    //try
                    //{
                    //    //da.Fill(idiomaAux);
                    //    /*jpb*/
                    //    da.Fill(idiomaAux);
                    //}
                    //catch (Exception ex)
                    //{
                    //    throw ex;
                    //}

                    /*jpb: Si el usuario no tiene configurado idioma, toma el de la base*/
                    //if (idiomaAux.Tables[0].Rows.Count > 0)
                    //    idio = idiomaAux.Tables[0].Rows[0].ItemArray[0].ToString();
                    //else
                    //    idio = Idioma;

                   
                    idio = getIdiomaUsuario(Usuario, Base);

                    //idioma tomado de la base de datos.
                    fila["lenguaje"] = idio;
 
            }
            else
            {
                fila["lenguaje"] = Idioma;
            }


            fila["MaxEmpl"] = "100";

            tablaSalida.Rows.Add(fila);

            DAL.AddLogEvent("Fin Login", EventLogEntryType.Information, 709);
            return tablaSalida;
        }
 

        //-------------------------------------------------------------------------------------
        //------------OTRA-------------------------------------------------------------------------

        






        [WebMethod(Description = "(HOME) Cambia la contraseña del Usuario con contraseña anterior PassOld a la nueva PassNew. Devuelve string que si es vacio entonces cambio password ok, sino devuelve el error.")]
        public string CambiarPass(string Usuario, string PassOld, string PassNew, string PassConfirm, string Base, string Idioma)
        {
            //Defino un objeto tipo DAL
            //DAL MiDAL = new DAL();   
 
            bool Termino = false;
            string Mess = "";
            long passHistoria = 0;
            //--------------------------------------------------------------------------------
            //Control de usuario valido
            //--------------------------------------------------------------------------------

            if (!Termino && !Password.usuarioValido(Usuario, Base))
            {
                //Mess = "Usuario no válido.";
                Mess = DAL.Error(12, Idioma, Base);
                Termino = true;
            }

            //--------------------------------------------------------------------------------
            //Control de cuenta bloqueada
            //--------------------------------------------------------------------------------

            if (!Termino && (Password.ctaBloqueada(Usuario, Base)))
            {
                //Mess = "Cuenta Bloqueada. Consulte con el administrador.";
                Mess = DAL.Error(13, Idioma, Base);
                Termino = true;
            }

            /*jpb: Si ingresa por LDAP no se debe hacer el control de politicas*/
            string LDAP_UseAuthentication = ConfigurationManager.AppSettings["LDAP_UseAuthentication"].ToString().ToLower().Trim();

            /*if (LDAP_UseAuthentication != "true")
            {              
            */

            //--------------------------------------------------------------------------------
                //Control de coincidencia con confirmacion
                //--------------------------------------------------------------------------------

                if (!Termino && (PassConfirm != PassNew))
                {
                    //Mess = "La confirmación de la contraseña no es coincidente.";
                    Mess = DAL.Error(11, Idioma, Base);
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
                        Mess = DAL.Error(14, Idioma, Base);
                        Termino = true;
                    }
                }

                //--------------------------------------------------------------------------------
                //Control de password anterior
                //--------------------------------------------------------------------------------
                if (LDAP_UseAuthentication != "true")
                {
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
                            Mess = DAL.Error(15, Idioma, Base);
                        }
                        else
                        {
                            Password.actLogFallidos(Usuario, intentosFallidos, Base);
                            //Mess = "Contraseña incorrecta.";
                            Mess = DAL.Error(16, Idioma, Base);
                        }
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
                        Mess = DAL.Error(17, Idioma, Base);
                        Termino = true;
                    }
                    else
                    {
                        if ((passLongitud > 0) && (PassNew.Length < passLongitud))
                        {
                            //Mess = "La longitud mínima es de " + passLongitud + " caracteres.";
                            Mess = DAL.Error(18, Idioma, Base) + passLongitud + DAL.Error(19, Idioma, Base);
                            Termino = true;
                        }
                    }
                }

                //--------------------------------------------------------------------------------
                //LED - Control de politica configurada para el usuario - Caracteres Especiales
                //--------------------------------------------------------------------------------
                long caract_especiales = (Password.valorPolCuenta(polNro, "pass_especiales", Base).Length != 0) ? Convert.ToInt64(Password.valorPolCuenta(polNro, "pass_especiales", Base)) : 0;

                if (!Termino)
                {
                    if (caract_especiales == -1)
                    {

                        long error_caract_especiales = Password.contieneCaractEspeciales(PassNew);
                        //Mess = "No se encontro la politica de cuenta del usuario.";
                        if (error_caract_especiales == 0)
                        {
                            Mess = DAL.Error(21, Idioma, Base);
                            Termino = true;
                        }
                    }
                }


                //--------------------------------------------------------------------------------
                //LED - Control de politica configurada para el usuario - Exigir Letras Mayusculas
                //--------------------------------------------------------------------------------
                long mayusculas = (Password.valorPolCuenta(polNro, "pass_mayuscula", Base).Length != 0) ? Convert.ToInt64(Password.valorPolCuenta(polNro, "pass_mayuscula", Base)) : 0;

                if (!Termino)
                {
                    if (mayusculas == -1)
                    {

                        long error_mayusculas = Password.contieneMayusculas(PassNew);
                        //Mess = "No se encontro la politica de cuenta del usuario.";
                        if (error_mayusculas == 0)
                        {
                            Mess = DAL.Error(22, Idioma, Base);
                            Termino = true;
                        }
                    }
                }

                //--------------------------------------------------------------------------------
                //LED - Control de politica configurada para el usuario - Exigir Letras Minusculas
                //--------------------------------------------------------------------------------
                long Minusculas = (Password.valorPolCuenta(polNro, "pass_minuscula", Base).Length != 0) ? Convert.ToInt64(Password.valorPolCuenta(polNro, "pass_minuscula", Base)) : 0;

                if (!Termino)
                {
                    if (Minusculas == -1)
                    {

                        long error_Minusculas = Password.contieneMinusculas(PassNew);
                        //Mess = "No se encontro la politica de cuenta del usuario.";
                        if (error_Minusculas == 0)
                        {
                            Mess = DAL.Error(23, Idioma, Base);
                            Termino = true;
                        }
                    }
                }

                //--------------------------------------------------------------------------------
                //LED - Control de politica configurada para el usuario - Exigir Numeros
                //--------------------------------------------------------------------------------
                long Numeros = (Password.valorPolCuenta(polNro, "pass_numeros", Base).Length != 0) ? Convert.ToInt64(Password.valorPolCuenta(polNro, "pass_numeros", Base)) : 0;

                if (!Termino)
                {
                    if (Numeros == -1)
                    {

                        long error_Numeros = Password.contieneNumeros(PassNew);
                        //Mess = "No se encontro la politica de cuenta del usuario.";
                        if (error_Numeros == 0)
                        {
                            Mess = DAL.Error(24, Idioma, Base);
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
                        //DAL.Error(21, Idioma);
                        Mess = DAL.Error(20, Idioma, Base);
                        Termino = true;
                    }

                }

            /*}*/

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
                    Password.CambiarPassBase(Usuario, Encriptar.Encrypt(DAL.EncrKy(), PassNew), Encriptar.Encrypt(DAL.EncrKy(), PassOld), Base, true);
                 
            }

            if (Mess == null)
                Mess = "";

            return Mess;        
        }

        //-------------------------------------------------------------------------------------
        //-------------------------------------------------------------------------------------
        
        [WebMethod(Description = "(ORG) Dado un codigo de empleado y un codigo de base retorna el codigo de empleado padre.")]
        public long Padre(long CodEmp, int Base)
        {
            //Defino un objeto tipo DAL
            //DAL MiDAL = new DAL();   

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
            //Defino un objeto tipo DAL
            //DAL MiDAL = new DAL();   

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
            //Defino un objeto tipo DAL
            //DAL MiDAL = new DAL();   

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
            //Defino un objeto tipo DAL
            //DAL MiDAL = new DAL();   

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
            //Defino un objeto tipo DAL
            //DAL MiDAL = new DAL();   

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
            //Defino un objeto tipo DAL
            //DAL MiDAL = new DAL();   

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
            //Defino un objeto tipo DAL
            //DAL MiDAL = new DAL();   

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
            //Defino un objeto tipo DAL
            //DAL MiDAL = new DAL();   

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
            //Defino un objeto tipo DAL
            //DAL MiDAL = new DAL();   

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
            //Defino un objeto tipo DAL
            //DAL MiDAL = new DAL();   

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
            //Defino un objeto tipo DAL
            //DAL MiDAL = new DAL();   

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
            //Defino un objeto tipo DAL
            //DAL MiDAL = new DAL();   

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
            //Defino un objeto tipo DAL
            //DAL MiDAL = new DAL();   

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
            //Defino un objeto tipo DAL
            //DAL MiDAL = new DAL();   
 

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
                case "NOTAS":
                    sql = "  select tnonro cod, tnodesabr descr from tiponota ORDER BY tnodesabr ";
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


        [WebMethod(Description = "(Interfaz Inteligente) Devuelve codigo y descripcion de la tabla plana de Erecruiting seleccionada.")]
        public DataTable TablaPlanaErecruiting(string Tabla, string Base)
        {

            string cn = DAL.constr(Base);
            string sql = "";
            DataSet ds = new DataSet();


            switch (Tabla.ToUpper())
            {

                case "CARRERA":
                    sql = "SELECT carredunro cod, carredudesabr descr FROM cap_carr_edu";
                    break;
                case "ESPECIALIZACION":
                    sql = "select espnro cod, espdesabr descr  from especializacion";
                    break;
                case "ELEMENTOESPEC":
                    sql = "select eltananro cod, eltanadesabr descr from eltoana";
                    break;
                case "ESTADOCIVIL":
                    sql = "select estcivnro cod, estcivdesabr descr from estcivil";
                    break;
                case "IDIOMA":
                    sql = "select idinro cod, ididesc descr  from idioma";
                    break;
                case "INSTITUCION":
                    sql = "select instnro cod, instdes descr from institucion";
                    break;
                case "LOCALIDAD":
                    sql = "select  locnro cod, locdesc descr from localidad";
                    break;
                case "NACIONALIDAD":
                    sql = "select nacionalnro cod, nacionaldes descr from nacionalidad";
                    break;
                case "NIVELESPC":                    
                    sql = "select espnivnro cod, espnivdesabr descr from espnivel";
                    break;
                case "IDIOMANIVEL":
                    sql = "select empidlee cod, idileedes descr  from idinivellee";
                    break;
                case "PAIS":
                    sql = "select paisnro cod, paisdesc descr from pais";
                    break;
                case "PROVINCIA":
                    sql = "select provnro cod, provdesc descr from provincia";
                    break;
                case "CARGO":
                    sql = "select estrnro cod, puestonro descr from recPUESTO";
                    break;
                case "SEXO":
                    sql = " select recTERSEX.tersex cod, recTERSEX.sexdes descr from recTERSEX ";
                    break;
                case "TIPODOC":
                    sql = "select tidnro cod, tidnom descr from tipodocu";
                    break;
                case "NIVELESTUDIO":
                    sql = "select nivnro cod, nivdesc descr from nivest";
                    break;
                case "TITULO":
                    sql = " select titulo.titnro cod,titulo.titdesabr descr from  titulo ";
                    break;
                case "NOTAS":
                    sql = "  select tnonro cod, tnodesabr descr from tiponota ORDER BY tnodesabr ";
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

        //**********************************************************
        //**********************************************************

        [WebMethod(Description = "(HOME) Retorna los lenguajes habilitados. Activos (-1), inactivos (0).")]
        public DataTable ComboBanderasMI(int Base)
        {
            //Defino un objeto tipo DAL
            //DAL MiDAL = new DAL();   

            string cn = DAL.constr(Base.ToString());
            string sql;
            DataSet ds = new DataSet();

            DataTable tablaSalida = new DataTable("table");

            DataColumn Columna1 = new DataColumn();
            Columna1.DataType = System.Type.GetType("System.String");
            Columna1.ColumnName = "lencod";
            Columna1.AutoIncrement = false;
            Columna1.Unique = false;
            tablaSalida.Columns.Add(Columna1);

            DataColumn Columna2 = new DataColumn();
            Columna2.DataType = System.Type.GetType("System.String");
            Columna2.ColumnName = "lendesabr";
            Columna2.AutoIncrement = false;
            Columna2.Unique = false;
            tablaSalida.Columns.Add(Columna2);

            
            sql = "SELECT lencod,lendesabr,paisdesc FROM lenguaje ";
            sql += " INNER JOIN pais ON pais.paisnro = lenguaje.paisnro";
            sql += " WHERE lenactivo <> 0";
            sql += " ORDER BY paisdef,paisdesc ASC";

            da = new OleDbDataAdapter(sql, cn);

            try
            {
                da.Fill(ds);

                //Cargo la tabla de salida
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    DataRow r = tablaSalida.NewRow();
                    r["lencod"] = ds.Tables[0].Rows[i]["lencod"].ToString().Substring(0, 2) + '-' + ds.Tables[0].Rows[i]["lencod"].ToString().Substring(2, 2);
                    r["lendesabr"] = ds.Tables[0].Rows[i]["lendesabr"].ToString().Substring(0, ds.Tables[0].Rows[i]["lendesabr"].ToString().IndexOf("-") - 1) + '-' + ds.Tables[0].Rows[i]["paisdesc"];
                    tablaSalida.Rows.Add(r);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return tablaSalida;
        }


        

        //**********************************************************
        [WebMethod(Description = "Retorna el string de conexion dada una base determinada")]
        public string constr(string NroBase)
        {
            //Defino un objeto tipo DAL
            //DAL MiDAL = new DAL();   

            return DAL.constr(NroBase);
        }
        
        //**********************************************************
        [WebMethod(Description = "Retorna un DataTable. Dado un string con la consulta y el numero de base")]
        public DataTable get_DataTable(string sql, string NroBase)
        {
            //Defino un objeto tipo DAL
            //DAL MiDAL = new DAL();   

            string cn = DAL.constr(NroBase);
            
            DataSet ds = new DataSet();            
            da = new OleDbDataAdapter(sql, cn);

            try
            {
                da.Fill(ds);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return ds.Tables.Count>0?ds.Tables[0]:null;
        }

        //**********************************************************
        [WebMethod(Description = "Retorna un DataSet. Dado un string con la consulta y el numero de base")]
        public DataSet get_DataSet(string sql, string NroBase)
        {   

            //Defino un objeto tipo DAL
            //DAL MiDAL = new DAL();   

            string cn = DAL.constr(NroBase);

            DataSet ds = new DataSet();

            da = new OleDbDataAdapter(sql, cn);

            try
            {
                da.Fill(ds);
            }
            catch (Exception ex)
            {
                throw ex;
            }            
            return ds;
            
        }


        
        //**********************************************************
        [WebMethod(Description = "Traduce una etiqueta")]
        //public DataTable get_Traduccion_Modulo(string EtiqLenguaje, string Etiqueta, string NroBase)
        public String get_Traduccion_Modulo(string EtiqLenguaje, string Etiqueta, string NroBase)
        { 

           return EtiquetasMI.EtiquetaErr(Etiqueta, EtiqLenguaje, NroBase);

            /*
            //Defino un objeto tipo DAL
            //DAL MiDAL = new DAL();              
            string cn = DAL.constr(NroBase);
            string sql;
                          

            if (DAL.TipoBase(NroBase).ToUpper() == "MSSQL")
            {
                sql = "SELECT  [" + EtiqLenguaje + "]  FROM menumstr ";
                sql += " WHERE menuname = '" + Etiqueta + "' AND [" + EtiqLenguaje + "]  IS NOT NULL ";

            }
            else
            {
                sql = "SELECT  '" + EtiqLenguaje.ToUpper() + "' FROM menumstr ";
                sql += " WHERE upper(menuname) = '" + Etiqueta.ToUpper() + "' AND '" + EtiqLenguaje.ToUpper() + "'  IS NOT NULL  ";
            }

           
            DataSet ds = new  DataSet();
            da = new OleDbDataAdapter(sql, cn);

            try
            {
                da.Fill(ds);
                return ds.Tables[0]; 

            }
            catch (Exception ex)
            {
               // throw ex;                
                return null;

            }
            */

           
        }


        //-------------------------------------------------------------------------------------
        //-------------------------------------------------------------------------------------

        [WebMethod(Description = "(HOME) Activa/Desactiva un determinado Gadgets. Valor: -1/0")]
        public Boolean Act_Desact_Gadget(int gadnro, int valor, string Base)
        {

            string sql = "";

            OleDbConnection cn = new OleDbConnection();
            cn.ConnectionString = DAL.constr(Base.ToString());

            try
            {              

                cn.Open();
                sql = "UPDATE Gadgets_User SET gadusractivo = " + valor + " WHERE gadusrnro=" + gadnro;                 
                OleDbCommand cmd = new OleDbCommand(sql, cn);                
                cmd.ExecuteNonQuery();                
                 

            }
            catch (Exception ex)
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

        [WebMethod(Description = "(HOME) Activa/Desactiva un determinado Gadgets. Valor: -1/0")]
        public Boolean Update_Pos_Gadget(int gadnro, int posicion, string Base)
        {
            string sql = "";            
            System.Data.OleDb.OleDbConnection cn = new System.Data.OleDb.OleDbConnection();
            cn.ConnectionString = DAL.constr(Base);                   

            try
            {
                cn.Open();

                sql = "UPDATE Gadgets_User SET gadusrposicion = " + posicion + " WHERE gadusrnro = " + gadnro;
                System.Data.OleDb.OleDbCommand cmd2 = new System.Data.OleDb.OleDbCommand(sql, cn);
                cmd2.ExecuteNonQuery();
            }
            catch
            {
                return false;
            }
            finally
            {
                if (cn.State == System.Data.ConnectionState.Open) cn.Close();
            }

            return true;
        }


        //-------------------------------------------------------------------------------------
        //-------------------------------------------------------------------------------------

        [WebMethod(Description = "(HOME) Ajusta el alto(Tipo 1) o ancho(Tipo 2) del gadget")]
        public Boolean Update_Dimension_Gadget(int gadnro, int valor, int Tipo, String Base, int gadtipo,int menumsnro)
        {
            //Tipo:1 es altura, 2 es ancho
            string sql = "";
            System.Data.OleDb.OleDbConnection cn = new System.Data.OleDb.OleDbConnection();
            cn.ConnectionString = DAL.constr(Base);

            try
            {
                cn.Open();
                if (gadtipo == 0)//Es un gadget de la portada principal
                {
                    if (Tipo == 1)
                        sql = " UPDATE Gadgets_User SET gadusraltofull = " + valor + " WHERE gadusrnro = " + gadnro;
                    else
                        sql = " UPDATE Gadgets_User SET gadusranchofull = " + valor + " WHERE gadusrnro = " + gadnro;
                }
                else//Si es un gadget de modulo
                {
                    if (Tipo == 1)
                        sql = " UPDATE Gadgets_User_Modulo SET gadusraltofull = " + valor + " WHERE gadusrnro = " + gadnro + " AND menumsnro =" + menumsnro;
                    else
                        sql = " UPDATE Gadgets_User_Modulo SET gadusranchofull = " + valor + " WHERE gadusrnro = " + gadnro + " AND menumsnro =" + menumsnro;
                }

             
                System.Data.OleDb.OleDbCommand cmd2 = new System.Data.OleDb.OleDbCommand(sql, cn);
                cmd2.ExecuteNonQuery();

 

            }
            catch 
            {   
                return false;
            }
            finally
            {
                if (cn.State == System.Data.ConnectionState.Open) cn.Close();
            }

            return true;
        }


        //-------------------------------------------------------------------------------------
        //-------------------------------------------------------------------------------------
        [WebMethod(Description = "(HOME) Retorna el nombre de la base de datos contenida en initial catalog en el string de conexion")]
        public string Initial_Catalog(string NroBase)
        {
            string InitialCatalog = "";

            try
            {   //Divido el ConnString para recuperar el Initial Catalog
                string[] ConnStr = Regex.Split(ConfigurationManager.ConnectionStrings[NroBase].ConnectionString.Trim().ToUpper(),"Initial Catalog=".Trim().ToUpper());
                string[] Catalog = Regex.Split(ConnStr[1], ";");
                InitialCatalog = Catalog[0];

            }
            catch {   }

            return InitialCatalog;
        }


        //-------------------------------------------------------------------------------------
        //-------------------------------------------------------------------------------------
        [WebMethod(Description = "(HOME) Retorna el nombre de la base de datos contenida en initial catalog en el string de conexion")]
        public string Data_Source(string NroBase)
        {
            string DS = "";

            try
            {   //Divido el ConnString para recuperar el Data Source
                string[] ConnStr = Regex.Split(ConfigurationManager.ConnectionStrings[NroBase].ConnectionString.Trim().ToUpper(), "Data Source=".Trim().ToUpper());
                string[] DataS = Regex.Split(ConnStr[1], ";");
                DS = DataS[0];
            }
            catch { }

            return DS;
        }

        //-------------------------------------------------------------------------------------
        //-------------------------------------------------------------------------------------
        [WebMethod(Description = "(HOME) Retorna el tipo de base de datos")]
        public string get_TipoBase(string NroBase)
        {
            return DAL.TipoBase(NroBase);
        }

        [WebMethod(Description = "(HOME) Asocia todos los gadget faltantes del primer acceso ")]
        public void Controlar_Gadget_EnLoguin(string User, string pass, string segNT,string Base)
        {
            String l_gadnro, l_gadtitulo, l_gadusranchofull, l_gadusraltofull, l_gadtipo, l_gaddefault;
            string TipoDB = DAL.TipoBase(Base).Trim().ToUpper();
             

            String sql = "";
            Boolean AlMenos_Uno = false;

            //Obtengo los tipos de gadget que le faltan al usuario
       
            sql = "  SELECT * FROM Gadgets_Tipo GT WHERE  ";
            sql += "   GT.gadnro NOT in (select GU.gadnro from Gadgets_User GU where GU.iduser='" + User + "') ";
            sql += "   AND GT.gadactivo = -1 ";
            sql += "   ORDER BY gadnro ";

            DataTable dt = get_DataTable(sql, Base);

            OleDbConnection cn3 = new OleDbConnection();
            cn3.ConnectionString = DAL.constr(Base);            
            cn3.Open();
            OleDbCommand cmd2;
            sql = "";
            foreach (DataRow GadgetFaltantes in dt.Rows)
            {
                try
                {
                    l_gadnro = Convert.ToString(GadgetFaltantes["gadnro"]);
                    l_gadtitulo = Convert.ToString(GadgetFaltantes["gadtitulo"]);
                    l_gadtipo = Convert.ToString(GadgetFaltantes["gadtipo"]);
                    l_gaddefault = Convert.ToString(GadgetFaltantes["gaddefault"]);
                    
                    if (l_gadnro == "1")
                    {
                        l_gadusranchofull = "-1";
                        l_gadusraltofull = "-1";
                    }
                    else
                    {
                        l_gadusranchofull = "0";
                        l_gadusraltofull = "0";
                    }

                    if (!AlMenos_Uno)
                    {
                        DataTable dt_existe = get_DataTable("select gadnro from Gadgets_User where Upper(iduser)=Upper('" + User + "')", Base);
                        AlMenos_Uno = (dt_existe.Rows.Count>0);
                    }
                    //Controlo que si existe al menos un registro para la relacion gadget - user para obtener la maxima posicion.
                   
                    if (AlMenos_Uno)                    
                    {                        
                        if (TipoDB == "MSSQL")
                        {
                            sql = " Declare @posicion int; ";
                            //JPB: La maxima posicion se calcula antes de hacer el insert
                            sql += " select @posicion = max(gadusrposicion) from Gadgets_User where Upper(iduser)=Upper('" + User + "');  ";
                            sql += " INSERT INTO Gadgets_User ( gadusractivo, gadusranchofull, gadusraltofull, gadusrposicion, iduser, gadusrmenudir, gadnro, gadusrdesabr,gadestado)";
                            sql += "       values (" + l_gaddefault + ", " + l_gadusranchofull + ", " + l_gadusraltofull;
                            sql += " ,@posicion + 1, '" + User + "',''," + l_gadnro + ",'" + l_gadtitulo + "',-1);";
                        }
                        else
                        {                            
                            sql = " INSERT INTO Gadgets_User ( gadusractivo, gadusranchofull, gadusraltofull, gadusrposicion, iduser, gadusrmenudir, gadnro, gadusrdesabr,gadestado)";
                            sql += "       values (" + l_gaddefault + ", " + l_gadusranchofull + ", " + l_gadusraltofull;
                            sql += " ,(select max(gadusrposicion) from Gadgets_User where Upper(iduser)=Upper('" + User + "')) + 1 ";
                            sql += ", '" + User + "',''," + l_gadnro + ",'" + l_gadtitulo + "',-1) ";
                            sql += " RETURNING gadusrnro INTO :currid ";
                        }                                                
                    }
                    else
                    {                         
                 
                        if (TipoDB == "MSSQL")
                        {                            
                            sql += " INSERT INTO Gadgets_User ( gadusractivo, gadusranchofull, gadusraltofull, gadusrposicion, iduser, gadusrmenudir, gadnro, gadusrdesabr,gadestado)";
                            sql += "       values (" + l_gaddefault + ", " + l_gadusranchofull + ", " + l_gadusraltofull;
                            sql += " ,1, '" + User + "',''," + l_gadnro + ",'" + l_gadtitulo + "',-1);";
                        }
                        else
                        {
                            sql = "     INSERT INTO Gadgets_User ( gadusractivo, gadusranchofull, gadusraltofull, gadusrposicion, iduser, gadusrmenudir, gadnro, gadusrdesabr,gadestado)";
                            sql += "       values (" + l_gaddefault + ", " + l_gadusranchofull + ", " + l_gadusraltofull + ",1, '" + User + "',''," + l_gadnro + ",'" + l_gadtitulo + "',-1) ";
                            sql += " RETURNING gadusrnro INTO :currid ";
                        }  
                    }

                    if (TipoDB == "MSSQL")
                        sql += " SELECT CAST(scope_identity() AS int) ";
                    else
                    {
                       // sql += " RETURNING gadusrnro INTO :currid ";
                       // sql += " select SEQ_Gadgets_User.CURRVAL  FROM DUAL ";                       
                    }
                   
                    /////*Ejecuto el query y recupero el ultimo id insertado*/

                    Int32 l_gadusrnro=0;

                    if (TipoDB == "MSSQL")
                    {
                        cmd2 = new OleDbCommand(sql, cn3);
                        l_gadusrnro = Convert.ToInt32(cmd2.ExecuteScalar());
                    }
                    else
                    {                        

                        OleDbParameter par = new OleDbParameter();
                            par.ParameterName = "currid";
                            par.DbType = DbType.Int32;
                            par.Direction = ParameterDirection.Output;

                        cmd2 = new OleDbCommand(sql, cn3);                      
                        cmd2.Parameters.Add(par);
                        cmd2.ExecuteNonQuery();

                        l_gadusrnro = Convert.ToInt32(cmd2.Parameters["currid"].Value);
                                              
                         //l_gadusrnro = Convert.ToInt32(cmdOra.Parameters[":currid"].Value);
 
                         //DataTable dt_o = get_DataTable("select SEQ_Gadgets_User.currval currid FROM DUAL", Base);
                         //if (dt_o.Rows.Count > 0)
                         //    l_gadusrnro = Convert.ToInt32(dt_o.Rows[0]["currid"]);
                         
                         
                    }
                     
                    if (l_gadtipo == "1")
                    {                           
                        if (l_gadusrnro > 0)
                        {
                            Asociar_Gadget_Modulos(User, Base, l_gadusrnro);
                        }
                    }
                    //18-02-2015 - JPB - Se agrega informe en log de windows
                    DAL.AddLogEvent("Se asociaron los gadgets al usuario " + User, EventLogEntryType.Information, 710);

                }
                catch (Exception ex)
                {
 
                    //18-02-2015 - JPB - Se agrega informe en log de windows                     
                    DAL.AddLogEvent("Error en Controlar_Gadget_EnLoguin:" + ex.Message, EventLogEntryType.Information, 711);
                    // throw ex;

                }
            }
            cn3.Close();
        }
 
                
        private void Asociar_Gadget_Modulos(string User, string Base, Int32 gadusrnro)
        {
            //System.Data.SqlClient.SqlTransaction transaction2 = null;
            try
            {
                               
                String sql = "";
                String sqlCommand = "";

                //Obtengo los modulos activos del sistema
                sql = " SELECT menumsnro  FROM menumstr WHERE Upper(parent) = Upper('rhpro') AND menuactivo=-1  ";
 
                DataTable dt = get_DataTable(sql, Base);
                
                System.Data.OleDb.OleDbConnection cn3 = new System.Data.OleDb.OleDbConnection();
                cn3.ConnectionString = DAL.constr(Base);
                cn3.Open();
                System.Data.OleDb.OleDbCommand cmd2;

                foreach (DataRow Modulo in dt.Rows)
                {
             
                    sqlCommand = " INSERT INTO Gadgets_User_Modulo ";
                    sqlCommand += " (gadusrnro,menumsnro,gadusranchofull,gadusraltofull) ";
                    sqlCommand +=  " VALUES ";
		            sqlCommand +=  " (" + gadusrnro + "," +Convert.ToString(Modulo["menumsnro"]) + ",0,0) ";                     

                    cmd2 = new System.Data.OleDb.OleDbCommand(sqlCommand, cn3);
                    cmd2.ExecuteNonQuery();
                }              
                
                cn3.Close();            
            }
            catch (Exception ex)
            {

                //18-02-2015 - JPB - Se agrega informe en log de windows
                DAL.AddLogEvent("Error en Asociar_Gadget_Modulos:" + ex.Message, EventLogEntryType.Information, 712);
             //   throw ex;

            }

        }






        /******************************************************************************************/
        /* CONTROL DE ACCESO  TEMPORAL Y POR ARMADO DE MENU   *********************************** */
        /******************************************************************************************/
        [WebMethod(Description = "(HOME) Retorna la lista de perfiles del usuario")]
        public List<String> getPerfilesUsuario(String idUser, String Base)
        //public List<String> getPerfilesUsuario(String idUser, String Base)
        {
            List<String> Lista = new List<string>();
            string[] Misplit1;
            string[] Misplit2;
          

            String fechaActual = Fecha.cambiaFecha(DateTime.Today.ToString("dd/MM/yyyy"), get_TipoBase(Base));
            //Busco el perfil del usuario
            string sql = "SELECT listperfnro ";
            sql = sql + " FROM user_perfil ";
            sql = sql + " WHERE Upper(user_perfil.iduser) = Upper('" + idUser + "') ";
            sql = sql + " UNION ALL ";
            sql = sql + " SELECT listperfnro from bk_perfil INNER JOIN bk_cab ON bk_cab.bkcabnro = bk_perfil.bkcabnro ";
            sql = sql + " AND (bk_cab.fdesde <= " + fechaActual + " AND (bk_cab.fhasta IS NULL OR bk_cab.fhasta >= " + fechaActual + " )) ";
            sql = sql + " AND upper(bk_perfil.iduser) = Upper('" + idUser + "')";

            DataTable dt =  get_DataTable(sql, Base);
            if (dt.Rows.Count > 0)
            {
                Misplit1 = (Convert.ToString(dt.Rows[0]["listperfnro"])).Split(','); 
                foreach (string Perfil in Misplit1)
                {
                    Lista.Add(Perfil);
                }
                if (dt.Rows.Count > 1)
                {
                    Misplit2 = (Convert.ToString(dt.Rows[1]["listperfnro"])).Split(','); 
                    foreach (string Perfil2 in Misplit2)
                    {
                        Lista.Add(Perfil2);
                    }
                }
            }


            return Lista;
        }


        public static bool Perfil_Habilitado(List<String> ListPerfUsr, String ListAccesos)
        {
            if (ListAccesos=="*") return true;

            String[] Misplit = ListAccesos.Split(',');
            foreach (String PerfUsr in ListPerfUsr)
            {
                foreach (String Acceso in Misplit)
                {
                    if ((PerfUsr == Acceso) || (Acceso == "*"))
                    {

                        return true;
                    }
                }
            }
            return false;
        }




        // public bool Menu_Habilitado(String menuaccess, int menumsnro, List<String>  ListaPerfUsr)
        [WebMethod(Description = "(HOME) Verifica si un menu se encuentra habilitado")]
        public bool Menu_Habilitado(String menuaccess, int menumsnro, String User, string Base )
        //public bool Menu_Habilitado(String menuaccess, int menumsnro, String User, string Base,String[] ListaPerfUsr)
        {
             
            DataTable grupo = Grupo_Restricciones(Base);
            List<String> ListaPerfUsr = getPerfilesUsuario(User, Base);

            //Si no encuentra un grupo de restriccion verifica si esta habilitado por perfil
            if ((grupo == null) || (DBNull.Value.Equals(grupo)))            
                return Perfil_Habilitado(ListaPerfUsr, menuaccess);            
            
            if (grupo.Rows.Count == 0) return Perfil_Habilitado(ListaPerfUsr, menuaccess);
            
            //Esta habilitado si esta habilitado por armado de menu, o por grupo de acceso            
            int Salida = Habilitado_Por_GrupoRestricciones(menumsnro, ListaPerfUsr, grupo);
             
            if (Salida < 1)
            {
                if (Salida == -1)
                    return true;
                else
                    return false;
            }
            else
            {
                return Perfil_Habilitado(ListaPerfUsr, menuaccess);
            }
            

        }


        [WebMethod(Description = "(HOME) Retorna el grupo de restricciones")]
        public DataTable Grupo_Restricciones(String Base)
        {

            String fechaActual = Fecha.cambiaFecha(DateTime.Today.ToString("dd/MM/yyyy"), get_TipoBase(Base));

            string sql = " ";
            sql = " SELECT A.alesch_frecrep,A.alesch_fecini,A.alesch_fecfin, A.frectipnro,A.schedhora,A.schedhora2, A.scheddesc ";
            sql += "     ,P.listperfnro,M.menumstrnros ,  MG.* ";
            sql += " FROM menugrp MG ";
            sql += "   inner join ale_sched A ON A.schednro = MG.schednro ";
            sql += "   inner join  menugrp_perf P ON P.menugnro = MG.menugnro ";
            sql += "   inner join  menugrp_menu M ON M.menugnro = MG.menugnro ";
            sql += " WHERE  ( (" + fechaActual + " >= A.alesch_fecini) AND (" + fechaActual + "<= A.alesch_fecfin)) ";


            return get_DataTable(sql, Base);
        }

        /// <summary>
        /// Verifica si tiene un grupo de acceso. 
        /// </summary>
        /// <param name="menumsnro"></param>
        /// <param name="ListaPerfUsr"></param>
        /// <param name="GrupoRestriccion"></param>
        /// <returns>-1: Puede Acceder, 0:No Puede Acceder, 1:No hay nada especificado</returns>
        public int Habilitado_Por_GrupoRestricciones(int menumsnro, List<String> ListaPerfUsr, DataTable GrupoRestriccion)
        {
            //Por cada restriccion configurada se verifica si se debe tomar en cuenta para restringir el acceso a ciertos usuasrios
            foreach (DataRow row in GrupoRestriccion.Rows)
            {
                //Verifico si el menumsnro esta asociado a la lista de menu en la restriccion
                if (!DBNull.Value.Equals(row["menumstrnros"]))
                {
                    string[] arrMenumsnro = Convert.ToString(row["menumstrnros"]).Split(',');

                    bool Existe = false;
                    foreach (string valor in arrMenumsnro)
                    {
                        if (valor != "")
                        {
                            if (Convert.ToInt32(valor) == menumsnro)
                            {
                                Existe = true;
                                break;
                            }
                        }
                    }

                    if (!Existe)
                        return 1;


                }


                //Verifico si estoy en el rango adecuado para realizar el control         
                if (!GrupoRestringido(Convert.ToInt32(row["frectipnro"]), row))
                {
                    return 1;
                }


                //Verifico el tipo de restriccion
                if (!DBNull.Value.Equals(row["frectipnro"]))
                {
                    //Controlo si los perfiles del usuario estan dentro del grupo y si ademas esta restringido segun el tipo de restriccion                                    
                    if (Perfil_Habilitado(ListaPerfUsr, Convert.ToString(row["listperfnro"])))
                        return -1;
                    else
                        return 0;
                }
            }

            //Si no pudo controlar, se asume que no tiene ninguna restriccion para ver el menu
            return 1;
        }

        protected bool Arr_Contain(String[] Arr, string elem)
        { 
            foreach(string s in Arr)
            {
            if (s==elem)
                return true;
            }
            return false;
        }


        protected bool GrupoRestringido(int tipo, DataRow row)
        {
            bool Habilitado = false;
            bool HoraHabilitada = ((DateTime.Now >= Convert.ToDateTime(row["schedhora"])) && (DateTime.Now <= Convert.ToDateTime(row["schedhora2"])));

            switch (tipo)
            {
                case 1://Diariamente                     
                    Habilitado = HoraHabilitada;
                    break;
                case 2://Semanalmente
                    //1 = Domingo // 2 = Lunes // 3 = Martes // 4 = Miercoles // 5 = Jueves // 6 = Viernes // 7 = Sabado  
                    int numero_dia = Convert.ToInt32(DateTime.Today.DayOfWeek) + 1;
                    Habilitado = (Convert.ToInt32(row["alesch_frecrep"]) == numero_dia) && (HoraHabilitada);
                    break;
                case 3://Mensualmente                  
                    if (!DBNull.Value.Equals(row["diassel"]))
                    {
                        String[] arrDias = Convert.ToString(row["diassel"]).Split(',');
                        Habilitado = (Arr_Contain(arrDias,Convert.ToString(DateTime.Today.Day))) && (HoraHabilitada);
                        //Habilitado = (arrDias.Contains(Convert.ToString(DateTime.Today.Day))) && (HoraHabilitada);
                    }
                    else //Si no hay dias configurados se adopta el mismo control que Diariamente
                        Habilitado = HoraHabilitada;

                    break;
                case 4://Temporalmente
                    Habilitado = Control_TemporalDias(Convert.ToDateTime(row["alesch_fecini"]), Convert.ToInt32(row["alesch_frecrep"])) && (HoraHabilitada);
                    break;
                default:

                    break;
            }

            return Habilitado;

        }

        /// <summary>
        /// Retorna verdadero si la fecha actual cae dentro de la configuracion temporal
        /// </summary>
        /// <param name="FechaInicio"></param>
        /// <param name="Incremento_Cada"></param>
        /// <returns></returns>
        protected bool Control_TemporalDias(DateTime FechaInicio, int Incremento_Cada)
        {

            bool Salida = true;
            String fecha = String.Format("{0:dd/MM/yyyy}", DateTime.Now.Date);
            DateTime fechaActual = Convert.ToDateTime(fecha);            
            DateTime fechaControl = Convert.ToDateTime(String.Format("{0:dd/MM/yyyy}", FechaInicio));//FechaInicio;            
            int comparacion;
            bool SeguirControl = true;

            while (SeguirControl)//Incremento la fecha de inicio hasta que sea mayor o igual que la fecha actual
            {
                fechaControl = fechaControl.AddDays(Incremento_Cada);
                //comparacion = DateTime.Compare(fechaControl,fechaActual);
                comparacion = DateTime.Compare(fechaActual, fechaControl);
                if (comparacion == 0)//Si las fechas son iguales
                {
                    SeguirControl = false;
                    Salida = true;
                }
                else
                    if (comparacion < 0)//Si ya se paso la suma quiere decir que el dia actual no cae dentro del espacio temporal
                    {
                        SeguirControl = false;
                        Salida = false;
                    }
            }

            return Salida;
        }




        /// <summary>
        /// Verifica si una IP esta bloqueada. Tambien actualiza los intentos fallidos en caso de que los campos no sean validos
        /// </summary>
        /// <param name="IP"></param>
        /// <param name="userName"></param>
        /// <param name="BaseSelec"></param>
        /// <param name="Validar_CamposLoguin"></param>
        /// <param name="DiasBloqueo"></param>
        /// <returns></returns>
        private  int IP_Bloqueada(string IP, string userName, string BaseSelec, bool Validar_CamposLoguin, int DiasBloqueo)
        {
            int Salida = 0;

            if (Control_Bloqueo_IP(userName, IP, BaseSelec, DiasBloqueo)) //Controla que la ip no este bloqueada
            {   
                //Ya esta bloqueado
                Salida = 1;
            }

            if ((Salida==0) && (!Validar_CamposLoguin))//Si no valida los campos de loguin no continuo el loguin
            {
                //Actualizo intentos fallidos por IP
                Actualizar_Bloqueos_IP(BaseSelec, userName, IP);
                Salida = 2;
            }

            return Salida;
        }
         
         
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
       // [WebMethod(Description = "(HOME) Verifica si la IP solicitante esta bloqueada")]
//        public bool Control_Bloqueo_IP(String userName, String IP, String BaseID, long DiasBloqueo)
        private bool Control_Bloqueo_IP(String userName, String IP, String BaseID, long DiasBloqueo)
        {
            String sql;
            int CantIntentos = 0;
            int MaxIntentos = 0;
            bool Salida = false;
            DataTable dt;
            Consultas cc = new Consultas();

            //Busco la cantidad maxima de intentos fallidos configurada en el confper
            sql = " SELECT confactivo, confint FROM confper WHERE confnro = 32 ";
            dt = cc.get_DataTable(sql, BaseID);
            if (dt.Rows.Count > 0)
            {
                if (Convert.ToInt32(dt.Rows[0]["confactivo"]) == -1)
                {
                    MaxIntentos = Convert.ToInt32(dt.Rows[0]["confint"]);

                    //Busco la cantidad de intentos fallidos de la IP de cliente
                    sql = " SELECT rhseglognro, rhseglogip, rhseglogpc, rhseglogHost, appnro, rhseglogfec, rhsegloghora, rhseglogcant ";
                    sql += " FROM rhpro_seg_login  WHERE rhseglogip='" + IP + "' AND rhseglogpc='" + userName + "'";
                    dt = cc.get_DataTable(sql, BaseID);

                    if (dt.Rows.Count > 0)
                    {
                        long diffDias = Fecha.DateDiff(DateInterval.Day, Convert.ToDateTime(dt.Rows[0]["rhseglogfec"]), DateTime.Today);                         

                        CantIntentos = Convert.ToInt32(dt.Rows[0]["rhseglogcant"]);
                        
                        Salida = ((CantIntentos > MaxIntentos) && (diffDias <= DiasBloqueo));
 
                        if (diffDias > DiasBloqueo)//Si ya pase los dias de bloqueo me vuelve a cero el contador
                        {                                                      
                            Password.ActLogFallidos_NTUser_IP(BaseID, ValidaIPNueva(BaseID, IP, userName), userName, IP, true);
                        }

                        return Salida;
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// Home: Actualiza bloqueos por IP
        /// </summary>
        /// <param name="Base"></param>
        /// <param name="EsNuevo"></param>
        /// <param name="AUTH_USER"></param>
        /// <param name="REMOTE_ADDR"></param>
  //      [WebMethod(Description = "(HOME) Verifica si la IP solicitante esta bloqueada")]
//        public void Actualizar_Bloqueos_IP(string Base,  String AUTH_USER, String REMOTE_ADDR)
        private void Actualizar_Bloqueos_IP(string Base, String AUTH_USER, String REMOTE_ADDR)
        {
            Password.ActLogFallidos_NTUser_IP(Base, ValidaIPNueva(Base,REMOTE_ADDR,AUTH_USER), AUTH_USER, REMOTE_ADDR, false);
        }



        private Boolean ValidaIPNueva(String Base, String REMOTE_ADDR, String AUTH_USER)
        {
            //JPB: Incremento los intentos fallidos desde la IP + NT/User
            String sql_consul;
            sql_consul = " SELECT rhseglogip FROM  rhpro_seg_login ";
            sql_consul += " WHERE rhseglogip = '" + REMOTE_ADDR + "' AND rhseglogpc ='" + AUTH_USER + "' ";

            return(TieneDatos(sql_consul, Base) == 0);
        }
         


 


    }
}