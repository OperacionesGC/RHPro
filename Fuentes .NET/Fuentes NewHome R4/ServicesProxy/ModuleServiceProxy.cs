using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using Entities;
using ServicesProxy.rhdesa;
using Common;

namespace ServicesProxy
{
    public static class ModuleServiceProxy
    {
        private const string DataBaseMenuNameFieldName = "menuname";

        private const string DataBaseMenuTitleFieldName = "menudesabr";
        private const string DataBaseMenuDetalleFieldName = "menudetalle";
        private const string DataBaseActionFieldName = "action";
        private const string DataBaseMenuObjetiveDetailFieldName = "menubeneficio";
        private const string DataBaseMenuObjetiveFieldName = "menuobjetivo";
        private const string DataBaseLinkManualFieldName = "linkmanual";
        private const string DataBaseLinkDvdFieldName = "linkdvd";

        private const string DataBaseMenuMsnroFieldName = "menumsnro";
        private const string DataBaseMenuRaizFieldName = "menuraiz";


        /// <summary>
        /// Busca lista de Module
        /// </summary>
        /// <param name="userName">Nombre de Usuario</param>
        /// <param name="dataBaseID">ID de la base</param>
        /// <returns>Lista de Module</returns>
        /// <param name="idioma">Idioma actual</param>
        public static List<Module> Find(string userName, string dataBaseID, string idioma)
        {
            List<Module> ModuleToReturn = new List<Module>();

            Consultas consultas = new Consultas();

            //consultas.Credentials = System.Net.CredentialCache.DefaultCredentials;
 
            DataTable modules = consultas.Modulos(userName, dataBaseID, idioma);

            /*JPB: Ordeno el DataTable que contiene los Modulos visibles
                   en el caso que un usuario logueado no tenga acceso a uno de los modulos, el action sera #. Por lo que al ordenar por action DESC 
                   aparecerán primero los modulos habilitados.
            */           
            DataView dv = modules.DefaultView;
            switch (Utils.Tipo_Ordenamiento_Modulos())
            {
                case "1": dv.Sort = "AccesosMRU DESC,menudesabr ASC"; break;
                case "2": dv.Sort = "menuponderacion DESC,menudesabr ASC"; break;
                case "3": dv.Sort = "menudesabr ASC"; break;
                default: dv.Sort = "menudesabr ASC"; break;                   
            }

            //JPB: Filtro el modulo HOME para que no se visualice
            dv.RowFilter = "  menuname <> 'HOME' ";

            modules = dv.ToTable();
            
            int posicion = 0;
            foreach (DataRow dataRow in modules.Rows)
            {
                Module footerPage = new Module
                {
                    Id = Guid.NewGuid(),
                    //Action = dataRow[DataBaseActionFieldName].ToString(),
                    Action = dataRow[DataBaseActionFieldName].ToString().Replace("abrirVentana('", "abrirVentana('../"),
                    LinkDvd = dataRow[DataBaseLinkDvdFieldName].ToString(),
                    LinkManual = dataRow[DataBaseLinkManualFieldName].ToString(),
                    MenuDetail = dataRow[DataBaseMenuDetalleFieldName].ToString(),
                    MenuObjective = dataRow[DataBaseMenuObjetiveFieldName].ToString(),
                    MenuObjectiveDetail = dataRow[DataBaseMenuObjetiveDetailFieldName].ToString(),
                    MenuTitle = dataRow[DataBaseMenuTitleFieldName].ToString(),
                    MenuName =  dataRow[DataBaseMenuNameFieldName].ToString(),
                    MenuMsnro = dataRow[DataBaseMenuMsnroFieldName].ToString(),
                    MenuRaiz = dataRow[DataBaseMenuRaizFieldName].ToString(),
                    pos = posicion,
                };
                posicion++;

                ModuleToReturn.Add(footerPage);
            }


            return ModuleToReturn;
        }


        /// <summary>
        /// Busca lista de Modulos. El ultimo parametro indica si es habilitado o inhabilitado
        /// </summary>
        /// <param name="userName">Nombre de Usuario</param>
        /// <param name="dataBaseID">ID de la base</param>
        /// <returns>Lista de Modulos habilitados</returns>
        /// <param name="idioma">Idioma actual</param>
        public static List<Module> Find_Modulos(string userName, string dataBaseID, string idioma, bool Habilitado)
        {
            List<Module> ModuleToReturn = new List<Module>();
            bool condicion;
            Consultas consultas = new Consultas();

            //consultas.Credentials = System.Net.CredentialCache.DefaultCredentials;

            DataTable modules = consultas.Modulos(userName, dataBaseID, idioma);

            //JPB: Ordeno el DataTable que contiene los Modulos visibles                            
            DataView dv = modules.DefaultView;            
            switch (Utils.Tipo_Ordenamiento_Modulos())
            {
                case "1": dv.Sort = "AccesosMRU DESC,menudesabr ASC"; break;
                case "2": dv.Sort = "menuponderacion DESC,menudesabr ASC"; break;
                case "3": dv.Sort = "menudesabr ASC"; break;
                default: dv.Sort = "menudesabr ASC"; break;   
            }

            //JPB: Filtro el modulo HOME para que no se visualice
            dv.RowFilter = "  menuname <> 'HOME' ";

            modules = dv.ToTable();

            int posicion = 0;
            foreach (DataRow dataRow in modules.Rows)
            {
                if (Habilitado)
                    condicion = ((dataRow[DataBaseActionFieldName].ToString() != "#") && (dataRow[DataBaseActionFieldName].ToString() != ""));
                else
                    condicion = ((dataRow[DataBaseActionFieldName].ToString() == "#") || (dataRow[DataBaseActionFieldName].ToString() == "")); 
                
                if (condicion)
                {
                    Module footerPage = new Module
                    {
                        Id = Guid.NewGuid(),
                        //Action = dataRow[DataBaseActionFieldName].ToString(),
                        Action = dataRow[DataBaseActionFieldName].ToString().Replace("abrirVentana('", "abrirVentana('../"),
                        LinkDvd = dataRow[DataBaseLinkDvdFieldName].ToString(),
                        LinkManual = dataRow[DataBaseLinkManualFieldName].ToString(),
                        MenuDetail = dataRow[DataBaseMenuDetalleFieldName].ToString(),
                        MenuObjective = dataRow[DataBaseMenuObjetiveFieldName].ToString(),
                        MenuObjectiveDetail = dataRow[DataBaseMenuObjetiveDetailFieldName].ToString(),
                        MenuTitle = dataRow[DataBaseMenuTitleFieldName].ToString(),
                        MenuName =  dataRow[DataBaseMenuNameFieldName].ToString(),
                        MenuMsnro = dataRow[DataBaseMenuMsnroFieldName].ToString(),
                        MenuRaiz = dataRow[DataBaseMenuRaizFieldName].ToString(),
                        pos = posicion,
                    };
                    posicion++;
                    ModuleToReturn.Add(footerPage);
                }

            }


            return ModuleToReturn;
        }



        /// <summary>
        /// Verifica si el usuario puede acceder a un determinado modulo
        /// </summary>
        /// <param name="userName">Nombre de Usuario</param>
        /// <param name="dataBaseID">ID de la base</param>
        /// <returns>Lista de Modulos habilitados</returns>
        /// <param name="idioma">Idioma actual</param>
        public static bool Puede_Acceder(string userName, string dataBaseID, string idioma, string menuname)
        {
            
            bool puede = false;
            Consultas consultas = new Consultas();

            //consultas.Credentials = System.Net.CredentialCache.DefaultCredentials;
            DataTable modules = consultas.Modulos(userName, dataBaseID, idioma);
            
            foreach (DataRow dataRow in modules.Rows)
            {
                //Verifico que el menuname sea el modulo que traigo de la consulta
                if (menuname == dataRow[DataBaseMenuNameFieldName].ToString())
                {   //Verifico que el action no esté vacio
                    if ((dataRow[DataBaseActionFieldName].ToString() != "#") && (dataRow[DataBaseActionFieldName].ToString() != ""))
                    {
                        puede = true;                       
                    }
                    break;
                }
            }

            return puede;
        }



        /// <summary>
        /// Busca lista de Module
        /// </summary>
        /// <param name="dataBaseID">ID de la base</param>
        /// <returns>Lista de Module</returns>
        /// <param name="idioma">Idioma actual</param>
        public static List<Module> Find(string dataBaseID, string idioma, bool integrateSecurity)
        {
            return Find(string.Empty, dataBaseID,idioma);
        }

    }
}
