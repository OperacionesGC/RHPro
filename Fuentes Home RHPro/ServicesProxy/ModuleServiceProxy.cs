using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using Entities;
using ServicesProxy.ar.com.rhpro.prueba;

namespace ServicesProxy
{
    public static class ModuleServiceProxy
    {
        private const string DataBaseMenuTitleFieldName = "menudesabr";
        private const string DataBaseMenuDetalleFieldName = "menudetalle";
        private const string DataBaseActionFieldName = "action";
        private const string DataBaseMenuObjetiveDetailFieldName = "menubeneficio";
        private const string DataBaseMenuObjetiveFieldName = "menuobjetivo";
        private const string DataBaseLinkManualFieldName = "linkmanual";
        private const string DataBaseLinkDvdFieldName = "linkdvd";


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

            consultas.Credentials = System.Net.CredentialCache.DefaultCredentials;

            DataTable modules = consultas.Modulos(userName, dataBaseID, idioma);

            foreach (DataRow dataRow in modules.Rows)
            {
                Module footerPage = new Module
                {
                    Id = Guid.NewGuid(),
                    Action = dataRow[DataBaseActionFieldName].ToString(),
                    LinkDvd = dataRow[DataBaseLinkDvdFieldName].ToString(),
                    LinkManual = dataRow[DataBaseLinkManualFieldName].ToString(),
                    MenuDetail = dataRow[DataBaseMenuDetalleFieldName].ToString(),
                    MenuObjective = dataRow[DataBaseMenuObjetiveFieldName].ToString(),
                    MenuObjectiveDetail = dataRow[DataBaseMenuObjetiveDetailFieldName].ToString(),
                    MenuTitle = dataRow[DataBaseMenuTitleFieldName].ToString(), 


                };

                ModuleToReturn.Add(footerPage);
            }

            return ModuleToReturn;
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
