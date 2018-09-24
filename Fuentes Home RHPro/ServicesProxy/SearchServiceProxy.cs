using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using Entities;
using ServicesProxy.ar.com.rhpro.prueba;

namespace ServicesProxy
{
    public static class SearchServiceProxy
    {
        private const string DataBaseModuleFieldName = "modulo";
        private const string DataBaseMenuDescriptionFieldName = "DescrMenu";
        private const string DataBaseActionFieldName = "accion";
        private const string DataBaseDescriptionFieldName = "DescrExt";

        /// <summary>
        /// Devuelve
        /// </summary>
        /// <param name="userName">El usuario logueado</param>
        /// <param name="palabra">La palabra a buscar</param>
        /// <param name="dataBaseID">La base de datos seleccionada</param>
        /// <param name="idioma">Idioma actual</param>
        /// <returns> Lista de Resultados</returns>
        public static List<Search> Find(string userName, string palabra, string dataBaseID, string idioma)
        {
            List<Search> SearchToReturn = new List<Search>();

            Consultas consultas = new Consultas();

            consultas.Credentials = System.Net.CredentialCache.DefaultCredentials;

            DataTable dataBases = consultas.Search(userName, palabra, dataBaseID, idioma);

            foreach (DataRow dataRow in dataBases.Rows)
            {
                Search search = new Search
                {
                    Action = dataRow[DataBaseActionFieldName].ToString(),
                    MenuDescription = dataRow[DataBaseMenuDescriptionFieldName].ToString(),
                    Module = dataRow[DataBaseModuleFieldName].ToString(),
                    Description = dataRow[DataBaseDescriptionFieldName].ToString()
                };

                SearchToReturn.Add(search);
            }

            return SearchToReturn;
        }

        public static List<Search> Find(string palabra, string dataBaseID, string idioma, bool integrateSecurity)
        {
            return Find(string.Empty, palabra, dataBaseID, idioma);
        }
    }
}
