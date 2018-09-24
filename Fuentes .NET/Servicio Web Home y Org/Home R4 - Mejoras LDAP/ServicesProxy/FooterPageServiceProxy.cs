using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using Entities;
using ServicesProxy.rhdesa;

namespace ServicesProxy
{
    public static class FooterPageServiceProxy
    {
        private const string DataBaseTitleFieldName = "hpptitulo";
        private const string DataBaseUrlFieldName = "hpppagina";
    
        /// <summary>
        /// Busca Footers
        /// </summary>
        /// <param name="userName">Nombre de Usuario</param>
        /// <param name="dataBaseID">ID de la base de datos</param>
        /// <returns>Lista de Footers</returns>
        /// <param name="idioma">Idioma actual</param>
        public static FooterPage Find(string userName, string dataBaseID, string idioma)
        { 
            FooterPage FooterPageToReturn = new FooterPage();

            Consultas consultas = new Consultas();

            consultas.Credentials = System.Net.CredentialCache.DefaultCredentials;

            DataTable dataBases = consultas.PagPie(userName, dataBaseID, idioma);

            foreach (DataRow dataRow in dataBases.Rows)
            {
                FooterPage footerPage = new FooterPage
                {
                    PageUrl = dataRow[DataBaseTitleFieldName].ToString(),
                    Title = dataRow[DataBaseUrlFieldName].ToString()
                    

                };

                FooterPageToReturn = footerPage;
            }

            return FooterPageToReturn;
        }

        /// <summary>
        /// Busca Footers
        /// </summary>
        /// <param name="dataBaseID">ID de la base de datos</param>
        /// <returns>Lista de Footers</returns>
        /// <param name="idioma">Idioma actual</param>
        public static FooterPage Find(string dataBaseID, string idioma, bool integrateSecurity)
        {
            return Find(string.Empty, dataBaseID, idioma);
        }

    }
}
