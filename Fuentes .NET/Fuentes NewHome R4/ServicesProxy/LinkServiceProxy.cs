using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using Entities;
using ServicesProxy.rhdesa;

namespace ServicesProxy
{
    public static class LinkServiceProxy
    {

        private const string DataBaseURLFieldName = "HlinkPagina";
        private const string DataBaseTituloFieldName = "HlinkTitulo";
        /// <summary>
        /// Busca los Link
        /// </summary>
        /// <param name="userName">Nombre de Usuario</param>
        /// <param name="databaseID">Id de la base de datos</param>
        /// <returns>Lista de Links</returns>
        /// <param name="idioma">Idioma actual</param>
        public static List<Link> Find(string userName, string databaseID, string idioma)
        {
            List<Link> LinksToReturn = new List<Link>();

            Consultas consultas = new Consultas();

            //consultas.Credentials = System.Net.CredentialCache.DefaultCredentials;

            DataTable dataBases = consultas.Link(userName ?? string.Empty, databaseID, idioma);

            foreach (DataRow dataRow in dataBases.Rows)
            {
                Link link = new Link
                                {
                                    Url = dataRow[DataBaseURLFieldName].ToString(),
                                    Title = dataRow[DataBaseTituloFieldName].ToString()
                                };
                LinksToReturn.Add(link);

            }

            return LinksToReturn;
        }




        public static List<Link> Find(string databaseID, string idioma, bool integrateSecurity)
        {
            return Find(string.Empty, databaseID, idioma);
        }
    }
}
