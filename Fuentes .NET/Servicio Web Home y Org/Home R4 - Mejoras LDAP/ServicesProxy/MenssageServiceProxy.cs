using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using Entities;
using ServicesProxy.rhdesa;

namespace ServicesProxy
{
    public static class MenssageServiceProxy
    {
    
        private const string DataBaseTitleFieldName = "hmsjtitulo";
        private const string DataBaseBodyFieldName = "hmsjcuerpo";

        /// <summary>
        /// Busca los Message
        /// </summary>
        /// <param name="dataBaseID">ID de la base de datos</param>
        /// <returns>Devuelve lista de Message</returns>
        /// <param name="idioma">Idioma actual</param>
        public static List<Message> Find(string dataBaseID, string idioma)
        {
            List<Message> MessagesToReturn = new List<Message>();

            Consultas consultas = new Consultas();

            consultas.Credentials = System.Net.CredentialCache.DefaultCredentials;

            DataTable dataBases = consultas.Mensaje(dataBaseID, idioma);

            foreach (DataRow dataRow in dataBases.Rows)
            {
                Message message = new Message
                {
                    Body = dataRow[DataBaseTitleFieldName].ToString(),
                    Title = dataRow[DataBaseBodyFieldName].ToString()


                };

                MessagesToReturn.Add(message);
            }

            return MessagesToReturn;
        }

    }
}
