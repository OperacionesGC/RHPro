using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using Entities;
using ServicesProxy.ar.com.rhpro.prueba;

namespace ServicesProxy
{
   public static class MRUServiceProxy
    {

        private const string DataBaseMenuNameFieldName = "menuname";
        private const string DataBaseRootFieldName = "raiz";
        private const string DataBaseActionFieldName = "action";

       /// <summary>
       /// Busca los MRU
       /// </summary>
       /// <param name="userName">El usuario logueado</param>
       /// <param name="cant">La cantidad de mensajes a mostrar </param>
       /// <param name="dataBaseID">El id de la base seleccionada</param>
       /// <returns>Lista de  MRU</returns>
        /// <param name="idioma">Idioma actual</param>
        public static List<MRU> Find(string userName, int cant, string dataBaseID, string idioma)
        {
            List<MRU> MruToReturn = new List<MRU>();

            Consultas consultas = new Consultas();

            consultas.Credentials = System.Net.CredentialCache.DefaultCredentials;

            DataTable dataBases = consultas.MRU(userName ?? string.Empty, cant, dataBaseID, idioma);

            foreach (DataRow dataRow in dataBases.Rows)
            {
                MRU mru = new MRU
                {
                    Action = dataRow[DataBaseActionFieldName].ToString(),
                    MenuName = dataRow[DataBaseMenuNameFieldName].ToString(),
                    Root = dataRow[DataBaseRootFieldName].ToString()


                };

                MruToReturn.Add(mru);
            }

            return MruToReturn;
        }
        public static List<MRU> Find(int cant, string dataBaseID, string idioma, bool integrateSecurity)
        {
            return Find(string.Empty, cant, dataBaseID, idioma);
        }

    }
}
