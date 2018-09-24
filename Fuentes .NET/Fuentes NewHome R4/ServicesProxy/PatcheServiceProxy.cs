using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using ServicesProxy.rhdesa;

namespace ServicesProxy
{
    public static class PatcheServiceProxy
    {
        /// <summary>
        /// Devuelve el patch
        /// </summary>
        /// <param name="dataBaseID">El id de la base seleccionada</param>
        /// <returns>El patch</returns>
        /// <param name="idioma">Idioma actual</param>
        public static string Find(string dataBaseID, string idioma)
        {
            Consultas consultas = new Consultas();

            //consultas.Credentials = System.Net.CredentialCache.DefaultCredentials;

            return consultas.Patch(dataBaseID, idioma);
        }
        
    }
}
