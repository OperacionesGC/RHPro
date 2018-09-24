using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using Common;
//using ServicesProxy.ar.com.rhpro.prueba;
using ServicesProxy.rhdesa;



using Entities;

namespace ServicesProxy
{
   public static class LoginServiceProxy
   {
       private const string DataBaseDescFieldName = "Ingresa";
       private const string DataBaseImageFieldName = "mensaje";
       private const string DataBaseDescextFieldName = "CambiarPass";
       private const string DataBaseLenguaje = "Lenguaje";
       private const string DataBaseMaxEmpl = "MaxEmpl";

       /// <summary>
       /// Devuelve Mensaje de Login
       /// </summary>
       /// <param name="userName">Nombre del Usuario</param>
       /// <param name="password">Contraseña del Usuario</param>
       /// <param name="integrateSecurity">Si posee Seguridad Integrada</param>
       /// <param name="dataBaseID">ID de la base de Datos</param>
       /// <returns>Lista de Mensajes</returns>
       /// <param name="idioma">Idioma actual</param>
       public static Login Find(string userName, string password, string encryptionKey, Utils.IntegrateSecurityConstants integrateSecurity, string dataBaseID, bool? encriptUserData, string idioma)
        {
 
           Login LoginToReturn = null;

            Consultas consultas = new Consultas();
            
            consultas.Credentials = System.Net.CredentialCache.DefaultCredentials;

            if (encriptUserData.HasValue && encriptUserData.Value)
            {
                userName = Encryptor.Encrypt(encryptionKey, userName);
                password = Encryptor.Encrypt(encryptionKey, password);
            }

            DataTable dataBases = consultas.Login(userName, password, integrateSecurity.ToString(), dataBaseID, idioma);

            foreach (DataRow dataRow in dataBases.Rows)
            {
                LoginToReturn  = new Entities.Login
                {
                    IsValid = bool.Parse(dataRow[DataBaseDescFieldName].ToString()),
                    Messege = dataRow[DataBaseImageFieldName].ToString(),
                    RequiredChangePassword = bool.Parse(dataRow[DataBaseDescextFieldName].ToString()),
                    Lenguaje = dataRow[DataBaseLenguaje].ToString(),
                    MaxEmpl = dataRow[DataBaseMaxEmpl].ToString(),
                };                
            }
 
            return LoginToReturn;            
        }

        public static Login Find(string password, string encryptionKey, Utils.IntegrateSecurityConstants integrateSecurity, string dataBaseID, bool? encriptUserData, string idioma)
        {
            return Find(null, password, encryptionKey, integrateSecurity, dataBaseID, encriptUserData, idioma);
        }
    }
}
