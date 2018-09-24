using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ServicesProxy.rhdesa;

namespace ServicesProxy
{
    public static class ChangePasswordServiceProxy
    {
        /// <summary>
        /// Cambiar el password de un usuario
        /// </summary>
        /// <param name="userName">Nombre del usuario</param>
        /// <param name="oldPassword">Contraseña actual</param>
        /// <param name="newPassword">Nueva contraseña</param>
        /// <param name="confirmPassword">Confirmacion de la nueva contraseña</param>
        /// <param name="dataBaseID">El Id de la base de datos correspondiente</param>
        /// <returns>Un mensage de error o vacio si se pudo hacer el cambio</returns>
        /// <param name="idioma">Idioma actual</param>
        public static string ChangePassword(string userName, string oldPassword, string newPassword, string confirmPassword, string dataBaseID, string idioma)
        {
            Consultas consultas = new Consultas();

            //consultas.Credentials = System.Net.CredentialCache.DefaultCredentials;
                        
            return consultas.CambiarPass(userName, oldPassword, newPassword, confirmPassword, dataBaseID, idioma);            
        }
    }
}
