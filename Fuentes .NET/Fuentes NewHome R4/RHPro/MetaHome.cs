using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Configuration;
using ServicesProxy.MetaHome;
using Common;
using System.Data;
using ServicesProxy.rhdesa;
using System.Security.Cryptography;

namespace RHPro
{
    public class MetaHome  
    {
        private int Conf_TiempoConexion;
        private int Conf_TiempoProcesamiento;
        

        MH_Externo MHome_ext = new MH_Externo();
         

        public  void Iniciar_Ws_Ext()
        {
            try{
                Conf_TiempoConexion = Convert.ToInt32(ConfigurationManager.AppSettings["Conf_TiempoConexion"]);
                Conf_TiempoProcesamiento = Convert.ToInt32(ConfigurationManager.AppSettings["Conf_TiempoProcesamiento"]);
            }
            catch
            {
                Conf_TiempoConexion = 2000;
                Conf_TiempoProcesamiento = 18000;
            }
            
            MHome_ext.Timeout = Conf_TiempoConexion;  
            MHome_ext.Url = ConfigurationManager.AppSettings["RootWS_MetaHome"];
            MHome_ext.Credentials = System.Net.CredentialCache.DefaultCredentials;
            //Actualizo el tiempo de procesamiento para que al llamar a un metodo de ws_ext pueda esperar un tiempo diferente al de conexion
            MHome_ext.Timeout = Conf_TiempoProcesamiento;      
        }


        public String[] get_Data_TokenAcceso(string token)
        {
            string TokenMeta = MHome_ext.GetToken();
            string keyToken = Armar_Key_Token(TokenMeta);
            return MHome_ext.get_Data_TokenAcceso(keyToken, token);
        }

//         public static string Hash(string ToHash)
        public  string Hash(string ToHash)
        {                       
            System.Text.Encoder enc = System.Text.Encoding.ASCII.GetEncoder();
         
            byte[] data = new byte[ToHash.Length];
            enc.GetBytes(ToHash.ToCharArray(), 0, ToHash.Length, data, 0, true);
         
            System.Security.Cryptography.MD5 md5 = new MD5CryptoServiceProvider();
            byte[] result = md5.ComputeHash(data);

            return BitConverter.ToString(result).Replace("-", "").ToLower();
        }


        /// <summary>
        /// Metodo que arma el token de credencial para pasarselo al web services de las alertas
        /// </summary>
        /// <param name="token"></param>
        /// <returns></returns>
        //public static string Armar_Key_Token(string token)
         public  string Armar_Key_Token(string token)
        {
            String UserName = "575648247C67616C103A0119";            
            String Password = "4A525F2D2125223853695749414A40B8BDB4BBA2AA929890";           

            String ToHash = UserName.ToUpper() + "|" + Password + "|" + token;
            return Hash(ToHash) + "|" + UserName;

        }


        public bool MetaHome_Activo()
        {
            return Convert.ToBoolean(ConfigurationManager.AppSettings["Meta_ModalidadSaaS"]);
        }
 
        public bool MetaHome_RegistraLoguin()
        {
            return MetaHome_Activo() && Convert.ToBoolean(ConfigurationManager.AppSettings["Meta_RegistraLogin"]);
        }

        public String MetaHome_TipoFiltroLogin()
        {
            return Convert.ToString(ConfigurationManager.AppSettings["Meta_TipoFiltroLogin"]);
        }

        public bool MetaH_Contenido(List<int> ListaContenidos)
        {
            Consultas consultas = new Consultas();
            DataTable bases = consultas.comboBase();
            foreach (DataRow dr in bases.Rows)
            {
                if (ListaContenidos.Contains(Convert.ToInt32(dr["key"]))  )
                    return true;                

            }
            return false;
        }

        public List<int> MetaHome_getBases(String URL, String Usuario, String Pass)
        {

            try
            {           
                string EncryptionKey = (String)ConfigurationManager.AppSettings["EncryptionKey"];
                int[] Arr= {};

                string TokenMeta = MHome_ext.GetToken();
                string keyToken = Armar_Key_Token(TokenMeta);
                Arr = MHome_ext.get_Bases(keyToken, URL, Usuario, Encryptor.Encrypt(EncryptionKey, Pass), MetaHome_TipoFiltroLogin());

                if (Arr == null) return null;
                 
                return Arr.ToList();                
            }            
            catch (Exception ex) {
                throw ex;    
            }
        }

        public List<String> MetaHome_fromLogin(string idTemp)
        {
            List<String> salida = new List<string>();
            string EncryptionKey = (String)ConfigurationManager.AppSettings["EncryptionKey"];

            string nroTemp = Encryptor.Decrypt(EncryptionKey, idTemp);
            string keyToken = Armar_Key_Token(MHome_ext.GetToken());
            DataSet DS = MHome_ext.get_UsuarioLogueado(keyToken, nroTemp);
            string UsrDecript;
            string PassDecript;
            foreach (DataRow fila in DS.Tables[0].Rows)
            {
                UsrDecript = Encryptor.Decrypt(EncryptionKey, (string)fila["usuario"]);
                PassDecript = Encryptor.Decrypt(EncryptionKey, (string)fila["password"]);
                salida.Add(UsrDecript);
                salida.Add(PassDecript);
                salida.Add(Convert.ToString(fila["base"]));
                break;               
            }

            return salida;
        }


        public void MetaHome_Logout()
        {
            try
            {
                string EncryptionKey = (String)ConfigurationManager.AppSettings["EncryptionKey"];               
                string idTemp = (String)Utils.SessionNroTempLogin;
                string nroTemp = Encryptor.Decrypt(EncryptionKey, idTemp);
                string keyToken = Armar_Key_Token(MHome_ext.GetToken());
                MHome_ext.logout_TempLogin(keyToken, nroTemp);
                Utils.SessionNroTempLogin = null;
            }
            catch (Exception ex) { throw ex; }
        }


      //  public bool MetaHome_ActualizaMultiplesBases(String usuario, String PassEncript, String InitialCatalog, String DataSource)
        public string MetaHome_ActualizaMultiplesBases(String usuario, String PassEncript, String InitialCatalog, String DataSource,
            String  keyConnStr,String URL)
        {
            try
            {                
                string keyToken = Armar_Key_Token(MHome_ext.GetToken());
               // bool valor = MHome_ext.WS_Actualizar_Multiples_Bases(keyToken, usuario, PassEncript, InitialCatalog, DataSource);
                string valor = MHome_ext.WS_Actualizar_Multiples_Bases(keyToken, usuario, PassEncript, InitialCatalog, DataSource,  keyConnStr, URL);
 
                return valor;
                 
            }
            catch (Exception ex) { 
            //    return false;
                return "Intento de conexión fallida";
            }
        }



    }
}
