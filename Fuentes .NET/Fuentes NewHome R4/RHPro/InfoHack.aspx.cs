using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Net;
using System.Net.NetworkInformation;
using Common;

namespace RHPro
{
    public partial class InfoHack : System.Web.UI.Page
    {
        public Lenguaje Obj_Lenguaje;



      
        protected void Page_Load(object sender, EventArgs e)
        {
            Obj_Lenguaje = new Lenguaje();
            Utils.Escribir_Log("Hack_" + DateTime.Now.Day + "_" + DateTime.Now.Month + "_" + DateTime.Now.Year + ".txt", SalidaError());

        }

        public string SalidaError()
        {
            string salida = "";

            string IP = Request.ServerVariables.Get("REMOTE_ADDR");
            string userName = "N/A";// System.Net.Dns.GetHostEntry(IP).HostName;//Request.ServerVariables.Get("AUTH_USER");           
            String Browser = Request.ServerVariables.Get("HTTP_USER_AGENT");

            salida += "********************************************************************************" + System.Environment.NewLine;
            salida += "HACKING: " + DateTime.Now + System.Environment.NewLine + System.Environment.NewLine;
            salida += "> NT User: " + userName + System.Environment.NewLine + System.Environment.NewLine;
            salida += "> IP del Cliente: " + IP + System.Environment.NewLine;
            salida += "> Navegador: " + Browser + System.Environment.NewLine;
            try
            {
               // salida += "> Host del Cliente: " + Dns.GetHostEntry(IP).HostName + System.Environment.NewLine;
            }
            catch { }

            salida += "********************************************************************************" + System.Environment.NewLine;  
         
            return salida;
        }

         

  

        /*
        static string GetInfoPlacaRed()  //Recupera la MAC
        {
            string macAddresses = "";
            foreach (NetworkInterface nic in NetworkInterface.GetAllNetworkInterfaces())
            {
                // Only consider Ethernet network interfaces, thereby ignoring any
                // loopback devices etc.
                if (nic.NetworkInterfaceType != NetworkInterfaceType.Ethernet) continue;
                if (nic.OperationalStatus == OperationalStatus.Up)
                {
                    macAddresses += nic.Description.ToString();
                    break;
                }
            }
            return macAddresses;
        }

 

        static string GetMacAddress()  //Recupera la MAC
        {
            string macAddresses = "";
            foreach (NetworkInterface nic in NetworkInterface.GetAllNetworkInterfaces())
            {            
                if (nic.NetworkInterfaceType != NetworkInterfaceType.Ethernet) continue;
                if (nic.OperationalStatus == OperationalStatus.Up)
                {
                    macAddresses += nic.GetPhysicalAddress().ToString();
                    break;
                }
               
            }
            return macAddresses;
        }

        public static string GetIpAddress()  //  Recupera la IP
        {
           
            string ip = "";
            IPHostEntry ipEntry = Dns.GetHostEntry(GetCompCode());
            IPAddress[] addr = ipEntry.AddressList;
            ip = addr[1].ToString();
            return ip;
        }


        public static string GetCompCode()  // Recupera el nombre de la computadora
        {
            string strHostName = "";
            strHostName = Dns.GetHostName();
            return strHostName;
        }
        */
         

        

    }
}
