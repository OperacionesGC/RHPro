using System;
using System.Collections.Generic;
using System.Web;
using System.Configuration;
using System.Net;
using System.DirectoryServices.Protocols;
using System.Security.Cryptography.X509Certificates;

namespace ConsultaBaseC
{
    public class LDAP
    {
        private string LDAP_Service;
        private string LDAP_Server;
        private string LDAP_Port;
        private string LDAP_SecureConnection;
        private string LDAP_SecurePort;
        private string LDAP_CheckCertificate;
        private string LDAP_GroupName;
        private string LDAP_MembershipAttributeName;
        private string LDAP_DNFormat;

        public LDAP()
        {
            LDAP_Service = ConfigurationManager.AppSettings["LDAP_Service"].ToString().Trim().ToLower();
            LDAP_Server = ConfigurationManager.AppSettings["LDAP_Server"].ToString().Trim();
            LDAP_Port = ConfigurationManager.AppSettings["LDAP_Port"].ToString().Trim();
            LDAP_SecureConnection = ConfigurationManager.AppSettings["LDAP_SecureConnection"].ToString().Trim().ToLower();
            LDAP_SecurePort = ConfigurationManager.AppSettings["LDAP_SecurePort"].ToString().Trim();
            LDAP_CheckCertificate = ConfigurationManager.AppSettings["LDAP_CheckCertificate"].ToString().Trim().ToLower();
            LDAP_GroupName = ConfigurationManager.AppSettings["LDAP_GroupName"].ToString().Trim().ToLower();
            LDAP_MembershipAttributeName = ConfigurationManager.AppSettings["LDAP_MembershipAttributeName"].ToString().Trim().ToLower();
            LDAP_DNFormat = ConfigurationManager.AppSettings["LDAP_DNFormat"].ToString().Trim();
        }

        public bool usuarioValido(string Usuario, string Password)
        {
            bool result = false;
            bool secureCn = false;
            string port;
            LdapConnection cn = new LdapConnection("");
            string dns = LDAP_DNFormat.Replace("USR", Usuario).Replace("\r\n", "");
            string dn = "";
            string[] units;
            bool connected = false;

            switch (LDAP_Service)
            {
                case "edirectory":

                    try
                    {
                        if (LDAP_SecureConnection == "true")
                        {
                            secureCn = true;
                            port = LDAP_SecurePort;
                        }
                        else
                        {
                            secureCn = false;
                            port = LDAP_Port;
                        }

                        units = dns.Split(";".ToCharArray());

                        //Itero por las distintas unidades organizacionales para iniciar sesión de usuario

                        for (int i = 0; i < units.Length; i++)
                        {
                            dn = units[i].Trim();

                            cn = new LdapConnection(LDAP_Server + ":" + port);
                            cn.SessionOptions.SecureSocketLayer = secureCn;
                            cn.SessionOptions.VerifyServerCertificate = new VerifyServerCertificateCallback(ServerCallback);
                            cn.AuthType = AuthType.Basic;
                            NetworkCredential nc = new NetworkCredential(dn, Password);
                            cn.Credential = nc;

                            if (Connect(cn))
                            {
                                connected = true;
                                break;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                    }

                    if (connected)
                    {
                        try
                        {
                            //Verifico si el usuario es miembro del grupo especificado

                            SearchRequest searchRequest = new SearchRequest(dn, "(objectclass=user)", SearchScope.Subtree, null);
                            SearchResponse searchResponse = (SearchResponse)cn.SendRequest(searchRequest);

                            if (searchResponse.Entries.Count != 0)
                            {
                                SearchResultAttributeCollection attributes = searchResponse.Entries[0].Attributes;

                                if (attributes.Count != 0)
                                {
                                    foreach (DirectoryAttribute attribute in attributes.Values)
                                    {
                                        string attr = attribute.Name;

                                        if (attr.ToLower().Trim() == LDAP_MembershipAttributeName)
                                        {
                                            for (int i = 0; i < attribute.Count; i++)
                                            {
                                                string grupo = attribute[i].ToString().Trim().ToLower();

                                                if (grupo.Contains(LDAP_GroupName))
                                                {
                                                    result = true;
                                                    break;
                                                }
                                            }
                                        }

                                        if (result)
                                            break;
                                    }
                                }
                            }

                            cn = null;
                            return result;

                        }
                        catch (Exception ex)
                        {
                            cn = null;
                            return false;
                        }
                    }
                    else
                    {
                        cn = null;
                        return false;
                    }
 
                case "activedirectory":

                    return false;
            }

            return false;
        }

        private bool ServerCallback(LdapConnection connection, X509Certificate certificate)
        {
            if (LDAP_CheckCertificate == "true")
            {
                //Implementar verificación de certificado
                
            }

            return true;
        }

        private bool Connect(LdapConnection connection)
        {
            try
            {
                connection.Bind();
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }
    }
}
