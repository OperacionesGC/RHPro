﻿using System;
using System.Collections.Generic;
using System.Web;
using System.Configuration;
using System.Net;
//using System.DirectoryServices.Protocols;
using System.Security.Cryptography.X509Certificates;
using System.Diagnostics;

using Novell.Directory.Ldap;
using Novell.Directory.Ldap.Utilclass;

namespace ConsultaBaseC
{
    public class LDAP
    {
        private string LDAP_Service;
        private string LDAP_Server;
        private string LDAP_Port;
        private string LDAP_UserForSearch;
        private string LDAP_PassForSearch;
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
            //Obtengo los datos del logueo
            LDAP_UserForSearch = "";
            LDAP_PassForSearch = "";
            try
            {
                LDAP_UserForSearch = Encriptar.Decrypt(DAL.EncrKy(), ConfigurationManager.AppSettings["LDAP_UserForSearch"].ToString().Trim());
                LDAP_PassForSearch = Encriptar.Decrypt(DAL.EncrKy(), ConfigurationManager.AppSettings["LDAP_PassForSearch"].ToString().Trim());
            }
            catch { }

            LDAP_SecureConnection = ConfigurationManager.AppSettings["LDAP_SecureConnection"].ToString().Trim().ToLower();
            LDAP_SecurePort = ConfigurationManager.AppSettings["LDAP_SecurePort"].ToString().Trim();
            LDAP_CheckCertificate = ConfigurationManager.AppSettings["LDAP_CheckCertificate"].ToString().Trim().ToLower();
            LDAP_GroupName = ConfigurationManager.AppSettings["LDAP_GroupName"].ToString().Trim().ToLower();
            LDAP_MembershipAttributeName = ConfigurationManager.AppSettings["LDAP_MembershipAttributeName"].ToString().Trim().ToLower();            
        }





        public bool usuarioValido(string usuario, string pwd)
        {
            DAL.AddLogEvent("Inicio Valización con LDAP", EventLogEntryType.Information, 900);
            /*Parametro que recupero del web.config*/            
            string LDAP_Service = ConfigurationManager.AppSettings["LDAP_Service"].ToString().Trim().ToLower();
            string LDAP_Server = ConfigurationManager.AppSettings["LDAP_Server"].ToString().Trim();
            int LDAP_Port = int.Parse(ConfigurationManager.AppSettings["LDAP_Port"]);
            string LDAP_SecureConnection = ConfigurationManager.AppSettings["LDAP_SecureConnection"].ToString().Trim().ToLower();
            int LDAP_SecurePort = int.Parse(ConfigurationManager.AppSettings["LDAP_SecurePort"]);
            string LDAP_CheckCertificate = ConfigurationManager.AppSettings["LDAP_CheckCertificate"].ToString().Trim().ToLower();
            string LDAP_GroupName = ConfigurationManager.AppSettings["LDAP_GroupName"].ToString().Trim().ToLower();
            string LDAP_MembershipAttributeName = ConfigurationManager.AppSettings["LDAP_MembershipAttributeName"].ToString().Trim();            
            string LDAP_SearchBase = ConfigurationManager.AppSettings["LDAP_SearchBase"].ToString().Trim();
            string LDAP_FilterClass = ConfigurationManager.AppSettings["LDAP_FilterClass"].ToString().Trim();
                        
            bool Salida = false;

            LdapConnection conn = new LdapConnection();
            conn.Connect(LDAP_Server, LDAP_Port);
            conn.Bind(LDAP_UserForSearch, LDAP_PassForSearch);
           
            string searchFilter = "(&(objectclass=" + LDAP_FilterClass + ")(cn=" + usuario + "))";

            try
            {                    
                //Inicio el esquema de busqueda con el filtro especifico
                LdapSchema dirschema = conn.FetchSchema(conn.GetSchemaDN());
                LdapSearchResults lsc = conn.Search(LDAP_SearchBase,
                                                LdapConnection.SCOPE_SUB,
                                                searchFilter,
                                                null,
                                                false);
                //Ciclo por cada entrada devuelta en el esquema de busqueda
                while (lsc.hasMore())
                {
                    LdapEntry nextEntry = null;
                    try
                    {
                        nextEntry = lsc.next();
                    }
                    catch (LdapException e)
                    {

                        continue;
                    }

                    string loginDN = nextEntry.DN.ToString();
                    string password = pwd;

                    //Creo la Conexion LDAP para validar
                    LdapConnection connCheck = new LdapConnection();
                    //Conecto con LDAP
                    connCheck.Connect(LDAP_Server, LDAP_Port);
                    connCheck.Bind(loginDN, password);

                    LdapAttributeSet attributeSet = nextEntry.getAttributeSet();

                    foreach (LdapAttribute attributeP in attributeSet)
                    {
                        string attributeName = attributeP.Name;
                        string[] ienum = attributeP.StringValueArray;

                        foreach (string valor in ienum)
                        {
                            if ((attributeName == LDAP_MembershipAttributeName) && ((string)valor.ToLower().Trim() == (string)LDAP_GroupName.ToLower().Trim()))
                            {
                                Salida = true;
                                break;
                            }
                        }
                    }

                }
                //Desconecta LDAP
                conn.Disconnect();

                return Salida;

            }
            catch (Exception ex)
            {
                DAL.AddLogEvent("Error en validación LDAP: "+ex.Message+"\n\n"+ex.StackTrace, EventLogEntryType.Error, 901);
            }

            return Salida;
        }


        public bool usuarioValido_OLD(string Usuario, string Password)
        {
            /*
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
                        throw ex; 
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
            */

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
        /*
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
                 throw ex;
              
            }
        }
         */ 
    }
}
