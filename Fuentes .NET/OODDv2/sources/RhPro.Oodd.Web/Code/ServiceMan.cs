using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using RhPro.Oodd.Web.OrgService;

public static class ServiceMan
{
    public static ConsultasSoapClient Get()
    {
        ConsultasSoapClient ret = null;

        string endPointConfigurationName = "ConsultasSoap";

        string remoteAddressService = AppSettingMan.GetChainValue
            ("remoteAddressService");

        if (!string.IsNullOrEmpty(remoteAddressService))
        {
            ret = new ConsultasSoapClient(endPointConfigurationName, 
                remoteAddressService);
        }

        return ret;
    }
}
