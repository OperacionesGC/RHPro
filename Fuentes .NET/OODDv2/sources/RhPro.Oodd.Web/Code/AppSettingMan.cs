using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;

public static class AppSettingMan
{
    public static bool GetBoolValue(string pKey)
    {
        bool ret = false;

        object enableDebug = ConfigurationManager.AppSettings.Get(pKey);
        if (enableDebug != null)
        {
            if (Verify.IsNumeric(enableDebug))
                enableDebug = Convert.ToInt16(enableDebug);
            else
                if (enableDebug.ToString().ToUpper().Equals("true".ToUpper())
                    || enableDebug.ToString().ToUpper().Equals("false".ToUpper()))
                    ret = Convert.ToBoolean(enableDebug);
        }

        return ret;
    }

    public static string GetChainValue(string pKey)
    {
        string ret = ConfigurationManager.AppSettings.Get(pKey);

        return ret;
    }
}

