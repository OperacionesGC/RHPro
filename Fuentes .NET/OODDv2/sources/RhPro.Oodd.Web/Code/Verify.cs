using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Globalization;
using System.Text.RegularExpressions;


public static class Verify
{
    static bool ret = false;

    public static bool IsNumeric (Object Expression)
    {
        if(Expression == null || Expression is DateTime)
            return false;

        if(Expression is Int16 || Expression is Int32 || Expression is Int64 || Expression is Decimal || 
            Expression is Single || Expression is Double || Expression is Boolean)
            return true;
   
        try 
        {
            if(Expression is string)
                Double.Parse(Expression as string);
            else
                Double.Parse(Expression.ToString());
                return true;
            } catch {}

            return false;
        }

    public static bool IsValidEmail(string pEmail)
    {
        ret = false;

        if (String.IsNullOrEmpty(pEmail))
            return false;

        // Use IdnMapping class to convert Unicode domain names.
        pEmail = Regex.Replace(pEmail, @"(@)(.+)$", DomainMapper);
        if (ret)
            return false;

        // Return true if pEmail is in valid e-mail format.
        return Regex.IsMatch(pEmail,
                @"^(?("")(""[^""]+?""@)|(([0-9a-z]((\.(?!\.))|[-!#\$%&'\*\+/=\?\^`\{\}\|~\w])*)(?<=[0-9a-z])@))" +
                @"(?(\[)(\[(\d{1,3}\.){3}\d{1,3}\])|(([0-9a-z][-\w]*[0-9a-z]*\.)+[a-z0-9]{2,17}))$",
                RegexOptions.IgnoreCase);
    }

    private static string DomainMapper(Match match)
    {
        // IdnMapping class with default property values.
        IdnMapping idn = new IdnMapping();

        string domainName = match.Groups[2].Value;
        try
        {
            domainName = idn.GetAscii(domainName);
        }
        catch (ArgumentException)
        {
            ret = true;
        }
        return match.Groups[1].Value + domainName;
    }
    
}

