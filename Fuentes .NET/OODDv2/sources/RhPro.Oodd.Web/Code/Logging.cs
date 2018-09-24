using System;
using System.Collections.Generic;
using System.Text;
using System.Configuration;
using System.Security;
using System.Diagnostics;

public class Logging
{
    public const string ServiceName = "RHPro Organization Chart";
    /// <summary>
    /// Updates the eventlog
    /// </summary>
    /// <param name="Message">Message for hte eventlog</param>
    /// <param name="msgType">Type of messange (warning, error, information, audit)</param>
    public static void UpdateLog(string Message, EventLogEntryType msgType)
    {
        try
        {
            if (EventLog.SourceExists("Application"))
            {
                System.Diagnostics.EventLog.WriteEntry("Application", Message, msgType);
            }
        }
        catch
        {
            //ignore
        }
    }
}

