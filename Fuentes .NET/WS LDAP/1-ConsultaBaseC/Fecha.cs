using System;
using System.Data;
using System.Configuration;
using System.Collections.Generic;
using System.Text;

namespace ConsultaBaseC
{
    public enum DateInterval
    {
        Day,
        DayOfYear,
        Hour,
        Minute,
        Month,
        Quarter,
        Second,
        Weekday,
        WeekOfYear,
        Year
    }

    public class Fecha
    {
        public static string cambiaFecha(string Actual, string TipoBase)
        {
            string Aux = "";

            switch (TipoBase.ToUpper())
            {
                case "SQL":
                    Aux = "'" + Actual.Substring(0, 2) + "/" + Actual.Substring(3, 2) + "/" + Actual.Substring(6, 4) + "'";
                    //auxi  = "'" & mid(actual,1,2) & "/" & mid(actual,4,2) & "/" & mid(actual,7,4) & "'"
                    break;
                case "ORA":
            //if mid(actual,1,1) = "T" then
            //   auxi  = "'" & mid(actual,13,2) & "/" & mid(actual,10,2) & "/" & mid(actual,16,4) & "'"
            //else
            //   auxi  = "TO_Date('" & mid(actual,4,2) & "/" & mid(actual,1,2) & "/" & mid(actual,7,4) & " 12:00:00 AM', 'MM/DD/YYYY HH:MI:SS AM')"
            //end if
                    if (Actual.Substring(0, 1) == "T") 
                        Aux = "'" + Actual.Substring(12, 2) + "/" + Actual.Substring(9, 2) + "/" + Actual.Substring(15, 4) + "'";
                    else
                        Aux = "TO_Date('" + Actual.Substring(3, 2) + "/" + Actual.Substring(0, 2) + "/" + Actual.Substring(6, 4) + " 12:00:00 AM', 'MM/DD/YYYY HH:MI:SS AM')";
                    break;
                default:
                    Aux = "'" + Actual.Substring(0, 2) + "/" + Actual.Substring(3, 2) + "/" + Actual.Substring(6, 4) + "'";
                    break;
            }

            return Aux;
        }

        public static long DateDiff(DateInterval interval, DateTime dt1, DateTime dt2)
        {
            return DateDiff(interval, dt1, dt2, System.Globalization.DateTimeFormatInfo.CurrentInfo.FirstDayOfWeek);
        } 

        private static int GetQuarter(int nMonth)
        {
            if (nMonth <= 3)
                return 1;
            if (nMonth <= 6)
                return 2;
            if (nMonth <= 9)
                return 3;
            return 4;
        }

        private static long Round(double dVal)
        {
            if (dVal >= 0)
                return (long)Math.Floor(dVal);
            return (long)Math.Ceiling(dVal);
        }

        public static long DateDiff(DateInterval interval, DateTime dt1, DateTime dt2, DayOfWeek eFirstDayOfWeek)
        {
            if (interval == DateInterval.Year)
                return dt2.Year - dt1.Year;

            if (interval == DateInterval.Month)
                return (dt2.Month - dt1.Month) + (12 * (dt2.Year - dt1.Year));

            TimeSpan ts = dt2 - dt1;

            if (interval == DateInterval.Day || interval == DateInterval.DayOfYear)
                return Round(ts.TotalDays);

            if (interval == DateInterval.Hour)
                return Round(ts.TotalHours);

            if (interval == DateInterval.Minute)
                return Round(ts.TotalMinutes);

            if (interval == DateInterval.Second)
                return Round(ts.TotalSeconds);

            if (interval == DateInterval.Weekday)
            {
                return Round(ts.TotalDays / 7.0);
            }

            if (interval == DateInterval.WeekOfYear)
            {
                while (dt2.DayOfWeek != eFirstDayOfWeek)
                    dt2 = dt2.AddDays(-1);
                while (dt1.DayOfWeek != eFirstDayOfWeek)
                    dt1 = dt1.AddDays(-1);
                ts = dt2 - dt1;
                return Round(ts.TotalDays / 7.0);
            }

            if (interval == DateInterval.Quarter)
            {
                double d1Quarter = GetQuarter(dt1.Month);
                double d2Quarter = GetQuarter(dt2.Month);
                double d1 = d2Quarter - d1Quarter;
                double d2 = (4 * (dt2.Year - dt1.Year));
                return Round(d1 + d2);
            }

            return 0;
        }
    }
}