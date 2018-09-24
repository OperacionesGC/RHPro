using System;
using System.Data;
using System.Configuration;

namespace RHPro.ReportesAFD.Clases
{
    public class AppSession
    {
        //private static string _rhproDBConnection = "LOCAL";
        private static string _rhproDBConnection = "1";
        public static string RHProDBConnection { get { return _rhproDBConnection; } set { _rhproDBConnection = value; } }

        /// <summary>
        /// Guarda id de la base en la que se encuentra conectado NG - 23/02/2016
        /// </summary>
        /// <param name="id"></param>
        public static void Base(string id)
        {
            _rhproDBConnection = id;
        }

    }

}