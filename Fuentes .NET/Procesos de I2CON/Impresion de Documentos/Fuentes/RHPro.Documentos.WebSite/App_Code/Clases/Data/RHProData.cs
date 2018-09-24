using System;
using System.Collections.Generic;
using System.Data;
using System.Configuration;
using i2Con.Data.Connection;
using RHPro.ReportesAFD.Clases;

namespace RHPro.ReportesAFD.DataLayer.Data
{
    public class RHProData
    {
        private static string _sql;
        public static DataTable ObtenerRegistros(bool EsVista, string NombreTablaVista, int CantidadRegistros, List<string> NombreCampos)
        {
            string aux = " ";
            _sql = "";

            if (CantidadRegistros == 0) _sql += " SELECT ";
            else _sql += " SELECT TOP " + CantidadRegistros;
            for (int index = 0; index < NombreCampos.Count; index++)
            {
                _sql += aux + NombreCampos[index];
                aux = ", ";
            }

            if (EsVista)
                _sql += " FROM ( " + NombreTablaVista + ") AS VistaAux ";
            else
                _sql += " FROM " + NombreTablaVista;

            DataTable dt = new DataTable();
            dt = I2Database.CreateDataSet(AppSession.RHProDBConnection, _sql).Tables[0];
            return dt;
        }
        public static int ObtenerCantRegistros(bool EsVista, string NombreTablaVista)
        {
            _sql = "";

            _sql += " SELECT COUNT (*) ";
            if (EsVista)
                _sql += " FROM ( " + NombreTablaVista + ") AS VistaAux ";
            else
                _sql += " FROM " + NombreTablaVista;

            DataTable dt = new DataTable();
            dt = I2Database.CreateDataSet(AppSession.RHProDBConnection, _sql).Tables[0];
            return int.Parse(dt.Rows[0][0].ToString());
        }
        public static string ObtenerNombrePlantillaXML(int NroReporte)
        {
            _sql = "";

            _sql += " SELECT repdesc ";
            _sql += " FROM reporte ";
            _sql += " WHERE repnro = " + NroReporte;

            DataTable dt = new DataTable();
            dt = I2Database.CreateDataSet(AppSession.RHProDBConnection, _sql).Tables[0];
            return dt.Rows[0][0].ToString();
        }
    }
}
