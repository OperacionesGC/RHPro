using System;
using System.Collections.Generic;
using System.Data;
using System.Configuration;
using RHPro.ReportesAFD.DataLayer.Data;

namespace RHPro.ReportesAFD.BussinesLayer.Biz
{
    public class RHProBiz
    {
        public static DataTable ObtenerRegistros(string NombreTablaVista, int CantidadRegistros, List<string> NombreCampos)
        {
            if (NombreTablaVista.Contains("SELECT"))
                return RHProData.ObtenerRegistros(true, NombreTablaVista, CantidadRegistros, NombreCampos);
            else
                return RHProData.ObtenerRegistros(false, NombreTablaVista, CantidadRegistros, NombreCampos);
        }
        public static int ObtenerCantRegistros(string NombreTablaVista)
        {
            if (NombreTablaVista.Contains("SELECT"))
                return RHProData.ObtenerCantRegistros(true, NombreTablaVista);
            else
                return RHProData.ObtenerCantRegistros(false, NombreTablaVista);
        }

        //public static int ObtenerCantRegistrosSegunTipoAFD(string NombreTablaVista,int tipoAFD)
        //{
        //    if (tipoAFD == -1) {
        //        if (NombreTablaVista.Contains("SELECT"))
        //            return RHProData.ObtenerCantRegistros(true, NombreTablaVista);
        //        else
        //            return RHProData.ObtenerCantRegistros(false, NombreTablaVista);
                   
        //    }

        //    if (tipoAFD == 0)
        //    {
        //        if (NombreTablaVista.Contains("SELECT"))
        //            return RHProData.ObtenerCantRegistrosAFD(true, NombreTablaVista);
        //        else
        //            return RHProData.ObtenerCantRegistrosAFD(false, NombreTablaVista);

        //    }

        //}

        public static string ObtenerNombrePlantillaXML(int NroReporte)
        {
            try
            {
                return RHProData.ObtenerNombrePlantillaXML(NroReporte);
            }
            catch
            {
                throw new Exception("No se ha encontrado el reporte cargado en la DB.");
            }
        }
    }
}
