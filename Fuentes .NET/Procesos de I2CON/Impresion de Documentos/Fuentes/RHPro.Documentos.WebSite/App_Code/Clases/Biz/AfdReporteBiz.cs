using System;
using System.Collections.Generic;
using System.Data;
using RHPro.ReportesAFD.BussinesLayer.Base;
using RHPro.ReportesAFD.DataLayer.Data;

namespace RHPro.ReportesAFD.BussinesLayer.Biz
{
	[Serializable]
    public class AfdReporteBiz
    {
		#region Atributos
        private AfdReporteDAO _afdReporteDao = new AfdReporteDAO();
		private IList< AfdReporteSBiz > _lista;
		private DataTable _tabla;
		#endregion
		
		#region Propiedades
		public IList< AfdReporteSBiz > Lista
		{
			get {return this._lista; }
		}
		public DataTable Tabla
		{
			get {return this._tabla; }
		}
		#endregion
		
		#region Metodos Privados
		private DataTable GetDTAfdReporteByFilters(int? Repnro, string Nomtag, string Tablequery, int? Cantreg, string CriterioOrden)
        {
            try
            {
                return this._afdReporteDao.GetAfdReporte(null, Repnro, Nomtag, Tablequery, Cantreg, CriterioOrden);
            }
            catch
            {
                throw;
            }
        }
		#endregion
		
		#region Metodos Publicos
        
		#endregion
		
		#region Constructores
		public AfdReporteBiz()
		{
			try
            {
				DataTable dt;
                dt = this.GetDTAfdReporteByFilters(null, null, null, null, null);
				
				this._lista = new List< AfdReporteSBiz >();
				foreach (DataRow dr in dt.Rows)
				{				
					AfdReporteSBiz objAfdReporte= new AfdReporteSBiz(dr);
					this._lista.Add(objAfdReporte);
				}
				this._tabla = dt;
            }
            catch
            {
                throw;
            }
		}
		public AfdReporteBiz(int? Repnro, string Nomtag, string Tablequery, int? Cantreg, string CriterioOrden)
		{
			try
            {
				DataTable dt;
                dt = this.GetDTAfdReporteByFilters(Repnro, Nomtag, Tablequery, Cantreg, CriterioOrden);
				
				this._lista = new List< AfdReporteSBiz >();
				foreach (DataRow dr in dt.Rows)
				{				
					AfdReporteSBiz objAfdReporte = new AfdReporteSBiz(dr);
					this._lista.Add(objAfdReporte);
				}
				this._tabla = dt;
            }
            catch
            {
                throw;
            }
		}
		#endregion
    }
}
