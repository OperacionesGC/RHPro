using System;
using System.Collections.Generic;
using System.Data;
using RHPro.ReportesAFD.BussinesLayer.Base;
using RHPro.ReportesAFD.DataLayer.Data;
		
namespace RHPro.ReportesAFD.BussinesLayer.Biz
{
	[Serializable]
    public class AfdReporteCampoBiz
    {
		#region Atributos
        private AfdReporteCampoDAO _afdReporteCampoDao = new AfdReporteCampoDAO();
		private IList< AfdReporteCampoSBiz > _lista;
		private DataTable _tabla;
		#endregion
		
		#region Propiedades
		public IList< AfdReporteCampoSBiz > Lista
		{
			get {return this._lista; }
		}
		public DataTable Tabla
		{
			get {return this._tabla; }
		}
		#endregion
		
		#region Metodos Privados
		private DataTable GetDTAfdReporteCampoByFilters(int? Nrotag, string Campo, string Alias, string CriterioOrden)
        {
            try
            {
                return this._afdReporteCampoDao.GetAfdReporteCampo(null, Nrotag, Campo, Alias, CriterioOrden);
            }
            catch
            {
                throw;
            }
        }
		#endregion
		
		#region Metodos Publicos
        public List<string> ObtenerCamposAlias()
        {
            List<string> camposAlias = new List<string>();
            foreach (AfdReporteCampoSBiz c in _lista)
                camposAlias.Add(c.Campo + " [" + c.Alias + "]");
            return camposAlias;
        }
		#endregion
		
		#region Constructores
		public AfdReporteCampoBiz()
		{
			try
            {
				DataTable dt;
                dt = this.GetDTAfdReporteCampoByFilters(null, null, null, null);
				
				this._lista = new List< AfdReporteCampoSBiz >();
				foreach (DataRow dr in dt.Rows)
				{				
					AfdReporteCampoSBiz objAfdReporteCampo= new AfdReporteCampoSBiz(dr);
					this._lista.Add(objAfdReporteCampo);
				}
				this._tabla = dt;
            }
            catch
            {
                throw;
            }
		}
		public AfdReporteCampoBiz(int? Nrotag, string Campo, string Alias, string CriterioOrden)
		{
			try
            {
				DataTable dt;
                dt = this.GetDTAfdReporteCampoByFilters(Nrotag, Campo, Alias, CriterioOrden);
				
				this._lista = new List< AfdReporteCampoSBiz >();
				foreach (DataRow dr in dt.Rows)
				{				
					AfdReporteCampoSBiz objAfdReporteCampo = new AfdReporteCampoSBiz(dr);
					this._lista.Add(objAfdReporteCampo);
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
