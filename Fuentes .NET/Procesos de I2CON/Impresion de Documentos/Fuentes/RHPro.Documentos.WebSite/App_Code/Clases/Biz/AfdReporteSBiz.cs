using System;
using System.Collections.Generic;
using System.Data;
using RHPro.ReportesAFD.BussinesLayer.Base;
using RHPro.ReportesAFD.DataLayer.Data;

namespace RHPro.ReportesAFD.BussinesLayer.Biz
{
	[Serializable]
    public class AfdReporteSBiz:AfdReporteBase
    {
		#region Atributos
        private AfdReporteDAO _afdReporteDao = new AfdReporteDAO();
		#endregion
		
		#region Metodos Privados
		private DataRow GetDRAfdReporteByPK(int Nrotag)
        {
            try
            {
				DataTable aux;
				aux = this._afdReporteDao.GetAfdReporte(Nrotag, null, null, null, null, null);
                if (aux.Rows.Count != 0)
					return aux.Rows[0]; 
				else
					return null;
            }
            catch
            {
                throw;
            }
        }
		#endregion
		
		#region Metodos Publicos
        public void Save()
        {
            try
            {	
				if (this._nrotag == int.MinValue)
				{
					long nuevoId;
					nuevoId = this._afdReporteDao.InsertarAfdReporte(this._repnro, this._nomtag, this._tablequery, this._cantreg);
					this._nrotag = (int) nuevoId;

				}
				else
				{
					this._afdReporteDao.ActualizarAfdReporte(this._nrotag, this._repnro, this._nomtag, this._tablequery, this._cantreg);
				}
            }
            catch
            {
                throw;
            }
        }
		public void Delete()
        {
            try
            {
                this._afdReporteDao.EliminarAfdReporte(this._nrotag);
            }
            catch
            {
                throw;
            }
        }
		#endregion
		
		#region Constructores
		public AfdReporteSBiz()
		{
		}
		public AfdReporteSBiz(DataRow dr)
		{
			try
            {
				if (dr != null)
				{
					
					if (dr["nrotag"] != DBNull.Value)
							this._nrotag = (int)dr["nrotag"];
					
					if (dr["repnro"] != DBNull.Value)
							this._repnro = (int)dr["repnro"];
					
					if (dr["nomtag"] != DBNull.Value)
							this._nomtag = (string)dr["nomtag"];
					
					if (dr["tablequery"] != DBNull.Value)
							this._tablequery = (string)dr["tablequery"];
					
					if (dr["cantreg"] != DBNull.Value)
							this._cantreg = (int)dr["cantreg"];
					
				}
            }
            catch
            {
                throw;
            }
		}
		public AfdReporteSBiz(int Nrotag)
		{
			try
            {
				DataRow dr;
				dr = this.GetDRAfdReporteByPK(Nrotag);

				if (dr != null)
				{
					
					if (dr["nrotag"] != DBNull.Value)
							this._nrotag = (int)dr["nrotag"];
					
					if (dr["repnro"] != DBNull.Value)
							this._repnro = (int)dr["repnro"];
					
					if (dr["nomtag"] != DBNull.Value)
							this._nomtag = (string)dr["nomtag"];
					
					if (dr["tablequery"] != DBNull.Value)
							this._tablequery = (string)dr["tablequery"];
					
					if (dr["cantreg"] != DBNull.Value)
							this._cantreg = (int)dr["cantreg"];
					
				}
            }
            catch
            {
                throw;
            }
		}
		#endregion
    }
}
