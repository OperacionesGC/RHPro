using System;
using System.Collections.Generic;
using System.Data;
using RHPro.ReportesAFD.BussinesLayer.Base;
using RHPro.ReportesAFD.DataLayer.Data;
		
namespace RHPro.ReportesAFD.BussinesLayer.Biz
{
	[Serializable]
    public class AfdReporteCampoSBiz:AfdReporteCampoBase
    {
		#region Atributos
        private AfdReporteCampoDAO _afdReporteCampoDao = new AfdReporteCampoDAO();
		#endregion
		
		#region Metodos Privados
		private DataRow GetDRAfdReporteCampoByPK(int Nrocampo)
        {
            try
            {
				DataTable aux;
				aux = this._afdReporteCampoDao.GetAfdReporteCampo(Nrocampo, null, null, null, null);
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
				if (this._nrocampo == int.MinValue)
				{
					long nuevoId;
					nuevoId = this._afdReporteCampoDao.InsertarAfdReporteCampo(this._nrotag, this._campo, this._alias);
					this._nrocampo = (int) nuevoId;

				}
				else
				{
					this._afdReporteCampoDao.ActualizarAfdReporteCampo(this._nrocampo, this._nrotag, this._campo, this._alias);
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
                this._afdReporteCampoDao.EliminarAfdReporteCampo(this._nrocampo);
            }
            catch
            {
                throw;
            }
        }
		#endregion
		
		#region Constructores
		public AfdReporteCampoSBiz()
		{
		}
		public AfdReporteCampoSBiz(DataRow dr)
		{
			try
            {
				if (dr != null)
				{
					
					if (dr["nrocampo"] != DBNull.Value)
							this._nrocampo = (int)dr["nrocampo"];
					
					if (dr["nrotag"] != DBNull.Value)
							this._nrotag = (int)dr["nrotag"];
					
					if (dr["campo"] != DBNull.Value)
							this._campo = (string)dr["campo"];
					
					if (dr["alias"] != DBNull.Value)
							this._alias = (string)dr["alias"];
					
				}
            }
            catch
            {
                throw;
            }
		}
		public AfdReporteCampoSBiz(int Nrocampo)
		{
			try
            {
				DataRow dr;
				dr = this.GetDRAfdReporteCampoByPK(Nrocampo);

				if (dr != null)
				{
					
					if (dr["nrocampo"] != DBNull.Value)
							this._nrocampo = (int)dr["nrocampo"];
					
					if (dr["nrotag"] != DBNull.Value)
							this._nrotag = (int)dr["nrotag"];
					
					if (dr["campo"] != DBNull.Value)
							this._campo = (string)dr["campo"];
					
					if (dr["alias"] != DBNull.Value)
							this._alias = (string)dr["alias"];
					
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
