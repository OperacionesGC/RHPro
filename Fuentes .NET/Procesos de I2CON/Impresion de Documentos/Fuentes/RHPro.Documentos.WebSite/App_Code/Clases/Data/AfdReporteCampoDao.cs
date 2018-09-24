using System;
using System.Data;
using System.Configuration;
using i2Con.Data.Connection;
using RHPro.ReportesAFD.Clases;

namespace RHPro.ReportesAFD.DataLayer.Data
{
	public class AfdReporteCampoDAO
	{		
		#region Atributos
    	private static string _sql;
		private static string _aux;
		#endregion
		
		#region Metodos Publicos
		public DataTable GetAfdReporteCampo(int? Nrocampo, int? Nrotag, string Campo, string Alias, string CriterioOrden)
		{
			try
			{	
        		_sql = "";
				_sql += " SELECT * ";
				_sql += " FROM afd_reporte_campo ";

				_aux = " WHERE ";
				if (Nrocampo != null)
				{
					_sql += _aux + " nrocampo = '" + Nrocampo.Value + "'";
					_aux = " AND ";
				}
				if (Nrotag != null)
				{
					_sql += _aux + " nrotag = '" + Nrotag.Value + "'";
					_aux = " AND ";
				}
				if (Campo != null)
				{
					_sql += _aux + " campo = '" + Campo + "'";
					_aux = " AND ";
				}
				if (Alias != null)
				{
					_sql += _aux + " alias = '" + Alias + "'";
					_aux = " AND ";
				}
				if (CriterioOrden != null)
				{
					_sql += " Order by " + CriterioOrden;
				}
				
				return I2Database.CreateDataSet(AppSession.RHProDBConnection, _sql).Tables[0];
			}    
			catch
			{
				throw;
			}
        }
        public long InsertarAfdReporteCampo(int Nrotag, string Campo, string Alias)
		{
			try
			{	
        		_sql = "";
				_sql = "SELECT MAX(nrocampo) FROM afd_reporte_campo";
				
				string lastId = I2Database.CreateDataSet(AppSession.RHProDBConnection, _sql).Tables[0].Rows[0][0].ToString();
                long newId = 0;
                if (lastId != "") newId = long.Parse(lastId) + 1;
				
				_sql = "";
				_sql += " INSERT INTO afd_reporte_campo(nrocampo, nrotag, campo, alias) VALUES (";
				_sql += " " + newId + " ";
				_sql += " ," + Nrotag + " ";
				_sql += " ,'" + Campo + "' ";
				_sql += " ,'" + Alias + "' ";
				_sql += ")";
				I2Database.Execute(AppSession.RHProDBConnection, _sql);
				return newId;
			}    
			catch
			{
				throw;
			}
		}
        public void EliminarAfdReporteCampo(int Nrocampo)
		{
			try
			{	
        		_sql = "";
				_sql += " DELETE FROM afd_reporte_campo WHERE ";
				_sql += " " + " nrocampo = " + Nrocampo + " ";
				I2Database.Execute(AppSession.RHProDBConnection, _sql);
			}    
			catch
			{
				throw;
			}
		}
        public void ActualizarAfdReporteCampo(int Nrocampo, int Nrotag, string Campo, string Alias)
		{
			try
			{	
        		_sql = "";
				_sql += " UPDATE afd_reporte_campo SET ";
				_sql += " " + " nrotag = " + Nrotag + " ";
				_sql += " ," + " campo = '" + Campo + "' ";
				_sql += " ," + " alias = '" + Alias + "' ";
				_sql += " WHERE ";
				_sql += " " + " nrocampo = " + Nrocampo + " ";
				I2Database.Execute(AppSession.RHProDBConnection, _sql);
			}    
			catch
			{
				throw;
			}
		}
		#endregion
	}
}
