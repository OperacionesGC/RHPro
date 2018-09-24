using System;
using System.Data;
using System.Configuration;
using i2Con.Data.Connection;
using RHPro.ReportesAFD.Clases;

namespace RHPro.ReportesAFD.DataLayer.Data
{
	public class AfdReporteDAO
	{		
		#region Atributos
    	private static string _sql;
		private static string _aux;
		#endregion
		
		#region Metodos Publicos
		public DataTable GetAfdReporte(int? Nrotag, int? Repnro, string Nomtag, string Tablequery, int? Cantreg, string CriterioOrden)
		{
			try
			{	
        		_sql = "";
				_sql += " SELECT * ";
				_sql += " FROM afd_reporte ";

				_aux = " WHERE ";
				if (Nrotag != null)
				{
					_sql += _aux + " nrotag = '" + Nrotag.Value + "'";
					_aux = " AND ";
				}
				if (Repnro != null)
				{
					_sql += _aux + " repnro = '" + Repnro.Value + "'";
					_aux = " AND ";
				}
				if (Nomtag != null)
				{
					_sql += _aux + " nomtag = '" + Nomtag + "'";
					_aux = " AND ";
				}
				if (Tablequery != null)
				{
					_sql += _aux + " tablequery = '" + Tablequery + "'";
					_aux = " AND ";
				}
				if (Cantreg != null)
				{
					_sql += _aux + " cantreg = '" + Cantreg.Value + "'";
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
        public long InsertarAfdReporte(int Repnro, string Nomtag, string Tablequery, int Cantreg)
		{
			try
			{	
        		_sql = "";
				_sql = "SELECT MAX(nrotag) FROM afd_reporte";
				
				string lastId = I2Database.CreateDataSet(AppSession.RHProDBConnection, _sql).Tables[0].Rows[0][0].ToString();
                long newId = 0;
                if (lastId != "") newId = long.Parse(lastId) + 1;
				
				_sql = "";
				_sql += " INSERT INTO afd_reporte(nrotag, repnro, nomtag, tablequery, cantreg) VALUES (";
				_sql += " " + newId + " ";
				_sql += " ," + Repnro + " ";
				_sql += " ,'" + Nomtag + "' ";
				_sql += " ,'" + Tablequery + "' ";
				_sql += " ," + Cantreg + " ";
				_sql += ")";
				I2Database.Execute(AppSession.RHProDBConnection, _sql);
				return newId;
			}    
			catch
			{
				throw;
			}
		}
        public void EliminarAfdReporte(int Nrotag)
		{
			try
			{	
        		_sql = "";
				_sql += " DELETE FROM afd_reporte WHERE ";
				_sql += " " + " nrotag = " + Nrotag + " ";
				I2Database.Execute(AppSession.RHProDBConnection, _sql);
			}    
			catch
			{
				throw;
			}
		}
        public void ActualizarAfdReporte(int Nrotag, int Repnro, string Nomtag, string Tablequery, int Cantreg)
		{
			try
			{	
        		_sql = "";
				_sql += " UPDATE afd_reporte SET ";
				_sql += " " + " repnro = " + Repnro + " ";
				_sql += " ," + " nomtag = '" + Nomtag + "' ";
				_sql += " ," + " tablequery = '" + Tablequery + "' ";
				_sql += " ," + " cantreg = " + Cantreg + " ";
				_sql += " WHERE ";
				_sql += " " + " nrotag = " + Nrotag + " ";
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
