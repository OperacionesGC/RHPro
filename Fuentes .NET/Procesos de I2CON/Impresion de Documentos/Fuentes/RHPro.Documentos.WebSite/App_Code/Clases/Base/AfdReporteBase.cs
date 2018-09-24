using System;
using System.Collections;
using System.Data;

namespace RHPro.ReportesAFD.BussinesLayer.Base
{
	public class AfdReporteBase
	{
		#region Atributos
		protected int _nrotag;
		protected int _repnro;
		protected string _nomtag;
		protected string _tablequery;
		protected int _cantreg;
		#endregion
		
		#region Propiedades
		public int Nrotag 
		{
			get {return this._nrotag; }
		}
		public int Repnro 
		{
			set { this._repnro = value; }
			get {return this._repnro; }
		}
		public string Nomtag 
		{
			set { this._nomtag = value; }
			get {return this._nomtag; }
		}
		public string Tablequery 
		{
			set { this._tablequery = value; }
			get {return this._tablequery; }
		}
		public int Cantreg 
		{
			set { this._cantreg = value; }
			get {return this._cantreg; }
		}
		#endregion
		
		#region Constructores
		protected AfdReporteBase()
		{
			this._nrotag = int.MinValue;
			
			
			
							
		}	
		#endregion
	}
}
