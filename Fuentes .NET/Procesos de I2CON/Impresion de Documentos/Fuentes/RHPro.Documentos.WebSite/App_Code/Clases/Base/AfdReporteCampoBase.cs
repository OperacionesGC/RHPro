using System;
using System.Collections;
using System.Data;

namespace RHPro.ReportesAFD.BussinesLayer.Base
{
	public class AfdReporteCampoBase
	{
		#region Atributos
		protected int _nrocampo;
		protected int _nrotag;
		protected string _campo;
		protected string _alias;
		#endregion
		
		#region Propiedades
		public int Nrocampo 
		{
			get {return this._nrocampo; }
		}
		public int Nrotag 
		{
			set { this._nrotag = value; }
			get {return this._nrotag; }
		}
		public string Campo 
		{
			set { this._campo = value; }
			get {return this._campo; }
		}
		public string Alias 
		{
			set { this._alias = value; }
			get {return this._alias; }
		}
		#endregion
		
		#region Constructores
		protected AfdReporteCampoBase()
		{
			this._nrocampo = int.MinValue;
			
			
							
		}	
		#endregion
	}
}
