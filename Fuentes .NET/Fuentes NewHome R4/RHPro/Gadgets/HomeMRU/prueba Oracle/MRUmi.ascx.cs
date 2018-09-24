using System;
using System.Configuration;
using System.Threading;
using System.Web.UI;
using Common;
using ServicesProxy;

namespace HomeMRU
{
    public partial class MRUmi : UserControl
    {

	    private  int MRUsAvailables = int.Parse(ConfigurationManager.AppSettings["CantidadMRUsVisibles"]);
		int posFila = 0;
		 
        public RHPro.Lenguaje ObjLenguaje;
		protected void Page_Load(object sender, EventArgs e)
        {
            ObjLenguaje = new RHPro.Lenguaje();
	 
        }
		
		public string background(){
			 string color = "#F7F7F7";
			 if (posFila%2==0) {
  			   color = "#FFFFFF";
			 }
 
			   posFila++;
			   return color;
		}

        protected void Page_PreRender(object sender, EventArgs e)
        {
            LoadMRU();				 
        }   
				
        /// <summary>
        /// Busca y carga los MRU
        /// </summary>
        internal void LoadMRU()
        {            
		  if (string.IsNullOrEmpty(Utils.SessionUserName)){
                mruImage.Visible = true;			 
                mruCompleto.Visible = false;
            }
            else
            {
                mruImage.Visible = false;
				mruCompleto.Visible = true;
				//SqlDataSourceMRU.SelectCommand = "SELECT top(" + MRUsAvailables + ") menumstr.menuname, menumstr.action, menuraiz.menunombre Root, menuraiz.menudir, menumstr.menuaccess FROM mru INNER JOIN menumstr ON menumstr.menumsnro = mru.menumsnro  INNER JOIN menuraiz ON menuraiz.menunro = mru.menuraiz  WHERE UPPER(mru.iduser) = 'rhpror3'  ORDER BY mrufecha DESC, mruhora DESC";                				 
//                MRURepeater.DataSourceID = "SqlDataSourceMRU";	
				
				
			/*Prueba de Oracle*/	
			System.Data.OleDb.OleDbConnection cn2 = new System.Data.OleDb.OleDbConnection();
            cn2.ConnectionString = "Provider=OraOLEDB.Oracle.1;Persist Security Info=True;Password=SOSR3;User ID=SOSR3;Data Source=rhoracle/SOSR3;";
            cn2.Open();
            string sqlSS = "ALTER SESSION SET CURRENT_SCHEMA = SOSR3";
            System.Data.OleDb.OleDbCommand cmd = new System.Data.OleDb.OleDbCommand(sqlSS, cn2);
            cmd.ExecuteNonQuery();
			
			
			sqlSS = "insert into sacar";
            cmd = new System.Data.OleDb.OleDbCommand(sqlSS, cn2);
            cmd.ExecuteNonQuery();
		    
            
            
				SqlDataSource1.ConnectionString = "Password=ess;Persist Security Info=True;User ID=ess;Initial Catalog=Base_0_R3_ARG;Data Source=RHDESA";
			//	SqlDataSource1.ConnectionString = "Persist Security Info=True;Password=ESS;User ID=ESS;Data Source=RHORACLE/SOSR3;";
				XX1.DataSource = SqlDataSource1;
				XX1.DataBind();
				
              /*---------------------------*/				
			  
			  
				MRURepeater.DataSource = MRUServiceProxy.Find(Utils.SessionUserName, MRUsAvailables, Utils.SessionBaseID, Thread.CurrentThread.CurrentCulture.Name);
				
                MRURepeater.DataBind();
						
            } 
        }
		
   public string Traducir(string palabra)
   { 
     RHPro.Lenguaje ObjLenguaje = new RHPro.Lenguaje();  
     return ObjLenguaje.Label_Home(palabra);
   }   
   
   public string Corregir(String action)
	{
 		string salida =  action.Replace("abrirVentana('","X_abrirVentana('"); 	
   	    return  salida;
    }		  
			
	}
}