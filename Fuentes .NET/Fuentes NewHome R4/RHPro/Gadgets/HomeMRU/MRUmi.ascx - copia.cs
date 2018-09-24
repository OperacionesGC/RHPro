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