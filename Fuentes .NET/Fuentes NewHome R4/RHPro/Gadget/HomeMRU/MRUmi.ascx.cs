using System;
using System.Configuration;
using System.Threading;
using System.Web.UI;
using Common;
using ServicesProxy;

/*using ServicesProxy.rhdesa;*/

namespace HomeMRU
{
    public partial class MRUmi : UserControl
    {
	  private  int MRUsAvailables = int.Parse(ConfigurationManager.AppSettings["CantidadMRUsVisibles"]);
		public int posFila = 0;
		public String BGColor="#FFFFFF";
		 
        public RHPro.Lenguaje ObjLenguaje;
		public RHPro.ConsultaDatos c_datos;
		public int Nro_Gadget;
		 
		protected void Page_Load(object sender, EventArgs e)
        {           
		  c_datos = new RHPro.ConsultaDatos();		   
		  Asignar_NroGadget(3);
        }
		
 
		public string Armar_Link_MRU(string menuaccess,int menumsnro, String action , int menuraiz,int menunro, String menuname, int mrucant, String menudir)
		{
			string salida="";
			//if (c_datos.MRU_Habilitado(menumsnro))	 
			if (c_datos.Menu_Habilitado(menuaccess , menumsnro))	 
			{
				if (posFila%2==0) 
				  BGColor = "#FFFFFF";  
				else  
				  BGColor ="#fff";			    
				posFila++;
                 
				salida =" <div onclick=\""+Corregir(action,Convert.ToString(menumsnro),Convert.ToString(menuraiz),Convert.ToString(menunro),menudir)+"\" class='MRUGeneral_Link'  style='background-color:"+BGColor+"'  title='"+ Traducir((String)Eval("menuname")) +"' > ";
 				salida+= Common.Utils.Armar_Icono("img/Modulos/SVG/LINK.svg", "IconoModuloMRUModulo",""," border='0' ", "");
				//salida+= "    <span style='font-size:7pt; float:right; margin-right:3px; vertical-align:middle; display:inline'>("+mrucant+") </span>   " + Traducir(menuname)  ;
				salida+=  Traducir(menuname)  ;
				salida+=" </div>  ";
				
			}  
			
			return salida;
		}
		 
		public void Asignar_NroGadget(int nro)
	    {
	 	  Nro_Gadget = nro;
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
              //  mruImage.Visible = true;			 
                mruCompleto.Visible = false;
            }
            else
            {
               try{

						mruCompleto.Visible = true; 	
						String sql = "";
						sql +=" SELECT top(" + MRUsAvailables + ") mru.mrucant,menumstr.menumsnro,menumstr.menuraiz,menuraiz.menunro, menumstr.menuname, menumstr.action, menuraiz.menunombre raiz, menuraiz.menudir, menumstr.menuaccess ";
						sql +="  ,menuraiz.menudir FROM mru ";
						sql +="  INNER JOIN menumstr ON menumstr.menumsnro = mru.menumsnro ";
						sql +="  INNER JOIN menuraiz ON menuraiz.menunro = mru.menuraiz ";					
						sql +="  WHERE UPPER(mru.iduser) = Upper('"+Utils.SessionUserName+"') ";						
						sql +="  ORDER BY mru.mrucant DESC,menuname ASC ";										
						
						ServicesProxy.rhdesa.Consultas cc = new ServicesProxy.rhdesa.Consultas();				
						System.Data.DataSet ds = cc.get_DataSet(sql,Utils.SessionBaseID);
						if (ds.Tables[0].Rows.Count>0)
						{ MRURepeater.DataSource = ds;
						  MRURepeater.DataBind();
						}
						else {					
							ScriptManager.RegisterStartupScript(Page, GetType(), "ControlMRU","Ocultar_MRU_Vacio();", true);
							MRURepeater.Controls.Clear();
						}
				}
				catch(Exception ex){throw ex;}
						
            } 
        }
		
   public string Traducir(string palabra)
   { 
     RHPro.Lenguaje ObjLenguaje = new RHPro.Lenguaje();  
     return ObjLenguaje.Label_Home(palabra);
   }   
   
   
   public string Corregir(String action,string menumsnro, string MenuRaiz,string menunro, string menudir)
	{		
		String salida=action;
		
		if (action != "#")
		   salida = Utils.ArmarAction(action,menudir, Convert.ToString(menumsnro), Convert.ToString(MenuRaiz) , Convert.ToString(menunro) );
		  
   	    return  salida;
    }	
      
	/*  
	public string Corregir(String action,String menumsnro, String MenuRaiz,String menunro)
	{		
		String salida=action;
		
		if (action != "#")
		  salida = Utils.ArmarAction(action,"", Convert.ToString(menumsnro), Convert.ToString(MenuRaiz) , Convert.ToString(menunro) );
         
   	    return  salida;
    }		  
		*/	
	 		  
			
  }
}