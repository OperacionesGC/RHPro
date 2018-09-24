using System;
using System.Configuration;
using System.Threading;
using System.Web.UI;
using Common;
using ServicesProxy;
using System.Data;
/*using ServicesProxy.rhdesa;*/

namespace Indicadores
{
    public partial class HomeIndicadores : UserControl
    {

	  //  private  int MRUsAvailables = int.Parse(ConfigurationManager.AppSettings["CantidadMRUsVisibles"]);
		
		
        public RHPro.Lenguaje ObjLenguaje;
		 
		protected void Page_Load(object sender, EventArgs e)
        {
            ObjLenguaje = new RHPro.Lenguaje();		    
			//TituloDescriptivo.Controls.Add(new LiteralControl(Traducir("Referencias más utilizadas del sistema")));
			//TituloDescriptivo.DataBind();	 
        }
		
		 
		  
        protected void Page_PreRender(object sender, EventArgs e)
        { 
            //LoadMRU();				      
        }   
				
        /// <summary>
        /// Busca y carga los MRU
        /// </summary>
        internal void LoadMRU()
        {            
		/*
		  if (string.IsNullOrEmpty(Utils.SessionUserName)){
                mruImage.Visible = true;			 
                mruCompleto.Visible = false;
            }
            else
            {
                mruImage.Visible = false;
				mruCompleto.Visible = true; 
				String sql = "";
				sql +=" SELECT menumstr.menuname, menumstr.action, menuraiz.menunombre raiz, menuraiz.menudir, menumstr.menuaccess ";
				sql +=" FROM mru ";
				sql +="  INNER JOIN menumstr ON menumstr.menumsnro = mru.menumsnro ";
				sql +="  INNER JOIN menuraiz ON menuraiz.menunro = mru.menuraiz ";
				sql +="  WHERE UPPER(mru.iduser) = Upper('"+Utils.SessionUserName+"') ";
				sql +="  AND menuraiz.menudir ='"+Utils.Session_ModuloActivo+"'  ";
				sql +="  ORDER BY mrufecha DESC, mruhora DESC ";
 ServicesProxy.rhdesa.Consultas cc = new ServicesProxy.rhdesa.Consultas();
				MRURepeater.DataSource = cc.get_DataSet(sql,Utils.SessionBaseID);//MRUServiceProxy.Find(Utils.SessionUserName, MRUsAvailables, Utils.SessionBaseID, Thread.CurrentThread.CurrentCulture.Name);
                MRURepeater.DataBind();
						
            } 
			*/
        }
		
   public string Traducir(string palabra)
   { 
     RHPro.Lenguaje ObjLenguaje = new RHPro.Lenguaje();  
     return ObjLenguaje.Label_Home(palabra);
   } 


  public string ArmarGrafLineal(String l_valor,String  l_indvalmin, String l_indvalmax,String  l_inddiv1, String l_inddiv2, 
            String l_inddiv3,String  l_inddiv4,String  l_inddiv5,String  l_indcolorrango1,String  l_indcolorrango2,String  l_indcolorrango3,
            String l_indcolorrango4, String l_indcolorrango5,String  l_indcolorrango6,String  l_link, String l_habDiv1,String  l_habDiv2,
            String l_habDiv3, String l_habDiv4, String l_habDiv5)
        {
            String l_htmlGraf ="";
            String strChartXML = "";

            //Genera el XML para el grafico
            strChartXML = "<Chart bgColor='f2f2f2' upperLimit='" + l_indvalmax + "' lowerLimit='" + l_indvalmin + "'";
            strChartXML +=  " showLimits='1' showValue='0' bgAlpha='0'";
            strChartXML +=  " BorderColor='B0BF9D' BorderThickness='1' baseFont='' baseFontSize='8' baseFontColor='' showColorNames='1'";
            strChartXML +=  " showTickValues='1' showTickMarks='1' majorTMNumber='5' numberSuffix='' numberScaleValue='1000,1000' numberScaleUnit='K,M'";
            strChartXML +=  " pointerSides='3' pointerBgColor='FF3333' pointerRadius='7' tickMarkDecimalPrecision='2'";
            strChartXML +=  " chartLeftMargin='20' chartRightMargin='20' decimalPrecision='2' thousandSeparator='' tickMarkGap='10'";
            if (l_link.Length != 0) 
            strChartXML +=  " clickURL='" + l_link + "'";

            strChartXML +=   ">";

            //Armo los rangos de colores segun lo configurado
            strChartXML +=   " <colorRange>";
            strChartXML += CalcularRangosColor(l_indvalmin, l_indvalmax, l_inddiv1, l_inddiv2, l_inddiv3, l_inddiv4, l_inddiv5, l_indcolorrango1, l_indcolorrango2, l_indcolorrango3, l_indcolorrango4, l_indcolorrango5, l_indcolorrango6, l_habDiv1, l_habDiv2, l_habDiv3, l_habDiv4, l_habDiv5);
            strChartXML +=   " </colorRange>";

            strChartXML +=   " <value>" + l_valor + "</value>";

            strChartXML +=   " </Chart>";


            //Armo el html que genera el grafico
            //"<object classid=""clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"" width=""250"" height=""110"">" +_
            l_htmlGraf = "<div align='center'>" ;
            l_htmlGraf =		 "<object type='application/x-shockwave-flash' width='250' height='110' data='/rhprox2/Shared/Charts/FI2_Linear.swf?chartWidth=250&chartheight=110'>"; 
            l_htmlGraf =            "<PARAM NAME='FlashVars' value='&dataXML=" + strChartXML + "'>" ;
            l_htmlGraf =			"<param name='movie' value='/rhprox2/Shared/Charts/FI2_Linear.swf?chartWidth=250&chartheight=110'>" ;
            l_htmlGraf =            "<param name='quality' value='high'>" ;
            l_htmlGraf =			"<PARAM NAME='wmode' VALUE='transparent'>" ;
            l_htmlGraf =         "</object>" ;
            l_htmlGraf =      "</div>";




            return l_htmlGraf;
        }

        
//=============================================================================================================
//Calcula los rangos de colores segun las divisiones configuradas
//=============================================================================================================
public string CalcularRangosColor(String l_indvalmin,String  l_indvalmax, String l_inddiv1,String  l_inddiv2,String  l_inddiv3,String  l_inddiv4,
    String l_inddiv5,String  l_indcolorrango1,String  l_indcolorrango2,String  l_indcolorrango3,String  l_indcolorrango4,String  l_indcolorrango5, 
    String l_indcolorrango6,String  l_habDiv1,String  l_habDiv2,String  l_habDiv3, String l_habDiv4, String l_habDiv5)
{
 
String l_sql = "";
String l_color1 = "";
String l_color2 = "";
String l_color3 = "";
String l_color4 = "";
String l_color5 = "";
String l_color6 = "";
String  l_salidaXML = "";


ServicesProxy.rhdesa.Consultas cc = new ServicesProxy.rhdesa.Consultas();
//Paso las credenciales al web service
cc.Credentials = System.Net.CredentialCache.DefaultCredentials;

 
	
	//Busco el color del rango
	l_color1 = "#FFFFFF";
	l_sql = "SELECT colorhtml ";
	l_sql  += " FROM colorhya ";
	l_sql  += " WHERE colornro = " + l_indcolorrango1;
	
    DataTable l_rsGraf = cc.get_DataTable(l_sql, Utils.SessionBaseID);
 
    if (l_rsGraf.Rows.Count>0)
        if (l_rsGraf.Rows[0]["colorhtml"]!=null)  
            l_color1 = Convert.ToString(l_rsGraf.Rows[0]["colorhtml"]);
	
	
	 
	//Primer Rango
	if (l_habDiv1 == "0")
		l_salidaXML += "<color minValue='" + l_indvalmin + "' maxValue='" + l_indvalmax + "' code='" + l_color1 + "'/>";
	else
    {l_salidaXML += "<color minValue='" + l_indvalmin + "' maxValue='" + l_inddiv1 + "' code='" + l_color1 + "'/>";
		
		//Segundo Rango
		l_color2 = "#FFFFFF";
		l_sql = "SELECT colorhtml ";
		l_sql  += " FROM colorhya ";
		l_sql  += " WHERE colornro = " + l_indcolorrango2;
		
        l_rsGraf = cc.get_DataTable(l_sql, Utils.SessionBaseID);
 
        if (l_rsGraf.Rows.Count>0)
        if (l_rsGraf.Rows[0]["colorhtml"]!=null)  
            l_color2 = Convert.ToString(l_rsGraf.Rows[0]["colorhtml"]);     
		
        if (l_habDiv2 == "0")
            l_salidaXML += "<color minValue='" + l_indvalmin + "' maxValue='" + l_indvalmax + "' code='" + l_color2 + "'/>";
        else
        { l_salidaXML += "<color minValue='" + l_indvalmin + "' maxValue='" + l_inddiv2 + "' code='" + l_color2 + "'/>";
			
			//Tercer Rango
			l_color3 = "#FFFFFF";
			l_sql = "SELECT colorhtml ";
			l_sql  +=  " FROM colorhya ";
			l_sql  +=  " WHERE colornro =" + l_indcolorrango3;

	        l_rsGraf = cc.get_DataTable(l_sql, Utils.SessionBaseID);
 
            if (l_rsGraf.Rows.Count>0)
              if (l_rsGraf.Rows[0]["colorhtml"]!=null)  
                l_color3 = Convert.ToString(l_rsGraf.Rows[0]["colorhtml"]);		

			 
			
            if (l_habDiv3 == "0")
                l_salidaXML += "<color minValue='" + l_indvalmin + "' maxValue='" + l_indvalmax + "' code='" + l_color3 + "'/>";
            else
            { l_salidaXML += "<color minValue='" + l_indvalmin + "' maxValue='" + l_inddiv3 + "' code='" + l_color3 + "'/>";
				
				//Cuarto Rango
				l_color4 = "#FFFFFF";
				l_sql = "SELECT colorhtml ";
				l_sql  += " FROM colorhya ";
				l_sql  += " WHERE colornro = " + l_indcolorrango4;
				
	            l_rsGraf = cc.get_DataTable(l_sql, Utils.SessionBaseID);
     
                if (l_rsGraf.Rows.Count>0)
                  if (l_rsGraf.Rows[0]["colorhtml"]!=null)  
                    l_color4 = Convert.ToString(l_rsGraf.Rows[0]["colorhtml"]);		
				
                if (l_habDiv4 == "0")
                    l_salidaXML += "<color minValue='" + l_indvalmin + "' maxValue='" + l_indvalmax + "' code='" + l_color4 + "'/>";
                else
                {   l_salidaXML += "<color minValue='" + l_indvalmin + "' maxValue='" + l_inddiv4 + "' code='" + l_color4 + "'/>";
    					
					//Quinto Rango
					l_color5 = "#FFFFFF";
					l_sql = "SELECT colorhtml ";
					l_sql  +=  " FROM colorhya ";
					l_sql  +=  " WHERE colornro = " + l_indcolorrango5;
					
                    l_rsGraf = cc.get_DataTable(l_sql, Utils.SessionBaseID);

                    if (l_rsGraf.Rows.Count>0)
                       if (l_rsGraf.Rows[0]["colorhtml"]!=null)  
                          l_color5 = Convert.ToString(l_rsGraf.Rows[0]["colorhtml"]);	
					
                    if (l_habDiv5 == "0")
                        l_salidaXML += "<color minValue='" + l_indvalmin + "' maxValue='" + l_indvalmax + "' code='" + l_color5 + "'/>";
                    else
                    { l_salidaXML += "<color minValue='" + l_indvalmin + "' maxValue='" + l_inddiv4 + "' code='" + l_color5 + "'/>";
					
						
						//Sexto Rango
						l_color6 = "#FFFFFF";
						l_sql = "SELECT colorhtml ";
						l_sql  +=  " FROM colorhya ";
						l_sql  +=  " WHERE colornro = " + l_indcolorrango6;
						
                        l_rsGraf = cc.get_DataTable(l_sql, Utils.SessionBaseID);
     
                        if (l_rsGraf.Rows.Count>0)
                           if (l_rsGraf.Rows[0]["colorhtml"]!=null)  
                             l_color6 = Convert.ToString(l_rsGraf.Rows[0]["colorhtml"]);	
						
						l_salidaXML +=  "<color minValue='" + l_inddiv5 + "' maxValue='"  + l_indvalmax + "' code='" + l_color6 + "'/>";
						
                        } //Rango 5
					
                    } //Rango 4
						
                } //Rango 3
			
            } //Rango 2
		
        } //Rango 1
	
  

        return l_salidaXML;
    }   
   
  		  
			
	}
}