using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using ServicesProxy.rhdesa;
using Common;
using System.Data;
using System.Collections;

namespace RHPro.Controls
{
    public partial class Contenedor_Favoritos : System.Web.UI.UserControl
    {
        public static String NombreModulo;
        public static Boolean EsPrimero;
        public RHPro.Lenguaje ObjLenguaje;
        public RHPro.ConsultaDatos c_datos;
      
        protected void Page_Load(object sender, EventArgs e)
        {
            ObjLenguaje = new Lenguaje();
            c_datos = new RHPro.ConsultaDatos();	
            EsPrimero = true;
            Imprimir_Favoritos();            
        }
        

        public void Refrescar(object sender, EventArgs e)
        {
            EsPrimero = true;
            Imprimir_Favoritos();
            //Habilito el armado de los submenues
            ScriptManager.RegisterStartupScript(this, typeof(Page), "Logo_InicializaMenu", "$(function() {  $('#main-menu').smartmenus({  subMenusSubOffsetX: 0, subMenusSubOffsetY: -1  }); });  ", true);
            ScriptManager.RegisterStartupScript(this, typeof(Page), "Logo_InicializaMenuTop", "$(function() {  $('#main-menuTop').smartmenus({  subMenusSubOffsetX: 0, subMenusSubOffsetY: -1, mainMenuSubOffsetX:0,mainMenuSubOffsetY:0,subMenusMinWidth:'60px', subMenusMaxWidth:'900px'  }); });  ", true);
            
            ScriptManager.RegisterStartupScript(this, typeof(Page), "InicializaMenu", "$(function() {  $('#main-menu').smartmenus({  subMenusSubOffsetX: 0, subMenusSubOffsetY: -1 , hideOnClick: false  }); });  ", true);
            ScriptManager.RegisterStartupScript(this, typeof(Page), "InicializaMenuTop", "$(function() {  $('#main-menuTopLoguin').smartmenus({  subMenusSubOffsetX: 0, subMenusSubOffsetY: -1, mainMenuSubOffsetX:0,mainMenuSubOffsetY:0,subMenusMinWidth:'60px', subMenusMaxWidth:'1060px', hideOnClick: true   }); });  ", true);
        }

        public void Imprimir_Favoritos()
        {

       
            if (Utils.IsUserLogin)
            {                
                NombreModulo = "";
                Consultas cc = new Consultas();
                //int AUX_menunro = 0;
                String sql = "";
                string Leng = Common.Utils.Lenguaje.Replace("-", ""); 

                if (cc.get_TipoBase(Utils.SessionBaseID) == "MSSQL")
                {
                    sql = "SELECT M.menuname,M.menudesabr,F.modulomenumsnro, F.favnro, F.favURL, F.favuser,F.favmodulo, F.favtitulo,F.anchoventana,F.altoventana,F.menumsnro,F.menunro,M.menuaccess ";
                    sql += " , (select top(1)  LE." + Leng + " from lenguaje_etiqueta LE where LE.etiqueta= (select menuname from menumstr MM where MM.menumsnro=F.menumsnro) COLLATE Modern_Spanish_CS_AS  and ( LE.modulo= (select menudir from menuraiz where menunro=F.menunro) Or LE.modulo is null) ORDER BY LE.modulo DESC ) TraduccionEtiqueta ";
                    sql += " , (select top(1)  LE.esAR from lenguaje_etiqueta LE where LE.etiqueta= (select menuname from menumstr MM where MM.menumsnro=F.menumsnro) COLLATE Modern_Spanish_CS_AS  and ( LE.modulo= (select menudir from menuraiz where menunro=F.menunro) Or LE.modulo is null) ORDER BY LE.modulo DESC ) TraduccionEtiquetaESAR ";
                    sql += " , (select menuname from menumstr MM where MM.menumsnro=F.menumsnro) Etiqueta ";
                    sql += " , (select top(1)  LE." + Leng + " from lenguaje_etiqueta LE where LE.etiqueta= M.menudesabr COLLATE Modern_Spanish_CS_AS  ) TraduccionEtiquetaModulo ";
                    sql += " , (select top(1)  LE.esAR from lenguaje_etiqueta LE where LE.etiqueta= M.menudesabr COLLATE Modern_Spanish_CS_AS  ) TraduccionEtiquetaModuloESAR ";
                    sql += " , (select menudir from menuraiz where menunro=F.menunro) MenuRaiz ";
                    sql += " , (select menudesc from menuraiz where menunro=F.menunro) MenuDesc ";
                    //sql += " , (select menudir from menuraiz where menunro=F.menunro) MenuRaiz ";

                    sql += " , (select menudesabr from menumstr MM where Upper(MM.parent)='RHPRO' and MM.menumsnro=F.modulomenumsnro) NombreMod";

                    sql += " FROM Home_Favoritos F ";
                    //sql += "  inner join menumstr M on M.menumsnro = F.modulomenumsnro ";
                    sql += "  inner join menumstr M on M.menumsnro = F.menumsnro ";
                    sql += " WHERE UPPER(favuser)=Upper('" + Utils.SessionUserName + "') ";
                    sql += " order by NombreMod, favtitulo";
                }
                else
                {
 


                    sql =" SELECT M.menuname,M.menudesabr,F.modulomenumsnro, F.favnro, F.favURL, F.favuser,F.favmodulo, ";
                    sql += " F.favtitulo,F.anchoventana,F.altoventana,F.menumsnro,F.menunro,M.menuaccess  ";
                    sql += "     , ( select  LE." + Leng + "  from lenguaje_etiqueta LE  where rownum =1 AND LE.etiqueta= (select menuname from menumstr MM  where rownum =1 AND MM.menumsnro=F.menumsnro)  and    ";
                    sql += "           ( LE.modulo= (select menudir from menuraiz where menunro=F.menunro) Or LE.modulo is null )  ";
                    sql += "     ) TraduccionEtiqueta  ";
                    sql += "     , (select    LE.esAR from lenguaje_etiqueta LE where rownum =1 AND LE.etiqueta= (select menuname from menumstr MM where rownum =1 AND MM.menumsnro=F.menumsnro)   and ( LE.modulo= (select menudir from menuraiz where rownum =1 AND menunro=F.menunro) Or LE.modulo is null)  ) TraduccionEtiquetaESAR  ";
                    sql += "     , (select menuname from menumstr MM where MM.menumsnro=F.menumsnro) Etiqueta  ";
                    sql += "     , (select    LE." + Leng + "  from lenguaje_etiqueta LE where rownum =1 AND LE.etiqueta= M.menudesabr  ) TraduccionEtiquetaModulo  ";
                    sql += "     , (select    LE.esAR from lenguaje_etiqueta LE where rownum =1 AND LE.etiqueta= M.menudesabr  ) TraduccionEtiquetaModuloESAR  ";
                    sql += "     , (select menudir from menuraiz where rownum =1 AND menunro=F.menunro) MenuRaiz  ";
                    sql += "     , (select menudesc from menuraiz where rownum =1 AND menunro=F.menunro) MenuDesc ";
                    sql += "     , (select menudesabr from menumstr MM where rownum =1 AND Upper(MM.parent)='RHPRO' and MM.menumsnro=F.modulomenumsnro) NombreMod ";
                    sql += "          FROM Home_Favoritos F ";
                    sql += " inner join menumstr M on M.menumsnro = F.menumsnro  ";
                    sql += " WHERE UPPER(F.favuser)=UPPER('" + Utils.SessionUserName + "')  ";
                    sql += "  order by NombreMod, favtitulo ";
                }

                
                //Paso las credenciales al web service
                //cc.Credentials = System.Net.CredentialCache.DefaultCredentials;
                
                DataTable dt = cc.get_DataTable(sql, Utils.SessionBaseID);
                
                String datos = "";
                String NM = "";
                String modulo = "";
                Contenedor_Fav.Controls.Clear();
                String Evento = "";
                String URL = "";
                String idFavo = "";
                String Titulo = "";
                String Traduccion = "";
                String TraduccionModulo = "";
                String MenuDesc = "";
                String MenuRaiz = "";
                

                int ModuloMenumsNroActivo=-1;
                int ModuloMenumsNro = -1;

                List<String> PerUsr;
                Usuarios Usr = new Usuarios();
                PerUsr = Usr.getPerfilesUsuario(Utils.SessionUserName);
                
                c_datos = new RHPro.ConsultaDatos();	

                //bool moduloActivo = (c_datos.get_ModulosHabilitados()).Contains(Convert.ToString(dr["menuname"]));
                bool moduloHabil = true;
                bool EntroAlmenosUnaVez = false;
                List<string> Lista_ModHabilidatos = c_datos.get_ModulosHabilitados();
                foreach (DataRow dr in dt.Rows)
                {                    
                    MenuDesc = Convert.ToString(dr["MenuDesc"]);
                    MenuRaiz = Convert.ToString(dr["MenuRaiz"]);

                    //Verifico si el modulo esta habilitado --------------------------------------------------------                    
                    moduloHabil = (Lista_ModHabilidatos).Contains(MenuDesc.ToUpper());
                    //----------------------------------------------------------------------------------------------

                    if (moduloHabil)
                    {
                        //Verifico si esta habilitado el menu
                       if (c_datos.Menu_Habilitado(Convert.ToString(dr["menuaccess"]), Convert.ToInt32(dr["menumsnro"])))                       
                        {

                            if (!DBNull.Value.Equals(dr["TraduccionEtiqueta"]))
                                Traduccion = Convert.ToString(dr["TraduccionEtiqueta"]);
                            else
                                if (!DBNull.Value.Equals(dr["TraduccionEtiquetaESAR"]))
                                    Traduccion = Convert.ToString(dr["TraduccionEtiquetaESAR"]);
                                else
                                    Traduccion = Convert.ToString(dr["Etiqueta"]);

                            if (!DBNull.Value.Equals(dr["TraduccionEtiquetaModulo"]))
                                TraduccionModulo = Convert.ToString(dr["TraduccionEtiquetaModulo"]);
                            else
                                if (!DBNull.Value.Equals(dr["TraduccionEtiquetaModuloESAR"]))
                                    TraduccionModulo = Convert.ToString(dr["TraduccionEtiquetaModuloESAR"]);
                                else
                                    TraduccionModulo = Convert.ToString(dr["menudesabr"]);

                             

                            //Limito la cantidad de caracteres a visualizar en el titulo del link
                            if (Convert.ToString(Traduccion).Length > 29)
                                Titulo = Convert.ToString(Traduccion).Substring(0, 29) + " ...";
                            else
                                Titulo = Convert.ToString(Traduccion);

                            idFavo = "Id_Fav_" + Convert.ToString(dr["favnro"]);
                            URL = ImprimirURL((String)dr["favURL"]);
                            Evento = "RHPROHome_Favorito_Iframe_Add.location='/rhprox2/shared/asp/home_add_favorito.asp?URL_FAVORITO=" + URL + "&ELIMINA_FAVORITO=-1&Id_Favorito=" + idFavo + "&favnro=" + Convert.ToString(dr["favnro"]) + "'";
                            NM = "";
                            //modulo = (String)dr["favmodulo"];
                            //modulo = TraduccionModulo;
                            modulo = ObjLenguaje.Label_Home(Convert.ToString(dr["NombreMod"]));
                            ModuloMenumsNro = Convert.ToInt32(dr["modulomenumsnro"]);                          

                            if (EsPrimero)
                            {
                                NM = "<DIV  class='SeccionFavorito' >";
                                if (modulo != "")
                                {
                                    //NM += "<DIV  class='EtiquetaFavoritoModulo' ><img src='img/modulos/SVG/" + MenuDesc + ".svg' class='IconoModulo_Favoritos'> <b>";                                   
                                    NM += "<DIV  class='EtiquetaFavoritoModulo' >" + Common.Utils.Armar_Icono("img/modulos/SVG/" + MenuDesc + ".svg", "IconoModuloGadget", "", "", "", "") + " <b>";                                                                       
                                    NM += modulo;
                                    NM += "</b></DIV>";
                                }
                                ModuloMenumsNroActivo = ModuloMenumsNro;
                                EsPrimero = false;
                                EntroAlmenosUnaVez = true;
                            }
                            else
                            {
                                //if (NombreModulo != modulo)
                                if (ModuloMenumsNroActivo != ModuloMenumsNro)
                                {
                                    NM = "</DIV><DIV  class='SeccionFavorito' >";
                                    if (modulo != "")
                                    {
                                        //NM += "<DIV  class='EtiquetaFavoritoModulo'  ><img src='img/modulos/SVG/" + MenuDesc + ".svg' class='IconoModulo_Favoritos'> <b>";
                                        NM += "<DIV  class='EtiquetaFavoritoModulo' >" + Common.Utils.Armar_Icono("img/modulos/SVG/" + MenuDesc + ".svg", "IconoModuloGadget", "", "", "", "") + " <b>";   
                                        //NM += "<img src='img/Modulos/SVG/" +  Convert.ToString(dr["MenuRaiz"]) + ".svg' class='IconoModulo'>";
                                        NM += modulo;
                                        NM += "</b></DIV>";
                                    }
                                    //NombreModulo = modulo;
                                    ModuloMenumsNroActivo = ModuloMenumsNro;
                                    EntroAlmenosUnaVez = true;
                                }
                            }

                            datos = " <span class='EtiquetaFavorito'  id='" + idFavo + "' title='" + Convert.ToString(dr["favtitulo"]) + "' onclick=\"abrirVentana('" + URL + "','','" + (String)dr["altoventana"] + "','" + (String)dr["anchoventana"] + "' )\">";
                            datos += Utils.Armar_Icono("img/Modulos/SVG/FAVORITO.svg", "IconoMRU", "", " border='0' style='cursor: pointer;'", "") + Titulo;
                            datos += "   <span class='cerrarVentana' style='margin-right:2px;' title='" + ObjLenguaje.Label_Home("Desanclar") + "' onclick=\"Javascript:" + Evento + "\"> X </span>";
                            datos += " </span>  ";
                           

                            //Agrego el link al contenedor de favoritos
                            Contenedor_Fav.Controls.Add(new LiteralControl(NM + datos));
                        }
                    }//Fin de cointrol modulo activo

                }
                if (EntroAlmenosUnaVez)
                    Contenedor_Fav.Controls.Add(new LiteralControl("</DIV>"));
            }

            
            
        }

        public String Imprimir_Modulo(String modulo)
        {
            String Salida = "";
            if (NombreModulo == "")
            {
                Salida = "<DIV  class='EtiquetaFavorito' ><b>";
                Salida += modulo;
                Salida += "</b></DIV>";
                NombreModulo = modulo;
            }
            else
            {
                if (NombreModulo != modulo)
                {
                    Salida = "<DIV  class='EtiquetaIdioma'  ><b>";
                    Salida += modulo;
                    Salida += "</b></DIV>";
                    NombreModulo = modulo;
                }
            }

            
            return Salida;

        }

        public String ImprimirURL(String URL)
        {
            String Salida = "";
            Salida = URL.Replace("_M=&","");
            Salida = Salida.Replace("_M=", "");
            return Salida;

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


Consultas cc = new Consultas();
//Paso las credenciales al web service
//cc.Credentials = System.Net.CredentialCache.DefaultCredentials;

 
	
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