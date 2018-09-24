using System;
using System.Collections;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Xml.Linq;
using ServicesProxy;
using Common;

namespace HomeGadget_Corporativa
{	
    public partial class Gadget_Corporativa : UserControl
    {
		/*
		 * Modificaciones:  Se ordenan los banners por id
		 */ 
		
    	public RHPro.Lenguaje ObjLenguaje;
        public static Hashtable confper;
        protected void Page_Load(object sender, EventArgs e)
        {      
 	
            int contador=0;
            confper = Imagenes_Por_Pais_Activa();
		    //Levanto inicialmente las imagenes del pais por defecto
            if (confper != null)
            {
                if (Convert.ToInt32(confper["confactivo"]) == -1)
                {
                    if (Convert.ToInt32(confper["confint"]) == -1)//Si esta en -1 debe mostrar primero las imagenes por pais y luego los banners
                    {
                        contador = Cargar_Banners_X_Pais(contador);
                        contador = Imprimir_Banners(contador);
                    }
                    else //Muestra primero los banners y luego las imagenes por pais
                    {
                        contador = Imprimir_Banners(contador);
                        contador = Cargar_Banners_X_Pais(contador);
                    }
                }
                else//Solo muestra los banners
                {
                    Imprimir_Banners(contador);
                }
            }
            else//Solo muestra los banners
            {
                Imprimir_Banners(contador);
            }

 
		 
		}

        protected int Imprimir_Banners(int contador)
        {
         
            String sql = "";
            String SalidaBanners = "";
            sql += " SELECT  DISTINCT hbanimage,hbannro ";
            sql += " FROM home_banner ";
            sql += "  where hbanactivo=-1 and rhpro=-1 ";
			sql += "  ORDER BY hbannro  ";
            //sql += "  ORDER BY hbanimage desc ";

            //Luego levanto las imagenes banners configuradas para el home
            ServicesProxy.rhdesa.Consultas cc = new ServicesProxy.rhdesa.Consultas();
            System.Data.DataSet ds = cc.get_DataSet(sql, Utils.SessionBaseID);
            if (ds.Tables[0].Rows.Count > 0)
            {
                foreach (DataRow dr in ds.Tables[0].Rows)
                {
                    SalidaBanners += " <li id='Imagen_Corporativa_Nro_" + contador + "'><img src='Images/Banner/" + dr["hbanimage"] + "'   /></li> ";
                    contador++;
                }
                BannersCorp.Controls.Add(new LiteralControl(SalidaBanners));
            }

            return contador;
        }
		
		
		protected Hashtable Imagenes_Por_Pais_Activa()
		{					   
		   String sql = "";
           Hashtable resultado = null;

           //if (Convert.ToString(Session["RHPRO_Imagenes_Por_Pais_Activa"]) == "")
           //{
           sql = " select confactivo,confint from confper where confnro = 20 ";
               ServicesProxy.rhdesa.Consultas cc = new ServicesProxy.rhdesa.Consultas();
               System.Data.DataSet ds = cc.get_DataSet(sql, Utils.SessionBaseID);

               if (ds.Tables[0].Rows.Count > 0)
               {
                   resultado = new Hashtable();

                   if (!resultado.Contains("confactivo")) resultado.Add("confactivo", Convert.ToInt32(ds.Tables[0].Rows[0]["confactivo"]));
                   if (!resultado.Contains("confint")) resultado.Add("confint", Convert.ToInt32(ds.Tables[0].Rows[0]["confint"]));
               }

           //    Session["RHPRO_Imagenes_Por_Pais_Activa"] = Convert.ToInt32(resultado["confactivo"]);
           //    Session["RHPRO_Imagenes_Por_Pais_Primera"] = Convert.ToInt32(resultado["confint"]);
           //}
           //else
           //{
           //    resultado.Add("confactivo", Convert.ToInt32(Session["RHPRO_Imagenes_Por_Pais_Activa"]));
           //    resultado.Add("confint", Convert.ToInt32(Session["RHPRO_Imagenes_Por_Pais_Primera"]));               
           //}


           return resultado;
						
		}
		
		//Metodo que iserta las imagenes por pais al carrousel
		 protected int Cargar_Banners_X_Pais(int contador)
        {
            String sql = "";
            String SalidaBanners="";          
            String Lenguaje_Pais = "";
			//int contador = 0; 
			//Busco la nomenclatura del idioma del pais por defecto 
			sql +=" select L.lencod from pais P ";
			sql +=" inner join lenguaje L on L.paisnro = P.paisnro  ";
			sql +="  where P.paisdef=-1 ";

            ServicesProxy.rhdesa.Consultas cc = new ServicesProxy.rhdesa.Consultas();
            System.Data.DataSet ds = cc.get_DataSet(sql, Utils.SessionBaseID);

            if (ds.Tables[0].Rows.Count > 0)
            {
				//Recupero la nomenclatura del idioma del pais por defecto
                Lenguaje_Pais = Convert.ToString(ds.Tables[0].Rows[0]["lencod"]);              						 
			    
				//Armo la url de las imagenes por pais
				String URL_Log = "./Images/Banner/"+Lenguaje_Pais;
				URL_Log = Server.MapPath(URL_Log);								 
				URL_Log = URL_Log.Replace("/","\\");
                string[] files;
				String NombreImagen="";
                try
                { 
				   //Recupero todas las imagenes del directorio de las imagenes del pais
				   files = System.IO.Directory.GetFiles(URL_Log);
				 
				   foreach (String imagen in files)                   
				   {
					NombreImagen = System.IO.Path.GetFileName(imagen);
					//Restringir solo imagenes
					if ( (NombreImagen.Contains(".png")) ||  (NombreImagen.Contains(".jpg")) ||  (NombreImagen.Contains(".gif"))||  (NombreImagen.Contains(".jpeg")) )
					{
					    SalidaBanners +=" <li id='Imagen_Corporativa_Nro_"+contador+"'><img src='Images/Banner/"+Lenguaje_Pais+"/"+NombreImagen+"'  /></li> ";
						contador++;	
				    }					 
				   }				     
                }
                catch(Exception ex)
                {                    
					  Response.Write("-"+ex.Message);
                }              

                BannersCorp.Controls.Add(new LiteralControl(SalidaBanners));
            }	

			return contador;
        }
		
    }
}