<%@ Page Title="" Language="C#" MasterPageFile="~/MasterPage.Master" AutoEventWireup="true" 
CodeBehind="Default.aspx.cs" Inherits="RHPro.Default" EnableViewState="true"  %>
 
 
<asp:Content ID="Content2" ContentPlaceHolderID="content" runat="server" > 

 

<head>
 

<!-- Hoja de Estilos construida desde un control -->    
    <cc:estilosnewhome id="EstilosNewHomes" runat="server"    >        
    </cc:estilosnewhome>
<!-- ------------------------------------------- -->    

<script type="text/javascript" src="/rhprox2/shared/js/fn_windowsmi.asp"></script>


       
<script type="text/javascript" >
  
  function Centrar_Globo(nombre) {

      var obj = document.getElementById(nombre);
      var Encab = document.getElementById("Encabezado");    
      var LeftObj;    
      LeftObj = (Encab.offsetWidth / 2) - (obj.offsetWidth / 2);
      obj.style.left = LeftObj + "px";      
    }

  

 
    function Abrir_Globo(nombre) {
         document.getElementById(nombre).style.visibility = "visible";      
         if (nombre == "Globo_Loguin") {
             Cerrar_Globo('Globo_Idiomas');
         } else {
                Cerrar_Globo("Globo_Loguin");
                Centrar_Globo('Globo_Idiomas');
                var fondo = document.getElementById("PopUp_FondoTransparenteLeng");
                fondo.style.display = "";          
         }
 }
 
 function Cerrar_Globo(nombre){
     document.getElementById(nombre).style.visibility = "hidden";
     /*****/
     var fondo = document.getElementById("PopUp_FondoTransparenteLeng");
     fondo.style.display = "none";
     /*****/
     document.cookie;
}

function AnchoPantalla() {  
   var width=0;
  if (self.screen)   // for NN4 and IE4
    width = screen.width;
  else
     if (self.java) {   // for NN3 with enabled Java
         var jkit = java.awt.Toolkit.getDefaultToolkit();
         var scrsize = jkit.getScreenSize();
         width = scrsize.width;
     } 
 return width;	 
}

function AltoPantalla() {  
  var height=0;  
  if (self.screen)   // for NN4 and IE4
    height = screen.height;
  else
     if (self.java) {   // for NN3 with enabled Java
         var jkit = java.awt.Toolkit.getDefaultToolkit();
         var scrsize = jkit.getScreenSize();
         height = scrsize.height;
     } 
 return height;	 
} 

function MostrarError(e){
document.write("<span  onclick=\"this.style.visibility = 'hidden'\" style='visibility:hidden; float:left;cursor:pointer; border:font-family:Arial; font-size:9pt; color:#333;border:4px #333333 solid; position: absolute; left:30px; top:30px; padding:6px; background-color:#FC9'><img src='img/error.png' align='absmiddle'> ERROR: <span Id='TextoError'></span> </span>"); 
//document.getElementById('ContenedorError').style.visibility = 'visible';
//document.getElementById('TextoError').innerHTML = e;
}

function RefrescarPagina(direccion,segundos){
setTimeout("location.href='"+direccion+"'",segundos*1000);
}

 
 /*
function Cerrar_Aplicacion(){
// Firefox || IE
       var e = window.event;
      // e = e || window.event;
       var y = e.pageY || e.clientY;

       if(y < 0)  //Verifica cuando se cierra la ventana        
       { 
        window.open("EndSession.aspx"); 
       }
       //else alert("Window refreshed");
}
 */

        
</script>




 
<!-- ---------------------------------------------------------------------- -->
<!-- SmartMenus jQuery plugin -->
 

 
<script type='text/javascript' >
    $(function() {
        $('#main-menu').smartmenus({
            subMenusSubOffsetX: 0,
            subMenusSubOffsetY: -1
        });
    });   
</script>  
<link href='css/sm-core-css.css' rel='stylesheet' type='text/css' /> 
<link href='css/sm-clean.css' rel='stylesheet' type='text/css' />
<!-- ---------------------------------------------------------------------- -->

<!-- #main-menu config - instance specific stuff not covered in the theme -->
<style type="text/css">
	#main-menu {
		position:relative;
		z-index:9999;
		width:auto;
	}
	#main-menu ul {
		width:12em; /* fixed width only please - you can use the "subMenusMinWidth"/"subMenusMaxWidth" script options to override this if you like */
	}
</style>


<style type="text/css">
	#main-menu {
		position:relative;
		z-index:9999;
		width:auto;
	}
	#main-menu ul {
		width:12em; /* fixed width only please - you can use the "subMenusMinWidth"/"subMenusMaxWidth" script options to override this if you like */
	}
</style>


 
<%     
    /*************************************************************************************/
    /*    ARMO EL TAG BODY                                                               */
    /*************************************************************************************/
    string ImpBody;
    ImpBody = " <body  leftmargin='0' topmargin='0' rightmargin='0' bottommargin='0' ";       
   /* 
    if (Common.Utils.IsUserLogin)
    {//Aqui controlo si viene logeado desde el MetaHome
        if ((Common.Utils.SessionNroTempLogin != null) && ((String)Common.Utils.SessionNroTempLogin != ""))
            ImpBody += "  onunload='Cerrar_Aplicacion()'   ";        
    }   
    */ 
    /*Aqui controlo si el sitio se va a refrescar para evitar la perdida de objetos en memoria. (Ver. Settings.config)*/
    if ((ConfigurationManager.AppSettings["TiempoRefrescoActivo"] != null) && (ConfigurationManager.AppSettings["TiempoRefresco"] != null))
    {
        if (bool.Parse(ConfigurationManager.AppSettings["TiempoRefrescoActivo"]))
        { 
            ImpBody += "onload=\"RefrescarPagina('Default.aspx'," + int.Parse(ConfigurationManager.AppSettings["TiempoRefresco"].ToString()) + ");\"";
        }
    }
    ImpBody += " >";
    /*************************************************************************************/
%>

<!--body  leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0" onmousemove='if (Mover()) {}'  -->
 
 
 
 <% Response.Write(ImpBody); %>
 
  
          
 
 <% 
    if (bool.Parse(ConfigurationManager.AppSettings["VisualizarLogo"]))
    {
  %> 
 <table   border="0" cellspacing="0" cellpadding="0" align="center" class="Encabezado" id="Encabezado">
 <% } else { %>
 <table width="100%" height="10" border="0" cellspacing="0" cellpadding="0" align="center" class="EncabezadoReducido" id="Encabezado"  >
 <% } %>
 
  <tr>
  
      
            <td class="TD_LOGO" >     
           
             <asp:LinkButton ID="BtnHome" OnCommand="Visualizar_Gadgets"  runat="server"    >                                                          
                        
              <% if (bool.Parse(ConfigurationManager.AppSettings["VisualizarLogo"]))
                { %>        
                    <img src="<%=ConfigurationManager.AppSettings["urlLogo"] %>" class="LOGO_EMPRESA"> 
                    
              <% } %>
              
                 </asp:LinkButton>  
              
            </td>
   
    
    
    <td class="TD_BARRA_TOP"    >           
 
       <TABLE border="0" cellspacing="0" cellpadding="0" class="TABLA_MENU_TOP" align="right">
       <TR>
        <%   if (bool.Parse(ConfigurationManager.AppSettings["VisualizarComplementos"]))
          { %> 
          
          
          <%  if (Common.Utils.IsUserLogin)
            {    %>  
               <TD class="TD_GADGET">
                    <span class="MENU_TOP_NAV idioma"  >                        
                       <img src="img/Modulos/SVG/GADGET.svg" border='0'   class="IconosBarraTop" onclick="AbrirModal()"   >                                
                     </span>       
               </TD>
       
          <% }%>
       
      <% } %>
         <TD class="idiomas">
              <span  id="Btn_Idioma" onclick="Abrir_Globo('Globo_Idiomas')"   class="MENU_TOP_NAV"  >        
                    
                    <img src="img/Modulos/SVG/IDIOMA.svg" border='0' class="IconosBarraTop" title="<%  Response.Write( ObjLenguaje.Label_Home("Idioma") ); %>"    >       
                      <%              
               
                          Response.Write(System.Web.HttpContext.Current.Session["ArgTitulo"]);   
               
           %> 
              </span>               
         </TD>  
         <TD class="user">              
            <span id="Btn_Loguin" onclick="PopUp_Abrir();" class="MENU_TOP_NAV">
             
                <img src="img/Modulos/SVG/USER.svg" border='0' class="IconosBarraTop"    >   
                <asp:Label id="Ingresar"  runat="server"     >         
           <%              
               if (!Common.Utils.IsUserLogin)
                 Response.Write( ObjLenguaje.Label_Home("Ingresar") ); 
               else
                   Response.Write(Common.Utils.SessionUserName);   
               
           %> 
          </asp:Label> 
        </span>                          
                         
      </TD>
      </TR>
      </TABLE>
 
    </td>
<tr> 
<td class="LineaEncabezado" colspan="2"> </td>   
</tr>
  </tr>
 </table>

<asp:UpdatePanel ID="Update_Modulos" runat="server" UpdateMode="Conditional">
       <ContentTemplate>                    
    
 <% 
    if (bool.Parse(ConfigurationManager.AppSettings["VisualizarLogo"]))
    {
  %> 
  <table   border="0" cellspacing="0" cellpadding="0" align="center" style="margin-top:-23px;"  class="Principal" >
 <% } else { %>
  <table   border="0" cellspacing="0" cellpadding="0" align="center" style="margin-top:0px;" class="Principal"   >
 <% } %>

  
 
   <tr   class="PreTop" >
    <td  >           
            <table border="0" cellpadding="0" cellspacing="0"  height="1" align="left" class="PreTop_Izq" >
            <tr> 
              <td valign="middle" align="center"    class="PreTop_IzqTop" style="white-space:nowrap" >        
                    
                <%=Traducir_Fecha()%> 
                
                  
                </td>
            </tr>
            </table>

    </td>
    
    <td    align="left" valign="middle"    class="PreTop_Der">
    
     
        
         <asp:PlaceHolder  id="InfoCambioPass" runat="server"></asp:PlaceHolder>
        
    </td>
  </tr>
  <tr>
    <td height="186" align="left" valign="top" class="PanelModulos" id="ContenedorModulos"  >
        
        <cc:Modulos id="Modulos" runat="server"     >
        </cc:Modulos>
       
     </td>
    
    <% if (bool.Parse(ConfigurationManager.AppSettings["VisualizarImgContenedorPpal"]))
       { %>
        <td   align="left" valign="top" width="100%"  style="height:250px; background: url(<%= ConfigurationManager.AppSettings["ImgContenedorPpal"]%>) no-repeat right bottom; background-color:#FFFFFF;" >
    <%  }
       else
       { %> 
        <td   align="left" valign="top" width="100%"  style="height:250px; background-color:#FFFFFF;" > 
    <% } %>
     
     
      
                  <cc:contenedorprincipal id="ContenedorPrincipal" runat="server">
                  </cc:contenedorprincipal>                  
      
              
      
     
      </td>

 
 
  </tr>
  
   <tr>
    
      
       <% if (bool.Parse(ConfigurationManager.AppSettings["VisualizarFooter"]))
          { %>
    <td colspan="2" class="Piso">      
              <table width="100%"  border="0" cellspacing="0" cellpadding="0" >
                <tr>
                  <td width="48%">
                  <p></p>
                   <p class="TitBase" >
                      <b>Versión:</b>  <label runat="server" id="versionMI" /> <b>Patch:</b> <label runat="server" id="patchMI" /> 
                   </p>
                    </td>
                  <td width="52%" rowspan="2" valign="top" align="right" style="height:70px" >
                    <span class="Frase" >Simplificamos su trabajo. Optimizamos su gestión.</span> 
                    </td>
                  </tr>
                <tr>
                  <td>
                    <p class="DetEmpresa">
                      Heidt & Asociados S.A.  <br>
                      Suipacha 72 - 4º A CP C1008AAB - Buenos Aires Argentina.<br />
                      Tel./Fax: +54 11 5252 7300 Email: ventas@rhpro.com 
        			  
                    </p></td>
                  </tr>
              </table>
      <% }        else       {%>
      <td colspan="3" style="height:10px; background-color:#262626"> 
      <% } %>
      
      </td>
  </tr>
</table>

</ContentTemplate>
        
     </asp:UpdatePanel>
    
   <cc:ConfigGadgets id="ConfigGadgets" runat="server"   >
   </cc:ConfigGadgets> 
   
    <cc:Idiomas  id="cIdiomas" runat="server" >
    </cc:Idiomas>
 
    <cc:CustomLogin id="cLogin" runat="server" ContentPlaceHolderID="cLogin"   >
    </cc:CustomLogin>   
 
<script language="javascript">
 document.getElementById("TG").style.left =  ((AnchoPantalla()/2) - 350) + "px";
 //document.getElementById("FondoTransparente").style.height =  window.document.body.scrollHeight + "px";
 
</script>




 
</body>
</asp:Content>


