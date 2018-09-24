<%@ Page Title="" Language="C#" MasterPageFile="~/MasterPage.Master" AutoEventWireup="true" 
CodeBehind="Default.aspx.cs" Inherits="RHPro.Default" EnableViewState="true"  %>
 
 
<asp:Content ID="Content2" ContentPlaceHolderID="content" runat="server" > 

 
 

<head>
 

<!-- Hoja de Estilos construida desde un control -->    
    <cc:estilosnewhome id="EstilosNewHomes" runat="server"    >        
    </cc:estilosnewhome>
<!-- ------------------------------------------- -->    

<%    if (Common.Utils.IsUserLogin)
      { %>
         <script type="text/javascript" src="/rhprox2/shared/js/fn_windowsmi.asp"></script>
   <% }
      else
      { %>
      <script type="text/javascript" src="/rhprox2/shared/js/fn_windows.js"></script>
   <% } %>
       
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
    // document.cookie;
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

var Modulos_Colapsados = false;
var Ancho_Contenedor_Modulos = "300";
var Ancho_MinContenedor_Modulos = "34";
window.onload = function() {

    var exp_mod = document.getElementById("ContenedorModulos");
    var Panel = document.getElementById("Fondo_Contenedor_Principal");
    if (Modulos_Colapsados) {
        exp_mod.style.width = Ancho_MinContenedor_Modulos + "px";
        Panel.style.borderLeft = "1px solid #cccccc";
    }
    else {
        exp_mod.style.width = Ancho_Contenedor_Modulos + "px";
        Panel.style.borderLeft = "0px";
        //exp_mod.style.width = "245px";
    }
}

function DesplazarBarraMenu() {
    var obj = document.getElementById("ContenedorModulos");
    var Panel = document.getElementById("Fondo_Contenedor_Principal");

    if (obj.style.width == Ancho_MinContenedor_Modulos + "px") {
        obj.style.width = Ancho_Contenedor_Modulos + "px";
        Modulos_Colapsados = false;
      Panel.style.borderLeft = "0px"; 
    }
    else {
        obj.style.width = Ancho_MinContenedor_Modulos + "px";
        Modulos_Colapsados = true;
         Panel.style.borderLeft = "1px solid #cccccc";
    }
}


function Abrir_PopUp_Generico(id) {
    //Contenedor_Favoritos
    //Contenedor_Estilos
    var pp = document.getElementById("PopUp_Generico");
    var cvg = document.getElementById(id);
    
    pp.style.display = "";
    cvg.style.display = "";
}

function Cerrar_PopUp_Generico(id) {
    var pp = document.getElementById("PopUp_Generico");
    var cvg = document.getElementById(id);
    pp.style.display = "none";
    cvg.style.display = "none";
}

        
</script>




 
<!-- ---------------------------------------------------------------------- -->
<!-- SmartMenus jQuery plugin -->
 

 
<script type='text/javascript' >
    $(function() {
        $('#main-menu').smartmenus({
            subMenusSubOffsetX: 0,
            subMenusSubOffsetY: -1
        });
        $(function() {
            $('#main-menuTop').smartmenus({
                subMenusSubOffsetX: 0,
                subMenusSubOffsetY: -1
            });
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
 
	#main-menuTop {
		position:relative;
		z-index:9999;
		width:auto;
	}
	#main-menuTop ul {
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

 
 
 <% Response.Write(ImpBody); %>
   
 
<asp:UpdatePanel ID="Update_Modulos" runat="server" UpdateMode="Conditional">
       <ContentTemplate>    
 <table   border="0" cellspacing="0" cellpadding="0" align="center" class="Encabezado" id="Encabezado">
 
  <tr>     
        <td class="TD_LOGO" >     
       
         <asp:LinkButton ID="BtnHome" OnCommand="Visualizar_Gadgets"  runat="server"    >                                                          
                    
          <% if (bool.Parse(ConfigurationManager.AppSettings["VisualizarLogo"]))
            { %>        
                <img  runat="server" id="Logo_Empresa" class="LOGO_EMPRESA"> 
                
          <% } %>
          
             </asp:LinkButton>  
          
        </td>
   
    <td class="TD_BARRA_TOP"    >        
    
    <ul id="main-menuTop" class="sm sm-blue">
       
 
       <TABLE border="0" cellspacing="0" cellpadding="0" class="TABLA_MENU_TOP" align="right">
       <TR>
        <%   if (bool.Parse(ConfigurationManager.AppSettings["VisualizarComplementos"]))
          { %> 
          
          
          <%  if (Common.Utils.IsUserLogin)
            {    %> 
               <TD class="TD_GADGET MENU_TOP_NAV idioma"  onclick="Abrir_PopUp_Generico('Contenedor_Estilos')" >                                             
                   <asp:Panel id="Info_Estilos" runat="server"></asp:Panel>  
               </TD>
               
               <TD class="TD_GADGET MENU_TOP_NAV idioma" id="BtnFavoritos"  onclick="Abrir_PopUp_Generico('Contenedor_Favoritos')"  >          
                    <asp:Panel id="Info_Favoritos" runat="server"></asp:Panel>                             
               </TD> 
               
               <TD class="TD_GADGET MENU_TOP_NAV idioma"   onclick="Abrir_PopUp_Generico('Contenedor_Gadgets')">                                             
                   <asp:Panel id="Info_Complementos" runat="server"></asp:Panel>  
               </TD>
       
          <% }%>
       
      <% } %>
         <TD class="idiomas MENU_TOP_NAV" id="Btn_Idioma" onclick="Abrir_Globo('Globo_Idiomas')"  >
              
              
                <asp:Panel id="Info_Lenguaje" runat="server"></asp:Panel>                
         </TD>  
         <TD class="user MENU_TOP_NAV" id="Btn_Loguin" onclick="PopUp_Abrir();">              
                      
                <asp:Panel id="Info_Usuario" runat="server"></asp:Panel>              
                                                         
         </TD>
      </TR>
      </TABLE>
 
 
 
 
    </td>
<tr> 
<td class="LineaEncabezado" colspan="2"> </td>   
</tr> 
 </table>
 
<table border="0" cellspacing="0" cellpadding="0 align="center"> 
 <tr   class="PreTop" >
    <td  >           
            <table border="0" cellpadding="0" cellspacing="0"  height="1" align="left" class="PreTop_Izq" >
            <tr>
              <td align="left"> 
                <div class='BarraDesplazamientoModulos' onclick='DesplazarBarraMenu()'> 
                  <img src="img/modulos/svg/CONTROLMODULOS.svg" class="IconosMaximizaModulos"> 
                </div> 
              </td>
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

</table>
                
    
 <% 
    if (bool.Parse(ConfigurationManager.AppSettings["VisualizarLogo"]))
    {
  %> 
  <table   border="0" cellspacing="0" cellpadding="0" align="center" style="margin-top:-23px;"  class="Principal" >
 <% } else { %>
  <table   border="0" cellspacing="0" cellpadding="0" align="center" style="margin-top:0px;" class="Principal"   >
 <% } %>

  
 
   
  <tr>
    <td height="186"  align="left" valign="top" class="PanelModulos"    >
        <DIV id="ContenedorModulos" class="ContenedorModulos">
          <cc:Modulos id="Modulos" runat="server"     >
          </cc:Modulos>
       </DIV>
     </td>
 
    
    <% if (bool.Parse(ConfigurationManager.AppSettings["VisualizarImgContenedorPpal"]))
       { %>
        <td   align="left" valign="top" width="100%"  style="height:250px; background: url(<%= ConfigurationManager.AppSettings["ImgContenedorPpal"]%>) no-repeat right bottom; background-color:#FFFFFF;" >
    <%  }
       else
       { %> 
        <td   align="left" valign="top" width="100%"  class='Fondo_Contenedor_Principal' id="Fondo_Contenedor_Principal"   > 
    <% } %>
     
     
      
                  <cc:contenedorprincipal id="ContenedorPrincipal" runat="server">
                  </cc:contenedorprincipal>                  
      
              
      
     
      </td>

 
 
  </tr>
  
   <tr>
    
      
       <% if (bool.Parse(ConfigurationManager.AppSettings["VisualizarFooter"]))
          { %>
    <td colspan="2" class="Piso">      
            
            
              <table width="100%"   height="100%" border="0" cellspacing="0" cellpadding="0" >
                <tr>
                   <td width="48%" valign="top">                  
                        <div class="TitBase" >
                        Versión:  <label runat="server" id="versionMI" /> 
                        </div>
                        <div class="TitBase" >
                        Patch:  <label runat="server" id="patchMI" /> 
                        </div>
                        <div id="NombreDeEmpresa" class="TitBase" runat="server"> </div>
                        <div id="DireccionEmpresa"   class="TitBase" runat="server"> </div>
                        <div id="TelefonoMailEmpresa" class="TitBase" runat="server"> </div> 
                   </td>
                   
                   <td class="DetRedes" nowrap valign="top">
                       <asp:Panel id="RedesSociales" runat="server"></asp:Panel>
                     </td>
                    <td width="52%" rowspan="2" valign="bottom" align="right" style="height:70px; padding:4px;" nowrap>
                      
                       <div class="Frase" id="Slogan" runat="server"  ></div>                      
                      <asp:Panel id="URL_Logo" runat="server"></asp:Panel>
                      
                     
                    </td>
                  
                  </tr>
                
              </table>
              
              
      <% }        else       {%>
      <td colspan="3" style="height:10px; background-color:#262626"> 
      <% } %>
      
      </td>
  </tr>
</table>


  
<cc:Idiomas  id="cIdiomas" runat="server" >
</cc:Idiomas>
 
    


    
</ContentTemplate> 
   
</asp:UpdatePanel>
 
 
    <!-- ############################## FONDO TRANSPARENTE ##################################----->        
     <DIV ID="PopUp_Generico"  Class="PopUp_FondoTransparente"  style="display:none" ></DIV>
    <!-- ############################## VENTANA FAVORITOS ##################################----->      
    <TABLE cellpadding="0" cellspacing="0" border="0" id="Contenedor_Favoritos"   class="Contenedor_Ventana_Generica" style="display:none">
          <tr class="PopUp_Cabecera">
           <td>   
              <span id="TituloIdi" runat="server"></span>    
               <asp:LinkButton id="BtnCloseGenerica"  onclick="Cerrar_PopUp_Generico('Contenedor_Favoritos')"    >   <span Class="cerrarVentana">X  </span>   </asp:LinkButton>                         
                      
           </td>
          </tr>
          <tr class="PopUp_DataUser">
           <td valign="top" style="text-align:center !important; ">  
               <cc:Favoritos id="CFavoritos" runat="server"   ></cc:Favoritos> 
               <asp:Panel id="Panel_Generico" runat="server"></asp:Panel>  
           </td>
          </tr>
        </TABLE>                      
  
  
        
  
   <!-- ############################## VENTANA ESTILOS ##################################----->      
    <TABLE cellpadding="0" cellspacing="0" border="0" id="Contenedor_Estilos"   class="Contenedor_Ventana_Estilos" style="display:none">
          <tr class="PopUp_Cabecera">
           <td valign="middle">   
              <span id="TituloEstilos" class='TituloVentana' runat="server"></span>    
               
                <asp:LinkButton id="BtnClose"  onclick="Cerrar_PopUp_Generico('Contenedor_Estilos')"   > <span class="cerrarVentana">X   </span>    </asp:LinkButton>                         
                     
           </td>
          </tr>
          <tr class="PopUp_DataUser">
           <td valign="top" style="text-align:center !important; ">  
               <cc:SelectorEstilos id="CEstilos" runat="server"   ></cc:SelectorEstilos>                
           </td>
          </tr>
        </TABLE>    

<!-- ############################## VENTANA REPOSITORIO DE COMPLEMENTOS ##################################----->      
    <TABLE cellpadding="0" cellspacing="0" border="0" id="Contenedor_Gadgets"   class="Contenedor_Ventana_Gadgets" style="display:none">
          <tr class="PopUp_Cabecera">
           <td valign="middle">   
              <span id="TituloGadgets" class='TituloVentana' runat="server"></span>    
               
                <asp:LinkButton id="BtnCloseGadgets"  onclick="Cerrar_PopUp_Generico('Contenedor_Gadgets')"   > <span class="cerrarVentana">X   </span>    </asp:LinkButton>                         
                     
           </td>
          </tr>
          <tr class="PopUp_DataUser">
           <td valign="top" style="text-align:center !important; ">  
            <cc:ConfigGadgets id="ConfigGadgets" runat="server"   ></cc:ConfigGadgets>           
           </td>
          </tr>
        </TABLE>            
        
        
        
    <cc:CustomLogin id="cLogin" runat="server" ContentPlaceHolderID="cLogin"   >
    </cc:CustomLogin>   
 
<script language="javascript">
 //document.getElementById("TG").style.left =  ((AnchoPantalla()/2) - 350) + "px";
 //document.getElementById("FondoTransparente").style.height =  window.document.body.scrollHeight + "px";
 
</script>

  <iframe id="ifrmEst" name="ifrmEst" style="display:none" runat="server"></iframe>
 
</body>
</asp:Content>


