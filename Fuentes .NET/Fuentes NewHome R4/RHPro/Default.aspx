<%@  Page Title="" Language="C#" MasterPageFile="~/MasterPage.Master" AutoEventWireup="true" 
CodeBehind="Default.aspx.cs" Inherits="RHPro.Default" EnableViewState="true"   %>
 
 
<asp:Content ID="Content2" ContentPlaceHolderID="content" runat="server"  > 
 

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
             //Cerrar_Globo('Globo_Idiomas');
         } else {
                Cerrar_Globo("Globo_Loguin");
               // Centrar_Globo('Globo_Idiomas');
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
 


function DesplazarBarraMenu_Ajax(valor) {

    var Tabla_Modulos = document.getElementById("CC_Tabla_Modulos");
    var Tabla_Modulos_Inactivos = document.getElementById("CC_Tabla_Modulos_Inactivos");
    var Tabla_Accesos = document.getElementById("CC_Tabla_Accesos");
    var PanelModulos = document.getElementById("CC_PanelModulos");


    if (valor == "1") {

        Tabla_Modulos.className = "Menu_Links_Colapsado";
        if (Tabla_Modulos_Inactivos)
            Tabla_Modulos_Inactivos.className = "Menu_Links_ColapsadoInact";
        Tabla_Accesos.className = "Menu_Links_Colapsado";
        PanelModulos.style.columnWidth = "29px";
    }
    else {
        Tabla_Modulos.className = "Menu_Links";
        if (Tabla_Modulos_Inactivos)
            Tabla_Modulos_Inactivos.className = "Menu_Links_Inact";
        Tabla_Accesos.className = "Menu_Links";
        PanelModulos.style.columnWidth = "";
    }
 
}


var ControlExpand = 1;
function DesplazarBarraMenu() {

    
     
    var Tabla_Modulos = document.getElementById("CC_Tabla_Modulos");
    var Tabla_Modulos_Inactivos = document.getElementById("CC_Tabla_Modulos_Inactivos");
    var Tabla_Accesos = document.getElementById("CC_Tabla_Accesos");
    var PanelModulos = document.getElementById("CC_PanelModulos");
     
  //  if (Tabla_Modulos.className == "Menu_Links") {
    if (ControlExpand==1) {
        ControlExpand = 0;
        Tabla_Modulos.className = "Menu_Links_Colapsado";
       
        if (Tabla_Modulos_Inactivos)
            Tabla_Modulos_Inactivos.className = "Menu_Links_ColapsadoInact";
       
        Tabla_Accesos.className = "Menu_Links_Colapsado";
        //PanelModulos.style.columnWidth = "39px";
        //Modifico el display de la clase para que se oculte
        $('.ASPlink  span,.ASPlink  div').css('display', 'none');
    }
    else {
        ControlExpand = 1;
        Tabla_Modulos.className = "Menu_Links";
       
        if (Tabla_Modulos_Inactivos)
            Tabla_Modulos_Inactivos.className = "Menu_Links_Inact";
       
        Tabla_Accesos.className = "Menu_Links";
        PanelModulos.style.columnWidth = "";
        //Modifico el display de la clase para que se restaure al tamaño original
        $('.ASPlink  span,.ASPlink  div').css('display', 'inline-block');
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


var isCtrl = false;

document.onkeyup = function(e) {
    var evento = e || window.event;
    var keyCode = evento.keyCode || evento.which;
    
    if (keyCode == 17) 
      isCtrl = false;
}

document.onkeydown = function(e) {
    var evento = e || window.event;
    var keyCode = evento.keyCode || evento.which;

    if (keyCode == 17) isCtrl = true;

    if (keyCode == 76 && isCtrl == true) {//Apertura del loguin con la tecla Ctrl+L
        
        document.getElementById("<%=Btn_Login_MenuTop.ClientID%>").click();
        return false;
    }
    /*
    if (e.which == 73 && isCtrl == true) {
    document.getElementById("Btn_Idiomas_MenuTop").click();
    return false;
    }

 
    if (e.which == 69 && isCtrl == true) {
    document.getElementById("Btn_Estilos_MenuTop").click();
    return false;
    }

    if (e.which == 71 && isCtrl == true) {
    document.getElementById("Btn_Gadgets_MenuTop").click();
    return false;
    }

    if (e.which == 70 && isCtrl == true) {
    document.getElementById("Btn_Favoritos_MenuTop").click();
    return false;
    }
    */

}

 
 
</script>




 
<!-- ---------------------------------------------------------------------- -->
<!-- SmartMenus jQuery plugin -->
 

 
<script type='text/javascript' >
 
    $(function() {
        $('#main-menu').smartmenus({
            subMenusSubOffsetX: 0,
            subMenusSubOffsetY: -1,
            hideOnClick: false  
        });
    });
/*
    $(function() {
        $('#main-menuTop').smartmenus({
            subMenusSubOffsetX: 0,
            subMenusSubOffsetY: 0,
            mainMenuSubOffsetX:0,
            mainMenuSubOffsetY:0,
            subMenusMinWidth:"60px",
            subMenusMaxWidth: "1060px" ,
            hideOnClick: false 
        });                
    });

 */
 
    $(function() {
    $('#main-menuTopLoguin').smartmenus({
            subMenusSubOffsetX: 0,
            subMenusSubOffsetY: 0,
            mainMenuSubOffsetX: 0,
            mainMenuSubOffsetY: 0,
            subMenusMinWidth: "60px",
            subMenusMaxWidth: "1060px",
            hideOnClick:true 
        });
    });
 
 
</script>  
<link href='css/sm-core-css.css' rel='stylesheet' type='text/css' /> 
<link href='css/sm-clean.aspx' rel='stylesheet' type='text/css' />
<link href='css/sm-cleanTop.css' rel='stylesheet' type='text/css' />

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
	
	#main-menuTopLoguin {
		position:relative;
		z-index:9999;
		width:auto;
	}
	#main-menuTopLoguin ul {
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
<!--
<span id="ControlaID"></span>
<span id="ControlaIDMod"></span>
 -->
 <table   border="0" cellspacing="0" cellpadding="0" align="center"style="width:100% !important" >
 <tr>
 <td align="center" valign="top">
 
 
 
 <asp:UpdatePanel ID="Update_Mod" runat="server"  >
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
   
     
    
         <TD class="MENU_TOP_NAV" align="right" nowrap="nowrap" >                
       <UL id="main-menuTopLoguin" class="sm sm-blueTop" onclick=" this.style.zIndex=400; if (document.getElementById('main-menu')){document.getElementById('main-menu').style.zIndex=300;  }"   >    
           <%  if (Common.Utils.IsUserLogin)
            {    %> 
          
            
            <% if (Habilitado("Favoritos","Default.aspx")) { %>
            
            <li>
                <a  id="Btn_Favoritos_MenuTop" href="#" class='BtnTransparenteTop' > <div class="LinkMega">    </div>   </a> 
                
                <div class='CajaBtnTop'>
                    <asp:Panel id="Info_Favoritos" runat="server"></asp:Panel>
                </div>
                <ul class="mega-menu"   style="width:833px !important; ">
                  <li>                    
                           <asp:Panel id="CFavoritos2" runat="server"></asp:Panel>
                           <asp:PlaceHolder id="Cfavoritos" runat="server"></asp:PlaceHolder>
                  </li>
                </ul>
            </li>
            
            <% } %>
            
            
           <% if (Habilitado("Gadgets","Default.aspx")) { %>
            <li>
                <a id="Btn_Gadgets_MenuTop" href="#" class='BtnTransparenteTop' > <div class="LinkMega">  </div>     </a> 
            
             <div class='CajaBtnTop'>
                <asp:Panel id="Info_Complementos" runat="server"></asp:Panel>
              </div>
                
                <ul class="mega-menu"  style="width:833px !important; ">
                  <li>     
                                    
                        <cc:ConfigGadgets id="ConfigGadgets" runat="server"   ></cc:ConfigGadgets>                                                                                      
                                    
                  </li>
                </ul>
            </li>
           <% } %> 
           
           <% if (Habilitado("Estilos","Default.aspx")) { %>
           <li>
               <a id="Btn_Estilos_MenuTop" href="#" class='BtnTransparenteTop' > <div class="LinkMega">    </div>   </a>
                        
                <div class='CajaBtnTop'>
                    <asp:Panel id="Info_Estilos" runat="server"></asp:Panel>
                </div>
                <ul class="mega-menu"  style="width:833px !important; "  >
                  <li>     
                                                      
                        <cc:SelectorEstilos id="CEstilos" runat="server"   ></cc:SelectorEstilos>                                                  
                                  
                  </li>
                </ul>
            </li>
           <% } %> 
            
        <% }%> 
        
        
            <li>
              <a id="Btn_Idiomas_MenuTop" href="#" class='BtnTransparenteTop' > <div class="LinkMega">  </div>  </a>
                             
              <div class='CajaBtnTop'>
                <asp:Panel id="Info_Lenguaje" runat="server"></asp:Panel>
              </div>
                <ul class="mega-menu"  style="width:833px !important; ">
                  <li>                          
                        <cc:Idiomas  id="cIdiomas" runat="server" ></cc:Idiomas>                                              
                  </li>
                </ul>
            </li>
            
            <li> 
                 <a id="Btn_Politicas_MenuTop" href="#" class='BtnTransparenteTop' >  <div class="LinkMega">   </div>        </a> 
                
                  <div class='CajaBtnTop'>
                    <asp:Panel id="Info_Politicas" runat="server"></asp:Panel>    
                  </div>
                <ul class="mega-menu"  style="width:833px !important; ">
                  <li>     
                       
                    <iframe src="PopUpPolitics.aspx" style="height:100%; width:100%; border:0" class='' ></iframe>                    
                    
                                           
                  </li>
                </ul>
            </li>
            
           
             <li>              
                 <a  id="Btn_Login_MenuTop" href="#" runat="server" class='BtnTransparenteTop' > 
                    <div class="LinkMega"  >  </div>
                 </a>
                  <div class='CajaBtnTop'>
                    <asp:Panel id="Info_Usuario" runat="server"></asp:Panel> 
                  </div>
                <ul class="mega-menu">
                  <li>  
                      <cc:CustomLogin id="cLogin" runat="server" ContentPlaceHolderID="cLogin"   >  </cc:CustomLogin>      
                            
                  </li>
                </ul> 
             </li>
           </UL>  
          
                                  
    
 
  
    </td>
<tr> 
<td class="LineaEncabezado" colspan="2"> </td>   
</tr> 
 </table>
 
 
 
<table border="0" cellspacing="0" cellpadding="0 align="center"   > 

 <tr   class="PreTop" > 
    <td  >           
            <table border="0" cellpadding="0" cellspacing="0"  height="1" align="left" class="PreTop_Izq" >
            <tr>
              <td align="left"> 
                
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
    <td height="186"  align="left" valign="top" class="PanelModulos" id="CC_PanelModulos"    >
   
                  <asp:Panel id="CModulos" runat="server"></asp:Panel>
                  
                                  

     
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
                   <td width="48%" valign="top" align="left" style="text-align:left !important; vertical-align:top !important">                  
                        <div class="TitBase" >
                          <label runat="server" id="versionMI" /> 
                        </div>
                        <div class="TitBase" >
                          <label runat="server" id="patchMI" /> 
                        </div>
                        <div id="NombreDeEmpresa" class="TitBase" runat="server"> </div>
                        <div id="DireccionEmpresa"   class="TitBase" runat="server"> </div>
                        <div id="TelefonoMailEmpresa" class="TitBase" runat="server"> </div> 
                   </td>
                   
                   <td class="DetRedes" nowrap valign="top" align="center">
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

    
          </ContentTemplate>                      
     </asp:UpdatePanel>
     
 

</TD>
  </TR>
</TABLE>
 
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
              
                <asp:Panel id="Panel_Generico" runat="server"></asp:Panel>    
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
                  
           </td>
          </tr>
        </TABLE>            
        
        
        
 
 
<script language="javascript">
 //document.getElementById("TG").style.left =  ((AnchoPantalla()/2) - 350) + "px";
 //document.getElementById("FondoTransparente").style.height =  window.document.body.scrollHeight + "px";

   
</script>

  <iframe id="ifrmEst" name="ifrmEst" style="display:none" runat="server"></iframe>
  <% 
  if (Common.Utils.IsUserLogin)
      {
  %>
        <script>
 

            //Control del refresco de las variables de sesion ASP
            window.onload = function() {
              FN_Control_Session_ASP();
           
            }
            
         
            function FN_Control_Session_ASP() {
                document.IfrmControlTimer.location = "../timer.asp";
                timerID = setTimeout("FN_Control_Session_ASP()", 10000);  
            }
        </script>
        
        <iframe src="../timer.asp" id="IfrmControlTimer" name="IfrmControlTimer" style="display:none"  ></iframe>
        
    <%}%>
</body>
</asp:Content>


