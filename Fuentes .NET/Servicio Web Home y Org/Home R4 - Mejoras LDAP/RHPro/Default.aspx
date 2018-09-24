<%@ Page Title="" Language="C#" MasterPageFile="~/MasterPage.Master" AutoEventWireup="true" 
CodeBehind="Default.aspx.cs" Inherits="RHPro.Default" EnableViewState="true"  %>
 

<asp:Content ID="Content2" ContentPlaceHolderID="content" runat="server" >
 
    
    <!-- -------------------------------------------------------------------->
 
     
<style>
 

 
.DetalleEmpresa{
	 color:#CCCCCC; font-size:7.5pt; font-family:Arial; margin-top:6px;  margin-left:6px;
	} 
.TituloBase{ color:#FFFFFF; font-size:7pt; font-family:Arial; margin-left:6px; margin-top:6px; }

.Detalle{ color:#CCC; font-size:12pt; font-family:Arial;}

.user{ color:#333; font-family:Arial; font-size:11pt; font-weight:bold;}

#user_izq { background:url(img/Loguin/user_izq.png) no-repeat right bottom;}
#user_centro1 {background:url(img/Loguin/user_centro.png) repeat-x center}
#user_centro2 {background:url(img/Loguin/user_centro.png) repeat-x center}
#user_der { background:url(img/Loguin/user_der.png) no-repeat right bottom;}

#idioma_izq { background:url(img/Loguin/user_izq.png) no-repeat right bottom;}
#idioma_centro1 {background:url(img/Loguin/user_centro.png) repeat-x center}
#idioma_centro2 {background:url(img/Loguin/user_centro.png) repeat-x center}
#idioma_der { background:url(img/Loguin/user_der.png) no-repeat right bottom;}
 
 
#Btn_Loguin{ cursor:pointer}
#Btn_Loguin:hover TD#user_izq{ background:url(img/Loguin/user_izq_hover.png) no-repeat right bottom;}	 
#Btn_Loguin:hover TD#user_centro1{ background:url(img/Loguin/user_centro_hover.png); color:  #999}	
#Btn_Loguin:hover TD#user_centro2{ background:url(img/Loguin/user_centro_hover.png); color: #999}	
#Btn_Loguin:hover TD#user_der{ background:url(img/Loguin/user_der_hover.png) no-repeat right bottom;}

#Btn_Idioma{ cursor:pointer}
#Btn_Idioma:hover TD#idioma_izq{ background:url(img/Loguin/user_izq_hover.png) no-repeat right bottom;}	 
#Btn_Idioma:hover TD#idioma_centro1{ background:url(img/Loguin/user_centro_hover.png); color: #999}	
#Btn_Idioma:hover TD#idioma_centro2{ background:url(img/Loguin/user_centro_hover.png); color: #999}	
#Btn_Idioma:hover TD#idioma_der{ background:url(img/Loguin/user_der_hover.png) no-repeat right bottom;}

.idioma{ color:#333; font-family:Arial; font-size:9pt; }
.fecha{ color:#333; font-family:Arial; font-size:9pt; }

Body
{
 /*EXPLORER:  */  /*  filter:progid:DXImageTransform.Microsoft.Gradient(endColorstr='#F3F4F6', startColorstr='#AAAEBB', gradientType='0');    */
/*SAFARI y CHROME:  */ /* background: -webkit-gradient(linear, left top, left bottom, from(#AAAEBB), to(#F3F4F6)); */
/*MOZILLA: */ /*  background: -moz-linear-gradient(top, #AAAEBB,#F3F4F6);*/
/*OPERA*/ /* background: -o-linear-gradient(#F3F4F6, #AAAEBB); */
 /*background:url(img/fondo.png) repeat;*/

 background:url(img/FONDO.jpg) no-repeat top center; 
 background-color:#c6c6c6;
 margin-top:0px;
 padding-top:0px;
}

.Gadget
{
	background:url(img/BotonGadget.png) no-repeat center; width:183px; height:36px;
	text-align:center;
	margin-left:3px; 
	vertical-align: bottom;
	
}

.BtnGadget
{
	color:#FFFFFF;
	font-family:Arial;
	font-size:14pt;
	font-weight:bold;	 
	text-decoration:none;
	vertical-align:bottom;

	 
}
.BtnGadget:hover
{
	color:#333333;
	 
}

    .style1
    {
        height: 53px;
        width: 239px;
    }
    
     
    
    .style3
    {
        height: 16px;
    }
  
     
    
    .style4
    {
        height: 10px;
        width: 40%;
    }
  
     
    
</style>
 

<script type="text/javascript" >
 
 function Abrir_Globo(nombre){   
     document.getElementById(nombre).style.visibility = "visible";
     if (nombre=="Globo_Loguin")
       Cerrar_Globo('Globo_Idiomas');        
     else
       Cerrar_Globo('Globo_Loguin');     
 }
 
 function Cerrar_Globo(nombre){  
     document.getElementById(nombre).style.visibility = "hidden";  
	 document.cookie
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

</script>

 
<%
    /*Aqui controlo si el sitio se va a refrescar para evitar la perdida de objetos en memoria. (Ver. Settings.config)*/
    string ImpBody = " <body  leftmargin='0' topmargin='0' rightmargin='0' bottommargin='0' >";
    if ((ConfigurationManager.AppSettings["TiempoRefrescoActivo"] != null) && (ConfigurationManager.AppSettings["TiempoRefresco"] != null))
    {

        if (bool.Parse(ConfigurationManager.AppSettings["TiempoRefrescoActivo"]))
        {
            ImpBody = " <body  leftmargin='0' topmargin='0' rightmargin='0' bottommargin='0' ";
            ImpBody += "onload=\"RefrescarPagina('Default.aspx'," + int.Parse(ConfigurationManager.AppSettings["TiempoRefresco"].ToString()) + ");\"";
            ImpBody += " >";
        }
    }  
%>

<!--body  leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0" onmousemove='if (Mover()) {}'  -->
 
 <% Response.Write(ImpBody); %>
 
 <% 
    if (bool.Parse(ConfigurationManager.AppSettings["VisualizarLogo"]))
    {
  %> 
 <table width="100%" height="81" border="0" cellspacing="0" cellpadding="0" align="center" style=" margin-top:0px; padding-top:0px; background:url(img/Header.png) no-repeat top ">
 <% } else { %>
 <table width="100%" height="10" border="0" cellspacing="0" cellpadding="0" align="center" style="margin-top:3px; padding-top:0px; margin-bottom:-12px ">
 <% } %>
 
  <tr>
    <td align="left" valign="bottom" class="style4" width="69%"> 
     <% if (bool.Parse(ConfigurationManager.AppSettings["VisualizarLogo"]))
        { %>
            <img src="<%=ConfigurationManager.AppSettings["urlLogo"] %>" height="60" style="margin-left:30px; margin-bottom:3px;"> 
      <% } %>
    </td>
    <td width="14%" align="right" valign="middle"     >      
      <cc:CustomLogin id="cLogin" runat="server" ContentPlaceHolderID="cLogin"   >
      </cc:CustomLogin>
    </td>
    <td width="17%"  align="left" valign="middle">   
      <!-- -- --------- IDIOMAS -------------------------------->
      <cc:Idiomas  id="cIdiomas" runat="server"  style="vertical-align:bottom;">
        </cc:Idiomas>
      <!-- -------------------------------------- -->  
      </td>    
  </tr>
 </table>


 <% 
    if (bool.Parse(ConfigurationManager.AppSettings["VisualizarLogo"]))
    {
  %> 
  <table   border="0" cellspacing="0" cellpadding="0" align="center" style="margin-top:-23px;"  >
 <% } else { %>
  <table   border="0" cellspacing="0" cellpadding="0" align="center" style="margin-top:0px;"  >
 <% } %>

  
 
   <tr>
    <td  align="right" valign="bottom"   class="style1" >       
    
    <table border="0" cellpadding="0" cellspacing="0" width="290" align="left" >
    <tr> <td nowrap align="right" valign="bottom" class="style3" ><img src="img/top-izq.png" border="0" align="bottom" height="20"  /></td>
    <td style="background-color:#6a6a6a; width:100%;"  valign="middle" align="center"   >
    
         <cc:ConfigGadgets id="ConfigGadgets" runat="server"   >
        </cc:ConfigGadgets>
   
        <span style="color:#CCC; font-family:Arial; font-size:8pt;"><%=Traducir_Fecha()%> </span>
     </td>
    </tr>
    </table>

    </td>
    
    <td   colspan="2" align="left" valign="bottom"style=" background-color:transparent;  " >
    
    
     <table border="0" cellpadding="0" cellspacing="0" width="100%" align="left">
    <tr> 
    <td  style="width:100%; background:url(img/top-centro.png) repeat-x  right;" align="left" valign="middle" >   
      
    </td>
    <td nowrap align="left" valign="bottom" >
    <img src="img/top-der.png" border="0" align="bottom"  height="20"/></td>    
    </tr>
    </table>
    
     
    </td>
  </tr>
  <tr>
    <td height="186" align="left" valign="top" style="background-color:#262626;  ">
      
      <cc:Modulos id="Modulos" runat="server"    >
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

  <td style="background:url(img/Contenedor-der.png) repeat-y right; background-color:#FFFFFF; width:2px" nowrap >    </td>
 
  </tr>
  <tr>
    
      
       <% if (bool.Parse(ConfigurationManager.AppSettings["VisualizarFooter"]))
          { %>
    <td colspan="3" style="height:80px; background-color:#262626">      
      <table width="100%"  border="0" cellspacing="0" cellpadding="0" style=" background:url(img/FooterRHPRO.png) no-repeat bottom right">
        <tr>
          <td width="48%">
          <p></p>
           <p class="TituloBase" >
              <b>Versión:</b>  <label runat="server" id="versionMI" /> <b>Patch:</b> <label runat="server" id="patchMI" /> 
           </p>
            </td>
          <td width="52%" rowspan="2" valign="top" align="right" style="height:70px" >
            <span class="Detalle" >Simplificamos su trabajo. Optimizamos su gestión.</span> 
            </td>
          </tr>
        <tr>
          <td>
            <p class="DetalleEmpresa">
              Heidt &amp; Asociados S.A. <br>
              Suipacha 72 - 4º A CP C1008AAB - Buenos Aires Argentina.<br />
              Tel./Fax: +54 11 5252 7300    Email: ventas@rhpro.com 
            </p></td>
          </tr>
      </table>
      <% }        else       {%>
      <td colspan="3" style="height:10px; background-color:#262626"> 
      <% } %>
      
      </td>
  </tr>
</table>


 




<script language="javascript">
 document.getElementById("TG").style.left =  ((AnchoPantalla()/2) - 350) + "px";
 document.getElementById("FondoTransparente").style.height =  window.document.body.scrollHeight + "px";
 
</script>
</body>
</asp:Content>


