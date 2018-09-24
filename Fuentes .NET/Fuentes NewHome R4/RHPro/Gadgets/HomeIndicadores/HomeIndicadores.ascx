<%@ Control Language="C#" AutoEventWireup="true" CodeFile="HomeIndicadores.ascx.cs" Inherits="Indicadores.HomeIndicadores"   %>

  
<% 
   if ( (String)Session["RHPRO_NombreModulo"]!="")
   {
 %>  
  
<DIV class="ContenedorGadget_NOIMG">
 
 <iframe frameborder=0 src="/rhprox2/ind/grafico_menu_ind_00_HOME.asp?menu=<%= (String)Session["RHPRO_NombreModulo"]%>" width="100%"  
 onload="AjustarIframe(this);" scrolling="no" >   </iframe>
</DIV>

 
 <%
 
 } 
 %>