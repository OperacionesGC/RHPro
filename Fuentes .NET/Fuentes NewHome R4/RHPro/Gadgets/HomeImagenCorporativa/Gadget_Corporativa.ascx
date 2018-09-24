<%@ Control Language="C#" AutoEventWireup="true" CodeFile="Gadget_Corporativa.ascx.cs" Inherits="HomeGadget_Corporativa.Gadget_Corporativa"   %>
 <asp:PlaceHolder id="ScriptJS" runat="server"></asp:PlaceHolder > 
 <asp:Panel id="ImagenCorpPais" runat="server"></asp:Panel>
  
 	<link rel="stylesheet" type="text/css" href="css/GaleriaBanners.css" /> 
	<script src="js/GaleriaBanners.js"></script>
	
<div id="slideshow">
	<ul class="slides"  >
	   <asp:PlaceHolder id="BannersCorp" runat="server"></asp:PlaceHolder>
    	
    </ul>

    <span class="arrow previous"></span>
    <span class="arrow next"></span>
</div>
