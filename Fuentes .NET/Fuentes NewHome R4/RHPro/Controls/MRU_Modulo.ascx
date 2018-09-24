<%@ Control Language="C#" AutoEventWireup="true" CodeFile="MRUmi_Modulo.ascx.cs" Inherits="HomeMRU_Modulo.MRUmi_Modulo"   %>

 
 <script>
 function X_abrirVentana(url, name, width, height,opc) 
{   
  var str = "height=" + height + ",width=" + width;
  str += ",resizable=yes,status=no,toolbar=no,location=no"
   
  if (opc != null)
	  str += opc
  var auxi = "";
 
  if (url!=undefined)  {
    auxi = url.substr(url.lastIndexOf('/')+1,url.length);   
    auxi = auxi.substr(0,auxi.indexOf(".asp"));       
    window.open("../"+url,"",str);   
  }   
 
}
 </script>

<DIV class="ContenedorGadget_NOIMG">
	 
	<div id="mruImage"  runat="server"  >
	   <img src='Gadgets/HomeMRU/img/MRU.png'   />
	   <div style="position: relative; margin-top:-58px; margin-left:20px;  text-align:left; font-family:Arial; font-size:12pt; font-weight:bold; color:#999999;   ">
	 
			 <asp:Panel id="TituloDescriptivo" runat="server"></asp:Panel>
			
		  
		</div>
	</div>

	<div id="mruCompleto" runat="server" >

	<div id="cuerpo"   >

	   
		<span style="color:Green" runat="server" id="InfoPolitica"> </span>
		
		<asp:Repeater runat="server" ID="MRURepeater"     >
			<ItemTemplate> 
	  
				<div style="text-align:left; height:26px; margin-left:13px; margin-top:5px; margin-right:5px; background-color:#ffffff"  >            
				<img src="img/Modulos/SVG/FAVORITO.svg" border='0' class="IconoModulo"    >     
				 
				   
					   <span  style="color: #666666; font-family::Verdana, Geneva, sans-serif; font-size:9pt;" >
					 
					   <%# Traducir((String)Eval("menuname")) %> :  
					  </span>
					 <span  style="color:#666666; font-family::Verdana, Geneva, sans-serif; font-size:8pt;" >
			  
					 
	  <img  onclick="<%# Corregir((String)Eval("action")) %>" src="img/Modulos/SVG/MORE.svg" border='0' style="cursor: pointer;"class="IconoMRU" >    
				 
				   
					  </span>
				  </div>     
			 
				
				
			</ItemTemplate>
	 
		</asp:Repeater>
	 
	 
		
	</div>
	 
	</div>

</DIV>
 