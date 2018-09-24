<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="ConfigGadgets.ascx.cs" Inherits="RHPro.Controls.ConfigGadgets" %>


<style>
 .TopeG{ margin-left:4px; color:#333333; font-family:Arial; font-size:10pt; font-weight:bold; text-align:left}
 .LinksG{color:#333333; font-family:Arial; font-size:10pt; text-align:left; }
 .LinksG div{ margin-bottom:5px; }
 .TopeContenedorG{ font-family:Arial; font-size:12pt; font-weight:bold; color:#FFFFFF; background:url(img/Fondo_Menu_Press4.png) repeat-x center}
 .MiniG {background:url(img/FondoGadget1.png) no-repeat center; margin-top:10px; }
 .MiniG span { margin-left:5px;}

 .TG{ border:1px solid #333;top:30mm; position:fixed; visibility:hidden; zoom:1%}
 .Transparente {position:absolute; width:100%; height:100%; background-color:#333; opacity:0.8;filter:alpha(opacity=80);  top:0px; left:0px; visibility:hidden}
 
 .BtnG {margin-left:5px; cursor:pointer}
 .BtnG:hover { color:#FF0000; }
</style>

<script type="text/javascript" >
 
 

function CerrarModal(){
 document.getElementById("TG").style.visibility = "hidden";
 document.getElementById("FondoTransparente").style.visibility = "hidden";  
 document.getElementById("TG").style.zoom = "10%";
}

function AbrirModal(){
 var t;
 document.getElementById("FondoTransparente").style.visibility = "visible";  
 document.getElementById("TG").style.visibility = "visible";
 
 var zoom = document.getElementById("TG").style.zoom;
 zoom = parseInt(zoom.replace("%",""));
 if (zoom<100) {
	 document.getElementById("TG").style.zoom = (zoom + 10) + "%";	  
     t =  setTimeout("AbrirModal()",1);  	 
  }
 else 
      clearTimeout(t); 
}

//Envia por medio da variables GET el gadget a activar
function Activar(gadnro){
CerrarModal();
   document.getElementById("ifrm2").src = "~/../Config_Gadget.aspx?gadnro="+gadnro+"&sube=0&desactiva=0&activa=-1";  
   
}

</script>
 




<!-- Seccion oculta. Visualiza los controles -->
 <DIV  class="Transparente" id="FondoTransparente"></DIV>
<table width="700" border="0" cellspacing="0" cellpadding="0" align="center" class="TG" id="TG" style="zoom:10%">
  <tr>
    <td width="37" height="30" align="left" valign="bottom" class="TopeContenedorG"  > <img src="img/ConfigSombra.png"  /> </td>
    <td width="621" class="TopeContenedorG"><% Response.Write(Obj_Lenguaje.Label_Home("Gadgets")); %></td>
    <td width="42" class="TopeContenedorG" valign="middle" align="center"><img src="img/Cerrar.png"  style="cursor:pointer;" onclick="CerrarModal()" /> </td>
  </tr>
  
  <TR><TD bgcolor="#FFFFFF" valign="top" align="center" colspan="3">
  
  
  <DIV style="overflow-y:scroll; overflow-x: hidden; width:100%;  height:280px">   
  
  <!-- -------------------------Armo el manu con los modulos activos -----------   -->      
 
  <TABLE>
  
      
      <asp:Repeater ID="Repeater1" runat="server" >
           <HeaderTemplate> </HeaderTemplate>
           <ItemTemplate >                                                                
               
                <% 
                    if (ContadorGadget % 3 == 0)
                        Response.Write("<tr>");
                 %>   
                      <td nowrap="nowrap" style="width:1px; background-color:Transparent" valign="middle" >                          
                          <table width="177" height="110" border="0" cellspacing="0" cellpadding="0" class="MiniG">
                              <tr>
                                <td colspan="2" align="left"><span class="TopeG"><%# Obj_Lenguaje.Label_Home((String)Eval("gadtitulo"))%>  </span></td>
                              </tr>
                              <tr>
                                <td width="1"><img src="img/EngranageRojo.png" width="80" height="80" /></td>
                                <td class="LinksG" >
                                   <div style="margin-left:5px;"><img src="img/Gder.png" align="absmiddle" width="8" height="7"> Detalle</div>
                                   <div onclick="Activar(<%# Eval("gadnro")%>)" class="BtnG">
                                       <img src="img/Gder.png" align="absmiddle" width="8" height="7"> Activar
                                   </div>
                                   <div style="margin-left:5px;"><img src="img/Gder.png" align="absmiddle" width="8" height="7"> Modificar</div>
                                   <div style="margin-left:5px;"><img src="img/Gder.png" align="absmiddle" width="8" height="7"> Eliminar</div>               
                                </td>
                              </tr>
                            </table>                                           
                          
                      </td>  
                
           <%ContadorGadget = ContadorGadget + 1; %>
           <% 
                    if (ContadorGadget % 3 == 0)
                        Response.Write("</tr>");
                 %> 
                                                  
           </ItemTemplate>
      </asp:Repeater>
      </TABLE>
      
<%  
    
    int CantGadgets = Repeater1.Items.Count;  
    
%>

    
    </DIV>
  </TD>
  </TR>
  <tr>
    <td colspan="3"  class="TopeContenedorG">&nbsp;</td>
  </tr>
</table>
    
<!-- --------------------------------------- -->
 <iframe src="" id='ifrm2' name='ifrm2' style="visibility:hidden; height:0; width:0;" ></iframe>