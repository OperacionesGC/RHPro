<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="ConfigGadgets.ascx.cs" Inherits="RHPro.Controls.ConfigGadgets" %>


 
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
    //CerrarModal();
    Cerrar_PopUp_Generico('Contenedor_Gadgets');
   document.getElementById("ifrm2").src = "~/../Config_Gadget.aspx?gadnro="+gadnro+"&sube=0&desactiva=0&activa=-1";  
   
}

</script>
 


 <DIV class='Contenedor_Gadgets'>
  
      
      <asp:Repeater ID="Repeater1" runat="server" >
           <HeaderTemplate> </HeaderTemplate>
           <ItemTemplate >                                                                
               
                  
                    <div class="DIV_Gadget_Config"  >                 
                         <div class="TopeG"><%# Obj_Lenguaje.Label_Home((String)Eval("gadtitulo"))%>  </div>                          
                            <div style="margen-left:8px; padding:8px">   
                               
                                   <div class="BtnG" ><img src="img/Gder.png" align="absmiddle" width="8" height="7"> Detalle</div>
                                   <div onclick="Activar(<%# Eval("gadnro")%>)" class="BtnG">
                                       <img src="img/Gder.png" align="absmiddle" width="8" height="7"> Activar
                                   </div>
                                   <div class="BtnG" ><img src="img/Gder.png" align="absmiddle" width="8" height="7"> Modificar</div>
                                   <div class="BtnG"><img src="img/Gder.png" align="absmiddle" width="8" height="7"> Eliminar</div>               
                           </div>
                                
                      </div>  
                
           <%ContadorGadget = ContadorGadget + 1; %>
           
                                                  
           </ItemTemplate>
      </asp:Repeater>
 
      </DIV>
<%  
    
    int CantGadgets = Repeater1.Items.Count;  
    
%>

     
    
<!-- --------------------------------------- -->
 <iframe src="" id='ifrm2' name='ifrm2' style="visibility:hidden; height:0px; width:0px;" ></iframe>