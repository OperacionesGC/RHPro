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
function ActivarGadget(gadnro){
    //CerrarModal();
    Cerrar_PopUp_Generico('Contenedor_Gadgets');
   document.getElementById("ifrm2").src = "~/../Config_Gadget.aspx?gadnro="+gadnro+"&sube=0&desactiva=0&activa=-1";  
   
}

//Envia por medio da variables GET el gadget a activar
function DesactivarGadget(gadnro) {
   
    Cerrar_PopUp_Generico('Contenedor_Gadgets');
    document.getElementById("ifrm2").src = "~/../Config_Gadget.aspx?gadnro=" + gadnro + "&sube=0&desactiva=-1&activa=0";

}

</script>
 
  


 <asp:UpdatePanel ID="Update_Gadgets" runat="server"   >
    <ContentTemplate>  
 <DIV class='Contenedor_Gadgets'>
  
      
      <asp:Repeater ID="Repeater1" runat="server" >
           <HeaderTemplate> </HeaderTemplate>
           <ItemTemplate >                                                                             
       
               
                   <div class="EtiquetaGadgets" > 
                      <div class="GadgetNombre" title="<%# Obj_Lenguaje.Label_Home((String)Eval("gadtitulo"))%> ">
                         <%#Imprimir_Led(Convert.ToInt32(Eval("gadusractivo")))%>            
                       <%# Obj_Lenguaje.Label_Home((String)Eval("gadtitulo"))%> </div>                        
                      
                       <%#Imprimir_Slider(Convert.ToInt32(Eval("gadusrnro")), Convert.ToInt32(Eval("gadusractivo")))%>                                          
                   </div>                     
                            
           <%ContadorGadget = ContadorGadget + 1; %>           
                                                  
           </ItemTemplate>
      </asp:Repeater>
 
      </DIV>
    </ContentTemplate>    
 </asp:UpdatePanel>      
<%  
    
    int CantGadgets = Repeater1.Items.Count;  
    
%>

     
    
<!-- --------------------------------------- -->
 <iframe src="" id='ifrm2' name='ifrm2' style="visibility:hidden; height:0px; width:0px;" ></iframe>