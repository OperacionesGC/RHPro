<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="Contenedor_Estilos.ascx.cs" Inherits="RHPro.Controls.Contenedor_Estilos" %>
 
 
 <asp:UpdatePanel ID="Update_Estilos" runat="server"  >
    <ContentTemplate>  
    
 <DIV class='ContenidoControlMenuTop'>
      <asp:Repeater ID="IteradorEstilos" runat="server"  >
           <ItemTemplate>                               
                
                       
                 <asp:LinkButton runat="server" OnCommand="Estilo_Click"  id="BtnCambioEstilo" CommandArgument='<%# Convert.ToString(Eval("idcarpetaestilo")) +"@@"+ Convert.ToString(Eval("estilocarpeta")) +"@@"+ Convert.ToString(Eval("codestilo")) %>'>
                   <div class="EtiquetaIdioma"> <div class="RGBNombre">  <%# Eval("estdesabr") %></div>  <div class="RGBEstilo" style="background-color:<%# Eval("estiloRGB")%>"> </div></div>                     
               </asp:LinkButton>     
                                   
                                        
           </ItemTemplate>
      </asp:Repeater>
</DIV>

    </ContentTemplate>    
 </asp:UpdatePanel>


 
     


 