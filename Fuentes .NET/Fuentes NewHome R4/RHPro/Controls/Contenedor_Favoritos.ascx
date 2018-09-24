<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="Contenedor_Favoritos.ascx.cs" Inherits="RHPro.Controls.Contenedor_Favoritos" %>

 

 <asp:UpdatePanel ID="Update_Favoritos" runat="server"  >
    <ContentTemplate>         
   
         <DIV style="text-align:left;" >
            <asp:Panel ID="Contenedor_Fav" CssClass="Contenedor_Fav" runat="server"></asp:Panel>                             
          </DIV>
          
          <DIV style="margin:0;padding:0">
          <table style="padding:0 !important; margin:0 !important; width:100%"> <tr class="PopUp_Piso"><td align="right">
          <span style="float:right">
             <asp:Button ID="BtnUpdate" runat="server"    CssClass="Boton_Cuadrado BtnOpc"   OnClick="Refrescar" Text="Refrescar" />
          </span>
          </td></tr>
          </table>
          </DIV>
    </ContentTemplate>
    <Triggers>        
       <asp:AsyncPostBackTrigger ControlID="BtnUpdate" EventName="Click" /> 
    </Triggers>
 </asp:UpdatePanel>


       
    
 <iframe src='' name='RHPROHome_Favorito_Iframe_Add' id='RHPROHome_Favorito_Iframe_Add' style='display:none; height:0px; width:0px; '></iframe>

 
 
 