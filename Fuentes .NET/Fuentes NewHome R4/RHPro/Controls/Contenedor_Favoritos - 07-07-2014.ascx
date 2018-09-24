<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="Contenedor_Favoritos.ascx.cs" Inherits="RHPro.Controls.Contenedor_Favoritos" %>

 


 <!-- CONTROL DE PROGRESO -->
    <span style="margin-left:40px; display:inline-block; text-align:center; vertical-align:top" >      
    <asp:UpdateProgress ID="UpdateProgress" AssociatedUpdatePanelID="Update_Favoritos" runat="server"  Visible="true"  >
    <ProgressTemplate>   
    <DIV id="TransparenteProgress"  Class="PopUp_FondoTransparente" style="z-index:1000 !important;"></DIV>               
       <div style="padding:10px; background-color:transparent; text-align:center; position:absolute; left:45%;top:16%;z-index:1001; font-size:9pt; font-family:Tahoma; color:#fff">
           <img src="img/LOGO_RHPRO_loader.png" align="absmiddle"/>                                
           <div style="width:108px; height:46px; overflow:hidden">  <img src="img/miniloaderplano.gif" align="absmiddle"/>   </div>
       </div>    
    </ProgressTemplate>               
    </asp:UpdateProgress>       
    </span>
 <!-- ------------------------------------------- -->
 

              
 <asp:UpdatePanel ID="Update_Favoritos" runat="server" UpdateMode="Conditional">
    <ContentTemplate>         
         <DIV style="text-align:left;" >
          <asp:Panel ID="Contenedor_Fav" CssClass="Contenedor_Fav" runat="server"></asp:Panel>                             
          </DIV>
    </ContentTemplate>
    <Triggers>        
       <asp:AsyncPostBackTrigger ControlID="BtnUpdate" EventName="Click" /> 
    </Triggers>
 </asp:UpdatePanel>

 <DIV class="BarraPisoFavoritos">
     <asp:LinkButton ID="BtnUpdate" runat="server"   CssClass="Boton_Cuadrado BtnOpc"   OnClick="Refrescar"   Text="Refrescar"  > </asp:LinkButton>  
 </DIV>               
    
 <iframe src='' name='RHPROHome_Favorito_Iframe_Add' id='RHPROHome_Favorito_Iframe_Add' style='display:none; height:0px; width:0px; '></iframe>

 
 
 