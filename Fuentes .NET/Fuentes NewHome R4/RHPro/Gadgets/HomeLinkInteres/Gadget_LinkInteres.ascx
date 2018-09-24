<%@ Control Language="C#" AutoEventWireup="true" Inherits="HomeLinkInteres.Gadget_LinkInteres" CodeFile ="Gadget_LinkInteres.ascx.cs"        %>
<style>
 .IconoModuloLinkModulo
{
	 width:19px !important; height:24px  !important; 
	 vertical-align:middle; text-align:center; 	 
	 cursor:pointer !important;  
	 text-align:center;
	 border:0;
     padding:0px; padding-left:0px; margin-left:2px; margin-right:0px;     
}

 .IconoModuloLinkModulo_PNG	 
	 {
	 width:25px !important;  
	 vertical-align:middle; text-align:center; 	 
	 cursor:pointer !important;  
	 text-align:center;
	 border:0;
     padding:0px; padding-left:0px; margin-left:4px; margin-right:0px;     
	 }
	 
</style>	 
 
   <DIV class="ContenedorGadget_NOIMG">
     
      <asp:SqlDataSource
          id="SqlDataSource11"
          runat="server"
          DataSourceMode="DataSet"
          ProviderName="System.Data.OracleClient"
          >
      </asp:SqlDataSource>     
    
      <asp:Repeater ID="Repeater11" runat="server"   >
           <ItemTemplate>  
              <div style="text-align:left; margin-left:13px; margin-top:8px; margin-right:5px; background-color:#FFFFFF" >                        
			    <%#Common.Utils.Armar_Icono("img/Modulos/SVG/LINK.svg", "IconoModuloLinkModulo",""," align='absmiddle' ", "") %>
                  <span  style="color: #666666; font-family::Verdana, Geneva, sans-serif; font-size:10pt;" >
				   <%# Eval("Title") %> : 
				  </span>
				 <span style="color:#666666; font-size:9pt;" >     
                     <a   class="HiperLink" href="http://<%#Eval("Url")%>" target="_blank">  <%# Eval("Url")%></a>
                  </span>
              </div>     
                          
           </ItemTemplate>
      </asp:Repeater>
      
      
      
      
 </DIV>