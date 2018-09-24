<%@ Control Language="C#" CodeFile ="Gadget_Mensajes.ascx.cs"  AutoEventWireup="true"  Inherits="HomeMensajes.Gadget_Mensajes"     %>
 
<DIV class="ContenedorGadget_NOIMG"> 
      <asp:SqlDataSource
          id="SqlDataSource1"
          runat="server"
          DataSourceMode="DataSet"           
          >
      </asp:SqlDataSource>     
      
      
      <asp:Repeater ID="Repeater1" runat="server"  >
           <ItemTemplate>                               
             <div style="text-align:left; margin-left:8px; margin-top:0px; margin-right:5px;   line-height:18px; width:100%" >            
            
			  
				   <%#Common.Utils.Armar_Icono("img/Modulos/SVG/MENSAJES.svg", "IconoModulo",""," border='0' ", "") %>
                  <span  style="color: #666666; " >
				  <%# Eval("Title") %> :                   
				  </span>
				 <span  style="color:#666666;  " >
                  <%# Eval("Body")%>
                  </span>
              </div>     
                                                                       
           </ItemTemplate>
      </asp:Repeater>
 </DIV>
  
