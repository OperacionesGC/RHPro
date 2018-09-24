<%@ Control Language="C#" CodeFile ="Gadget_Mensajes.ascx.cs"  AutoEventWireup="true"  Inherits="HomeMensajes.Gadget_Mensajes"     %>
 
 
 
      <asp:SqlDataSource
          id="SqlDataSource1"
          runat="server"
          DataSourceMode="DataSet"  >
      </asp:SqlDataSource>     
      
      
      <asp:Repeater ID="Repeater1" runat="server"  DataSourceID="SqlDataSource1" Visible="false"  >
           <ItemTemplate>                               
             <div style="text-align:left; margin-left:13px; margin-top:5px; margin-right:5px;" >            
            <img src="Gadgets/HomeMensajes/comentario.png" align="baseline"> 
                  <span  style="color: #666666; font-family::Verdana, Geneva, sans-serif; font-size:9pt;" >
				  <b><%# Eval("hmsjtitulo") %> :</b>
				  </span>
				 <span  style="color:#666666; font-family::Verdana, Geneva, sans-serif; font-size:8pt;" >
                  <%# Eval("hmsjcuerpo")%>
                  </span>
              </div>     
              <br>                                                         
           </ItemTemplate>
      </asp:Repeater>
 
 