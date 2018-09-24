<%@ Control Language="C#" AutoEventWireup="true" ClassName="Gadget_Mensajes"  CodeBehind="Gadget_Mensajes.ascx.cs"   %>
 
 
   <% 
   RHPro.Lenguaje ObjLenguaje = new RHPro.Lenguaje();
   RHPro.Gadget G = new RHPro.Gadget();    
   %>
   
   
   <%  Response.Write(G.TopeModulo(ObjLenguaje.Label_Home("Mensajes"),"350"));  %>
      <asp:SqlDataSource
          id="SqlDataSource1"
          runat="server"
          DataSourceMode="DataSet"
          ConnectionString="Password=ess;Persist Security Info=True;User ID=ess;Initial Catalog=Base_0_R3_ARG;Data Source=RHDESA"
          SelectCommand="SELECT  * FROM home_mensaje WHERE rhpro=-1  ORDER BY hmsjfecalta DESC">
      </asp:SqlDataSource>     
      
      <asp:Repeater ID="Repeater1" runat="server"  DataSourceID="SqlDataSource1">
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
<%  Response.Write(G.PisoModulo());  %>