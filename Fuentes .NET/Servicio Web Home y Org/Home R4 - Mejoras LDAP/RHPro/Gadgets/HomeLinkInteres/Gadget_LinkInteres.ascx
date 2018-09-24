<%@ Control Language="C#" AutoEventWireup="true" ClassName="Gadget_LinkInteres"  CodeBehind="Gadget_LinkInteres.ascx.cs"   %>
 
 
    <% 
   RHPro.Lenguaje ObjLenguaje = new RHPro.Lenguaje();
   RHPro.Gadget G = new RHPro.Gadget();    
   %>
   
      <%  Response.Write(G.TopeModulo(ObjLenguaje.Label_Home("Link de Interes"),"350"));  %>
      <asp:SqlDataSource
          id="SqlDataSource11"
          runat="server"
          DataSourceMode="DataSet"
          ConnectionString="Password=ess;Persist Security Info=True;User ID=ess;Initial Catalog=Base_0_R3_ARG;Data Source=RHDESA"
          SelectCommand="select * from home_link">
      </asp:SqlDataSource>     
    
      <asp:Repeater ID="Repeater11" runat="server"  DataSourceID="SqlDataSource11" >
           <ItemTemplate>  
              <div style="text-align:left; margin-left:13px; margin-top:8px; margin-right:5px; background-color:#FFFFFF" >            
               <img src="Gadgets/HomeLinkInteres/link.gif" align="absmiddle"> 
                  <span  style="color: #666666; font-family::Verdana, Geneva, sans-serif; font-size:10pt;" >
				  <b><%# Eval("hlinktitulo") %> :</b>
				  </span>
				 <span  style="color:#666666; font-family::Verdana, Geneva, sans-serif; font-size:10pt;" >     
                     <a href="http://<%#Eval("hlinkpagina")%>" target="_blank">  <%# Eval("hlinkpagina")%></a>
                  </span>
              </div>     
                          
           </ItemTemplate>
      </asp:Repeater>
<%  Response.Write(G.PisoModulo());  %>