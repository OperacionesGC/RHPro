<%@ Control Language="C#" AutoEventWireup="true" ClassName="Gadget_Links" 
CodeBehind="Gadget_Links.ascx.cs" Inherits="RHPro.Gadgets.Gadget_Links"   %>
 
 
   <% RHPro.Gadget G = new RHPro.Gadget();%>
   <%  Response.Write(G.TopeModulo());  %>
      <asp:SqlDataSource
          id="SqlDataSource1"
          runat="server"
          DataSourceMode="DataSet"
          ConnectionString="Password=ess;Persist Security Info=True;User ID=ess;Initial Catalog=Base_0_R3_ARG;Data Source=RHDESA"
          SelectCommand="select * from home_link">
      </asp:SqlDataSource>     
      LINKS DE INTERES:<br>
      <asp:Repeater ID="Repeater1" runat="server"  DataSourceID="SqlDataSource1">
           <ItemTemplate>                               
             <div onmouseover="this.style.background='url(img/ItemIdioma.png) repeat-y center';" onmouseout="this.style.background='';" 
               style="width:85%; text-align:left; margin-left:20px; font-size:10pt; color:#CCCCCC">            
           
                 
                <asp:LinkButton ID="LinkButton1"     runat="server"  >
                  <%# Eval("hlinktitulo") + ":" + Eval("hlinkpagina")%>
               </asp:LinkButton>             
               
              </div>                      
                                        
           </ItemTemplate>
      </asp:Repeater>
<%  Response.Write(G.PisoModulo());  %>