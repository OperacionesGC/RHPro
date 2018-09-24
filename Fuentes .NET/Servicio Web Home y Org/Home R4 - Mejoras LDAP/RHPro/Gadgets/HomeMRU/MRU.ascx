<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="MRU.ascx.cs" Inherits="RHPro.Controls.MRU" %>


  <% 
   RHPro.Lenguaje ObjLenguaje = new RHPro.Lenguaje();
   RHPro.Gadget G = new RHPro.Gadget();    
   %>
<%  Response.Write(G.TopeModulo(ObjLenguaje.Label_Home("Accesos más usados"),"350"));  %>

<div id="mruImage"  runat="server"  ></div>
<div id="mruCompleto" runat="server" >
 
 
<div id="cuerpo"  >


 
<asp:SqlDataSource
          id="SqlDataSource1"
          runat="server"
          DataSourceMode="DataSet"
          ConnectionString="Password=ess;Persist Security Info=True;User ID=ess;Initial Catalog=Base_0_R3_ARG;Data Source=RHDESA"
          SelectCommand="SELECT menumstr.menuname, menumstr.action, menuraiz.menunombre Root, menuraiz.menudir, menumstr.menuaccess FROM mru INNER JOIN menumstr ON menumstr.menumsnro = mru.menumsnro  INNER JOIN menuraiz ON menuraiz.menunro = mru.menuraiz  WHERE UPPER(mru.iduser) = 'rhpror3'  ORDER BY mrufecha DESC, mruhora DESC" >
</asp:SqlDataSource>     

    <asp:Repeater runat="server" ID="MRURepeater"  >
        <ItemTemplate> 
            
             <div style="text-align:left; margin-left:13px; margin-top:5px; margin-right:5px;" >            
            <img src="Gadgets/HomeMRU/img/favorito.png" align="baseline"> 
                  <span  style="color: #666666; font-family::Verdana, Geneva, sans-serif; font-size:9pt;" >
				  <b><%# Eval("menuname") %> :</b>
				  </span>
				 <span  style="color:#666666; font-family::Verdana, Geneva, sans-serif; font-size:8pt;" >
          
                 <%# Eval("Root") %> :  <a onclick="<%# Eval("action") %>" style="cursor: pointer;">acceder</a>
                 
                  </span>
              </div>     
              <br>        
            
        </ItemTemplate>
        <SeparatorTemplate>
            <div class="separador">
            </div>
        </SeparatorTemplate>
    </asp:Repeater>
</div>
<div id="inferior">
</div>
</div>

<% Response.Write(G.PisoModulo());  %>