<%@ Control Language="C#" AutoEventWireup="true" CodeFile="MRUmi.ascx.cs" Inherits="HomeMRU.MRUmi"   %>

 
 <script>
 function X_abrirVentana(url, name, width, height,opc) 
{   
  var str = "height=" + height + ",width=" + width;
  str += ",resizable=yes,status=no,toolbar=no,location=no"
   
  if (opc != null)
	  str += opc
  var auxi = "";
 
  if (url!=undefined)  {
    auxi = url.substr(url.lastIndexOf('/')+1,url.length);   
    auxi = auxi.substr(0,auxi.indexOf(".asp"));       
    window.open("../"+url,"",str);   
  }   
 
}
 </script>
 
<div id="mruImage"  runat="server"  >
   <img src='Gadgets/HomeMRU/img/MRU.png'   />
   <div style="position: relative; margin-top:-58px; margin-left:20px;  text-align:left; font-family:Arial; font-size:12pt; font-weight:bold; color:#999999;   ">
      <% Response.Write(Traducir("Referencias más utilizadas del sistema.")); %>  
    </div>
</div>

<div id="mruCompleto" runat="server" >

<div id="cuerpo"   >

   
 
    <asp:Repeater runat="server" ID="MRURepeater"     >
        <ItemTemplate> 
  
            <div style="text-align:left; height:26px; margin-left:13px; margin-top:5px; margin-right:5px; background-color:<% Response.Write(background());%>"  >            
            <img src="Gadgets/HomeMRU/img/favorito.png" align="baseline"> 
               
                   <span  style="color: #666666; font-family::Verdana, Geneva, sans-serif; font-size:9pt;" >
				 
                  <b><%# Traducir((String)Eval("menuname")) %> :</b>
				  </span>
				 <span  style="color:#666666; font-family::Verdana, Geneva, sans-serif; font-size:8pt;" >
          
                 <%# Traducir((String)Eval("Root")) %> :  <a  onclick="<%# Corregir((String)Eval("action")) %>" style="cursor: pointer;">   
                   <% Response.Write(Traducir("Acceder"));  %>  </a>
               
                  </span>
              </div>     
              
            
        </ItemTemplate>
 
    </asp:Repeater>
 
 
 <asp:SqlDataSource
          id="SqlDataSource1"
          runat="server"
          DataSourceMode="DataSet"
        
          SelectCommand="SELECT * FROM Gadgets" >
</asp:SqlDataSource>  
 
 
  <asp:Repeater runat="server" ID="XX1"     >
        <ItemTemplate> 
  
             <%# (String)Eval("gaddesabr") %>   
                   
            
        </ItemTemplate>
 
    </asp:Repeater>
    
</div>
 
</div>

 