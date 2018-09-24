<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="Idiomas.ascx.cs" Inherits="RHPro.Controls.Idiomas" %>


 

<style type="text/css">
#Globo_Idiomas{ visibility:hidden; float:left; z-index:1001; position: absolute; margin-left:-25px /* visibility:hidden; position:absolute; float:left;z-index:1000 */}
#Globo_Idiomas_Centro{background:url(img/Loguin/Globo_Centro.png) repeat-y center;} 
.Cerrar {background:url(img/Loguin/Globo_Centro.png) repeat-y center;} 

#Globo_Idiomas_Centro a{ text-decoration:none; font-family: Verdana; font-size:10pt; color:#333333;}
#Globo_Idiomas_Centro a:hover{ font-weight:bold;}
#Globo_Idiomas_Centro a:visited{  }


#Btn_Idioma{ cursor:pointer}
#Btn_Idioma:hover TD#idioma_izq{ background:url(img/Loguin/user_izq_hover.png) no-repeat right bottom;}	 
#Btn_Idioma:hover TD#idioma_centro1{ background:url(img/Loguin/user_centro_hover.png); color: #999}	
#Btn_Idioma:hover TD#idioma_centro2{ background:url(img/Loguin/user_centro_hover.png); color: #999}	
#Btn_Idioma:hover TD#idioma_der{ background:url(img/Loguin/user_der_hover.png) no-repeat right bottom;}

.idioma{ color:#333; font-family:Tahoma; font-size:9pt; }
#idioma_izq { background:url(img/Loguin/user_izq.png) no-repeat right bottom;}
#idioma_centro1 {background:url(img/Loguin/user_centro.png) repeat-x center}
#idioma_centro2 {background:url(img/Loguin/user_centro.png) repeat-x center}
#idioma_der { background:url(img/Loguin/user_der.png) no-repeat right bottom;}

</style>

<script type="text/javascript"> 
 
function Listar_Idiomas(){
     
	if (document.getElementById("Globo_Idiomas").style.visibility!="visible")
	    document.getElementById("Globo_Idiomas").style.visibility = "visible";
	 else
	 	document.getElementById("Globo_Idiomas").style.visibility = "hidden";	
		
}

</script>

 <!-- ##############################CONTENEDOR##################################----->
 <TABLE cellpadding="0" cellspacing="0" border="0">
 <TR> 
  <TD > 
 <!-- ---------------------- -->

 
  <table    border="0" cellspacing="0" cellpadding="0" class="idioma" id="Btn_Idioma" onclick="Abrir_Globo('Globo_Idiomas');" style="vertical-align:bottom">
      <tr>
        <td width="5" align="right" valign="middle" id="idioma_izq"><div style="width:5px; height:33px">&nbsp;</div></td>
        <td width="29"  id="idioma_centro1">
          <asp:Image runat="server" ImageUrl="~/img/Flags/flag_esES.png"  ID="Bandera" />         
           
        </td>
        <td width="145" align="center" nowrap="nowrap"  id="idioma_centro2"> 
        <span style="padding-left:3px; padding-right:3px">
        <asp:Label ID="Idioma" runat="server" >
        Español
        </asp:Label>
        </span></td>
        <td width="14" align="left" valign="middle" id="idioma_der"><div style="width:14px"></div></td>
      </tr>     
  </table>
     <% 
         /*
         if ((String)System.Web.HttpContext.Current.Session["Lenguaje"] != "")
         {
             RefrescarComboIdioma((String)System.Web.HttpContext.Current.Session["Lenguaje"], (String)System.Web.HttpContext.Current.Session["ArgTitulo"], (String)System.Web.HttpContext.Current.Session["ArgUrlImagen"]);
            
         }

         RefrescarComboIdioma("enUS", "Inglesito", "~/img/Flags/flag_enUS.png");
      */
          
         %>
     

  </TD>  
 </TR>
 <TR>
     <TD align="left" valign="bottom">  

       
     <!-- -- --------- IDIOMAS -------------------------------->
    <table width="250" border="0" cellspacing="0" cellpadding="0" id="Globo_Idiomas"  >    
    <tr>
      <td align="center" valign="bottom"><img src="img/Loguin/Globo_Tope.png" align="bottom" ></td>
    </tr>
    
    <tr>
      <td align="center" valign="bottom" class="Cerrar">
        <img src="img/Close.png" style=" margin-left:0px; cursor:pointer" onclick="Cerrar_Globo('Globo_Idiomas');" 
         onmouseover="this.src = 'img/Close-hover.png'" onmouseout="this.src = 'img/Close.png'" >
      </td>
    </tr>
    
    
    <tr>
      <td height="48" id="Globo_Idiomas_Centro" align="left" valign="top">
    
<!-- --------------------------ARMO EL COMBO DE IDIOMAS ACTIVOS-----------------------------------------------------------   -->      
   
     
      <asp:Repeater ID="Repeater1" runat="server"  >
           <ItemTemplate>                               
             <div onmouseover="this.style.background='url(img/ItemIdioma.png) repeat-y center';" onmouseout="this.style.background='';" 
               style="width:85%; text-align:left; margin-left:20px; font-size:10pt; color:#CCCCCC">            
           
               <img src="img/Flags/flag_<%# Eval("lencod") %>.png" align="absmiddle" border="0" style="margin-right:3px"   />    
                <asp:LinkButton     runat="server" OnClick="Idioma_Click"  
                 CommandArgument='<%# Eval("lencod") +"@"+ Eval("lendesabr") + "@~/img/Flags/flag_" + Eval("lencod") + ".png" %>' >
                  
                  <%# Eval("lendesabr") %>
               </asp:LinkButton>             
               
              </div>                      
                                        
           </ItemTemplate>
      </asp:Repeater>

<!-- --------------------------------------------------------------------------------------------------------------   -->
    
 
      </td>
    </tr>
    <tr>
     <td align="center" valign="top"><img src="img/Loguin/Globo_Piso.png" /></td>
   </tr>
  </table>
 

     </TD>
 </TR>
 </TABLE>
 
