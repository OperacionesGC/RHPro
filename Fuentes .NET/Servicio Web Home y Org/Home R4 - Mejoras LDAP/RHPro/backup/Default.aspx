<%@ Page Title="" Language="C#" MasterPageFile="~/MasterPage.Master" AutoEventWireup="true" 
CodeBehind="Default.aspx.cs" Inherits="RHPro.Default" EnableViewState="true"  %>

 
 
 
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
</asp:Content>
 


<asp:Content ID="Content2" ContentPlaceHolderID="content" runat="server" >


    <div id="contenido" style="margin-top:0px;"  >   
        </div>
       
    
<!-- -------------------------------------------------------------------->
 
     
<style>
 

.Separador { color:#FFF; font-family:Arial; font-size:11pt; background:url(img/Fondo_Menu.png) repeat-x top; height:37px; cursor:pointer }
.Menu_Ppal TR { color:#FFF; font-family:Arial; font-size:11pt; background:url(img/Fondo_Menu.png) repeat-x top; height:37px; cursor:pointer }
.Menu_Links TR { color:#FFF; font-family:Arial; font-size:9pt; background:url(img/Fondo_Menu.png) repeat-x top; height:37px; cursor:pointer }

.DetalleEmpresa{
	 color:#CCCCCC; font-size:7.5pt; font-family:Arial; margin-top:4px;
	} 
.TituloBase{ color:#FFFFFF; font-size:7pt; font-family:Arial;}

.Detalle{ color:#CCC; font-size:12pt; font-family:Arial;}

.user{ color:#333; font-family:Arial; font-size:11pt; font-weight:bold;}

#user_izq { background:url(img/Loguin/user_izq.png) no-repeat right bottom;}
#user_centro1 {background:url(img/Loguin/user_centro.png) repeat-x center}
#user_centro2 {background:url(img/Loguin/user_centro.png) repeat-x center}
#user_der { background:url(img/Loguin/user_der.png) no-repeat right bottom;}

#idioma_izq { background:url(img/Loguin/user_izq.png) no-repeat right bottom;}
#idioma_centro1 {background:url(img/Loguin/user_centro.png) repeat-x center}
#idioma_centro2 {background:url(img/Loguin/user_centro.png) repeat-x center}
#idioma_der { background:url(img/Loguin/user_der.png) no-repeat right bottom;}
 
 
#Btn_Loguin{ cursor:pointer}
#Btn_Loguin:hover TD#user_izq{ background:url(img/Loguin/user_izq_hover.png) no-repeat right bottom;}	 
#Btn_Loguin:hover TD#user_centro1{ background:url(img/Loguin/user_centro_hover.png); color:  #999}	
#Btn_Loguin:hover TD#user_centro2{ background:url(img/Loguin/user_centro_hover.png); color: #999}	
#Btn_Loguin:hover TD#user_der{ background:url(img/Loguin/user_der_hover.png) no-repeat right bottom;}

#Btn_Idioma{ cursor:pointer}
#Btn_Idioma:hover TD#idioma_izq{ background:url(img/Loguin/user_izq_hover.png) no-repeat right bottom;}	 
#Btn_Idioma:hover TD#idioma_centro1{ background:url(img/Loguin/user_centro_hover.png); color: #999}	
#Btn_Idioma:hover TD#idioma_centro2{ background:url(img/Loguin/user_centro_hover.png); color: #999}	
#Btn_Idioma:hover TD#idioma_der{ background:url(img/Loguin/user_der_hover.png) no-repeat right bottom;}

.idioma{ color:#333; font-family:Arial; font-size:9pt; }
.fecha{ color:#333; font-family:Arial; font-size:9pt; }

Body
{
 /*EXPLORER:  */  /* filter:progid:DXImageTransform.Microsoft.Gradient(endColorstr='#F3F4F6', startColorstr='#AAAEBB', gradientType='0');   */
/*SAFARI y CHROME:  */ /* background: -webkit-gradient(linear, left top, left bottom, from(#AAAEBB), to(#F3F4F6));*/
/*MOZILLA: */ /*  background: -moz-linear-gradient(top, #AAAEBB,#F3F4F6);*/
/*OPERA*/ /* background: -o-linear-gradient(#F3F4F6, #AAAEBB); */
 background:url(img/fondo.png) repeat;

}
</style>

<script type="text/javascript">

var activo = 0;
 
var CantidadLink = 8;
var LinkSeleccionado = "";
 

/*------Link-----*/	
function AbrirLink(Link){
	alert(Link);
}

function Seleccionar(id,Link){
  for (i=1; i<=CantidadLink; i++) {	 
	document.getElementById("Link"+i).style.background = "url(img/Fondo_Menu.png)"; 
   } 
   LinkSeleccionado = id;
   document.getElementById(id).style.background = "url(img/Fondo_Menu_Press.png)"; 
 }

function InfoLink(id,Link) {  
  Seleccionar(id,Link);
}

function Sobre(obj){   
	  obj.style.background = "url(img/Fondo_Menu_Press.png)"; 	 
	}
	
function Sale(obj){
   if (LinkSeleccionado!=obj.id)
	  obj.style.background = "url(img/Fondo_Menu.png)";
	}
 function Abrir_Globo(){
	if (document.getElementById("Globo_Idiomas").style.visibility=="visible")
	 document.getElementById("Globo_Idiomas").style.visibility = "hidden";
	
	if (document.getElementById("Globo_Loguin").style.visibility!="visible")
	 document.getElementById("Globo_Loguin").style.visibility = "visible";
	 else
	 	 document.getElementById("Globo_Loguin").style.visibility = "hidden";
}

 function Listar_Idiomas(){
    if (document.getElementById("Globo_Loguin").style.visibility=="visible")
	   document.getElementById("Globo_Loguin").style.visibility = "hidden";
	 
 
	if (document.getElementById("Globo_Idiomas").style.visibility!="visible")
	 document.getElementById("Globo_Idiomas").style.visibility = "visible";
	 else
	 	 document.getElementById("Globo_Idiomas").style.visibility = "hidden";}

</script>

 
 

<table width="955" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td width="544" align="left" valign="middle" style="height:90px" rowspan="2"> <img src="img/logoRHPRO.png"> </td>
    <td width="185" align="right" valign="middle"  >
    
    <!-- ACA BORRE LOGUIN-->
    <div id="login">
         <cc:CustomLogin id="cLogin" runat="server" ContentPlaceHolderID="cLogin"   >
        </cc:CustomLogin>
    </div>
    
   
    
    </td>
    <td width="226" align="right" valign="middle">   
     <!-- -- --------- IDIOMAS -------------------------------->
         <cc:Idiomas  id="cIdiomas" runat="server"  >
         </cc:Idiomas>
     <!-- -------------------------------------- -->   
      
  
    
    </td>
  </tr>
     <tr>
     <td></td>
     <td valign="bottom" align="right" style="height:10px" class="fecha">
     <%=DateTime.Now.ToLongDateString() %>
     </td>
  </tr>
</table>
<table   border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td style="width:199px; height:53px; background:url(img/top_izq.png) no-repeat right bottom" >&nbsp;</td>
    <td align="right" valign="middle" style="background:url(img/top_der.png) no-repeat left bottom; width:756px;" >
    <div style="float:left;">
            <ul class="menu_list">
            <div id="botones">
                <asp:Repeater runat="server" ID="menuRepeater" DataSourceID="menuDataSource">
                    <ItemTemplate>
                        <li><a href="<%# Eval("url") %>" target="_blank">
                            <%# Eval("title") %>
                        </a></li>
                    </ItemTemplate>
                </asp:Repeater>
             </div>
           </ul>
         </div>
    </td>
  </tr>
  <tr>
    <td height="186" align="left" valign="top">
    
    
    
    
    
  
  
   
    <!-- ------------------------- MENUES  -->
    
    </td>
    <td style="background:url(img/FondoContenedor.png) repeat-y center" align="center" valign="top">
     
     
             
         
        <div id="centralArriba">
                    <cc:Modules Id="mlsMain" runat="server">
                    </cc:Modules>
        </div>
    
    </td>
  </tr>
  <tr>
    <td colspan="2" style="height:80px; background-color:#262626">
    
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="48%"><span class="TituloBase" >
        Versión:Base 0 ARG - R3 Multidioma   Patch:Update Mejora Nº 0000097</span>
        </td>
        <td width="52%" rowspan="2" valign="top" align="right">
        <span class="Detalle" >Simplificamos su trabajo. Optimizamos su gestión.</span> 
        </td>
      </tr>
      <tr>
        <td>
         <span class="DetalleEmpresa">
         Heidt &amp; Asociados S.A. <br>
          Suipacha 72 - 4º A CP C1008AAB - Buenos Aires Argentina.<br />
          Tel./Fax: +54 11 5252 7300    Email: ventas@rhpro.com 
          </span></td>
      </tr>
    </table></td>
  </tr>
</table>

 

</asp:Content>

