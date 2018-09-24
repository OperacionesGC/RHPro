 
 

<style>
 

.Separador { color:#FFF; font-family:Arial; font-size:11pt; background:url(img/Fondo_Menu.png) repeat-x top; height:37px; cursor:pointer }
.Menu_Ppal TR { color:#FFF; font-family:Arial; font-size:11pt; background:url(img/Fondo_Menu.png) repeat-x top; height:37px; cursor:pointer }
.Menu_Links TR { color:#FFF; font-family:Arial; font-size:9pt; background:url(img/Fondo_Menu.png) repeat-x top; height:37px; cursor:pointer }

.DetalleEmpresa{
	 color:#CCCCCC; font-size:7.5pt; font-family:Arial; margin-top:4px;
	} 
.TituloBase{ color:#FFFFFF; font-size:7pt; font-family:Arial;}

.Detalle{ color:#CCC; font-size:12pt; font-family:Arial;  }

.user{ color:#333; font-family:Arial; font-size:11pt; font-weight:bold;}
.idioma{ color:#333; font-family:Arial; font-size:9pt; }
</style>

<script>

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
 

</script>

<body style="background:url(img/fondo.png)" bottommargin="0" topmargin="0" leftmargin="0" rightmargin="0">
<table width="955" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td width="409" align="left" valign="middle" style="height:90px"> <img src="img/logoRHPRO.png"> </td>
    <td width="218">
    
    <table width="90" border="0" cellspacing="0" cellpadding="0" class="user">
      <tr>
        <td width="6" align="right" valign="middle"><img src="img/user_izq.png" border="0"></td>
         <td width="4" style="background:url(img/user_centro.png) repeat-x center"> </td>
        <td width="4" style="background:url(img/user_centro.png) repeat-x center">rhpror3</td>
        <td width="14" align="left" valign="middle"><img src="img/user_der.png" border="0"></td>
      </tr>
    </table>
    
    </td>
    <td width="328"><table   border="0" cellspacing="0" cellpadding="0" class="idioma">
      <tr>
        <td width="5" align="right" valign="middle"><img src="img/user_izq.png" border="0" /></td>
        <td width="34" style="background:url(img/user_centro.png) repeat-x center"><img src="img/banderas/argentina.png" border="0" align="absmiddle" /> </td>
        <td width="140" style="background:url(img/user_centro.png) repeat-x center" nowrap="nowrap">
            Español (Argentina)</td>
        <td width="14" align="left" valign="middle"><img src="img/user_der.png" border="0" /></td>
      </tr>
    </table></td>
  </tr>
</table>
<table   border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td style="width:199px; height:53px; background:url(img/top_izq.png) no-repeat right bottom" >
        &nbsp;</td>
    <td align="right" valign="middle" style="background:url(img/top_der.png) no-repeat left bottom; width:756px;" >
        &nbsp;</td>
  </tr>
  <tr>
    <td height="186" align="left" valign="top">
    <!-- MODULOS ---------------------------->
     <table width="100%" border="0" cellspacing="0" cellpadding="0" class="Separador">
      <tr>
        <td  valign="middle" align="right"nowrap="nowrap" colspan="3" style="color:#999; font-family:Arial; font-size:8pt;">
         <span  style="margin-right:5px; vertical-align:middle"> Modulos <img src="img/flecha.png" border='0' align="middle" /></span>
         </td>
         
      </tr>
    </table>
    <table width="100%" border="0" cellspacing="0" cellpadding="0" class="Menu_Links">
      <tr onclick="InfoLink('Link1','adp')" id='Link1' onmouseover="Sobre(this)" onmouseout="Sale(this)">
        <td nowrap="nowrap"><img src="img/adp.png" border='0' align="absmiddle" style="margin-left: 4px;"></td>
        <td style="width:100%"><span style="margin-left:3px;"> Adm. de Personal</span></td>
        <td><img src="img/plus.png" border='0' align="absmiddle" onmouseover="this.src = 'img/plus_hover.png'"  onmouseout="this.src = 'img/plus.png'" 
        onclick="AbrirLink('adp')"></td>
      </tr>
      
       <tr onclick="InfoLink('Link2','adp')" id='Link2' onmouseover="Sobre(this)" onmouseout="Sale(this)"> 
         <td nowrap="nowrap"><img src="img/adp.png" border='0' align="absmiddle" style="margin-left: 4px;"></td>
        <td><span style="margin-left:3px;">Liquidación de Haberes</span></td>
        <td><img src="img/plus.png" border='0' align="absmiddle" onmouseover="this.src = 'img/plus_hover.png'"  onmouseout="this.src = 'img/plus.png'" onclick="AbrirLink('adp')"></td>
      </tr>
      
       <tr onclick="InfoLink('Link3','adp')" id='Link3' onmouseover="Sobre(this)" onmouseout="Sale(this)">
         <td nowrap="nowrap"><img src="img/adp.png" border='0' align="absmiddle" style="margin-left: 4px;"></td>
        <td ><span style="margin-left:3px;">Gestión de Tiempos</span> </td>
        <td><img src="img/plus.png" border='0' align="absmiddle" onmouseover="this.src = 'img/plus_hover.png'"  onmouseout="this.src = 'img/plus.png'" onclick="AbrirLink('adp')"></td>
      </tr>
    </table>
    
    <table width="100%" border="0" cellspacing="0" cellpadding="0" class="Separador">
      <tr>
        <td  valign="middle" align="right"nowrap="nowrap" colspan="3" style="color:#999; font-family:Arial; font-size:8pt;"><span  style="margin-right:5px; vertical-align:middle"> 
            Accesos <img src="img/flecha.png" border='0' align="middle" /></span></td>
      </tr>
    </table>
    
    
   
    <!-- ------------------------- MENUES  -->
      
    <table width="100%" border="0" cellspacing="0" cellpadding="0" class="Menu_Ppal">
      <tr id="Link4" onclick="Seleccionar('Link4','adp')"   onmouseover="Sobre(this)" onmouseout="Sale(this)">
        <td nowrap="nowrap"> <img src="img/link.png" border='0' align="absmiddle" style="margin-left: 4px;"/></td>
        <td style="width:100%"><span style="margin-left:3px;"> Sitio</span></td>
        <td>&nbsp;</td>
      </tr>
      <tr id="Link5" onclick="Seleccionar('Link5','adp')"   onmouseover="Sobre(this)" onmouseout="Sale(this)">
        <td nowrap="nowrap"><img src="img/link.png" border='0' align="absmiddle"  style="margin-left: 4px;"/></td>
        <td><span style="margin-left:3px;"> E-recruiting</span></td>
        <td>&nbsp;</td>
      </tr>
       <tr id="Link6" onclick="Seleccionar('Link6','adp')"   onmouseover="Sobre(this)" onmouseout="Sale(this)">
        <td nowrap="nowrap"><img src="img/link.png" border='0' align="absmiddle"  style="margin-left: 4px;"/></td>
        <td><span style="margin-left:3px;"> E-learning</span></td>
        <td>&nbsp;</td>        
      </tr>
       <tr id="Link7" onclick="Seleccionar('Link7','adp')"   onmouseover="Sobre(this)" onmouseout="Sale(this)">
        <td nowrap="nowrap"><img src="img/link.png" border='0' align="absmiddle"  style="margin-left: 4px;"/></td>
        <td><span style="margin-left:3px;"> CRM</span></td>
        <td>&nbsp;</td>
      </tr>
       <tr id="Link8" onclick="Seleccionar('Link8','adp')"   onmouseover="Sobre(this)" onmouseout="Sale(this)">
        <td nowrap="nowrap"><img src="img/link.png" border='0' align="absmiddle"  style="margin-left: 4px;"/></td>
        <td><span style="margin-left:3px;">Patch</span></td>
        <td>&nbsp;</td>
      </tr>     
    </table>
  
  
   
    <!-- ------------------------- MENUES  -->
    
    </td>
    <td style="background:url(img/FondoContenedor.png) repeat-y center" align="center" 
                valign="top">
                
                <cc:CustomLogin id="cLogin" runat="server">
                    </cc:CustomLogin>
                
                </td>
  </tr>
  <tr>
    <td colspan="2" style="height:80px; background-color:#262626">
    
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="48%"><span class="TituloBase" >
            Versión:Base 0 ARG - R3 Multidioma Patch:Update Mejora Nº 0000097</span>
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
            Tel./Fax: +54 11 5252 7300 Email: ventas@rhpro.com 
          </span></td>
      </tr>
    </table></td>
  </tr>
</table>
<br>
</body>
 

  