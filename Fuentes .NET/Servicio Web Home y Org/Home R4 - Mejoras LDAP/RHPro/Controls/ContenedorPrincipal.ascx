<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="ContenedorPrincipal.ascx.cs" 
Inherits="RHPro.Controls.ContenedorPrincipal" %>
 
<script language="javascript">
function abrirVentana(url, name, width, height,opc) 
{ 
 
 /*
  var str = "height=" + height + ",innerHeight=" + height;
  str += ",width=" + width + ",innerWidth=" + width;
  */
  var str = "height=" + height + ",width=" + width;
   
  //str += ",resizable=yes"
   str += ",resizable=yes,status=no,toolbar=no,location=no"
   
  if (name=="ESS")
    str += ",scrollbars=1,menubar=no";
  else
    str += ",scrollbars=no,menubar=no";
   
  
  if (opc != null)
	  str += opc
  var auxi = "";
 
   if (url!=undefined)  {
    auxi = url.substr(url.lastIndexOf('/')+1,url.length);
   
    auxi = auxi.substr(0,auxi.indexOf(".asp"));  
      
    window.open(url,"",str);   
  }   
 
  //window.open(url, auxi, str); 
 


}



function AbrirLink(Link,menuname){ 

	 if (menuname=="ESS") //En el caso que se desea abrir el modulo de ESS, el tamaño de ventana va a ser mayor
 	abrirVentana(Link,"ESS",AnchoPantalla()-100,AltoPantalla()-150,"");
  else
     abrirVentana(Link,"",700,500,"");
}
</script>

 

<script type="text/javascript" src="~/../Js/Drag.js"></script>

<style>
    
    
.BordeGris{ border:1px #999999 solid; margin-top:4px; margin-bottom:5px;   margin-left:0px;   }
.PisoGris{ border-bottom:1px #999999 solid; background-color:#f0f0f0; height:30px; text-align:left; color:#666666; font-family:Arial; width:100%;  }
.TopeGris{ border-top:1px #999999 solid; background-color:#f0f0f0; height:0px}
.ContenedorGadget{overflow-y:scroll; height:200px; width:350px; color:#333333; font-family:Arial; font-size:9pt;}
.ContenedorGadget a{color:#811e1e; font-family:Arial; font-size:9pt; text-decoration:none}
.ContenedorGadget a:hover{color:#000000;}

	 .InfoModulos {
		 font-family:Arial;
		 font-size:10pt;
		 color:#333;
		 background-color: transparent;
		 width:600px;
		 text-align:justify;
		 margin-top:9pt;
		 
		 border:6px solid #CCCCCC;
		 background-color:#FFFFFF; 
		 padding:4px;
		 margin-left:12px; 		 	
		//margin-left:0px; 		
		 }
		 
	.TopeInfoModulos {
		 font-family:Arial;
		 font-size:11pt;
		 color:#333;
		 background-color: transparent; 
		 background: url(img/Modulos/TopeDescripcionModulos.png) no-repeat top center;
		 margin-top:8px;
		}	 
.tooltiphelp {
	position:absolute;
	visibility:hidden;	
	overflow: visible;
	background-color:transparent;
 
	margin-left:10px;  /*-5px;*/
	margin-top: 1px;  /*-10px;*/
	
	//margin-left: -101px  ;/*-60px;  */  /*IE*/
	//margin-top: 17px ;/*10px;    */  /*IE*/
}
.contenidoTooltip {
	background-color: transparent;   
	color:#FFFFFF;
	font-family:Arial;
	font-size:8pt;
	font-weight:bold;	}

.contenidoTooltip a{	   
	color:#FFFFFF;
	font-family:Arial;
	font-size:8pt;
	font-weight:bold;
    text-decoration:none;	}

.contenidoTooltip a:hover{ color:#333333; }

.tool{ border:1px #666 solid; background-color:#333;opacity:0.8;filter:alpha(opacity=80); }

.CabeceraDrag {width:100%;top:0px; left:0px;}

</style>
 
<script language"javascript">
 
 function Subir(gadnro){
    document.getElementById("ifrm2").src = "~/../Config_Gadget.aspx?gadnro="+gadnro+"&gadnro1=-1&gadnro2=-1&sube=-1&desactiva=0&activa=0";  
 }
 
 function Bajar(gadnro){
    document.getElementById("ifrm2").src = "~/../Config_Gadget.aspx?gadnro="+gadnro+"&gadnro1=-1&gadnro2=-1&sube=0&desactiva=0&activa=0";  
 }
 
 function Intercambiar(gadnro1,gadnro2){
    document.getElementById("ifrm2").src = "~/../Config_Gadget.aspx?gadnro1="+gadnro1+"&gadnro2="+gadnro2+"&sube=0&desactiva=0&activa=0";  
 }
 
 
  function Desactivar(gadnro,titulo){ 
    if (confirm(titulo))
      document.getElementById("ifrm2").src = "~/../Config_Gadget.aspx?gadnro="+gadnro+"&gadnro1=-1&gadnro2=-1&sube=0&desactiva=-1&activa=0";  
 }
 
 function AbrirTooltipHelp(obj) {
   document.getElementById(obj).style.visibility = 'visible';
}
 
function CerrarTooltipHelp(obj) {
     document.getElementById(obj).style.visibility = 'hidden';
}

//funcion que se encarga de redimensionar un iframe a su contenido
function Redimensionar(id)
{
  if (!window.opera && document.all && document.getElementById) {
    id.style.height=id.contentWindow.document.body.scrollHeight;
  }
  else if(document.getElementById) {
          id.style.height=id.contentDocument.body.scrollHeight+"px";
       }
} 
 
</script>

 
 



<!-- ------- IMPRIME GADGET TIPO CONTROL ------------------------------------- -->

<TABLE cellpadding="0" cellspacing="0" border="0" align="left" width="100%"  >
<TR>
<TD align="center" valign="top" style=" width:100%;" id="ContPpal" >       
      
      <div  id="Cuerpo" runat="server"   visible="false" style=" text-align:center; margin-left:10px;"></div>
      <asp:Panel ID="MiPanel" runat="server"></asp:Panel> 
      <asp:Panel ID="Panel1" runat="server"></asp:Panel> 
      
      <iframe src="" id='IfrmAccesos' name='IfrmAccesos' onload="Redimensionar(this);" style="visibility:hidden; height:0%"  scrolling="no" frameborder="0"></iframe>
      
      <iframe src="" id='ifrm' name='ifrm' style="visibility:hidden; height:0px; width:0px;" ></iframe>
      
</TD>
</TR>
</TABLE>
 
 

 
 
 

 
 