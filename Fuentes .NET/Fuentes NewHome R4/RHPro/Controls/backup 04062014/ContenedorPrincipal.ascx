<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="ContenedorPrincipal.ascx.cs" 
Inherits="RHPro.Controls.ContenedorPrincipal" %>
 
<script language="javascript">
function abrirVentana(url, name, width, height,opc) { 
 /*
  var str = "height=" + height + ",innerHeight=" + height;
  str += ",width=" + width + ",innerWidth=" + width;
  */
    if (name != "ESS") {
        height = "500px";
        width = "950px";
    }  
    
         
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

 
 
<script language"javascript">
 
 function Subir(gadnro){ 
 if (gadnro!="") {
    Abrir_Progreso();    
    document.getElementById("ifrm2").src = "~/../Config_Gadget.aspx?gadnro="+gadnro+"&gadnro1=-1&gadnro2=-1&sube=-1&desactiva=0&activa=0";  
    }
     
   
 }
 
 function Bajar(gadnro){
 if (gadnro!="") {
    Abrir_Progreso();  
    document.getElementById("ifrm2").src = "~/../Config_Gadget.aspx?gadnro="+gadnro+"&gadnro1=-1&gadnro2=-1&sube=0&desactiva=0&activa=0";  
  }
 }
 
 function Intercambiar(gadnro1,gadnro2){
 if ( (gadnro1!="") &&  (gadnro2!="") ) 
  {
    Abrir_Progreso();  
    document.getElementById("ifrm2").src = "~/../Config_Gadget.aspx?gadnro1="+gadnro1+"&gadnro2="+gadnro2+"&sube=0&desactiva=0&activa=0";  
   }
 }
 
 
  function Desactivar(gadnro,titulo){ 
    if (gadnro!="") {
    if (confirm(titulo))
    {
       Abrir_Progreso();  
       document.getElementById("ifrm2").src = "~/../Config_Gadget.aspx?gadnro="+gadnro+"&gadnro1=-1&gadnro2=-1&sube=0&desactiva=-1&activa=0";  
    }
    }
     
      
 }
 
 function AbrirTooltipHelp(obj) {
     document.getElementById(obj).style.visibility = 'visible';
     document.getElementById(obj).style.display = '';
}
 
function CerrarTooltipHelp(obj) {
    document.getElementById(obj).style.visibility = 'hidden';
    document.getElementById(obj).style.display = 'none';
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

function Abrir_Progreso(){
  var prog = document.getElementById("Progreso");
  var trans = document.getElementById("Transparente");
  prog.style.visibility = "visible";
  trans.style.visibility = "visible";
} 
 
function Cerrar_Progreso(){
  var prog = document.getElementById("Progreso");
  var trans = document.getElementById("Transparente");
  prog.style.visibility = "hidden";
  trans.style.visibility = "hidden";
} 

</script>

 
<style> 
 *{
font-family:sans-serif;
list-style:none;
text-decoration:none;
margin:0;
padding:0;

}

.ContenedorBarraNavegacion { width:100%; background-color:#efefef; height:30px; border-bottom:1px solid #999999 }

.BarraNavegacion > li {
    float:left;
    font-family:Tahoma;
    font-size:9pt;
}

.BarraNavegacion li   {
    background:#efefef;
    color:#777777;
    display:block;
    border:0px solid #888888;
    padding:5;
    border-right:1px solid #ccc;
    height:20px;
}

.BarraNavegacion li .flecha{
    font-size: 9pt;
    padding-left: 6px;
    display: none;
}

.BarraNavegacion li:not(:last-child) .flecha {
display: inline;
}

.BarraNavegacion li:hover {
background:#cccccc;
}

.BarraNavegacion li {
    position:relative;

}

.BarraNavegacion li ul {
    display:none;
    position:absolute;
    min-width:190px;
    top:30px;
    border:1px solid #ccc;
}

.BarraNavegacion li:hover > ul {
    display:block;
    cursor:default;
}

.BarraNavegacion li ul li ul {
    right: -191px; 
    top:0;
}

 
 

</style>


<%  RHPro.Lenguaje Obj_Lenguaje;    
    Obj_Lenguaje = new RHPro.Lenguaje();    
%>
<!-- ------- IMPRIME GADGET TIPO CONTROL ------------------------------------- -->
<DIV id="Transparente"></DIV>
<DIV id="Progreso"><img src="~/../img/loader.gif"> <p> <% Response.Write(Obj_Lenguaje.Label_Home("Cargando")); %></p></DIV>

<asp:Panel ID="MenuPrincipalModulo" runat="server">
</asp:Panel>

<TABLE cellpadding="0" cellspacing="0" border="0" align="left" width="100%"  >
<TR>
<TD align="center" valign="top" id="ContPpal" style="width:100%;" >       
      
      <div  id="Cuerpo" runat="server"   visible="false" style=" text-align:center; margin-left:10px;"></div>
      <asp:Panel ID="MiPanel" runat="server" CssClass="ContenedorMenuPrincipal">      
      </asp:Panel> 
      
      
      <iframe src="" id='IfrmAccesos' name='IfrmAccesos' onload="Redimensionar(this);" style="visibility:hidden; height:0%"  scrolling="no" frameborder="0"></iframe>
      
      <iframe src="" id='ifrm' name='ifrm' style="visibility:hidden; height:0px; width:0px;" ></iframe>
      
</TD>
</TR>
</TABLE>
 
 

 
 
 

 
 