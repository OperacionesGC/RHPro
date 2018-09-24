<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="ContenedorPrincipal.ascx.cs" 
Inherits="RHPro.Controls.ContenedorPrincipal" %>
 
<script language="javascript">
 



function AbrirLink(Link,menuname){ 

	 if (menuname=="ESS") //En el caso que se desea abrir el modulo de ESS, el tamaño de ventana va a ser mayor
 	abrirVentana(Link,"ESS",AnchoPantalla()-100,AltoPantalla()-150,"");
  else
     abrirVentana(Link,"",700,500,"");
}
</script>

 

<script type="text/javascript" src="~/../Js/Drag.js"></script>


 
<script language"javascript">

 function Subir(gadnro) { 
 
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


 function Desactivar(gadnro, titulo)
  { 
    if (gadnro!="") {
    if (confirm(titulo))
    {
       Abrir_Progreso();  
       document.getElementById("ifrm2").src = "~/../Config_Gadget.aspx?gadnro="+gadnro+"&gadnro1=-1&gadnro2=-1&sube=0&desactiva=-1&activa=0";  
    }
    }
}

function ExpandirAltura(gadnro, alturaActual, gadtipo) {
    
    var dimension;
    
    if (alturaActual == 0)
    {  dimension = -1; }
    else {  dimension = 0; }

    if (gadnro != "") {
        Abrir_Progreso();

        document.getElementById("ifrm2").src = "~/../Config_Gadget.aspx?gadnro=" + gadnro + "&gadnro1=-1&gadnro2=-1&sube=0&desactiva=0&activa=0&ModificoAlto=-1&alturaActual=" + dimension + "&gadtipo=" + gadtipo;
    }

}

function ExpandirAncho(gadnro, anchoActual, gadtipo) {
 
    var dimension;    
    
    if (anchoActual == 0)
        dimension = -1;
    else dimension = 0;
    
    if (gadnro != "") {
        Abrir_Progreso();
        document.getElementById("ifrm2").src = "~/../Config_Gadget.aspx?gadnro=" + gadnro + "&gadnro1=-1&gadnro2=-1&sube=0&desactiva=0&activa=0&ModificoAncho=-1&anchoActual=" + dimension + "&gadtipo=" + gadtipo;
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


var GADNRO_MRU_MODULO = "";
var GADNRO_INDICADORES = "";

function Ocultar_MRU_Vacio() {
   
    if (document.getElementById(GADNRO_MRU_MODULO)) {
        document.getElementById(GADNRO_MRU_MODULO).style.display = 'none';
        }
}


function Ocultar_Indicador_Vacio() {
 
    if (document.getElementById(GADNRO_INDICADORES)) {
        document.getElementById(GADNRO_INDICADORES).style.display = 'none';
    }
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
 

</style>


<%  RHPro.Lenguaje Obj_Lenguaje;    
    Obj_Lenguaje = new RHPro.Lenguaje();    
%>

 


<!-- ------- IMPRIME GADGET TIPO CONTROL ------------------------------------- -->
<DIV id="Transparente"  Class="PopUp_FondoTransparente" style="z-index:1000 !important;"></DIV>               
<div  id="Progreso"  style="padding:10px; background-color:transparent; text-align:center; position:absolute; left:45%;top:16%;z-index:1001; font-size:9pt; font-family:Tahoma; color:#fff">
   <img src="img/LOGO_RHPRO_loader.png" align="absmiddle"/>                                
   <div style="width:108px; height:46px; overflow:hidden">  <img src="img/miniloaderplano.gif" align="absmiddle"/>   </div>
</div>    


<asp:Panel ID="MenuPrincipalModulo" runat="server" CssClass="ContenedorGadgets">
</asp:Panel>


   
       
<TABLE cellpadding="0" cellspacing="0" border="0" align="left" width="100%"  >
<TR>
<TD align="left" valign="top" id="ContPpal" style="width:100%;" >       
      
      <asp:Panel ID="Gadgets_Del_Modulo" runat="server" CssClass="Gadgets_Del_Modulo"  >
      </asp:Panel> 
      
      <div  id="Cuerpo" runat="server"   visible="false" style=" text-align:center; "> </div>           
      
      <asp:Panel ID="MiPanel" runat="server"  >            
      </asp:Panel> 
       <!-- CONTROL DE PROGRESO -->
                        <span style="margin-left:40px; display:inline-block; text-align:center; vertical-align:top" >      
                        <asp:UpdateProgress ID="UpdateProgress" AssociatedUpdatePanelID="Update_Mod" runat="server"  Visible="true"  >
                        <ProgressTemplate>   
                        
                        <DIV id="TransparenteProgress"  Class="PopUp_FondoTransparente" style="z-index:1000 !important;"></DIV>               
                        
                        <div style="padding:10px; background-color:transparent; text-align:center; position:absolute; left:45%;top:16%;z-index:1001; font-size:9pt; font-family:Tahoma; color:#fff">
                               <img src="img/LOGO_RHPRO_loader.png" align="absmiddle"/>                                
                               <div style="width:108px; height:46px; overflow:hidden">  <img src="img/miniloaderplano.gif" align="absmiddle"/>   </div>
                           </div>    
                        </ProgressTemplate>               
                        </asp:UpdateProgress>       
                        </span>
      <!-- ------------------------------------------- -->
      
      <!--iframe src="" id='IfrmAccesos' name='IfrmAccesos' onload="Redimensionar(this);"  scrolling="no" frameborder="0"></iframe-->
      
      <iframe src="" id='ifrm' name='ifrm' style="display:none; height:0px; width:0px;" ></iframe>
     
        <input type="hidden" id="idIdentidicador" value="" runat="server"></input>
        <input type="hidden" id="idMRU" name="idMRU" value="" runat="server" ></input>
         
     
  
</TD>
</TR>
</TABLE>
     
 

 
 
 

 
 