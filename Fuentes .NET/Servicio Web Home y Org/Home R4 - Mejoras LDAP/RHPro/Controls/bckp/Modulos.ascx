
<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="Modulos.ascx.cs" Inherits="RHPro.Controls.Modulos" %>

 

<style>
.Separador { color:#FFF; font-family:Arial; font-size:11pt; background:url(img/Fondo_Menu.png) repeat-x top; height:37px; cursor:pointer }
.Menu_Ppal TR { color:#FFF; font-family:Arial; font-size:11pt; background:url(img/Fondo_Menu.png) repeat-x top; height:37px; cursor:pointer }
.Menu_Links TR { color:#FFF; font-family:Arial; font-size:9pt; background:url(img/Fondo_Menu.png) repeat-x top; height:27px; cursor:pointer }
 
.ASPlink 
{ 
 background-color:transparent; text-decoration:none; font-family:Arial; font-size:10pt; color:#FFFFFF; 
 cursor: pointer;
}

.ASPlink:hover
{ 
  color:#CCCCCC;  
}
.ListaOculta
{
    height:0%;
	width:0%;
	position:absolute;
	left:0; top:0;
	visibility:hidden;
}

.FondoOculto
{
	height:155%;
	width:100%;
	position:absolute;
	left:0; top:0;	 
	vertical-align:middle;
	text-align:center;
	background-color:#000000;
	opacity:0.4;
    filter:alpha(opacity=40);
     z-index:2000;
}
.ListaGadgets 
{
	 color:#FFFFFF; 
	 font-size:15pt; 
	 position: absolute; 
	 top: 100px; 
	 vertical-align:middle; 
	 text-align:center; 
	 z-index:2001;
} 
</style>


<script type="text/javascript">

var activo = 0;
 
var CantidadLink = 16;
var LinkSeleccionado = "";


function abrirVentana(url, name, width, height,opc) 
{
  var str = "height=" + height + ",innerHeight=" + height;
  str += ",width=" + width + ",innerWidth=" + width;
  str += ",resizable=yes"
  if (opc != null)
	  str += opc
  var auxi;
  auxi = url.substr(url.lastIndexOf('/')+1,url.length);
  auxi = auxi.substr(0,auxi.indexOf(".asp"));
  window.open(url,"",str); 
 //window.open(url, auxi, str); 
}


 
function AbrirModulo(Link, menuname){  
 if (menuname=="ESS") //En el caso que se desea abrir el modulo de ESS, el tamaño de ventana va a ser mayor
 	abrirVentana(Link,"",AnchoPantalla()-100,AltoPantalla()-150,"");
  else
     abrirVentana(Link,"",700,500,"");
}

function Seleccionar(id,Link){
 
  for (i=0; i<CantidadLink; i++) {	 
	document.getElementById("Link"+i).style.background = "url(img/Fondo_Menu.png)"; 
   } 
   LinkSeleccionado = id;
   document.getElementById(id).style.background = "url(img/Fondo_Menu_Press.png)"; 
 }

function InfoLink(id,Link) {  
  Seleccionar(id,Link);
}

function Sobre(obj){   
// for (i=0;i<obj.cells.length;i++) {
//   obj.cells[i].style.background = "url(img/Fondo_Menu_Press.png)";	  
// }  
	   obj.style.background = "url(img/Fondo_Menu_Press.png)"; 	 
	  
	  
	}
	
function Sale(obj){
   if (LinkSeleccionado!=obj.id)
	   obj.style.background = "url(img/Fondo_Menu.png)";
//	 for (i=0;i<obj.cells.length;i++) {
//       obj.style.background = "url(img/Fondo_Menu.png)";  
//     }
	}
</script>
 
  <!-- MODULOS ---------------------------->

 <DIV id="ListaGadgets" class="ListaGadgets" style="visibility:hidden"  >LISTADO DE GADGET</DIV>
 <DIV id="FondoOculto" class="FondoOculto"  style="visibility:hidden" ></DIV>

 
    <table width="100%" border="0" cellspacing="0" cellpadding="0" class="Menu_Links">
 
      <tr style="height:20px" >
        <td  valign="middle" align="right"nowrap="nowrap"  colspan="3" style="color:#999; font-family:Arial; font-size:8pt; ">
         <span  style="margin-right:5px; vertical-align:middle">         
            <% Response.Write(ObjLenguaje.Label_Home("Modulos")); %>
            <img src="img/flecha.png" border='0' align="middle" />           
          </span>
         </td>         
      </tr>

<!-- -------------------------Armo el manu con los modulos activos -----------   -->      
      <asp:SqlDataSource
          id="SqlDataSource1"
          runat="server"
          DataSourceMode="DataSet"
          ConnectionString="Password=ess;Persist Security Info=True;User ID=ess;Initial Catalog=Base_0_R3_ARG;Data Source=RHDESA"
          SelectCommand="SELECT ROW_NUMBER() OVER(ORDER BY menudesabr ASC)  'pos' ,* FROM menumstr  WHERE menuraiz = 74 AND menuactivo = -1 ORDER BY menudesabr">
      </asp:SqlDataSource>     
      
      <asp:Repeater ID="Repeater1" runat="server"  >
           <HeaderTemplate> </HeaderTemplate>
           <ItemTemplate >                                                                
                <tr id='Link<%#Eval("pos") %>' onmouseover="Sobre(this)" onmouseout="Sale(this)">
                      <td nowrap="nowrap" style="width:1px; background-color:Transparent" valign="middle" >
                        <img src="img/Modulos/<%#Eval("menuname") %>.png" border='0' align="bottom" style="margin-left: 4px; margin-right:9px; ">          
                      </td>
                      <td nowrap="nowrap" align="left" valign="middle" >                   
                       <span style="margin-left:8px;"  >                        
                        <asp:LinkButton OnCommand="ActualizarContenedor"  CommandArgument='<%# Eval("menuname") %>'
                         runat="server" CssClass="ASPlink"   >                           
                              <DIV style=" width:200px;">  
                               
                                <%# ObjLenguaje.Label_Home((String)Eval("menudesabr"))%>                                
                              </DIV>
                        </asp:LinkButton>   
                       </span>
                      </td>
                      <td>    
                      <%  //Si algun usuario esta logeado muestra el acceso a los modulos 
                          if (Common.Utils.IsUserLogin)
                          {  
                      %>                      
                            <img src='img/plus.png' border='0' align='absmiddle' onmouseover="this.src = 'img/plus_hover.png'"  onmouseout="this.src = 'img/plus.png'" onclick="AbrirModulo('../<%# Eval("action")%>','<%#Eval("menuname") %>')">  
                    
                      <% }
                          else
                          { %>
                       <img src='img/plus.png' border='0' align='absmiddle' style='visibility:<%# Visibilidad( (String) Eval("menuname") == "ESS" )  %>' onmouseover="this.src = 'img/plus_hover.png'"  onmouseout="this.src = 'img/plus.png'" onclick="AbrirModulo('../<%# Eval("action")%>','<%#Eval("menuname") %>')">  
                       <%} %>
                      </td>
               </tr>              
                
                                                            
           </ItemTemplate>
      </asp:Repeater> 
<%  
    //Actualizo la ultima posicion de links para continuar con los accesos basicos
    //ActualizarPosmenu();
%>
<!-- --------------------------------------------------------------------------------------------------------------   -->
</table>
 
<!-- ------------------------- ACCESOS BASICOS  -->
    
    <table width="100%" border="0" cellspacing="0" cellpadding="0" class="Menu_Ppal">
        
              <tr style="background:url(img/Fondo_Menu_Press4.png) repeat-x">
              <td nowrap="nowrap" style="width:1px; " >
                <div style="background:url(img/Fondo_Menu_Press2.png) no-repeat left;height:35px; width:34px; vertical-align: bottom">
                   <div style="//margin-top:8px">
                      <img src="img/Modulos/Gadget2.png" border='0' align="absmiddle" style="margin-left:4px; margin-right:9px; margin-top:8px;//margin-top:0px">          
                    </div>
                </div>                 
              </td>
                      <td  nowrap    >                   
                       <span style="margin-left:8px;"  >                        
                        <asp:LinkButton OnCommand="ActualizaGadgets"  
                         runat="server" CssClass="ASPlink"  style="font-size:14pt;"  >   
                                                 
                               <%= ObjLenguaje.Label_Home("Gadgets")%>
                        </asp:LinkButton>   
                       </span>
            </td>
             <td align="right">
                      <%  //Si algun usuario esta logeado muestra el acceso a los modulos 
                          if (Common.Utils.IsUserLogin)  {  
                      %>
                    <img src='img/plusG.png' border='0' align='absmiddle'  onclick="AbrirModal()" style=" margin-right:4px;">  
                      <% } %>
            </td>
       </tr>
      
        <tr style="height:20px" >
        <td  valign="middle" align="right"nowrap="nowrap"  colspan="3" style="color:#999; font-family:Arial; font-size:8pt; ">
         <span  style="margin-right:5px; vertical-align:middle">         
            <% Response.Write(ObjLenguaje.Label_Home("Accesos")); %> <img src="img/flecha.png" border='0' align="middle" />           
          </span>
         </td>         
      </tr>
          
<!-- -------------------------- Se busca los accesos en el archivo Accesos_Home.xml -------------------------------------------->
<asp:repeater id="rpMyRepeater" runat="server">
  <ItemTemplate>
  
    <%# Construir_Acceso((String)DataBinder.Eval(Container.DataItem, "Activo"), (String)DataBinder.Eval(Container.DataItem, "Nombre"), (String)DataBinder.Eval(Container.DataItem, "URL"), (String)DataBinder.Eval(Container.DataItem, "isLogin"))%>   
  </ItemTemplate> 
</asp:repeater> 

        
    </table>
    
    <script type="text/javascript">
      var CantidadLink =  <%= posmenu %>;            
    </script>
  
 