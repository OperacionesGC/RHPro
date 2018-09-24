
<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="Modulos.ascx.cs" Inherits="RHPro.Controls.Modulos" %>
 


<script type="text/javascript">

var activo = 0;
 
var CantidadLink = 16;
var LinkSeleccionado = "";

 /*
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
 */
 


function AbrirModulo(Link, menuname) {
  
 if (menuname=="ESS") //En el caso que se desea abrir el modulo de ESS, el tamaño de ventana va a ser mayor
 	abrirVentana(Link,"ESS",AnchoPantalla()-100,AltoPantalla()-150,"");
  else
   abrirVentana(Link,"Modulo",AnchoPantalla()-160,500,"");
   // abrirVentana(Link,"",700,500,"");
/* 
Link = String(Link).replace("abrirVentana('","abrirVentana('../");
eval(Link);
*/
}

function AbrirMRU(menumsnro, menuraiz)  {
    //ifrmModulos.location = "/rhprox2/shared/asp/mru_00.asp?menumsnro=" + menumsnro + "&menuraiz=" + menuraiz;
}

function Formatear(acceso){ 
 return acceso.replace("'","\"");
}

function Seleccionar(id,Link){
 /*
  for (i=1; i<=CantidadLink; i++) {	  
	document.getElementById("Link"+i).style.background = "url(img/Fondo_Menu.png)";
} 
   
   LinkSeleccionado = id;
   document.getElementById(id).style.background = "url(img/Fondo_Menu_Press.png)"; */
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


function ExpandInactivos(item){ 
 
 obj=document.getElementById(item);
 texto=document.getElementById("txtInactivos");
 flecha=document.getElementById("imgInactivos");
 
 if (obj.style.display=="none") {
  obj.style.display="block";
  flecha.src = "img/up.png";
  texto.innerHTML = "Ocultar Inactivos";
 }
 else {
  obj.style.display="none"; 
  flecha.src = "img/down.png";
  texto.innerHTML = "Mostrar Inactivos";
 }
 
}
	
</script>
 
  <!-- MODULOS ---------------------------->

 <DIV id="ListaGadgets" class="ListaGadgets" style="visibility:hidden"  >LISTADO DE GADGET</DIV>
 <DIV id="FondoOculto" class="FondoOculto"  style="visibility:hidden" ></DIV>

 
    <table width="100%" border="0" cellspacing="0" cellpadding="0" class="Menu_Links">
     
     <%   if (bool.Parse(ConfigurationManager.AppSettings["VisualizarComplementos"]))
          { %>
     <tr >
      <td nowrap="nowrap" style="width:1px; " valign="middle" align="center" >
        
           
              <!--img src="img/Modulos/Gadget2.png" border='0' align="absmiddle"  class="IconoModulo"   -->          
              <img src="img/Modulos/SVG/HOME.svg" border='0' class="IconoModulo"    >          
        
      </td>
     <td  nowrap    >                   
                                  
                <asp:LinkButton ID="LinkButton1" OnCommand="ActualizaGadgets"  runat="server" CssClass="ASPlink"     >                                            
                      
                </asp:LinkButton>   
              
    </td>
     <td align="center" valign="center">
              <%  //Si algun usuario esta logeado muestra el acceso a los modulos 
         if (Common.Utils.IsUserLogin)
         {  
              %>
            <!--img src='img/plusG.png' border='0' align='absmiddle'  onclick="AbrirModal()" style=" margin-right:4px;"-->  
             <!--span class="AperturaModulo" onclick="AbrirModal()" > » </span-->
            
             <img src="img/Modulos/SVG/APERTURA_MODULO.svg" border='0' class="IconoAperturaModulo" onclick="AbrirModal()"   >           
              <% } %>
     </td>
    </tr>
    <% } %>



      <tr style="height:20px" >
        <td  valign="middle" align="right"nowrap="nowrap"  colspan="3" class="SeparadorMenu"  >
         <span  style="margin-right:5px; vertical-align:middle">         
            <% Response.Write(ObjLenguaje.Label_Home("Modulos")); %>
            <img src="img/flecha.png" border='0' align="middle" />           
          </span>
         </td>         
      </tr>

<!-- -------------------------Armo el manu con los modulos activos -----------   -->            
  <%  //En el caso que sea un usuario anonimo
      if (!Common.Utils.IsUserLogin)
      {  
  %>
  
  
      <asp:Repeater ID="Repeater1" runat="server"  >
           <HeaderTemplate> </HeaderTemplate>
           <ItemTemplate >                                                                
                  <tr id='Link<%=posmenu %>'  >
                      <td nowrap="nowrap"  valign="middle"  class="FONDO_SVG"  >                        
                        <asp:LinkButton   OnCommand="ActualizarContenedor"  CommandArgument='<%# Eval("MenuName") %>' runat="server"     >    
                            <img src="img/Modulos/SVG/<%#Eval("MenuName") %>.svg" border='0' class="IconoModulo" >  
                          </asp:LinkButton>   
                      </td>
                      <td nowrap="nowrap" align="left" valign="middle" >                   
                                            
                        <asp:LinkButton   OnCommand="ActualizarContenedor"  CommandArgument='<%# Eval("MenuName") %>'
                         runat="server" CssClass="ASPlink"   >                           
                             <DIV> 
                              <% posmenu++; %>                              
                              <%# ObjLenguaje.Label_Home((String)Eval("MenuTitle"))%>                                 
                               </DIV>
                        </asp:LinkButton>   
                        
                      </td>
                      
               </tr>                                                                     
           </ItemTemplate>
      </asp:Repeater> 
      <%  }  else  { %>
      <!-- RECUPERO TODOS LOS MODULOS HABIITADOS PARA EL USUARIO -->
           <asp:Repeater ID="Repeater2" runat="server"  >
           <HeaderTemplate> </HeaderTemplate>
           <ItemTemplate >                                                                
                  <tr id='Link<%=posmenu %>'  >
                      <td nowrap="nowrap"  valign="middle" class="FONDO_SVG" >
                        <img src="img/Modulos/SVG/<%#Eval("MenuName") %>.svg" border='0' class="IconoModulo" >          
                      </td>
                      <td nowrap="nowrap" align="left" valign="middle" >                   
                       <span style="margin-left:8px;"  >                        
                        <asp:LinkButton ID="LinkButton3"   OnCommand="ActualizarContenedor"  CommandArgument='<%# Eval("MenuName") %>'
                         runat="server" CssClass="ASPlink"   >                           
                              <DIV style=" width:200px;">  
                              <% posmenu++; %>                              
                              <%# ObjLenguaje.Label_Home((String)Eval("MenuTitle"))%>                                  
                              </DIV>
                        </asp:LinkButton>   
                       </span>
                      </td>
                      <td align="center" valign="center">  
                       
                      
                            <%# Imprimir_Action((string)Eval("Action"), (String)Eval("MenuName"), (string)Eval("menuraiz"), (string)Eval("menumsnro"))%>         
                 
                      </td>
               </tr>                                                                     
           </ItemTemplate>
      </asp:Repeater> 
      
      <% if (bool.Parse(ConfigurationManager.AppSettings["VisualizarModulosInhabilitados"]))
         { %>
              <!-- RECUPERO TODOS LOS MODULOS INHABIITADOS PARA EL USUARIO -->
               <tr style="height:20px" onclick="ExpandInactivos('ModInactivos')" >
                <td  valign="middle" align="right"nowrap="nowrap"  colspan="3"  >
                 <span  style="margin-right:5px; vertical-align:middle"  >         
                   <span id="txtInactivos"> <% Response.Write(ObjLenguaje.Label_Home("Ocultar Inactivos")); %></span>
                    <img src="img/up.png" border='0' align="absmiddle" width="6" height="6" id="imgInactivos" />           
                  </span>
                 </td>         
              </tr>
              
              <tr>
                <td colspan="3" style="height=1px">
                  <TABLE cellpadding="0" cellspacing="0" border="0" id="ModInactivos"  width="100%">
         
                 <asp:Repeater ID="Repeater3" runat="server"    >
                   <HeaderTemplate> </HeaderTemplate>
                   <ItemTemplate >                                                                
                          <tr id='Link<%=posmenu %>'  >
                              <td nowrap="nowrap"   valign="middle" >                               
                                 <img src="img/Modulos/SVG/<%#Eval("MenuName") %>.svg" border='0' class="IconoModulo" >     
                              </td>
                              <td nowrap="nowrap" align="left" valign="middle" width="100%" >                   
                               <span style="margin-left:8px;"  >                        
                                <asp:LinkButton ID="LinkButton3"   OnCommand="ActualizarContenedor"  CommandArgument='<%# Eval("MenuName") %>'
                                 runat="server" CssClass="ASPlink"   >                           
                                      <DIV style=" width:200px;">  
                                      <% posmenu++; %>                              
                                      <%# ObjLenguaje.Label_Home((String)Eval("MenuTitle"))%>                                       
                                      </DIV>
                                </asp:LinkButton>   
                               </span>
                              </td>
                               
                       </tr>                                                                     
                   </ItemTemplate>
              </asp:Repeater> 
               </table>
               </td>
               </tr>
       <% } %>
       
       
       
      <% } %>
 <!-- --------------------------------------------------------------------------------------------------------------   -->
</table>
 
<!-- ------------------------- ACCESOS BASICOS  -->
      
<table width="100%" border="0" cellspacing="0" cellpadding="0" class="Menu_Links">      
 

<tr style="height:20px" > 
<td  valign="middle" align="right"nowrap="nowrap"  colspan="3"   >
 <span  style="margin-right:5px; vertical-align:middle">         
    <% Response.Write(ObjLenguaje.Label_Home("Accesos")); %> <img src="img/flecha.png" border='0' align="middle" />           
  </span>
 </td>         
</tr>

 

<!-- Busco los accesos  -->           
  <asp:Repeater ID="RepAccesos" runat="server"  >
           <HeaderTemplate> </HeaderTemplate>
           <ItemTemplate>       
           <tr id='Link<%=posmenu %>' onclick="Seleccionar('Link<%=posmenu %>','')"      > 
                    <td nowrap='nowrap'>
                       <!--img src='img/link.png' border='0' align='absmiddle'  style="margin-left: 4px;"/-->
                        <img src="img/Modulos/SVG/LINK.svg" border='0' class="IconoModulo"    >
                    </td>
                    <td>     
                      <asp:LinkButton ID="LinkButton2"   OnCommand="Actualizar_Accesos_XML"  CommandArgument='<%#DataBinder.Eval(Container.DataItem, "nroAcceso")%>'
                         runat="server" CssClass="ASPlink"   >                           
                           <DIV style=" width:200px;">  
                             <% posmenu++; %>                                                           
                                <%# DataBinder.Eval(Container.DataItem, "Nombre")%> 
                              
                                 <%//#DataBinder.Eval(Container.DataItem, "nroAcceso")%>  
                                 <%//#DataBinder.Eval(Container.DataItem, "URL")%>  
                                  <%//#DataBinder.Eval(Container.DataItem, "isLogin")%>
                           </DIV>
                        </asp:LinkButton> 
                    
                    </td> 
                    <td align='right'> 
                        <%# ImprimirLink(Boolean.Parse((String)DataBinder.Eval(Container.DataItem, "isLogin")), (string)DataBinder.Eval(Container.DataItem, "URL"))%>                    

                     </td> 
              </tr> 
                                                                                         
           </ItemTemplate>
      </asp:Repeater> 
<!-- -------------------------- Se busca los accesos en el archivo Accesos_Home.xml -------------------------------------------->
<% 
   //  posmenu++;
   //  Response.Write(Leer_XML()); 
 %>        
   </table>
   
   <iframe src="" name='ifrmModulos' id="ifrmModulos" style="display:none;"></iframe>
 
    
    <script type="text/javascript">
      var CantidadLink =  <%= posmenu %> - 1;            
   
    </script>
  
 