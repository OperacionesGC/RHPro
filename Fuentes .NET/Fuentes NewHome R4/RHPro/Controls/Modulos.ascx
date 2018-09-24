
<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="Modulos.ascx.cs" Inherits="RHPro.Controls.Modulos" %>
 


<script type="text/javascript">

var activo = 0;
 
var CantidadLink = 16;
var LinkSeleccionado = "";

 

function AbrirModulo(Link, menuname) {
/*
    if (menuname == "ESS") //En el caso que se desea abrir el modulo de ESS, el tamaño de ventana va a ser mayor
      abrirVentana(Link,"ESS",AnchoPantalla()-100,AltoPantalla()-150,"");         
    else
        abrirVentana(Link, "Modulo", AnchoPantalla() - 160, 500, "");
*/
    window.open(Link, '_blank', 'location=yes, toolbar=yes, scrollbars=yes, resizable=yes, width=800, height=600');
   
}

function AbrirMRU(menumsnro, menuraiz) {
    alert(menumsnro);
    alert(menuraiz);
    ifrmModulos.location = "/rhprox2/shared/asp/mru_00.asp?menumsnro=" + menumsnro + "&menuraiz=" + menuraiz;
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
 texto=document.getElementById("<%=txtInactivos.ClientID%>");
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


      <asp:UpdatePanel ID="Update_Mod" runat="server"  >
       <ContentTemplate> 
 
<table width="39" border="0" cellspacing="0" cellpadding="0" class="Menu_Links"  id="CC_Tabla_Modulos">
     
   <tr>
   <th style="white-space:nowrap" > 
    
        <span class="BotonCabeceraModulos" > 
            <a class='BtnTransparenteOcultar' onclick='DesplazarBarraMenu()'> </a>    
            <%=Common.Utils.Armar_Icono("img/modulos/SVG/CONTROLMODULOS.svg", "IconosMaximizaModulos", "", "", "", "")%>
            <span><%= ObjLenguaje.Label_Home("Modulos")%>  </span>
        </span> 
            
    </th>
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
                      <td nowrap="nowrap" align="left" valign="middle" >                                                               
                     
                        <asp:LinkButton  id="Btn_AccModuloInfo" OnCommand="ActualizarContenedor" AutoPostBack="False"  CommandArgument='<%# Eval("MenuName")+"@"+Eval("MenuMsnro") %>'
                         runat="server" CssClass="ASPlink"   >   
                               
                               <%#Common.Utils.Armar_Icono("img/modulos/SVG/" + Convert.ToString(Eval("MenuName"))+".svg", "IconoModulo", ObjLenguaje.Label_Home((String)Eval("MenuTitle")), "", "")%>
                               
                               <span> 
                              <% posmenu++; %>                              
                              <%# ObjLenguaje.Label_Home((String)Eval("MenuTitle"))%>                           
                               </span>
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
                      <td nowrap="nowrap" align="left" valign="middle" width="100%"  >                                                
                       
                        <asp:LinkButton ID="LinkButton3"   OnCommand="ActualizarContenedor" AutoPostBack="False"   CommandArgument='<%# Eval("MenuName")+"@"+Eval("MenuMsnro") %>'
                         runat="server" CssClass="ASPlink"   >                                                               
                                 <%#Common.Utils.Armar_Icono("img/modulos/SVG/" + Convert.ToString(Eval("MenuName")) + ".svg", "IconoModulo", ObjLenguaje.Label_Home((String)Eval("MenuTitle")), "", "")%>
                              <div>  
                              <% posmenu++; %>   
                              <%# Agregar_Modulo_Visibilidad((String)Eval("MenuName"), true)%>                                        
                              <%# ObjLenguaje.Label_Home((String)Eval("MenuTitle"))%>                                                                 
                              </div>
                        </asp:LinkButton>   
                      
                      </td>
                      
               </tr>                                                                     
           </ItemTemplate>
      </asp:Repeater> 
      
         
       
      <% } %>
 <!-- --------------------------------------------------------------------------------------------------------------   -->
</table>
 
<%  if (Common.Utils.IsUserLogin)
    {
        if (bool.Parse(ConfigurationManager.AppSettings["VisualizarModulosInhabilitados"]))
        {
             %>
<asp:Panel runat="server" id="Menu_Links_Inact">   
         
<table width="100%"  border="0" cellspacing="0" cellpadding="0" class="Menu_Links_Inact"  id="CC_Tabla_Modulos_Inactivos">
 <!-- RECUPERO TODOS LOS MODULOS INHABIITADOS PARA EL USUARIO -->
 <tr  >
    <th  valign="middle" align="left"nowrap="nowrap"  >
         
          <span class="BotonCabeceraModulos"  style="margin-left:0px !important">             
            <a class='BtnTransparenteOcultar' onclick='DesplazarBarraMenu()'> </a>    
            <%=Common.Utils.Armar_Icono("img/modulos/SVG/CONTROLMODULOS.svg", "IconosMaximizaModulos", "", "", "", "")%>
            
            <span id="txtInactivos" runat="server">   </span>                              
            <!--img src="img/up.png" border='0' align="absmiddle" width="6" height="6" id="imgInactivos"    onclick="ExpandInactivos('ModInactivos')" /--> 
          </span>
          
    </th>         
 </tr>
 <tr>
  <td valign="top" align="left"> 
   <TABLE cellpadding="0" cellspacing="0" border="0" id="ModInactivos" class="Menu_Links"  style="border:0px !important; margin:0px !important; padding:0px !important; " width="100%" align="left">
         
                 <asp:Repeater ID="RepeaterModulosInactivos" runat="server"    >
                   <HeaderTemplate> </HeaderTemplate>
                   <ItemTemplate >  
                          <tr id='Link<%=posmenu %>'  >
                              <td nowrap="nowrap" align="left" valign="middle" width="100%"  >                   
                                                    
                                <asp:LinkButton ID="LinkButton3"   OnCommand="ActualizarContenedor"  CommandArgument='<%# Eval("MenuName")+"@"+Eval("MenuMsnro") %>'
                                 runat="server" CssClass="ASPlink"   >                                       
                                     
                                     <%#Common.Utils.Armar_Icono("img/modulos/SVG/"+Convert.ToString(Eval("MenuName"))+".svg", "IconoModulo",ObjLenguaje.Label_Home((String)Eval("MenuTitle")),"", "") %>
                                 
                                      <div  >                                             
                                      <% posmenu++;%>
                                      <%# Agregar_Modulo_Visibilidad((String)Eval("MenuName"), false)%>  
                                                                  
                                      <%# ObjLenguaje.Label_Home((String)Eval("MenuTitle"))%>                                    
                                  
                                      </div>
                                </asp:LinkButton>   
                             
                              </td>
                               
                       </tr>                                                                     
                   </ItemTemplate>
                 </asp:Repeater> 
    </TABLE>
</td>
</tr>
</table>
</asp:Panel>  
<% }
    } %>
<!-- ------------------------- ACCESOS BASICOS  -->
      
<table width="100%" border="0" cellspacing="0" cellpadding="0" class="Menu_Links" id="CC_Tabla_Accesos">      
 

<tr style="height:20px" > 
<th  valign="middle" align="right"nowrap="nowrap"     >
 <span class='BotonCabeceraModulos' > 
  <a class='BtnTransparenteOcultar' onclick='DesplazarBarraMenu()'> </a>    
    <%=Common.Utils.Armar_Icono("img/modulos/SVG/CONTROLMODULOS.svg", "IconosMaximizaModulos", "", "", "", "")%>
     <span  style="margin-right:5px; vertical-align:middle" runat="server" id="Seccion_Accesos">  </span>  
 </span>  

 </th>         
</tr>


<!-- Busco los accesos  -->           
  <asp:Repeater ID="RepAccesos" runat="server"  >
           <HeaderTemplate> </HeaderTemplate>
           <ItemTemplate>       
           <tr id='Link<%=posmenu %>' onclick="Seleccionar('Link<%=posmenu %>','')"      > 
                   
                    <td  nowrap="nowrap" align="left" valign="middle" width="100%" >     
                      <asp:LinkButton ID="LinkButton2"   OnCommand="Actualizar_Accesos_XML"  CommandArgument='<%# (Int32)Eval("nroAcceso")%>'
                         runat="server" CssClass="ASPlink"    >                                                        
                            
                            <%#Common.Utils.Armar_Icono("img/modulos/SVG/LINK.svg", "IconoModulo", Convert.ToString(DataBinder.Eval(Container.DataItem, "Nombre")), "", "")%>
                           
                           <span >  
                             <% posmenu++; %>                                                           
                             <%#  ObjLenguaje.Label_Home(Convert.ToString(DataBinder.Eval(Container.DataItem, "Nombre")))%> 
                              
                           </span>
                           
                         
                         
                        </asp:LinkButton> 
                  
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
                  
      </ContentTemplate>                      
     </asp:UpdatePanel>
   <iframe src="" name='ifrmModulos' id="ifrmModulos" style="display:none;"></iframe>
 
    
    <script type="text/javascript">
      var CantidadLink =  <%= posmenu %> - 1;            
   
    </script>
  
 