<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="CustomLogin.ascx.cs"
    EnableViewState="true" Inherits="RHPro.Controls.CustomLogin" %>

 

<script type="text/javascript">
  

    //Abre popUp de Politics
    function Politic_show() {
        window.open("PopUpPolitics.aspx", "Ventana", 'height=215,width=450,status=yes,toolbar=no,menubar=no,location=no,resizable=yes,scrollbars=yes,left=5,top=5');
    }

    function KeyPress(e) {
   
        var esIE = (document.all);
        tecla = (esIE) ? event.keyCode : e.which;

        if (tecla == 13)
            javascript: __doPostBack('doLogin_Click', '');

    };
    
    
</script>
<style>
    .Clase_TextBox     { border:0px; background-color: #CCCCCC; }
</style>

<style>
 

.Separador { color:#FFF; font-family:Tahoma; font-size:11pt; background:url(img/Fondo_Menu.png) repeat-x top; height:37px; cursor:pointer }
.Menu_Ppal TR { color:#FFF; font-family:Tahoma; font-size:11pt; background:url(img/Fondo_Menu.png) repeat-x top; height:37px; cursor:pointer }
.Menu_Links TR { color:#FFF; font-family:Tahoma; font-size:9pt; background:url(img/Fondo_Menu.png) repeat-x top; height:37px; cursor:pointer }

.DetalleEmpresa{
	 color:#CCCCCC; font-size:7.5pt; font-family:Tahoma; margin-top:4px;
	} 
.TituloBase{ color:#FFFFFF; font-size:7pt; font-family:Tahoma;}

.Detalle{ color:#CCC; font-size:12pt; font-family:Tahoma;  }

.user{}
 
 
 
#Btn_Loguin{ cursor:pointer; } 

#Globo_Loguin{  display:none; float:left; z-index:1001; position: absolute; margin-left:-80px;}
#Glogo_Centro {background:url(img/Loguin/Globo_Centro.png) repeat-y center;} 


 
.fecha{ color:#333; font-family:Tahoma; font-size:9pt; }

.Input{ background-color:transparent; border-width:0px; border-color:transparent;}
.Input_Centro{background:url(img/input/input_centro.png) repeat-x center;} 

#Ingresar{ font-family:Arial; font-size:10pt; font-weight:normal !important}
.Select{ background-color:none; border-width:0px; border-color:none;} 

.BotonGris{background:url(img/BotonGris/btn_centro.png) repeat-x center;width:50px; text-align:center; } 
.BotonGris a:link{ color: #333333; font-family:Tahoma; font-size: 8pt; text-decoration:none } 
.BotonGris a:hover{ color:#FF0000; } 
.BotonGris a:visited{ color: #333333; font-family:Tahoma; font-size: 8pt; text-decoration:none } 

.info{ color:#333333; font-family:Tahoma; font-size:9pt}

.Err_MH{visibility:hidden} 
.Err_MH_Visible{cursor:pointer;position:absolute; top:10%; left:40%; padding:12px; border:1px solid #333333; background-color:#FFFFFF; visibility:visible}

 
 
</style>
 

 <script>
     function Ocultar(obj) {
         obj.innerHTML = "";
         obj.style.visibility = "hidden";
     }


     function PopUp_Cerrar() {
         var obj = document.getElementById("PopUp_NewHome");
         var fondo = document.getElementById("PopUp_FondoTransparente");
         obj.style.display = "none";
         fondo.style.display = "none";
     }

     function PopUp_Abrir() {
         
         var obj = document.getElementById("PopUp_NewHome");
         var fondo = document.getElementById("PopUp_FondoTransparente");
         var campofoco = document.getElementById("ctl00_content_cLogin_txtUserName");
    
        // Cerrar_Globo('Globo_Idiomas');
         obj.style.display = "";
         fondo.style.display = "";
         campofoco.focus();
     }
 </script>
 
 
<!-- ############################## FONDO TRANSPARENTE ##################################----->
<DIV   ID="PopUp_FondoTransparente"  Class="PopUp_FondoTransparente" style='display:none'></DIV>
<!-- ############################## CONTENEDOR FLOTANTE ##################################----->

<DIV id="PopUp_NewHome"     style='display:none'>  
    
  <TABLE cellpadding="0" cellspacing="0" border="0" class="PopUp_NewHome">
  <tr class="PopUp_Cabecera">
   <td>
   
   <asp:PlaceHolder ID="PopUp_Cabecera" runat="server"></asp:PlaceHolder>  
   
   <span class="cerrarVentana"  onclick="PopUp_Cerrar()">X</span>  
   
   </td>
  </tr>
  <tr class="PopUp_DataUser">
   <td>
     <asp:Panel ID="Campos_Formulario" runat="server">
         
         <span id="LabelUsr" runat="server"></span>            
         
         <DIV class="popInput">         
             <input ID="txtUserName" name="txtUserName"  runat="server"   type="text"    />                                       
             
         </DIV>
        
         <span id="LabelPass" runat="server"></span>    
         
         <DIV class="popInput">         
               <input id="txtPassword" name="txtPassword" runat="server" type="password"    onkeypress="KeyPress(event)"  />             
        </DIV>
        
        </asp:Panel>
      <!-- CONTROL DE PROGRESO -->
        <span >      
        <asp:UpdateProgress ID="UpdateProgress" runat="server"  Visible="true"  >
        <ProgressTemplate>                  
               <img  src="img/miniloader.gif" align="absmiddle"/>                      
        </ProgressTemplate>               
        </asp:UpdateProgress>       
        </span>
      <!-- ------------------------------------------- -->   
      <asp:Panel ID="Info_User_Logueado" runat="server">  
      </asp:Panel>
      <asp:Panel ID="Info_Base_Seleccionada" runat="server">  
      </asp:Panel> 
      
      
   </td>
  </tr>
  <tr class="PopUp_BD">
   <td>
   
      <asp:UpdatePanel ID="Update_Bases" runat="server" UpdateMode="Conditional">
            <ContentTemplate> 
                   <asp:Panel ID="Combo_Bases_Formulario" runat="server">
                         <asp:PlaceHolder ID="PopUp_ImagenUsr" runat="server"><!--img src="img/PopUp_User.png" ></img--></asp:PlaceHolder>         
                         <asp:Label ID="TituloSelBase" runat="server"> </asp:Label>
                         <asp:ListBox ID="cmbDatabase" runat="server"  CssClass="PopUp_BD_Combo"   > </asp:ListBox>                              
                  </asp:Panel>
                   <asp:Panel ID="Informes_Error" runat="server" CssClass="PopUp_Informes_Error">  
      </asp:Panel>
      
       </ContentTemplate>
        <Triggers>        
           <asp:AsyncPostBackTrigger ControlID="PopUp_BotonControlar" EventName="Click" /> 
        </Triggers>
     </asp:UpdatePanel>
 
  
   </td>
  </tr>
  <tr  class="PopUp_Piso">
   <td>
        
        <asp:UpdatePanel runat="server" ID="miup">
         <ContentTemplate>
            
              <asp:LinkButton ID="PopUp_BotonLogin" runat="server"  CssClass="Boton_Cuadrado"  OnClick="doLogin_Click" Text=""   >                    
              </asp:LinkButton>  
              <asp:LinkButton ID="PopUp_BotonControlar" runat="server"   CssClass="Boton_Cuadrado"   OnClick="doLogin_Control"   Text=""  >
              </asp:LinkButton>   
              <asp:LinkButton ID="CerrarSesion"  OnClick="btnLogOut_Click" CssClass="Boton_Cuadrado"   runat="server"  Text=""> 
              </asp:LinkButton>      
              
            </ContentTemplate>
         <Triggers>
             <asp:AsyncPostBackTrigger ControlID="PopUp_BotonControlar" EventName="Click" />      

         </Triggers>
        </asp:UpdatePanel>
         
                             
   </td>
  </tr>
  </TABLE>
</DIV>
         
        
 
 
 <!-- ##############################CONTENEDOR##################################----->
 
 
 <!-- ##############################FIN CONTENEDOR##################################----->
    
<!-- -- --------- LOGUEO-------------------------------->
<table width="250" border="0" cellspacing="0" cellpadding="0" id="Globo_Loguin"  align="left" >
  <tr>
    <td align="center" valign="bottom"><img src="img/Loguin/Globo_Tope.png"></td>
  </tr>
  <tr>
    <td height="48" id="Glogo_Centro" align="center">
      
     <img src="img/Close.png" style=" margin-left:5px; cursor:pointer" onclick="Cerrar_Globo('Globo_Loguin');"
        onmouseover="this.src = 'img/Close-hover.png'" onmouseout="this.src =' img/Close.png'" >
    
 <!-- -------------------------------------------->
 

<!-- -------------------------------Loguin OFF ------------------->
 <div id="LoginON" runat="server">
 </div> 
<!-- -----------------------------Loguin ON ----------------- -->
    <div id="LoginOFF" runat="server" style="display: block;" >

<table width="188" border="0" cellspacing="0" cellpadding="0" align="center" valign="top"   >
  <tr>
    <td  align="left" valign="middle">&nbsp;</td>
    <td width="100%" align="left" valign="middle" nowrap="nowrap">
     
     <asp:Label  ID="Bienvenido" runat="server">  </asp:Label>
      
     <br>    <br>
     </td>
  </tr>
  <tr>
    <td align="left" valign="middle">&nbsp;</td>
    <td align="left" valign="middle" nowrap="nowrap" class="info">
     <img src="img/usuario.png"  align="absmiddle" style="margin-right:6px;">    
        <label id="lblUser"   runat="server">   </label>
    </td>
  </tr>
  <tr>
    <td align="left" valign="middle">&nbsp;</td>
    <td align="left" valign="middle" nowrap="nowrap" class="info">
       <img src="img/base.png"  align="absmiddle" style="margin-right:6px;">   
       <asp:Label ID="LabelBaseSeleccionada" runat="server" ></asp:Label>
    </td>
  </tr>
  <tr>
    <td colspan="2" align="center" valign="middle"> 
  
   <div id="linkLogOff">
  <table   border="0" cellspacing="0" cellpadding="0" style="display:inline-block;">
   <tr>
      <td width="1" align="right" valign="middle" ><img src="img/BotonGris/btn_izq.png" ></td>
      <td width="60" align="center" valign="middle" class="BotonGris" nowrap="nowrap" >   
                        
                <asp:LinkButton ID="btnLogOut"  OnClick="btnLogOut_Click" runat="server" Text=""  > 
                     
                </asp:LinkButton>         
             
       </td>                       
       <td width="1" align="left" valign="middle"><img src="img/BotonGris/btn_der.png" ></td>
  </tr>
</table>    
</div>                  
            
            

    </td>
  </tr>
</table>
 </div>
 
 
 
 <!-- -------------FIN LOGUEO ---------->
  </td>
  </tr>
  <tr> 
    <td align="center" valign="top"><img src="img/Loguin/Globo_Piso.png" /></td>
  </tr>
</table>

  
 
 <!-- ---------------------- -->
 
  
   <DIV id="Err_MH" runat="server"  class="Err_MH" onclick="Ocultar(this)">  
 </DIV>
 