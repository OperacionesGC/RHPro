<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="CustomLogin.ascx.cs"
    EnableViewState="true" Inherits="RHPro.Controls.CustomLogin" %>

 

<script type="text/javascript">
    /*  $(document).ready(function() {
        $('#ctl00_content_<%= this.ID %>_cmbDatabase').sSelect();
    }); */




    //Abre popUp de Politics
    function Politic_show() {
        window.open("PopUpPolitics.aspx", "Ventana", 'height=215,width=450,status=yes,toolbar=no,menubar=no,location=no,resizable=yes,scrollbars=yes,left=5,top=5');
    }
    
  
    
    
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

.user{ color:#333; font-family:Tahoma; font-size:11pt; font-weight:bold;}

#user_izq { background:url(img/Loguin/user_izq.png) no-repeat right bottom;}
#user_centro1 {background:url(img/Loguin/user_centro.png) repeat-x center}
#user_centro2 {background:url(img/Loguin/user_centro.png) repeat-x center}
#user_der { background:url(img/Loguin/user_der.png) no-repeat right bottom;  }

 
 
#Btn_Loguin{ cursor:pointer; }
#Btn_Loguin:hover TD#user_izq{ background:url(img/Loguin/user_izq_hover.png) no-repeat right bottom;}	 
#Btn_Loguin:hover TD#user_centro1{ background:url(img/Loguin/user_centro_hover.png); color:  #999}	
#Btn_Loguin:hover TD#user_centro2{ background:url(img/Loguin/user_centro_hover.png); color: #999}	
#Btn_Loguin:hover TD#user_der{ background:url(img/Loguin/user_der_hover.png) no-repeat right bottom;}

 


#Globo_Loguin{ visibility:hidden; float:left; z-index:1001; position: absolute; margin-left:-80px;}
#Glogo_Centro {background:url(img/Loguin/Globo_Centro.png) repeat-y center;} 


 
.fecha{ color:#333; font-family:Tahoma; font-size:9pt; }

.Input{ background-color:transparent; border-width:0px; border-color:transparent;}
.Input_Centro{background:url(img/input/input_centro.png) repeat-x center;} 
 
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
         
 </script>
 
  
 <DIV id="Err_MH" runat="server"  class="Err_MH" onclick="Ocultar(this)">  
 </DIV>
 
 
 
 
 <!-- ##############################CONTENEDOR##################################----->
 <TABLE cellpadding="0" cellspacing="0" border="0"  >
 <TR> 
  <TD> 
 <!-- ---------------------- -->
 <table width="90" border="0" cellspacing="0" cellpadding="0" class="user" id="Btn_Loguin" onclick="Abrir_Globo('Globo_Loguin');"  >
      <tr>
        <td width="6" align="right" valign="middle" id="user_izq"><div style="width:5px; height:33px">&nbsp;</div> </td>
         <td   id="user_centro1" ><asp:Image  ImageUrl="~/img/login.png"  ID="Candado"  runat="server" style=" margin-left:2px; margin-right:2px;" /></td>
        <td    id="user_centro2" nowrap align="center" > 
         <span style="margin-right:10px; margin-left:5px; text-align:center">
          <asp:Label id="Ingresar"  runat="server"     >         
           <%              
               if (!Common.Utils.IsUserLogin)
                 Response.Write( ObjLenguaje.Label_Home("Ingresar") ); 
               else
                   Response.Write(Common.Utils.SessionUserName);   
               
           %> 
          </asp:Label> 
          
              </span>
             
        </td>
        <td width="14" align="left" valign="middle" id="user_der" > <div style="width:14px"></div></td>
      
      </tr>
</table>
<!-- ---------------------- -->
  </TD>  
 </TR>
 <TR>
     <TD align="left" valign="bottom">  
 <!-- ---------------------- -->
 
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
         
            <asp:Panel ID="panelLogin" runat="server" DefaultButton="btnLogin" meta:resourcekey="panelLoginResource1" Font-Size="XX-Small">
            
            
<table width="200" border="0" cellspacing="0" cellpadding="0" align="left" valign="top"  style="margin-left:20px;"  >
  <tr>
    <td width="24" align="left" valign="middle">    </td>
    <td width="163" align="left" valign="middle">            
       <label>
             
            <asp:Label id="TitIngresar"  runat="server"   > 
     <% Response.Write( ObjLenguaje.Label_Home("Ingresar") ); %> 
            </asp:Label>
        </label>
    </td>
  </tr>
  <tr>
    <td width="24" align="left" valign="middle">    </td>
    <td width="163" align="left" valign="middle">            
       <asp:TextBox ID="TextBox1" runat="server" Style="display: none;" meta:resourcekey="TextBoxResource1"></asp:TextBox>

<!-- USUARIO -->
<table   border="0" cellspacing="0" cellpadding="0" style="margin-top:4px">
  <tr>
   <td width="1" align="left" valign="middle" >
    <img src="img/usuario.png"  align="absmiddle" style="margin-right:6px;">       
    </td>
    <td width="1" align="right" valign="middle" ><img src="img/input/input_izq.png" ></td>
    <td width="1" align="left" valign="middle" class="Input_Centro">
        
        <input ID="txtUserName"    runat="server" class="Input"  style="width: 150px;   margin: 5px 0px 5px 0px;" type="text" />
    </td>
    <td width="1" align="left" valign="middle"><img src="img/input/input_der.png" ></td>
  </tr>
</table>
       
       
    </td>
  </tr>
  <tr>
    <td align="left" valign="middle">&nbsp;</td>
    <td align="left" valign="middle"> 
  
<!-- PASSWORD -->   
<table   border="0" cellspacing="0" cellpadding="0" style="margin-top:10px"  >
  <tr>
  <td width="1" align="left" valign="middle" >
    <img src="img/password.png"  align="absmiddle" style="margin-right:6px;">       
    </td>
    <td width="1" align="right" valign="middle" ><img src="img/input/input_izq.png" ></td>
    <td width="1" align="left" valign="middle" class="Input_Centro">
    <input id="txtPassword" runat="server" type="password" style="width: 150px; margin: 0px 0px 5px 0px;" class="Input" />
       
    </td>
    <td width="1" align="left" valign="middle"><img src="img/input/input_der.png" ></td>
  </tr>
</table>
   
    </td>
  </tr>
  <tr>
    <td align="center" valign="middle">
   
    </td>
    <td align="left" valign="middle">

<!-- SELECCION DE LAS BASES -->
<table  border="0" cellspacing="0" cellpadding="0" style="margin-top:10px">
  <tr>
    <td width="1" align="left" valign="middle" >
    <img src="img/base.png"  align="absmiddle" style="margin-right:6px;">       
    </td>    
    <td width="1" align="right" valign="middle" ><img src="img/input/input_izq.png" ></td>
    <td width="1" align="left" valign="middle" class="Input_Centro">            
      <asp:DropDownList ID="cmbDatabase" runat="server"  
                    EnableTheming="True" meta:resourcekey="cmbDatabaseResource1" name="cmbDatabase" 
                    Width="154px" CssClass="Select" 
                    OnSelectedIndexChanged="cmbDatabase_SelectedIndexChanged" >
            </asp:DropDownList>         
            <asp:Panel ID="PanellstDatabase" runat="server" Height="55px" Width="165px" 
                        Font-Size="XX-Small" DefaultButton="btnChangeDB">
                        &nbsp;&nbsp;<asp:ListBox ID="lstDatabase" runat="server" CssClass="borde" Height="55px" 
                            OnSelectedIndexChanged="lstDatabase_SelectedIndexChanged" Width="154px">
                        </asp:ListBox>                
            </asp:Panel>      
      <td width="1" align="left" valign="middle"><img src="img/input/input_der.png" ></td>
  </tr>
</table>      
                 
    </td>
  </tr>
  
  <tr>
    <td colspan="2" align="center" valign="middle"  nowrap="nowrap" width="100%" >      
<TABLE cellpadding="0" cellspacing="0" border="0" align="center">    
<TR><TD valign="top" align="center">

 <!-- ACEPTAR -->
 <table  border="0" cellspacing="0" cellpadding="0" style="margin-top:10px; ">
  <tr>
    <td width="1" align="right" valign="middle" ><img src="img/BotonGris/btn_izq.png" ></td>
    <td width="1" align="left" valign="middle" class="BotonGris">
     
             <asp:LinkButton ID="btnLogin" runat="server"  SkinID="Aceptar"  OnClick="doLogin_Click"    >
                <% Response.Write( ObjLenguaje.Label_Home("Aceptar") ); %> 
             </asp:LinkButton>              
     
     <td width="1" align="left" valign="middle"><img src="img/BotonGris/btn_der.png" ></td>
  </tr>
</table>     
</TD>         
<TD valign="top" align="center">
         
<!-- POLITICAS --> 
 <table   border="0" cellspacing="0" cellpadding="0" style="margin-top:10px;">
  <tr>
    <td width="1" align="right" valign="middle" ><img src="img/BotonGris/btn_izq.png" ></td>
    <td width="1" align="left" valign="middle" class="BotonGris"> 
              <asp:LinkButton ID="btPolitics" runat="server"                       
                    OnClientClick="Politic_show();return false;"  SkinID="Políticas"    >
                <% Response.Write( ObjLenguaje.Label_Home("Políticas") ); %> 
             </asp:LinkButton> 
  <td width="1" align="left" valign="middle"><img src="img/BotonGris/btn_der.png" ></td>
  </tr>
</table>  

</TD>
</TR>   
</TABLE>        
               <span style="padding-left: 0px;">
                  <asp:LinkButton ID="btnChangeDB" runat="server" Font-Size="XX-Small" 
                  Style="padding-left: 0px;" 
                   OnClick="doChangeDB_Click"/>
                </span>
            
    </td>
  </tr>
</table>
            </asp:Panel>
        
        <input type="text" style="display: none" />
    </div>
 


<!-- -----------------------------Loguin ON ----------------- -->
    <div id="LoginOFF" runat="server" style="display: block;" >

<table width="188" border="0" cellspacing="0" cellpadding="0" align="center" valign="top"   >
  <tr>
    <td  align="left" valign="middle">&nbsp;</td>
    <td width="100%" align="left" valign="middle" nowrap="nowrap">
     
     <asp:Label  ID="Bienvenido" runat="server">
        <% Response.Write( ObjLenguaje.Label_Home("Bienvenido") ); %> 
       </asp:Label>
      
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
                        
                <asp:LinkButton ID="btnLogOut"  OnClick="btnLogOut_Click" runat="server"  > 
                       <% Response.Write( ObjLenguaje.Label_Home("Salir") ); %> 
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
 
 
    </TD>
 </TR>
 </TABLE>
 
  