<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="CustomLogin.ascx.cs"
    EnableViewState="true" Inherits="RHPro.Controls.CustomLogin" %>

 

<script type="text/javascript" >
  

    //Abre popUp de Politics
    function Politic_show() {
        window.open("PopUpPolitics.aspx", "Ventana", 'height=215,width=450,status=yes,toolbar=no,menubar=no,location=no,resizable=yes,scrollbars=yes,left=5,top=5');
    }

    var ErrorFaltaUsr = '<%= ObjLenguaje.Label_Home("Falta Usuario") %>';
    var ErrorFaltaPass = '<%= ObjLenguaje.Label_Home("Falta Contraseña") %>';

    function KeyPress(e) {
    
        var esIE = (document.all);
        tecla = (esIE) ? event.keyCode : e.which;

        //ctl00_content_cLogin_txtUserName
        //ctl00_content_cLogin_txtPassword

        if (tecla == 13) {
            if ((document.getElementById("ctl00_content_cLogin_txtUserName").value == "") || (document.getElementById("ctl00_content_cLogin_txtPassword").value == ""))
                if (document.getElementById("ctl00_content_cLogin_txtUserName").value == "")
                alert(ErrorFaltaUsr);
            else {
                alert(ErrorFaltaPass);
                document.getElementById("ctl00_content_cLogin_txtPassword").focus();
            }
            else {               
                javascript: __doPostBack('doLogin_Click', 'loginJS');                            
            }
        // javascript: __doPostBack('UpdatePanel_LGN', '');
        
        
        
       
        }

    };
    
    
    
    
</script>
<style>
    .Clase_TextBox     { border:0px; background-color: #CCCCCC; }
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
         //var campofoco = document.getElementById("<%//=txtUserName.ClientID%>");
         var campofoco = document.getElementById("ctl00_content_cLogin_txtUserName");
          
        
         obj.style.display = "";
         fondo.style.display = "";
         campofoco.focus();
     }

     function FocoUser() {
         document.getElementById("<%=txtUserName.ClientID%>").focus(); 
     }

     function FocoPass() {
         document.getElementById("<%=txtPassword.ClientID%>").focus();
     }
     
 
 
 </script>
 
       
        <asp:UpdatePanel ID="UpdatePanel_LGN" runat="server"      >
         <ContentTemplate> 

<!--DIV id="PopUp_NewHome"  -->  
    
  <TABLE cellpadding="0" cellspacing="0" border="0" class="PopUp_NewHome" >
  <tr class="PopUp_DataUser">
   <td>
     <asp:Panel ID="Campos_Formulario" runat="server">         
         <span id="LabelUsr" runat="server"></span>            
         
         <DIV class="popInput">         
            
            <%=Common.Utils.Armar_Icono("img/modulos/SVG/LOGINUSER.svg", "IconoLogin", "", "align='absmiddle'  onfocus=\"FocoUser()\"", "")%>
               <input ID="txtUserName" name="txtUserName" runat="server" onclick="this.focus();"   onkeypress="KeyPress(event)"   autocomplete="off" type="text"    />                                       
                
               
         </DIV>
        
         <span id="LabelPass" runat="server"></span>    
         
         <DIV class="popInput">
            
            <%=Common.Utils.Armar_Icono("img/modulos/SVG/LOGINPASS.svg", "IconoLogin", "", "align='absmiddle'  onfocus=\"FocoPass()\" ", "")%>      
            <input id="txtPassword" name="txtPassword"  runat="server" type="password" onclick="this.focus()"  onkeypress="KeyPress(event)"    />             
            
        </DIV>
        
     </asp:Panel>
      <!-- CONTROL DE PROGRESO -->
        <span >      
        <asp:UpdateProgress ID="UpdateProgress" runat="server"  Visible="true" AssociatedUpdatePanelID="UpdatePanel_LGN" >
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
   
      
                   <asp:Panel ID="Combo_Bases_Formulario" runat="server">
                         <asp:PlaceHolder ID="PopUp_ImagenUsr" runat="server"><!--img src="img/PopUp_User.png" ></img--></asp:PlaceHolder>         
                           
                           <%=Common.Utils.Armar_Icono("img/modulos/SVG/LOGINBASE.svg", "IconoLogin", "", "align='absmiddle'", "IconoBases")%>
                           
                           <asp:Label ID="TituloSelBase" runat="server"> </asp:Label>
                        
                         <asp:ListBox ID="cmbDatabase" runat="server"  CssClass="PopUp_BD_Combo" SelectionMode="Single"   >
                         </asp:ListBox>                              
                         
                  </asp:Panel>
                   <asp:Panel ID="Informes_Error" runat="server" CssClass="PopUp_Informes_Error">  
      </asp:Panel>
        
 
  
   </td>
  </tr>
  <tr  class="PopUp_Piso" >
   <td >
    
            
              <asp:LinkButton ID="PopUp_BotonLogin" runat="server"  CssClass="Boton_Cuadrado"  OnClick="doLogin_Click" Text=""    >                    
              </asp:LinkButton>  
 
              
              <asp:LinkButton ID="PopUp_BotonControlar" runat="server"   CssClass="Boton_Cuadrado"   OnClick="doLogin_Control"   Text=""  >
              </asp:LinkButton> 
              
              <asp:LinkButton ID="PopUp_Limpiar"  OnClick="btnLogOut_Limpiar" CssClass="Boton_Cuadrado"   runat="server"  Text=""> 
              </asp:LinkButton>      

                
              <asp:LinkButton ID="CerrarSesion"  OnClick="btnLogOut_Click" CssClass="Boton_Cuadrado"   runat="server"  Text=""> 
              </asp:LinkButton>      
              
      
         
      
   </td>
  </tr>
  </TABLE>
  
  
  
        </ContentTemplate>
 
 
        </asp:UpdatePanel>
<!--/DIV-->
         
        
 
<DIV id="Err_MH" runat="server"  class="Err_MH" onclick="Ocultar(this)">  </DIV>
 