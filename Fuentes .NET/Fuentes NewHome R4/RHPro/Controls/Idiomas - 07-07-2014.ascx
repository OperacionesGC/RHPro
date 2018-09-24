<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="Idiomas.ascx.cs" Inherits="RHPro.Controls.Idiomas" %>

 
<script type="text/javascript">

    function Listar_Idiomas() {

        if (document.getElementById("Globo_Idiomas").style.visibility != "visible")
            document.getElementById("Globo_Idiomas").style.visibility = "visible";
        else
            document.getElementById("Globo_Idiomas").style.visibility = "hidden";
    }


   

 
</script>

 
 
 
<!-- ############################## FONDO TRANSPARENTE ##################################----->
<DIV ID="PopUp_FondoTransparenteLeng"  Class="PopUp_FondoTransparente" style="display:none"></DIV>

 <!-- ##############################CONTENEDOR##################################----->
 
 
       
   <!-- -- --------- IDIOMAS -------------------------------->
  <TABLE cellpadding="0" cellspacing="0" border="0" id="Globo_Idiomas"  class="Globo_Lenguajes TablaIdioma">
  <tr class="PopUp_Cabecera">
   <td>   
     <span id="TituloIdi" runat="server"></span>   
     <span class="cerrarVentana" id="BtnCierraVentana" runat="server"  onclick="Cerrar_Globo('Globo_Idiomas');">X</span>     
   </td>
  </tr>
  <tr class="PopUp_DataUser">
   <td align="center">
    
<!-- --------------------------ARMO EL COMBO DE IDIOMAS ACTIVOS-----------------------------------------------------------   -->      
     
      <asp:Repeater ID="Repeater1" runat="server"  >
           <ItemTemplate>                               
            <asp:LinkButton  id="Btn_IdiomaAcc"   runat="server" OnClick="Idioma_Click"  
                 CommandArgument='<%# Eval("lencod") +"@"+ Eval("lendesabr") + "@~/img/Flags/flag_" + Eval("lencod") + ".png" %>' >                      
              <div class="EtiquetaIdioma">
               <img src="img/Flags/flag_<%# Eval("lencod") %>.png" align="absmiddle" border="0" style="margin-right:3px"   />                    
                  <%# Eval("lendesabr") %>                     
               </div>  
          </asp:LinkButton>                     
                                        
           </ItemTemplate>
      </asp:Repeater>

<!-- --------------------------------------------------------------------------------------------------------------   -->
    
                          
   </td>
  </tr>
  </TABLE>
 