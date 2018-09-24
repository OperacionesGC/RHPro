<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="PopUpChangePassword.aspx.cs"
    Inherits="RHPro.PopUp" StylesheetTheme="" meta:resourcekey="PageResource1" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>RHPro | Cambio de Contraseña</title>
 
<style type="text/css">
    
*{ margin:0; padding:0;}

body{  background:#ffffff; margin:0; padding:0;  }

form,input,select,textarea{margin:0; padding:0; color:#333;}



div.box {
margin:0 auto;
width:100%;
 
}

div.box h1 { 
color:#555;
font-size:10pt;
text-transform:uppercase;
 
border:0;
 
background:#ccc;
 
 
vertical-align:middle; 

}

div.box label {
width:100%;
display: block;
background:#ffffff;
border:0;
padding:10px 0 10px 0;
}

div.box label span {
 
color:#333;
font-size:9pt;
float:left;
width:210px;
text-align:right;
padding:5px 20px 0 0;
}

div.box label Err {
 
color:#FF0000;
font-size:15px;
float:left;
width:100%;
text-align:right;

}

div.box .input_text {
padding:5px;
width:200px;
background:#fff;
border: 1px solid #ccc;
 
}
/*
div.box .boton
{	
background:#cccccc;
border:1px solid #FFFFFF;
  
color: #333333;
text-decoration:none;
vertical-align:middle;
text-align:right;

padding-left:8px;
padding-right:8px;
padding-top:4px;
padding-bottom:4px; 
margin-left:2px;
margin-right:8px;
  
}
*/

.errorMess{
display: block;
color: #ff0000 !important;
font-size:12px;
float:left;
width:100%;
text-align:right;
 
margin-top:10px;
margin-bottom:10px;

 
}



  #passwordDescription {margin-bottom:1px; color:#000000;}
        #passwordStrength{ height:6px; display:block;float:left; font-size:4px; }
        .strength00{height:6px;        width:0%; }
        .strength0{height:6px;        width:10%; background-color:#000000;        }
        .strength1{height:6px;        width:25%;        background:#000000;}
        .strength2{height:6px;        width:45%;            background:#FF0;}
        .strength3{height:6px;        width:65%;        background:#FF0;}
        .strength4{height:6px;        width:85%;background:#399800;}
        .strength5{height:6px;        width:100%;background:#399800;}
        #limitador { width:135px; background-color:#666}
        #fuerza { width:10px !important; height:6px; margin-left:5px; font-size:10px; color:#333; background-color:#ffffff }
        .ayudapass { text-decoration:none; font-weight:bold; color:#000000;}
        
.Boton2 {  padding:4px !important; color:#ffffff; background-color: #666; border:1px #ffffff solid; margin:2px; text-decoration:none; 
         font-family:Calibri,Arial; font-size:10pt;  margin-right:2px; margin-top:4px; display: inline-block; margin-left:5px; width:50px; text-align:center }
.Boton2:hover { color:#FFFFFF; background-color:#0099CC }
    
.ListaErrores { width:100%; background-color:#ffffff; height:60px; }
            
</style>

<script type="text/javascript">

var ayudapass = '';

function tips() {
  return showModalDialog('../shared/asp/tips.asp','', 'center:yes;dialogWidth:25;dialogHeight:9');
}

function passwordStrength(password){
	var desc = new Array();
	desc[0] = "<asp:Label runat="server" Text="Insegura" meta:resourcekey="lbinsegura"></asp:Label>";
	desc[1] = "<asp:Label runat="server" Text="Débil" meta:resourcekey="lbdebil"></asp:Label>";
	desc[2] = "<asp:Label runat="server" Text="Regular" meta:resourcekey="lbregular"></asp:Label>";
	desc[3] = "<asp:Label runat="server" Text="Aceptable" meta:resourcekey="lbaceptable"></asp:Label>";
	desc[4] = "<asp:Label runat="server" Text="Fuerte" meta:resourcekey="lbfuerte"></asp:Label>";
	desc[5] = "<asp:Label runat="server" Text="Optima" meta:resourcekey="lboptima"></asp:Label>";
	var score   = 0;
	
	// si no hay nada escrito, borro todo
	if (password.length == 0) {
		document.getElementById("passwordDescription").innerHTML = '';
		document.getElementById("passwordStrength").className = "strength00";
		document.getElementById("fuerza").style.display = "none";
		return;
	} else {
		document.getElementById("fuerza").style.display = "inline";
	}
	
	//if password bigger than 6 give 1 point
	if (password.length > 6) score++;
	
	//if password has both lower and uppercase characters give 1 point 
	if ( ( password.match(/[a-z]/) ) && ( password.match(/[A-Z]/) ) ) score++;
	
	//if password has at least one number give 1 point
	if (password.match(/\d+/)) score++;
	
	//if password has at least one special caracther give 1 point
	if ( password.match(/.[!,@,#,$,%,^,&,*,?,_,~,-,(,)]/) ) score++;
	
	//if password bigger than 12 give another 1 point
	if (password.length > 12) score++;
	
	document.getElementById("passwordDescription").innerHTML = ayudapass + desc[score];
	document.getElementById("passwordStrength").className = "strength" + score;	
	
	document.getElementById("errorMess").innerHTML = "";
}
</script>


</head>
<body>
 

<!-- ************************************************ -->
 <div class="box">
     <form id="form1" runat="server" style="font-family:Arial; color:#666 !important;" defaultbutton="btnConfirmar">
            <h1> 
            <!--img src="img/Modulos/Gadget2.png" border="0" align="absmiddle" style="margin-left:4px; margin-right:9px; margin-top:0px;//margin-top:0px"-->
            <div style=" background-color:#cccccc; border-bottom:1px solid #666; height:30px; width:100%; vertical-align: middle; line-height:2; ">
            
                      
             <span style="margin-left:5px"> <asp:Label ID="title" runat="server"   ></asp:Label> </span>
             
        </div> 
                  
            </h1>
            <label>
               <span>  <% Response.Write(ObjLenguaje.Label_Home("Ingrese su anterior contraseña")); %>:  </span>
               <!--input type="text" class="input_text" name="name" id="name"/-->               
               <input type="password" id="txtOldPassword" runat="server"  class="input_text" />
            </label>
             <label>
               <span> <% Response.Write(ObjLenguaje.Label_Home("Ingrese su nueva contraseña")); %>:</span>
               <!--input type="text" class="input_text" name="email" id="email"/-->                
                <input type="password" id="txtNewPassword" runat="server"   class="input_text" onKeyUp="passwordStrength(this.value)"/>
            </label>
             <label>
                <span>  <% Response.Write(ObjLenguaje.Label_Home("Vuelva a ingresar su nueva contraseña")); %>:</span>
                <!--input type="text" class="input_text" name="subject" id="subject"/-->
                 <input type="password" id="txtVerifyPassword" runat="server" class="input_text"/>
                
            </label>
            
            
            <label class="ListaErrores">
             
             <div style="display:none;"  id="fuerza"> 
                     <div style="width:130px">     Fuerza de Contraseña:   </div>
                   
                     <div id="limitador">
                        <div id="passwordStrength" class="strength00"></div>
                     </div>
                     
                     <div id="passwordDescription"></div> 
                     
              </div>                     
              <asp:Label ID="errorMess" runat="server" CssClass="errorMess" />              
            </label>  
            <table cellpadding="0" cellspacing="0" border="0" width="100%" align="center" >
                 <tr>
               
                 <td valign="middle" align="right" style="height:40px; border-top:1px solid #999; background-color:#cccccc " >  
                   
                           
                      <asp:LinkButton ID="btnConfirmar" runat="server" OnClick="btnConfirmar_Click" Text="Guardar"    CssClass="Boton2">Guardar</asp:LinkButton>
                      <asp:LinkButton runat="server" CausesValidation="False" ID="btnCancel" Text="Cancelar"               
                OnClientClick="javascript:window.close();" CssClass="Boton2" >Salir</asp:LinkButton>
             
             
             
             
                  </td>
                  
                 </tr>
              </table>
           
           
    </form>
       
            
         </div>
<!-- ************************************************ -->

 

 




</body>
</html>
