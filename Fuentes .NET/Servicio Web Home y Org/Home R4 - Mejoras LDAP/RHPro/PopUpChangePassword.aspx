<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="PopUpChangePassword.aspx.cs"
    Inherits="RHPro.PopUp" StylesheetTheme="" meta:resourcekey="PageResource1" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>RHPro | Cambio de Contraseña</title>
 
<style type="text/css">
    
*{ margin:0; padding:0;}

body{ font:100% normal Arial, Helvetica, sans-serif; background:#161712; margin:0; padding:0;  }

form,input,select,textarea{margin:0; padding:0; color:#ffffff;}

div.box {
margin:0 auto;
width:500px;
background:#222222;
position:relative;
top:20px;
border:1px solid #262626;
}

div.box h1 { 
color:#ffffff;
font-size:13px;
text-transform:uppercase;
/*padding:5px 0 5px 5px;*/
border-bottom:1px solid #161712;
border-top:1px solid #161712; 
background:#262626;
 /*background-image: url(img/top-centro.png);*/
 background-image: url(img/top-centro.png);
 

vertical-align:middle; 
//vertical-align: top; 
}

div.box label {
width:100%;
display: block;
background:#1C1C1C;
border-top:1px solid #262626;
border-bottom:1px solid #161712;
padding:10px 0 10px 0;
}

div.box label span {
display: block;
color:#bbbbbb;
font-size:15px;
float:left;
width:210px;
text-align:right;
padding:5px 20px 0 0;
}

div.box label Err {
display: block;
color:#FF0000;
font-size:15px;
float:left;
width:100%;
text-align:right;

}

div.box .input_text {
padding:10px 10px;
width:200px;
background:#262626;
border-bottom: 1px double #171717;
border-top: 1px double #171717;
border-left:1px double #333333;
border-right:1px double #333333;
}

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


div.box .errorMess{
display: block;
color: Red;
font-size:12px;
float:left;
width:100%;
text-align:right;
padding:5px 20px 0 0;
margin-top:10px;
margin-bottom:10px;
}



  #passwordDescription {margin-bottom:1px; color:#000000;}
        #passwordStrength{ height:2px; display:block;float:left; font-size:4px; }
        .strength00{height:2px;        width:0%; }
        .strength0{height:2px;        width:10%; background-color:#000000;        }
        .strength1{height:2px;        width:25%;        background:#000000;}
        .strength2{height:2px;        width:45%;            background:#FF0;}
        .strength3{height:2px;        width:65%;        background:#FF0;}
        .strength4{height:2px;        width:85%;background:#399800;}
        .strength5{height:2px;        width:100%;background:#399800;}
        #limitador { width:135px; background-color:#666}
        #fuerza { width:10px !important; height:2px; margin-left:5px; font-size:10px; }
        .ayudapass { text-decoration:none; font-weight:bold; color:#000000;}
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
    <form id="form1" runat="server" style="font-family:Arial; color: #FFFFFF;" defaultbutton="btnConfirmar">

<!-- ************************************************ -->
 <div class="box">
            <h1> 
            <!--img src="img/Modulos/Gadget2.png" border="0" align="absmiddle" style="margin-left:4px; margin-right:9px; margin-top:0px;//margin-top:0px"-->
            <div style="background:url(img/Fondo_Menu_Press2.png) no-repeat left;height:28px; width:100%; vertical-align: top">
            
              <img src="img/Modulos/Gadget2.png" border='0' align="bottom"style="margin-left:4px; margin-right:9px;  margin-top:0px; //margin-top:2px">          
              <asp:Label ID="title" runat="server"   ></asp:Label> 
             
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
            
            
            <label>
             
             <span style="width:100%;display:none;" id="fuerza"> 
                     <div style="width:130px">     Fuerza de Contraseña:   </div>
                     <div id="limitador">
                        <div id="passwordStrength" class="strength00"></div>
                     </div>
                     <div id="passwordDescription"></div>
 
              </span>                              
              <asp:Label ID="errorMess" runat="server" CssClass="errorMess" />
             
              
            </label>  
            
       
                 
              <table cellpadding="0" cellspacing="0" border="0" width="100%" align="center" style="background:url(img/FooterRHPROizq.png) no-repeat TOP LEFT" >
                 <tr>
                 <!--td align="left">
                   
                      <img src="img/LogoRHPRO1.png" border='0' align="absmiddle" width="130" style="/*margin-left:10px;*/"  >    
                  
                 </td-->
                 <td valign="middle" align="right" style="height:39px;" >  
                   
                           
                      <asp:LinkButton ID="btnConfirmar" runat="server" OnClick="btnConfirmar_Click" Text="Guardar"    CssClass="boton">Guardar</asp:LinkButton>
                      <asp:LinkButton runat="server" CausesValidation="False" ID="btnCancel" Text="Cancelar"               
                OnClientClick="javascript:window.close();" CssClass="boton" >Salir</asp:LinkButton>
             
                  </td>
                  
                 </tr>
              </table>
           
            
            
         </div>
<!-- ************************************************ -->

 
    
    </form>
 





</body>
</html>
