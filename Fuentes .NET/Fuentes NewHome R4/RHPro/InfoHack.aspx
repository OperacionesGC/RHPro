<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="InfoHack.aspx.cs" Inherits="RHPro.InfoHack" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title></title>
<style>
.ErrHack { text-align:center; font-family:Calibri; font-size:20pt; vertical-align:middle !important; display:inline-block; 
           padding:18px; border:1px solid #fff;   color:#fff; width:40%;
           position:absolute; top:20%; left:28%}
body { background-color:#333}            
</style>    
    
</head>
<body>
    <form id="form1" runat="server">
     
    
    <div class="ErrHack">
       <img src="img/logoRHPRO_X2Blanco.png">
       <div>
       <%  Response.Write(Obj_Lenguaje.Label_Home("Credenciales invalidas")); %>
       </div>
       <div>
       <%  Response.Write(Obj_Lenguaje.Label_Home("Cierre el navegador y vuelva a ingresar")); %>
       </div>
    </div>
    
 
    
    </form>
</body>
</html>
