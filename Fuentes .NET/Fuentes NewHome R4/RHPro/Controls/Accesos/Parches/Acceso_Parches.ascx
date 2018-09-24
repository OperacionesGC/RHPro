<%@ Control Language="C#" CodeFile ="Acceso_Parches.ascx.cs"  AutoEventWireup="true"  Inherits="Accesos.Acceso_Parches"     %>
 
 
<%
 
string Lenguaje =  ObjLenguaje.Idioma();// (String) System.Web.HttpContext.Current.Session["Lenguaje"];
string Texto = ""; 
switch (Lenguaje)
{
    case "enUS":
        Texto = "Reference where the latest versions of Pro HR patches To download, you must log into the website..";
        break;
    case "esAR":
        Texto = "Referencia donde se encuentran las ultimas versiones de los parches de RH Pro. Para poder descargarlos, deberá loguearse en el sitio web. ";
        break;
    case "ptPT": 
	    Texto = "Referência, onde as versões mais recentes patches de Pro RH Para fazer o download, você deve entrar no site..";
		break;
    case "ptBR": 
	    Texto = "Referência, onde as versões mais recentes patches de Pro RH Para fazer o download, você deve entrar no site..";
		break;		
    default:
          Texto = "Referencia donde se encuentran las ultimas versiones de los parches de RH Pro. Para poder descargarlos, deberá loguearse en el sitio web.. ";
       break;
}


%> 
 
<style> 
#contenedor { margin-top:5px}
</style>
   <DIV class="InfoModulos">
<table width="95%"   border="0" cellspacing="0" cellpadding="0" align="center" id="contenedor"  >
 <tr>
   <td    valign="top" align="left"  >
    <%=Texto%>    
   </td>
 </tr>
  <tr>
   <td    valign="top" align="center"  >
     <img src = "Accesos/Parches/img/ACCESOS_Patch.png" border="0" style="margin-top:10px">
   </td>
 </tr>
 

 
</table>   </DIV>