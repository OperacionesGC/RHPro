<%@ Control Language="C#" CodeFile ="Acceso_CRM.ascx.cs"  AutoEventWireup="true"  Inherits="Accesos.Acceso_CRM"     %>
 
 
<%
 
string Lenguaje =  ObjLenguaje.Idioma();// (String) System.Web.HttpContext.Current.Session["Lenguaje"];
string Texto = ""; 
switch (Lenguaje)
{
    case "enUS":
        Texto = "Microsoft Dynamics CRM is a management software tool for Customer Relationship.";		
        break;
    case "esAR":
        Texto = " Microsoft Dynamics CRM es un herramienta software para Gestión de las Relaciones con Clientes ";
        break;
    case "ptPT": 
	    Texto = "Microsoft Dynamics CRM é uma ferramenta de software de gestão para Relacionamento com o Cliente.";
		break;
    case "ptBR": 
	    Texto = "Microsoft Dynamics CRM é uma ferramenta de software de gestão para Relacionamento com o Cliente.";
		break;		
    default:
          Texto = "Microsoft Dynamics CRM es un herramienta software para Gestión de las Relaciones con Clientes. ";
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
     <img src = "Accesos/CRM/img/ACCESOS_CRM.png" border="0" style="margin-top:10px">
   </td>
 </tr>
 

 
</table>   </DIV>