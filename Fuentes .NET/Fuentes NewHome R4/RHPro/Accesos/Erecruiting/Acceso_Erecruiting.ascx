<%@ Control Language="C#" CodeFile ="Acceso_Erecruiting.ascx.cs"  AutoEventWireup="true"  Inherits="Accesos.Acceso_Erecruiting"     %>
 
 
<%
 
string Lenguaje =  ObjLenguaje.Idioma();// (String) System.Web.HttpContext.Current.Session["Lenguaje"];
string Texto = ""; 
switch (Lenguaje)
{
    case "enUS":
        Texto = "The E-recruiting system provides all the tools necessary to resume management via WEB, armed with CV models and building searches in order to incorporate the most qualified staff.<br>";
		Texto +="<p><b>Objectives:</b><br>";
		Texto +="• Incorporation of C.V. via WEB.";
		Texto +="• Administration of outstanding queries.";
		Texto +="• Having a proprietary basis Seekers";
		Texto +="• Provide the results of the searches.";
		Texto +="• Incorporation of the CV module of RHPro Jobs and Job Seekers in a transparent manner.</p>";
 	    Texto +="<p><b>Benefits;</b><br>";
		Texto +="• Full Control Base C.V.<br>"; 
		Texto +="• Integration Module for Jobs and Job Seekers RHPro.<br>";
		Texto +="• Upload CV accessing a Web site from the Home Page of the Company.<br>";
		Texto +="• Look and Field idem Company Home Page.<br></p>";	
        break;
		
		
    case "esAR":
        Texto = " El Sistema de E-recruiting brinda todas las herramientas necesarias para la administración de Currículum Vitae vía WEB, armado de modelos de CV y creación";
		Texto += " de Búsquedas con el objetivo de incorporar el personal mejor calificado.";
		Texto += " <p> <b>Objetivos</b> <br />";
		Texto += "   •	Incorporación de C.V. vía WEB. <br />";
	    Texto += "  •	Administración de Búsquedas destacadas. <br />";
		Texto += "   •	Tener una base propietaria de Postulantes <br />";
		Texto += "   •	Brindar el resultado de las Búsquedas. <br />";
		Texto += "   •	Incorporación de los CV al Modulo de Empleos y Postulantes de RHPro en forma trasparente. </p>";
		Texto += " <p> <b>Beneficios</b> <br />";
		Texto += "   •	Control total de la Base de C.V. <br />";
		Texto += "   •	Integración con Modulo de Empleos y Postulantes de RHPro. <br />";
		Texto += "   •	Carga de CV accediendo a un Sitio WEB desde la Home Page de la Empresa <br /> ";
		Texto += "   •	Look and Field idem Home Page de la Empresa </p>";
        break;
		
		
    case "ptPT": 
	    Texto = "The E-recruiting system provides all the tools necessary to resume management via WEB, armed with CV models and building searches in order to incorporate the most qualified staff.";	
		Texto += " <p> <b>Objetivos</b> <br />";
		Texto += "• Incorporação de C.V. via WEB.<br />";
		Texto += "• Administração de consultas pendentes.<br />";
		Texto += "• Ter uma base de propriedade Seekers<br />";
		Texto += "• Fornecer os resultados das pesquisas.<br />";
		Texto += "• Incorporação do módulo CV de Empregos RHPro e candidatos a emprego de uma forma transparente</p>";
 		Texto += " <p> <b>Benefícios</b> <br />";
		Texto += "• Controle total da Base C.V.<br /> ";
		Texto += "• Módulo de Integração de Empregos e RHPro Candidatos.<br /> ";
		Texto += "• Enviar CV acessar um site da Home Page da Empresa<br /> ";
		Texto += "• Olhe e Campo idem Página Inicial Empresa</p> ";		
		break;
		
    case "ptBR": 
	    Texto = "The E-recruiting system provides all the tools necessary to resume management via WEB, armed with CV models and building searches in order to incorporate the most qualified staff.";	
		Texto += " <p> <b>Objetivos</b> <br />";
		Texto += "• Incorporação de C.V. via WEB.<br />";
		Texto += "• Administração de consultas pendentes.<br />";
		Texto += "• Ter uma base de propriedade Seekers<br />";
		Texto += "• Fornecer os resultados das pesquisas.<br />";
		Texto += "• Incorporação do módulo CV de Empregos RHPro e candidatos a emprego de uma forma transparente</p>";
 		Texto += " <p> <b>Benefícios</b> <br />";
		Texto += "• Controle total da Base C.V.<br /> ";
		Texto += "• Módulo de Integração de Empregos e RHPro Candidatos.<br /> ";
		Texto += "• Enviar CV acessar um site da Home Page da Empresa<br /> ";
		Texto += "• Olhe e Campo idem Página Inicial Empresa</p> ";
		break;		
    default:
        Texto = " El Sistema de E-recruiting brinda todas las herramientas necesarias para la administración de Currículum Vitae vía WEB, armado de modelos de CV y creación";
		Texto += " de Búsquedas con el objetivo de incorporar el personal mejor calificado.</p>";
		Texto += " <p> <b>Objetivos</b> <br />";
		Texto += "   •	Incorporación de C.V. vía WEB. <br />";
	    Texto += "  •	Administración de Búsquedas destacadas. <br />";
		Texto += "   •	Tener una base propietaria de Postulantes <br />";
		Texto += "   •	Brindar el resultado de las Búsquedas. <br />";
		Texto += "   •	Incorporación de los CV al Modulo de Empleos y Postulantes de RHPro en forma trasparente </p>";
		Texto += " <p> <b>Beneficios</b> <br />";
		Texto += "   •	Control total de la Base de C.V. <br />";
		Texto += "   •	Integración con Modulo de Empleos y Postulantes de RHPro. <br />";
		Texto += "   •	Carga de CV accediendo a un Sitio WEB desde la Home Page de la Empresa <br /> ";
		Texto += "   •	Look and Field idem Home Page de la Empresa </p>";
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
     <p><%=Texto%>    
       
 </tr>
  <tr>
   <td    valign="top" align="center"  >
     <img src = "Accesos/Erecruiting/img/ACCESOS_Erecruiting.png" border="0" style="margin-top:10px">
   </td>
 </tr>
 

 
</table>   </DIV>