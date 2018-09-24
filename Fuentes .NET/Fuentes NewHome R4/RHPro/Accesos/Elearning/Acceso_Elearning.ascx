<%@ Control Language="C#" CodeFile ="Acceso_Elearning.ascx.cs"  AutoEventWireup="true"  Inherits="Accesos.Acceso_Elearning"     %>
 
 
<%
 
string Lenguaje =  ObjLenguaje.Idioma();// (String) System.Web.HttpContext.Current.Session["Lenguaje"];
string Texto = ""; 
switch (Lenguaje)
{
    case "enUS":
        Texto = "Our platform allows you to train comfortably elearning from a PC with an internet connection, when you want, wherever you want.<BR><BR>";
        Texto +=" Interactive and dynamic courses, using various resources to capture students' attention and evaluation of knowledge, with the same benefits you get with";
		Texto +="classroom training courses. <BR><BR>";
        Texto +="A solution for everyone to learn and take advantage";
        Texto +="the most of the tools provided to HR Pro,";
        Texto +="and improve the daily performance immediately.";
		
		
        break;
    case "esAR": 
		Texto = " Nuestra plataforma de elearning le permite capacitarse cómodamente desde una PC con una conexión a internet, cuándo quiera, dónde quiera. ";
        Texto += " Cursos interactivos y dinámicos, con la utilización diversos recursos para capturar la atención del alumno y evaluación de conocimientos, con los mismos ";
        Texto += "beneficios que obtiene con los cursos de capacitación presencial.<BR>";
        Texto += "Una solución al alcance de todos para conocer y poder aprovechar ";
        Texto += "al máximo las herramientas provistas para RH Pro, ";
        Texto +=" y mejorar el desempeño diario de forma inmediata.";	
		
        break;
    case "ptPT": 
	    Texto = "A nossa plataforma permite que você treine confortavelmente e-learning a partir de um PC com ligação à Internet, quando quiser, onde quiser.<BR><BR>";
        Texto += "Cursos interativos e dinâmicos, utilizando vários recursos para capturar a atenção dos alunos e avaliação de conhecimentos, com os mesmos benefícios que você começa com cursos de formação em sala de aula.<BR><BR>";
        Texto += "Uma solução para que todos possam aprender e tirar proveito";
        Texto += "ao máximo as ferramentas fornecidas para RH Pro,";
        Texto += "e melhorar o desempenho diário imediatamente.";	     
		break;
		
    case "ptBR": 
	    Texto = "A nossa plataforma permite que você treine confortavelmente e-learning a partir de um PC com ligação à Internet, quando quiser, onde quiser.<BR><BR>";
        Texto += "Cursos interativos e dinâmicos, utilizando vários recursos para capturar a atenção dos alunos e avaliação de conhecimentos, com os mesmos benefícios que você começa com cursos de formação em sala de aula.<BR><BR>";
        Texto += "Uma solução para que todos possam aprender e tirar proveito";
        Texto += "ao máximo as ferramentas fornecidas para RH Pro,";
        Texto += "e melhorar o desempenho diário imediatamente.";
		break;		
		
    default:
         		Texto = " Nuestra plataforma de elearning le permite capacitarse cómodamente desde una PC con una conexión a internet, cuándo quiera, dónde quiera. ";
        Texto += " Cursos interactivos y dinámicos, con la utilización diversos recursos para capturar la atención del alumno y evaluación de conocimientos, con los mismos ";
        Texto += "beneficios que obtiene con los cursos de capacitación presencial.<BR>";
        Texto += "Una solución al alcance de todos para conocer y poder aprovechar ";
        Texto += "al máximo las herramientas provistas para RH Pro, ";
        Texto +=" y mejorar el desempeño diario de forma inmediata.";
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
     <img src = "Accesos/Elearning/img/ACCESOS_Elearning.png" border="0" style="margin-top:10px">
   </td>
 </tr>
 

 
</table>   </DIV>