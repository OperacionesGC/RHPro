<%@ Control Language="C#" CodeFile ="Acceso_ESS.ascx.cs"  AutoEventWireup="true"  Inherits="Accesos.Acceso_ESS"     %>
 
 
<%
 
string Lenguaje =  ObjLenguaje.Idioma();// (String) System.Web.HttpContext.Current.Session["Lenguaje"];
string Texto = ""; 
switch (Lenguaje)
{
    case "enUS":
        Texto = "Through it, accessible for quick and easy information to the company's human capital depends not only you but also to the staff with dependents.<BR> Self Management (ESS) will allow you to manage your data and participate in processes authorized for this purpose. <BR> <br />In case of staff assignment, you can access the Manager (MSS) where you can manage data and participate in the processes of staff in charge. <BR><br /> To enter the module, indicating on the dial to the left of the window the user and password that has been given time by the Human Resources area of the organization.";
        break;
    case "esAR":
        Texto = "A través del mismo, podrá acceder de manera ágil y sencilla a la información del   Capital humano de la empresa relacionada no sólo con usted sino también con el   personal que posee a cargo. <br /> <br />La Auto Gestión (ESS) permitirá que Ud.   pueda gestionar sus datos y participe en los procesos habilitados a tal fin. <br /><br /> En caso de tener personal a cargo asignado, podrá acceder al Manager   (MSS) donde podrá gestionar datos y participar en los procesos de su personal a   cargo.   <br />Para ingresar al módulo, indique en el cuadrante ubicado a la izquierda   de la ventana el usuario y clave que le ha sido entregado oportunamente por el   área de Recursos Humanos de la organización. ";
        break;
    case "ptPT": 
	    Texto = "Através dela, acessível a informação rápida e fácil ao capital humano da empresa depende não só você mas também para o pessoal com seus dependentes.<BR> Auto Gestão (ESS), vai permitir que você gerencie seus dados e participar de processos autorizados para esse fim.<BR><br />Em caso de cessão de pessoal, você pode acessar o Manager (MSS), onde você pode gerenciar os dados e participar nos processos de pessoal no cargo.<BR><br />Para entrar no módulo, indicando no mostrador para a esquerda da janela o usuário ea senha que foi dado tempo pela área de Recursos Humanos da organização.";
		break;
    case "ptBR": 
	    Texto = "Através dela, acessível a informação rápida e fácil ao capital humano da empresa depende não só você mas também para o pessoal com seus dependentes.<BR> Auto Gestão (ESS), vai permitir que você gerencie seus dados e participar de processos autorizados para esse fim.<BR><br />Em caso de cessão de pessoal, você pode acessar o Manager (MSS), onde você pode gerenciar os dados e participar nos processos de pessoal no cargo.<BR><br />Para entrar no módulo, indicando no mostrador para a esquerda da janela o usuário ea senha que foi dado tempo pela área de Recursos Humanos da organização.";
		break;		
    default:
          Texto = "A través del mismo, podrá acceder de manera ágil y sencilla a la información del   Capital humano de la empresa relacionada no sólo con usted sino también con el   personal que posee a cargo. <br /> <br />La Auto Gestión (ESS) permitirá que Ud.   pueda gestionar sus datos y participe en los procesos habilitados a tal fin. <br /><br /> En caso de tener personal a cargo asignado, podrá acceder al Manager   (MSS) donde podrá gestionar datos y participar en los procesos de su personal a   cargo.   <br />Para ingresar al módulo, indique en el cuadrante ubicado a la izquierda   de la ventana el usuario y clave que le ha sido entregado oportunamente por el   área de Recursos Humanos de la organización. ";
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
     <img src = "Accesos/ESS/img/ACCESOS_ESS.png" border="0" style="margin-top:10px">
   </td>
 </tr>
 

 
</table>   </DIV>