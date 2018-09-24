<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="SVG.aspx.cs" Inherits="RHPro.css.SVG" %>


<% 
    Response.ContentType = "text/css"; 
 %> 

<%
 
 String l_coloricono;
 String l_coloriconomenutop;
 String l_fuenteFecha;
 String l_fuentePiso;
  
 if (Session["EstiloR4_coloricono"]!="" )  
   l_coloricono = (String) Session["EstiloR4_coloricono"];
 else
   l_coloricono = "#FFFFFF";
 
 if (Session["EstiloR4_coloriconomenutop"]!="" )  
   l_coloriconomenutop = (String) Session["EstiloR4_coloriconomenutop"];
 else
   l_coloriconomenutop = "#FFFFFF";
 
  
 if ( Session["EstiloR4_fuenteFecha"]!="" )
   l_fuenteFecha = (String) Session["EstiloR4_fuenteFecha"];
 else
   l_fuenteFecha = "#ffffff";

 if (Session["EstiloR4_fuentePiso"] != "")
     l_fuentePiso = (String)Session["EstiloR4_fuentePiso"];
 else
     l_fuentePiso = "#ffffff";
 
  
%>

 
.SVG{    
 height:20px !important; width:20px !important; vertical-align:middle; text-align:center; padding:2px;  
 cursor:pointer !important;
 
}
.PATH {fill:<%=l_coloricono%> !important;      }
 
.SVG_LOGIN{    
 height:35px !important; width:29px !important; vertical-align:middle; text-align:center; padding:3px;  
 cursor:pointer !important;
 
}
.PATH_LOGIN {fill:<%=l_coloricono%> !important;      } 
 
.SVG_MENU{    
 height:22px !important; width:22px !important; vertical-align:middle; text-align:center; padding:2px; 
}
.PATH_MENU {fill:<%=l_coloricono%> !important;      }

  
.FondoSVG{ fill:#AAAAAA; fill-opacity:1; stroke-width:1px; stroke-linejoin:round;    }

.SVG_APERTURA { height:16px !important; width:16px !important; vertical-align:middle; text-align:center; padding-right:16px !important;  }
.PATH_APERTURA {fill:#999999; vertical-align:middle; text-align:center; fill-hover:#F00 }


.SVG_GADGET { 
  height:14px !important; width:14px !important; vertical-align:middle; text-align:center; padding:5px !important;
  /*stroke-opacity:0; */
 }
 
.PATH_GADGET {
  fill:#999999; vertical-align:middle; text-align:center;  cursor:pointer !important;
  /*stroke-opacity:0.5;*/
 }


.SVG_GADGET_INTERNO { height:19px !important; width:19px !important;    }
.PATH_GADGET_INTERNO {fill:<%=l_coloricono%>; vertical-align:middle; text-align:center;  cursor:pointer !important;}

.SVG_BARRA_TOP { height:28px !important; width:28px !important; vertical-align:middle; text-align:center; padding:0px !important;  }
.PATH_BARRA_TOP {fill:<%=l_coloriconomenutop%>; vertical-align:middle; text-align:center;  cursor:pointer !important;}

.SVG_MODULOS_CAB{height:37px !important; width:37px !important;}
.PATH_MODULOS_CAB {fill:<%=l_coloricono%>; vertical-align:middle; text-align:center;  cursor:pointer !important;}


.SVG_GADGET_ENG {height:60px !important; width:60px !important; vertical-align:middle; text-align:center; padding:9px !important; }
.PATH_GADGET_ENG {fill:#999999; vertical-align:middle; text-align:center;  cursor:pointer !important;}


.SVG_CONTROL_MODULOS{height:26px !important; width:26px !important;  }
.PATH_CONTROL_MODULOS{fill:<%=l_coloricono%>; vertical-align:middle; text-align:center;  cursor:pointer !important;}


.SVG_REDES{height:32px !important; width:32px !important;  }
.PATH_REDES{fill:<%=l_fuentePiso%>; vertical-align:middle; text-align:center;  cursor:pointer !important;}

.sm-blue{ background-color:#FF0000 !important;}
 