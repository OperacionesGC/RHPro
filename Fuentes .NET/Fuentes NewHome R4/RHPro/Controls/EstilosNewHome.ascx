<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="EstilosNewHome.ascx.cs" Inherits="RHPro.Controls.EstilosNewHome" %>

<% 
    AnchoPagina = "100%";
    RadioGadget = "0px";
    //FuenteCabeceraGadget_Color = "#4d515c";
    FuenteCabeceraGadget_Font = "Arial";
    FuenteCabeceraGadget_Size = "9pt";
    AnchoMenuLinks = "270px";
    //BackgroundCabeceraGadget = "#f6f8fb";
%>  

 <style>
  
*{
 
	scrollbar-3dlight-color:#CDCDCD;
	scrollbar-arrow-color:#999999;
	scrollbar-base-color:#EFEFEF;
	scrollbar-darkshadow-color:#CDCDCD;
	scrollbar-face-color:#cccccc;
	scrollbar-highlight-color:#efefef;
	scrollbar-shadow-color:#EFEFEF; 
}
   
     
body
{ 
    margin: 0 !important;
    padding:0 !important; 
    width:100%!important;
    height:100%!important;
    background: #cccccc; /* Old browsers */
    vertical-align:top !important;          
}



#PopUpChangePassword {  width:480px; color:White; margin-left:30px; background-color:#ee2e24; padding-left:70px; }
#PopUpChangePassword .title{ font-size:20px;  margin-bottom:20px;  margin-left:50px; }
#PopUpChangePassword .text {float:left;  }
#PopUpChangePassword .text span{font-size:15px; margin-top:5px; color:White; }
#PopUpChangePassword .fields {float:right;  }
#PopUpChangePassword .fields input{margin-top:5px;  }

#PopUpChangePassword .table dt{ width:240px; padding-bottom:10px;}
#PopUpChangePassword .table dd{ margin-left:10px; margin-bottom:10px; width:100px;}
#PopUpChangePassword .table dd input{  width:100px;}
#PopUpChangePassword a { text-decoration:none;}
#PopUpChangePassword .linkButton span{ color: white; font-family: Arial; font-size: 15px; font-weight: bold;text-decoration:none; }

 #SearchPopUp{ font-family:Arial; margin: 0px; width:400px; margin-top:0px;  margin-left:20px; background-color:#aaa; background-image:none; }
 #SearchPopUp #tblSearch{ background-color:White; }
#SearchPopUp #tblSearch .titulo td {color:Red; font-size:18px; text-align:center; }
#SearchPopUp #tblSearch td{background-color:#ccc; padding:5px;}
#SearchPopUp #tblSearch .izquierda {color:Red;}
#SearchPopUp #tblSearch .derecha a{ text-decoration:none; color:White;}
#SearchPopUp #logoSearch { background-image: url(Images/Logo.gif); height:91px; width:184px; margin-left:00px; margin-bottom:15px; }
#SearchPopUp #lbtitulo { text-decoration:underline; font-size:12px; }
#SearchPopUp #lbPagina{  font-size:12px; }


#SearchPopUp #modulo{ font-weight:bold; color:Black; padding-bottom:5px;}
#SearchPopUp #descripcion{ padding-left:15px;}
#SearchPopUp #link a{ color:Red; font-weight:bold;}
#SearchPopUp #text{ padding-left:15px; padding-bottom:5px;}


.barraSuperior{ background-color: #4a4a4a;height: 18px; vertical-align: middle; }
.barraSuperior div{ color: #999999; font-family: Tahoma; font-size: 11px; text-align: right; vertical-align: middle; padding-top: 2px; margin-right: Auto; width:950px; margin-right:auto; margin-left:auto; }
.barraSuperior div img{border: 0px; padding-left: 0px; }
  

        
.form{ width: 100%; text-align: left;  margin:0 !important; padding:0 !important; vertical-align:top !important }
    
#logo{ background-image: url('Images/TopDemo.png');
background-repeat: no-repeat;height: 90px;margin-top: 15px; width:950px; margin-bottom:15px; }
        
#botonera #superior{ background-repeat: no-repeat; background-image: url('Images/Botonera/newBotoneraSup.png');
height: 6px; line-height: 1px; font-size: 1px; }
#botonera #medio{ background-repeat: repeat-y; background-image: url('Images/Botonera/newBotoneraMedio.png'); 
height:15px; padding-top:2px; }
#botonera #inferior{ background-repeat: no-repeat; background-image: url('Images/Botonera/newBotoneraInf.png');
height: 6px;  }

.menu_list{ list-style-type: none;font-family: arial;font-size: 12px;margin: 0;padding: 0; }

.menu_list li {float: left;padding: 0 10px 0px 10px;color: #ee2e24;border-right: solid 1px #999999;color: red;font-family: Arial;font-size: 11px;font-weight: bold; text-decoration:none; }
.menu_list li a{ float: left;color: #ee2e24;color: red;font-family: Arial;font-size: 11px;font-weight: bold; text-decoration:none; }

.clear{ clear: both; }
.inputSearch{ background-color:white; border:solid 1px #000 !important; color:red;}
.searchButton{ color: white !important; border: 0px; font-size: 14px !important; margin-top: -1px; border:none;}

#contenido {margin-top: 15px; width: 950px;}


#Izquierda{width: 150px; float: left;}
.mruImage{ background-image: url(Images/Banner/MruImage.gif); height:380px;}
#Menu1Izq #superior { background-repeat: no-repeat; background-image: url('Images/MRU/IzqSupGris1.gif'); 
height: 9px; line-height: 1px; font-size: 1px; }
#Menu1Izq #titulo   { background-repeat: no-repeat; background-color: #6f6f6f; height: 50px; }                  
#Menu1Izq #titulo span { color: white; font-family: Arial; font-size: 18px;  margin: 0 0 0 0; padding-left: 20px }
#Menu1Izq #cuerpo   { background-repeat: no-repeat; background-color: #dedede; overflow:auto; }
#Menu1Izq #cuerpo div{  width: 125px; margin-left: 8px; padding-top: 7px; padding-bottom: 5px; color: Red; font-family: Arial, Helvetica, sans-serif; font-size: 12px; font-weight: bold}
#Menu1Izq #cuerpo  .separador{border-bottom: dotted 1px black; padding-top: 0px;}
#Menu1Izq #cuerpo div .descripcion{color: Black; font-size: 11px; font-weight: normal; }
.MruCuerpo{ width:150px;height:312px; }
#Menu1Izq #inferior {background-repeat: no-repeat; background-image: url('Images/MRU/IzqInfGris1.gif');
height: 9px; }
#Menu2Izq {margin-top: 15px;}
#Menu2Izq a img{height:240px; width:150px;}

#Centro{ float: left; margin-left: 5px; }
#centralArriba{ margin: 0px 10px 0px 10px; width: 600px; }
#centralArriba #superior{ background-repeat: no-repeat; background-image: url('Images/Module/CentralSup.gif');
height: 8px; line-height: 1px; font-size: 1px; }
#centralArriba #titulo{ background-repeat: no-repeat; background-color: #4b4a4a; margin-top: -1px;color: White; padding-left: 10px; height: 35px; font-size:20px; }
#centralArriba #medio{ background-repeat: no-repeat; background-color: white; margin-top: -1px;padding-left: 5px;  }
#centralArriba #medio .CentroArribaIzq{ 
background-repeat: no-repeat;
background-color: #cccccc;
margin: 5px;
color: Red;
padding: 5px 30px 5px 5px;
font-size: 12px; 
font-family:arial;  
}
#centralArriba #medio .CentroArribaIzq a { text-decoration:none; color:Red; font-weight:bold; font-family:arial; font-size:9pt;
   }
#centralArriba #medio .CentroArribaDerTitulo{
  background-repeat: no-repeat;
  background-color: #d9d9d9;
  margin: 5px;
  border-bottom: solid 1px white; 
  color: red; 
  padding: 5px;   
}
#centralArriba #medio .CentroArribaDerTitulo a {text-decoration:none; color:Red; }
#centralArriba #medio .CentroArribaDerDescripcion{ margin-left: 5px; margin-right:5px;height: 265px; overflow: auto; font-size:11px;}
#centralArriba #medio .CentroArribaDerInfo{ background-repeat: no-repeat; background-color: #999999; margin:0px 5px 5px 5px; border-bottom: solid 1px white; color: white; text-align: right; font-size: 13px;padding: 5px; height: 15px; }
#centralArriba #medio .CentroArribaDerInfo a { text-decoration:none; color:White;}
#centralArriba #inferior{ background-repeat: no-repeat; background-image: url('Images/Module/CentralInf.gif');
height: 8px;}

.divScroll { float: left; width: 270px; height: 325px; overflow: auto; margin-top:1px;  }

.listaLink{border:none;  width:270px; }
.listaLink option{ margin-bottom:10px; padding-top:5px; padding-bottom:5px; color:Red;font-weight: bold; background-color:#cccccc; }

#Derecha { width: 170px; float: right; margin-left:5px; }
#login #superior{ background-repeat: no-repeat; background-image: url('Images/Login/newLoginSup.png');
height: 9px; line-height: 1px; font-size: 1px; }
#login #medio{ background-repeat: no-repeat; background-color: #ee2e24; margin-top: -1px; padding-left:2px; height:150px; }

.btnLoginOFF { margin:30px 0px 0px 10px; font-size:10px; }


.ErrorMessegeON {margin-left:10px;   margin-top:30px; display:block; font-size:12px !important;}
.ErrorMessegeOFF {display:none; }
.btnLoginON {margin-top:0px; }
.linksLogin { margin-top:25px;}
.linksLogin a { text-decoration:none; color:White; font-size:10px;  }

.errorMessage {color:Black; font-size:15px;}
.btnPolitic { padding-left: 0px; margin-left:10px;}
#login #medio span a , #login #medio span a:selected { color: white; font-family: Arial; font-size: 18px; padding-left: 0px;text-decoration: none ;}
#login #medio label{ color: white; font-family: Arial; font-size: 18px; text-decoration:none; margin-left:10px; }
 #Derecha #login #medio span{ color: white; font-family: Arial; font-size: 18px; text-decoration:none;  }
 #linkLogOff { padding-top:20px; }
 #linkLogOff span a { text-decoration:none; font-size:10px; color:White;}
 
#login #inferior{ background-repeat: no-repeat; background-image: url('Images/Login/newLoginInf.png');height: 9px; }
#Menu1Der{ margin-top: 15px; }
#Menu1Der #superior { background-repeat: no-repeat; background-image: url('Images/Message/DerSupGris1.gif');
height: 9px; line-height: 1px; font-size: 1px; }
#Menu1Der #medioTitulo{ background-repeat: no-repeat; background-color: #6f6f6f; height: 50px; }
#Menu1Der #medioTitulo span{ color: white;font-family: Arial;font-size: 18px;margin: 0 0 0 0;padding-left: 15px;  }
#Menu1Der #medioCuerpo{ background-repeat: no-repeat; background-color: #dedede;height:131px; overflow:auto; }

#Menu1Der #medioCuerpo div{ width: 125px; margin-left: 15px; padding-top: 7px; padding-bottom: 5px; color: Red; font-family: Arial, Helvetica, sans-serif; font-size: 12px; font-weight: bold}
#Menu1Der #medioCuerpo  .separador{border-bottom: dotted 1px black; padding-top: 0px;}
#Menu1Der #medioCuerpo div .descripcion{ color: Black; font-size: 10px; }
#Menu1Der #inferior{ background-repeat: no-repeat; background-image: url('Images/Message/DerInfGris1.gif');
height: 9px; }


 #Menu2Der{ margin-top: 15px; }
#Menu2Der #superior{ background-repeat: no-repeat; background-image: url('Images/Link/DerSupGris1.gif');
height: 9px; line-height: 1px; font-size: 1px;
}
#Menu2Der #medioTitulo{ background-repeat: no-repeat; background-color: #6f6f6f; height: 30px; }
#Menu2Der #medioTitulo span{ color: white;font-family: Arial;font-size: 18px;margin: 0 0 0 0;padding-left: 15px;  }
#Menu2Der #medioCuerpo{ background-repeat: no-repeat; background-color: #eeeeee; height:193px; overflow:auto; }
#Menu2Der #medioCuerpo div{ width: 125px; margin-left: 15px; padding-bottom: 5px; color: Red; font-family: Arial; font-size: 12px; font-weight: bold }
#Menu2Der #medioCuerpo div .descripcion{ color: Black; font-size: 10px; text-decoration:none; }
#Menu2Der #inferior{ background-repeat: no-repeat; background-image: url('Images/Link/DerInfGris2.gif');
height: 9px; }

.ComboData { margin-left:5px; }
.iframeCentral {padding: 0px; margin-left:10px; margin-top:14px; width:600px; height:241px;}
#imageFrame { background-repeat: no-repeat; background-image: url('Images/ImageCentral.gif'); height:241px; margin-top:15px; margin-left:10px; width:600px;}

.imgBanner{ height:240px; width:150px;}
#footer{}

#footer #superior{background-repeat: no-repeat; background-image: url('Images/Footer/FootSup.gif');height: 8px; line-height: 1px; font-size: 1px; margin-top:15px;}
#footer #medio{ background-color:#ee2e24; color:White; font-family:Tahoma;height: 15px;}
#footer #medio .titulo{  padding:0px 0px 0px 10px; font-size:14px; }
#footer #medio .version{ font-size:9px;  padding-top:3px; padding-right:10px;}
#footer #inferior{background-repeat: no-repeat; background-image: url('Images/Footer/FootInf.gif');height: 9px; line-height: 1px; font-size: 1px; margin-bottom:10px;}

.inputs{ border: 1px solid #fff;background: white;height: 15px;width: 50px;border-color: #000;  }

.InputDeshab { background-color:#efefef !important; color:#999 !important}

.Combo input{ border-style: none;
    border-color: inherit;
    border-width: 0px;
    margin: 0;
background-image: url('Images/botonCombo.png');
background-position: left top;
    padding: 2px 0px 0px 8px;    font-size: 11px;    height: 17px;    width: 150px; }
.Combo button{ border: solid 1px white;height: 19px;width: 19px; }

.borde { border-style:solid; border-width:1px; border-color:black; background: white; color:Red;}

.Boton_Gadgets TD{
 /*filter:progid:DXImageTransform.Microsoft.Gradient(endColorstr='#cccccc', startColorstr='#cccccc', gradientType='0') !important;            
 SAFARI y CHROME:  background: -webkit-gradient(linear, left top, left bottom, from(#333333), to(#d3d2d2)); 
 MOZILLA:  background: -moz-linear-gradient(top, #333333,#333333);	 
*/
background-color:#CCCCCCC !important;
color:#333333 !important;
}


.TopeContenedorG{ filter:progid:DXImageTransform.Microsoft.Gradient(endColorstr='#222222', startColorstr='#222222', gradientType='0') !important;            
 SAFARI y CHROME:  background: -webkit-gradient(linear, left top, left bottom, from(#A8A8A8), to(#d3d2d2)); 
 MOZILLA:  background: -moz-linear-gradient(top, #A8A8A8,#d3d2d2);	 
color:#FFFFFF !important;}

.BotonMenPpal {height:30px; width:34px; vertical-align: bottom;
filter:progid:DXImageTransform.Microsoft.Gradient(endColorstr='#222222', startColorstr='#222222', gradientType='0') !important;            
}

.Separador { color:#FFF; font-family:Tahoma; font-size:11pt;
filter:progid:DXImageTransform.Microsoft.Gradient(endColorstr='#222222', startColorstr='#222222', gradientType='0') !important;            
   cursor:pointer }
 
.Menu_Ppal TR  TD{ color:#FFF; font-family:Tahoma; font-size:11pt; 
 /*filter:progid:DXImageTransform.Microsoft.Gradient(endColorstr='#222222', startColorstr='#222222', gradientType='0') !important;            */
 background-color: #fff;/*JPB*/
 height:15px; cursor:pointer;  vertical-align:middle !important; }

 
.ListaModInactivos TD DIV
{
	 background-color:<%=fondoModulos%> !important;  
}

 
.TablaInactivos TR:hover TD 
{
	 background-color:#ccc !important;  
}

.OcultaInactivos span{margin-left:10px; vertical-align:middle; margin-right:2px;}
.ListaModInactivos TR:hover TD {  background-color:#ccc !important;  }

.SeparadorIconoMenu {padding-left:31px;} 



.Btn_FuncionGadget{cursor:pointer;  padding:2px; margin:1px; display:inline-block}  
.IconoModulo, .IconoInfoModulo  ,.IconoLogin
{
	 width:24px !important; height:24px  !important; 
	 vertical-align:middle; text-align:center; 	 
	 cursor:pointer !important;  
	 text-align:center;
	 border:0;
     padding-top:4px;
     padding-bottom:4px;
     padding-left:0px; 
     margin-left:5px; margin-right:2px;     
}

.IconoModulo_PNG , .IconoInfoModulo_PNG  ,.IconoLogin_PNG
{
	 width:23px !important; 
	 vertical-align:middle; text-align:center; 	 
	 cursor:pointer !important;  
	 text-align:center;
	 border:0;
     padding-top:5px;
     padding-bottom:5px;
     padding-left:0px; margin-left:2px; margin-right:4px;     
}

.IconoInfoModulo { margin-left:10px; margin-right:5px}

.IconoLogin{ width:35px !important; height:41px  !important; }
.IconoLogin_PNG{  height:28px  !important; }

.IconosMaximizaModulos
{
	width:24px !important; 
	height:22px  !important; vertical-align:middle; 
	text-align:left; 
/*	padding:5px; */
    padding:5px; 
	padding-left:0px; margin-left:0px; margin-right:0px;  
	cursor:pointer;
}
 
 
.IconosMaximizaModulos_PNG
{
    height:29px  !important; vertical-align:middle; 
	text-align:left;  
    padding:5px; 
	padding-left:0px; margin-left:0px; margin-right:0px;  
	cursor:pointer;
}
.IconoModulo_Favoritos {  width:22px; height:22px; vertical-align:middle; text-align:center;  cursor:pointer !important; margin-left:-4px} 


.IconoAperturaModulo {  width:19px; height:19px; vertical-align:middle; text-align:center;   }
.IconoAperturaModulo_PNG {height:19px; vertical-align:middle; text-align:center;   }
.IconoModuloGadget {  width:24px; height:24px; vertical-align:middle; text-align:center;  cursor:pointer !important;
}
.IconoModuloAcceso { width:38px !important; height:19px !important; }
.IconoModuloGadget_PNG {  height:22px; vertical-align:middle; text-align:center;  cursor:pointer !important; margin-right:4px;}
                 
.IconoCabModulo { width:32px; height:32px; vertical-align:middle; text-align:center;  cursor:pointer !important; }

.IconoCabModulo_PNG { height:32px; vertical-align:middle; text-align:center;  cursor:pointer !important; }

.IconoBarraTopUser
{
	  height:45px; vertical-align:middle; text-align:center; 
	  border:2px solid <%=coloriconomenutop%>;  
	 -webkit-border-radius: <%=RadioGadget%>;
     -moz-border-radius: <%=RadioGadget%>;
     border-radius: <%=RadioGadget%>;
}

.IconoBarraTopUser_PNG
{
	  height:45px; vertical-align:middle; text-align:center; 
	  border:2px solid <%=coloriconomenutop%>;  
	 -webkit-border-radius: <%=RadioGadget%>;
     -moz-border-radius: <%=RadioGadget%>;
     border-radius: <%=RadioGadget%>;
}

.IconosBarraAyuda{width:32px; height:24px; vertical-align:middle; text-align:center;
                  margin-left:7px; margin-top:4px;  margin-bottom:5px;   }
.IconosBarraTop{width:32px; height:28px; vertical-align:middle; text-align:center;   }
.IconosBarraTop_PNG{ height:33px; vertical-align:middle; text-align:center;   }

.IconoAcceso{width:19px; vertical-align:middle; text-align:center; padding:0px; margin:0; margin-left:0px; margin-right:6px;  }
.IconoAcceso_PNG{width:19px; vertical-align:middle; text-align:center; padding:0px; margin:0; margin-left:0px; margin-right:6px;  }

.IconoMRU {  width:19px; height:19px; vertical-align:middle; text-align:center;  cursor:pointer !important;   }
.IconoMRU_PNG {  height:22px; vertical-align:middle; text-align:center;  cursor:pointer !important; margin-right:2px;   }

.IconoREDES { width:32px; height:32px; vertical-align:middle; text-align:center;  cursor:pointer !important; padding:5px; margin:2px;   }
.IconoREDES_PNG{ height:42px; vertical-align:middle; text-align:center;  cursor:pointer !important; padding:5px; margin:2px;   }

.IconoMenuPPAL {  width:26px; height:26px; vertical-align:middle; text-align:center;  cursor:pointer !important;     }
.IconoMenuPPAL_PNG {  height:26px; vertical-align:middle; text-align:center;  cursor:pointer !important;     }
 
.Item_Links { height:12px !important; }
.Item_Links:hover TD {  
    filter:progid:DXImageTransform.Microsoft.Gradient(endColorstr='#111111', startColorstr='#111111', gradientType='0') !important;  
 }
 
  
 
 img:hover .DetRedes
 {
 	border:1px solid #efefef;
 	 -webkit-border-radius: <%=RadioGadget%>;
     -moz-border-radius: <%=RadioGadget%>;
     border-radius: <%=RadioGadget%>;
 }
 
 .FONDO_SVG{ vertical-align:middle; text-align:center; }
 
 embed,object { background-color:none !important}
 
.SeparadorMenu{ height:10px !important; padding:0px !important; font-size:8pt !important; color:#666 !important;    }

  
   

/*AGREGADOS*/

.Encabezado
{
    margin-top:0px !important;  margin-bottom:0px  !important; 
    height:45px;
    background-color:<%=fondoCabecera%>; 
    border:0px; 
    width:<%=AnchoPagina%>;
    padding:0px !important;   
    /*
    background: rgb(107,107,107);
    background: url(data:image/svg+xml;base64,PD94bWwgdmVyc2lvbj0iMS4wIiA/Pgo8c3ZnIHhtbG5zPSJodHRwOi8vd3d3LnczLm9yZy8yMDAwL3N2ZyIgd2lkdGg9IjEwMCUiIGhlaWdodD0iMTAwJSIgdmlld0JveD0iMCAwIDEgMSIgcHJlc2VydmVBc3BlY3RSYXRpbz0ibm9uZSI+CiAgPGxpbmVhckdyYWRpZW50IGlkPSJncmFkLXVjZ2ctZ2VuZXJhdGVkIiBncmFkaWVudFVuaXRzPSJ1c2VyU3BhY2VPblVzZSIgeDE9IjAlIiB5MT0iMCUiIHgyPSIxMDAlIiB5Mj0iMCUiPgogICAgPHN0b3Agb2Zmc2V0PSIwJSIgc3RvcC1jb2xvcj0iIzZiNmI2YiIgc3RvcC1vcGFjaXR5PSIxIi8+CiAgICA8c3RvcCBvZmZzZXQ9IjEwMCUiIHN0b3AtY29sb3I9IiNhYWFhYWEiIHN0b3Atb3BhY2l0eT0iMSIvPgogIDwvbGluZWFyR3JhZGllbnQ+CiAgPHJlY3QgeD0iMCIgeT0iMCIgd2lkdGg9IjEiIGhlaWdodD0iMSIgZmlsbD0idXJsKCNncmFkLXVjZ2ctZ2VuZXJhdGVkKSIgLz4KPC9zdmc+);
    background: -moz-linear-gradient(left,  rgba(107,107,107,1) 0%, rgba(170,170,170,1) 100%);
    background: -webkit-gradient(linear, left top, right top, color-stop(0%,rgba(107,107,107,1)), color-stop(100%,rgba(170,170,170,1)));
    background: -webkit-linear-gradient(left,  rgba(107,107,107,1) 0%,rgba(170,170,170,1) 100%);
    background: -o-linear-gradient(left,  rgba(107,107,107,1) 0%,rgba(170,170,170,1) 100%);
    background: -ms-linear-gradient(left,  rgba(107,107,107,1) 0%,rgba(170,170,170,1) 100%);
    background: linear-gradient(to right,  rgba(107,107,107,1) 0%,rgba(170,170,170,1) 100%);
    filter: progid:DXImageTransform.Microsoft.gradient( startColorstr='#6b6b6b', endColorstr='#aaaaaa',GradientType=1 );
*/
}

.LineaEncabezado
{
	/*
	background:url(img/FONDO_MENU_BAR.png) repeat-x bottom right; height:1px !important;  
 	filter: alpha(opacity=50);
	-moz-opacity:0.5;
	-khtml-opacity: 0.5;
	opacity: 0.5;
	*/
}

.TD_LOGO{ text-align:left !important; vertical-align:middle !important; width:1px;}
.TD_LOGO a{ text-decoration:none; border:0; }
.TD_LOGO img{ text-decoration:none; border:0; }
 
.TD_BARRA_TOP{  text-align:right !important; vertical-align:bottom !important;  width:100%;}


.TABLA_MENU_TOP{ text-align:right; width:1px; height:100%;}
.TABLA_MENU_TOP TR TD{ height:45px; vertical-align: middle; text-align:center; white-space:nowrap !important;  
                        color:<%=coloriconomenutop%>;     }
.TABLA_MENU_TOP TR TD:hover 
{
	/*background-color:#78BFE4;*/
	/*background:url(img/FONDO_MENU_BAR.png) repeat-x bottom right;*/
	color:<%=coloriconomenutop%>    !important; }

.MENU_TOP_NAV
{
	/*cursor:pointer;*/
	font-family:Tahoma; font-size:9pt  !important; 
	height:45px !important;
    vertical-align:bottom; 
	border:0px solid #ffffff;
	/*padding:6px;*/
	margin:0px;
	/*border-left:1px solid <%=coloriconomenutop%>;*/
	color:<%=coloriconomenutop%>; 
	width:100%;
	white-space:nowrap;
     
 }
.MENU_TOP_NAV div  
{
	/*margin-right:8px; margin-left:4px;*/
 }

 
SPAN.MENU_TOP_NAV:hover ,.MENU_TOP_NAV:hover
{
	/*background:url(img/FONDO_MENU_BAR.png) repeat-x bottom right; */
	/*background-color:#5477FF;*/
	/*color:#fff;*/
    
/*
    background: rgb(198,198,198)  !important;  
    background: url(data:image/svg+xml;base64,PD94bWwgdmVyc2lvbj0iMS4wIiA/Pgo8c3ZnIHhtbG5zPSJodHRwOi8vd3d3LnczLm9yZy8yMDAwL3N2ZyIgd2lkdGg9IjEwMCUiIGhlaWdodD0iMTAwJSIgdmlld0JveD0iMCAwIDEgMSIgcHJlc2VydmVBc3BlY3RSYXRpbz0ibm9uZSI+CiAgPGxpbmVhckdyYWRpZW50IGlkPSJncmFkLXVjZ2ctZ2VuZXJhdGVkIiBncmFkaWVudFVuaXRzPSJ1c2VyU3BhY2VPblVzZSIgeDE9IjAlIiB5MT0iMCUiIHgyPSIwJSIgeTI9IjEwMCUiPgogICAgPHN0b3Agb2Zmc2V0PSIwJSIgc3RvcC1jb2xvcj0iI2M2YzZjNiIgc3RvcC1vcGFjaXR5PSIxIi8+CiAgICA8c3RvcCBvZmZzZXQ9IjEwMCUiIHN0b3AtY29sb3I9IiNmMmYyZjIiIHN0b3Atb3BhY2l0eT0iMSIvPgogIDwvbGluZWFyR3JhZGllbnQ+CiAgPHJlY3QgeD0iMCIgeT0iMCIgd2lkdGg9IjEiIGhlaWdodD0iMSIgZmlsbD0idXJsKCNncmFkLXVjZ2ctZ2VuZXJhdGVkKSIgLz4KPC9zdmc+)  !important;
    background: -moz-linear-gradient(top,  rgba(198,198,198,1) 0%, rgba(242,242,242,1) 100%)  !important; 
    background: -webkit-gradient(linear, left top, left bottom, color-stop(0%,rgba(198,198,198,1)), color-stop(100%,rgba(242,242,242,1)))  !important; 
    background: -webkit-linear-gradient(top,  rgba(198,198,198,1) 0%,rgba(242,242,242,1) 100%) !important ; 
    background: -o-linear-gradient(top,  rgba(198,198,198,1) 0%,rgba(242,242,242,1) 100%)  !important; 
    background: -ms-linear-gradient(top,  rgba(198,198,198,1) 0%,rgba(242,242,242,1) 100%)  !important; 
    background: linear-gradient(to bottom,  rgba(198,198,198,1) 0%,rgba(242,242,242,1) 100%)  !important;  
    filter: progid:DXImageTransform.Microsoft.gradient( startColorstr='#c6c6c6', endColorstr='#f2f2f2',GradientType=0 ) !important;  
 */
} 
 
.TD_GADGET{  /*background-color:#333333;*/   font-family:Arial; font-size:11pt; font-weight:normal; }
.idiomas, .idiomas a{  /*background-color:Green;  */  font-family:Arial; font-size:11pt; font-weight:normal;}

.Gadgets_Del_Modulo{ vertical-align:top !important; text-align:left !important;}

.user{   font-family:Arial; font-size:11pt; font-weight:normal; margin-left:2px;   }
.user:hover{background:url(img/FONDO_MENU_BAR.png) repeat-x bottom right; }

.Info_Log_Usr{height:10px; line-height:2}
.Info_Log_Usr TR  TD{ height:10px !important; white-space:nowrap; padding:2px; text-align:left; color:<%=coloriconomenutop%> }

.Log_Detalle{   display:inline-block }
.Log_Nombre{ font-size:8pt; margin-left:3px;   }
.Log_User{font-size:11pt; margin-left:3px;  }
 
.LOGO_EMPRESA{margin-left:10px; vertical-align:middle; padding:0;}


.EncabezadoReducido{margin-top:3px !important; padding-top:0px !important; margin-bottom:-12px !important;
background:#fff; vertical-align:middle;}

.Principal 
{
	/*border:1px solid #333 !important; */
	border:0px solid #ccc !important; 
	border-top:0px !important; 
	margin-top:0px !important; 
	width:<%=AnchoPagina%>;/*:72%;*/  
}

.PreTop TD  
{
	height:16px; vertical-align: bottom; text-align:center; padding-top:0px;
	/*border-top:1px solid #999999;*/
	/*border-top:1px solid #999999;*/
 }

.PreTop_Izq {width:300px; height:22px !important;border:0px;  }

.PreTop_Izq TR TD
{
    height:22px !important;padding:0; margin:0;
    color:<%=fuenteFecha%>;     
    background-color:<%=fondoFecha%>; 
    font-family:Tahoma; font-size:8pt; vertical-align:middle; text-align:left; 
    padding-left:3px;
    border:0px;
    /*text-transform: capitalize;  */ 
        
}
 
.PreTop_Der
{
	 
    width:100% !important; height:22px !important; 
    padding:0; margin:0; 
    padding-right:10px;
    text-align:right !important;font-size:9pt; 
    vertical-align:middle !important;   font-family:Tahoma;         
    color:<%=fuenteFecha%>;     
    background-color:<%=fondoFecha%>;
   /*   background:url(img/fondoFecha.png) repeat-x bottom left */
     
   
}

.PreTop_Der a{ text-decoration:none;    color:<%=fuenteFecha%>;      font-size:8pt; margin-left:8px; }
.PreTop_Der a:hover{ color:#333;}
.PreTop_Der img{ height:16px;}
 
 
 
.Piso 
{
	height:80px;
	background-color:<%=fondoPiso%> !important;
	border-bottom:0px ;
	border-top:22px solid <%=fondoFecha%>; 
	color:<%=fuentePiso%>  !important; 
	width:<%=AnchoPagina%>;
	height:160px;
	vertical-align:top;
}
.Piso table 
{
	/*background:url(img/FooterRHPRO.png) no-repeat bottom right; */
    color:<%=fuentePiso%>  !important;
 }


.DetEmpresa{	   font-size:8pt; font-family: Tahoma; margin-bottom:6px; 	} 

.TitBase{  font-size:9pt; font-family:Tahoma; margin-left:6px; margin-top:6px; vertical-align:top; }

.Frase{   font-size:9pt; font-family:Tahoma; height:100px; vertical-align:top}

 
  

/******************************/

/*NUEVOS*/
.PopUp_NewHome
{
	 width:300px; height:120px; background-color:#FFFFFF !important;
	 /* position: fixed; z-index:10010; left:38%; top:16%;*/
	 border:0px solid #444; 
	 
	 font-family:Tahoma,Arial; font-size:11pt; color:#888; white-space: normal !important; 
}
.PopUp_Cabecera TD
{
    background-color:#efefef !important; color:#555555!important;  height:36px; 
    font-family:Verdana !important; font-size:10pt !important;    
    padding-left:15px;
    border-bottom:1px solid #aaa !important;
    -webkit-border-top-left-radius: <%=RadioGadget%>;
    -webkit-border-top-right-radius: <%=RadioGadget%>;
    -moz-border-radius-topleft: <%=RadioGadget%>;
    -moz-border-radius-topright: <%=RadioGadget%>;
    border-top-left-radius: <%=RadioGadget%>;
    border-top-right-radius: <%=RadioGadget%>;
    vertical-align:middle !important;
 }
 
 
 
                    
.PopUp_DataUser 
{
	  text-align:center !important;background-color:#FFFFFF !important;
	 
}

.PopUp_DataUser TD 
{
    vertical-align:top !important; text-align:center !important;white-space: normal !important;    
    -webkit-border-bottom-right-radius: <%=RadioGadget%>;
    -webkit-border-bottom-left-radius: <%=RadioGadget%>;
    -moz-border-radius-bottomright:<%=RadioGadget%>;
    -moz-border-radius-bottomleft: <%=RadioGadget%>;
    border-bottom-right-radius: <%=RadioGadget%>;
    border-bottom-left-radius: <%=RadioGadget%>; 
    padding:6px;      
  background-color:#ffffff; 
  border:0px solid #ccc;
 border-bottom:0px;

}

 

.PopUp_BD  TD  
{
	 vertical-align:top !important; text-align:center !important; 
	background-color:#efefef; border:1px solid #ccc; 
	 
	border-bottom:0px;
	/*
	height:10px; 
	padding:4px;
	padding-bottom:13px;
	*/
	 
	}

.PopUp_BD_Combo { width:90% !important; font-family:Tahoma !important; color:#777 !important; font-size:9pt !important; height:160px !important;
                  border:1px solid #ccc !important; font-weight:normal;  
                 }
option{ height:32px !important;   font-size:10pt !important; font-family:Tahoma !important;   }
option:hover { background-color:#efefef !important; color:#333; cursor:pointer } 

.PopUp_Piso TD
{
    /*background-color:#efefef !important; 
    color:#333333 !important;
    -webkit-border-bottom-right-radius: <%=RadioGadget%>;
    -webkit-border-bottom-left-radius: <%=RadioGadget%>;
    -moz-border-radius-bottomright:<%=RadioGadget%>;
    -moz-border-radius-bottomleft: <%=RadioGadget%>;
    border-bottom-right-radius: <%=RadioGadget%>;
    border-bottom-left-radius: <%=RadioGadget%>;
    */
    height:36px !important;  
    vertical-align:middle !important; text-align:center !important; border:0px solid #aaa;
    padding:0px !important;
    margin:0px !important;
    
    background-color:#efefef;
     border:1px solid #ccc;
     border-top:0px;
     
    
 }

.PopUp_FondoTransparente{  
	 position: fixed; 
	 width:100% !important; height:100% !important; background-color:#000000 !important; 
     opacity:0.6;
     filter:alpha(opacity=60) !important;  
     top:0px;
     left:0px;
     z-index:10009;     
             
}

.Fondo_Contenedor_Principal
{
	height:250px; 
	/*background-color:#eef2f5;*/
	background-color: <%=fondocontppal%>;
}

.Contenedor_Ventana_Gadgets
{
	position: fixed; 
	width:50% !important; 
	height:250px !important; 	
    top:10%;
    left:26%;
    z-index:9300 !important; 
  
    -webkit-border-bottom-right-radius: <%=RadioGadget%>;
    -webkit-border-bottom-left-radius: <%=RadioGadget%>;
    -moz-border-radius-bottomright:<%=RadioGadget%>;
    -moz-border-radius-bottomleft: <%=RadioGadget%>;
    border-bottom-right-radius: <%=RadioGadget%>;
    border-bottom-left-radius: <%=RadioGadget%>;
}
.Contenedor_Ventana_Generica
{
	position: fixed; 
	width:70% !important; 
	/*height:80% !important; 	*/
    top:10%;
    left:14%;
    z-index:9300 !important; 
  
    -webkit-border-bottom-right-radius: <%=RadioGadget%>;
    -webkit-border-bottom-left-radius: <%=RadioGadget%>;
    -moz-border-radius-bottomright:<%=RadioGadget%>;
    -moz-border-radius-bottomleft: <%=RadioGadget%>;
    border-bottom-right-radius: <%=RadioGadget%>;
    border-bottom-left-radius: <%=RadioGadget%>;
}


.Contenedor_Ventana_Estilos
{
	position: fixed; 
	width:35% !important; 
	height:200px !important; 	
    top:10%;
    left:35%;
    z-index:9300 !important; 
  
    -webkit-border-bottom-right-radius: <%=RadioGadget%>;
    -webkit-border-bottom-left-radius: <%=RadioGadget%>;
    -moz-border-radius-bottomright:<%=RadioGadget%>;
    -moz-border-radius-bottomleft: <%=RadioGadget%>;
    border-bottom-right-radius: <%=RadioGadget%>;
    border-bottom-left-radius: <%=RadioGadget%>;
}

.PopUp_Informes_Error{ color:#F00; }

.Boton_Cuadrado {        
    background-color: <%=fondoFecha%> !important; 
    padding: 5px !important; 
    display: inline-block !important;  
    text-decoration: none !important; 
    margin: 2px !important; 
    border: 1px solid #666 !important; 
    cursor: pointer !important; 
    font-family: Tahoma !important; 
    font-size: 9pt !important; 
    color: <%=fuenteFecha%> !important; 
    -webkit-border-radius: <%=RadioGadget%> !important; 
    -moz-border-radius: <%=RadioGadget%> !important; 
    border-radius: <%=RadioGadget%> !important; 
 
    
}


input:focus
{ 
  border:1px solid #999 !important;
  color:#333 !important;
}

input
{   
  color:#333 !important;
  background-color:#ffffff !important;
}

input:-webkit-autofill {
  -webkit-box-shadow: 0 0 0px 1000px white inset;
}


.Boton_Cuadrado:hover{ background-color: <%=fondoCabecera%> !important; color:<%=coloriconomenutop%> !important; }

.popInput{  margin:0; padding:0; text-align:center !important; }

.popInput input{ height:29px; border:1px solid #ccc; width:80% !important; font-size:12pt; color:#333; font-family:Tahoma,Arial; text-align:center }

.cerrarVentana
{
	 float:right; margin-right:10px; font-weight:bold; cursor:pointer;color:#888; padding:6px; 
	 display:block;	 
	 /*
     -webkit-border-radius: 6px;
     -moz-border-radius:6px;
     border-radius:6px;  
     */
}
.cerrarVentana a{ float:right; margin-right:10px; font-weight:bold; cursor:pointer;color:#888; padding:6px; display:block;  	 }

.cerrarVentana:hover{color:#444; border:1px solid #aaa; padding:5px; }
.cerrarVentana:hover a{color:#444; border:1px solid #aaa; padding:5px; }




.TituloVentana{padding:0px !important; margin:0px !important; vertical-align:middle; }


.info_Login{ padding:2px; font-family:Tahoma; font-size:11pt; color:#555; vertical-align:middle; text-align:left; } 
.info_Login img{ vertical-align:middle}

.CuadradoEstilo{ border:1px solid #333; height:18px !important; width:18px !important; padding:2px; margin:10px; cursor:pointer;}

.RGBEstilo{ height:26px; width:26px; float:right; display: inline-block; border:1px solid #333; margin:2}
.RGBNombre {display:inline-block; height:30px; vertical-align:middle; margin-left:2px;}

.LinkMega 
{
	line-height:5 !important; 
	margin-right:6px !important;}
/****************************/
.Globo_Lenguajes 
{
	                
     position: fixed; width:400px; height:120px; background-color:#FFFFFF !important; z-index:9001; left:35%; top:16%;
	 border:1px solid #444; 
	 -webkit-border-radius: <%=RadioGadget%>;
     -moz-border-radius: <%=RadioGadget%>;
     border-radius: <%=RadioGadget%>; 
	 font-family:Tahoma,Arial; font-size:11pt; color:#888; white-space: normal !important; 
	/* visibility:hidden; */
                  
   }
.Globo_Lenguajes a{ color:#999999; font-family:Tahoma; font-size:9pt; text-decoration:none}   
.Link_Lenguajes {width:99%; text-align:left; padding-left:4px; margin-right:2px; margin-top:3px;margin-bottom:4px; font-size:10pt;  }
.Link_Lenguajes:hover{ background-color:#cccccc;  }
.Link_Lenguajes:hover a{   color:#333333 !important;}
.DIV_Globo_Idiomas{ /*position:absolute;   top:0px */}

 
.ContenidoControlMenuTop a:hover{ background:none !important; border:0 !important}

.TablaIdioma{ height: 200px !important; width: 680px  !important; border: 1px solid #333;}

.TablaIdioma TD{ text-align:left }

.EtiquetaIdioma  {
    white-space: nowrap;        
    padding: 2px; 
    vertical-align:middle;
   /* color: #333333; */
    font-size: 9pt; font-family: Arial;
    /*display: inline-block;*/
    display:inline-block !important;
    width:185px;    
    margin: 8px; 
    /*margin-right:13px;*/
    overflow:hidden;     
    /*background-color:#efefef;       */
    border: 1px solid #cccccc;
    /*
    -webkit-border-radius: 6px;    
    -moz-border-radius: 6px;
    border-radius: 6px;
    */
    text-align:left;
    cursor:pointer;
    float:left;
    line-height:2.5;
    color:#333;
    }
.EtiquetaIdioma a {   color:#333333;} 
.EtiquetaIdioma:hover 
{
	 background-color:#efefef;  
}

.BtnOpc{ margin-right:10px;   }

.BarraPisoFavoritos  
{
	border-top:1px solid #aaa; width:100%; height:35px; padding:0px; background-color:#efefef; text-align:right;
	-webkit-border-bottom-right-radius: <%=RadioGadget%>;
    -webkit-border-bottom-left-radius:<%=RadioGadget%>;
    -moz-border-radius-bottomright: <%=RadioGadget%>;
    -moz-border-radius-bottomleft: <%=RadioGadget%>;
    border-bottom-right-radius: <%=RadioGadget%>;
    border-bottom-left-radius: <%=RadioGadget%>;
}

.Contenedor_Fav 
{
/*	height:400px; overflow:auto;*/
 white-space: normal;
}

.Contenedor_Estilos
{
	/*height:255px; */
	/*overflow-y:auto; */	
	/*width:500; */
	text-align:left;
}

.Contenedor_Estilos a
{ display:inline-block}

.Contenedor_Gadgets
{
/*	height:300px; overflow-y:auto; width:100%;*/ text-align:left;
 
}
.TituloDetModulo{ margin-left:10px;  vertical-align:middle !important; display:inline-block; padding:3px; line-height:2  }
.TituloDetModulo:hover{ border:1px dashed  #ccc; padding:2px;}

.BotonCabeceraGadget
{
	  cursor:pointer;color:#888; padding:4px; 	 
     -webkit-border-radius: 6px;
     -moz-border-radius:6px;
     border-radius:6px;       
     display:inline-block;
     vertical-align:middle;
     text-align:center;
      
}
  
.BotonCabeceraGadget:hover{color:#444; border:1px solid #aaa; padding:3px; }
.BotonCabeceraGadget:hover span{color:#444; border:1px solid #aaa; padding:3px; }

.BotonCabeceraModulos
{
	 cursor: default;    
     margin-left:0px !important;
     padding:0px; 	 
     display: table-caption;
     margin-left:0px !important; 	
    color:<%=FuenteCabeceraGadget_Color%>;
	font-family:<%=FuenteCabeceraGadget_Font%>;
    font-size:<%=FuenteCabeceraGadget_Size%>;    
    background-color:<%=BackgroundCabeceraGadget%>;
    
       
}

.BotonCabeceraModulos span img
{ margin:0px !important; padding:0px !important;      color:<%=coloricono%>; }
  
 

 
 

.TituloCabeceraModulo
{
	  margin-left:0px !important;
     padding:0px; 	 
     display: table-caption;
     margin-left:0px !important;
     white-space: nowrap;
 
     }

.SeccionFavorito
{
    white-space: nowrap;        
     color: #333333; 
    font-size: 8pt; font-family: Arial;
    display:  inline-block;
    width:200px;    
    margin: 3px; 
   /* margin-right:9px;*/
    overflow:hidden;
    border:1px solid #cccccc; 
    vertical-align:top;
   /*
    -webkit-border-radius: 6px;
    -moz-border-radius: 6px;
    border-radius: 6px;
    */
    text-align:left;
/*    float: left;*/
}

.EtiquetaFavoritoModulo
{
	white-space: nowrap;        
    color: #333333; 
    font-size: 8pt; font-family: Arial;
    display: list-item;
    width:216px;    
    margin: 0px;  
    overflow:hidden;  
    background-color:#efefef; 
    margin:0;
    padding:4px; 
    border-bottom:1px solid #ccc;
}


.EtiquetaFavorito
{
	white-space: nowrap;        
    padding: 2px; color: #333333; 
    font-size: 8pt; font-family: Arial;
    display: list-item; 
    width:200px;    
    margin:0px; 
    margin-right:0px;
    overflow:hidden;
     cursor:pointer;
  
}

.EtiquetaFavorito:hover 
{
	 background-color:#cccccc;  
}

 


.DIV_Gadget_Config  
{
 /*  white-space: nowrap;        
    padding: 5px; 
    color: #333333; 
    font-size: 9pt; font-family: Arial;
 
    display:inline-block !important;
    width:170px;    
    margin: 8px; 
    margin-right:13px;
    overflow:hidden;     
    background-color:#efefef;       
    border: 1px solid #cccccc;
    -webkit-border-radius: 6px;    
    -moz-border-radius: 6px;
    border-radius: 6px;
    text-align:left;
    cursor:pointer;
 */
                     
        margin: 2px !important;            
        border: 1px solid #ccc !important;
        color: #333333 !important;
        font-size: 9pt !important;
        font-family: Arial !important;
        display: inline-block !important;          
        color: #333 !important;
        width:160px !important;
        padding:0px !important; 
        float:left !important;
         /*  width:185px !important;   */
        
    }


 
.DIV_Gadget_Config a{ color:#333 !important; }

.Globo_Idiomas_Centro
{
	/*background:url(img/Loguin/Globo_Centro.png) repeat-y center;*/
	} 

.Cerrar {background:url(img/Loguin/Globo_Centro.png) repeat-y center;} 

.Globo_Idiomas_Centro a{ text-decoration:none; font-family: Verdana; font-size:10pt; color:#333333;}
.Globo_Idiomas_Centro a:hover{ font-weight:bold;}
.Globo_Idiomas_Centro a:visited{  }
 

/*****************************/

.ContenedorPrincipal{background:url(img/Contenedor-der.png) repeat-y right; background-color:#FFFFFF; width:2px}
.SeccionCentral{ white-space:nowrap !important}

.AperturaModulo{ font-family: Tahoma; font-size:12pt; color:#999; font-weight:bold; padding:5px; margin-right:5px}
span.AperturaModulo:hover{color:#008ABE !important; border:1px solid #999; padding:4px; }

/***************************************************/

.ASPAccesos
{ 
 background-color:transparent; text-decoration:none; font-family: Tahoma; font-size:11pt; color:#FFFFFF; 
 cursor: pointer; margin-left:10px;
}


.ASPAccesos:hover
{ 
  color:#CCCCCC;  
}

.ListaOculta
{
    height:0%;
	width:0%;
	position:absolute;
	left:0; top:0;
	visibility:hidden;
}

.FondoOculto
{
	height:100%;
	width:100%;
	position:absolute;
	left:0; top:0;	 
	vertical-align:middle;
	text-align:center;
	background-color:#000000;
	opacity:0.4;
    filter:alpha(opacity=40);
     z-index:2000;
}
.ListaGadgets 
{
	 color:#FFFFFF; 
	 font-size:15pt; 
	 position: absolute; 
	 top: 100px; 
	 vertical-align:middle; 
	 text-align:center; 
	 z-index:2001;
} 
 
/*****************************************************/

/**/

.BotonMas img{ margin:4px  !important;   }
.BotonMas:hover img{  border:1px solid #555;margin:3px  !important; background-color:#000  }
.DescModulos {color:#777 !important; font-family:Arial !important; font-size:8pt !important;}

.ASPlink 
{ 
 text-decoration:none; font-family: Arial; font-size:10pt; color:<%=fuenteModulos%> !important;  
 cursor: pointer;
 vertical-align:middle !important;
 display: table-caption;
 
 width:100% !important;
 float:left !important;
 
 
}

 
.ASPlink  span,.ASPlink  div
{  
 width:100%; line-height:27px;  
 margin-left:3px; vertical-align:middle !important; text-align:left; 
 display:inline-block;
}

.ASPlink  img
{  
 margin-left:3px; vertical-align:middle; text-align:center
}
 

/* Estilos - Configura Gadgets */


.ContenedorPrincipalDeGadgets{ width:99%;    }
/*.BordeGris{ border:1px #dddddd solid; margin-top:4px; margin-bottom:5px;   margin-left:0px;     }
.PisoGris{ border-bottom:1px #cccccc solid; background-color:#efefef; height:30px; text-align:left; color:#666666; font-family:Arial; width:100%;  }
.TopeGris{ border-top:1px #999999 solid; background-color:#f0f0f0; height:0px}
*/


.GadgetFlotante  {
        white-space: nowrap;        
        padding: 3px; color: #333333;  border: 0px solid #888888;
        font-size: 9pt; font-family: Arial;
        display: inline-block;        
        overflow:hidden;        
        vertical-align:top;
        float:left;
        position:relative;
        padding-left:1px;
        padding-right:2px;
        margin-left:1px;
        margin-right:2px;
        margin-bottom:2px !important;  
        margin-top:5px !important;  
        
       
    }

.BordeGris
{
    border:1px #CCC5C4 solid; 
    margin-bottom:0px;   margin-left:0px;  margin-top:0px;/*0.6px;*/
    -webkit-border-top-left-radius: <%=RadioGadget%>;
    -webkit-border-top-right-radius: <%=RadioGadget%>;
    -moz-border-radius-topleft:  <%=RadioGadget%>;
    -moz-border-radius-topright:  <%=RadioGadget%>;
    border-top-left-radius:  <%=RadioGadget%>;
    border-top-right-radius: <%=RadioGadget%>;
    border-bottom:0px;    
 }
 
 .BordeGris TD{white-space:normal;}
 
.BordeGris a{ text-decoration:none; color:<%=coloricono%>;}
 
.PisoGris 
{
    border-bottom:1px #ccc solid;
    height:30px;	
    color:<%=FuenteCabeceraGadget_Color%>;
	font-family:<%=FuenteCabeceraGadget_Font%>;
    font-size:<%=FuenteCabeceraGadget_Size%>;    
    background-color:<%=BackgroundCabeceraGadget%>;
    width:100%;
    
}
.PisoGris TD
{
	font-weight:bold; 
	border:1px solid <%=BackgroundCabeceraGadget%>;
    
    -webkit-border-top-left-radius:  <%=RadioGadget%>;
    -webkit-border-top-right-radius:  <%=RadioGadget%>;
    -moz-border-radius-topleft:  <%=RadioGadget%>;
    -moz-border-radius-topright:  <%=RadioGadget%>;
    border-top-left-radius:  <%=RadioGadget%>;
    border-top-right-radius:  <%=RadioGadget%>;
    
}

.PisoGris  TD, .BordeGris TD{text-align:left;   }

.HiperLink{ color:<%=coloricono%> !important;}

.TopeGris
{
    border-top:1px #CCC5C4 solid; 
    background-color:#f0f0f0; height:0px;
 
    
}

.InfoModulo_MenuTop
{
   
    -webkit-border-top-left-radius:  <%=RadioGadget%>;
    -webkit-border-bottom-left-radius:  <%=RadioGadget%>;
    -moz-border-radius-topleft:  <%=RadioGadget%>;
    -moz-border-radius-bottomleft:  <%=RadioGadget%>;
    border-top-left-radius: <%=RadioGadget%>;
    border-bottom-left-radius:  <%=RadioGadget%>;
}
.InfoModulo_MenuTop a
{
color:<%=coloricono%> !important;
font-weight: bold;
font-size:13pt;

}

.InfoModulo_MenuTop:hover
{
	background-color:none !important;
}
.CabeceraDrag 
{
    width:100%;top:0px; left:0px;
    background-color: <%=BackgroundCabeceraGadget%>;
    -webkit-border-top-left-radius: <%=RadioGadget%>;
    -webkit-border-top-right-radius: <%=RadioGadget%>;
    -moz-border-radius-topleft:  <%=RadioGadget%>;
    -moz-border-radius-topright:  <%=RadioGadget%>;
    border-top-left-radius:  <%=RadioGadget%>;
    border-top-right-radius: <%=RadioGadget%>;

    border-bottom:1px solid #cccccc;
}

.ContenedorGadget_Alto
{ 
	height:330px; 
	
  }
 
.ContenedorGadget
{
	overflow-x:hidden !important;
	overflow-y:auto !important;
	/*height:304px; */
	height:296px;
	width:100%; 
	color:#cccccc !important; font-family:Tahoma  !important; font-size:9pt; font-weight:none !important;
	

}
.ContenedorGadgetAltoFull
{
	width:100%; 
	color:#cccccc !important; font-family:Tahoma  !important; font-size:9pt; font-weight:none !important;
}
 
.ContenedorGadget a{color:#333333; font-family:Tahoma  !important; font-size:9pt; text-decoration:none; font-weight:none !important;}
 
.ContenedorGadget a:hover{color:#000000;}

/*.ContenedorGadget DIV{ margin:0px !important; padding:0px !important; margin-left:6px !important; margin-top:6px !important; }*/
.ContenedorGadget_NOIMG{  margin:0px !important; padding:0px !important; margin-left:2px !important; margin-top:4px !important;  }


.PanelModulos
{
 /*border-right:1px solid #aaa; */
 /*background-color:<%=fondoModulos%>;*/
 /*background-color:#eef2f5;*/
 
 	background-color: <%=fondoModulos%>;
 
 /*width:20px;*/
  width:28px !important;
 overflow:hidden;
 padding:8px !important;
 padding-top:0px !important;
}
.ContenedorModulos{ width:100%;}
 
.ContenedorModulo
 {
    border:1px #CCC5C4 solid; 
    margin-bottom:0px;   margin-left:0px;  
    -webkit-border-radius: <%=RadioGadget%>;
    -moz-border-radius: <%=RadioGadget%>;
    border-radius: <%=RadioGadget%>;  
    border-bottom:0px;
    width:99% !important;
    margin-top:8px !important;
    margin-bottom:6px;
    margin-right:5px;
  
    
 }
.ContenedorModulo_Cab   
{/*
    -webkit-border-top-left-radius: <%=RadioGadget%>;
    -webkit-border-top-right-radius: <%=RadioGadget%>;
    -moz-border-radius-topleft:  <%=RadioGadget%>;
    -moz-border-radius-topright:  <%=RadioGadget%>;
    border-top-left-radius:  <%=RadioGadget%>;
    border-top-right-radius: <%=RadioGadget%>;*/
    background-color:<%=BackgroundCabeceraGadget%>;
} 

.ContenedorModulo_Cab TD  
{
 
    border-bottom:1px solid #ccc;
    padding:1px;
    height:32px;
    
      color:<%=FuenteCabeceraGadget_Color%>;
	font-family:<%=FuenteCabeceraGadget_Font%>;
    font-size:<%=FuenteCabeceraGadget_Size%>;    
    background-color:<%=BackgroundCabeceraGadget%>;
}
 
.ContenedorModulo_Info TD
{	    
    color:#777777; font-family:Tahoma; font-size:10pt; 
    text-align:justify;
    padding:8px;
    height:27px;
    border-bottom:1px solid #CCC5C4;
    -webkit-border-bottom-right-radius: <%=RadioGadget%>;
    -webkit-border-bottom-left-radius: <%=RadioGadget%>;
    -moz-border-radius-bottomright: <%=RadioGadget%>;
    -moz-border-radius-bottomleft:  <%=RadioGadget%>;
    border-bottom-right-radius:  <%=RadioGadget%>;
    border-bottom-left-radius:  <%=RadioGadget%>;
    background-color:#FFFFFF;
   
}  

.BarraDesplazamientoModulos{ width:29px;    }  

.BarraDesplazamientoModulos:hover{ background:url(img/FONDO_MENU_BAR.png) repeat-x bottom left; cursor:pointer; width:29px;  }  

.ContenedorBarraNavegacion
{
    background-color:<%=BackgroundCabeceraGadget%>;    
    color:<%=fuenteModulos%>; 
    -webkit-border-radius: <%=RadioGadget%>;
    -moz-border-radius:<%=RadioGadget%>;
    border-radius: <%=RadioGadget%>;
    margin:0px;
    margin-right:0px !important;
    margin-left:2px !important;    
    margin-top:8px !important;
    width:99% !important;    
    border:1px solid #ccc;     
    height:25px;
    border-bottom:1px solid #cccccc;
    display:inline-table;
}


.sm-blue { border:0;}

.ContenedorBarraNavegacion a{   color:<%=fuenteModulos%>;}

 
 
.InfoModulos {
		 font-family:Tahoma !important;
		 font-size:10pt;
		 color:#333;
		 
		 width:600px;
		 text-align:justify;
		 margin-top:9pt;		 
		 border:1px solid #CCCCCC;
		 background-color:#FFFFFF; 
		 padding:4px;
		 margin-left:12px; 		 	
		//margin-left:0px; 		
		 }
		 
.TopeInfoModulos {
		 font-family:Arial;
		 font-size:11pt;
		 color:#333;
		 background-color: transparent; 	 
		 margin-top:8px;
		}	 

 

.Menu_Links, .Menu_Links_Inact,.Menu_Links_Colapsado,.Menu_Links_ColapsadoInact 
{
    border:1px #CCC5C4 solid;     
    -webkit-border-top-left-radius:  <%=RadioGadget%> !important;
    -webkit-border-top-right-radius:  <%=RadioGadget%> !important;
    -moz-border-radius-topleft:  <%=RadioGadget%> !important;
    -moz-border-radius-topright:  <%=RadioGadget%> !important;
    border-top-left-radius:  <%=RadioGadget%> !important;
    border-top-right-radius:  <%=RadioGadget%> !important;
    margin-left:0px;  
    margin-top:8px !important; 
    margin-bottom:8px;  
    width:<%=AnchoMenuLinks%>;
} 

 

 
    
.Menu_Links TR TH,.Menu_Links_Inact TR TH,.Menu_Links_Colapsado TR TH, .Menu_Links_ColapsadoInact TR TH
{
    cursor:default;
    color:<%=FuenteCabeceraGadget_Color%>;
    font-family:<%=FuenteCabeceraGadget_Font%>;
    font-size:<%=FuenteCabeceraGadget_Size%>;
    background-color:<%=BackgroundCabeceraGadget%>;
    text-align:left;
    padding:0px;	
    padding-left:6px;    
    -webkit-border-top-left-radius:  <%=RadioGadget%> !important;
    -webkit-border-top-right-radius:  <%=RadioGadget%> !important;
    -moz-border-radius-topleft:  <%=RadioGadget%> !important;
    -moz-border-radius-topright:  <%=RadioGadget%> !important;
    border-top-left-radius:  <%=RadioGadget%>  !important;
    border-top-right-radius:  <%=RadioGadget%>  !important;    
    border:0px !important;
    border-bottom:1px solid #ccc !important;
  /*height:40px !important; */
    overflow:hidden !important;
	white-space:nowrap !important;
}
 
 
 
.Menu_Links TR  TD ,.Menu_Links_Inact TR TD, .Menu_Links_Colapsado TR TD, .Menu_Links_ColapsadoInact TR TD
{
    color:<%=fuenteModulos%> !important;
    font-family:Verdana; 
    font-size:9pt; 
    border-bottom:0px solid #f2f2f2;
    cursor:pointer;    
    background-color:#FFFFFF;
    vertical-align:middle !important;  
    /*height:37px !important; */
    overflow :hidden !important;
    
	white-space:nowrap !important;
}

.Menu_Links_Colapsado,.Menu_Links_ColapsadoInact {width:36px !important} 

.Menu_Links_Colapsado TR TD,.Menu_Links_Colapsado TR TH ,.Menu_Links_ColapsadoInact TR TH ,.Menu_Links_ColapsadoInact TR TD
{
	column-width:30px !important;
		 
	overflow:hidden !important;
	white-space:nowrap !important;
	height:32px !important;
 
}
 
.Menu_Links TR:hover TD ,.Menu_Links TR:hover TD  IMG,.Menu_Links_Colapsado TR:hover TD
{    
 background-color: #cccccc !important; 
 border:0; 
 
}
 
.Menu_Links TR TD a,.Menu_Links_Inact TR TD a, .Menu_Links_ColapsadoInact TR TD a{color:<%=fuenteModulos%> !important;  }


 
		
.tooltiphelp {
	position:absolute;
	visibility:hidden;	
	overflow: visible;
	background-color:transparent;
 
	margin-left:10px;  /*-5px;*/
	margin-top: 1px;  /*-10px;*/
	display:none;
	//margin-left: -101px  ;/*-60px;  */  /*IE*/
	//margin-top: 17px ;/*10px;    */  /*IE*/
}
.contenidoTooltip {
	background-color: transparent;   
	color:#FFFFFF;
	font-family:Arial;
	font-size:8pt;
	font-weight:bold;	}

.contenidoTooltip a{	   
	color:#FFFFFF;
	font-family:Arial;
	font-size:8pt;
	font-weight:bold;
    text-decoration:none;	}

.contenidoTooltip a:hover{ color:#333333; }

.tool{ border:1px #666 solid; background-color:#333;opacity:0.8;filter:alpha(opacity=80); }

 
 
#Transparente { width:100% !important;
                height:100% !important;
                margin:0px !important;
                padding:0px !important;
                background-color:#000000;
                z-index:10002;
                position: fixed;
                top:0px;
                left:0px;
                visibility: hidden;
                opacity:0.6;
                filter:alpha(opacity=60); /* For IE8 and earlier */
                }

#Progreso 
{
	visibility: hidden;
	font-family: Arial;
	font-size:10pt;
	color:#333333;
	position: fixed;
	z-index:10001;
	top:15%;
	left:45%;
	height:80px;
	width:100px;
	background-color:#FFFFFF;
 
	text-align:center;
	padding-top:8px;	
}



.InfoCambioPass
{
	color:#333;  font-family:Tahoma,Arial; font-size:9pt; margin-left:8px; margin-bottom:2px; padding:0;
}  
.InfoCambioPass a
{
	color:#333;  font-family:Tahoma,Arial; margin-left:8px; text-decoration:none;  
}  


.TopeG
{
	  color:#333333; font-family:Arial; font-size:10pt; font-weight:bold; text-align:left; 
	 background-color:#efefef; padding:3px;
}

 .LinksG{color:#333333; font-family:Arial; font-size:10pt; text-align:left !important; }
 .LinksG div{ margin-bottom:5px; text-align:left; margin-top:3px; }
 .TopeContenedorG 
 {
 	font-family:Arial; font-size:12pt; font-weight:normal; color:#333333 !important; background-color:#cccccc !important 
 	
 }
 
 
 .MiniG 
 {
 	/*background:url(img/FondoGadget1.png) no-repeat center;*/
 	border:1px #bbb solid; margin-bottom:0px;   margin-left:0px;  margin-top:6px;/*0.6px;*/
    /*
    -webkit-border-radius: 16px;
    -moz-border-radius: 16px;
    border-radius: 16px;
    
    padding:5px;
    */
    margin:2px;
  
 }
 .MiniG span { margin-left:5px;}

 .TG  
 {
    border:1px solid #555;top:15%; position:fixed; visibility:hidden;    
    -webkit-border-radius: <%=RadioGadget%>;
    -moz-border-radius: <%=RadioGadget%>;
    border-radius:<%=RadioGadget%>;
   z-index:10002;
   left:30%;
 }
 
 .TG_Encab td
 {	 border:1px solid #ccc;     
    -webkit-border-top-left-radius: <%=RadioGadget%>;
    -webkit-border-top-right-radius: <%=RadioGadget%>;
    -moz-border-radius-topleft:  <%=RadioGadget%>;
    -moz-border-radius-topright:  <%=RadioGadget%>;
    border-top-left-radius:  <%=RadioGadget%>;
    border-top-right-radius: <%=RadioGadget%>;
    border-bottom:0px;
    background-color:#cccccc;
    
 }
 
 .TG_Piso td
{	border:1px solid #ccc;
 	background-color:#cccccc; 	   
    -webkit-border-bottom-right-radius: <%=RadioGadget%>;
    -webkit-border-bottom-left-radius:<%=RadioGadget%>;
    -moz-border-radius-bottomright: <%=RadioGadget%>;
    -moz-border-radius-bottomleft: <%=RadioGadget%>;
    border-bottom-right-radius: <%=RadioGadget%>;
    border-bottom-left-radius: <%=RadioGadget%>;
}
 
 .Transparente 
 {
 	position:fixed; width:100%; height:100%; background-color:#000; opacity:0.6;filter:alpha(opacity=60);  top:0px; left:0px; visibility:hidden;
 	   z-index:10001;
 	}
 
 /*.BtnG {margin-left:5px; cursor:pointer; padding:4px;}
 .BtnG:hover { color:#FF0000; }*/
 .BtnG {  cursor:pointer; padding:0px; display:inline-block; float:right;color:#444; background-color:#555 !important}
 .BtnG:hover {/* color:#FF0000; border:1px solid #aaa; padding:1px;*/ }
 
 

.EtiquetaGadgets  {
    white-space: nowrap;        
    padding: 2px; 
    vertical-align:middle; 
    font-size: 9pt;
    font-family: Arial; 
    display:inline-block !important;
    width:185px;
    margin: 8px;  
    overflow:hidden;      
    border: 1px solid #cccccc; 
    text-align:left;
    cursor:default;
    float:left;
    line-height:2.3;
    color:#333;
    }
    
 
 
.SLIDER_BTN
{
	line-height:1.0;
	width:56px;
	text-align:center;    
} 
.SLIDER_Inact,.SLIDER_Act{  padding:1px; border:0px solid #333;
-webkit-border-radius: 8px;
-moz-border-radius:8px;
border-radius: 8px;
}

.SLIDER_Inact{ background-color:#ccc}
.SLIDER_Act{background-color:Green}


.SLIDER{margin:1px; margin-top:3px;line-height:1.6; cursor:pointer; padding:0px; 
        display:inline-block; float:right; vertical-align:middle !important;       
   
}

.SLIDER TR TD{width:22px; padding:1px}

.SLIDER TR:hover TD{ color:#333 !important;}
   
.SliderOscuroON{background-color:#555 !important; display:inline-block; line-height:1.6;
-webkit-border-top-left-radius: 4px;
-webkit-border-bottom-left-radius: 4px;
-moz-border-radius-topleft: 4px;
-moz-border-radius-bottomleft: 4px;
border-top-left-radius: 4px;
border-bottom-left-radius: 4px;

 
}

.SliderClaroON{background-color:#83b64e; color:#fff; vertical-align:middle; text-align:center; line-height:1.6; font-size:8pt; font-weight:bold;
               -webkit-border-top-right-radius: 4px;
-webkit-border-bottom-right-radius: 4px;
-moz-border-radius-topright: 4px;
-moz-border-radius-bottomright: 4px;
border-top-right-radius: 4px;
border-bottom-right-radius: 4px;
border:1px solid #555;
 
}

.SliderOscuroOFF{background-color:#555 !important; display:inline-block;   line-height:1.6;  
                 -webkit-border-top-right-radius: 4px;
-webkit-border-bottom-right-radius: 4px;
-moz-border-radius-topright: 4px;
-moz-border-radius-bottomright: 4px;
border-top-right-radius: 4px;
border-bottom-right-radius: 4px; 
 }

.SliderClaroOFF{background-color:#999999; color:#fff; vertical-align:middle; text-align:center; line-height:1.6;  font-size:8pt; font-weight:bold;
-webkit-border-top-left-radius: 4px;
-webkit-border-bottom-left-radius: 4px;
-moz-border-radius-topleft: 4px;
-moz-border-radius-bottomleft: 4px;
border-top-left-radius: 4px;
border-bottom-left-radius: 4px;
border:1px solid #555;
 }

 
 

.GadgetNombre {display:inline-block; height:30px; vertical-align:middle; margin-left:2px; 
               width:128px; 
            /*   width:110px; */
               overflow:hidden}
 
 
.BtnTransparenteTop {cursor:pointer !important; position:absolute !important;
                     height:100% !important;margin:0 !important;padding:0 !important; width:100% !important; 
                     background-color:transparent !important}  
 
 
 .BtnTransparenteAyuda {cursor:pointer !important; position:absolute !important;
                     height:100% !important;margin:0 !important;padding:0 !important; width:100% !important; 
                     background-color:transparent !important;}  

.BtnTransparenteOcultar {cursor:pointer !important; position:absolute !important;
                     height:30px !important;margin:0 !important;padding:0 !important; width:30px !important; 
                     background-color:transparent !important}  
 
 
.CajaBtnTop{line-height:60px; padding-left:2px; padding-right:2px;} 
.CajaBtnTop div {min-width:90px; text-align:center; margin-right:5px;}

 
 </style>
 
 