﻿<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="sm-clean.aspx.cs" Inherits="RHPro.css.sm_clean" %>

 <% 
    Response.ContentType = "text/css"; 
 %> 
 
 <% 
     String l_coloricono;
     String l_coloriconomenutop;
     String l_fuenteFecha;
     String l_fondoFecha;
     String l_fuentePiso;

     if (Session["EstiloR4_coloricono"] != "")
         l_coloricono = (String)Session["EstiloR4_coloricono"];
     else
         l_coloricono = "#FFFFFF";

     if (Session["EstiloR4_coloriconomenutop"] != "")
         l_coloriconomenutop = (String)Session["EstiloR4_coloriconomenutop"];
     else
         l_coloriconomenutop = "#FFFFFF";


     if (Session["EstiloR4_fuenteFecha"] != "")
         l_fuenteFecha = (String)Session["EstiloR4_fuenteFecha"];
     else
         l_fuenteFecha = "#ffffff";

     if (Session["EstiloR4_fondoFecha"] != "")
         l_fondoFecha = (String)Session["EstiloR4_fondoFecha"];
     else
         l_fondoFecha = "#999999";
     %>


/* Menu box
===================*/
#main-menu{ display:-ms-inline-flexbox;position:fixed;    }
 
 
	/*
	.sm-blue {
	 
		background:#7d8594;   
	}
	*/
	.sm-blue-vertical {
		 
	}
	.sm-blue ul {	
		border:1px solid #a9a9a9;
		padding:0px 0;	
		background:#fff;
		/*-moz-border-radius:0 0 4px 4px;
		-webkit-border-radius:0 0 4px 4px;
		border-radius:0 0 4px 4px;*/
		-moz-box-shadow:0 5px 12px rgba(0,0,0,0.3);		
		-webkit-box-shadow:0 5px 12px rgba(0,0,0,0.3);
		box-shadow:0 2px 16px rgba(0,0,0,0.3);
	}
	/* for vertical main menu subs and 2+ level horizontal main menu subs round all corners */
	.sm-blue-vertical ul,
	.sm-blue ul ul {
		 
	}

/*******************/

 

.BarraNavegacion > li {
    float:left;
    font-family:Tahoma;
    font-size:9pt;
}

.BarraNavegacion li   {
    background:#efefef;
    color:#777777;
    display:block;
    border:0px solid #888888;
    padding:6px;
    border-right:1px solid #ccc !important; 
    height:18px;
}

.SubBarraNavegacion{ margin-left:-6px;}
.SubBarraNavegacion li   {
     height:14px;
     border-right:0px;
     
}

.BarraNavegacion li .flecha{
    font-size: 9pt;
    padding-left:0px;
    display: none;
}

.BarraNavegacion li:not(:last-child) .flecha {
display: inline;
}

.BarraNavegacion li:hover 
{
cursor:default;
color:#333333; 
 
background:url(img/FONDO_MENU_BAR.png) repeat-x bottom right #FFFFFF;
 
 
 
}

.BarraNavegacion li {
    position:relative;

}

.BarraNavegacion li ul {
    display:none;
    position:absolute;
    min-width:190px;
    top:30px;
    border:1px solid #ccc;
 
}

.BarraNavegacion li:hover > ul {
    display:block;
    cursor:default;
}

.BarraNavegacion li ul li ul {
    right: -192px; 
    top:0;
}

 
/*********************/
/* Menu items
===================*/

	.sm-blue a {
		padding:5px; 
		padding-bottom:2px; 
		font-size:8pt;
		line-height:26px;
		font-family:Verdana;		 
		text-decoration:none;
		 	
		/*border-bottom:1px solid #efefef; */
		 
	}
	.sm-blue a:hover, .sm-blue a:focus, .sm-blue a:active,
	.sm-blue a.highlighted {	 
		color:#000000;
		
		/*background-color:#cccccc;*/
		background:<%=l_fuenteFecha%>; /*#999999;*/  /*JPB: sobre los submenues*/
		color:<%=l_fondoFecha%>;
		
		cursor:pointer;
	}
	.sm-blue-vertical a {
		padding:9px 40px 8px 23px;
		background:#efefef; /* Old browsers */
		
		 
	}
	.sm-blue ul a {
		padding:0px;
		padding-right:25px;
		background:transparent;
		color:#666666;
		font-size:9pt;
		text-shadow:none;
	}
	
	.sm-blue ul a:hover, .sm-blue ul a:focus, .sm-blue ul a:active,
	.sm-blue ul a.highlighted {
		background:<%=l_fondoFecha%>; /*#999999;*/  /*JPB: sobre los submenues*/
		color:<%=l_fuenteFecha%>;
		cursor:pointer;
	 
	}
	/* current items - add the class manually to some item or check the "markCurrentItem" script option */
	.sm-blue a.current, .sm-blue a.current:hover, .sm-blue a.current:focus, .sm-blue a.current:active,
	.sm-blue ul a.current, .sm-blue ul a.current:hover, .sm-blue ul a.current:focus, .sm-blue ul a.current:active {
		background:#555555;
		 
		color:#fff;
		 
	}
	/* round the left corners of the first item for horizontal main menu */
	.sm-blue > li:first-child > a {
		 
	}
	/* round the corners of the first and last items for vertical main menu */
	.sm-blue-vertical > li:first-child > a {
		 
	}
	.sm-blue-vertical > li:last-child > a {
		 
	}
	.sm-blue a.has-submenu {
        /*border-bottom:1px solid #CCC;*/
	}

 
/* Sub menu indicators
===================*/

	.sm-blue a span.sub-arrow {
		position:absolute;
		bottom:2px;
		left:50%;
		margin-left:0px;
		/* we will use one-side border to create a triangle so that we don't use a real background image, of course, you can use a real image if you like too */
		width:0;
		height:0;
		overflow:hidden;
	
		 
	}
 	.sm-blue-vertical a span.sub-arrow,
 	.sm-blue ul a span.sub-arrow {
		bottom:auto;
		top:50%;
		margin-top:-5px;
		right:0px;
		left:auto;
		margin-left:0;
		border-width:3px;
		border-style:solid solid solid solid;
		border-color:transparent transparent transparent #111111;
	}


 	.sm-blue ul a:hover span.sub-arrow {
 
		border-color:transparent transparent transparent #ffffff;
	}

/* Items separators
===================*/

	.sm-blue li {
		border-right:1px solid #ccc;
	 
	}
	.sm-blue li:first-child,
	.sm-blue-vertical li,
	.sm-blue ul li {
		border-left:0;
	}


/* Scrolling arrows containers for tall sub menus - test sub menu: "Sub test" -> "more..." -> "more..." in the default download package
===================*/
/*
	.sm-blue span.scroll-up, .sm-blue span.scroll-down {
		position:absolute;
		display:none;
		visibility:hidden;
		overflow:hidden;
		background:#ffffff;
		height:20px;	
	}
 */

 
	.sm-blue span.scroll-up, .sm-blue span.scroll-down {
		position:absolute;
		display:none;
		visibility:hidden;
		overflow:hidden;
		background:#ffffff;
		height:20px;
		/* width and position will be automatically set by the script */
	}
	.sm-blue span.scroll-up-arrow, .sm-blue span.scroll-down-arrow {
		position:absolute;
		top:-2px;
		left:50%;
		margin-left:-8px;
		/* we will use one-side border to create a triangle so that we don't use a real background image, of course, you can use a real image if you like too */
		width:0;
		height:0;
		overflow:hidden;
		border-width:8px; /* tweak size of the arrow */
		border-style:dashed dashed solid dashed;
		border-color:transparent transparent #555 transparent;
	}
	.sm-blue span.scroll-down-arrow {
		top:6px;
		border-style:solid dashed dashed dashed;
		border-color:#333 transparent transparent transparent;
	}





/*
---------------------------------------------------------------
  Responsiveness
  These will make the sub menus collapsible when the screen width is too small.
---------------------------------------------------------------*/


/* decrease horizontal main menu items left/right padding to avoid wrapping */
@media screen and (max-width: 1px) {
	.sm-blue:not(.sm-blue-vertical) > li > a {
		padding-left:1px;
		padding-right:1px;
	}
}
@media screen and (max-width: 1px) {
	.sm-blue:not(.sm-blue-vertical) > li > a {
		padding-left:1px;
		padding-right:1px;
	}
}

@media screen and (max-width: 1px) {

	/* The following will make the sub menus collapsible for small screen devices (it's not recommended editing these) */
	ul.sm-blue{width:auto !important;}
	ul.sm-blue ul{display:none;position:static !important;top:auto !important;left:auto !important;margin-left:0 !important;margin-top:0 !important;width:auto !important;min-width:0 !important;max-width:none !important;}
	ul.sm-blue>li{float:none;}
	ul.sm-blue>li>a,ul.sm-blue ul.sm-nowrap>li>a{white-space:normal;}
	ul.sm-blue iframe{display:none;}
 
	
	/* Uncomment this rule to disable completely the sub menus for small screen devices */
	/*.sm-blue ul, .sm-blue span.sub-arrow, .sm-blue iframe {
		display:none !important;
	}*/


/* Menu box
===================*/

	.sm-blue {
		background:transparent;
	}
	.sm-blue ul {
		border:0;
		padding:0;
		background:#fff;
		 
	}
	.sm-blue ul ul {
		/* darken the background of the 2+ level sub menus and remove border rounding */
		background:rgba(100,100,100,0.1);
		 
	}


/* Menu items
===================*/

	.sm-blue a {
		padding:10px 5px 10px 28px !important; /* add some additional left padding to make room for the sub indicator */
		background:#888888 !important; /* Old browsers */
		 
		color:#fff !important;
	}
	.sm-blue ul a {
		background:transparent !important;
		color:#444444 !important;
		text-shadow:none !important;
	}
	.sm-blue a.current {
		color:#fff !important;
	}
	/* add some text indentation for the 2+ level sub menu items */
	.sm-blue ul a {
		border-left:8px solid transparent;
	}
	.sm-blue ul ul a {
		border-left:16px solid transparent;
	}
	.sm-blue ul ul ul a {
		border-left:24px solid transparent;
	}
	.sm-blue ul ul ul ul a {
		border-left:32px solid transparent;
	}
	.sm-blue ul ul ul ul ul a {
		border-left:40px solid transparent;
	}
	/* round the corners of the first and last items */
	.sm-blue > li:first-child > a {
		 
	}
	/* presume we have 4 levels max */
	.sm-blue > li:last-child > a,
	.sm-blue > li:last-child > ul > li:last-child > a,
	.sm-blue > li:last-child > ul > li:last-child > ul > li:last-child > a,
	.sm-blue > li:last-child > ul > li:last-child > ul > li:last-child > ul > li:last-child > a,
	.sm-blue > li:last-child > ul > li:last-child > ul > li:last-child > ul > li:last-child > ul > li:last-child > a,
	.sm-blue > li:last-child > ul,
	.sm-blue > li:last-child > ul > li:last-child > ul,
	.sm-blue > li:last-child > ul > li:last-child > ul > li:last-child > ul,
	.sm-blue > li:last-child > ul > li:last-child > ul > li:last-child > ul > li:last-child > ul,
	.sm-blue > li:last-child > ul > li:last-child > ul > li:last-child > ul > li:last-child > ul > li:last-child > ul {
		 
	}
	/* highlighted items, don't need rounding since their sub is open */
	.sm-blue > li:last-child > a.highlighted,
	.sm-blue > li:last-child > ul > li:last-child > a.highlighted,
	.sm-blue > li:last-child > ul > li:last-child > ul > li:last-child > a.highlighted,
	.sm-blue > li:last-child > ul > li:last-child > ul > li:last-child > ul > li:last-child > a.highlighted,
	.sm-blue > li:last-child > ul > li:last-child > ul > li:last-child > ul > li:last-child > ul > li:last-child > a.highlighted {
		 
	}


/* Sub menu indicators
===================*/
 

	.sm-blue a span.sub-arrow,
	.sm-blue ul a span.sub-arrow {
		top:50%;
		margin-top:-9px;
		right:auto;
		left:6px;
		margin-left:0;
		width:17px;
		height:17px;
		font:bold 16px/16px monospace !important;
		text-align:center;
		border:0;
		text-shadow:none;
		background:rgba(0,0,0,0.1);
		 
	}
	/* Hide sub indicator "+" when item is expanded - we enable the item link when it's expanded */
	.sm-blue a.highlighted span.sub-arrow {
		display:none !important;
        /*border-bottom:0px solid #ccc !important;*/
	}
	


/* Items separators
===================*/

	.sm-blue li {
		border-left:0;
	}
	.sm-blue ul li {
		border-top:1px solid rgba(0,0,0,0.05);
	}
	.sm-blue ul li:first-child {
		border-top:0;
	}

}


