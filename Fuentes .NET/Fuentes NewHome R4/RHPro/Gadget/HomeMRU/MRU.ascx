<%@ Control Language="C#" AutoEventWireup="true" CodeFile="MRUmi.ascx.cs" Inherits="HomeMRU.MRUmi"   %>

 
 
 <style>
 
 .Divisor_MRUS { 
 /*
    columns:3 !important;
    -webkit-columns:100px 3; 
    -moz-columns:100px 3;
	padding-top:18px !important;
	*/
 -moz-column-count: 3;
    -moz-column-gap: 20px;
    -ms-column-count: 3;
    -ms-column-gap: 20px;
    -webkit-column-count: 3;
    -webkit-column-gap: 20px;
    column-count: 3;
 column-gap:0px; 
 
 
	
}

 .MRUGeneral_Link{
 color:#666;text-align:left; height:26px; margin-left:0; margin-top:0px; margin-right:0; 
 font-size:9pt; cursor:pointer !important; overflow:hidden !important;} 
 
 .MRUGeneral_Link:hover{ color:#000000; background-color:#E0E0E0 !important; }
 .IconoModuloMRUModulo
{
	 width:19px !important; height:24px  !important; 
	 vertical-align:middle; text-align:center; 	 
	 cursor:pointer !important;  
	 text-align:center;
	 border:0;
     padding:5px; padding-left:0px; margin-left:2px; margin-right:0px;     
}

 .IconoModuloMRUPpal_PNG	 
	 {
	 width:25px !important;  
	 vertical-align:middle; text-align:center; 	 
	 cursor:pointer !important;  
	 text-align:center;
	 border:0;
     padding:0px; padding-left:0px; margin-left:4px; margin-right:0px;     
	 }
	 
 </style>

	<div id="mruCompleto" runat="server" class="Divisor_MRUS" >
		<asp:Repeater runat="server" ID="MRURepeater"     >
			<ItemTemplate> 
 
<%#Armar_Link_MRU((string)Eval("menuaccess"),(int)Eval("menumsnro"),(string)Eval("action"),(int)Eval("menuraiz"),(int)Eval("menunro"),(string)Eval("menuname"),(int)Eval("mrucant"),(string)Eval("menudir")  ) %>

				
			</ItemTemplate>			
		</asp:Repeater>
	</div>
