<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="PopUpPolitics.aspx.cs" Inherits="RHPro.PopUpPolitics"
StylesheetTheme="Default" meta:resourcekey="PageResource1" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title></title>
    <style>
         .TextoPoliticas 
         {
         	margin:15px 0px 0px 15px; font-family:Arial; font-size:10pt;
         	color:#666666;
         	line-height:1.5;
         	
         }
    </style>
</head>
<body>
    <form id="form1" runat="server">
    <div class="TextoPoliticas">
      <span style="color:#333333;">
        <% Imprimir_Politica(); %>
     
      </span>
    </div>
    </form>
</body>
</html>
