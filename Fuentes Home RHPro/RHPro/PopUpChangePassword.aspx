<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="PopUpChangePassword.aspx.cs"
    Inherits="RHPro.PopUp" StylesheetTheme="" meta:resourcekey="PageResource1" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>RHPro | Cambio de Contraseña</title>
    <style type="text/css">
         body
        {
            margin-top: 0px;
            margin-bottom: 0px;
        }
        .style6
        {
            width: 985px;
            height: 25px;
        }
        .style7
        {
            width: 269px;
            height: 25px;
        }
        .style11
        {
            width: 269px;
            height: 32px;
        }
        .style12
        {
            height: 32px;
        }
        .style13
        {
            height: 32px;
            width: 985px;
        }
        .style14
        {
            width: 985px;
        }
        .style15
        {
            height: 50px;
        }
        .style16
        {
            height: 28px;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server" style="font-family:Arial; color: #FFFFFF;" 
    defaultbutton="btnConfirmar">

    <div >
        &nbsp;<table style="margin: 0px; padding: 0px; width: 500px; height: 188px; " 
            cellpadding="0" cellspacing="0" bgcolor="#EE2E24"
            align="center">
            <tr>
                <td align="center" colspan="2" bgcolor="White">
                    <img alt="" src="Images/newPasswordSup.GIF" /></td>
            </tr>
            <tr>
                <td align="center" colspan="2" valign="middle">
                    <asp:Label ID="title" runat="server" Text="Usted debe cambiar su contraseña"
                    meta:resourcekey="lbPasswordResource1" Font-Bold="True" Font-Size="Medium" 
                        ForeColor="White"></asp:Label>
                </td>
            </tr>
            <tr>
                <td class="style6">
                </td>
                <td align="center" class="style7">
                    &nbsp;</td>
            </tr>
            <tr>
                <td align="right" class="style13" valign="middle">
                    <asp:Label ID="lbPassword" Text="Ingrese su anterior contraseña" runat="server" 
                    AssociatedControlID="txtOldPassword" meta:resourcekey="lbPasswordResource1" 
                        Font-Size="Medium" ForeColor="White" />
                &nbsp;&nbsp;&nbsp;
                </td>
                <td align="left" class="style12" valign="middle" >
                    <input type="password" id="txtOldPassword" runat="server" 
                        style="border:solid 1px black; color: #FF0000;" tabindex="0" /></td>
            </tr>
            <tr>
                <td align="right" class="style14" valign="middle">
                    <asp:Label ID="Label1" Text="Ingrese su nueva contraseña" runat="server" 
                    AssociatedControlID="txtOldPassword" meta:resourcekey="Label1Resource1" 
                        Font-Size="Medium" ForeColor="White" />
                &nbsp;&nbsp;&nbsp;
                </td>
                <td class="style11" valign="middle">
                    <input type="password" id="txtNewPassword" runat="server" 
                        style="border:solid 1px black; color: #FF0000;" tabindex="0"/></td>
            </tr>
            <tr>
                <td align="right" class="style14" valign="middle">
                <asp:Label ID="Label2" Text="Vuelva a ingresar su nueva contraseña" runat="server"
                    AssociatedControlID="txtOldPassword" meta:resourcekey="Label2Resource1" 
                        Font-Size="Medium" ForeColor="White" />
                &nbsp;&nbsp;&nbsp;
                </td>
                <td class="style11" valign="middle">
                <input type="password" id="txtVerifyPassword" runat="server" 
                        style="border:solid 1px black; color: #FF0000;" tabindex="0"/></td>
            </tr>
            <tr>
                <td class="style15" colspan="2" valign="middle" align="center">
     <asp:Label ID="errorMess" runat="server" 
            meta:resourcekey="errorMessResource1" Font-Bold="True" Font-Size="Medium" 
                        ForeColor="#FFFF66" Height="16px" />
                </td>
            </tr>
            <tr>
                <td align="center" class="style16" colspan="2">
            <asp:LinkButton ID="btnConfirmar" runat="server" 
                OnClick="btnConfirmar_Click" meta:resourcekey="btnConfirmarResource1" 
                        Font-Size="Small" ForeColor="White"></asp:LinkButton>
            <asp:LinkButton runat="server" CausesValidation="False" ID="btnCancel" 
                style="margin-left:20px;"   
                OnClientClick="javascript:window.close();" 
                meta:resourcekey="btnCancelResource1" Font-Size="Small" ForeColor="White"></asp:LinkButton>
                </td>
            </tr>
            <tr>
                <td align="center" colspan="2" bgcolor="White">
                    <img alt="" src="Images/newPasswordInf.GIF" /></td>
            </tr>
            </table>
    
    
    </div>
    
    </form>

</body>
</html>
