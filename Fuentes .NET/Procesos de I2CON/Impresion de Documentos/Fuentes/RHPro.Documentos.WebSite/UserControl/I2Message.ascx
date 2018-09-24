<%@ control language="C#" autoeventwireup="true" inherits="I2Message, App_Web_nmvvfdyy" %>

<table cellpadding="0" cellspacing="0" width="100%" style="border-bottom: #08225a 6px solid">
    <tr>
        <td valign="top" style="width: 37px; padding-right: 5px; padding-left: 5px; padding-bottom: 5px; padding-top: 5px;">
            <asp:Image ID="ImageError" runat="server" ImageUrl="~/Images/error.gif" Visible="False" /><asp:Image ID="ImageInfo" runat="server" ImageUrl="~/Images/info.gif" /><asp:Image ID="ImageWarning" runat="server" ImageUrl="~/Images/warning.gif" Visible="False" /></td>
        <td valign="top" style="padding-right: 5px; padding-left: 5px; border-left-width: 1px; border-left-color: firebrick; padding-bottom: 5px; padding-top: 5px;">
            <asp:Label ID="LabelMessage" runat="server" Text="Mensaje para usuario." Font-Names="Verdana" Font-Size="9pt" ForeColor="Gray"></asp:Label>
        </td>
    </tr>
</table>
