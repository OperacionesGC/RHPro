<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="Footer.ascx.cs" Inherits="RHPro.Controls.Footer" %>
<div id="superior">
</div>
<div id="medio">
    <div style="float: left;" class="titulo">
        <asp:literal ID="title" runat="server"  meta:resourcekey="titleResource1"
             /></div>
    <div style="float: right;" class="version">
        <asp:literal ID="ltVersion" runat="server"  Text="Version:" meta:resourcekey="ltVersionResource1" 
             /><label runat="server" id="version" />
        &nbsp 
        <asp:literal ID="ltPatch" runat="server"  Text="Patch:" meta:resourcekey="ltPatchResource1" 
            /><label runat="server" id="patch" /></div>
</div>
<div id="inferior">
</div>
