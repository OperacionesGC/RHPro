<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="CustomLogin.ascx.cs"
    EnableViewState="true" Inherits="RHPro.Controls.CustomLogin" %>

<script type="text/javascript">
    /*$(document).ready(function() {
        $('#ctl00_content_<%= this.ID %>_cmbDatabase').sSelect();
    });*/




    //Abre popUp de Politics
    function Politic_show() {
        window.open("PopUpPolitics.aspx", "Ventana", 'height=215,width=450,status=yes,toolbar=no,menubar=no,location=no,resizable=yes,scrollbars=yes,left=5,top=5');
    }
</script>

<div id="superior">
</div>
<div id="medio">
    <div id="LoginON" runat="server">
        <label>
            <asp:Literal ID="title" runat="server" Text="Login" meta:resourcekey="titleResource1" /></label>
        <asp:Panel ID="panelLogin" runat="server" DefaultButton="btnLogin" 
            meta:resourcekey="panelLoginResource1" Font-Size="XX-Small">
            <asp:TextBox runat="server" Style="display: none;" meta:resourcekey="TextBoxResource1"></asp:TextBox>
            &nbsp;&nbsp;<input ID="txtUserName" runat="server" class="borde" style="width: 150px;
                margin: 5px 0px 5px 0px;" type="text" /><br />
            &nbsp;
            <input id="txtPassword" runat="server" type="password" style="width: 150px;
                margin: 0px 0px 5px 0px;" class="borde" /><br />
            &nbsp;
            <asp:DropDownList ID="cmbDatabase" runat="server" AutoPostBack="True" 
                    EnableTheming="True" meta:resourcekey="cmbDatabaseResource1" name="cmbDatabase" 
                    Width="154px" CssClass="borde" 
                    OnSelectedIndexChanged="cmbDatabase_SelectedIndexChanged">
                    </asp:DropDownList>        
                    <asp:Panel ID="PanellstDatabase" runat="server" Height="55px" Width="165px" 
                        Font-Size="XX-Small" DefaultButton="btnChangeDB">
                        &nbsp;&nbsp;<asp:ListBox ID="lstDatabase" runat="server" CssClass="borde" Height="55px" 
                            OnSelectedIndexChanged="lstDatabase_SelectedIndexChanged" Width="154px">
                        </asp:ListBox>                
                    </asp:Panel>
                
                <span style="font-size: 10px;">
            &nbsp;&nbsp;<asp:LinkButton ID="btnLogin" runat="server" Font-Size="XX-Small" 
                    meta:resourcekey="sendResource1" OnClick="doLogin_Click" />
                </span><span style="padding-left: 0px;">
                <asp:LinkButton ID="btnClean" runat="server" Font-Size="XX-Small" 
                    meta:resourcekey="cleenResource1" Style="padding-left: 0px;" />
                </span><span style="font-size: 10px; padding-left: 0px;">
                <asp:LinkButton ID="btPolitics" runat="server" Font-Size="XX-Small" 
                    meta:resourcekey="Literal2Resource1" 
                    OnClientClick="Politic_show();return false;" CssClass="btnLoginON" />
                <span style="padding-left: 0px;">
            <asp:LinkButton ID="btnChangeDB" runat="server" Font-Size="XX-Small" 
                Style="padding-left: 0px;" 
                OnClick="doChangeDB_Click"/>
            </span>
                </span>
        </asp:Panel>
        <input type="text" style="display: none" />
    </div>
    <div id="LoginOFF" runat="server" style="display: block;">
        <label><asp:Literal ID="login" runat="server" Text="Bienvenido" meta:resourcekey="loginResource1" /></label>
        <br />
        <b>
            <label id="lblUser" style="color: #abc;" runat="server">
            </label>
        <br />
        <br />
        &nbsp;
        <asp:Label ID="LabelBaseSeleccionada" runat="server" Font-Size="Small"></asp:Label>
        </b>
        <br />
        <div id="linkLogOff">
            <span style="font-size: 10px;">
                &nbsp;
                <asp:LinkButton ID="btnLogOut" OnClick="btnLogOut_Click" runat="server"
                    meta:resourcekey="logoutResource1" Font-Size="XX-Small" />
            </span>
        </div>
    </div>
</div>
<div id="inferior">
</div>

