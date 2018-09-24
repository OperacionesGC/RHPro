<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="Modules.ascx.cs" Inherits="RHPro.Controls.Modules" %>
<input type="hidden" id="divScroll_scrollValue" runat="server" />
<input type="hidden" id="linkSelected" runat="server" />
<div id="superior">
</div>
<div id="titulo">
<asp:Literal ID="title" Text="Modulos <b>RHPro X2</b>" runat="server" 
        meta:resourcekey="titleResource1"  />
</div>
<div id="medio" style="height: 332px">
    <asp:UpdatePanel  ID="updatePnlModules" runat="server" UpdateMode="Conditional">
        <ContentTemplate>
            <div id="divScroll" class="divScroll" runat="server">
                <asp:Repeater ID="rprModules" runat="server" OnItemCommand="rprModules_ItemCommand"
                    OnItemDataBound="rprModules_ItemDataBound">
                    <ItemTemplate>
                        <div class="CentroArribaIzq">
                           <table>
                            <tr>
                            <td align="left" style="width:170px;">
                            <asp:LinkButton ID="btnLink" runat="server" CommandName="btnLink" 
                                    ><%# Eval("MenuTitle") %></asp:LinkButton>
                            </td>
                            <td align="right" style="width: 25px;"><asp:ImageButton  ID="btnManual" 
                                    runat="server" ImageUrl="~/Images/document.png" CommandName="btnManual"
                                Visible="False" meta:resourcekey="btnManualResource1" /></td>
                                <td align="right" style="width: 25px;">
                            <asp:ImageButton ID="btnDVD" runat="server" ImageUrl="~/Images/new_video.png" CommandName="btnDVD"
                                Visible="False" meta:resourcekey="btnDVDResource1" /></td>
                                </tr>
                                </table>                            
                        </div>
                    </ItemTemplate>
                </asp:Repeater>
            </div>
            <div style="float: right; width: 320px; margin-right: 0px; height: 320px">
                <div class="CentroArribaDerTitulo">
                    <b>
                        <label id="lblModuleTitle"  runat="server">
                        </label>
                        <asp:LinkButton ID="lklModuleTitle" runat="server" 
                        OnClientClick="onLnkModuleLinkClick()" 
                        meta:resourcekey="lnkModuleLinkResource1"></asp:LinkButton>
                    </b>
                </div>
                <div class="CentroArribaDerDescripcion">
                    <label id="lblModuleDescription" runat="server">
                    </label>
                </div>
                <div class="CentroArribaDerInfo">
                    <b>
                        <asp:LinkButton ID="lnkModuleLink" runat="server" 
                        OnClientClick="onLnkModuleLinkClick()" 
                        meta:resourcekey="lnkModuleLinkResource1"></asp:LinkButton></b>
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</div>
<div id="inferior">
</div>
