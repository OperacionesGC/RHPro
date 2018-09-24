<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="Link.ascx.cs" Inherits="RHPro.Controls.Link" %>
 <div id="superior">
                    </div>
<div id="medioTitulo">
<asp:Literal ID="title" runat="server" Text="<span>Links de interes</span>" 
        meta:resourcekey="titleResource1" />
            </div>
            <div id="medioCuerpo" >
                <asp:Repeater runat="server" ID="linkRepeater">
                    <HeaderTemplate>
                        <div style="height: 10px;">
                        </div>
                    </HeaderTemplate>
                    <ItemTemplate>
                        <div>
                            <label>
                                <%# Eval ("Title") %></label>
                            <br />
                            <a class="descripcion" target="_blank" 
                                <%# Eval("Url","href='http://{0}'") %>="">
                                <%# Eval ("Url") %></a>
                        </div>
                    </ItemTemplate>
                </asp:Repeater>
            </div>
            <div id="inferior">
                    </div>