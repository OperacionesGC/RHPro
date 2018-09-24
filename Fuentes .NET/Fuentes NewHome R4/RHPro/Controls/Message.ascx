<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="Message.ascx.cs" Inherits="RHPro.Controls.Message" %>
<div id="superior">
                    </div>
<div id="medioTitulo"><asp:Literal ID="title" runat="server" 
        Text="<span>Comunidad</span><br /><span><b>RHPro</b></span>" 
        meta:resourcekey="titleResource1" />
    </div>
    <div id="medioCuerpo" >
        <asp:Repeater runat="server" ID="messageRepeater">
            <ItemTemplate>
                <div>
                    <label>
                        <%# Eval("Title") %>
                    </label>
                    <br />
                    <label class="descripcion">
                        <%# Eval("Body") %></label>
                </div>
            </ItemTemplate>
            <SeparatorTemplate>
            <div class="separador">
            </div>
        </SeparatorTemplate>
        </asp:Repeater>
    </div>
    <div id="inferior">
                    </div>