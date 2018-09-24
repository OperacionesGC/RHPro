<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="MRU.ascx.cs" Inherits="RHPro.Controls.MRU"     %>
<div id="mruImage" class="mruImage"  runat="server"  ></div>
<div id="mruCompleto" runat="server" >
<div id="superior">
</div>

<div id="titulo">
<asp:Literal ID="title" runat="server"  
        Text="<span>Menues mas</span><br /><span>usados</span>" 
        meta:resourcekey="titleResource1" />
</div>
<div id="cuerpo" class="MruCuerpo">
    <asp:Repeater runat="server" ID="MRURepeater">
        <ItemTemplate>
            <div>
                <label>
                    <%# Eval("MenuName") %>
                </label>
                <br />
                <label class="descripcion">
                    <%# Eval("Root") %>
                </label>
                <br />
                <a onclick="<%# Eval("Action") %>" style="cursor: pointer;">>></a>
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
</div>