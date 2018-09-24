<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="PopUpSearch.aspx.cs" Inherits="RHPro.PopUpSearch" meta:resourcekey="PageResource1" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Search Page</title>
    <style type="text/css">
        body
        {
            margin-left: 0px;
            margin-right: 0px;
            margin-top: 0px;
            margin-bottom: 0px;            
        }
        #logoSearch
        {
            height: 91px;
        }
        .resultados
        {
           margin-left: 0px;
           margin-top: 10px;
        }
        
        a:link {text-decoration: none}
           
    </style>
</head>
<body id="SearchPopUp" bgcolor="#dedede">
    <form id="form" runat="server">
    <asp:ScriptManager ID="scriptManager" runat="server">
        <Scripts>
            <asp:ScriptReference Path="~/Js/Utils.js" />
        </Scripts>
    </asp:ScriptManager>

    <div id="logoSearch" style="background-color: #FFFFFF">
        <asp:Image ID="Image1" runat="server" ImageUrl="~/Images/Logo.gif" />
    </div>
        
    <div style="font-family: Arial; font-size: small; background-color:Gray; color:White; font-weight:bold; margin-top:0px;" class="resultados">
        
        <br />
        &nbsp;&nbsp;
        <asp:Label  ID="lbtitulo" runat="server" meta:resourcekey="lbtituloResource1"></asp:Label><br />
        <br />
        &nbsp;&nbsp;
        <asp:Label ID="lbPagina" runat="server" meta:resourcekey="lbPaginaResource1"></asp:Label> 
        <br />
        <br />
        &nbsp;&nbsp;
        <asp:button id="cmdFirst" runat="server" 
                     OnClick="cmdFirst_Click" text=" << " 
            meta:resourcekey="cmdFirstResource1">
        </asp:button>
        &nbsp;
        <asp:button id="cmdPrev" runat="server"
                     OnClick="cmdPrev_Click" text="  <  " 
            meta:resourcekey="cmdPrevResource1">
        </asp:button>
        &nbsp;
        <asp:button id="cmdNext" runat="server" 
                     OnClick="cmdNext_Click" text="  >  " 
            meta:resourcekey="cmdNextResource1">
        </asp:button>            
        &nbsp;
        <asp:button id="cmdLast" runat="server" 
                     OnClick="cmdLast_Click" text=" >> " 
            meta:resourcekey="cmdLastResource1">
        </asp:button>
            
        <br />
        <br />
        
        <div id="ContenedorRepeater" style="background-color:#dedede; padding-top:5px;" >
        <asp:Repeater runat="server" ID="searchRepeater" OnItemDataBound="searchRepeater_ItemDataBound">
            <ItemTemplate> 
                <div id="modulo" style="color:Red; font-weight:bold; margin-left:11px; margin-top:10px; margin-bottom:5px;"> 
                    <asp:Label ID="lblModuloItem" runat="server" 
                        meta:resourcekey="lblModuloItemResource1"></asp:Label>
                </div>
                <div id="descripcion" style="background-color:#dedede; margin-left:20px; margin-top:10px">
                    <div id="link">
                        <asp:LinkButton ID="linkMenuItem" runat="server" CommandName="linkMenuItem" 
                            meta:resourcekey="linkMenuItemResource1"></asp:LinkButton></div>
                    <div id="text">
                        <asp:Label ID="lbDescripcion" runat="server" ForeColor="Black" Font-Bold="false"
                            meta:resourcekey="lbDescripcionResource1"></asp:Label></div>
                </div>
            </ItemTemplate>
        </asp:Repeater>
        </div>
    </div>
    </form>
</body>
</html>
