<%@ Page Title="" Language="C#" MasterPageFile="~/MasterPage.Master" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="RHPro.Default" EnableViewState="true"  %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="content" runat="server">


    <div id="contenido">
        <div id="Izquierda">
            <div id="Menu1Izq">
            
                <cc:MRU id="mruMain" runat="server" visible="true">
                </cc:MRU>
            </div>
            <div id="Menu2Izq">
                <cc:Banner id="bannerMain" runat="server" ></cc:Banner>
            </div>
        </div>
        <div id="centroDerecha">
            <div id="Centro">
                <div id="centralArriba">
                    <cc:Modules Id="mlsMain" runat="server">
                    </cc:Modules>
                </div>
                <cc:FooterPage id="cFooterPage" runat="server"></cc:FooterPage>
                
                
            </div>
            <div id="Derecha">
                <div id="login">
                    <cc:CustomLogin id="cLogin" runat="server">
                    </cc:CustomLogin>
                </div>
                <div id="Menu1Der">
                    <cc:Message id="messageMain" runat="server">
                    </cc:Message>
                </div>
                <div id="Menu2Der">
                    <cc:Link id="linksMain" runat="server">
                    </cc:Link>
                </div>
            </div>
        </div>
        <div class="clear">
        </div>
        </div>
        
    



</asp:Content>
