<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="PrintPreview.aspx.cs" Inherits="OD.Web.PrintPreview" %>

<%@ Register assembly="Microsoft.ReportViewer.WebForms, Version=9.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a" namespace="Microsoft.Reporting.WebForms" tagprefix="rsweb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>RHPro X2 - Organigrama Dinámico (Vista previa de impresión)</title>
    <script type="text/javascript">
        function postBackPrint() {
            <%= Page.ClientScript.GetPostBackEventReference(this, "") %>
        }
    </script>
</head>
<body>
    <form id="form1" runat="server">
    <asp:HiddenField ID="txtPNGBytes" runat="server" />
    <asp:ScriptManager ID="ScriptManager1" runat="server" /> 
    <div>
    
    </div>
    <rsweb:ReportViewer ID="ReportViewer1" runat="server" Font-Names="Verdana" 
        Font-Size="8pt" Height="600px" Width="900px">
        <LocalReport ReportPath="OrganigramaDinamico.rdlc">
            <DataSources>
                <rsweb:ReportDataSource DataSourceId="ObjectDataSource1" 
                    Name="ReportDataSource_ImageData" />
            </DataSources>
        </LocalReport>
    </rsweb:ReportViewer>
    <asp:ObjectDataSource ID="ObjectDataSource1" runat="server" 
        SelectMethod="GetData" 
        TypeName="OD.Web.ReportDataSourceTableAdapters.">
    </asp:ObjectDataSource>
    </form>
</body>
</html>
