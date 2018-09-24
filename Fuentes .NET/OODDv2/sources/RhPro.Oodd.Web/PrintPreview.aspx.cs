using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Microsoft.Reporting.WebForms;
using System.Data;
using OD4.Web;

namespace OD.Web
{
    public partial class PrintPreview : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!Page.IsPostBack)
                FillReport();
        }

        private void FillReport()
        {
            string bytes64 = Request["txtPNGBytes"];
            byte[] imageBytes = System.Convert.FromBase64String(bytes64);

            DSReportPrintImage ds = new DSReportPrintImage();
            DataRow drImage = ds.Tables[0].NewRow();
            drImage["ImageBytes"] = imageBytes;
            ds.Tables[0].Rows.Add(drImage);

            ReportViewer1.LocalReport.ReportPath = "OrganigramaDinamico.rdlc";

            ReportDataSource src = new ReportDataSource("DSReportPrintImage_ImageData", ds.Tables[0]);
            ReportViewer1.LocalReport.DataSources.Add(src);

            ReportViewer1.SizeToReportContent = true;

            //ReportViewer1.ShowBackButton = true;
            //ReportViewer1.ShowDocumentMapButton = true;
            ReportViewer1.ShowPageNavigationControls = false;
            ReportViewer1.ShowParameterPrompts = false;
            ReportViewer1.ShowPrintButton = true;
            //ReportViewer1.ShowPromptAreaButton = true;
            //ReportViewer1.ShowReportBody = true;
            ReportViewer1.ShowZoomControl = true;
            //ReportViewer1.ShowRefreshButton = false

            ReportViewer1.LocalReport.Refresh();
            
        }
    }
}
