using System;
using System.Collections.Generic;
using System.Threading;
using System.Web.UI;
using System.Web.UI.WebControls;
using Common;
using Entities;
using ServicesProxy;
using System.IO;
using System.Web;


namespace RHPro.Controls
{
    public partial class Modules : UserControl
    {
        /// <summary>
        /// Modulos disponibles
        /// </summary>
        protected List<Module> AvailableModules
        {
            get { return ViewState["AvailableModules"] as List<Module>; }
            set { ViewState["AvailableModules"] = value; }
        }

        private int RprModulesItemIndexSelected
        {
            get { return int.Parse(ViewState["RprModulesItemIndexSelected"].ToString()); }
            set { ViewState["RprModulesItemIndexSelected"] = value.ToString(); }
        }

        public string BaseID
        {
            get
            {
                if(Session["BaseModule"]!=null)
            {
                return Session["BaseModule"].ToString();
            }
                return "";
            }
            set { Session["BaseModule"] = value; }
        }

        protected void Page_Load(object sender, EventArgs e)
        {
            
            if (!IsPostBack)
            {
                LoadModule();    
                divScroll.Attributes.Add("onscroll", string.Format("saveScrollPosition('{0}','{1}')", divScroll.UniqueID.Replace("$", "_"), divScroll_scrollValue.UniqueID.Replace("$", "_")));
            }

            if (divScroll_scrollValue.Value.Length > 0)
            {
                ScriptManager.RegisterStartupScript(Page, GetType(), "SetScroll", string.Format("setScrollPosition('{0}','{1}');", divScroll.UniqueID.Replace("$", "_"), divScroll_scrollValue.UniqueID.Replace("$", "_")), true);
            }
        }
        protected void Page_PreRender(object sender, EventArgs e)
        {
            if (Utils.SessionBaseID!=BaseID)
            {
                BaseID = Utils.SessionBaseID;
                LoadModule();    
            }
            
        }

        public  void LoadModule()
        {
            AvailableModules = ModuleServiceProxy.Find(Utils.SessionUserName, Utils.SessionBaseID, Thread.CurrentThread.CurrentCulture.Name);

            RprModulesItemIndexSelected = 0;
            ShowModule(AvailableModules[0]);
            rprModules.DataSource = AvailableModules;
            rprModules.DataBind();

        }

        protected void rprModules_ItemDataBound(object sender, RepeaterItemEventArgs e)
        {
            LinkButton linkButton = ((LinkButton)e.Item.FindControl("btnLink"));

            if (e.Item.ItemIndex == RprModulesItemIndexSelected)
            {
                SetItemSelected(e.Item);
            }

            linkButton.CommandArgument = ((Module)e.Item.DataItem).Id.ToString();

            if (!String.IsNullOrEmpty(((Module)e.Item.DataItem).LinkManual))
            {
                ImageButton btnManual = (ImageButton)e.Item.FindControl("btnManual");
                btnManual.CommandArgument = "~/../" + ((Module)e.Item.DataItem).LinkManual;
                btnManual.Visible = true;
            }

            if (!String.IsNullOrEmpty(((Module)e.Item.DataItem).LinkDvd))
            {
                ImageButton btnDVD = (ImageButton)e.Item.FindControl("btnDVD");
                btnDVD.Attributes.Add("OnClick", String.Format("javascript:popvideo('{0}','{1}','{2}'); return false;", ((Module)e.Item.DataItem).LinkDvd, ((Module)e.Item.DataItem).MenuTitle, e.Item.ClientID)); 
                btnDVD.Visible = true;
            }
        }


        protected void rprModules_ItemCommand(object source, RepeaterCommandEventArgs e)
        {
            if (e.CommandName == "btnManual" || e.CommandName == "btnDVD")
            {
                Utils.Redirect(e.CommandArgument.ToString(), "_blank", String.Empty);
            }

            SetItemUnSelected(rprModules.Items[RprModulesItemIndexSelected]);
            SetItemSelected(e.Item);

            Module selectedModule = AvailableModules.Find(M => M.Id.ToString().Equals(e.CommandArgument));
            ShowModule(selectedModule);
        }

        private void SetItemSelected(RepeaterItem item)
        {
            LinkButton linkButtonSelected = ((LinkButton)item.FindControl("btnLink"));
            linkButtonSelected.Style.Add(HtmlTextWriterStyle.Color, "blue");
            RprModulesItemIndexSelected = item.ItemIndex;
        }

        private void SetItemUnSelected(RepeaterItem item)
        {
            LinkButton linkButtonSelected = ((LinkButton)item.FindControl("btnLink"));
            linkButtonSelected.Style.Add(HtmlTextWriterStyle.Color, "red");
        }

        private void ShowModule(Module module)
        {
            if (module != null)
            {
                lklModuleTitle.Text= module.MenuTitle;
                lblModuleTitle.InnerText = module.MenuTitle;
                if (lklModuleTitle.Text.Length > 31)
                {
                    lklModuleTitle.Text = lklModuleTitle.Text.Substring(0, 28) + "...";
                    lblModuleTitle.InnerText = lblModuleTitle.InnerText.Substring(0, 28) + "...";
                }
                
                lblModuleDescription.InnerText = module.MenuDetail;
                if (!string.IsNullOrEmpty(module.Action))
                {
                    lnkModuleLink.Visible = true;
                    lblModuleTitle.Visible = false;
                    lklModuleTitle.Visible = true;
                    //lnkModuleLink.Text = "+ Ingresar";
                    ScriptManager.RegisterStartupScript(Page, GetType(), "LnkModuleLinkClick", string.Concat("function onLnkModuleLinkClick(){", string.Format("{0}", module.Action.ToLower().Replace("javascript:", string.Empty)), "}"), true);
                }
                else
                {
                    lblModuleTitle.Visible = true;
                    lklModuleTitle.Visible = false;
                    lnkModuleLink.Visible = false;
                }
            }
        }
    }
}
