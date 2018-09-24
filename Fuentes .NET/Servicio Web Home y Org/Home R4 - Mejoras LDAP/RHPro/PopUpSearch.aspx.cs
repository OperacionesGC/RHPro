using System;
using System.Data;
using System.Text;
using System.Threading;
using System.Web.UI;
using Entities;
using ServicesProxy;
using System.Collections.Generic;
using System.Web.UI.WebControls;

namespace RHPro
{
    public partial class PopUpSearch : Page
    {
        #region Properties

        private string module = "@@@"; //pongo este strint para poder agrupar por OTROS
        public bool showModule = true; 

        public PopUpSearchData PopUpSearchData
        {
            get { return ViewState["PopUpSearchData"] as PopUpSearchData; }
            set { ViewState["PopUpSearchData"] = value; }
        }

        public int CurrentPage
        {
            get
            {
                // look for current page in ViewState
                object o = this.ViewState["_CurrentPage"];
                if (o == null)
                    return 0; // default page index of 0
                else
                    return (int) o;
            }

            set { this.ViewState["_CurrentPage"] = value; }
        }

        public List<Search> listSearch
        {
            get
            {
                object o = this.ViewState["_listSearch"];
                if (o == null)
                    return null;
                else
                    return (List<Search>) o;
            }

            set { this.ViewState["_listSearch"] = value; }
        }

        #endregion

        #region Page Handles

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                PopUpSearchData = Session["PopUpSearchData"] as PopUpSearchData;
                Session.Remove("PopUpSearchData");

                if (PopUpSearchData != null)
                {
                    listSearch = SearchServiceProxy.Find(PopUpSearchData.UserName, PopUpSearchData.WordToFind,
                                                         PopUpSearchData.DataBase, Thread.CurrentThread.CurrentCulture.Name);

                    ShowInfo();

                    StringBuilder titulo = new StringBuilder();


                    titulo.Append(GetLocalResourceObject("Find1.Text").ToString());
                    titulo.Append(listSearch.Count);
                    titulo.Append(GetLocalResourceObject("Find2.Text").ToString());
                            
                    titulo.Append(PopUpSearchData.WordToFind);
                    titulo.Append("'");
                    lbtitulo.Text = titulo.ToString();
                }
                else
                {
                    Page.ClientScript.RegisterStartupScript(GetType(), "Mensaje", String.Format("javascript:alert('{0}');", GetLocalResourceObject("FormaIncorrecta.Text").ToString()), true);
                    Page.ClientScript.RegisterStartupScript(GetType(), "cerrarPopUp", "javascript:window.close()", true);
                }
            }
        }

        protected void cmdFirst_Click(object sender, EventArgs e)
        {
            CurrentPage = 0;
            ShowInfo();
        }

        protected void cmdPrev_Click(object sender, EventArgs e)
        {
            CurrentPage -= 1;
            ShowInfo();
        }

        protected void cmdNext_Click(object sender, EventArgs e)
        {
            CurrentPage += 1;
            ShowInfo();
        }

        protected void cmdLast_Click(object sender, EventArgs e)
        {
            CurrentPage = 9999999;
            ShowInfo();
        }

        private void ShowInfo()
        {
            PagedDataSource pdsPaginado = new PagedDataSource();

            pdsPaginado.DataSource = listSearch;
            pdsPaginado.AllowPaging = true;

            pdsPaginado.PageSize = 6;

            if (CurrentPage > pdsPaginado.PageCount-1)
                CurrentPage = pdsPaginado.PageCount-1;

            pdsPaginado.CurrentPageIndex = CurrentPage;

            searchRepeater.DataSource = pdsPaginado;

            searchRepeater.DataBind();

            StringBuilder pagina = new StringBuilder();

            pagina.Append(GetLocalResourceObject("Page1.Text").ToString());
            pagina.Append(CurrentPage + 1);
            pagina.Append(GetLocalResourceObject("Page2.Text").ToString());                   
    
            pagina.Append(pdsPaginado.PageCount);

            lbPagina.Text = pagina.ToString();

            cmdFirst.Enabled = !pdsPaginado.IsFirstPage;
            cmdPrev.Enabled = !pdsPaginado.IsFirstPage;
            cmdNext.Enabled = !pdsPaginado.IsLastPage;
            cmdLast.Enabled = !pdsPaginado.IsLastPage;
        }

        #endregion

        #region Controls Handles

        protected void searchRepeater_ItemDataBound(object sender, RepeaterItemEventArgs e)
        {            
            if (e.Item.DataItem != null)
            {
                Search search = (Search) e.Item.DataItem;

                if (showModule)
                {
                    if (module != search.Module)
                    {
                        Label lblModuloItem = (Label) e.Item.FindControl("lblModuloItem");
                        lblModuloItem.Text = search.Module != "" ? search.Module : "Otros";
                        module = search.Module;
                    }
                }

                LinkButton linkMenuItem = (LinkButton)e.Item.FindControl("linkMenuItem");
                Label description = (Label)e.Item.FindControl("lbDescripcion");
                linkMenuItem.Text = search.MenuDescription;
                description.Text = search.Description;

                linkMenuItem.Attributes.Add("OnClick",
                                            String.Format("{0}; return false;",
                                                          search.Action.Replace("javascript:", string.Empty))); 
                
              
                showModule = !showModule;                    
            }            
        }

        #endregion
    }
}