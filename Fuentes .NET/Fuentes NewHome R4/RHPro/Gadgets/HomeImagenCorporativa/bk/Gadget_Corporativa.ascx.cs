using System;
using System.Collections;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Xml.Linq;

namespace HomeGadget_Corporativa
{	
    public partial class Gadget_Corporativa : UserControl
    {
    	public RHPro.Lenguaje ObjLenguaje;
        protected void Page_Load(object sender, EventArgs e)
        {          
			//ImagenCorpPais.Controls.Add(new LiteralControl("<img src='Gadgets/HomeImagenCorporativa/img/Corporativas/corp_esPE.png' width='100%' >"));
			//ImagenCorpPais.Controls.Add(new LiteralControl("<img src='Gadgets/HomeImagenCorporativa/img/BannerCorporativo3.png' width='100%' >"));
			//ImagenCorpPais.Controls.Add(new LiteralControl("<img src='Gadgets/HomeImagenCorporativa/img/Corporativas/corp_esPE.png' width='100%' >"));
			ImagenCorpPais.Controls.Add(new LiteralControl("<img src='Gadgets/HomeImagenCorporativa/img/BannerCorporativo4.png' width='100%' >"));
			ImagenCorpPais.DataBind();
          
      
        }
    }
}