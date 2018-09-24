﻿using System;
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

using Common;
using System.Threading;
using ServicesProxy;

namespace HomeMensajes
{
	
    public partial class Gadget_Mensajes : System.Web.UI.UserControl
    {
    	public RHPro.Lenguaje ObjLenguaje;
        protected void Page_Load(object sender, EventArgs e)
        {
		  	  			 
		   Repeater1.DataSource =  MenssageServiceProxy.Find(Utils.SessionBaseID, Utils.Lenguaje);              		  
		   Repeater1.Visible = true;
		   Repeater1.DataBind();
		   
        }
    }
}