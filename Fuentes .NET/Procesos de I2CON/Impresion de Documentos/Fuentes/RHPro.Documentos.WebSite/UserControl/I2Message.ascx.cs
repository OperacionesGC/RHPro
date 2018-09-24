using System;
using System.Data;
using System.Configuration;
using System.Collections;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;

public partial class I2Message : System.Web.UI.UserControl
{
    public enum I2MessageType {Information, Warning, Error};
    private I2MessageType _type;
    
    protected void Page_Load(object sender, EventArgs e)
    {
    }

    public I2MessageType Type 
    {
        get { return _type; }
        set
        {
            _type = value;
            ImageError.Visible = (_type == I2MessageType.Error);
            ImageInfo.Visible = (_type== I2MessageType.Information);
            ImageWarning.Visible = (_type==I2MessageType.Warning);
        }
    }

    public string Text
    {
        get { return LabelMessage.Text; }
        set {LabelMessage.Text = (value.Trim() == "" ? "Mensaje para usuario." : value); }
    }

    public void Show(I2MessageType type, string message)
    {
        Type = type;
        Text = message;
    }
}
