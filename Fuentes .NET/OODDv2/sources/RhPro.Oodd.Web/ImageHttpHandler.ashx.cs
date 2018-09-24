using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Services;
using System.Drawing;
using System.Drawing.Imaging;
using System.ComponentModel;
using System.IO;
using System.Windows.Media.Imaging;
using System.Net;

namespace OD.Web
{
    /// <summary>
    /// Fetch an Image and convert it to jpg format
    /// </summary>
    [WebService(Namespace = "http://tempuri.org/")]
    [WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
    public class ImageHttpHandler : IHttpHandler
    {

        public void ProcessRequest(HttpContext context)
        {
            Uri uri = context.Request.Url;

            string imageName = "";

            if (!string.IsNullOrEmpty(context.Request.QueryString["imageName"]))
            {
                imageName = context.Request.QueryString["imageName"];
            }
            else
            {
                throw new ArgumentException("No parameter specified");
            }

            // Fetch the image 
            Bitmap newBmp = InitializeStreamBitmap(uri, imageName);
            
            if (newBmp != null)
            {
                newBmp.Save(context.Response.OutputStream, ImageFormat.Jpeg);
                newBmp.Dispose();
            }

        }

        private Bitmap InitializeStreamBitmap(Uri uri, string imageName)
        {
            try
            {
                
                string imagePath = System.Web.Configuration.WebConfigurationManager.AppSettings["imagePath"];
                
                if (!imagePath.Substring(imagePath.Length - 1, 1).Equals("/"))
                {
                    imagePath += "/";
                }

                Uri imageUri = new Uri(uri, imagePath + imageName);
                WebRequest request = WebRequest.Create(imageUri);
                WebResponse response = request.GetResponse();
                Stream responseStream = response.GetResponseStream();
                Bitmap bitmap = new Bitmap(responseStream);
                return bitmap;
            }
            catch (WebException)
            {
                return null;
            }
        }

        public bool IsReusable
        {
            get
            {
                return false;
            }
        }
    }
}
