<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="Banner.ascx.cs" Inherits="RHPro.Controls.Banner" %>
<img id="img"  class="imgBanner" />        
<script type="text/javascript">
 
    setInterval(ChangeImage, <%= int.Parse(ConfigurationManager.AppSettings["Time"].ToString()) * 1000 %>)
    currentImage = 0;
    setImage(images[currentImage]);

 
    function ChangeImage() {
        currentImage++;
           
        if (images.length <= currentImage) {
            currentImage = 0
        }
        
        setImage(images[currentImage]);
    }
    
    function setImage(imageURL)
    {
        document.getElementById("img").src = '<%= ConfigurationManager.AppSettings["RootImagenes"] %>' + imageURL ;
    }
    
</script>
