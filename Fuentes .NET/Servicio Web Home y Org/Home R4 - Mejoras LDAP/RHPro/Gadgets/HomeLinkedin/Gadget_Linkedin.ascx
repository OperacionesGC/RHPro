<%@ Control Language="C#" AutoEventWireup="true" ClassName="Gadget_Linkedin"  CodeBehind="Gadget_Linkedin.ascx.cs"   %>
 
 
   <% 
   RHPro.Lenguaje ObjLenguaje = new RHPro.Lenguaje();
   RHPro.Gadget G = new RHPro.Gadget();    
   %>
      
   <%  Response.Write(G.TopeModulo(ObjLenguaje.Label_Home("Empleados RHPro"),"680"));  %>
     <script src="//platform.linkedin.com/in.js" type="text/javascript"></script>
<script type="IN/CompanyInsider" data-id="121732" data-modules="innetwork,newhires,jobchanges"></script>
   <%  Response.Write(G.PisoModulo());  %>