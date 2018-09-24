<%@ Page Language="C#" ContentType="text/html" ResponseEncoding="utf-8" CodeFile="CopyAspSessionToAspNetSession.aspx.cs" Inherits="CopyAspSessionToAspNetSession" %>
<%
//'Descripción	: CopyAspSessionToAspNetSession.aspx
//'Fecha		: 07/02/2013
//'Autor		: Lisandro Moro
//'Modificado	: 28/02/2013 - Lisandro Moro - Agrego pasaje de variables de asp a asp.net
//'Modificado	: 27/06/2014 - Gonzalez Nicolás - Se modifco el charset a ANSI
//'Modificado	: 22/08/2014 - Lisandro Moro - Se agrego codefile

string c_seed = "56238";
string abandon = Encriptar.Encrypt(c_seed,"Abandon");
string uno = Encriptar.Encrypt(c_seed,"1");
string returnURL = Encriptar.Encrypt(c_seed,"returnURL");
string PASSWORD = Encriptar.Encrypt(c_seed,"Password");
if (Request.Form[abandon] == uno){
	Session.Abandon();
}else{
	foreach (string key in Request.Form.Keys){
	//left(ucase(Trim(param(0))),9) <> Ucase("RHpro_log")
	  
		if (key == PASSWORD){
			Session[Encriptar.Decrypt(c_seed,key.ToString())] = Request.Form[key].ToString();
		}else{
			Session[Encriptar.Decrypt(c_seed,key.ToString())] = Encriptar.Decrypt(c_seed,Request.Form[key].ToString());
		}
	 	 //Response.Write("<br><b>"+Encriptar.Decrypt(c_seed,key.ToString()) + "</b> : "+ Encriptar.Decrypt(c_seed,Request.Form[key].ToString()));
		//Session[key] = Request.Form[key];
	  
	    
	}
}

//string c_seed = "56238";
//string returnURL = Encriptar.Encrypt(c_seed,"returnURL");

Response.Redirect(Encriptar.Decrypt(c_seed,Request.Form[returnURL].ToString()));
%>
