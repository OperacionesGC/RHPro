<%
'Descripción	: CopyAspNetSessionToAspSession.asp 
'Fecha		: ??
'Autor		: ??
'Modificado	: 22/01/2013 - Lisandro Moro - Agrego versionado y session abandon
'Modificado	: 28/02/2013 - 16660 - Lisandro Moro - Agrego pasaje de variables de session a rhpro.net
'Modificado	: 09/01/2014 - Gonzalez Nicolás - Se cambió la forma de separar  los parametros recibidos.Soluciona errro con @ en las contraseñas.
'             29/05/2014 - JPB - Levanto el estilo del home y las cargo en las variables de session
'			  25/06/2014 - MDZ - CAS-20270 - no se opian variables con log de fuentes y querys (RHpro_log?)
'             01/09/2014 - JPB - CAS-20903 - Heidt & Asoc. - Ingreso clientes desde MetaHome [Entrega2]
'                              - Se evita el pasaje de variables de sesion con gran tamaño y especificas que provienen del home.
%>
<!--#include virtual="/rhprox2/shared/inc/encrypt.inc"-->
<html>
<head>
</head>
<body onLoad="document.datos.submit()">
<form target="_self" name="datos" method="post" action="CopyAspSessionToAspNetSession.aspx">
<%
on error goto 0
'Dim p, param, params, array_param, l_abandon
Dim p,  params, array_param, l_abandon
Dim param(1)
Dim arrQS 'lista con las variables del querystring o form
Dim c_seed

c_seed = session("c_seed1")
if c_seed = "" then
	c_seed = "56238"
end if

l_abandon = false

Session.LCID = 11274
'Session.LCID = 2108

params =  request("params")

'hago un split para separar cada uno de los parametros
array_param = Split(params,"_")

 

<!--#include virtual="/rhprox2/shared/inc/sec.inc"-->
<!--#include virtual="/rhprox2/shared/inc/const.inc"-->
 
for i=0 to UBound(array_param)
    'Decripto cada uno de los parametros y se lo asigno a la session
    p = Decrypt(c_seed, array_param(i))	

	'09/01/2014 - NG	
	param(0) = Mid(p,1,Instr(1,p,"@")-1)
	param(1) = Mid(p,Instr(1,p,"@")+1,len(p))

    '09/01/2014 - NG
	'param = Split(p,"@")
	
	'response.write param(0) & " " & param(1) & "<br>"
	arrQS = arrQS & param(0) & ","
	'response.Write("<!--input type=""hidden"" name=""" & enc(param(0))& """ value=""" & enc(param(1)) & """-->")
	
	if ((ucase(param(0))<> Ucase("VisualizaModulos"))   and (left(ucase(Trim(param(0))),9) <> Ucase("RHpro_log")) and (ucase(param(0))<> Ucase("RHPRO_Home_MenuPrincipal")) ) then	
		if ucase(param(0)) = "PASSWORD" then
			'Session(param(0)) = enc(param(1))			
			Session(param(0)) = enc(param(1))
			'response.Write("<input type=""hidden"" name=""" & enc(param(0))& """ value=""" & enc(param(1)) & """>")
		else
		' Response.Write("<BR><b>VALOR:</b>"&param(0))
			Session(param(0)) = param(1)
			'response.Write("<input type=""hidden"" name=""" & enc(param(0)) & """ value=""" & enc(param(1)) & """>")
		end if

		response.Write("<input type=""hidden"" name=""" & enc(param(0)) & """ value=""" & enc(param(1)) & """>")

	end if
	
	if ucase(param(0)) = ucase("UserName") AND param(1) = "" then
		l_abandon = true
		Session.Abandon
		response.Write("<input type=""hidden"" name=""" & enc("Abandon") & """ value=""" & enc("1") & """>")
	 
	end if    

Next
response.Write("<input type=""hidden"" name=""" &  enc("returnURL") & """ value=""" &  enc(Request("returnURL")) & """>")
 
 	'if  instr("RHprao_log@", "RHpro_log") > 0 then
	'if (left(ucase(Trim("RHprao_log@")),9) <> Ucase("RHpro_log")) then	
	'		response.Write("***  *********************************************************")
	'end if
	
 
 
if not l_abandon then

	%>
	<!--#include virtual="/rhprox2/shared/db/conn_db.inc"-->	
	<%
	 
	'Response.Write("<script>alert(2);</script>")
	Dim item, itemloop
	For Each item in Session.Contents	
	   
		if (left(ucase(Trim(item)),9) <> Ucase("RHpro_log")) then 'Restrinjo variables de session
	  
			if clng(VarType(Session(item)))<>clng(8204) then
				rhprodebug " copiar : " & item & " = "  & Session(item)
			else
				rhprodebug " copiar : " & item & " = "  & TypeName(Session(item))
			end if
			
			If IsArray(Session(item)) then
				For itemloop = LBound(Session(item)) to UBound(Session(item))
					rhprodebug " sub item copiar : " & enc(itemloop)
					if notIn(itemloop) then
						response.Write("<input type=""hidden"" name=""" &  enc(itemloop) & """ value=""" &  enc(Session(item)(itemloop)) & """>")
					end if
				Next
			Else
				if notIn(item) then
					response.Write("<input type=""hidden"" name=""" &  enc(item) & """ value=""" &  enc(Session.Contents(item)) & """>")
				end if
			End If
		end if	
	Next
 	 
end if

function notIn(que)
	dim arr, i
	notIn = true
	arr = split(arrQS,",")
	for i=0 to ubound(arr)
		if ucase(que)= ucase(arr(i)) then
			notIn = false
		end if
	next
end function
 

function enc(text)
	enc = Encrypt(c_seed,text)
end function

function dec(text)
	dec = Decrypt(c_seed,text)
end function

%>
</form>
</body>
</html>
