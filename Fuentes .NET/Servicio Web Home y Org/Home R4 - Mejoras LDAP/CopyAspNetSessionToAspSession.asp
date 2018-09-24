<html>
<head>

</head>
<body>
 

<%

Dim p,param,params, array_param  

Session.LCID = 11274

params =  request("params")

'hago un split para separar cada uno de los parametros
array_param = Split(params,"_")

for i=0 to UBound(array_param)
    'Decripto cada uno de los parametros y se lo asigno a la session

    p = Decrypt("56238", array_param(i))
    param = Split(p,"@")
	response.write param(0) & " " & param(1) & "<br>"
	if ucase(param(0)) = "PASSWORD" then
    	Session(param(0)) = Encrypt("56238", param(1))
	else
		Session(param(0)) = param(1)
	end if
    
Next

Response.Redirect(Request("returnURL"))

'Encriptar un string
Private Function Encrypt(ByVal strEncryptionKey, ByVal strTextToEncrypt)
    Dim outer, inner, Key, strTemp, buffer

    For outer = 1 To Len(strEncryptionKey)
		key = Asc(Mid(strEncryptionKey, outer, 1))
		For inner = 1 To Len(strTextToEncrypt)
            strTemp = strTemp & Chr(Asc(Mid(strTextToEncrypt, inner, 1)) Xor key)
            key = (key + Len(strEncryptionKey)) Mod 256
        Next
        strTextToEncrypt = strTemp
        strTemp = ""
    Next

    strTextToEncrypt = CadenaHex(strTextToEncrypt)	

	Encrypt = strTextToEncrypt
End Function

'Desencriptar un string
Private Function Decrypt(ByVal strEncryptionKey, ByVal strTextToEncrypt)
    Dim outer, inner, Key, strTemp, buffer
	
	strTextToEncrypt = CadenaAscii(strTextToEncrypt)

    'Response.Write("strTextToEncrypt" + strTextToEncrypt)

    For outer = 1 To Len(strEncryptionKey)
		key = Asc(Mid(strEncryptionKey, outer, 1))
		For inner = 1 To Len(strTextToEncrypt)
        
          'Response.Write (Mid(strTextToEncrypt, inner, 1) + "   ")
          'Response.Write (cstr(Asc(Mid(strTextToEncrypt, inner, 1))) + "<br />")
        
            strTemp = strTemp & Chr(Asc(Mid(strTextToEncrypt, inner, 1)) Xor key)
            key = (key + Len(strEncryptionKey)) Mod 256
        Next
        strTextToEncrypt = strTemp
        strTemp = ""
    Next

	Decrypt = strTextToEncrypt
End Function

Function CadenaHex(ByVal strTextToEncrypt)
    Dim buffer,outer,auxi
	buffer = ""
	For outer = 1 To Len(strTextToEncrypt)
		auxi = Hex(Asc(Mid(strTextToEncrypt, outer, 1)))
		if len(auxi) < 2 then auxi = "0" & auxi
		buffer = buffer & auxi
    Next
	CadenaHex = buffer
end function

Function CadenaAscii(ByVal strTextToEncrypt)
    
    'Response.write("CADENA ASCII<BR /><BR /><BR /><BR />")
    Dim buffer,outer
	buffer = ""
	For outer = 1 To Len(strTextToEncrypt) step 2
        'Response.write(cstr( Mid(strTextToEncrypt, outer, 2)) + " $$ " +  cstr( CLng("&h" & Mid(strTextToEncrypt, outer, 2))  ) + " $$ " +  chrw(CLng("&h" & Mid(strTextToEncrypt, outer, 2))) + "<br />")
		buffer = buffer & chr(CLng("&h" & Mid(strTextToEncrypt, outer, 2)))
    Next
	CadenaAscii = buffer
    'Response.write("FIN CADENA ASCII<BR /><BR /><BR /><BR />")
end function

function obf (ByVal Src)
    Dim StrLen
	Dim CheckSum
	Dim I
	StrLen = Len(Src)
    CheckSum = 0
    For I = 1 To StrLen
        If I Mod 2 = 1 Then
            CheckSum = CheckSum + (Asc(Mid(Src, StrLen - I + 1, 1)) - Asc("0")) * 3333
        Else
            CheckSum = CheckSum + Asc(Mid(Src, StrLen - I + 1, 1)) - Asc("0")
        End If
    Next
    obf = CheckSum
End Function
%>

</body>
</html>
