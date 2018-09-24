Attribute VB_Name = "MdlEncrypt"
'Encriptar un string
Public Function Encrypt(ByVal strEncryptionKey, ByVal strTextToEncrypt)
    Dim outer, inner, Key, strTemp, buffer

    For outer = 1 To Len(strEncryptionKey)
        Key = Asc(Mid(strEncryptionKey, outer, 1))
        For inner = 1 To Len(strTextToEncrypt)
            strTemp = strTemp & Chr(Asc(Mid(strTextToEncrypt, inner, 1)) Xor Key)
            Key = (Key + Len(strEncryptionKey)) Mod 256
        Next
        strTextToEncrypt = strTemp
        strTemp = ""
    Next

    strTextToEncrypt = CadenaHex(strTextToEncrypt)

    Encrypt = strTextToEncrypt
End Function


'Desencriptar un string
Public Function Decrypt(ByVal strEncryptionKey, ByVal strTextToEncrypt)
    Dim outer, inner, Key, strTemp, buffer
    
    strTextToEncrypt = CadenaAscii(strTextToEncrypt)

    For outer = 1 To Len(strEncryptionKey)
        Key = Asc(Mid(strEncryptionKey, outer, 1))
        For inner = 1 To Len(strTextToEncrypt)
            strTemp = strTemp & Chr(Asc(Mid(strTextToEncrypt, inner, 1)) Xor Key)
            Key = (Key + Len(strEncryptionKey)) Mod 256
        Next
        strTextToEncrypt = strTemp
        strTemp = ""
    Next

    Decrypt = strTextToEncrypt
End Function


Public Function CadenaHex(ByVal strTextToEncrypt)
    Dim buffer, outer, auxi
    buffer = ""
    For outer = 1 To Len(strTextToEncrypt)
        auxi = Hex(Asc(Mid(strTextToEncrypt, outer, 1)))
        If Len(auxi) < 2 Then auxi = "0" & auxi
        buffer = buffer & auxi
    Next
    CadenaHex = buffer
End Function


Public Function CadenaAscii(ByVal strTextToEncrypt)
    Dim buffer, outer
    buffer = ""
    For outer = 1 To Len(strTextToEncrypt) Step 2
        buffer = buffer & Chr(CLng("&h" & Mid(strTextToEncrypt, outer, 2)))
    Next
    CadenaAscii = buffer
End Function

