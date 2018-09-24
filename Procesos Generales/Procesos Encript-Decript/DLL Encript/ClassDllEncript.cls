VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Autor: Fernando Zwenger
'FEcha: 11/02/2009
'Descripcion:
'   Programa para encriptar y desencriptar
'----------------------------------------
Option Explicit

'---------------------------------------------------------------------------------------------------
'Versiones
Const Version = "1.00"    'Inicial.
Const FechaVersion = "11/02/2009"

'---------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------

'Declaraciones
'Const MaxPendientes = 1000
'Const ForReading = 1
'Const ForAppending = 8
'Const ForWriting = 2
'Const FormatoInternoFecha = "dd/mm/yyyy HH:mm:ss"
'Const FormatoInternoHora = "HH:mm:ss"

Public Function Encrypt(ByVal strEncryptionKey, ByVal strTextToEncrypt)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento para Encriptar un string.
' Autor      : FGZ
' Fecha      : 11/02/2009
' Ultima Mod.:
' ---------------------------------------------------------------------------------------------
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


Public Function Decrypt(ByVal strEncryptionKey, ByVal strTextToEncrypt)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento para Desencriptar un string.
' Autor      : FGZ
' Fecha      : 11/02/2009
' Ultima Mod.:
' ---------------------------------------------------------------------------------------------
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

Function CadenaHex(ByVal strTextToEncrypt)
' ---------------------------------------------------------------------------------------------
' Descripcion: Funcion complementaria para Encriptar un string.
' Autor      : FGZ
' Fecha      : 11/02/2009
' Ultima Mod.:
' ---------------------------------------------------------------------------------------------
Dim buffer, outer, auxi
    buffer = ""
    For outer = 1 To Len(strTextToEncrypt)
        auxi = Hex(Asc(Mid(strTextToEncrypt, outer, 1)))
        If Len(auxi) < 2 Then auxi = "0" & auxi
        buffer = buffer & auxi
    Next
    CadenaHex = buffer
End Function

Function CadenaAscii(ByVal strTextToEncrypt)
' ---------------------------------------------------------------------------------------------
' Descripcion: Funcion complementaria para Desencriptar un string.
' Autor      : FGZ
' Fecha      : 11/02/2009
' Ultima Mod.:
' ---------------------------------------------------------------------------------------------
Dim buffer, outer
    buffer = ""
    For outer = 1 To Len(strTextToEncrypt) Step 2
        buffer = buffer & Chr(CLng("&h" & Mid(strTextToEncrypt, outer, 2)))
    Next
    CadenaAscii = buffer
End Function



