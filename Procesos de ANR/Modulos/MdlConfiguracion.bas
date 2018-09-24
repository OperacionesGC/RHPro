Attribute VB_Name = "MdlConfiguracion"
Option Explicit


'constantes para obtener la información regional
Public Const SORT_DEFAULT = &H0
Public Const LOCALE_FONTSIGNATURE = &H58
Public Const LOCALE_ICENTURY = &H24
Public Const LOCALE_ICOUNTRY = &H5 'código de país
Public Const LOCALE_ICURRDIGITS = &H19 'nº de decimales en las monedas
Public Const LOCALE_ICURRENCY = &H1B 'posición del simbolo de moneda respecto al _
    número,0=delante, 1=detrás, 2=delante con un blanco, 3=detras con un blanco
Public Const LOCALE_IDATE = &H21
Public Const LOCALE_IDAYLZERO = &H26 '1=días con dos dígitos en fecha corta
Public Const LOCALE_IDEFAULTCODEPAGE = &HB 'página de códigos por defecto
Public Const LOCALE_IDEFAULTCOUNTRY = &HA 'código de país por defecto
Public Const LOCALE_IDEFAULTLANGUAGE = &H9 'codigo de lenguaje por defecto
Public Const LOCALE_IDIGITS = &H11 'nº de decimales en los numeros
Public Const LOCALE_IINTLCURRDIGITS = &H1A
Public Const LOCALE_ILANGUAGE = &H1 'codigo del lenguaje
Public Const LOCALE_ILDATE = &H22
Public Const LOCALE_ILZERO = &H12
Public Const LOCALE_IMEASURE = &HD 'sistema de medida, 0=metrico, 1 =EE.UU.
Public Const LOCALE_IMONLZERO = &H27
Public Const LOCALE_INEGCURR = &H1C 'formato nº negativo en las monedas
Public Const LOCALE_INEGSEPBYSPACE = &H57 'un espacio entre el nº y la moneda en los _
    negativos
Public Const LOCALE_INEGSIGNPOSN = &H53 'posicion del signo en las monedas negativas, _
    0=no se pone, 1=antes del numero, 2=despues del numero,3=antes de la moneda, _
    4=despues de la monea
Public Const LOCALE_INEGSYMPRECEDES = &H56
Public Const LOCALE_IPOSSEPBYSPACE = &H55
Public Const LOCALE_IPOSSIGNPOSN = &H52
Public Const LOCALE_IPOSSYMPRECEDES = &H54
Public Const LOCALE_ITIME = &H23
Public Const LOCALE_ITLZERO = &H25 '1=horas con dos digitos
Public Const LOCALE_NOUSEROVERRIDE = &H80000000
Public Const LOCALE_S1159 = &H28 'simbolo a.m.
Public Const LOCALE_S2359 = &H29 'simbolo p.m.
Public Const LOCALE_SABBREVCTRYNAME = &H7 'nombre abreviado del país
Public Const LOCALE_SABBREVDAYNAME1 = &H31 'nombre abreviado de los días de la semana
Public Const LOCALE_SABBREVDAYNAME2 = &H32 'en el idioma del país
Public Const LOCALE_SABBREVDAYNAME3 = &H33
Public Const LOCALE_SABBREVDAYNAME4 = &H34
Public Const LOCALE_SABBREVDAYNAME5 = &H35
Public Const LOCALE_SABBREVDAYNAME6 = &H36
Public Const LOCALE_SABBREVDAYNAME7 = &H37
Public Const LOCALE_SABBREVLANGNAME = &H3 'nombre a breviado del lenguaje
Public Const LOCALE_SABBREVMONTHNAME1 = &H44  'nombre abreviado de los meses del año
Public Const LOCALE_SABBREVMONTHNAME10 = &H4D 'en el idioma del país
Public Const LOCALE_SABBREVMONTHNAME11 = &H4E
Public Const LOCALE_SABBREVMONTHNAME12 = &H4F
Public Const LOCALE_SABBREVMONTHNAME13 = &H100F
Public Const LOCALE_SABBREVMONTHNAME2 = &H45
Public Const LOCALE_SABBREVMONTHNAME3 = &H46
Public Const LOCALE_SABBREVMONTHNAME4 = &H47
Public Const LOCALE_SABBREVMONTHNAME5 = &H48
Public Const LOCALE_SABBREVMONTHNAME6 = &H49
Public Const LOCALE_SABBREVMONTHNAME7 = &H4A
Public Const LOCALE_SABBREVMONTHNAME8 = &H4B
Public Const LOCALE_SABBREVMONTHNAME9 = &H4C
Public Const LOCALE_SCOUNTRY = &H6 'nombre del país en inglés
Public Const LOCALE_SCURRENCY = &H14 'símbolo de la moneda
Public Const LOCALE_SDATE = &H1D 'separador de fechas
Public Const LOCALE_SDAYNAME1 = &H2A 'nombre de los días día de la semana
Public Const LOCALE_SDAYNAME2 = &H2B 'en el idioma del país
Public Const LOCALE_SDAYNAME3 = &H2C
Public Const LOCALE_SDAYNAME4 = &H2D
Public Const LOCALE_SDAYNAME5 = &H2E
Public Const LOCALE_SDAYNAME6 = &H2F
Public Const LOCALE_SDAYNAME7 = &H30
Public Const LOCALE_SDECIMAL = &HE 'separador decimal
Public Const LOCALE_SENGCOUNTRY = &H1002
Public Const LOCALE_SENGLANGUAGE = &H1001
Public Const LOCALE_SGROUPING = &H10 'nº de dígitos en grupo
Public Const LOCALE_SINTLSYMBOL = &H15 'simbolo internacional del pais
Public Const LOCALE_SLANGUAGE = &H2 'lenguaje selecionado en conf.reg.
Public Const LOCALE_SLIST = &HC 'separador de listas
Public Const LOCALE_SLONGDATE = &H20 'formato de fecha larga
Public Const LOCALE_SMONDECIMALSEP = &H16 'separador decimal en las monedas
Public Const LOCALE_SMONGROUPING = &H18 'nº de dígitos en grupo para las monedas
Public Const LOCALE_SMONTHNAME1 = &H38  'nombres de los meses
Public Const LOCALE_SMONTHNAME10 = &H41 'en el idioma del país
Public Const LOCALE_SMONTHNAME11 = &H42
Public Const LOCALE_SMONTHNAME12 = &H43
Public Const LOCALE_SMONTHNAME2 = &H39
Public Const LOCALE_SMONTHNAME3 = &H3A
Public Const LOCALE_SMONTHNAME4 = &H3B
Public Const LOCALE_SMONTHNAME5 = &H3C
Public Const LOCALE_SMONTHNAME6 = &H3D
Public Const LOCALE_SMONTHNAME7 = &H3E
Public Const LOCALE_SMONTHNAME8 = &H3F
Public Const LOCALE_SMONTHNAME9 = &H40
Public Const LOCALE_SMONTHOUSANDSEP = &H17 'separador de miles en las monedas
Public Const LOCALE_SNATIVECTRYNAME = &H8 'nombre del país en el idioma del país
Public Const LOCALE_SNATIVEDIGITS = &H13 'digitos empleados en el país
Public Const LOCALE_SNATIVELANGNAME = &H4 'idioma del país en el idioma del país
Public Const LOCALE_SNEGATIVESIGN = &H51 'simbolo de signo negativo
Public Const LOCALE_SPOSITIVESIGN = &H50 'simbolo de signo positivo
Public Const LOCALE_SSHORTDATE = &H1F 'formato de fecha corta
Public Const LOCALE_STHOUSAND = &HF 'separador de miles
Public Const LOCALE_STIME = &H1E 'separador de horas
Public Const LOCALE_STIMEFORMAT = &H1003 'formato de horas
Public Const LOCALE_SYSTEM_DEFAULT = &H800 'presentar información del sistema
Public Const LOCALE_USER_DEFAULT = &H400 'presentar información del usuario

Public Const LANG_NEUTRAL = &H0
Public Const SUBLANG_NEUTRAL = &H0
Public Const SUBLANG_DEFAULT = &H1
Public Const SUBLANG_SYS_DEFAULT = &H2

Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Long
Declare Function GetUserDefaultLCID% Lib "kernel32" ()

Public Function MAKELCID(ByVal wLanguageId As Long, ByVal wSortId As Long) As Long
    MAKELCID = wSortId * &H10000 + wLanguageId
End Function
Public Function MAKELANGID(ByVal usPrimaryLanguage As Long, ByVal usSubLanguage _
        As Long) As Long
    MAKELANGID = (usSubLanguage * 1024) Or usPrimaryLanguage
End Function

Public Sub SetearConfiguracionRegional()
Dim LCID As Long
Dim Lang As Long


     'LCID = GetUserDefaultLCID
'    LCID = MAKELCID(MAKELANGID(&H2C0A, &H7F), SORT_DEFAULT)
    'Lang = MAKELANGID(&H2C0A, &H7F)
    'Lang = MAKELANGID(LANG_NEUTRAL, SUBLANG_NEUTRAL)
    Lang = MAKELANGID(&H2C0A, SUBLANG_NEUTRAL)
    
    LCID = MAKELCID(Lang, SORT_DEFAULT)
    
    'Configuración del número
    SetLocaleInfo LCID, LOCALE_SDECIMAL, Nuevo_NumeroSeparadorDecimal
    SetLocaleInfo LCID, LOCALE_STHOUSAND, Nuevo_NumeroSeparadorMiles
    'Configuración de la moneda
    SetLocaleInfo LCID, LOCALE_SMONDECIMALSEP, Nuevo_MonedaSeparadorDecimal
    SetLocaleInfo LCID, LOCALE_SMONTHOUSANDSEP, Nuevo_MonedaSeparadorMiles
   
End Sub

Public Sub ObtenerConfiguracionRegional()
Dim Ret As String
Dim buffer As String
    
    buffer = String$(256, 0)
    'Configuración del número
    Ret = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SDECIMAL, buffer, Len(buffer))
    If Ret > 0 Then
        NumeroSeparadorDecimal = Left$(buffer, Ret - 1)
    End If
    Ret = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_STHOUSAND, buffer, Len(buffer))
    If Ret > 0 Then
        NumeroSeparadorMiles = Left$(buffer, Ret - 1)
    End If
    'Configuración de la moneda
    Ret = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SMONDECIMALSEP, buffer, Len(buffer))
    If Ret > 0 Then
        MonedaSeparadorDecimal = Left$(buffer, Ret - 1)
    End If
    
    Ret = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SMONTHOUSANDSEP, buffer, Len(buffer))
    If Ret > 0 Then
        MonedaSeparadorMiles = Left$(buffer, Ret - 1)
    End If
    
    Ret = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SSHORTDATE, buffer, Len(buffer))
    If Ret > 0 Then
        FormatoDeFechaCorto = Left$(buffer, Ret - 1)
    End If

End Sub

'Public Sub EjecutarSQL(ByVal Conexion As ADODB.Connection, ByVal TmpStrSql As String, ByRef rs_Temp As ADODB.Recordset)
'   Dim Err As Error
'
'   ' Ejecuta el sql y levanta los posibles errores (chequeando la Coleccion de Errores)
'   On Error GoTo ME_Execute
'   'cmdTemp.Execute
'   Conexion.Execute TmpStrSql, , adExecuteNoRecords
'   On Error GoTo 0
'
'   ' Actualiza el recordset
'   rs_Temp.Requery
'
'   Exit Sub
'
'ME_Execute:
'    ' Notifica al usuario de cualquier error proveniente de la ejecucuion del SQL
'    If rs_Temp.ActiveConnection.Errors.Count > 0 Then
'        For Each Err In rs_Temp.ActiveConnection.Errors
'            Flog.writeln "Error: " & Err.Number & vbCr & Err.Description
'        Next Err
'    End If
'    Resume Next
'End Sub






Private Function Encrypt(ByVal strEncryptionKey, ByVal strTextToEncrypt)
'Encriptar un string
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


Private Function Decrypt(ByVal strEncryptionKey, ByVal strTextToEncrypt)
'Desencriptar un string
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
Dim buffer, outer
    buffer = ""
    For outer = 1 To Len(strTextToEncrypt) Step 2
        buffer = buffer & Chr(CLng("&h" & Mid(strTextToEncrypt, outer, 2)))
    Next
    CadenaAscii = buffer
End Function
