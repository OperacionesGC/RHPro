Attribute VB_Name = "mdlDataAccess"
Option Explicit

' definicion de variables globales de configuracion basica
Global strformatoFservidor As String
Global strconexion As String
Global EncriptStrconexion As Boolean
Global Error_Encrypt As Boolean
Global c_seed As String
Global Ya_Encripto As Boolean
Global TipoBD As String
                        ' DB2 = 1
                        ' Informix = 2
                        ' SQL Server = 3
Global PathFLog As String

'FGZ - 18/06/2004
Global Setea_Configuracion_Regional As Boolean
Global NumeroSeparadorDecimal As String
Global NumeroSeparadorMiles As String
Global MonedaSeparadorDecimal As String
Global MonedaSeparadorMiles As String
Global FormatoDeFechaCorto As String
Global Nuevo_NumeroSeparadorDecimal As String
Global Nuevo_NumeroSeparadorMiles As String
Global Nuevo_MonedaSeparadorDecimal As String
Global Nuevo_MonedaSeparadorMiles As String
Global Nuevo_FormatoDeFechaCorto As String

' -------------------------------
' FGZ - 28/10/2003 Unificacion del .ini
Global PathSAP As String
Global PathProcesos As String

Global objConn As New ADODB.Connection
Global objconn2 As New ADODB.Connection
Global CnTraza As New ADODB.Connection
Global objRs As New ADODB.Recordset
Global dummy As Long
Global TransactionRunning As Boolean
Global T2Running As Boolean
Global Reg_Afected
Global SCHEMA As String
Global Usa_Estadisticas As Boolean



Public Sub CargarConfiguracionesBasicas()
' carga las configuraciones basicas para los procesos
Const ForReading = 1
Const ForAppending = 8

Dim f, fs
Dim strline As String
Dim pos1 As Integer
Dim pos2 As Integer
Dim Encontro As Boolean

    On Error Resume Next
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.OpenTextFile(App.Path & "\rhproprocesos.ini", ForReading, 0)
    If Err.Number <> 0 Then
        On Error GoTo 0
        Set f = fs.OpenTextFile(App.Path & "\rhproappsrv.ini", ForReading, 0)
    End If
    If Not EsNulo(Etiqueta) Then
        Encontro = False
        Do While Not f.AtEndOfStream And Not Encontro
            strline = f.ReadLine()
            If InStr(1, UCase(strline), UCase(Etiqueta)) > 0 Then
                Encontro = True
            End If
        Loop
    End If
    
    ' Path del Proceso de SAP (lo usa el SAP)
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        PathSAP = Mid(strline, pos1, pos2 - pos1)
        If Right(PathSAP, 1) <> "\" Then PathSAP = PathSAP & "\"
    End If
    
    ' Path de los ejecutables de los procesos (lo usa el AppSrv)
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        PathProcesos = Mid(strline, pos1, pos2 - pos1)
        If Right(PathProcesos, 1) <> "\" Then PathProcesos = PathProcesos & "\"
    End If
    
    ' seteo del path del archivo de Log
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        PathFLog = Mid(strline, pos1, pos2 - pos1)
        If Right(PathFLog, 1) <> "\" Then PathFLog = PathFLog & "\log\"
    End If
    
    ' seteo del formato de Fecha del Servidor
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        strformatoFservidor = Mid(strline, pos1, pos2 - pos1)
    End If

    ' seteo del string de conexion
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        strconexion = Mid(strline, pos1, pos2 - pos1)
    End If
    
    ' seteo del tipo de Base de datos
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        TipoBD = Mid(strline, pos1, pos2 - pos1)
    End If

    'Etiqueta (solo para el AppSrv)
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
    End If

    'FGZ - 23/06/2008 - Se agregó este parametro
    ' seteo del schema de Base de datos
    SCHEMA = ""
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        SCHEMA = Mid(strline, pos1, pos2 - pos1)
    End If

    f.Close

End Sub


Public Sub CargarConfiguracionesBasicas_old3()
' carga las configuraciones basicas para los procesos
Const ForReading = 1
Const ForAppending = 8

Dim f, fs
Dim strline As String
Dim pos1 As Long
Dim pos2 As Long
Dim Encontro As Boolean
Dim UltimaVersion As Boolean
Dim Seteado As Boolean

    On Error Resume Next
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.OpenTextFile(App.Path & "\rhproprocesos.ini", ForReading, 0)
    If Err.Number <> 0 Then
        UltimaVersion = False
        On Error GoTo 0
        Set f = fs.OpenTextFile(App.Path & "\rhproappsrv.ini", ForReading, 0)
    Else
        UltimaVersion = True
    End If
    If Not EsNulo(Etiqueta) Then
        Encontro = False
        Do While Not f.AtEndOfStream And Not Encontro
            strline = f.ReadLine()
            If InStr(1, UCase(strline), UCase(Etiqueta)) > 0 Then
                Encontro = True
            End If
        Loop
    End If
    
    ' Path del Proceso de SAP (lo usa el SAP)
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        PathSAP = Mid(strline, pos1, pos2 - pos1)
        If Right(PathSAP, 1) <> "\" Then PathSAP = PathSAP & "\"
    End If
    
    ' Path de los ejecutables de los procesos (lo usa el AppSrv)
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        PathProcesos = Mid(strline, pos1, pos2 - pos1)
        If Right(PathProcesos, 1) <> "\" Then PathProcesos = PathProcesos & "\"
    End If
    
    ' seteo del path del archivo de Log
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        PathFLog = Mid(strline, pos1, pos2 - pos1)
        If Right(PathFLog, 1) <> "\" Then PathFLog = PathFLog & "\"
    End If
    
    ' seteo del formato de Fecha del Servidor
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        strformatoFservidor = Mid(strline, pos1, pos2 - pos1)
    End If

    ' seteo del string de conexion
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        strconexion = Mid(strline, pos1, pos2 - pos1)
    End If
    
    ' seteo del tipo de Base de datos
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        TipoBD = Mid(strline, pos1, pos2 - pos1)
    End If

'    'seteo la Configuracion Regional?
'    If Not f.AtEndOfStream Then
'        strline = f.ReadLine()
'        If Not UltimaVersion Then
'            If Not f.AtEndOfStream Then
'                strline = f.ReadLine()
'                Seteado = True
'            Else
'                Seteado = False
'            End If
'        Else
'            Seteado = True
'        End If
'    Else
'        Seteado = True
'    End If
    Seteado = False
    If Seteado Then
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        If pos2 > pos1 Then
            Setea_Configuracion_Regional = CBool(Mid(strline, pos1, pos2 - pos1))
        Else
            Setea_Configuracion_Regional = False
        End If
    Else
        Setea_Configuracion_Regional = False
    End If
    

'configuracion regional
If Setea_Configuracion_Regional Then
    'Numero
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        Nuevo_NumeroSeparadorDecimal = Mid(strline, pos1, pos2 - pos1)
    End If
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        Nuevo_NumeroSeparadorMiles = Mid(strline, pos1, pos2 - pos1)
    End If
    
    Nuevo_MonedaSeparadorDecimal = Nuevo_NumeroSeparadorDecimal
    Nuevo_MonedaSeparadorMiles = Nuevo_NumeroSeparadorMiles
    
    'Call SetearConfiguracionRegional
End If


''Moneda
'    If Not f.AtEndOfStream Then
'        strline = f.ReadLine()
'        pos1 = InStr(1, strline, "[") + 1
'        pos2 = InStr(1, strline, "]")
'        Nuevo_MonedaSeparadorDecimal = Mid(strline, pos1, pos2 - pos1)
'    End If
'    If Not f.AtEndOfStream Then
'        strline = f.ReadLine()
'        pos1 = InStr(1, strline, "[") + 1
'        pos2 = InStr(1, strline, "]")
'        Nuevo_MonedaSeparadorMiles = Mid(strline, pos1, pos2 - pos1)
'    End If
        
    f.Close

End Sub


Public Sub CargarConfiguracionesBasicas_old2()
' carga las configuraciones basicas para los procesos
Const ForReading = 1
Const ForAppending = 8

Dim f, fs
Dim strline As String
Dim pos1 As Long
Dim pos2 As Long

    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.OpenTextFile(App.Path & "\rhproappsrv.ini", ForReading, 0)
    
    ' Path del Proceso de SAP (lo usa el SAP)
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        PathSAP = Mid(strline, pos1, pos2 - pos1)
        If Right(PathSAP, 1) <> "\" Then PathSAP = PathSAP & "\"
    End If
    
    ' Path de los ejecutables de los procesos (lo usa el AppSrv)
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        PathProcesos = Mid(strline, pos1, pos2 - pos1)
        If Right(PathProcesos, 1) <> "\" Then PathProcesos = PathProcesos & "\"
    End If
    
    ' seteo del path del archivo de Log
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        PathFLog = Mid(strline, pos1, pos2 - pos1)
        If Right(PathFLog, 1) <> "\" Then PathFLog = PathFLog & "\"
    End If
    
    ' seteo del formato de Fecha del Servidor
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        strformatoFservidor = Mid(strline, pos1, pos2 - pos1)
    End If

    ' seteo del string de conexion
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        strconexion = Mid(strline, pos1, pos2 - pos1)
    End If
    
    ' seteo del tipo de Base de datos
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        TipoBD = Mid(strline, pos1, pos2 - pos1)
    End If

''FGZ - 18/06/2004
''configuracion regional
''Numero
'    If Not f.AtEndOfStream Then
'        strline = f.ReadLine()
'        pos1 = InStr(1, strline, "[") + 1
'        pos2 = InStr(1, strline, "]")
'        Nuevo_NumeroSeparadorDecimal = Mid(strline, pos1, pos2 - pos1)
'    End If
'    If Not f.AtEndOfStream Then
'        strline = f.ReadLine()
'        pos1 = InStr(1, strline, "[") + 1
'        pos2 = InStr(1, strline, "]")
'        Nuevo_NumeroSeparadorMiles = Mid(strline, pos1, pos2 - pos1)
'    End If
''Moneda
'    If Not f.AtEndOfStream Then
'        strline = f.ReadLine()
'        pos1 = InStr(1, strline, "[") + 1
'        pos2 = InStr(1, strline, "]")
'        Nuevo_MonedaSeparadorDecimal = Mid(strline, pos1, pos2 - pos1)
'    End If
'    If Not f.AtEndOfStream Then
'        strline = f.ReadLine()
'        pos1 = InStr(1, strline, "[") + 1
'        pos2 = InStr(1, strline, "]")
'        Nuevo_MonedaSeparadorMiles = Mid(strline, pos1, pos2 - pos1)
'    End If
        
    f.Close

End Sub


Public Sub CargarConfiguracionesBasicas_OLD()
' carga las configuraciones basicas para los procesos
Const ForReading = 1
Const ForAppending = 8

Dim f, fs
Dim strline As String
Dim pos1 As Long
Dim pos2 As Long

    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.OpenTextFile(App.Path & "\procesos.INI", ForReading, 0)
    
    ' seteo del path del archivo de Log
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        PathFLog = Mid(strline, pos1, pos2 - pos1)
        If Right(PathFLog, 1) <> "\" Then PathFLog = PathFLog & "\"
    End If
    
    ' seteo del formato de Fecha del Servidor
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        strformatoFservidor = Mid(strline, pos1, pos2 - pos1)
    End If

    ' seteo del string de conexion
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        strconexion = Mid(strline, pos1, pos2 - pos1)
    End If
    
    ' seteo del tipo de Base de datos
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        TipoBD = Mid(strline, pos1, pos2 - pos1)
    End If

        
    f.Close

End Sub


Public Function TamañoCampoBD(ByVal Tabla As String, ByVal Campo As String) As Long
' ---------------------------------------------------------------------------------------------
' Descripcion: Retorna la longitud del campo de la tabla de la BD
' Autor      : FGZ
' Fecha      : 24/02/2006
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs As New ADODB.Recordset

    Select Case TipoBD
    Case 1 'DB2
        StrSql = ""
    Case 2 'Informix
        StrSql = ""
    Case 3 'SQL Server
        StrSql = "SELECT DISTINCT b.name AS tablename, LEFT(a.name, 30) AS fieldname, left(c.name,20) AS datatype, a.length AS size"
        StrSql = StrSql & "FROM syscolumns a"
        StrSql = StrSql & "INNER JOIN systypes c ON a.xtype = c.xusertype"
        StrSql = StrSql & "INNER JOIN sysobjects b ON a.id = b.id"
        StrSql = StrSql & "WHERE (b.name = '" & Tabla & "') AND (a.name = '" & Campo & "')"
        StrSql = StrSql & "ORDER BY size desc, tablename, fieldname"
    Case 4 'Oracle
        StrSql = ""
    End Select
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        TamañoCampoBD = rs!Size
    Else
        TamañoCampoBD = -1
    End If

End Function

Public Function getLastIdentity(ByRef objConn As ADODB.Connection, ByVal NombreTabla As String) As Variant
Dim objRs As New ADODB.Recordset
    Dim StrSql As String
    Dim sqlAux As String
    Dim VersionBD
    
    Select Case TipoBD
    Case 1 'db2
        StrSql = "SELECT identity_val_local() as Codigo FROM sysibm.sysdummy1"
    Case 2 ' Informix
        StrSql = "select unique DBINFO('sqlca.sqlerrd1') as codigo from " & NombreTabla
    Case 3 ' sql server
        'seba desde aca
        sqlAux = " SELECT SERVERPROPERTY('productversion') nroversion "
        OpenRecordset sqlAux, objRs
        If Not objRs.EOF Then
            VersionBD = objRs!nroversion
            VersionBD = Split(VersionBD, ".")
            VersionBD = VersionBD(0)
        Else
            VersionBD = 0
            Flog.writeline "No se encontro la version de sql server"
        End If
        objRs.Close
            
            Select Case VersionBD
                Case "11"
                    StrSql = "SELECT IDENT_CURRENT('" & NombreTabla & "') as codigo"
                Case "10"
                    StrSql = "SELECT IDENT_CURRENT('" & NombreTabla & "') as codigo"
                Case "10"
                    StrSql = "SELECT IDENT_CURRENT('" & NombreTabla & "') as codigo"
                'case "2005"
                'case "2000"
                'case "7.0"
                'case "6.5"
                Case Else 'default
                    StrSql = "select @@IDENTITY as codigo FROM " & NombreTabla
            End Select
                
            'hasta aca
    
            'StrSql = "SELECT SCOPE_IDENTITY() as Codigo "
            'StrSql = "SELECT IDENT_CURRENT('" & NombreTabla & "') as codigo"
            'StrSql = "SELECT @@IDENTITY as Codigo FROM " & NombreTabla & ""
            'StrSql = "SELECT @@IDENTITY as Codigo"
    Case 4 ' Oracle 9
        StrSql = "select SEQ_" & UCase(NombreTabla) & ".CURRVAL as Codigo FROM DUAL"
    End Select
    OpenRecordset StrSql, objRs
    If Not (objRs.EOF And objRs.BOF) Then
        If Not EsNulo(objRs.Fields("codigo").Value) Then
            getLastIdentity = objRs.Fields("codigo").Value
        Else
            getLastIdentity = -1
        End If
    Else
        getLastIdentity = -1
    End If
End Function

Public Function getLastIdentity2(ByRef Conn As ADODB.Connection, ByVal NombreTabla As String) As Variant
Dim objRs As New ADODB.Recordset
    Dim StrSql As String
    Select Case TipoBD
    Case 1 'db2
        StrSql = "SELECT identity_val_local() as Codigo FROM sysibm.sysdummy1"
    Case 2 ' Informix
        StrSql = "select unique DBINFO('sqlca.sqlerrd1') as codigo from " & NombreTabla
    Case 3 ' sql server
            StrSql = "SELECT SCOPE_IDENTITY() as Codigo "
            'StrSql = "SELECT @@IDENTITY as Codigo FROM " & NombreTabla & ""
            'StrSql = "SELECT @@IDENTITY as Codigo"
    Case 4 ' Oracle 9
        StrSql = "select SEQ_" & UCase(NombreTabla) & ".CURRVAL as Codigo FROM DUAL"
    End Select
    OpenRecordsetWithConn StrSql, objRs, Conn
    If Not (objRs.EOF And objRs.BOF) Then
        If Not EsNulo(objRs.Fields("codigo").Value) Then
            getLastIdentity2 = objRs.Fields("codigo").Value
        Else
            getLastIdentity2 = -1
        End If
    Else
        getLastIdentity2 = -1
    End If
End Function

Public Function Str2SQLField(ByVal cadena As String) As String
    'Transforma una cadena de caracteres en compatible para SQL (transforma el apostrofo por un acento francés)
    'PARA VALIDAR ENTRADAS DE CAMPOS
         
    Str2SQLField = Replace(cadena, "'", "`")
End Function

Public Sub OpenConnection(strConnectionString As String, ByRef objConn As ADODB.Connection)
    Error_Encrypt = False
    On Error GoTo Manejador
    If Not Ya_Encripto Then
        If EncriptStrconexion Then
            'strconexion = Encrypt(c_seed, strconexion)
            strConnectionString = Decrypt(c_seed, strConnectionString)
            Ya_Encripto = True
        End If
    End If
    
    If objConn.State <> adStateClosed Then objConn.Close
    'objConn.CursorLocation = adUseServer
    objConn.CursorLocation = adUseClient
    'objconn.IsolationLevel =
    objConn.IsolationLevel = adXactCursorStability
    'objConn.IsolationLevel = adXactBrowse
    objConn.CommandTimeout = 60
    objConn.ConnectionTimeout = 60
    objConn.Open strConnectionString
'   If TipoBD = 2 Then
'       objConn.Execute "SET LOCK MODE TO WAIT 60"
'   End If
    Select Case TipoBD
    Case 2:
        objConn.Execute "SET LOCK MODE TO WAIT 60"
    Case 4:
    If Not EsNulo(SCHEMA) Then
        'objConn.Execute "ALTER SESSION SET NLS_SORT = BINARY"
        objConn.Execute "ALTER SESSION SET CURRENT_SCHEMA = " & SCHEMA
    End If
   Case Else
   End Select
Exit Sub
Manejador:
    'Flog.writeline "La conexion no se puede desencriptar. Revise el string de conexion configurado en el .ini"
    Error_Encrypt = True
End Sub


Public Sub OpenConnection_old(strConnectionString As String, ByRef objConn As ADODB.Connection)
    If objConn.State <> adStateClosed Then objConn.Close
    'objConn.CursorLocation = adUseServer
    objConn.CursorLocation = adUseClient
    
    'objConn.IsolationLevel = adXactCursorStability
    'Indica que desde una transacción se pueden ver cambios que no se han producido
    'en otras transacciones.
    objConn.IsolationLevel = adXactReadUncommitted
    
    'objConn.IsolationLevel = adXactBrowse
    objConn.CommandTimeout = 3600 'segundos
    objConn.ConnectionTimeout = 60 'segundos
    objConn.Open strConnectionString
    If TipoBD = 2 Then
        objConn.Execute "SET LOCK MODE TO WAIT 60"
    End If
   
End Sub
Public Sub OpenRecordset(strSQLQuery As String, ByRef objRs As ADODB.Recordset, Optional lockType As LockTypeEnum = adLockReadOnly)
Dim pos1 As Long
Dim pos2 As Long
Dim aux As String

    'Abre un recordset con la consulta strSQLQuery
    If objRs.State <> adStateClosed Then
        If objRs.lockType <> adLockReadOnly Then objRs.UpdateBatch
        objRs.Close
    End If
    
    'Algunas propiedades de prueba
'    objRs.CursorType = 0 'adForwardOnly
'    objRs.CursorLocation = adUseServer
'    objRs.lockType = adLockReadOnly
    objRs.CacheSize = 500
    objRs.Open strSQLQuery, objConn, adOpenDynamic, lockType, adCmdText
    
    Cantidad_de_OpenRecordset = Cantidad_de_OpenRecordset + 1
    
If Usa_Estadisticas Then
    If InStr(1, strSQLQuery, "*", vbTextCompare) > 0 Then
        Cantidad_Select = Cantidad_Select + 1
    End If
    'Parta recolectar estadisticas
    pos1 = InStr(1, strSQLQuery, "from", vbTextCompare) + 5
    'pos1 = InStr(1, strSQLQuery, "* from", vbTextCompare) + 5
    If pos1 > 5 Then
        pos2 = InStr(pos1, strSQLQuery, " ")
        If pos2 = 0 Then
            pos2 = Len(strSQLQuery) + 1
        End If
        aux = UCase(Mid(strSQLQuery, pos1, pos2 - pos1))
              
        Select Case aux
        Case "HIS_ESTRUCTURA":
            Cantidad_His_Estructura = Cantidad_His_Estructura + 1
        Case "DETLIQ":
            Cantidad_Detliq = Cantidad_Detliq + 1
        Case "NOVGRAL":
            Cantidad_NovG = Cantidad_NovG + 1
        Case "NOVESTR"
            Cantidad_NovE = Cantidad_NovE + 1
        Case "NOVEMP"
            Cantidad_NovI = Cantidad_NovI + 1
        Case "NOVAJU"
            Cantidad_NovA = Cantidad_NovA + 1
        Case "ACU_MES":
            Cantidad_Acumes = Cantidad_Acumes + 1
        Case "ACU_LIQ":
            Cantidad_Acu_liq = Cantidad_Acu_liq + 1
        Case "#WF_TPA"
            Cantidad_WF_Tpa = Cantidad_WF_Tpa + 1
        Case "#WF_IMPPROARG":
            Cantidad_WF_Impproarg = Cantidad_WF_Impproarg + 1
        Case "#WF_RETROACTIVO":
            Cantidad_WF_Retro = Cantidad_WF_Retro + 1

        End Select
    End If
End If
    
End Sub

Public Sub OpenRecordsetLectura(strSQLQuery As String, ByRef objRs As ADODB.Recordset, Optional lockType As LockTypeEnum = adLockReadOnly)
Dim pos1 As Long
Dim pos2 As Long
Dim aux As String

    'Abre un recordset con la consulta strSQLQuery
    If objRs.State <> adStateClosed Then
        If objRs.lockType <> adLockReadOnly Then objRs.UpdateBatch
        objRs.Close
    End If
    
    'Algunas propiedades de prueba
'    objRs.CursorType = 0 'adForwardOnly
'    objRs.CursorLocation = adUseServer
    objRs.lockType = adLockReadOnly
    objRs.CacheSize = 500

    objRs.Open strSQLQuery, objConn, adOpenDynamic, lockType, adCmdText
    
'    pos1 = InStr(1, strSQLQuery, "from", vbTextCompare) + 5
'    If pos1 > 5 Then
'        pos2 = InStr(pos1, strSQLQuery, " ")
'        If pos2 = 0 Then
'            pos2 = Len(strSQLQuery)
'        End If
'        aux = Mid(strSQLQuery, pos1, pos2 - pos1)
'        Flog.writeline Espacios(Tabulador * 4) & "Tabla: " & aux
'    End If
    Cantidad_de_OpenRecordset = Cantidad_de_OpenRecordset + 1
    
End Sub

Public Sub OpenRecordsetRH(strSQLQuery As String, ByRef objRs As ADODB.Recordset, Optional lockType As LockTypeEnum = adLockReadOnly)
    'Abre un recordset con la consulta strSQLQuery
    If objRs.State <> adStateClosed Then
        If objRs.lockType <> adLockReadOnly Then objRs.UpdateBatch
        objRs.Close
    End If
    objRs.Open strSQLQuery, objConn, adOpenDynamic, lockType, adCmdText
End Sub

Public Sub OpenRecordsetWithConn(strSQLQuery As String, ByRef objRs As ADODB.Recordset, ByRef Conn As ADODB.Connection, Optional lockType As LockTypeEnum = adLockReadOnly)
    'Abre un recordset con la consulta strSQLQuery, usando la conexion Conn
    If objRs.State <> adStateClosed Then
        If objRs.lockType <> adLockReadOnly Then objRs.UpdateBatch
        objRs.Close
    End If
    objRs.CacheSize = 500
    objRs.Open strSQLQuery, Conn, adOpenDynamic, lockType, adCmdText
End Sub

Public Sub MyBeginTransLiq()
    If Not TransactionRunning Then
            'FGZ - 04/06/2012 ----------------------------------------
            If TipoBD = 3 Then
                StrSql = "SET TRANSACTION ISOLATION LEVEL SNAPSHOT "
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
            'FGZ - 04/06/2012 ----------------------------------------
            objConn.BeginTrans
            
        TransactionRunning = True
    End If
End Sub
Public Sub MyCommitTransLiq()
    If TransactionRunning Then
        objConn.CommitTrans
        TransactionRunning = False
        'FGZ - 04/06/2012 ----------------------------------------
        If TipoBD = 3 Then
            StrSql = "SET TRANSACTION ISOLATION LEVEL READ COMMITTED "
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        'FGZ - 04/06/2012 ----------------------------------------
    End If
End Sub
Public Sub MyRollbackTransliq()
    If TransactionRunning Then
        objConn.RollbackTrans
        TransactionRunning = False
        'FGZ - 04/06/2012 ----------------------------------------
        If TipoBD = 3 Then
            StrSql = "SET TRANSACTION ISOLATION LEVEL READ COMMITTED "
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        'FGZ - 04/06/2012 ----------------------------------------
    End If
End Sub


Public Sub MyBeginTrans()
    If Not TransactionRunning Then
        objConn.BeginTrans
        TransactionRunning = True
    End If
End Sub
Public Sub MyCommitTrans()
    If TransactionRunning Then
        objConn.CommitTrans
        TransactionRunning = False
    End If
End Sub
Public Sub MyRollbackTrans()
    If TransactionRunning Then
        objConn.RollbackTrans
        TransactionRunning = False
    End If
End Sub




Public Sub BeginT()
    'If Not T2Running Then
    '    objconn2.BeginTrans
    '    T2Running = True
    'End If
End Sub

Public Sub CommitT()
    'If T2Running Then
    '    objconn2.CommitTrans
    '    T2Running = False
    'End If
End Sub

Public Sub RollbackT()
    'If T2Running Then
    '    objconn2.RollbackTrans
    '    T2Running = False
    'End If
End Sub

Public Function ConvFecha(ByVal dteFecha As Date) As String
    
    ConvFecha = "'" & Format(dteFecha, strformatoFservidor) & "'"
    
End Function

Public Function C_Date(ByVal Fecha) As Date
    'C_Date = Format(CDate(Fecha), strformatoFservidor)
    C_Date = Format(CDate(Fecha), "dd/mm/yyyy")
End Function


Public Function ConvNumero(ByVal Numero) As String
    If NumeroSeparadorDecimal = "," Then
        'elimino los puntos innecesarios (separador de miles)
        Numero = Replace(Numero, ".", "")
        'cambio coma por punto
        ConvNumero = "'" & Replace(Numero, ",", ".") & "'"
    End If
End Function


Public Function Pad(ByVal strTexto As String, bytLargo As Byte, Optional ByVal chrRelleno As String = "0") As String
    Dim intResto As Long
    Dim strResult As String

    intResto = bytLargo - Len(Trim(strTexto))
    If intResto > 0 Then Pad = String(intResto, chrRelleno) & Trim(strTexto) Else Pad = Trim(strTexto)
End Function

'EAM- Crea una vista temporal con el usuario que disparo el proceso y en el caso de Oracle una tabla temporal.
Public Sub CreaVistaEmpleadoProceso(ByVal NombreTabla As String, ByVal usuario As String)
 Dim objRs As New ADODB.Recordset
 Dim StrSql As String
 Dim Sql_VistaEmpleado As String
 On Error GoTo CE
    Select Case TipoBD
        Case 1 'db2
            StrSql = ""
        Case 2 ' Informix
            StrSql = ""
        Case 3 ' sql server
                StrSql = "SELECT view_definition FROM INFORMATION_SCHEMA.[VIEWS] v WHERE table_name= '" & UCase(NombreTabla) & "'"
                OpenRecordset StrSql, objRs
                
                If Not objRs.EOF Then
                    Sql_VistaEmpleado = Mid(objRs!view_definition, InStr(1, objRs!view_definition, "SELECT"), Len(objRs!view_definition))
                    'Sql_VistaEmpleado = Replace(Sql_VistaEmpleado, "SUSER_SNAME()", usuario, 1) mdf
                    Sql_VistaEmpleado = Replace(Sql_VistaEmpleado, "SUSER_SNAME()", "'" & usuario & "'", 1, -1, vbTextCompare)
                Else
                    Flog.writeline Espacios(0) & "No se puedo encontrar la Vista: "
                    Flog.writeline Espacios(0) & "SQL: " & StrSql
                    Exit Sub
                End If
            StrSql = "SELECT * FROM sys.views WHERE name = 'v_empleadoproc' AND schema_id = SCHEMA_ID('dbo')" 'mdf
            OpenRecordset StrSql, objRs
            If objRs.EOF Then 'mdf si la vista no existe se crea
                StrSql = "CREATE VIEW v_empleadoproc AS " & Sql_VistaEmpleado
                objConn.Execute StrSql
                
             Else 'mdf elimino y creo la vista en caso que ya existiera
                 StrSql = "DROP VIEW v_empleadoproc"
                 objConn.Execute StrSql
                 StrSql = "CREATE VIEW v_empleadoproc AS " & Sql_VistaEmpleado
                 objConn.Execute StrSql
                 Flog.writeline "La vista temporal ya existia. Se elimina y se vuelve a crear"
              End If
             objRs.Close
        Case 4 ' Oracle 9
            
            StrSql = "SELECT text FROM user_views WHERE view_name = '" & UCase(NombreTabla) & "'"
            OpenRecordset StrSql, objRs
            
            If Not objRs.EOF Then
                Sql_VistaEmpleado = objRs(0)
                Sql_VistaEmpleado = Replace(Sql_VistaEmpleado, "SUSER_SNAME()", usuario, 1)
            Else
                'Flog.writeline Espacios(Tabulador * 0) & "No se puedo encontrar la Vista: "
                'Flog.writeline Espacios(Tabulador * 0) & "SQL: " & StrSql
                Exit Sub
            End If
                      
                StrSql = "CREATE GLOBAL TEMPORARY TABLE v_empleadoproc " & _
                        " ON COMMIT PRESERVE ROWS " & _
                        " AS " & objRs(0)
                objConn.Execute StrSql
                objRs.Close
            
    End Select
    
    Flog.writeline Espacios(0) & "Utilizando seguridad por vistas."
    
Exit Sub
CE:
    Select Case TipoBD
        Case 1 'db2
            StrSql = ""
        Case 2 ' Informix
            StrSql = ""
        Case 3 ' sql server
            objConn.Execute "DROP VIEW v_empleadoproc"
            objConn.Execute StrSql
                        
        Case 4 ' Oracle 9
            'StrSql = "SELECT text FROM user_views WHERE view_name = '" & UCase(NombreTabla) & "'"
            'OpenRecordset StrSql, objRs
            'Flog.writeline Espacios(Tabulador * 0) & "Vista Utilizada: " & "DROP table v_empleadoproc"
            objConn.Execute "DROP table v_empleadoproc"
            objConn.Execute StrSql
    End Select
    Flog.writeline "Error:" & Err.Description
    Flog.writeline "Ultima SQL ejecuatada: " & StrSql
    Flog.writeline Espacios(0) & "Vista Utilizada por error: " & Sql_VistaEmpleado
    
    If objRs.State = adStateOpen Then objRs.Close
    Set objRs = Nothing
End Sub




Public Sub CreaVista(ByVal Nombre As String, ByVal SQLString As String)
Dim borra As String
    On Error GoTo CE
    If objConn.State = adStateOpen Then
        SQLString = "CREATE VIEW " & Nombre & " AS " & SQLString
        objConn.Execute SQLString
        Flog.writeline Espacios(0) & "crea Vista Utilizada: " & SQLString
    End If
Exit Sub
CE:
    Select Case Err.Number
    Case -2147217900
        ' la vista ya existe -> la borro
        objConn.Execute "DROP TABLE " & Nombre
        ' y la creo...
        objConn.Execute SQLString
    End Select
End Sub
Public Sub BorraVista(ByVal Nombre As String)
    On Error GoTo CE:
    If objConn.State = adStateOpen Then
        objConn.Execute "DROP VIEW " & Nombre
    End If
Exit Sub
CE:
    If Err.Number = -214721786 Then
        ' la vista no existe -> no la borro
    End If
End Sub

Public Function Encrypt(ByVal strEncryptionKey, ByVal strTextToEncrypt)
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


Public Function Decrypt(ByVal strEncryptionKey, ByVal strTextToEncrypt)
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
Dim buffer, outer, AuxI
    buffer = ""
    For outer = 1 To Len(strTextToEncrypt)
        AuxI = Hex(Asc(Mid(strTextToEncrypt, outer, 1)))
        If Len(AuxI) < 2 Then AuxI = "0" & AuxI
        buffer = buffer & AuxI
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


Public Function NOLOCK() As String
' ---------------------------------------------------------------------------------------------
' Descripcion: Retorna funcion nolock de acuerdo al tipo de base de datos
' Autor      : Martin
' Fecha      : 11/06/2009
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim aux As String

    Select Case TipoBD
        Case 1 'DB2
            aux = " "
        Case 2 'Informix
            aux = " "
        Case 3 'SQL Server
            aux = " WITH(NOLOCK) "
        Case 4 'Oracle
            aux = " "
        Case Else 'Default en SQL
            aux = " WITH(NOLOCK) "
    End Select

    NOLOCK = aux
    
End Function

Public Function TOP(ByVal cant As Long) As String
' ---------------------------------------------------------------------------------------------
' Descripcion: Retorna funcion nolock de acuerdo al tipo de base de datos
' Autor      : Martin
' Fecha      : 11/06/2009
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim aux As String

    Select Case TipoBD
        Case 1 'DB2
            aux = " "
        Case 2 'Informix
            aux = " "
        Case 3 'SQL Server
            aux = " TOP(" & cant & ") "
        Case 4 'Oracle
            aux = " TOP " & cant & " "
        Case Else 'Default en SQL
            aux = " TOP(" & cant & ") "
    End Select

    TOP = aux
    
End Function



