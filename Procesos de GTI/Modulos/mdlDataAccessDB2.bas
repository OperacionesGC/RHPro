Attribute VB_Name = "mdlDataAccess"
Option Explicit

'  ---- para Informix
'Global Const strformatoFservidor As String = "DD/MM/YYYY"
'Global Const strformatoFservidor As String = "MM/DD/YYYY"
'Global Const strconexion = "DSN=informix"
'  ---- para Informix oledb
'Global Const strFormatoFServidor As String = "YYYY/MM/DD"
'Global Const strconexion = "Provider=Ifxoledbc;Password=rhpro;Persist Security Info=True;User ID=informix;Data Source=rhproexp@rhsco;"

'  ---- para DB2
'Global Const strformatoFservidor As String = "DD/MM/YYYY"
'Global Const strconexion = "dsn=db;database=;uid=;pwd="

' ----para sql server
'Global Const strformatoFservidor As String = "DD/MM/YYYY"
'Global Const strConexion = "Provider=sqloledb;server=DESARROLLO;database=rhpro;uid=sa;pwd="
'Global Const strformatoFservidor As String = "MM/DD/YYYY"
'Global Const strconexion = "dsn=rhpro;database=schneider;uid=sa;pwd="
'Global Const strconexion = "dsn=rhpro;database=rhpro;uid=sa;pwd="
' -------------------------------


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
                        ' Oracle9 = 4
Global PathFLog As String
' -------------------------------

' FGZ - 28/10/2003 Unificacion del .ini
Global PathSAP As String
Global PathProcesos As String

Global Const blnEsInformix As Boolean = True
Global objConn As New ADODB.Connection
Global objConnProgreso As New ADODB.Connection
Global objConnRhpro As New ADODB.Connection
Global CnTraza As New ADODB.Connection

Global objRs As New ADODB.Recordset
Global objRsAD As New ADODB.Recordset
Global TransactionRunning As Boolean
Global StrSql As String
Global StrSqlDatos As String
Global dummy As Long

' Variables que contienen los nombres de las tablas
' temporales de acuerdo a la DB
Global TTempWFDia As String
Global TTempWFDiaLaboral As String
Global TTempWFTurno As String
Global TTempWFEmbudo As String
Global TTempWFJustif As String
Global TTempWFAuspa As String
Global TTempWFAd As String
Global TTempLstaEmple As String
Global TTempWFLecturas As String
Global TTempWFInputFT As String

' -------- Variables de control de tiempo ------------
' minutos que espera el spool antes de abortar el proceso
Global TiempoDeEsperaNoResponde As Integer

' minutos que espera el spool antes de poner al proceso que se esta ejecutando en estado de No Responde
Global TiempoDeEsperaSinProgreso As Integer

' Tiempo entre lectura y lectura
Global TiempoDeLecturadeRegistraciones As Integer

' Tiempo de Dormida del Spool
Global TiempodeDormida As Integer

' Variable booleana que maneja si se usa Lectura de Registraciones o no
Global UsaLecturaRegistraciones As Boolean

'Maximo nro de Procesos Concurrentes
Global MaxConcurrentes As Integer

'FGZ - 19/05/2004
Global UltimaRegInsertadaWFTurno As String  '(N) - Ninguna, (E) - Entrada y (S) - Salida


'FGZ - 18/06/2004
Global Setea_Configuracion_Regional As Boolean
Global NumeroSeparadorDecimal As String
Global NumeroSeparadorMiles As String
Global MonedaSeparadorDecimal As String
Global MonedaSeparadorMiles As String
Global Nuevo_NumeroSeparadorDecimal As String
Global Nuevo_NumeroSeparadorMiles As String
Global Nuevo_MonedaSeparadorDecimal As String
Global Nuevo_MonedaSeparadorMiles As String
'FGZ - 06/02/2008 - se agregaron estas 2 variables
Global FormatoDeFechaCorto As String
Global Nuevo_FormatoDeFechaCorto As String
Global SCHEMA As String
Global PrimerConexion As Boolean

'FGZ - 04/02/2015 ----
Global UltimaRegHora As String
Global UltimaRegFecha As Date



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


Public Sub CargarConfiguracionesBasicas_old()
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
    If Not IsNull(Etiqueta) Then
        If Etiqueta <> "" Then
            Encontro = False
            Do While Not f.AtEndOfStream And Not Encontro
                strline = f.ReadLine()
                If InStr(1, UCase(strline), UCase(Etiqueta)) > 0 Then
                    Encontro = True
                End If
            Loop
        End If
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

    ' seteo de la etiqueta de BD a pasar como parametro a los procesos disparados
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        Etiqueta = Mid(strline, pos1, pos2 - pos1)
    End If
        
    f.Close

End Sub


Public Sub CargarConfiguracionesBasicasAppSrv()
' carga las configuraciones basicas para los procesos
Const ForReading = 1
Const ForAppending = 8

Dim f, fs
Dim strline As String
Dim pos1 As Integer
Dim pos2 As Integer

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

    ' seteo de la etiqueta de BD a pasar como parametro a los procesos disparados
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        If pos2 > pos1 Then
            Etiqueta = Mid(strline, pos1, pos2 - pos1)
        End If
    End If
        
    'FGZ - 23/06/2008 - Se agregó este parametro
    ' seteo del schema de Base de datos
    SCHEMA = ""
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        If pos2 >= pos1 Then
            SCHEMA = Mid(strline, pos1, pos2 - pos1)
        End If
    End If
        
    f.Close

End Sub

'Public Function getLastIdentity(ByRef objConn As ADODB.Connection, ByVal NombreTabla As String) As Variant
'    Dim objrs As New ADODB.Recordset
'
''
'    OpenRecordset StrSql, objrs
'    If Not (objrs.EOF And objrs.BOF) Then
'        If Not IsNull(objrs.Fields("next_id").Value) Then
'            getLastIdentity = objrs.Fields("next_id").Value
'        Else
'            getLastIdentity = -1
'        End If
'    Else
'        getLastIdentity = -1
'    End If
'End Function

Public Function getLastIdentity(ByRef objConn As ADODB.Connection, ByVal NombreTabla As String) As Variant
Dim objRs As New ADODB.Recordset
    
    Select Case TipoBD
    Case 1 'db2
        StrSql = "SELECT identity_val_local() as Codigo FROM sysibm.sysdummy1"
    Case 2 ' Informix
        StrSql = "select unique DBINFO('sqlca.sqlerrd1') as codigo from " & NombreTabla
    Case 3 ' sql server
            'StrSql = "SELECT @@IDENTITY as Codigo FROM " & NombreTabla & ""
            StrSql = "SELECT SCOPE_IDENTITY() as Codigo "
            'StrSql = "SELECT @@IDENTITY as Codigo"
    Case 4 ' Oracle 9
        StrSql = "select SEQ_" & UCase(NombreTabla) & ".CURRVAL as Codigo FROM DUAL"
    End Select
    OpenRecordset StrSql, objRs
    If Not (objRs.EOF And objRs.BOF) Then
        If Not IsNull(objRs.Fields("codigo").Value) Then
            getLastIdentity = objRs.Fields("codigo").Value
        Else
            getLastIdentity = -1
        End If
    Else
        getLastIdentity = -1
    End If
End Function

Public Function getLastIdentityMoño(ByRef objConn As ADODB.Connection, ByVal NombreTabla As String) As Variant
Dim objRs As New ADODB.Recordset

    
    'FGZ - 01/10/2003
    Flog.writeline "Entró en el GetLastIdentity"
    Flog.writeline objConn
    Flog.writeline NombreTabla
    
    Select Case TipoBD
    Case 1 'db2
        Flog.writeline "tipo de BD: DB2"
        StrSql = "SELECT crpnnro as codigo FROM " & NombreTabla & " ORDER BY codigo DESC"
    Case 2 ' Informix
        StrSql = "select unique DBINFO('sqlca.sqlerrd1') as codigo from " & NombreTabla
    Case 3 ' sql server
            StrSql = "SELECT @@IDENTITY as Codigo FROM " & NombreTabla & ""
            'StrSql = "SELECT @@IDENTITY as Codigo"
    End Select
    'FGZ - 01/10/2003
    Flog.writeline StrSql
    
    OpenRecordset StrSql, objRs
    If Not (objRs.EOF And objRs.BOF) Then
        Flog.writeline "encontro algo"
        If Not IsNull(objRs.Fields("codigo").Value) Then
            getLastIdentityMoño = objRs.Fields("codigo").Value
        Else
            Flog.writeline "el valor es nulo"
            getLastIdentityMoño = -1
        End If
    Else
        Flog.writeline "No encontró nada"
        getLastIdentityMoño = -1
    End If
End Function

Public Function getLastIdentityExpo(ByRef objConn As ADODB.Connection, ByVal NombreTabla As String) As Variant
Dim objRs As New ADODB.Recordset
    
    Select Case TipoBD
    Case 1 'db2
        StrSql = "SELECT identity_val_local() as Codigo FROM sysibm.sysdummy1"
    Case 2 ' Informix
        StrSql = "select unique DBINFO('sqlca.sqlerrd1') as codigo from " & NombreTabla
    Case 3 ' sql server
            StrSql = "SELECT @@IDENTITY as Codigo FROM " & NombreTabla & ""
            'StrSql = "SELECT @@IDENTITY as Codigo"
    Case 4 ' Oracle 9
        If NombreTabla = "expo_Lotes" Then
            StrSql = "select SEQ_NRO_LOTE.CURRVAL as Codigo FROM DUAL"
        Else
            StrSql = "select SEQ_" & UCase(NombreTabla) & ".CURRVAL as Codigo FROM DUAL"
        End If
    End Select
    'OpenRecordset StrSql, objRs
    OpenRecordsetNexus objConn, StrSql, objRs
    If Not (objRs.EOF And objRs.BOF) Then
        If Not IsNull(objRs.Fields("codigo").Value) Then
            getLastIdentityExpo = objRs.Fields("codigo").Value
        Else
            getLastIdentityExpo = -1
        End If
    Else
        getLastIdentityExpo = -1
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
    objConn.CommandTimeout = 120
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
    If EncriptStrconexion Then
        Flog.writeline "No se puede establecer conexion. Revise el string de conexion configurado en el .ini. Posible problema de encriptación."
    Else
        Flog.writeline "No se puede establecer conexion. Revise el string de conexion configurado en el .ini."
    End If
    Error_Encrypt = True
End Sub


Public Sub OpenRecordset(strSQLQuery As String, ByRef objRs As ADODB.Recordset, Optional lockType As LockTypeEnum = adLockReadOnly)
    'Abre un recordset con la consulta strSQLQuery
    If objRs.State <> adStateClosed Then
        If objRs.lockType <> adLockReadOnly Then objRs.UpdateBatch
        objRs.Close
    End If
    objRs.Open strSQLQuery, objConn, adOpenDynamic, lockType, adCmdText
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
Public Sub OpenRecordsetNexus(objConn As ADODB.Connection, strSQLQuery As String, ByRef objRs As ADODB.Recordset, Optional lockType As LockTypeEnum = adLockReadOnly)
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

Public Function ConvFecha(ByVal dteFecha As Date) As String
    
'    If blnEsInformix Then
        ConvFecha = "'" & Format(C_Date(dteFecha), strformatoFservidor) & "'"
 '   Else
  '      ConvFecha = "#" & Format(dteFecha, strFormatoFServidor) & "#"
  '  End If
End Function

Public Function ConvFecha2(ByVal dteFecha As Date) As String
'FGZ - 22/09/2003
If Not IsNull(dteFecha) Then
    ConvFecha2 = "'" & Format(dteFecha, strformatoFservidor) & "'"
Else
    ConvFecha2 = "' '"
End If
End Function

Public Function C_Date(ByVal Fecha) As Date
    'C_Date = Format(CDate(Fecha), strformatoFservidor)
    C_Date = Format(CDate(Fecha), "dd/mm/yyyy")
End Function


Public Function NuloaCero(ByVal nro) As Integer
'FGZ - 22/09/2003
If Not IsNull(nro) Then
    NuloaCero = nro
Else
    NuloaCero = 0
End If
End Function

Public Function NuloaVacio(ByVal X) As String
'FGZ - 22/09/2003
If Not IsNull(X) Then
    NuloaVacio = "'" & X & "'"
Else
    NuloaVacio = "' '"
End If
End Function


Public Function Pad(ByVal strTexto As String, bytLargo As Byte, Optional ByVal chrRelleno As String = "0") As String
    Dim intResto As Integer
    Dim strResult As String

    intResto = bytLargo - Len(Trim(strTexto))
    If intResto > 0 Then Pad = String(intResto, chrRelleno) & Trim(strTexto) Else Pad = Trim(strTexto)
End Function

Public Sub CreaVista(ByVal Nombre As String, ByVal SQLString As String)
Dim borra As String
    On Error GoTo ce
    If objConn.State = adStateOpen Then
        SQLString = "CREATE VIEW " & Nombre & " AS " & SQLString
        objConn.Execute SQLString
    End If
Exit Sub
ce:
    Select Case Err.Number
    Case -2147217900
        ' la vista ya existe -> la borro
        objConn.Execute "DROP TABLE " & Nombre
        ' y la creo...
        objConn.Execute SQLString
    End Select
End Sub
Public Sub BorraVista(ByVal Nombre As String)
    On Error GoTo ce:
    If objConn.State = adStateOpen Then
        objConn.Execute "DROP VIEW " & Nombre
    End If
Exit Sub
ce:
    If Err.Number = -214721786 Then
        ' la vista no existe -> no la borro
    End If
End Sub

'-------------------------------------------------
'Descripción: Procedimiento generico para Eliminar
'             las tablas temporales
'DE:          Nombre de la tabla temporal a Eliminar
'--------------------------------------------------

Public Sub BorrarTempTable(NombreTabla As String)
Dim Tabla As String

Select Case UCase(NombreTabla)
    Case "LSTA_EMPLE", "#LSTA_EMPLE":
        Tabla = TTempLstaEmple
    Case "WF_DIALABORAL", "#WF_DIALABORAL":
            Tabla = TTempWFDiaLaboral
    Case "WF_DIA", "#WF_DIA":
            Tabla = TTempWFDia
    Case "WF_EMBUDO", "#WF_EMBUDO":
            Tabla = TTempWFEmbudo
    Case "WF_TURNO", "#WF_TURNO":
            Tabla = TTempWFTurno
    Case "WF_JUSTIF", "#WF_JUSTIF":
            Tabla = TTempWFJustif
    Case "WF_AUSPA", "#WF_AUSPA":
            Tabla = TTempWFAuspa
    Case "WF_AD", "#WF_AD":
            Tabla = TTempWFAd
    Case "WF_LECTURAS", "#WF_LECTURAS":
            Tabla = TTempWFLecturas
    Case "WF_INPUTFT", "#WF_INPUTFT":
        Tabla = TTempWFInputFT
End Select
        
If TipoBD = 4 Then
    StrSql = "TRUNCATE TABLE " & Tabla
    objConn.Execute StrSql, , adExecuteNoRecords
Else
    StrSql = "DROP TABLE " & Tabla
    objConn.Execute StrSql, , adExecuteNoRecords
End If

End Sub

'-------------------------------------------------
'Descripción: Procedimiento generico para crear las
'               tablas temporales
'DE:          Nombre de la tabla temporal a crear
'--------------------------------------------------

Public Sub CreateTempTable(NombreTabla As String)
Dim cadena As String

On Error GoTo ce

Select Case TipoBD
Case 1: ' DB2
        Select Case UCase(NombreTabla)
        Case "LSTA_EMPLE", "#LSTA_EMPLE":
            cadena = TTempLstaEmple & "(empleg integer,ternro integer)"
        Case "WF_DIALABORAL", "#WF_DIALABORAL":
            cadena = TTempWFDiaLaboral & "(Codigo integer,Fecha_entrada date,Hora_entrada char(4),Fecha_salida date,Hora_salida char(4),entrada_salida integer, nrojustif integer)"
        Case "WF_DIA", "#WF_DIA":
            cadena = TTempWFDia & "(Codigo integer,Fecha date,Hora char(4),entrada integer)"
        Case "WF_EMBUDO", "#WF_EMBUDO":
            cadena = TTempWFEmbudo & "(Codigo integer,Fecha_desde date,Hora_desde char(4),Fecha_hasta date,Hora_hasta char(4))"
        Case "WF_TURNO", "#WF_TURNO":
            cadena = TTempWFTurno & "(ternro integer,Fecha date,Hora char(4),regnro integer,thnro integer,entrada integer,jusnro integer,par integer,evenro integer,anornro integer)"
        Case "WF_JUSTIF", "#WF_JUSTIF":
            cadena = TTempWFJustif & "(canths double,codext integer,desde date,hasta date,maxhoras double,orden integer,nro integer,sigla char(10),ternro integer,tjusnro integer" & _
                 ",horadesde char(4),horahasta char(4),thnro integer)"
        Case "WF_AUSPA", "#WF_AUSPA":
            cadena = TTempWFAuspa & "(fecha date,hora char(4),par integer)"
        Case "WF_AD", "#WF_AD":
            cadena = TTempWFAd & "(thnro integer,horas char(4),Cant_hs double,Acumula integer)"
        Case "WF_LECTURAS", "#WF_LECTURAS":
            cadena = TTempWFLecturas & "(ternro integer,fecha Date)"
        Case "WF_INPUTFT", "#WF_INPUTFT":
            cadena = TTempWFInputFT & "(idnro integer,idtipoinput integer, origen integer)"
        End Select
        
        StrSql = "CREATE TABLE " & cadena
        objConn.Execute StrSql, , adExecuteNoRecords
        Exit Sub
        
Case 2: ' Informix
        Select Case UCase(NombreTabla)
        Case "LSTA_EMPLE", "#LSTA_EMPLE":
            cadena = TTempLstaEmple & "(empleg integer,ternro integer)"
        Case "WF_DIALABORAL", "#WF_DIALABORAL":
            cadena = TTempWFDiaLaboral & "(Codigo integer,Fecha_entrada date,Hora_entrada char(4),Fecha_salida date,Hora_salida char(4),entrada_salida integer, nrojustif integer)"
        Case "WF_DIA", "#WF_DIA":
            cadena = TTempWFDia & "(Codigo integer,Fecha date,Hora char(4),entrada integer)"
        Case "WF_EMBUDO", "#WF_EMBUDO":
            cadena = TTempWFEmbudo & "(Codigo integer,Fecha_desde date,Hora_desde char(4),Fecha_hasta date,Hora_hasta char(4))"
        Case "WF_TURNO", "#WF_TURNO":
            cadena = TTempWFTurno & "(ternro integer,Fecha date,Hora char(4),regnro integer,thnro integer,entrada integer,jusnro integer,par integer,evenro integer,anornro integer)"
        Case "WF_JUSTIF", "#WF_JUSTIF":
            cadena = TTempWFJustif & "(canths decimal(5,2),codext integer,desde date,hasta date,maxhoras decimal(5,2),orden integer,nro integer,sigla char(10),ternro integer,tjusnro integer" & _
                 ",horadesde char(4),horahasta char(4),thnro integer)"
        Case "WF_AUSPA", "#WF_AUSPA":
            cadena = TTempWFAuspa & "(fecha date,hora char(4),par integer)"
        Case "WF_AD", "#WF_AD":
            cadena = TTempWFAd & "(thnro integer,horas char(10),Cant_hs decimal(5,2),Acumula integer)"
        Case "WF_LECTURAS", "#WF_LECTURAS":
            cadena = TTempWFLecturas & "(ternro integer,fecha Date)"
        Case "WF_INPUTFT", "#WF_INPUTFT":
            cadena = TTempWFInputFT & "(idnro integer,idtipoinput integer, origen integer)"
        End Select
        
        StrSql = "CREATE TEMP TABLE " & cadena
        objConn.Execute StrSql, , adExecuteNoRecords
        Exit Sub
        
Case 3: ' SQL Server
        Select Case UCase(NombreTabla)
        Case "LSTA_EMPLE", "#LSTA_EMPLE":
            cadena = TTempLstaEmple & "(empleg integer,ternro integer)"
        Case "WF_DIALABORAL", "#WF_DIALABORAL":
            cadena = TTempWFDiaLaboral & "(Codigo integer,Fecha_entrada datetime,Hora_entrada char(4),Fecha_salida datetime,Hora_salida char(4),entrada_salida integer, nrojustif integer)"
        Case "WF_DIA", "#WF_DIA":
            cadena = TTempWFDia & "(Codigo integer,Fecha datetime,Hora char(4),entrada integer)"
        Case "WF_EMBUDO", "#WF_EMBUDO":
            cadena = TTempWFEmbudo & "(Codigo integer,Fecha_desde datetime,Hora_desde char(4),Fecha_hasta datetime,Hora_hasta char(4))"
        Case "WF_TURNO", "#WF_TURNO":
            cadena = TTempWFTurno & "(ternro integer,Fecha datetime,Hora char(4),regnro integer,thnro integer,entrada integer,jusnro integer,par integer,evenro integer,anornro integer)"
        Case "WF_JUSTIF", "#WF_JUSTIF":
            cadena = TTempWFJustif & "(canths decimal(5,2),codext integer,desde datetime,hasta datetime,maxhoras decimal(5,2),orden integer,nro integer,sigla char(10),ternro integer,tjusnro integer" & _
                            ",horadesde char(4),horahasta char(4),thnro integer)"
        Case "WF_AUSPA", "#WF_AUSPA":
            cadena = TTempWFAuspa & "(fecha datetime,hora char(4),par integer)"
        Case "WF_AD", "#WF_AD":
            cadena = TTempWFAd & "(thnro integer, horas char(10),Cant_hs decimal(5,2),Acumula integer)"
        Case "WF_LECTURAS", "#WF_LECTURAS":
            cadena = TTempWFLecturas & "(ternro integer,fecha Datetime)"
        Case "WF_INPUTFT", "#WF_INPUTFT":
            cadena = TTempWFInputFT & "(idnro integer,idtipoinput integer, origen integer)"
        End Select
        
        StrSql = "CREATE TABLE " & cadena
        objConn.Execute StrSql, , adExecuteNoRecords
        Exit Sub
Case 4: ' Oracle 9
        Select Case UCase(NombreTabla)
        Case "LSTA_EMPLE", "#LSTA_EMPLE":
            cadena = TTempLstaEmple & "(empleg double precision,ternro double precision)"
        Case "WF_DIALABORAL", "#WF_DIALABORAL":
            cadena = TTempWFDiaLaboral & "(Codigo double precision,Fecha_entrada date,Hora_entrada char(4),Fecha_salida date,Hora_salida char(4),entrada_salida double precision, nrojustif double precision)"
        Case "WF_DIA", "#WF_DIA":
            cadena = TTempWFDia & "(Codigo double precision,Fecha date,Hora char(4),entrada double precision)"
        Case "WF_EMBUDO", "#WF_EMBUDO":
            cadena = TTempWFEmbudo & "(Codigo double precision,Fecha_desde date,Hora_desde char(4),Fecha_hasta date,Hora_hasta char(4))"
        Case "WF_TURNO", "#WF_TURNO":
            cadena = TTempWFTurno & "(ternro double precision,Fecha date,Hora char(4),regnro double precision,thnro double precision,entrada double precision,jusnro double precision,par double precision,evenro double precision,anornro double precision)"
        Case "WF_JUSTIF", "#WF_JUSTIF":
            cadena = TTempWFJustif & "(canths FLOAT(63),codext double precision,desde date,hasta date,maxhoras FLOAT(63),orden double precision,nro double precision,sigla char(10),ternro double precision,tjusnro double precision" & _
                            ",horadesde char(4),horahasta char(4),thnro double precision)"
        Case "WF_AUSPA", "#WF_AUSPA":
            cadena = TTempWFAuspa & "(fecha date,hora char(4),par double precision)"
        Case "WF_AD", "#WF_AD":
            cadena = TTempWFAd & "(thnro double precision,horas char(10),Cant_hs FLOAT(63),Acumula double precision)"
        Case "WF_LECTURAS", "#WF_LECTURAS":
            cadena = TTempWFLecturas & "(ternro double precision,fecha Date)"
        Case "WF_INPUTFT", "#WF_INPUTFT":
            cadena = TTempWFInputFT & "(idnro double precision,idtipoinput double precision, origen double precision)"
        End Select
        
        StrSql = "CREATE GLOBAL TEMPORARY TABLE " & cadena & " ON COMMIT PRESERVE ROWS"
        objConn.Execute StrSql, , adExecuteNoRecords
        Exit Sub
End Select

ce:
    If TipoBD = 4 Then
        Call BorrarTempTable(NombreTabla)
    Else
        Call BorrarTempTable(NombreTabla)
        Call CreateTempTable(NombreTabla)
    End If

End Sub

Public Sub InsertarWF_Lecturas(ByVal Tercero As Long, ByVal Fecha As Date)
    StrSql = "INSERT INTO " & TTempWFLecturas & "(ternro,fecha) VALUES (" & _
             Tercero & "," & ConvFecha(Fecha) & ")"
    
    objConn.Execute StrSql, , adExecuteNoRecords
End Sub



Public Sub InsertarWFDia(Codigo As Integer, Hora As String, Fecha As Date, Optional entrada As Boolean = False)
    Select Case TipoBD
    Case 1: ' DB2
        StrSql = "INSERT INTO " & TTempWFDia & "(Codigo,Fecha,Hora,Entrada) VALUES (" & _
             Codigo & "," & ConvFecha(Fecha) & ",'" & Hora & "'," & CInt(entrada) & ")"
    Case 2: ' Informix
         StrSql = "INSERT INTO " & TTempWFDia & "(Codigo,Fecha,Hora,Entrada) VALUES (" & _
             Codigo & "," & ConvFecha(Fecha) & ",'" & Hora & "'," & CInt(entrada) & ")"
    Case 3: ' SQL Server
        StrSql = "INSERT INTO " & TTempWFDia & "(Codigo,Fecha,Hora,Entrada) VALUES (" & _
             Codigo & "," & ConvFecha(Fecha) & ",'" & Hora & "'," & CInt(entrada) & ")"
    Case 4: ' Oracle 9
        StrSql = "INSERT INTO " & TTempWFDia & "(Codigo,Fecha,Hora,Entrada) VALUES (" & _
             Codigo & "," & ConvFecha(Fecha) & ",'" & Hora & "'," & CInt(entrada) & ")"
    End Select
    
    objConn.Execute StrSql, , adExecuteNoRecords
    
End Sub


Public Sub InsertarWFDiaLaboral(Codigo As Long, Fecha_entrada As Date, Fecha_Salida As Date, hora_entrada As String, hora_salida As String, entrada_salida As Integer, Optional NroJustif As Long = 1)
     
    Select Case TipoBD
    Case 1: ' DB2
        StrSql = "INSERT INTO " & TTempWFDiaLaboral & "(codigo,Fecha_entrada,Fecha_Salida,hora_entrada,hora_salida,entrada_salida,NroJustif) VALUES (" & _
              Codigo & "," & ConvFecha(Fecha_entrada) & "," & ConvFecha(Fecha_Salida) & ",'" & hora_entrada & "','" & hora_salida & "'," & CInt(entrada_salida) & "," & NroJustif & ")"
    Case 2: ' Informix
         StrSql = "INSERT INTO " & TTempWFDiaLaboral & "(codigo,Fecha_entrada,Fecha_Salida,hora_entrada,hora_salida,entrada_salida,NroJustif) VALUES (" & _
              Codigo & "," & ConvFecha(Fecha_entrada) & "," & ConvFecha(Fecha_Salida) & ",'" & hora_entrada & "','" & hora_salida & "'," & CInt(entrada_salida) & "," & NroJustif & ")"
    Case 3: ' SQL Server
        StrSql = "INSERT INTO " & TTempWFDiaLaboral & "(codigo,Fecha_entrada,Fecha_Salida,hora_entrada,hora_salida,entrada_salida,NroJustif) VALUES (" & _
              Codigo & "," & ConvFecha(Fecha_entrada) & "," & ConvFecha(Fecha_Salida) & ",'" & hora_entrada & "','" & hora_salida & "'," & CInt(entrada_salida) & "," & NroJustif & ")"
    Case 4: ' Oracle 9
        StrSql = "INSERT INTO " & TTempWFDiaLaboral & "(codigo,Fecha_entrada,Fecha_Salida,hora_entrada,hora_salida,entrada_salida,NroJustif) VALUES (" & _
              Codigo & "," & ConvFecha(Fecha_entrada) & "," & ConvFecha(Fecha_Salida) & ",'" & hora_entrada & "','" & hora_salida & "'," & CInt(entrada_salida) & "," & NroJustif & ")"
    End Select
     
    objConn.Execute StrSql, , adExecuteNoRecords
    
End Sub


Public Sub InsertarWFEmbudo(Codigo As Long, fecha_desde As Date, hora_desde As String, fecha_hasta As Date, hora_hasta As String)
    
    Select Case TipoBD
    Case 1: ' DB2
        StrSql = "INSERT INTO " & TTempWFEmbudo & "(Codigo,Fecha_desde,Hora_desde,Fecha_hasta,Hora_hasta) VALUES (" & _
             Codigo & "," & ConvFecha(fecha_desde) & ",'" & hora_desde & "'," & ConvFecha(fecha_hasta) & ",'" & hora_hasta & "')"
    Case 2: ' Informix
         StrSql = "INSERT INTO " & TTempWFEmbudo & "(Codigo,Fecha_desde,Hora_desde,Fecha_hasta,Hora_hasta) VALUES (" & _
             Codigo & "," & ConvFecha(fecha_desde) & ",'" & hora_desde & "'," & ConvFecha(fecha_hasta) & ",'" & hora_hasta & "')"
    Case 3: ' SQL Server
        StrSql = "INSERT INTO " & TTempWFEmbudo & "(Codigo,Fecha_desde,Hora_desde,Fecha_hasta,Hora_hasta) VALUES (" & _
             Codigo & "," & ConvFecha(fecha_desde) & ",'" & hora_desde & "'," & ConvFecha(fecha_hasta) & ",'" & hora_hasta & "')"
    Case 4: ' Oracle 9
        StrSql = "INSERT INTO " & TTempWFEmbudo & "(Codigo,Fecha_desde,Hora_desde,Fecha_hasta,Hora_hasta) VALUES (" & _
             Codigo & "," & ConvFecha(fecha_desde) & ",'" & hora_desde & "'," & ConvFecha(fecha_hasta) & ",'" & hora_hasta & "')"
    End Select
    
    objConn.Execute StrSql, , adExecuteNoRecords
End Sub


Public Sub InsertarWFTurno(Ternro As Long, Hora As String, Fecha As Date, Regnro As Long, thnro As Integer, anornro As Integer, entrada As Boolean, jusnro As Long, par As Integer, evenro As Integer)
    
    ' FGZ  - 19/05/2004
    Select Case UltimaRegInsertadaWFTurno
        Case "N":   entrada = True
                    UltimaRegInsertadaWFTurno = "E"
        Case "E":   entrada = False
                    UltimaRegInsertadaWFTurno = "S"
        Case "S":   entrada = True
                    UltimaRegInsertadaWFTurno = "E"
    End Select
    
    Select Case TipoBD
    Case 1: ' DB2
        StrSql = "INSERT INTO " & TTempWFTurno & "(Ternro,Fecha,Hora,regnro,thnro,anornro,Entrada,jusnro,par,evenro) VALUES (" & _
             Ternro & "," & ConvFecha(Fecha) & ",'" & Hora & "'," & Regnro & "," & thnro & "," & anornro & "," & entrada & "," & jusnro & "," & par & "," & evenro & ")"
    Case 2: ' Informix
         StrSql = "INSERT INTO " & TTempWFTurno & "(Ternro,Fecha,Hora,regnro,thnro,anornro,Entrada,jusnro,par,evenro) VALUES (" & _
             Ternro & "," & ConvFecha(Fecha) & ",'" & Hora & "'," & Regnro & "," & thnro & "," & anornro & "," & entrada & "," & jusnro & "," & par & "," & evenro & ")"
    Case 3: ' SQL Server
        StrSql = "INSERT INTO " & TTempWFTurno & "(Ternro,Fecha,Hora,regnro,thnro,anornro,Entrada,jusnro,par,evenro) VALUES (" & _
             Ternro & "," & ConvFecha(Fecha) & ",'" & Hora & "'," & Regnro & "," & thnro & "," & anornro & "," & entrada & "," & jusnro & "," & par & "," & evenro & ")"
    Case 4: ' Oracle 9
        StrSql = "INSERT INTO " & TTempWFTurno & "(Ternro,Fecha,Hora,regnro,thnro,anornro,Entrada,jusnro,par,evenro) VALUES (" & _
             Ternro & "," & ConvFecha(Fecha) & ",'" & Hora & "'," & Regnro & "," & thnro & "," & anornro & "," & entrada & "," & jusnro & "," & par & "," & evenro & ")"
    End Select
    
    objConn.Execute StrSql, , adExecuteNoRecords
End Sub

'--------------------------------------------------------------
' Procedimiento que setea los nombres de las tablas temporales
' de acuerdo al tipo de DB
' -------------------------------------------------------------
Public Sub CargarNombresTablasTemporales()

Select Case TipoBD
    Case 1: ' DB2
            TTempWFDia = "WF_Dia"
            TTempWFDiaLaboral = "WF_Dialaboral"
            TTempWFTurno = "WF_Turno"
            TTempWFEmbudo = "WF_Embudo"
            TTempWFJustif = "WF_Justif"
            TTempWFAuspa = "WF_Auspa"
            TTempWFAd = "WF_Ad"
            TTempLstaEmple = "Lsta_Emple"
            TTempWFLecturas = "WF_Lecturas"
            TTempWFInputFT = "WF_INPUTFT"
    Case 2: ' Informix
            TTempWFDia = "WF_Dia"
            TTempWFDiaLaboral = "WF_Dialaboral"
            TTempWFTurno = "WF_Turno"
            TTempWFEmbudo = "WF_Embudo"
            TTempWFJustif = "WF_Justif"
            TTempWFAuspa = "WF_Auspa"
            TTempWFAd = "WF_Ad"
            TTempLstaEmple = "Lsta_Emple"
            TTempWFLecturas = "WF_Lecturas"
            TTempWFInputFT = "WF_INPUTFT"
    Case 3: ' SQL Server
            TTempWFDia = "#WF_Dia"
            TTempWFDiaLaboral = "#WF_Dialaboral"
            TTempWFTurno = "#WF_Turno"
            TTempWFEmbudo = "#WF_Embudo"
            TTempWFJustif = "#WF_Justif"
            TTempWFAuspa = "#WF_Auspa"
            TTempWFAd = "#WF_Ad"
            TTempLstaEmple = "#Lsta_Emple"
            TTempWFLecturas = "#WF_Lecturas"
            TTempWFInputFT = "#WF_INPUTFT"
    Case 4: ' Oracle 9
            TTempWFDia = "WF_Dia"
            TTempWFDiaLaboral = "WF_Dialaboral"
            TTempWFTurno = "WF_Turno"
            TTempWFEmbudo = "WF_Embudo"
            TTempWFJustif = "WF_Justif"
            TTempWFAuspa = "WF_Auspa"
            TTempWFAd = "WF_Ad"
            TTempLstaEmple = "Lsta_Emple"
            TTempWFLecturas = "WF_Lecturas"
            TTempWFInputFT = "WF_INPUTFT"
    End Select
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

'EAM- Carga las configuraciones basicas para los procesos SrvDefaults (v6.00)
Public Sub SetarDefaultsReducido()
 Const ForReading = 1

 Dim f, fs
 Dim strline As String
 Dim pos1 As Integer
 Dim pos2 As Integer

    On Error GoTo ME_CargaDef

    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.OpenTextFile(App.Path & "\rhproappsrvDefaults.ini", ForReading, 0)
    
    'Minutos que espera el spool antes de abortar el proceso
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        TiempoDeEsperaNoResponde = Mid(strline, pos1, pos2 - pos1)
    End If
        
    GoTo datosOK
    
ME_CargaDef:
    Flog.writeline "No se pudo leer el archivo de configuracion (" & App.Path & "\RHProappSrvDefaults.ini)."
    TiempoDeEsperaNoResponde = 5
    Flog.writeline
    Exit Sub
datosOK:
    f.Close
End Sub

