Attribute VB_Name = "mdlDataAccess"
Option Explicit

' definicion de variables globales de configuracion basica
Global strformatoFservidor As String
Global strconexion As String
Global TipoBD As String
                        ' DB2 = 1
                        ' Informix = 2
                        ' SQL Server = 3
Global PathFLog As String

'FGZ - 18/06/2004
Global NumeroSeparadorDecimal As String
Global NumeroSeparadorMiles As String
Global MonedaSeparadorDecimal As String
Global MonedaSeparadorMiles As String
Global Nuevo_NumeroSeparadorDecimal As String
Global Nuevo_NumeroSeparadorMiles As String
Global Nuevo_MonedaSeparadorDecimal As String
Global Nuevo_MonedaSeparadorMiles As String

' -------------------------------
' FGZ - 28/10/2003 Unificacion del .ini
Global PathSAP As String
Global PathProcesos As String

Global objConn As New ADODB.Connection
Global CnTraza As New ADODB.Connection
Global objRs As New ADODB.Recordset
Global dummy As Long
Global TransactionRunning As Boolean

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


Public Sub CargarConfiguracionesBasicas_old2()
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


Public Sub CargarConfiguracionesBasicas_old_old()
' carga las configuraciones basicas para los procesos
Const ForReading = 1
Const ForAppending = 8

Dim f, fs
Dim strline As String
Dim pos1 As Integer
Dim pos2 As Integer

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

Public Function getLastIdentity(ByRef objConn As ADODB.Connection, ByVal NombreTabla As String) As Variant
Dim objRs As New ADODB.Recordset
    Dim StrSql As String
    Select Case TipoBD
    Case 1 'db2
        StrSql = "SELECT identity_val_local() as Codigo FROM sysibm.sysdummy1"
    Case 2 ' Informix
        StrSql = "select unique DBINFO('sqlca.sqlerrd1') as codigo from " & NombreTabla
    Case 3 ' sql server
            StrSql = "SELECT @@IDENTITY as Codigo FROM " & NombreTabla & ""
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
Public Function Str2SQLField(ByVal cadena As String) As String
    'Transforma una cadena de caracteres en compatible para SQL (transforma el apostrofo por un acento franc�s)
    'PARA VALIDAR ENTRADAS DE CAMPOS
         
    Str2SQLField = Replace(cadena, "'", "`")
End Function
Public Sub OpenConnection(strConnectionString As String, ByRef objConn As ADODB.Connection)
    If objConn.State <> adStateClosed Then objConn.Close
    'objConn.CursorLocation = adUseServer
    objConn.CursorLocation = adUseClient
    
    'objConn.IsolationLevel = adXactCursorStability
    'Indica que desde una transacci�n se pueden ver cambios que no se han producido
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
Dim pos1 As Integer
Dim pos2 As Integer
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
    
    ConvFecha = "'" & Format(dteFecha, strformatoFservidor) & "'"
    
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
    Dim intResto As Integer
    Dim strResult As String

    intResto = bytLargo - Len(Trim(strTexto))
    If intResto > 0 Then Pad = String(intResto, chrRelleno) & Trim(strTexto) Else Pad = Trim(strTexto)
End Function

Public Sub CreaVista(ByVal nombre As String, ByVal SQLString As String)
Dim borra As String
    On Error GoTo CE
    If objConn.State = adStateOpen Then
        SQLString = "CREATE VIEW " & nombre & " AS " & SQLString
        objConn.Execute SQLString
    End If
Exit Sub
CE:
    Select Case Err.Number
    Case -2147217900
        ' la vista ya existe -> la borro
        objConn.Execute "DROP TABLE " & nombre
        ' y la creo...
        objConn.Execute SQLString
    End Select
End Sub
Public Sub BorraVista(ByVal nombre As String)
    On Error GoTo CE:
    If objConn.State = adStateOpen Then
        objConn.Execute "DROP VIEW " & nombre
    End If
Exit Sub
CE:
    If Err.Number = -214721786 Then
        ' la vista no existe -> no la borro
    End If
End Sub
