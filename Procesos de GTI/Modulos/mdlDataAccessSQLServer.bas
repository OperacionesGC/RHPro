Attribute VB_Name = "mdlDataAccess"
Option Explicit


'Global Const strFormatoFServidor As String = "DD/MM/YYYY"
Global Const strFormatoFServidor As String = "MM/DD/YYYY"
Global Const blnEsInformix As Boolean = True
Global objConn As New ADODB.Connection
Global CnTraza As New ADODB.Connection
Global objRs As New ADODB.Recordset
Global TransactionRunning As Boolean
Global StrSql As String
Global dummy As Long

Public Function getLastIdentity(ByRef objConn As ADODB.Connection, ByVal NombreTabla As String) As Variant
Dim objRs As New ADODB.Recordset
    Dim StrSql As String
    StrSql = "SELECT @@IDENTITY as Codigo FROM " & NombreTabla & ""
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
Public Function Str2SQLField(ByVal Cadena As String) As String
    'Transforma una cadena de caracteres en compatible para SQL (transforma el apostrofo por un acento francés)
    'PARA VALIDAR ENTRADAS DE CAMPOS
        
    Str2SQLField = Replace(Cadena, "'", "`")
End Function
Public Sub OpenConnection(strConnectionString As String, ByRef objConn As ADODB.Connection)
    If objConn.State <> adStateClosed Then objConn.Close
    'objConn.CursorLocation = adUseServer
    objConn.CursorLocation = adUseClient
    'objconn.IsolationLevel =
    objConn.IsolationLevel = adXactCursorStability
    objConn.CommandTimeout = 60
    objConn.ConnectionTimeout = 60
    objConn.Open strConnectionString
'    objConn.Execute "SET LOCK MODE TO WAIT 60"
End Sub
Public Sub OpenRecordset(strSQLQuery As String, ByRef objRs As ADODB.Recordset, Optional lockType As LockTypeEnum = adLockReadOnly)
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
    
'    If blnEsInformix Then
        ConvFecha = "'" & Format(dteFecha, strFormatoFServidor) & "'"
 '   Else
  '      ConvFecha = "#" & Format(dteFecha, strFormatoFServidor) & "#"
  '  End If
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
