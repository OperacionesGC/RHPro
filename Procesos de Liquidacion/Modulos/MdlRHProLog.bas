Attribute VB_Name = "MdlRHProLog"
Option Explicit
' ---------------------------------------------------------------------------------------------
' Descripcion: Modulo oara creacion de Logs.
' Autor      : Mauricio Zwenger
' Fecha      : 14/12/2016
' ---------------------------------------------------------------------------------------------

Public Sub createRHProLog(ByVal appId As String, ByVal bpronro As Long, ByVal archivo As String, ByRef FLog As Variant)

    On Error GoTo 0
    
    Dim StrSql As String
    Dim rs As New ADODB.Recordset
    Dim Nombre_Arch As String
    Dim cnString As String
    Dim Base As String
    Dim tipo_proceso As Long
   
    Nombre_Arch = PathFLog & "bpronro_" & bpronro & ".log"
    
    If bpronro = CLng(-1) Then
        'bpronro    = -1    ==> Planificador
        Nombre_Arch = PathFLog & "planificador " & Format(Date, "dd-mm-yyyy") & ".log"
        tipo_proceso = -1
    ElseIf bpronro = CLng(0) Then
        'bpronro    = 0     ==> AppServer
        Nombre_Arch = PathFLog & "RHProAppSrv " & Format(Date, "dd-mm-yyyy") & ".log"
        tipo_proceso = 0
    Else
        StrSql = "select btprarchlog, batch_proceso.btprcnro from batch_proceso inner join batch_tipproc on batch_proceso.btprcnro=batch_tipproc.btprcnro "
        StrSql = StrSql & " WHERE bpronro=" & bpronro
        OpenRecordset StrSql, rs
        If Not rs.EOF Then
            Nombre_Arch = PathFLog & rs("btprarchlog") & bpronro & ".log"
            tipo_proceso = rs("btprcnro")
        Else
            rs.Close
            Exit Sub
        End If
        rs.Close
    End If
    
    StrSql = "SELECT cnstring, cndesc FROM confper "
    StrSql = StrSql & " INNER JOIN conexion ON confper.confint = conexion.cnnro "
    StrSql = StrSql & " WHERE confnro=39 and confactivo=-1"
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        cnString = rs.Fields.Item("cnString")
        If InStr(cnString, "Provider") > 0 Then
            Dim aa, aux As String, i As Long
            aa = Split(cnString, ";")
            For i = 0 To UBound(aa)
                If InStr(aa(i), "Provider") = 0 Then
                    aux = aux & aa(i) & ";"
                End If
            Next
            cnString = aux
        End If
        
        Base = rs.Fields.Item("cndesc")
        rs.Close
    Else
        rs.Close
        Exit Sub
    End If
    
    'elimino el archivo de log generado hasta aca
    FLog.Close
    Set FLog = Nothing
    
    Dim fslog As Object
    Set fslog = CreateObject("Scripting.filesystemobject")
    If fslog.FileExists(archivo) Then
        fslog.DeleteFile archivo, True
    End If
    

    Set FLog = New Logger
    
    Dim cmLog As RHProLog.logConnManager
    Set cmLog = New RHProLog.logConnManager
    cmLog.addConnection 0, 1, cnString
    cmLog.addConnection 5, 0, Nombre_Arch
    If bpronro <= 0 Then
        FLog.Append = True
    End If
    FLog.Name = Nombre_Arch
 
    Call FLog.Connect(cmLog, "rhpro", Base, appId, tipo_proceso, bpronro, 100)
    
       
End Sub




