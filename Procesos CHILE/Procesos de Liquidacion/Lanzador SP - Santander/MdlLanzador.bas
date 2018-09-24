Attribute VB_Name = "MdlLanzador"
Option Explicit

Const Version = 1.01
Const FechaVersion = "29/02/2012" ' Juan A. Zamarbide
        

        
Global Mes As Integer
Global Anio As Integer
Global Empresa As String


Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial del Lanzador.
' Autor      : JAZ
' Fecha      : 29/12/2011
' Ultima Mod.:
' Descripcion: Procedimiento Inicial del Lanzador.
' ---------------------------------------------------------------------------------------------
Dim objconnMain As New ADODB.Connection
Dim strCmdLine
Dim Nombre_Arch As String
Dim HuboError As Boolean
Dim rs_batch_proceso As New ADODB.Recordset
Dim PID As String
Dim bprcparam As String
Dim ArrParametros

    strCmdLine = Command()
    ArrParametros = Split(strCmdLine, " ", -1)
    If UBound(ArrParametros) > 0 Then
        If IsNumeric(ArrParametros(0)) Then
            NroProcesoBatch = ArrParametros(0)
            Etiqueta = ArrParametros(1)
        Else
            Exit Sub
        End If
    Else
        If IsNumeric(strCmdLine) Then
            NroProcesoBatch = strCmdLine
        Else
            Exit Sub
        End If
    End If
    
    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas
    
On Error GoTo MAINE

    'Abro la conexion
    OpenConnection strconexion, objConn
    OpenConnection strconexion, objconnProgreso
    
    Nombre_Arch = PathFLog & "Lanzador_SP" & "-" & NroProcesoBatch & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Version = " & Version
    Flog.writeline "Fecha   = " & FechaVersion
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline
    Flog.writeline "PID = " & PID
    
    'Cambio el estado del proceso a Procesando y el PID
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 317 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        'IdUser = rs_batch_proceso!IdUser
        'Fecha = rs_batch_proceso!bprcfecha
        'hora = rs_batch_proceso!bprchora
        bprcparam = rs_batch_proceso!bprcparam
        
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        
        Call LevantarParametros(NroProcesoBatch, bprcparam)
    End If
    
    TiempoFinalProceso = GetTickCount
    Flog.writeline "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    
    If Not HuboError Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
    End If
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Flog.Close
    
    objConn.Close
    objconnProgreso.Close
    
    Exit Sub
    
MAINE:
    Flog.writeline "Error Main: " & Err.Description
    Flog.writeline "Ultimo sql Ejecutado: " & StrSql
    
End Sub


Public Sub Lanzador()
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de Lanzamiento del SP
' Autor      : JAZ
' Fecha      : 29/12/2011
' Ult. Mod   :
'
' --------------------------------------------------------------------------------------------

Dim objCmd As ADODB.Command


'Creo el Comando para ejecutar el SP
Set objCmd = New ADODB.Command
objCmd.CommandType = adCmdStoredProc
objCmd.CommandText = "Revisiones_RHPRO..sp_rel_AsientoContable" '"sp_rel_AsientoContable"
objCmd.CommandTimeout = 99999999
Set objCmd.ActiveConnection = objConn

Flog.writeline "Inicia Lanzador"

On Error GoTo CE

' Comienzo la transaccion
MyBeginTrans

StrSql = "UPDATE batch_proceso SET bprcprogreso = 10.0000  WHERE bpronro = " & NroProcesoBatch
objconnProgreso.Execute StrSql, , adExecuteNoRecords


'Inserto los Parámetros
objCmd.Parameters.Append objCmd.CreateParameter("@Mes", adInteger, adParamInput, 0, Mes)
objCmd.Parameters.Append objCmd.CreateParameter("@Anio", adInteger, adParamInput, 0, Anio)
objCmd.Parameters.Append objCmd.CreateParameter("@Empresa", adVarChar, adParamInput, 4, Empresa)
'Set Module1.backgroundWorker = background

'background.RunWorkerSync AddressOf BackgroundWork, 5000

objCmd.Execute

'background.CancelAsync
Flog.writeline "Terminó Lanzador"

StrSql = "UPDATE batch_proceso SET bprcprogreso = 100.0000  WHERE bpronro = " & NroProcesoBatch
objconnProgreso.Execute StrSql, , adExecuteNoRecords

'Fin de la transaccion
MyCommitTrans

'objCmd = Nothing



Exit Sub
CE:
    HuboError = True
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo sql Ejecutado: " & StrSql
    MyRollbackTrans

End Sub


Public Sub LevantarParametros(ByVal bpronro As Long, ByVal parametros As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento para levantar los parametros pasados en batch_proceso en bprcparam
' Autor      : JAZ
' Fecha      :
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim pos1 As Integer
Dim pos2 As Integer

Flog.writeline "Levanto los parámetros = " & parametros

' Levanto cada parametro por separado, el separador de parametros es "."
If Not IsNull(parametros) Then
        
    If Len(parametros) >= 1 Then
    
        Flog.writeline "Parametros: " & parametros
    
        pos1 = 1
        pos2 = InStr(pos1, parametros, ".") - 1
        Mes = CInt(Mid(parametros, pos1, pos2))
    
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, ".") - 1
        Anio = CInt(Mid(parametros, pos1, pos2 - pos1 + 1))
        
        pos1 = pos2 + 2
        pos2 = Len(parametros)
        Empresa = Mid(parametros, pos1, pos2 - pos1 + 1)
        
        
    End If
End If

Lanzador

End Sub

