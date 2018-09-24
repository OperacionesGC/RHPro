Attribute VB_Name = "MdlBajaMasivaProc"
Option Explicit

Global Const Version = "1.00"
Global Const FechaVersion = "23/10/2010"
Global Const UltimaModificacion = ""
Global Const UltimaModificacion1 = ""



Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial.
' Autor      : Martin Ferraro
' Fecha      : 23/09/2010
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim objconnMain As New ADODB.Connection
Dim strCmdLine
Dim Nombre_Arch As String
Dim rs_batch_proceso As New ADODB.Recordset
Dim PID As String
Dim idUser As String
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
    
    Nombre_Arch = PathFLog & "BajaMasivaProc-" & NroProcesoBatch & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaVersion
    Flog.writeline "Modificacion = " & UltimaModificacion
    Flog.writeline "Modificacion = " & UltimaModificacion1
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "-------------------------------------------------"
    Flog.writeline
    Flog.writeline

    On Error Resume Next
    'Abro la conexion
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
     Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
     Exit Sub
    End If
        
    On Error Resume Next
    OpenConnection strconexion, objconnProgreso
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    On Error GoTo 0
    
    On Error GoTo ME_Main

    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0 ,bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Flog.writeline "Pone el estado en procesando"
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 273 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        bprcparam = rs_batch_proceso!bprcparam
        idUser = rs_batch_proceso!idUser
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        Call GenerarBajaMasiva(NroProcesoBatch, bprcparam, idUser)
    Else
        Flog.writeline "no encontró el proceso"
    End If
    
    TiempoFinalProceso = GetTickCount
    Flog.writeline "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    
    If Not HuboError Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
    End If
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
Fin:
    Flog.Close
    If objconnProgreso.State = adStateOpen Then objconnProgreso.Close
Exit Sub

ME_Main:
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
        Flog.writeline Espacios(Tabulador * 0) & " Error General " & Err.Description
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "SQL Ejecutado: " & StrSql
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
        Flog.writeline
        
    'Actualizo el progreso
    MyBeginTrans
        StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Error General', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
    MyCommitTrans
End Sub


Public Sub GenerarBajaMasiva(ByVal bpronro As Long, ByVal parametros As String, ByVal idUser As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion de procesos de baja masiva
' Autor      : Martin Ferraro
' Fecha      : 28/09/2010
' --------------------------------------------------------------------------------------------
Dim listaProc As String
Dim arrProc
Dim indProc As Long
Dim paramProc As String
Dim codBproNro As Long


'RecordSet
Dim rs_Empleados As New ADODB.Recordset
Dim rs_Consult As New ADODB.Recordset

'Variables

' Inicio codigo ejecutable
On Error GoTo CE

' Levanto cada parametro por separado, el separador de parametros es "@"
Flog.writeline Espacios(Tabulador * 0) & "levantando parametros" & parametros
If Not IsNull(parametros) Then
    listaProc = parametros
Else
    Flog.writeline Espacios(Tabulador * 0) & "ERROR. No se encuentran los paramentros."
    Exit Sub
End If
Flog.writeline

arrProc = Split(listaProc, ",")

CEmpleadosAProc = UBound(arrProc) + 1

'seteo de las variables de progreso
Progreso = 0
If CEmpleadosAProc = 0 Then
   Flog.writeline "no hay procesos"
   CEmpleadosAProc = 1
End If
IncPorc = (100 / CEmpleadosAProc)
        
Flog.writeline
Flog.writeline
Flog.writeline
Flog.writeline Espacios(Tabulador * 0) & "--------------------------------------------------------"
Flog.writeline Espacios(Tabulador * 0) & "Comienza a procesar los Procesos de liquidacion"
Flog.writeline Espacios(Tabulador * 0) & "--------------------------------------------------------"
Flog.writeline


'Comienzo a procesar los empleados
For indProc = 0 To UBound(arrProc)
    
    MyBeginTrans
    
    Flog.writeline Espacios(Tabulador * 1) & "Analizando Proceso " & arrProc(indProc)
    
    paramProc = arrProc(indProc)
    paramProc = paramProc & ".0"  'l_mantliq
    paramProc = paramProc & ".0"  'l_guardarnov
    paramProc = paramProc & ".0"  'l_anadet
    paramProc = paramProc & ".0"  'l_todos
    paramProc = paramProc & ".0"  'l_liqhis
    paramProc = paramProc & ".-1" 'l_borrarliq
    paramProc = paramProc & ".0"  'l_usadebug
    paramProc = paramProc & ".-1" 'l_BorrarProc
    
    StrSql = generarSQLProc(3, paramProc, idUser)
    objConn.Execute StrSql, , adExecuteNoRecords
    codBproNro = getLastIdentity(objConn, "batch_proceso")
    
    StrSql = "SELECT empleado FROM cabliq WHERE pronro = " & arrProc(indProc)
    OpenRecordset StrSql, rs_Empleados
    Do While Not rs_Empleados.EOF
        
       StrSql = "INSERT INTO batch_empleado"
       StrSql = StrSql & " (bpronro, ternro, estado)"
       StrSql = StrSql & " VALUES (" & codBproNro & "," & rs_Empleados!Empleado & ",null)"
       objConn.Execute StrSql, , adExecuteNoRecords
       
       rs_Empleados.MoveNext
    Loop
    rs_Empleados.Close
    
    MyCommitTrans
    
    'Actualizo el progreso
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
             "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
Next


    

If rs_Empleados.State = adStateOpen Then rs_Empleados.Close
If rs_Consult.State = adStateOpen Then rs_Consult.Close
Set rs_Empleados = Nothing
Set rs_Consult = Nothing

Exit Sub

CE:
    Flog.writeline "=================================================================="
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="
    
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
             "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    HuboError = True

End Sub


Function generarSQLProc(ByVal tipoPorc As Long, ByVal parametros As String, ByVal id As String)

Dim hora As String
Dim Dia As String
Dim sqlp As String

hora = Mid(Time, 1, 8)
Dia = ConvFecha(Date)

sqlp = " INSERT INTO batch_proceso "
sqlp = sqlp & " (btprcnro, bprcfecha, iduser, bprchora, bprcfecdesde, bprcfechasta, bprcparam, "
sqlp = sqlp & " bprcestado, bprcprogreso, bprcfecfin, bprchorafin, bprctiempo, empnro, bprcempleados,bprcurgente) "
sqlp = sqlp & " VALUES (" & tipoPorc & "," & Dia & ", '" & id & "','" & hora & "' "
sqlp = sqlp & " ,null,null"
sqlp = sqlp & " , '" & parametros & "', 'Pendiente', null , null, null, null, 0, null,0)"

generarSQLProc = sqlp

End Function 'generarSQLProc()
