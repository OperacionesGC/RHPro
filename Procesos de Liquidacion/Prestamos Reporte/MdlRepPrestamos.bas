Attribute VB_Name = "MdlRepPrestamos"
Option Explicit

'Const Version = 1.1 ' Gustavo Ring - Reporte Estado de Préstamos
'Const FechaVersion = "27/12/2006"
Const tiprep_nro = 151

Global Const Version = 1.2
Global Const FechaVersion = "21/10/2009"
Global Const UltimaModificacion = "Encriptacion de conexion"
Global Const UltimaModificacion1 = "Manuel Lopez"

'------------------------------------------------------------
'------------------------------------------------------------
'------------------------------------------------------------
Global IdUser As String
Global Fecha As Date
Global Hora As String

Global Aux_Autoriz_Apenom As String
Global Aux_Autoriz_Docu As String
Global Aux_Autoriz_Prov_Emis As String

Global Aux_Certifi_Corresponde As String
Global Aux_Certifi_Doc_Tipo As String
Global Aux_Certifi_Doc_Nro As String
Global Aux_Certifi_Expedida As String



Public Sub Main()
    ' ---------------------------------------------------------------------------------------------
    ' Descripcion: Procedimiento inicial del Generador de Reportes.
    ' Autor      : FGZ
    ' Fecha      : 17/02/2004
    ' Ultima Mod.:
    ' Descripcion:
    ' ---------------------------------------------------------------------------------------------
    Dim objconnMain As New ADODB.Connection
    Dim strCmdLine
    Dim Nombre_Arch As String
    Dim HuboError As Boolean
    Dim rs_batch_proceso As New ADODB.Recordset
    Dim bprcparam As String
    Dim PID As String
    Dim ArrParametros
    
    strCmdLine = Command()
    'strCmdLine = "10602"
    
    
    strCmdLine = Command()
    ArrParametros = Split(strCmdLine, " ", -1)
    If UBound(ArrParametros) > 1 Then
        If IsNumeric(ArrParametros(0)) Then
            NroProcesoBatch = ArrParametros(0)
            Etiqueta = ArrParametros(1)
            EncriptStrconexion = CBool(ArrParametros(2))
            c_seed = ArrParametros(2)
        Else
            Exit Sub
        End If
    Else
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
    End If
    
    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas

    
    'Abro la conexion
    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    
    On Error Resume Next
    OpenConnection strconexion, CnTraza
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    
    Nombre_Arch = PathFLog & "Reporte_prestamos" & "-" & NroProcesoBatch & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Version        = " & Version
    Flog.writeline "Fecha          = " & FechaVersion
    Flog.writeline "Modificacion   = " & UltimaModificacion
    Flog.writeline "                 " & UltimaModificacion1
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline
    Flog.writeline "PID = " & PID
    
    
    'Cambio el estado del proceso a Procesando y el PID
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objConn.Execute StrSql, , adExecuteNoRecords
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE (btprcnro = 151 ) AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        IdUser = rs_batch_proceso!IdUser
        Fecha = rs_batch_proceso!bprcfecha
        Hora = rs_batch_proceso!bprchora
        bprcparam = rs_batch_proceso!bprcparam
        
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        
        'Call Generar_Reporte(NroProcesoBatch, aux_sucursal, aux_mes, aux_anio, aux_empr)
        Call LevantarParamteros(NroProcesoBatch, bprcparam)
    Else
        Flog.writeline "-----------------------------------------------------------------"
        Flog.writeline "No existen empleados para procesar"
        Flog.writeline "-----------------------------------------------------------------"
    End If
    
    TiempoFinalProceso = GetTickCount
    Flog.writeline "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    
    If Not HuboError Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
    End If
    objConn.Execute StrSql, , adExecuteNoRecords

    Flog.Close
    objConn.Close
    
End Sub

Public Sub Generar_Reporte(ByVal bpronro As Long, ByVal mes_desde As Integer, ByVal Anio_desde As Integer, ByVal mes_hasta As Integer, ByVal anio_hasta As Integer, ByVal titulo As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del Reporte estado de préstamos
' Autor      : Gustavo Ring
' Fecha      : 19/12/2006
' Ult. Mod   :
'
' --------------------------------------------------------------------------------------------

'Variables auxiliares

'Parametros
Dim aux_sucursal As String
Dim aux_mes As Integer
Dim aux_anio As Integer
Dim aux_periodo As String

Dim HFecha As Date

'Variables
Dim Aux_Repnro As String
Dim Aux_Empleado_Nombre As String
Dim Aux_Empleado_Ternro As String
Dim Aux_Periodo_Mesdesde As Integer
Dim aux_periodo_Meshasta As Integer
Dim aux_periodo_Aniodesde As Integer
Dim aux_periodo_Aniohasta As Integer
Dim Aux_Empleado_Empleg As Integer

'Cuotas
Dim aux

'Registros
Dim rs_Empleados As New ADODB.Recordset
Dim rs_prestamos As New ADODB.Recordset
Dim rs_prestamos_det As New ADODB.Recordset
Dim rs_batch_empleado As New ADODB.Recordset
Dim rs_aux As New ADODB.Recordset

On Error GoTo CE

' Comienzo la transaccion
    'MyBeginTrans
    
    StrSql = "SELECT batch_empleado.ternro,empleado.empleg from batch_empleado inner join empleado on empleado.ternro = batch_empleado.ternro where batch_empleado.bpronro = " & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_empleado
    
    If rs_batch_empleado.EOF Then
        StrSql = "SELECT ternro,empleg from empleado"
        OpenRecordset StrSql, rs_batch_empleado
    End If
    'Seteo el progreso
    Progreso = 0
    CEmpleadosAProc = rs_batch_empleado.RecordCount
    If CEmpleadosAProc = 0 Then
        Flog.writeline "------------------------------"
        Flog.writeline "No hay empleados que procesar "
        Flog.writeline "------------------------------"
        CEmpleadosAProc = 1
    End If
    IncPorc = (100 / CEmpleadosAProc)
    
'Loopeo por cada empleado en la tabla batch_empleado con el ID del proceso

'mes_desde = 1
'Anio_desde = 2005

'mes_desde = 6
'Anio_desde = 2005


Do While Not rs_batch_empleado.EOF

    MyBeginTrans
    
    Flog.writeline "Procesando empleado id " & rs_batch_empleado!ternro
    Flog.writeline "-----------------------------------------------------------------"
        
    'Traigo los valores del empleado
    StrSql = "select * from prestamo inner join pre_cuota on pre_cuota.prenro=prestamo.prenro "
    StrSql = StrSql & " inner join empleado on prestamo.ternro=empleado.ternro "
    StrSql = StrSql & " where empleado.ternro = " & rs_batch_empleado!ternro
    
    OpenRecordset StrSql, rs_Empleados
        
    Aux_Empleado_Nombre = ""
    aux_periodo = ""
    
    If Not rs_Empleados.EOF Then
         Aux_Empleado_Ternro = rs_batch_empleado!ternro
         If Not IsNull(rs_Empleados!terape) Then Aux_Empleado_Nombre = rs_Empleados!terape & " " & rs_Empleados!terape2 & " " & rs_Empleados!ternom & " " & rs_Empleados!ternom2
         If Not IsNull(rs_Empleados!empleg) Then Aux_Empleado_Empleg = rs_Empleados!empleg
    End If

    'Aux_Repnro = getLastIdentity(objConn, "rep_prestamos")
    
    StrSql = "select max(repnro) as c from rep_prestamos"
    OpenRecordset StrSql, rs_aux
        
    If rs_aux.EOF Or IsNull(rs_aux!c) Then
          Aux_Repnro = 1
    Else
          Aux_Repnro = rs_aux!c + 1
    End If
    
    
    'Aux_Repnro = getLastIdentity(objConn, "rep_prestamos")
        
    StrSql = "SELECT prestamo.prenro,cuomes,cuoano,cuocancela,cuonrocuo,cuototal,predesc FROM prestamo "
    StrSql = StrSql & " INNER JOIN pre_cuota on prestamo.prenro=pre_cuota.prenro "
    StrSql = StrSql & " WHERE ternro= " & rs_batch_empleado!ternro
    StrSql = StrSql & " and (((pre_cuota.cuomes >= " & mes_desde & " and pre_cuota.cuoano >= " & Anio_desde & ") "
    StrSql = StrSql & " and (pre_cuota.cuomes <= " & mes_hasta & " and pre_cuota.cuoano <= " & anio_hasta & ")) "
    StrSql = StrSql & " or ((pre_cuota.cuoano > " & Anio_desde & ") and (pre_cuota.cuomes <= " & mes_hasta & " and pre_cuota.cuoano <= " & anio_hasta
    StrSql = StrSql & " or (pre_cuota.cuoano > " & Anio_desde & " and pre_cuota.cuoano < " & anio_hasta & ")"
    StrSql = StrSql & "))) ORDER BY prestamo.prenro,cuonrocuo"
    
    OpenRecordset StrSql, rs_prestamos
    
    StrSql = "INSERT INTO rep_prestamos "
    StrSql = StrSql & "(bpronro,fecha, hora, ternro,nombres,periodo,repnro,empleg,prodesc)"
    StrSql = StrSql & " VALUES ("
    StrSql = StrSql & "'" & NroProcesoBatch & "',"
    StrSql = StrSql & "'" & Date & "',"
    StrSql = StrSql & "'" & Time & "',"
    StrSql = StrSql & "'" & rs_batch_empleado!ternro & "',"
    StrSql = StrSql & "'" & Aux_Empleado_Nombre & "',"
    StrSql = StrSql & "'" & aux_periodo & "',"
    StrSql = StrSql & "'" & Aux_Repnro & "',"
    StrSql = StrSql & "'" & Aux_Empleado_Empleg & "',"
    StrSql = StrSql & "'" & titulo & "')"
  
    
    If Not rs_prestamos.EOF Then
        objConn.Execute StrSql, , adExecuteNoRecords
    Else
        Flog.writeline "Procesando empleado id " & rs_batch_empleado!ternro & " no tiene prestamos asociados"
        Flog.writeline "-------------------------------------------------------------------------------------"
    End If
    
    If rs_Empleados.State = adStateOpen Then rs_Empleados.Close
                    
    
    Dim Aux_Prestamo_Prenro As Integer
    Dim Aux_Prestamo_Cuomes As Integer
    Dim Aux_Prestamo_Cuoano As Integer
    Dim Aux_Prestamo_Cuocancela As Boolean
    Dim Aux_Prestamo_Cuonrocuo As Integer
    Dim Aux_Prestamo_Cuoimp As Double
    Dim Aux_prestamo_predesc As String
    
    Dim Aux_Estado As Integer
    
    'loop de cuotas
    
    
    Do While Not rs_prestamos.EOF
        If Not IsNull(rs_prestamos!prenro) Then Aux_Prestamo_Prenro = rs_prestamos!prenro
        If Not IsNull(rs_prestamos!predesc) Then Aux_prestamo_predesc = rs_prestamos!predesc
        If Not IsNull(rs_prestamos!cuomes) Then Aux_Prestamo_Cuomes = rs_prestamos!cuomes
        If Not IsNull(rs_prestamos!cuoano) Then Aux_Prestamo_Cuoano = rs_prestamos!cuoano
        If Not IsNull(rs_prestamos!cuocancela) Then Aux_Prestamo_Cuocancela = rs_prestamos!cuocancela
        If Not IsNull(rs_prestamos!cuonrocuo) Then Aux_Prestamo_Cuonrocuo = rs_prestamos!cuonrocuo
        If Not IsNull(rs_prestamos!cuototal) Then Aux_Prestamo_Cuoimp = rs_prestamos!cuototal
        
        If Aux_Prestamo_Cuocancela Then
             Aux_Estado = -1
        Else
            Aux_Estado = 0
        End If
        
        rs_prestamos.MoveNext
        
        Flog.writeline "Procesando cuota id " & Aux_Prestamo_Cuonrocuo & " del prestamo " & Aux_Prestamo_Prenro & " de " & Aux_Prestamo_Cuomes & "/" & Aux_Prestamo_Cuoano
        Flog.writeline "---------------------------------------------------------------------------------------"
    
 
        StrSql = "INSERT INTO rep_prestamos_det "
        StrSql = StrSql & "(repnro,prenro,cuomes,cuoanio,cuocancela,cuonrocuo,cuoimp,predesc)"
        StrSql = StrSql & " VALUES ("
        StrSql = StrSql & "'" & Aux_Repnro & "',"
        StrSql = StrSql & "'" & Aux_Prestamo_Prenro & "',"
        StrSql = StrSql & "'" & Aux_Prestamo_Cuomes & "',"
        StrSql = StrSql & "'" & Aux_Prestamo_Cuoano & "',"
        StrSql = StrSql & "'" & Aux_Estado & "',"
        StrSql = StrSql & "'" & Aux_Prestamo_Cuonrocuo & "',"
        StrSql = StrSql & "'" & Aux_Prestamo_Cuoimp & "',"
        StrSql = StrSql & "'" & Aux_prestamo_predesc & "')"
        
        objConn.Execute StrSql, , adExecuteNoRecords
        
        If rs_Empleados.State = adStateOpen Then rs_prestamos.Close
        
    Loop
    If rs_prestamos.State = adStateOpen Then rs_prestamos.Close
    
    rs_batch_empleado.MoveNext
        
    If Not HuboError Then
            StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
    Else
            StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
    End If
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
                 ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
                 "' WHERE bpronro = " & NroProcesoBatch
    CnTraza.Execute StrSql, , adExecuteNoRecords
    MyCommitTrans
Loop

If rs_batch_empleado.State = adStateOpen Then rs_batch_empleado.Close

'Fin de la transaccion
'MyCommitTrans

Set rs_Empleados = Nothing
Set rs_batch_empleado = Nothing

Exit Sub
CE:
    HuboError = True
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo sql Ejecutado: " & StrSql
    MyRollbackTrans
End Sub

Public Sub LevantarParamteros(ByVal bpronro As Long, ByVal parametros As String)
' ----------------------------------------------------------------------------------------------
' Descripcion: Procedimiento para levantar los parametros pasados en batch_proceso en bprcparam
' Autor      : FGZ
' Fecha      :
' Ult. Mod   :
' Fecha      :
'-----------------------------------------------------------------------------------------------
Dim pos1 As Integer
Dim pos2 As Integer
Dim pos3 As Integer
Dim pos4 As Integer
Dim aux_sucursal
Dim aux_mes_desde
Dim aux_anio_desde
Dim aux_mes_hasta
Dim aux_anio_hasta
Dim aux_titulo
Dim ArrParametros

Dim aux As String

Dim HFecha As Date
Dim Aux_Separador As String

Aux_Separador = "@"
' Levanto cada parametro por separado, el separador de parametros es Aux_Separador

If Not IsNull(parametros) Then
    If Len(parametros) >= 1 Then
        ArrParametros = Split(parametros, Aux_Separador, -1)
        aux_mes_desde = ArrParametros(0)
        aux_anio_desde = ArrParametros(1)
        aux_mes_hasta = ArrParametros(2)
        aux_anio_hasta = ArrParametros(3)
        aux_titulo = Left(ArrParametros(4), 200)
    End If
End If

'Reporte Individual del personal
Call Generar_Reporte(bpronro, aux_mes_desde, aux_anio_desde, aux_mes_hasta, aux_anio_hasta, aux_titulo)
End Sub





