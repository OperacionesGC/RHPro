Attribute VB_Name = "MdlDistribucion"
Option Explicit

Global Const Version = 1
Global Const FechaVersion = "05/08/2009"   'Encriptacion de string connection
Global Const UltimaModificacion = "Manuel Lopez"
Global Const UltimaModificacion1 = "Encriptacion de string connection"

'Global inx             As Integer
'Global inxfin          As Integer
'
'Global vec_testr1(50)  As Integer
'Global vec_testr2(50)  As Integer
'Global vec_testr3(50)  As Integer
'
'Global vec_jor(50) As Single
'
'Global Descripcion As String
'Global Cantidad As Single
'
'
'Global rs_Proc_Vol As New ADODB.Recordset
'Global rs_Mod_Linea As New ADODB.Recordset
'Global rs_Empleado As New ADODB.Recordset
'Global rs_Mod_Asiento As New ADODB.Recordset
'
'Global BUF_mod_linea As New ADODB.Recordset
'Global BUF_temp As New ADODB.Recordset

Global CantidadEmpleados As Long
Global PrimeraVez As Boolean


Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial.
' Autor      : FGZ
' Fecha      : 16/01/2003
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim objconnMain As New ADODB.Connection
Dim strCmdLine
Dim Nombre_Arch As String
Dim HuboError As Boolean
Dim rs_Batch_Proceso As New ADODB.Recordset
Dim PID As String
Dim bprcparam As String
Dim ArrParametros

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
    
    Nombre_Arch = PathFLog & "Asiento_Contable" & "-" & NroProcesoBatch & ".log"
    
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
    
    'Abro la conexion
    On Error Resume Next
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
    
    Nombre_Arch = PathFLog & "Distribucion_Costos" & "-" & NroProcesoBatch & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline Espacios(Tabulador * 0) & "PID = " & PID
    TiempoInicialProceso = GetTickCount
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & ",bprctiempo = 0 WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 76 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_Batch_Proceso
    
    If Not rs_Batch_Proceso.EOF Then
        bprcparam = rs_Batch_Proceso!bprcparam
        rs_Batch_Proceso.Close
        Set rs_Batch_Proceso = Nothing
        Call LevantarParamteros(NroProcesoBatch, bprcparam)
    End If
    
    Flog.writeline Espacios(Tabulador * 0) & "Actualizo estado del proceso"
    TiempoFinalProceso = GetTickCount
    If Not HuboError Then
        StrSql = "UPDATE batch_proceso SET bprctiempo = " & (TiempoFinalProceso - TiempoInicialProceso) & ", bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado', bprcprogreso = 100 WHERE bpronro = " & NroProcesoBatch
    Else
        StrSql = "UPDATE batch_proceso SET bprctiempo = " & (TiempoFinalProceso - TiempoInicialProceso) & ", bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto', bprcprogreso = 100 WHERE bpronro = " & NroProcesoBatch
    End If
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Flog.writeline Espacios(Tabulador * 0) & "Fin proceso"
    objConn.Close
    objconnProgreso.Close
    Flog.Close

End Sub

Public Sub LevantarParamteros(ByVal bpronro As Long, ByVal parametros As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento para levantar los parametros pasados en batch_proceso en bprcparam
' Autor      : FGZ
' Fecha      :
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim pos1 As Integer
Dim pos2 As Integer
Dim Separador As String

Dim FechaDesde As Date
Dim FechaHasta As Date
Dim Empresa As Long

Flog.writeline Espacios(Tabulador * 1) & "Levanto los parametros " & parametros
'Orden de los parametros

Separador = "@"
' Levanto cada parametro por separado
If Not IsNull(parametros) Then
    If Len(parametros) >= 1 Then
        pos1 = 1
        pos2 = InStr(pos1, parametros, Separador) - 1
        NroProc = Mid(parametros, pos1, pos2 - pos1 + 1)
        ListaNroProc = Replace(NroProc, "-", ",")
        
        pos1 = pos2 + 2
        pos2 = Len(parametros)
        Empresa = Mid(parametros, pos1, pos2 - pos1 + 1)
    End If
End If

Call Generacion_Distribucion(bpronro, ListaNroProc, Empresa)

End Sub


Public Sub Generacion_Distribucion(ByVal bpronro As Long, ByVal Lista_Procesos As String, ByVal Empresa As Long)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento principal de Distribucion de Costos
' Autor      : FGZ
' Fecha      : 10/01/2005
' --------------------------------------------------------------------------------------------
Dim Cantidad_Cabeceras As Long
Dim Ultimo_Empleado As Long
Dim Tenro1 As Long
Dim Tenro2 As Long
Dim Tenro3 As Long

Dim LogicaDistribucion As Integer
Dim Razon_MesCalendario As Boolean  'False ==> Meses de 30 dias
                                    'True  ==> Meses Normales
Dim rs_Cabeceras As New ADODB.Recordset

Flog.writeline Espacios(Tabulador * 1) & "Generacion de la Distribucion de Costos para los procesos " & Lista_Procesos

StrSql = "SELECT * FROM  proceso "
StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
StrSql = StrSql & " WHERE proceso.pronro in ( " & Lista_Procesos & ")"
StrSql = StrSql & " ORDER BY cabliq.empleado, proceso.pronro, cabliq.cliqnro "
OpenRecordset StrSql, rs_Cabeceras

Cantidad_Cabeceras = IIf(rs_Cabeceras.RecordCount <> 0, rs_Cabeceras.RecordCount, 1)
IncPorc = 99 / Cantidad_Cabeceras
Progreso = 0

Ultimo_Empleado = -1
LogicaDistribucion = 0
Do While Not rs_Cabeceras.EOF
        If Ultimo_Empleado <> rs_Cabeceras!Empleado Then
            LogicaDistribucion = 0
            Flog.writeline Espacios(Tabulador * 1) & "Tercero " & rs_Cabeceras!Empleado
            Flog.writeline Espacios(Tabulador * 2) & "Busco el alcance y Resuelvo el tipo de Distribucion a aplicar"
            'Busco el alcance y Resuelvo el tipo de Distribucion a aplicar
            Call Resolver_Alcenace(rs_Cabeceras!Empleado, rs_Cabeceras!profecini, rs_Cabeceras!profecfin, LogicaDistribucion, Razon_MesCalendario, Tenro1, Tenro2, Tenro3)
            Ultimo_Empleado = rs_Cabeceras!Empleado
        End If
        
        Select Case LogicaDistribucion
        Case 1: 'Distribucion en base a historico de estructuras
            Flog.writeline Espacios(Tabulador * 2) & "Distribucion en base a historico de estructuras"
            Call Distrubucion_Historica(rs_Cabeceras!Pronro, rs_Cabeceras!cliqnro, rs_Cabeceras!Empleado, rs_Cabeceras!profecini, rs_Cabeceras!profecfin, Tenro1, Tenro2, Tenro3, Razon_MesCalendario)
        Case 2: 'Distribucion en base a porcentajes fijos
            Flog.writeline Espacios(Tabulador * 2) & "Distribucion en base a porcentajes fijos"
            'Call Distrubucion_Historica(rs_Cabeceras!Pronro, rs_Cabeceras!cliqnro, rs_Cabeceras!Empleado, rs_Cabeceras!profecini, rs_Cabeceras!profecfin, Tenro1, Tenro2, Tenro3, Razon_MesCalendario)
        Case 3: 'Distribucion en base a Medida de GTI
            Flog.writeline Espacios(Tabulador * 2) & "Distribucion en base a Medida de GTI"
            'Call Distrubucion_Historica(rs_Cabeceras!Pronro, rs_Cabeceras!cliqnro, rs_Cabeceras!Empleado, rs_Cabeceras!profecini, rs_Cabeceras!profecfin, Tenro1, Tenro2, Tenro3, Razon_MesCalendario)
        Case Else   'Error
            Flog.writeline Espacios(Tabulador * 2) & "No se encontro alcance de distribucion"
            'No se encontro alcance de distribucion
        End Select

        Flog.writeline Espacios(Tabulador * 2) & "Actualizar el progreso"
        'Actualizar el progreso
        TiempoFinalProceso = GetTickCount
        Progreso = Progreso + IncPorc
        StrSql = "UPDATE batch_proceso SET bprctiempo = " & (TiempoFinalProceso - TiempoInicialProceso) & ", bprcprogreso = " & Progreso & " WHERE bpronro = " & NroProcesoBatch
        objConn.Execute StrSql, , adExecuteNoRecords
    
    rs_Cabeceras.MoveNext
Loop

If rs_Cabeceras.State = adStateOpen Then rs_Cabeceras.Close

Set rs_Cabeceras = Nothing

Exit Sub

CE:
    MyRollbackTrans
    HuboError = True
    Flog.writeline " Error: " & Err.Description
End Sub


Public Sub Resolver_Alcenace(ByVal ternro As Long, ByVal FechaDesde As Date, ByVal FechaHasta As Date, ByRef LogicaDistribucion As Integer, ByRef Razon_MesCalendario As Boolean, ByRef Tipo1 As Long, ByRef Tipo2 As Long, ByRef Tipo3 As Long)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que resuelve el tipo de alcance de Distribucion Contable.
' Autor      : FGZ
' Fecha      : 10/01/2005
' --------------------------------------------------------------------------------------------
Dim rs_Alcance As New ADODB.Recordset
Dim rs_TipoDistribucion As New ADODB.Recordset


'Busco alcance Individual
Flog.writeline Espacios(Tabulador * 3) & "Busco alcance Individual"
StrSql = "SELECT * FROM tdist_ind "
StrSql = StrSql & " WHERE ternro = " & ternro
OpenRecordset StrSql, rs_Alcance

If Not rs_Alcance.EOF Then
    'Busco el tipo de distribucion
    Flog.writeline Espacios(Tabulador * 3) & "Busco el tipo de distribucion"
    StrSql = "SELECT * FROM tipodistrib "
    StrSql = StrSql & " WHERE tdistnro = " & rs_Alcance!tdistnro
    OpenRecordset StrSql, rs_TipoDistribucion
    
    If Not rs_TipoDistribucion.EOF Then
        Razon_MesCalendario = CBool(rs_TipoDistribucion!tdistcalend)
        LogicaDistribucion = rs_TipoDistribucion!tdistlogica
        Tipo1 = IIf(Not EsNulo(rs_TipoDistribucion!tdistTenro1), rs_TipoDistribucion!tdistTenro1, 0)
        Tipo2 = IIf(Not EsNulo(rs_TipoDistribucion!tdistTenro2), rs_TipoDistribucion!tdistTenro2, 0)
        Tipo3 = IIf(Not EsNulo(rs_TipoDistribucion!tdistTenro3), rs_TipoDistribucion!tdistTenro3, 0)
        
        If Razon_MesCalendario Then
            Flog.writeline Espacios(Tabulador * 4) & "Razon Mes Calendario "
        Else
            Flog.writeline Espacios(Tabulador * 4) & "Razon de 30 d�as "
        End If
        Flog.writeline Espacios(Tabulador * 4) & "Logica de Distribuci�n " & LogicaDistribucion
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de estructura 1 " & Tipo1
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de estructura 2 " & Tipo2
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de estructura 3 " & Tipo3
    Else
        LogicaDistribucion = 0
        Razon_MesCalendario = False
        Flog.writeline Espacios(Tabulador * 4) & "No se encontro tipo de alcance de distribucion " & rs_Alcance!idtipodist
    End If
Else
    'Busco alcance por estructura
    Flog.writeline Espacios(Tabulador * 3) & "Busco alcance por estructura"
    StrSql = "SELECT * FROM tdist_est "
    StrSql = StrSql & " INNER JOIN his_estructura ON tdist_est.tdestrestrnro = his_estructura.estrnro "
    StrSql = StrSql & " WHERE his_estructura.htetdesde <= " & ConvFecha(FechaHasta)
    StrSql = StrSql & " AND (his_estructura.htethasta >= " & ConvFecha(FechaDesde)
    StrSql = StrSql & " OR his_estructura.htethasta IS NULL)"
    StrSql = StrSql & " AND his_estructura.ternro =" & ternro
    StrSql = StrSql & " ORDER BY tdist_est.tdestrorden "
    OpenRecordset StrSql, rs_Alcance
    If Not rs_Alcance.EOF Then
        'Busco el tipo de distribucion
        Flog.writeline Espacios(Tabulador * 3) & "Busco el tipo de distribucion"
        StrSql = "SELECT * FROM tipodistrib "
        StrSql = StrSql & " WHERE tdistnro = " & rs_Alcance!tdistnro
        OpenRecordset StrSql, rs_TipoDistribucion
        
        If Not rs_TipoDistribucion.EOF Then
            Razon_MesCalendario = CBool(rs_TipoDistribucion!tdistcalend)
            LogicaDistribucion = rs_TipoDistribucion!tdistlogica
            Tipo1 = IIf(Not EsNulo(rs_TipoDistribucion!tdistTenro1), rs_TipoDistribucion!tdistTenro1, 0)
            Tipo2 = IIf(Not EsNulo(rs_TipoDistribucion!tdistTenro2), rs_TipoDistribucion!tdistTenro2, 0)
            Tipo3 = IIf(Not EsNulo(rs_TipoDistribucion!tdistTenro3), rs_TipoDistribucion!tdistTenro3, 0)
            
            If Razon_MesCalendario Then
                Flog.writeline Espacios(Tabulador * 4) & "Razon Mes Calendario "
            Else
                Flog.writeline Espacios(Tabulador * 4) & "Razon de 30 d�as "
            End If
            Flog.writeline Espacios(Tabulador * 4) & "Logica de Distribuci�n " & LogicaDistribucion
            Flog.writeline Espacios(Tabulador * 4) & "Tipo de estructura 1 " & Tipo1
            Flog.writeline Espacios(Tabulador * 4) & "Tipo de estructura 2 " & Tipo2
            Flog.writeline Espacios(Tabulador * 4) & "Tipo de estructura 3 " & Tipo3
        Else
            LogicaDistribucion = 0
            Razon_MesCalendario = False
            Flog.writeline Espacios(Tabulador * 4) & "No se encontro tipo de alcance de distribucion " & rs_Alcance!idtipodist
        End If
    Else
        'Busco alcance Global
        Flog.writeline Espacios(Tabulador * 3) & "Busco alcance Global"
        'Busco el tipo de distribucion
        Flog.writeline Espacios(Tabulador * 3) & "Busco el tipo de distribucion"
        StrSql = "SELECT * FROM tipodistrib "
        StrSql = StrSql & " WHERE tdistglobal = -1"
        OpenRecordset StrSql, rs_TipoDistribucion
        
        If Not rs_TipoDistribucion.EOF Then
            Razon_MesCalendario = CBool(rs_TipoDistribucion!tdistcalend)
            LogicaDistribucion = rs_TipoDistribucion!tdistlogica
            Tipo1 = IIf(Not EsNulo(rs_TipoDistribucion!tdistTenro1), rs_TipoDistribucion!tdistTenro1, 0)
            Tipo2 = IIf(Not EsNulo(rs_TipoDistribucion!tdistTenro2), rs_TipoDistribucion!tdistTenro2, 0)
            Tipo3 = IIf(Not EsNulo(rs_TipoDistribucion!tdistTenro3), rs_TipoDistribucion!tdistTenro3, 0)
            
            If Razon_MesCalendario Then
                Flog.writeline Espacios(Tabulador * 4) & "Razon Mes Calendario "
            Else
                Flog.writeline Espacios(Tabulador * 4) & "Razon de 30 d�as "
            End If
            Flog.writeline Espacios(Tabulador * 4) & "Logica de Distribuci�n " & LogicaDistribucion
            Flog.writeline Espacios(Tabulador * 4) & "Tipo de estructura 1 " & Tipo1
            Flog.writeline Espacios(Tabulador * 4) & "Tipo de estructura 2 " & Tipo2
            Flog.writeline Espacios(Tabulador * 4) & "Tipo de estructura 3 " & Tipo3
        Else
            LogicaDistribucion = 0
            Razon_MesCalendario = False
            Flog.writeline Espacios(Tabulador * 4) & "No se encontro ningun tipo de alcance de distribucion "
        End If
    End If
End If

If rs_Alcance.State = adStateOpen Then rs_Alcance.Close
Set rs_Alcance = Nothing
End Sub


Public Sub Distrubucion_Historica(ByVal Pronro As Long, ByVal Cabecera As Long, ByVal Tercero As Long, ByVal FechaDesde As Date, ByVal FechaHasta As Date, ByVal TipoE1 As Long, ByVal TipoE2 As Long, TipoE3 As Long, ByVal Razon_MesCalendario As Boolean)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de Distribucion en base a Historicos de estructuras
' Autor      : FGZ
' Fecha      : 10/01/2005
' --------------------------------------------------------------------------------------------
Dim Monto As Single
Dim Cantidad As Single
Dim Aux_Monto As Single
Dim Aux_Cantidad As Single
Dim MontoAplicado As Single
Dim CantAplicada As Single

Dim TipoOrigen As Integer
Dim Origen As Long
Dim TotalDias As Integer
Dim dias As Integer
Dim Aux_Desde As Date
Dim Aux_Hasta As Date
Dim Aux_Desde1 As Date
Dim Aux_Hasta1 As Date
Dim Aux_Desde2 As Date
Dim Aux_Hasta2 As Date
Dim Aux_Desde3 As Date
Dim Aux_Hasta3 As Date
Dim Estructura1 As Long
Dim Estructura2 As Long
Dim Estructura3 As Long
Dim Ultima_Estructura1 As Long
Dim Ultima_Estructura2 As Long
Dim Ultima_Estructura3 As Long

Dim rs_Fases As New ADODB.Recordset
Dim rs_Detliq As New ADODB.Recordset
Dim rs_Aculiq As New ADODB.Recordset
Dim rs_Estructura1 As New ADODB.Recordset
Dim rs_Estructura2 As New ADODB.Recordset
Dim rs_Estructura3 As New ADODB.Recordset

'Limpio la tabla detcostos para esta cabecera y proceso
Flog.writeline Espacios(Tabulador * 3) & "Limpio la tabla detcostos para el proceso " & Pronro & " y cabecera " & Cabecera
StrSql = "DELETE FROM detcostos "
StrSql = StrSql & " WHERE cliqnro = " & Cabecera
StrSql = StrSql & " AND pronro = " & Pronro
objConn.Execute StrSql, , adExecuteNoRecords


'Buscar las fases y calcular la cantidad de dias efectivos entra elas fechas
Flog.writeline Espacios(Tabulador * 3) & "Buscar las fases y calcular la cantidad de dias efectivos entra elas fechas "
TotalDias = 0
StrSql = "SELECT * FROM fases WHERE empleado = " & Tercero
StrSql = StrSql & " AND sueldo = -1 "
StrSql = StrSql & " AND altfec <= " & ConvFecha(FechaHasta)
StrSql = StrSql & " AND ( bajfec IS NULL or bajfec > " & ConvFecha(FechaDesde) & ")"
StrSql = StrSql & " ORDER BY altfec "
OpenRecordset StrSql, rs_Fases
If rs_Fases.EOF Then
    Flog.writeline Espacios(Tabulador * 4) & "No se encontraron fases "
    Aux_Desde = FechaDesde
    Aux_Hasta = FechaHasta
    TotalDias = DateDiff("d", Aux_Desde, Aux_Hasta) + 1
End If
Do While Not rs_Fases.EOF
    Aux_Desde = IIf(rs_Fases!altfec < FechaDesde, FechaDesde, rs_Fases!altfec)
    
    If Not EsNulo(rs_Fases!bajfec) Then
        If rs_Fases!bajfec >= FechaDesde And rs_Fases!bajfec <= FechaHasta Then
            Aux_Hasta = rs_Fases!bajfec
        Else
            Aux_Hasta = FechaHasta
        End If
    Else
        Aux_Hasta = FechaHasta
    End If
    Flog.writeline Espacios(Tabulador * 4) & "Fases desde " & Aux_Desde & " hasta " & Aux_Hasta
    TotalDias = TotalDias + (DateDiff("d", Aux_Desde, Aux_Hasta) + 1)
    rs_Fases.MoveNext
Loop
Flog.writeline Espacios(Tabulador * 4) & "Total de dias usados " & TotalDias
If Not Razon_MesCalendario Then
    If TotalDias > 28 Then
        If TotalDias > 30 Then
            TotalDias = 30
        End If
        If Month(Aux_Desde) = 2 Then
            If Biciesto(Year(Aux_Desde)) Then
                If TotalDias = 29 Then
                    TotalDias = 30
                End If
            Else
                If TotalDias = 28 Then
                    TotalDias = 30
                End If
            End If
        End If
    End If
    Flog.writeline Espacios(Tabulador * 4) & "Topeo de dias " & TotalDias
End If

FechaHasta = Aux_Hasta
FechaDesde = Aux_Desde

'Conceptos
Flog.writeline Espacios(Tabulador * 3) & "Conceptos ... "
TipoOrigen = 1
StrSql = "SELECT * FROM detliq "
StrSql = StrSql & " INNER JOIN concepto ON detliq.concnro = concepto.concnro AND concepto.concapertura = -1"
StrSql = StrSql & " WHERE cliqnro =" & Cabecera
OpenRecordset StrSql, rs_Detliq

Do While Not rs_Detliq.EOF
    MontoAplicado = 0
    CantAplicada = 0
    Origen = rs_Detliq!concnro
    Monto = rs_Detliq!dlimonto
    Cantidad = rs_Detliq!dlicant
    Flog.writeline Espacios(Tabulador * 4) & "Concepto " & rs_Detliq!Conccod
    Flog.writeline Espacios(Tabulador * 4) & "Monto " & Monto
    Flog.writeline Espacios(Tabulador * 4) & "Cantidad " & Cantidad
    
    'ciclo por los tres tipos de estructura
    Flog.writeline Espacios(Tabulador * 4) & "Ciclo por los tres tipos de estructura"
    
    If TipoE1 <> 0 Then
        Flog.writeline Espacios(Tabulador * 5) & "Tipo de estructura 1: " & TipoE1
        StrSql = " SELECT * FROM his_estructura " & _
                 " WHERE ternro = " & Tercero & " AND " & _
                 " tenro =" & TipoE1 & " AND " & _
                 " (htetdesde <= " & ConvFecha(FechaHasta) & ") AND " & _
                 " ((" & ConvFecha(FechaDesde) & " <= htethasta) or (htethasta is null))" & _
                 " ORDER BY htetdesde "
        OpenRecordset StrSql, rs_Estructura1
        If rs_Estructura1.EOF Then
            Flog.writeline Espacios(Tabulador * 5) & "No se encontr� ninguna estructura de tipo 1 entre " & FechaDesde & " y " & FechaHasta & ". Inserto el total"
            Aux_Desde = FechaDesde
            Aux_Hasta = FechaHasta
            dias = CantidadDeDias(FechaDesde, FechaHasta, Aux_Desde, Aux_Hasta)
            Aux_Monto = (dias * Monto) / TotalDias
            Aux_Cantidad = (dias * Cantidad) / TotalDias
            MontoAplicado = MontoAplicado + Aux_Monto
            CantAplicada = CantAplicada + Aux_Cantidad
            Call Insertar_Detcostos(Cabecera, Pronro, TipoOrigen, Origen, Aux_Monto, Aux_Cantidad, TipoE1, Estructura1, TipoE2, Estructura2, TipoE3, Estructura3)
        End If

        Do While Not rs_Estructura1.EOF
            Estructura1 = rs_Estructura1!estrnro
            Ultima_Estructura1 = Estructura1
            Aux_Desde1 = IIf(rs_Estructura1!htetdesde < FechaDesde, FechaDesde, rs_Estructura1!htetdesde)
            If Not IsNull(rs_Estructura1!htethasta) Then
                Aux_Hasta1 = IIf(rs_Estructura1!htethasta > FechaHasta, FechaHasta, rs_Estructura1!htethasta)
            Else
                Aux_Hasta1 = FechaHasta
            End If

            If TipoE2 <> 0 Then
                Flog.writeline Espacios(Tabulador * 6) & "Tipo de estructura 2: " & TipoE2
                StrSql = " SELECT * FROM his_estructura " & _
                         " WHERE ternro = " & Tercero & " AND " & _
                         " tenro =" & TipoE2 & " AND " & _
                         " (htetdesde <= " & ConvFecha(FechaHasta) & ") AND " & _
                         " ((" & ConvFecha(FechaDesde) & " <= htethasta) or (htethasta is null))" & _
                         " ORDER BY htetdesde "
                OpenRecordset StrSql, rs_Estructura2
                If rs_Estructura2.EOF Then
                    Flog.writeline Espacios(Tabulador * 6) & "No se encontr� ninguna estructura de tipo 2 entre " & Aux_Desde & " y " & Aux_Hasta & ". Inserto el total"
                    dias = CantidadDeDias(FechaDesde, FechaHasta, Aux_Desde1, Aux_Hasta1)
                    If dias > 0 Then
                        Aux_Monto = (dias * Monto) / TotalDias
                        Aux_Cantidad = (dias * Cantidad) / TotalDias
                        MontoAplicado = MontoAplicado + Aux_Monto
                        CantAplicada = CantAplicada + Aux_Cantidad
                        Call Insertar_Detcostos(Cabecera, Pronro, TipoOrigen, Origen, Aux_Monto, Aux_Cantidad, TipoE1, Estructura1, TipoE2, Estructura2, TipoE3, Estructura3)
                    End If
                End If

                Do While Not rs_Estructura2.EOF
                    Estructura2 = rs_Estructura2!estrnro
                    Ultima_Estructura2 = Estructura2
                    Aux_Desde2 = IIf(rs_Estructura2!htetdesde < FechaDesde, FechaDesde, rs_Estructura2!htetdesde)
                    If Not IsNull(rs_Estructura2!htethasta) Then
                        Aux_Hasta2 = IIf(rs_Estructura2!htethasta > FechaHasta, FechaHasta, rs_Estructura2!htethasta)
                    Else
                        Aux_Hasta2 = FechaHasta
                    End If

                    If TipoE3 <> 0 Then
                        Flog.writeline Espacios(Tabulador * 7) & "Tipo de estructura 3: " & TipoE3
                        StrSql = " SELECT * FROM his_estructura " & _
                                 " WHERE ternro = " & Tercero & " AND " & _
                                 " tenro =" & TipoE3 & " AND " & _
                                 " (htetdesde <= " & ConvFecha(FechaHasta) & ") AND " & _
                                 " ((" & ConvFecha(FechaDesde) & " <= htethasta) or (htethasta is null))" & _
                                 " ORDER BY htetdesde "
                        OpenRecordset StrSql, rs_Estructura3
                        If rs_Estructura3.EOF Then
                            Flog.writeline Espacios(Tabulador * 7) & "No se encontr� ninguna estructura de tipo 3 entre " & Aux_Desde & " y " & Aux_Hasta & ". Inserto el total"
                            Aux_Desde = mayorFecha(Aux_Desde1, Aux_Desde2, Aux_Desde3, 2)
                            Aux_Hasta = menorFecha(Aux_Hasta1, Aux_Hasta2, Aux_Hasta3, 2)
                            dias = CantidadDeDias(FechaDesde, FechaHasta, Aux_Desde, Aux_Hasta)
                            If dias > 0 Then
                                Aux_Monto = (dias * Monto) / TotalDias
                                Aux_Cantidad = (dias * Cantidad) / TotalDias
                                MontoAplicado = MontoAplicado + Aux_Monto
                                CantAplicada = CantAplicada + Aux_Cantidad
                                Call Insertar_Detcostos(Cabecera, Pronro, TipoOrigen, Origen, Aux_Monto, Aux_Cantidad, TipoE1, Estructura1, TipoE2, Estructura2, TipoE3, Estructura3)
                            End If
                        End If

                        Do While Not rs_Estructura3.EOF
                            Estructura3 = rs_Estructura3!estrnro
                            Ultima_Estructura3 = Estructura3
                            Aux_Desde3 = IIf(rs_Estructura3!htetdesde < FechaDesde, FechaDesde, rs_Estructura3!htetdesde)
                            If Not IsNull(rs_Estructura3!htethasta) Then
                                Aux_Hasta3 = IIf(rs_Estructura3!htethasta > FechaHasta, FechaHasta, rs_Estructura3!htethasta)
                            Else
                                Aux_Hasta3 = FechaHasta
                            End If

                            Aux_Desde = mayorFecha(Aux_Desde1, Aux_Desde2, Aux_Desde3, 3)
                            Aux_Hasta = menorFecha(Aux_Hasta1, Aux_Hasta2, Aux_Hasta3, 3)
                            dias = CantidadDeDias(FechaDesde, FechaHasta, Aux_Desde, Aux_Hasta)
                            If dias > 0 Then
                                Aux_Monto = (dias * Monto) / TotalDias
                                Aux_Cantidad = (dias * Cantidad) / TotalDias
                                MontoAplicado = MontoAplicado + Aux_Monto
                                CantAplicada = CantAplicada + Aux_Cantidad
                                Call Insertar_Detcostos(Cabecera, Pronro, TipoOrigen, Origen, Aux_Monto, Aux_Cantidad, TipoE1, Estructura1, TipoE2, Estructura2, TipoE3, Estructura3)
                            End If

                            rs_Estructura3.MoveNext
                        Loop
                    Else
                        Aux_Desde = mayorFecha(Aux_Desde1, Aux_Desde2, Aux_Desde3, 2)
                        Aux_Hasta = menorFecha(Aux_Hasta1, Aux_Hasta2, Aux_Hasta3, 2)
                        dias = CantidadDeDias(FechaDesde, FechaHasta, Aux_Desde, Aux_Hasta)
                        If dias > 0 Then
                            Aux_Monto = (dias * Monto) / TotalDias
                            Aux_Cantidad = (dias * Cantidad) / TotalDias
                            MontoAplicado = MontoAplicado + Aux_Monto
                            CantAplicada = CantAplicada + Aux_Cantidad
                            Call Insertar_Detcostos(Cabecera, Pronro, TipoOrigen, Origen, Aux_Monto, Aux_Cantidad, TipoE1, Estructura1, TipoE2, Estructura2, TipoE3, Estructura3)
                        End If
                    End If

                    rs_Estructura2.MoveNext
                Loop
            Else
                dias = CantidadDeDias(FechaDesde, FechaHasta, Aux_Desde1, Aux_Hasta1)
                Aux_Monto = (dias * Monto) / TotalDias
                Aux_Cantidad = (dias * Cantidad) / TotalDias
                MontoAplicado = MontoAplicado + Aux_Monto
                CantAplicada = CantAplicada + Aux_Cantidad
                Call Insertar_Detcostos(Cabecera, Pronro, TipoOrigen, Origen, Aux_Monto, Aux_Cantidad, TipoE1, Estructura1, TipoE2, Estructura2, TipoE3, Estructura3)
            End If

            rs_Estructura1.MoveNext
        Loop
    Else
        Flog.writeline Espacios(Tabulador * 4) & "No hay apertura. Inserto el total "
        Aux_Desde = FechaDesde
        Aux_Hasta = FechaHasta
        dias = CantidadDeDias(FechaDesde, FechaHasta, Aux_Desde, Aux_Hasta)
        Aux_Monto = (dias * Monto) / TotalDias
        Aux_Cantidad = (dias * Cantidad) / TotalDias
        MontoAplicado = MontoAplicado + Aux_Monto
        CantAplicada = CantAplicada + Aux_Cantidad
        Call Insertar_Detcostos(Cabecera, Pronro, TipoOrigen, Origen, Aux_Monto, Aux_Cantidad, TipoE1, Estructura1, TipoE2, Estructura2, TipoE3, Estructura3)
    End If

    'Reviso si qued� saldo
    Flog.writeline Espacios(Tabulador * 4) & "Reviso si qued� saldo "
    If MontoAplicado <> Monto Then
        Aux_Monto = Monto - MontoAplicado
        Aux_Cantidad = Cantidad - CantAplicada
        Flog.writeline Espacios(Tabulador * 4) & "Actualizo saldo. Monto " & Aux_Monto & " Cantidad " & Aux_Cantidad
        Call Actualizar_Detcostos(Cabecera, Pronro, TipoOrigen, Origen, Aux_Monto, Aux_Cantidad, TipoE1, Ultima_Estructura1, TipoE2, Ultima_Estructura2, TipoE3, Ultima_Estructura3)
    End If

    rs_Detliq.MoveNext
Loop


'Acumuladores
Flog.writeline Espacios(Tabulador * 3) & "Acumuladores ... "
TipoOrigen = 2
StrSql = "SELECT * FROM acu_liq "
StrSql = StrSql & " INNER JOIN acumulador ON acu_liq.acunro = acumulador.acunro AND acumulador.acuapertura = -1"
StrSql = StrSql & " WHERE cliqnro =" & Cabecera
OpenRecordset StrSql, rs_Aculiq

Do While Not rs_Aculiq.EOF
    MontoAplicado = 0
    CantAplicada = 0
    Origen = rs_Aculiq!acuNro
    Monto = rs_Aculiq!almonto
    Cantidad = rs_Aculiq!alcant
    Flog.writeline Espacios(Tabulador * 4) & "Acumulador " & rs_Aculiq!acuNro
    Flog.writeline Espacios(Tabulador * 4) & "Monto " & Monto
    Flog.writeline Espacios(Tabulador * 4) & "Cantidad " & Cantidad
    
    'ciclo por los tres tipos de estructura
    If TipoE1 <> 0 Then
        Flog.writeline Espacios(Tabulador * 5) & "Tipo de estructura 1: " & TipoE1
        StrSql = " SELECT * FROM his_estructura " & _
                 " WHERE ternro = " & Tercero & " AND " & _
                 " tenro =" & TipoE1 & " AND " & _
                 " (htetdesde <= " & ConvFecha(FechaHasta) & ") AND " & _
                 " ((" & ConvFecha(FechaDesde) & " <= htethasta) or (htethasta is null))" & _
                 " ORDER BY htetdesde "
        OpenRecordset StrSql, rs_Estructura1
        If rs_Estructura1.EOF Then
            Flog.writeline Espacios(Tabulador * 5) & "No se encontr� ninguna estructura de tipo 1 entre " & FechaDesde & " y " & FechaHasta & ". Inserto el total"
            Aux_Desde = FechaDesde
            Aux_Hasta = FechaHasta
            dias = CantidadDeDias(FechaDesde, FechaHasta, Aux_Desde, Aux_Hasta)
            Aux_Monto = (dias * Monto) / TotalDias
            Aux_Cantidad = (dias * Cantidad) / TotalDias
            MontoAplicado = MontoAplicado + Aux_Monto
            CantAplicada = CantAplicada + Aux_Cantidad
            Call Insertar_Detcostos(Cabecera, Pronro, TipoOrigen, Origen, Aux_Monto, Aux_Cantidad, TipoE1, Estructura1, TipoE2, Estructura2, TipoE3, Estructura3)
        End If

        Do While Not rs_Estructura1.EOF
            Estructura1 = rs_Estructura1!estrnro
            Ultima_Estructura1 = Estructura1
            Aux_Desde1 = IIf(rs_Estructura1!htetdesde < FechaDesde, FechaDesde, rs_Estructura1!htetdesde)
            If Not IsNull(rs_Estructura1!htethasta) Then
                Aux_Hasta1 = IIf(rs_Estructura1!htethasta > FechaHasta, FechaHasta, rs_Estructura1!htethasta)
            Else
                Aux_Hasta1 = FechaHasta
            End If

            If TipoE2 <> 0 Then
                Flog.writeline Espacios(Tabulador * 6) & "Tipo de estructura 2: " & TipoE2
                StrSql = " SELECT * FROM his_estructura " & _
                         " WHERE ternro = " & Tercero & " AND " & _
                         " tenro =" & TipoE2 & " AND " & _
                         " (htetdesde <= " & ConvFecha(FechaHasta) & ") AND " & _
                         " ((" & ConvFecha(FechaDesde) & " <= htethasta) or (htethasta is null))" & _
                         " ORDER BY htetdesde "
                OpenRecordset StrSql, rs_Estructura2
                If rs_Estructura2.EOF Then
                    Flog.writeline Espacios(Tabulador * 6) & "No se encontr� ninguna estructura de tipo 2 entre " & Aux_Desde & " y " & Aux_Hasta & ". Inserto el total"
                    dias = CantidadDeDias(FechaDesde, FechaHasta, Aux_Desde1, Aux_Hasta1)
                    If dias > 0 Then
                        Aux_Monto = (dias * Monto) / TotalDias
                        Aux_Cantidad = (dias * Cantidad) / TotalDias
                        MontoAplicado = MontoAplicado + Aux_Monto
                        CantAplicada = CantAplicada + Aux_Cantidad
                        Call Insertar_Detcostos(Cabecera, Pronro, TipoOrigen, Origen, Aux_Monto, Aux_Cantidad, TipoE1, Estructura1, TipoE2, Estructura2, TipoE3, Estructura3)
                    End If
                End If

                Do While Not rs_Estructura2.EOF
                    Estructura2 = rs_Estructura2!estrnro
                    Ultima_Estructura2 = Estructura2
                    Aux_Desde2 = IIf(rs_Estructura2!htetdesde < FechaDesde, FechaDesde, rs_Estructura2!htetdesde)
                    If Not IsNull(rs_Estructura2!htethasta) Then
                        Aux_Hasta2 = IIf(rs_Estructura2!htethasta > FechaHasta, FechaHasta, rs_Estructura2!htethasta)
                    Else
                        Aux_Hasta2 = FechaHasta
                    End If

                    If TipoE3 <> 0 Then
                        Flog.writeline Espacios(Tabulador * 7) & "Tipo de estructura 3: " & TipoE3
                        StrSql = " SELECT * FROM his_estructura " & _
                                 " WHERE ternro = " & Tercero & " AND " & _
                                 " tenro =" & TipoE3 & " AND " & _
                                 " (htetdesde <= " & ConvFecha(FechaHasta) & ") AND " & _
                                 " ((" & ConvFecha(FechaDesde) & " <= htethasta) or (htethasta is null))" & _
                                 " ORDER BY htetdesde "
                        OpenRecordset StrSql, rs_Estructura3
                        If rs_Estructura3.EOF Then
                            Flog.writeline Espacios(Tabulador * 7) & "No se encontr� ninguna estructura de tipo 3 entre " & Aux_Desde & " y " & Aux_Hasta & ". Inserto el total"
                            Aux_Desde = mayorFecha(Aux_Desde1, Aux_Desde2, Aux_Desde3, 2)
                            Aux_Hasta = menorFecha(Aux_Hasta1, Aux_Hasta2, Aux_Hasta3, 2)
                            dias = CantidadDeDias(FechaDesde, FechaHasta, Aux_Desde, Aux_Hasta)
                            If dias > 0 Then
                                Aux_Monto = (dias * Monto) / TotalDias
                                Aux_Cantidad = (dias * Cantidad) / TotalDias
                                MontoAplicado = MontoAplicado + Aux_Monto
                                CantAplicada = CantAplicada + Aux_Cantidad
                                Call Insertar_Detcostos(Cabecera, Pronro, TipoOrigen, Origen, Aux_Monto, Aux_Cantidad, TipoE1, Estructura1, TipoE2, Estructura2, TipoE3, Estructura3)
                            End If
                        End If

                        Do While Not rs_Estructura3.EOF
                            Estructura3 = rs_Estructura3!estrnro
                            Ultima_Estructura3 = Estructura3
                            Aux_Desde3 = IIf(rs_Estructura3!htetdesde < FechaDesde, FechaDesde, rs_Estructura3!htetdesde)
                            If Not IsNull(rs_Estructura3!htethasta) Then
                                Aux_Hasta3 = IIf(rs_Estructura3!htethasta > FechaHasta, FechaHasta, rs_Estructura3!htethasta)
                            Else
                                Aux_Hasta3 = FechaHasta
                            End If

                            Aux_Desde = mayorFecha(Aux_Desde1, Aux_Desde2, Aux_Desde3, 3)
                            Aux_Hasta = menorFecha(Aux_Hasta1, Aux_Hasta2, Aux_Hasta3, 3)
                            dias = CantidadDeDias(FechaDesde, FechaHasta, Aux_Desde, Aux_Hasta)
                            If dias > 0 Then
                                Aux_Monto = (dias * Monto) / TotalDias
                                Aux_Cantidad = (dias * Cantidad) / TotalDias
                                MontoAplicado = MontoAplicado + Aux_Monto
                                CantAplicada = CantAplicada + Aux_Cantidad
                                Call Insertar_Detcostos(Cabecera, Pronro, TipoOrigen, Origen, Aux_Monto, Aux_Cantidad, TipoE1, Estructura1, TipoE2, Estructura2, TipoE3, Estructura3)
                            End If

                            rs_Estructura3.MoveNext
                        Loop
                    Else
                        Aux_Desde = mayorFecha(Aux_Desde1, Aux_Desde2, Aux_Hasta3, 2)
                        Aux_Hasta = menorFecha(Aux_Hasta1, Aux_Hasta2, Aux_Hasta3, 2)
                        dias = CantidadDeDias(FechaDesde, FechaHasta, Aux_Desde, Aux_Hasta)
                        If dias > 0 Then
                            Aux_Monto = (dias * Monto) / TotalDias
                            Aux_Cantidad = (dias * Cantidad) / TotalDias
                            MontoAplicado = MontoAplicado + Aux_Monto
                            CantAplicada = CantAplicada + Aux_Cantidad
                            Call Insertar_Detcostos(Cabecera, Pronro, TipoOrigen, Origen, Aux_Monto, Aux_Cantidad, TipoE1, Estructura1, TipoE2, Estructura2, TipoE3, Estructura3)
                        End If
                    End If

                    rs_Estructura2.MoveNext
                Loop
            Else
                dias = CantidadDeDias(FechaDesde, FechaHasta, Aux_Desde1, Aux_Hasta1)
                Aux_Monto = (dias * Monto) / TotalDias
                Aux_Cantidad = (dias * Cantidad) / TotalDias
                MontoAplicado = MontoAplicado + Aux_Monto
                CantAplicada = CantAplicada + Aux_Cantidad
                Call Insertar_Detcostos(Cabecera, Pronro, TipoOrigen, Origen, Aux_Monto, Aux_Cantidad, TipoE1, Estructura1, TipoE2, Estructura2, TipoE3, Estructura3)
            End If

            rs_Estructura1.MoveNext
        Loop
    Else
        Flog.writeline Espacios(Tabulador * 4) & "No hay apertura. Inserto el total "
        Aux_Desde = FechaDesde
        Aux_Hasta = FechaHasta
        dias = CantidadDeDias(FechaDesde, FechaHasta, Aux_Desde, Aux_Hasta)
        Aux_Monto = (dias * Monto) / TotalDias
        Aux_Cantidad = (dias * Cantidad) / TotalDias
        MontoAplicado = MontoAplicado + Aux_Monto
        CantAplicada = CantAplicada + Aux_Cantidad
        Call Insertar_Detcostos(Cabecera, Pronro, TipoOrigen, Origen, Aux_Monto, Aux_Cantidad, TipoE1, Estructura1, TipoE2, Estructura2, TipoE3, Estructura3)
    End If
    
    'Reviso si qued� saldo
    Flog.writeline Espacios(Tabulador * 4) & "Reviso si qued� saldo "
    If MontoAplicado <> Monto Then
        Aux_Monto = Monto - MontoAplicado
        Aux_Cantidad = Cantidad - CantAplicada
        Flog.writeline Espacios(Tabulador * 4) & "Actualizo saldo. Monto " & Aux_Monto & " Cantidad " & Aux_Cantidad
        Call Actualizar_Detcostos(Cabecera, Pronro, TipoOrigen, Origen, Aux_Monto, Aux_Cantidad, TipoE1, Ultima_Estructura1, TipoE2, Ultima_Estructura2, TipoE3, Ultima_Estructura3)
    End If
    
    rs_Aculiq.MoveNext
Loop


'Cierro y libero
If rs_Detliq.State = adStateOpen Then rs_Detliq.Close
Set rs_Detliq = Nothing
If rs_Aculiq.State = adStateOpen Then rs_Aculiq.Close
Set rs_Aculiq = Nothing
If rs_Estructura1.State = adStateOpen Then rs_Estructura1.Close
Set rs_Estructura1 = Nothing
If rs_Estructura2.State = adStateOpen Then rs_Estructura2.Close
Set rs_Estructura2 = Nothing
If rs_Estructura3.State = adStateOpen Then rs_Estructura3.Close
Set rs_Estructura3 = Nothing

End Sub


Public Sub Insertar_Detcostos(ByVal Cabecera As Long, ByVal Pronro As Long, ByVal TipoOrigen As Integer, Origen As Long, ByVal Monto As Single, ByVal Cantidad As Single, ByVal TipoE1 As Long, ByVal Estructura1 As Long, ByVal TipoE2 As Long, ByVal Estructura2 As Long, ByVal TipoE3 As Long, ByVal Estructura3 As Long)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que inserta el detalle en detcostos.
' Autor      : FGZ
' Fecha      : 11/01/2005
' --------------------------------------------------------------------------------------------

    'Inserto en detcostos
    Flog.writeline Espacios(Tabulador * 8) & "Inserto en detcostos Monto: " & Monto & " y cantidad: " & Cantidad
    StrSql = "INSERT INTO detcostos (pronro,cliqnro,tipoorigen,origen"
    If TipoE1 <> 0 Then
        StrSql = StrSql & ",tenro1,estrnro1"
    End If
    If TipoE2 <> 0 Then
        StrSql = StrSql & ",tenro2,estrnro2"
    End If
    If TipoE3 <> 0 Then
        StrSql = StrSql & ",tenro3,estrnro3"
    End If
    StrSql = StrSql & ",monto,cant"
    StrSql = StrSql & ")"
    StrSql = StrSql & " VALUES ("
    StrSql = StrSql & Pronro
    StrSql = StrSql & "," & Cabecera
    StrSql = StrSql & "," & TipoOrigen
    StrSql = StrSql & "," & Origen
    If TipoE1 <> 0 Then
        StrSql = StrSql & "," & TipoE1
        StrSql = StrSql & "," & Estructura1
    End If
    If TipoE2 <> 0 Then
        StrSql = StrSql & "," & TipoE2
        StrSql = StrSql & "," & Estructura2
    End If
    If TipoE3 <> 0 Then
        StrSql = StrSql & "," & TipoE3
        StrSql = StrSql & "," & Estructura3
    End If
    StrSql = StrSql & "," & Monto
    StrSql = StrSql & "," & Cantidad
    StrSql = StrSql & ")"
    objConn.Execute StrSql, , adExecuteNoRecords

End Sub

Public Sub Actualizar_Detcostos(ByVal Cabecera As Long, ByVal Pronro As Long, ByVal TipoOrigen As Integer, Origen As Long, ByVal Monto As Single, ByVal Cantidad As Single, ByVal TipoE1 As Long, ByVal Estructura1 As Long, ByVal TipoE2 As Long, ByVal Estructura2 As Long, ByVal TipoE3 As Long, ByVal Estructura3 As Long)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que inserta el detalle en detcostos.
' Autor      : FGZ
' Fecha      : 11/01/2005
' --------------------------------------------------------------------------------------------

    'Actualizo detcostos
    Flog.writeline Espacios(Tabulador * 5) & "Actualizo detcostos Monto: " & Monto & " y cantidad: " & Cantidad
    StrSql = "UPDATE detcostos set monto = monto + " & Monto
    StrSql = StrSql & ", cant = cant + " & Cantidad
    StrSql = StrSql & " WHERE cliqnro =" & Cabecera
    StrSql = StrSql & " AND pronro =" & Pronro
    StrSql = StrSql & " AND tipoorigen =" & TipoOrigen
    StrSql = StrSql & " AND origen =" & Origen
    If TipoE1 <> 0 Then
        StrSql = StrSql & "AND tenro1 = " & TipoE1
        StrSql = StrSql & "AND estrnro1 = " & Estructura1
    End If
    If TipoE2 <> 0 Then
        StrSql = StrSql & "AND tenro2 = " & TipoE2
        StrSql = StrSql & "AND estrnro2 = " & Estructura2
    End If
    If TipoE3 <> 0 Then
        StrSql = StrSql & "AND tenro3 = " & TipoE3
        StrSql = StrSql & "AND estrnro3 = " & Estructura3
    End If
    objConn.Execute StrSql, , adExecuteNoRecords

End Sub



Function mayorFecha(ByVal F1 As Date, ByVal F2 As Date, ByVal F3 As Date, ByVal Tipo As Integer)
  Dim salida
  
  If Tipo = 1 Then
     salida = F1
  Else
     If Tipo = 2 Then
        If F1 > F2 Then
           salida = F1
        Else
           salida = F2
        End If
     Else
        If F1 > F2 Then
           If F1 > F3 Then
              salida = F1
           Else
              salida = F3
           End If
        Else
           If F2 > F3 Then
              salida = F2
           Else
              salida = F3
           End If
        End If
     End If
  End If
  
  mayorFecha = salida
End Function


Function menorFecha(ByVal F1 As Date, ByVal F2 As Date, ByVal F3 As Date, ByVal Tipo As Integer)
  Dim salida
  
  If Tipo = 1 Then
     salida = F1
  Else
     If Tipo = 2 Then
        If F1 < F2 Then
           salida = F1
        Else
           salida = F2
        End If
     Else
        If F1 < F2 Then
           If F1 < F3 Then
              salida = F1
           Else
              salida = F3
           End If
        Else
           If F2 < F3 Then
              salida = F2
           Else
              salida = F3
           End If
        End If
     End If
  End If
  
  menorFecha = salida
End Function

