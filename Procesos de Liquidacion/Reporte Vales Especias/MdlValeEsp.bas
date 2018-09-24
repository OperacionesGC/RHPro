Attribute VB_Name = "MdlValeEsp"
Option Explicit

Const Version = 1.01
Const FechaVersion = "06/02/2007"
'Autor = Martin Ferraro
'Reporte Vales Especias AGD




Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial del Generador de Reporte AFP.
' Autor      : Martin Ferraro
' Fecha      : 31/01/2007
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim objconnMain As New ADODB.Connection
Dim strCmdLine
Dim Nombre_Arch As String
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

    On Error Resume Next
    'Abro la conexion
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Then
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If
    OpenConnection strconexion, objconnProgreso
    If Err.Number <> 0 Then
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If
    On Error GoTo 0
    
    On Error GoTo ME_Main
    
    Nombre_Arch = PathFLog & "Generacion_Reporte_Vales_Especias" & "-" & NroProcesoBatch & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
        
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaVersion
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-------------------------------------------------"
    Flog.writeline
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0 ,bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Flog.writeline "Pone el estado en procesando"
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 157 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        bprcparam = rs_batch_proceso!bprcparam
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        Call GenerarReporte(NroProcesoBatch, bprcparam)
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
    'If objConn.State = adStateOpen Then objConn.Close
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


Public Function ArmarListaPeriodo(ByRef Desde As Date, ByRef Hasta As Date) As String
Dim rs_periodo As New ADODB.Recordset
Dim aux As String

On Error GoTo ErrorArmarListaPeriodo

    aux = "0"
    
    StrSql = "SELECT * FROM periodo"
    StrSql = StrSql & " WHERE " & ConvFecha(Desde) & " <= periodo.pliqdesde"
    StrSql = StrSql & " AND periodo.pliqhasta <= " & ConvFecha(Hasta)
    OpenRecordset StrSql, rs_periodo
    
    Do While Not rs_periodo.EOF
        aux = aux & "," & rs_periodo!pliqnro
        rs_periodo.MoveNext
    Loop

    ArmarListaPeriodo = aux
    
If rs_periodo.State = adStateOpen Then rs_periodo.Close
Set rs_periodo = Nothing

Exit Function

ErrorArmarListaPeriodo:
Flog.writeline "Error en ArmarListaPeriodo: " & Err.Description
Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    
End Function


Public Sub GenerarReporte(ByVal bpronro As Long, ByVal Parametros As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del reporte de Vales de especias
' Autor      : Martin Ferraro
' Fecha      : 31/01/2007
' --------------------------------------------------------------------------------------------

'Arreglo que contiene los parametros
Dim arrParam
Dim i As Long


'Parametros desde ASP
Dim FiltroSql As String
Dim TipoVale As Long
Dim Revisado As Long
Dim PeriodoDesde As Long
Dim PeriodoHasta As Long
Dim Desde As Date
Dim Hasta As Date
Dim Tenro1 As Long
Dim Estrnro1 As Long
Dim Tenro2 As Long
Dim Estrnro2 As Long
Dim Tenro3 As Long
Dim Estrnro3 As Long
Dim FecEstr As Date
Dim TituloFiltro As String
Dim OrdenSql As String
Dim Leyenda As String

'RecordSet
Dim rs_Empleados As New ADODB.Recordset
Dim rs_Consult As New ADODB.Recordset

'Variables
Dim Domicilio As String
Dim CPLoc As String
Dim TerApe As String
Dim TerNom As String
Dim Orden As Long

' Inicio codigo ejecutable
On Error GoTo CE

' Levanto cada parametro por separado, el separador de parametros es "@"
Flog.writeline Espacios(Tabulador * 0) & "levantando parametros" & Parametros
If Not IsNull(Parametros) Then
    
    arrParam = Split(Parametros, "@")
    
    If UBound(arrParam) = 16 Then
    
        FiltroSql = arrParam(0)
        TipoVale = CLng(arrParam(1))
        Revisado = arrParam(2)
        PeriodoDesde = CLng(arrParam(3))
        PeriodoHasta = CLng(arrParam(4))
        Desde = CDate(arrParam(5))
        Hasta = CDate(arrParam(6))
        Tenro1 = CLng(arrParam(7))
        Estrnro1 = CLng(arrParam(8))
        Tenro2 = CLng(arrParam(9))
        Estrnro2 = CLng(arrParam(10))
        Tenro3 = CLng(arrParam(11))
        Estrnro3 = CLng(arrParam(12))
        FecEstr = CDate(arrParam(13))
        TituloFiltro = arrParam(14)
        OrdenSql = arrParam(15)
        Leyenda = arrParam(16)

    
        Flog.writeline Espacios(Tabulador * 1) & "Filtro = " & FiltroSql
        Flog.writeline Espacios(Tabulador * 1) & "TipoVale = " & TipoVale
        Flog.writeline Espacios(Tabulador * 1) & "Revisado = " & Revisado
        Flog.writeline Espacios(Tabulador * 1) & "Periodo Desde = " & PeriodoDesde
        Flog.writeline Espacios(Tabulador * 1) & "Periodo Hasta = " & PeriodoHasta
        Flog.writeline Espacios(Tabulador * 1) & "Desde = " & Desde
        Flog.writeline Espacios(Tabulador * 1) & "Hasta = " & Hasta
        Flog.writeline Espacios(Tabulador * 1) & "TE1 = " & Tenro1
        Flog.writeline Espacios(Tabulador * 1) & "Estr1 = " & Estrnro1
        Flog.writeline Espacios(Tabulador * 1) & "TE2 = " & Tenro2
        Flog.writeline Espacios(Tabulador * 1) & "Estr2 = " & Estrnro2
        Flog.writeline Espacios(Tabulador * 1) & "TE3 = " & Tenro3
        Flog.writeline Espacios(Tabulador * 1) & "Estr3 = " & Estrnro3
        Flog.writeline Espacios(Tabulador * 1) & "Fecha Estr =" & FecEstr
        Flog.writeline Espacios(Tabulador * 1) & "Titulo = " & TituloFiltro
        Flog.writeline Espacios(Tabulador * 1) & "Leyenda = " & Leyenda
        Flog.writeline Espacios(Tabulador * 1) & "Orden = " & OrdenSql
        
    Else
        Flog.writeline Espacios(Tabulador * 0) & "ERROR. La cantidad de parametros no es la esperada."
        Exit Sub
        
    End If
Else
    Flog.writeline Espacios(Tabulador * 0) & "ERROR. No se encuentran los paramentros."
    Exit Sub
End If

Flog.writeline

'Comienzo la transaccion
MyBeginTrans

'Inicializacion de variables
Orden = 0

'---------------------------------------------------------------------------------
'Consulta Principal
'---------------------------------------------------------------------------------

StrSql = " SELECT vales.valnro, vales.empleado, vales.valmonto, vales.tvalenro, vales.pliqdto, vales.valfecped, "
StrSql = StrSql & " tipovale.tvaledesabr, tipovale.tvaledesext,"
StrSql = StrSql & " empleado.terape, empleado.terape2, empleado.ternom,"
StrSql = StrSql & " empleado.ternom2, empleado.empleg, periodo.pliqhasta,"
StrSql = StrSql & " estrsuc.estrdabr sucdesabr, estrsuc.estrnro sucnro, sucursal.ternro sucursalternro, "
StrSql = StrSql & " estrempresa.estrdabr empresadesabr, estrempresa.estrnro empresanro "
StrSql = StrSql & " FROM vales"
'Filtro tipo vale
StrSql = StrSql & " INNER JOIN tipovale ON tipovale.tvalenro = vales.tvalenro"
StrSql = StrSql & " AND tipovale.tvalenro = " & TipoVale
StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = vales.empleado"
StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = vales.pliqdto"
'Filtra que el empleado tenga sucursal a fin periodo del vale, ademas para el ordenamiento
StrSql = StrSql & " INNER JOIN his_estructura suc ON empleado.ternro = suc.ternro"
StrSql = StrSql & " AND suc.tenro = 1"
StrSql = StrSql & " AND suc.htetdesde <= periodo.pliqhasta AND (suc.htethasta IS NULL OR suc.htethasta >= periodo.pliqhasta )"
StrSql = StrSql & " INNER JOIN estructura estrsuc ON estrsuc.estrnro = suc.estrnro"
StrSql = StrSql & " INNER JOIN sucursal ON sucursal.estrnro = estrsuc.estrnro"
'Filtra que el empleado tenga empresa a fin periodo del vale
StrSql = StrSql & " INNER JOIN his_estructura empresa ON empleado.ternro = empresa.ternro"
StrSql = StrSql & " AND empresa.tenro = 10"
StrSql = StrSql & " AND empresa.htetdesde <= periodo.pliqhasta AND (empresa.htethasta IS NULL OR empresa.htethasta >= periodo.pliqhasta )"
StrSql = StrSql & " INNER JOIN estructura estrempresa ON estrempresa.estrnro = empresa.estrnro"
'Filtros de niveles de estructura
If Tenro1 <> 0 Then
    StrSql = StrSql & " INNER JOIN his_estructura tenro1 ON empleado.ternro = tenro1.ternro "
    StrSql = StrSql & " AND tenro1.tenro = " & Tenro1
    StrSql = StrSql & " AND tenro1.htetdesde <= " & ConvFecha(FecEstr) & " AND (tenro1.htethasta IS NULL OR tenro1.htethasta >= " & ConvFecha(FecEstr) & ") "
    If Estrnro1 <> 0 Then
        StrSql = StrSql & " AND tenro1.estrnro = " & Estrnro1
    End If
End If
If Tenro2 <> 0 Then
    StrSql = StrSql & " INNER JOIN his_estructura tenro2 ON empleado.ternro = tenro2.ternro "
    StrSql = StrSql & " AND tenro2.tenro = " & Tenro2
    StrSql = StrSql & " AND tenro2.htetdesde <= " & ConvFecha(FecEstr) & " AND (tenro2.htethasta IS NULL OR tenro2.htethasta >= " & ConvFecha(FecEstr) & ") "
    If Estrnro2 <> 0 Then
        StrSql = StrSql & " AND tenro2.estrnro = " & Estrnro2
    End If
End If
If Tenro3 <> 0 Then
    StrSql = StrSql & " INNER JOIN his_estructura tenro3 ON empleado.ternro = tenro3.ternro "
    StrSql = StrSql & " AND tenro3.tenro = " & Tenro3
    StrSql = StrSql & " AND tenro3.htetdesde <= " & ConvFecha(FecEstr) & " AND (tenro3.htethasta IS NULL OR tenro3.htethasta >= " & ConvFecha(FecEstr) & ") "
    If Estrnro3 <> 0 Then
        StrSql = StrSql & " AND tenro3.estrnro = " & Estrnro3
    End If
End If
'Filtro Empleados
StrSql = StrSql & " WHERE " & FiltroSql
'Filtro Periodo
StrSql = StrSql & " AND vales.pliqdto in (" & ArmarListaPeriodo(Desde, Hasta) & ")"
'Filtro Revisado
If Revisado = "-1" Then
    StrSql = StrSql & " AND vales.valrevis = -1"
End If
StrSql = StrSql & " ORDER BY sucdesabr , " & OrdenSql
OpenRecordset StrSql, rs_Empleados


'seteo de las variables de progreso
Progreso = 0
CEmpleadosAProc = rs_Empleados.RecordCount
If CEmpleadosAProc = 0 Then
   Flog.writeline "no hay empleados"
   CEmpleadosAProc = 1
End If
IncPorc = (100 / CEmpleadosAProc)
        
Flog.writeline
Flog.writeline
Flog.writeline
Flog.writeline Espacios(Tabulador * 0) & "--------------------------------------------------------"
Flog.writeline Espacios(Tabulador * 0) & "Comienza el procesamiento de empleados."
Flog.writeline Espacios(Tabulador * 0) & "--------------------------------------------------------"
Flog.writeline


'Comienzo a procesar los empleados
Do While Not rs_Empleados.EOF
    
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "PROCESANDO: " & rs_Empleados!empleg & "  - " & rs_Empleados!TerApe & " " & rs_Empleados!TerNom & " Vale: " & rs_Empleados!valnro
    Flog.writeline Espacios(Tabulador * 0) & "-------------------------------------------------------------"
    
    Domicilio = ""
    CPLoc = ""
    TerApe = IIf(EsNulo(rs_Empleados!TerApe), "", rs_Empleados!TerApe) & IIf(EsNulo(rs_Empleados!terape2), "", " " & rs_Empleados!terape2)
    TerNom = IIf(EsNulo(rs_Empleados!TerNom), "", rs_Empleados!TerNom) & IIf(EsNulo(rs_Empleados!ternom2), "", " " & rs_Empleados!ternom2)
    
    Orden = Orden + 1
    
    Flog.writeline Espacios(Tabulador * 1) & "Buscando datos de domicilio de sucursal " & rs_Empleados!sucdesabr
    StrSql = " SELECT detdom.calle,detdom.nro,localidad.locdesc,piso,oficdepto,codigopostal"
    StrSql = StrSql & " From cabdom"
    StrSql = StrSql & " INNER JOIN detdom ON detdom.domnro = cabdom.domnro AND cabdom.ternro = " & rs_Empleados!sucursalternro
    StrSql = StrSql & " INNER JOIN localidad ON detdom.locnro = localidad.locnro"
    StrSql = StrSql & " WHERE cabdom.domdefault = -1"
    OpenRecordset StrSql, rs_Consult
    
    If Not rs_Consult.EOF Then
        Domicilio = IIf(EsNulo(rs_Consult!calle), "", rs_Consult!calle)
        Domicilio = Domicilio & IIf(EsNulo(rs_Consult!nro), "", " " & rs_Consult!nro)
        Domicilio = Domicilio & IIf(EsNulo(rs_Consult!piso), "", " Piso " & rs_Consult!piso)
        Domicilio = Domicilio & IIf(EsNulo(rs_Consult!oficdepto), " ", " Dpto. " & rs_Consult!oficdepto)
        CPLoc = IIf(EsNulo(rs_Consult!codigopostal), "", "(" & rs_Consult!codigopostal & ")")
        CPLoc = CPLoc & IIf(EsNulo(rs_Consult!locdesc), "", " " & rs_Consult!locdesc)
    Else
        Flog.writeline Espacios(Tabulador * 2) & "Domicilio No encontado."
    End If
    
    
    Flog.writeline Espacios(Tabulador * 0) & "Guardando en base."
    StrSql = " INSERT INTO rep_vales_esp "
    StrSql = StrSql & " ("
    StrSql = StrSql & " bpronro,"
    StrSql = StrSql & " ternro,"
    StrSql = StrSql & " legajo,"
    StrSql = StrSql & " nombre,"
    StrSql = StrSql & " apellido,"
    StrSql = StrSql & " emprnro,"
    StrSql = StrSql & " empdesc,"
    StrSql = StrSql & " domicilio,"
    StrSql = StrSql & " cp,"
    StrSql = StrSql & " sucnro,"
    StrSql = StrSql & " sucdesc,"
    StrSql = StrSql & " tvalenro,"
    StrSql = StrSql & " tvaledesabr,"
    StrSql = StrSql & " monto,"
    StrSql = StrSql & " fecha,"
    StrSql = StrSql & " orden,"
    StrSql = StrSql & " leyenda,"
    StrSql = StrSql & " titulo"
    StrSql = StrSql & " )"
    StrSql = StrSql & " VALUES ("
    StrSql = StrSql & NroProcesoBatch
    StrSql = StrSql & " , " & rs_Empleados!Empleado
    StrSql = StrSql & " , " & rs_Empleados!empleg
    StrSql = StrSql & " , '" & Mid(TerNom, 1, 100) & "'"
    StrSql = StrSql & " , '" & Mid(TerApe, 1, 100) & "'"
    StrSql = StrSql & " , " & rs_Empleados!empresanro
    StrSql = StrSql & " , '" & Mid(rs_Empleados!empresadesabr, 1, 100) & "'"
    StrSql = StrSql & " , '" & Mid(Domicilio, 1, 100) & "'"
    StrSql = StrSql & " , '" & Mid(CPLoc, 1, 100) & "'"
    StrSql = StrSql & " , " & rs_Empleados!sucnro
    StrSql = StrSql & " , '" & Mid(rs_Empleados!sucdesabr, 1, 100) & "'"
    StrSql = StrSql & " , " & rs_Empleados!tvalenro
    StrSql = StrSql & " , '" & Mid(rs_Empleados!tvaledesabr, 1, 100) & "'"
    StrSql = StrSql & " , " & rs_Empleados!valmonto
    If EsNulo(rs_Empleados!valfecped) Then
        StrSql = StrSql & " , null"
    Else
        StrSql = StrSql & " , " & ConvFecha(rs_Empleados!valfecped)
    End If
    StrSql = StrSql & " , " & Orden
    StrSql = StrSql & " , '" & Mid(Leyenda, 1, 100) & "'"
    StrSql = StrSql & " , '" & Mid(TituloFiltro, 1, 200) & "'"
    StrSql = StrSql & " )"
    objConn.Execute StrSql, , adExecuteNoRecords
    
    
    'Actualizo el progreso
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
             "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    'Paso a siguiente vale
    rs_Empleados.MoveNext
    
Loop

'Fin de la transaccion
If Not HuboError Then
    MyCommitTrans
End If


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
    
    MyRollbackTrans
    
    MyBeginTrans
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
             "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    MyCommitTrans
    
    HuboError = True

End Sub

