Attribute VB_Name = "MdlValeInd"
Option Explicit

'Const Version = 1.01
'Const FechaVersion = "15/02/2007"
'Autor = Martin Ferraro - Version Inicial
'Reporte Vales IND AGD

'Const Version = 1.02
'Const FechaVersion = "02/03/2007" 'MAF Buscar el campo del firmante en campo nuevo de tabla vales

Const Version = "1.03"
Const FechaVersion = "31/07/2009" 'MAF - Encriptacion de string connection




Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial del Generador de Reporte Vales.
' Autor      : Martin Ferraro
' Fecha      : 15/02/2007
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

'    strCmdLine = Command()
'    ArrParametros = Split(strCmdLine, " ", -1)
'    If UBound(ArrParametros) > 0 Then
'        If IsNumeric(ArrParametros(0)) Then
'            NroProcesoBatch = ArrParametros(0)
'            Etiqueta = ArrParametros(1)
'        Else
'            Exit Sub
'        End If
'    Else
'        If IsNumeric(strCmdLine) Then
'            NroProcesoBatch = strCmdLine
'        Else
'            Exit Sub
'        End If
'    End If

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

    On Error GoTo ME_Main
    
    Nombre_Arch = PathFLog & "Generacion_Reporte_Vales_Ind" & "-" & NroProcesoBatch & ".log"
    
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
    
    On Error Resume Next
    'Abro la conexion
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If
    OpenConnection strconexion, objconnProgreso
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If
    On Error GoTo 0
    
    On Error GoTo ME_Main
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0 ,bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Flog.writeline "Pone el estado en procesando"
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 158 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        bprcparam = rs_batch_proceso!bprcparam
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        Call GenerarReporte(NroProcesoBatch, bprcparam)
    Else
        Flog.writeline "no encontr� el proceso"
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
' Descripcion: Procedimiento de generacion del reporte de Vales Ind de AGD
' Autor      : Martin Ferraro
' Fecha      : 15/02/2007
' --------------------------------------------------------------------------------------------

'Arreglo que contiene los parametros
Dim arrParam
Dim I As Long


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


'RecordSet
Dim rs_Empleados As New ADODB.Recordset
Dim rs_Consult As New ADODB.Recordset

'Variables
Dim Domicilio As String
Dim Provincia As String
Dim CPLoc As String
Dim TerApe As String
Dim TerNom As String
Dim Orden As Long
Dim EmpLogo As String
Dim EmpLogoAlto As Integer
Dim EmpLogoAncho As Integer

' Inicio codigo ejecutable
On Error GoTo CE

' Levanto cada parametro por separado, el separador de parametros es "@"
Flog.writeline Espacios(Tabulador * 0) & "levantando parametros" & Parametros
If Not IsNull(Parametros) Then
    
    arrParam = Split(Parametros, "@")
    
    If UBound(arrParam) = 15 Then
    
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
'MyBeginTrans

'Inicializacion de variables
Orden = 0

'---------------------------------------------------------------------------------
'Consulta Principal
'---------------------------------------------------------------------------------

StrSql = " SELECT vales.valnro, vales.empleado, vales.valmonto, vales.tvalenro, vales.pliqdto, vales.valfecped, vales.valusuario,"
StrSql = StrSql & " tipovale.tvaledesabr, tipovale.tvaledesext,"
StrSql = StrSql & " empleado.terape, empleado.terape2, empleado.ternom,"
StrSql = StrSql & " empleado.ternom2, empleado.empleg, periodo.pliqhasta,"
StrSql = StrSql & " estrsuc.estrdabr sucdesabr, estrsuc.estrnro sucnro, sucursal.ternro sucursalternro, "
StrSql = StrSql & " estrempresa.estrdabr empresadesabr, estrempresa.estrnro empresanro, empresa.ternro empresaternro"
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
StrSql = StrSql & " INNER JOIN his_estructura his_empresa ON empleado.ternro = his_empresa.ternro"
StrSql = StrSql & " AND his_empresa.tenro = 10"
StrSql = StrSql & " AND his_empresa.htetdesde <= periodo.pliqhasta AND (his_empresa.htethasta IS NULL OR his_empresa.htethasta >= periodo.pliqhasta )"
StrSql = StrSql & " INNER JOIN estructura estrempresa ON estrempresa.estrnro = his_empresa.estrnro"
StrSql = StrSql & " INNER JOIN empresa ON empresa.estrnro = estrempresa.estrnro"
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
    Provincia = ""
    TerApe = IIf(EsNulo(rs_Empleados!TerApe), "", rs_Empleados!TerApe) & IIf(EsNulo(rs_Empleados!terape2), "", " " & rs_Empleados!terape2)
    TerNom = IIf(EsNulo(rs_Empleados!TerNom), "", rs_Empleados!TerNom) & IIf(EsNulo(rs_Empleados!ternom2), "", " " & rs_Empleados!ternom2)
    
    Orden = Orden + 1
    
    Flog.writeline Espacios(Tabulador * 1) & "Buscando datos de domicilio de la empresa " & rs_Empleados!empresadesabr
    StrSql = " SELECT detdom.calle,detdom.nro,localidad.locdesc,piso,oficdepto,codigopostal,provdesc"
    StrSql = StrSql & " From cabdom"
    StrSql = StrSql & " INNER JOIN detdom ON detdom.domnro = cabdom.domnro AND cabdom.ternro = " & rs_Empleados!empresaternro
    StrSql = StrSql & " INNER JOIN localidad ON detdom.locnro = localidad.locnro"
    StrSql = StrSql & " INNER JOIN provincia ON detdom.provnro = provincia.provnro"
    StrSql = StrSql & " WHERE cabdom.domdefault = -1"
    OpenRecordset StrSql, rs_Consult
    
    If Not rs_Consult.EOF Then
        Domicilio = IIf(EsNulo(rs_Consult!calle), "", rs_Consult!calle)
        Domicilio = Domicilio & IIf(EsNulo(rs_Consult!nro), "", " " & rs_Consult!nro)
        Domicilio = Domicilio & IIf(EsNulo(rs_Consult!piso), "", " Piso " & rs_Consult!piso)
        Domicilio = Domicilio & IIf(EsNulo(rs_Consult!oficdepto), " ", " Dpto. " & rs_Consult!oficdepto)
        CPLoc = IIf(EsNulo(rs_Consult!codigopostal), "", rs_Consult!codigopostal)
        CPLoc = CPLoc & IIf(EsNulo(rs_Consult!locdesc), "", " " & rs_Consult!locdesc)
        Provincia = IIf(EsNulo(rs_Consult!provdesc), "", rs_Consult!provdesc)
    Else
        Flog.writeline Espacios(Tabulador * 2) & "Domicilio No encontado."
    End If
    rs_Consult.Close
    
    'Consulta para buscar el logo de la empresa
    StrSql = "SELECT ter_imag.terimnombre, tipoimag.tipimdire, tipoimag.tipimanchodef, tipoimag.tipimaltodef"
    StrSql = StrSql & " From ter_imag"
    StrSql = StrSql & " INNER JOIN tipoimag ON tipoimag.tipimnro = 6 AND tipoimag.tipimnro = ter_imag.tipimnro"
    StrSql = StrSql & " AND ter_imag.ternro =" & rs_Empleados!empresaternro
    OpenRecordset StrSql, rs_Consult
    If rs_Consult.EOF Then
        Flog.writeline "No se encontr� el Logo de la Empresa"
        'Exit Sub
        EmpLogo = ""
        EmpLogoAlto = 0
        EmpLogoAncho = 0
    Else
        EmpLogo = rs_Consult!tipimdire & rs_Consult!terimnombre
        EmpLogoAlto = rs_Consult!tipimaltodef
        EmpLogoAncho = rs_Consult!tipimanchodef
    End If
    rs_Consult.Close

    Flog.writeline Espacios(Tabulador * 0) & "Guardando en base."
    StrSql = " INSERT INTO rep_vales_ind "
    StrSql = StrSql & " ("
    StrSql = StrSql & " bpronro,"
    StrSql = StrSql & " ternro,"
    StrSql = StrSql & " valnro,"
    
    StrSql = StrSql & " legajo,"
    StrSql = StrSql & " nombre,"
    StrSql = StrSql & " apellido,"
    StrSql = StrSql & " emprnro,"
    StrSql = StrSql & " empdesc,"
    StrSql = StrSql & " domicilio,"
    StrSql = StrSql & " cp,"
    StrSql = StrSql & " sucnro,"
    StrSql = StrSql & " sucdesc,"
    StrSql = StrSql & " provincia,"
    
    StrSql = StrSql & " tvalenro,"
    StrSql = StrSql & " tvaledesabr,"
    StrSql = StrSql & " monto,"
    StrSql = StrSql & " fecha,"
    StrSql = StrSql & " orden,"
    StrSql = StrSql & " obs1,"
    StrSql = StrSql & " obs2,"
    StrSql = StrSql & " obs3,"
    
    StrSql = StrSql & " titulo,"
    StrSql = StrSql & " emplogo,"
    StrSql = StrSql & " emplogoalto,"
    StrSql = StrSql & " emplogoancho,"
    StrSql = StrSql & " firmante"
    
    
    StrSql = StrSql & " )"
    StrSql = StrSql & " VALUES ("
    StrSql = StrSql & NroProcesoBatch
    StrSql = StrSql & " , " & rs_Empleados!Empleado
    StrSql = StrSql & " , " & rs_Empleados!valnro
    StrSql = StrSql & " , " & rs_Empleados!empleg
    StrSql = StrSql & " , '" & Mid(TerNom, 1, 100) & "'"
    StrSql = StrSql & " , '" & Mid(TerApe, 1, 100) & "'"
    StrSql = StrSql & " , " & rs_Empleados!empresanro
    StrSql = StrSql & " , '" & Mid(rs_Empleados!empresadesabr, 1, 100) & "'"
    StrSql = StrSql & " , '" & Mid(Domicilio, 1, 100) & "'"
    StrSql = StrSql & " , '" & Mid(CPLoc, 1, 100) & "'"
    StrSql = StrSql & " , " & rs_Empleados!sucnro
    StrSql = StrSql & " , '" & Mid(rs_Empleados!sucdesabr, 1, 100) & "'"
    StrSql = StrSql & " , '" & Mid(Provincia, 1, 100) & "'"
    StrSql = StrSql & " , " & rs_Empleados!tvalenro
    StrSql = StrSql & " , '" & Mid(rs_Empleados!tvaledesabr, 1, 100) & "'"
    StrSql = StrSql & " , " & rs_Empleados!valmonto
    If EsNulo(rs_Empleados!valfecped) Then
        StrSql = StrSql & " , null"
    Else
        StrSql = StrSql & " , " & ConvFecha(rs_Empleados!valfecped)
    End If
    StrSql = StrSql & " , " & Orden
    StrSql = StrSql & " , '" & Mid("", 1, 100) & "'"
    StrSql = StrSql & " , '" & Mid("", 1, 100) & "'"
    StrSql = StrSql & " , '" & Mid("", 1, 100) & "'"
    
    StrSql = StrSql & " , '" & Mid(TituloFiltro, 1, 200) & "'"
    StrSql = StrSql & " , '" & Mid(EmpLogo, 1, 100) & "'"
    StrSql = StrSql & " , " & EmpLogoAlto
    StrSql = StrSql & " , " & EmpLogoAncho
    StrSql = StrSql & " , '" & IIf(EsNulo(rs_Empleados!valusuario), "", Mid(rs_Empleados!valusuario, 1, 100)) & "'"
    
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
'If Not HuboError Then
'    MyCommitTrans
'End If


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
    
    'MyRollbackTrans
    
    'MyBeginTrans
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
             "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    'MyCommitTrans
    
    HuboError = True

End Sub

