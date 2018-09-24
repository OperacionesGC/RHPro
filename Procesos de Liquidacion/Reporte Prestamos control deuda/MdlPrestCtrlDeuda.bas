Attribute VB_Name = "MdlPrestCtrlDeuda"
Option Explicit

'Const Version = "1.00"   'Martin Ferraro - Version Inicial
'Const FechaVersion = "17/08/2007"

Global Const Version = "1.01" ' Cesar Stankunas
Global Const FechaVersion = "05/08/2009"
Global Const UltimaModificacion = ""    'Encriptacion de string connection

Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial Reporte Prestamos Control Estado Deuda.
' Autor      : Martin Ferraro
' Fecha      : 17/08/2007
' Ultima Mod.:
' Descripcion:
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
    
    Nombre_Arch = PathFLog & "Prestamos Control Deuda" & "-" & NroProcesoBatch & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaVersion
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-----------------------------------------------------------------"
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
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 194 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        bprcparam = rs_batch_proceso!bprcparam
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        Call Generar_Reporte(NroProcesoBatch, bprcparam)
    End If
    
    TiempoFinalProceso = GetTickCount
    Flog.writeline "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    
    If Not HuboError Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcprogreso = 100, bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcprogreso = 100, bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
    End If
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Flog.Close
    objConn.Close
    objconnProgreso.Close
End Sub


Public Sub Generar_Reporte(ByVal bpronro As Long, ByVal Parametros As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion de Vales
' Autor      : FGZ
' Fecha      : 28/06/2004
' Ult. Mod   :
' --------------------------------------------------------------------------------------------
'Arreglo que contiene los parametros
Dim arrParam
Dim I As Long

'Parametros
Dim FiltroSql As String
Dim FecEstr As Date
Dim TeNro1 As Long
Dim EstrNro1 As Long
Dim TeNro2 As Long
Dim EstrNro2 As Long
Dim TeNro3 As Long
Dim EstrNro3 As Long
Dim OrdenSql As String
Dim Empresa As Long
Dim LineaPre As Long
Dim EstadoPre As Long
Dim MesCuota As Long
Dim AnioCuota As Long
Dim Titulo As String
Dim OrdenCorte As String

'RecordSets
Dim rs_Consult As New ADODB.Recordset

'Variables
Dim LineaAnt As Long
Dim EmpleadoAnt As Long
Dim MontoPrestamoAcum As Double
Dim MontoCuotasAcum As Double
Dim MontoCuotas As Double

Dim TernroSQL As Long
Dim LegajoSQL As Long
Dim ApellidoSQL As String
Dim NombreSQL As String
Dim TpnroSQL As Long
Dim TpdesabrSQL As String
Dim LnprenroSQL As Long
Dim LnpredabrSQL As String
Dim EstnroSQL As Long
Dim EstdabrSQL As String
Dim TeNro1SQL As Long
Dim Tenro1DescSQL As String
Dim EstrNro1SQL As Long
Dim Estrnro1DescSQL As String
Dim TeNro2SQL As Long
Dim Tenro2DescSQL As String
Dim EstrNro2SQL As Long
Dim Estrnro2DescSQL As String
Dim TeNro3SQL As Long
Dim Tenro3DescSQL As String
Dim EstrNro3SQL As Long
Dim Estrnro3DescSQL As String
Dim OrdenReg As Long


On Error GoTo E_Generar_Reporte
    
' Levanto cada parametro por separado, el separador de parametros es "@"
'l_filtro "@" l_fecestr "@" l_tenro1 "@" l_estrnro1 "@" l_tenro2 "@" l_estrnro2 "@" l_tenro3 "@"
'l_estrnro3 "@" l_orden "@" l_empresa "@" l_lnprenro "@" l_estnro "@" l_mes "@" l_anio "@" l_titulofiltro
Flog.writeline Espacios(Tabulador * 0) & "levantando parametros" & Parametros
If Not IsNull(Parametros) Then

    arrParam = Split(Parametros, "@")
    
    If UBound(arrParam) = 14 Then
        
        FiltroSql = arrParam(0)
        FecEstr = CDate(arrParam(1))
        TeNro1 = CLng(arrParam(2))
        EstrNro1 = CLng(arrParam(3))
        TeNro2 = CLng(arrParam(4))
        EstrNro2 = CLng(arrParam(5))
        TeNro3 = CLng(arrParam(6))
        EstrNro3 = CLng(arrParam(7))
        OrdenSql = arrParam(8)
        Empresa = CLng(arrParam(9))
        LineaPre = CLng(arrParam(10))
        EstadoPre = CLng(arrParam(11))
        MesCuota = CLng(arrParam(12))
        AnioCuota = CLng(arrParam(13))
        Titulo = arrParam(14)
        
        
        Flog.writeline Espacios(Tabulador * 1) & "Filtro = " & FiltroSql
        Flog.writeline Espacios(Tabulador * 1) & "Fecha Estruct = " & FecEstr
        Flog.writeline Espacios(Tabulador * 1) & "TE 1 = " & TeNro1
        Flog.writeline Espacios(Tabulador * 1) & "EST 1 = " & EstrNro1
        Flog.writeline Espacios(Tabulador * 1) & "TE 2 = " & TeNro2
        Flog.writeline Espacios(Tabulador * 1) & "EST 2 = " & EstrNro2
        Flog.writeline Espacios(Tabulador * 1) & "TE 3 = " & TeNro3
        Flog.writeline Espacios(Tabulador * 1) & "EST 3 = " & EstrNro3
        Flog.writeline Espacios(Tabulador * 1) & "Orden = " & OrdenSql
        Flog.writeline Espacios(Tabulador * 1) & "Empresa = " & Empresa
        Flog.writeline Espacios(Tabulador * 1) & "Linea Prest. = " & LineaPre
        Flog.writeline Espacios(Tabulador * 1) & "Estado Prest. = " & EstadoPre
        Flog.writeline Espacios(Tabulador * 1) & "Mes Saldo Hasta = " & MesCuota
        Flog.writeline Espacios(Tabulador * 1) & "Año Saldo Hasta = " & AnioCuota
        Flog.writeline Espacios(Tabulador * 1) & "Titulo = " & Titulo
        
    Else
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. La cantidad de parametros no es la esperada."
        Exit Sub
        
    End If
Else
    Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encuentran los paramentros."
    Exit Sub
End If

Flog.writeline

OrdenCorte = ""

Flog.writeline Espacios(Tabulador * 0) & "-----------------------------------------------------"
Flog.writeline Espacios(Tabulador * 0) & "CONSULTA PRINCIPAL"
Flog.writeline Espacios(Tabulador * 0) & "-----------------------------------------------------"
Flog.writeline
'--------------------------------------------------------------------------------------
'CONSULTA PRINCIPAL
'Busco Todos los detliq de los conceptos de los empleados
'--------------------------------------------------------------------------------------
StrSql = " SELECT "
StrSql = StrSql & " empleado.terape, empleado.terape2, empleado.ternom, empleado.ternom2, empleado.empest, empleado.empleg, empleado.ternro,"
StrSql = StrSql & " prestamo.prenro, prestamo.predesc, prestamo.preimp,"
StrSql = StrSql & " estadopre.estnro, estadopre.estdabr,"
StrSql = StrSql & " pre_linea.lnprenro ,pre_linea.lnpredabr,"
StrSql = StrSql & " tipoprestamo.tpnro , tipoprestamo.tpdesabr"
If TeNro1 <> 0 Then
    StrSql = StrSql & " ,tenro1.tenro tenro1, tenro1.estrnro estrnro1, estructura1.estrdabr estrdabr1, tipoest1.tedabr tipoesttedabr1"
End If
If TeNro2 <> 0 Then
    StrSql = StrSql & " ,tenro2.tenro tenro2, tenro2.estrnro estrnro2, estructura2.estrdabr estrdabr2, tipoest2.tedabr tipoesttedabr2"
End If
If TeNro3 <> 0 Then
    StrSql = StrSql & " ,tenro3.tenro tenro3, tenro3.estrnro estrnro3, estructura3.estrdabr estrdabr3, tipoest3.tedabr tipoesttedabr3"
End If

StrSql = StrSql & " FROM prestamo"
StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = prestamo.ternro"

'Filtro por estado
StrSql = StrSql & " INNER JOIN estadopre  ON estadopre.estnro = prestamo.estnro"
If EstadoPre <> 0 Then StrSql = StrSql & " AND estadopre.estnro = " & EstadoPre

'Filtro por linea
StrSql = StrSql & " INNER JOIN pre_linea  ON pre_linea.lnprenro = prestamo.lnprenro"
If LineaPre <> 0 Then StrSql = StrSql & " AND pre_linea.lnprenro = " & LineaPre

'Filtro tipo prestamo
StrSql = StrSql & " INNER JOIN tipoprestamo ON tipoprestamo.tpnro = pre_linea.tpnro"

'Filtro por empresa
StrSql = StrSql & " INNER JOIN his_estructura empre ON empre.ternro = empleado.ternro AND empre.tenro = 10"
StrSql = StrSql & " AND (empre.htetdesde<=" & ConvFecha(FecEstr) & " AND (empre.htethasta is null or empre.htethasta>=" & ConvFecha(FecEstr) & "))"
StrSql = StrSql & " AND empre.estrnro =" & Empresa

'Filtro de tres niveles de estructura
If TeNro1 <> 0 Then
    StrSql = StrSql & " INNER JOIN his_estructura tenro1 ON empleado.ternro = tenro1.ternro"
    StrSql = StrSql & " AND tenro1.tenro = " & TeNro1
    StrSql = StrSql & " AND tenro1.htetdesde <= " & ConvFecha(FecEstr) & " AND (tenro1.htethasta IS NULL OR tenro1.htethasta >= " & ConvFecha(FecEstr) & ")"
    If EstrNro1 <> 0 Then
        StrSql = StrSql & " AND tenro1.estrnro = " & EstrNro1
    End If
    StrSql = StrSql & " INNER JOIN tipoestructura tipoest1 ON tenro1.tenro = tipoest1.tenro"
    StrSql = StrSql & " INNER JOIN estructura estructura1 ON estructura1.estrnro = tenro1.estrnro"
    OrdenCorte = OrdenCorte & "tenro1,estrdabr1,"
End If

If TeNro2 <> 0 Then
    StrSql = StrSql & " INNER JOIN his_estructura tenro2 ON empleado.ternro = tenro2.ternro "
    StrSql = StrSql & " AND tenro2.tenro = " & TeNro2
    StrSql = StrSql & " AND tenro2.htetdesde <= " & ConvFecha(FecEstr) & " AND (tenro2.htethasta IS NULL OR tenro2.htethasta >= " & ConvFecha(FecEstr) & ") "
    If EstrNro2 <> 0 Then
        StrSql = StrSql & " AND tenro2.estrnro = " & EstrNro2
    End If
    StrSql = StrSql & " INNER JOIN tipoestructura tipoest2 ON tenro2.tenro = tipoest2.tenro"
    StrSql = StrSql & " INNER JOIN estructura estructura2 ON estructura2.estrnro = tenro2.estrnro "
    OrdenCorte = OrdenCorte & "tenro2,estrdabr2,"
End If

If TeNro3 <> 0 Then
    StrSql = StrSql & " INNER JOIN his_estructura tenro3 ON empleado.ternro = tenro3.ternro "
    StrSql = StrSql & " AND tenro3.tenro = " & TeNro3
    StrSql = StrSql & " AND tenro3.htetdesde <= " & ConvFecha(FecEstr) & " AND (tenro3.htethasta IS NULL OR tenro3.htethasta >= " & ConvFecha(FecEstr) & ") "
    If EstrNro3 <> 0 Then
        StrSql = StrSql & " AND tenro3.estrnro = " & EstrNro3
    End If
    StrSql = StrSql & " INNER JOIN tipoestructura tipoest3 ON tenro3.tenro = tipoest3.tenro"
    StrSql = StrSql & " INNER JOIN estructura estructura3 ON estructura3.estrnro = tenro3.estrnro "
    OrdenCorte = OrdenCorte & "tenro3,estrdabr3,"
End If
 
StrSql = StrSql & " WHERE " & FiltroSql

StrSql = StrSql & " ORDER BY " & OrdenCorte & "pre_linea.lnprenro, " & OrdenSql

OpenRecordset StrSql, rs_Consult

'Progreso
Progreso = 0
CEmpleadosAProc = rs_Consult.RecordCount
If CEmpleadosAProc = 0 Then
   CEmpleadosAProc = 1
End If
IncPorc = CEmpleadosAProc / 100


If Not rs_Consult.EOF Then

'Inicializacion de Variables
LineaAnt = rs_Consult!lnprenro
EmpleadoAnt = rs_Consult!ternro
MontoPrestamoAcum = 0
MontoCuotasAcum = 0
MontoCuotas = 0

'Guardo en varibles los datos a guardar en la rep para el caso del ultimo registro
TernroSQL = rs_Consult!ternro
LegajoSQL = rs_Consult!empleg
If EsNulo(rs_Consult!terape) Then ApellidoSQL = "" Else ApellidoSQL = rs_Consult!terape
If Not EsNulo(rs_Consult!terape2) Then ApellidoSQL = ApellidoSQL & " " & rs_Consult!terape2
If EsNulo(rs_Consult!ternom) Then NombreSQL = "" Else NombreSQL = rs_Consult!ternom
If Not EsNulo(rs_Consult!ternom2) Then NombreSQL = NombreSQL & " " & rs_Consult!ternom2
TpnroSQL = rs_Consult!tpnro
If EsNulo(rs_Consult!tpdesabr) Then TpdesabrSQL = "" Else TpdesabrSQL = rs_Consult!tpdesabr
LnprenroSQL = rs_Consult!lnprenro
If EsNulo(rs_Consult!lnpredabr) Then LnpredabrSQL = "" Else LnpredabrSQL = rs_Consult!lnpredabr
EstnroSQL = rs_Consult!estnro
If EsNulo(rs_Consult!estdabr) Then EstdabrSQL = "" Else EstdabrSQL = rs_Consult!estdabr
If TeNro1 <> 0 Then
    TeNro1SQL = rs_Consult!TeNro1
    If EsNulo(rs_Consult!tipoesttedabr1) Then Tenro1DescSQL = "" Else Tenro1DescSQL = rs_Consult!tipoesttedabr1
    EstrNro1SQL = rs_Consult!EstrNro1
    If EsNulo(rs_Consult!estrdabr1) Then Estrnro1DescSQL = "" Else Estrnro1DescSQL = rs_Consult!estrdabr1
Else
    TeNro1SQL = 0
    Tenro1DescSQL = ""
    EstrNro1SQL = 0
    Estrnro1DescSQL = ""
End If
If TeNro2 <> 0 Then
    TeNro2SQL = rs_Consult!TeNro2
    If EsNulo(rs_Consult!tipoesttedabr2) Then Tenro2DescSQL = "" Else Tenro2DescSQL = rs_Consult!tipoesttedabr2
    EstrNro2SQL = rs_Consult!EstrNro2
    If EsNulo(rs_Consult!estrdabr2) Then Estrnro2DescSQL = "" Else Estrnro2DescSQL = rs_Consult!estrdabr2
Else
    TeNro2SQL = 0
    Tenro2DescSQL = ""
    EstrNro2SQL = 0
    Estrnro2DescSQL = ""
End If
If TeNro3 <> 0 Then
    TeNro3SQL = rs_Consult!TeNro3
    If EsNulo(rs_Consult!tipoesttedabr3) Then Tenro3DescSQL = "" Else Tenro3DescSQL = rs_Consult!tipoesttedabr3
    EstrNro3SQL = rs_Consult!EstrNro3
    If EsNulo(rs_Consult!estrdabr3) Then Estrnro3DescSQL = "" Else Estrnro3DescSQL = rs_Consult!estrdabr3
Else
    TeNro3SQL = 0
    Tenro3DescSQL = ""
    EstrNro3SQL = 0
    Estrnro3DescSQL = ""
End If

    Do While Not rs_Consult.EOF
    
        Flog.writeline Espacios(Tabulador * 1) & "Procesando " & rs_Consult!empleg & " " & rs_Consult!terape & " " & rs_Consult!ternom & " prestamo " & rs_Consult!PreNro & " " & rs_Consult!predesc & " Importe " & rs_Consult!preimp
        
        'Control de corte para sumarizar
        If CambioCorte(LineaAnt, EmpleadoAnt, rs_Consult!lnprenro, rs_Consult!ternro) Then
        'Cambio de empleado o linea
            
            OrdenReg = OrdenReg + 1
            
            'Guardo en la rep de BD
            StrSql = " INSERT INTO rep_prestdeuda "
            StrSql = StrSql & " (bpronro,ternro,Legajo,"
            StrSql = StrSql & " apellido,Nombre,"
            StrSql = StrSql & " tpnro,tpdesabr,"
            StrSql = StrSql & " lnprenro,lnpredabr,"
            StrSql = StrSql & " estnro,estdabr,"
            StrSql = StrSql & " tenro1,tenro1Desc,"
            StrSql = StrSql & " estrnro1,estrnro1Desc,"
            StrSql = StrSql & " tenro2,tenro2Desc,"
            StrSql = StrSql & " estrnro2,estrnro2Desc,"
            StrSql = StrSql & " tenro3,tenro3Desc,"
            StrSql = StrSql & " estrnro3,estrnro3Desc,"
            StrSql = StrSql & " otorgado,amortizado,saldo,"
            StrSql = StrSql & " mes,anio,"
            StrSql = StrSql & " titulo, orden)"
            StrSql = StrSql & " VALUES"
            StrSql = StrSql & "(" & NroProcesoBatch
            StrSql = StrSql & "," & TernroSQL
            StrSql = StrSql & "," & LegajoSQL
            StrSql = StrSql & ",'" & Mid(ApellidoSQL, 1, 100) & "'"
            StrSql = StrSql & ",'" & Mid(NombreSQL, 1, 100) & "'"
            StrSql = StrSql & "," & TpnroSQL
            StrSql = StrSql & ",'" & Mid(TpdesabrSQL, 1, 40) & "'"
            StrSql = StrSql & "," & LnprenroSQL
            StrSql = StrSql & ",'" & Mid(LnpredabrSQL, 1, 30) & "'"
            StrSql = StrSql & "," & EstnroSQL
            StrSql = StrSql & ",'" & Mid(EstdabrSQL, 1, 30) & "'"
            StrSql = StrSql & "," & TeNro1SQL
            StrSql = StrSql & ",'" & Mid(Tenro1DescSQL, 1, 50) & "'"
            StrSql = StrSql & "," & EstrNro1SQL
            StrSql = StrSql & ",'" & Mid(Estrnro1DescSQL, 1, 50) & "'"
            StrSql = StrSql & "," & TeNro2SQL
            StrSql = StrSql & ",'" & Mid(Tenro2DescSQL, 1, 50) & "'"
            StrSql = StrSql & "," & EstrNro2SQL
            StrSql = StrSql & ",'" & Mid(Estrnro2DescSQL, 1, 50) & "'"
            StrSql = StrSql & "," & TeNro3SQL
            StrSql = StrSql & ",'" & Mid(Tenro3DescSQL, 1, 50) & "'"
            StrSql = StrSql & "," & EstrNro3SQL
            StrSql = StrSql & ",'" & Mid(Estrnro3DescSQL, 1, 50) & "'"
            StrSql = StrSql & "," & MontoPrestamoAcum
            StrSql = StrSql & "," & MontoCuotasAcum
            StrSql = StrSql & "," & MontoPrestamoAcum - MontoCuotasAcum
            StrSql = StrSql & "," & MesCuota
            StrSql = StrSql & "," & AnioCuota
            StrSql = StrSql & ",'" & Mid(Titulo, 1, 300) & "'"
            StrSql = StrSql & "," & OrdenReg & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            
            'Busco las cuotas del prestamo
            Call BuscarCuotasCanceladas(rs_Consult!PreNro, MesCuota, AnioCuota, MontoCuotas)
            
            'Inicilizo Acumuladores
            MontoCuotasAcum = MontoCuotas
            MontoPrestamoAcum = rs_Consult!preimp
            
            'Actualizo las variables de corte
            LineaAnt = rs_Consult!lnprenro
            EmpleadoAnt = rs_Consult!ternro
            
            'Guardo los datos para el caso del insert fuera del loop
            TernroSQL = rs_Consult!ternro
            LegajoSQL = rs_Consult!empleg
            If EsNulo(rs_Consult!terape) Then ApellidoSQL = "" Else ApellidoSQL = rs_Consult!terape
            If Not EsNulo(rs_Consult!terape2) Then ApellidoSQL = ApellidoSQL & " " & rs_Consult!terape2
            If EsNulo(rs_Consult!ternom) Then NombreSQL = "" Else NombreSQL = rs_Consult!ternom
            If Not EsNulo(rs_Consult!ternom2) Then NombreSQL = NombreSQL & " " & rs_Consult!ternom2
            TpnroSQL = rs_Consult!tpnro
            If EsNulo(rs_Consult!tpdesabr) Then TpdesabrSQL = "" Else TpdesabrSQL = rs_Consult!tpdesabr
            LnprenroSQL = rs_Consult!lnprenro
            If EsNulo(rs_Consult!lnpredabr) Then LnpredabrSQL = "" Else LnpredabrSQL = rs_Consult!lnpredabr
            EstnroSQL = rs_Consult!estnro
            If EsNulo(rs_Consult!estdabr) Then EstdabrSQL = "" Else EstdabrSQL = rs_Consult!estdabr
            If TeNro1 <> 0 Then
                TeNro1SQL = rs_Consult!TeNro1
                If EsNulo(rs_Consult!tipoesttedabr1) Then Tenro1DescSQL = "" Else Tenro1DescSQL = rs_Consult!tipoesttedabr1
                EstrNro1SQL = rs_Consult!EstrNro1
                If EsNulo(rs_Consult!estrdabr1) Then Estrnro1DescSQL = "" Else Estrnro1DescSQL = rs_Consult!estrdabr1
            Else
                TeNro1SQL = 0
                Tenro1DescSQL = ""
                EstrNro1SQL = 0
                Estrnro1DescSQL = ""
            End If
            If TeNro2 <> 0 Then
                TeNro2SQL = rs_Consult!TeNro2
                If EsNulo(rs_Consult!tipoesttedabr2) Then Tenro2DescSQL = "" Else Tenro2DescSQL = rs_Consult!tipoesttedabr2
                EstrNro2SQL = rs_Consult!EstrNro2
                If EsNulo(rs_Consult!estrdabr2) Then Estrnro2DescSQL = "" Else Estrnro2DescSQL = rs_Consult!estrdabr2
            Else
                TeNro2SQL = 0
                Tenro2DescSQL = ""
                EstrNro2SQL = 0
                Estrnro2DescSQL = ""
            End If
            If TeNro3 <> 0 Then
                TeNro3SQL = rs_Consult!TeNro3
                If EsNulo(rs_Consult!tipoesttedabr3) Then Tenro3DescSQL = "" Else Tenro3DescSQL = rs_Consult!tipoesttedabr3
                EstrNro3SQL = rs_Consult!EstrNro3
                If EsNulo(rs_Consult!estrdabr3) Then Estrnro3DescSQL = "" Else Estrnro3DescSQL = rs_Consult!estrdabr3
            Else
                TeNro3SQL = 0
                Tenro3DescSQL = ""
                EstrNro3SQL = 0
                Estrnro3DescSQL = ""
            End If
            
        
        Else
        'Sigue en el mismo empleado y linea
            
            'Busco las cuotas del prestamo
            Call BuscarCuotasCanceladas(rs_Consult!PreNro, MesCuota, AnioCuota, MontoCuotas)
            
            'Acumulo
            MontoCuotasAcum = MontoCuotasAcum + MontoCuotas
            MontoPrestamoAcum = MontoPrestamoAcum + rs_Consult!preimp
        
        End If
        
        rs_Consult.MoveNext
        
        
        'Actualizo el progreso
        Progreso = Progreso + IncPorc
        TiempoAcumulado = GetTickCount
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
                 ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
                 "' WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
        
        
    Loop
    
    rs_Consult.Close
    
    OrdenReg = OrdenReg + 1
    
    'Insertar el ultimo registro
    StrSql = " INSERT INTO rep_prestdeuda "
    StrSql = StrSql & " (bpronro,ternro,Legajo,"
    StrSql = StrSql & " apellido,Nombre,"
    StrSql = StrSql & " tpnro,tpdesabr,"
    StrSql = StrSql & " lnprenro,lnpredabr,"
    StrSql = StrSql & " estnro,estdabr,"
    StrSql = StrSql & " tenro1,tenro1Desc,"
    StrSql = StrSql & " estrnro1,estrnro1Desc,"
    StrSql = StrSql & " tenro2,tenro2Desc,"
    StrSql = StrSql & " estrnro2,estrnro2Desc,"
    StrSql = StrSql & " tenro3,tenro3Desc,"
    StrSql = StrSql & " estrnro3,estrnro3Desc,"
    StrSql = StrSql & " otorgado,amortizado,saldo,"
    StrSql = StrSql & " mes,anio,"
    StrSql = StrSql & " titulo, orden)"
    StrSql = StrSql & " VALUES"
    StrSql = StrSql & "(" & NroProcesoBatch
    StrSql = StrSql & "," & TernroSQL
    StrSql = StrSql & "," & LegajoSQL
    StrSql = StrSql & ",'" & Mid(ApellidoSQL, 1, 100) & "'"
    StrSql = StrSql & ",'" & Mid(NombreSQL, 1, 100) & "'"
    StrSql = StrSql & "," & TpnroSQL
    StrSql = StrSql & ",'" & Mid(TpdesabrSQL, 1, 40) & "'"
    StrSql = StrSql & "," & LnprenroSQL
    StrSql = StrSql & ",'" & Mid(LnpredabrSQL, 1, 30) & "'"
    StrSql = StrSql & "," & EstnroSQL
    StrSql = StrSql & ",'" & Mid(EstdabrSQL, 1, 30) & "'"
    StrSql = StrSql & "," & TeNro1SQL
    StrSql = StrSql & ",'" & Mid(Tenro1DescSQL, 1, 50) & "'"
    StrSql = StrSql & "," & EstrNro1SQL
    StrSql = StrSql & ",'" & Mid(Estrnro1DescSQL, 1, 50) & "'"
    StrSql = StrSql & "," & TeNro2SQL
    StrSql = StrSql & ",'" & Mid(Tenro2DescSQL, 1, 50) & "'"
    StrSql = StrSql & "," & EstrNro2SQL
    StrSql = StrSql & ",'" & Mid(Estrnro2DescSQL, 1, 50) & "'"
    StrSql = StrSql & "," & TeNro3SQL
    StrSql = StrSql & ",'" & Mid(Tenro3DescSQL, 1, 50) & "'"
    StrSql = StrSql & "," & EstrNro3SQL
    StrSql = StrSql & ",'" & Mid(Estrnro3DescSQL, 1, 50) & "'"
    StrSql = StrSql & "," & MontoPrestamoAcum
    StrSql = StrSql & "," & MontoCuotasAcum
    StrSql = StrSql & "," & MontoPrestamoAcum - MontoCuotasAcum
    StrSql = StrSql & "," & MesCuota
    StrSql = StrSql & "," & AnioCuota
    StrSql = StrSql & ",'" & Mid(Titulo, 1, 300) & "'"
    StrSql = StrSql & "," & OrdenReg & ")"
    objConn.Execute StrSql, , adExecuteNoRecords

            
Else
    Flog.writeline Espacios(Tabulador * 1) & "No se encontraron prestamos para el filtro."
End If

Fin:

'Cierro todo y libero
If rs_Consult.State = adStateOpen Then rs_Consult.Close
Set rs_Consult = Nothing

Exit Sub

E_Generar_Reporte:
    
    HuboError = True
    Flog.writeline
    Flog.writeline "**********************************************************"
    Flog.writeline " Error en Generar_Reporte: "
    Flog.writeline " Error: " & Err.Description
    Flog.writeline "**********************************************************"
    Flog.writeline
End Sub


Public Function CambioCorte(ByVal LineaAnterior As Long, ByVal EmpleadoAnterior As Long, ByVal LineaActual As Long, ByVal EmpleadoActual As Long) As Boolean
    
    If ((LineaAnterior <> LineaActual) Or (EmpleadoAnterior <> EmpleadoActual)) Then
        CambioCorte = True
    Else
        CambioCorte = False
    End If
    
End Function


Public Sub BuscarCuotasCanceladas(ByVal PreNro As Long, ByVal mes As Long, ByVal Anio As Long, ByRef Monto As Double)

Dim rs_Cuota As New ADODB.Recordset
Dim ImporteAux As Double

On Error GoTo E_BuscarCuotasCanceladas

    ImporteAux = 0

    StrSql = "SELECT SUM(cuoimp) importe"
    StrSql = StrSql & " FROM pre_cuota"
    StrSql = StrSql & " WHERE PreNro = " & PreNro
    StrSql = StrSql & " And cuocancela = -1"
    StrSql = StrSql & " And ("
    StrSql = StrSql & " (cuomes <= " & mes & " And cuoano = " & Anio & ")"
    StrSql = StrSql & " OR"
    StrSql = StrSql & " cuoano < " & Anio
    StrSql = StrSql & " )"
    OpenRecordset StrSql, rs_Cuota
    If Not rs_Cuota.EOF Then
        If Not EsNulo(rs_Cuota!importe) Then ImporteAux = rs_Cuota!importe
    End If
    
    Flog.writeline Espacios(Tabulador * 2) & "Monto Cuotas Canceladas " & ImporteAux
    Monto = ImporteAux

If rs_Cuota.State = adStateOpen Then rs_Cuota.Close
Set rs_Cuota = Nothing

Exit Sub

E_BuscarCuotasCanceladas:
    
    HuboError = True
    Flog.writeline
    Flog.writeline "**********************************************************"
    Flog.writeline " Error en BuscarCuotasCanceladas: "
    Flog.writeline " Error: " & Err.Description
    Flog.writeline "**********************************************************"
    Flog.writeline

End Sub
