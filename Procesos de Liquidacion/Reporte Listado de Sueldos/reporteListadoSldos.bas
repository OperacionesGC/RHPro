Attribute VB_Name = "repListadoSldos"
Option Explicit

Global Const Version = "1.01"
Global Const FechaModificacion = "20/09/2006"
Global Const UltimaModificacion = " " 'Martin Ferraro - en todas las estructuras se agrego buscar por fecha
                                      'hasta del periodo


Public Type TColumnas
    TipoOrigen As String
    Origen As Long
    OrigenTxt As String
    Valor As Double
End Type


Dim fs, f
Global Flog

Dim NroLinea As Long
Dim crpNro As Long
Dim RegLeidos As Long
Dim RegError As Long
Dim RegFecha As Date
Dim NroProceso As Long

Global Path As String
Global NArchivo As String
Global Rta
Global HuboErrores As Boolean
Global EmpErrores As Boolean

Global pliqnro1 As Integer
Global pliqdesc1 As String
Global pliqmesanio1 As String
Global pronro1 As String
Global pliqnro2 As Integer
Global pliqdesc2 As String
Global pliqmesanio2 As String
Global pronro2 As String
Global listaconcnro As String
Global Tabulador As Long

Global IdUser As String
Global Fecha As Date
Global Hora As String

Private Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento Principal.
' Autor      : FGZ
' Fecha      : 29/07/2005
' Ultima Mod :
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim strCmdLine As String
Dim Nombre_Arch As String
Dim rs_Batch_Proceso As New ADODB.Recordset
Dim PID As String
Dim Parametros As String
Dim ArrParametros

    strCmdLine = Command()
    ArrParametros = Split(strCmdLine, " ", -1)
    If UBound(ArrParametros) > 0 Then
        If IsNumeric(ArrParametros(0)) Then
            NroProceso = ArrParametros(0)
            Etiqueta = ArrParametros(1)
        Else
            Exit Sub
        End If
    Else
        If IsNumeric(strCmdLine) Then
            NroProceso = strCmdLine
        Else
            Exit Sub
        End If
    End If
    
    On Error GoTo M_Error
    
    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas

    TiempoInicialProceso = GetTickCount
    OpenConnection strconexion, objConn
    OpenConnection strconexion, objconnProgreso
    
    HuboErrores = False
    
    Nombre_Arch = PathFLog & "ReporteListadoSldos" & "-" & NroProceso & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Version = " & Version
    Flog.writeline "Modificacion = " & UltimaModificacion
    Flog.writeline "Fecha = " & FechaModificacion
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline
    
    Flog.writeline "Inicio Proceso: " & Now
    Flog.writeline "Cambio el estado del proceso a Procesando"
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "PID = " & PID
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & ", bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Obtengo los datos del proceso"
    TiempoAcumulado = GetTickCount
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 100 AND bpronro = " & NroProceso
    OpenRecordset StrSql, rs_Batch_Proceso
    
    If Not rs_Batch_Proceso.EOF Then
        IdUser = rs_Batch_Proceso!IdUser
        Fecha = rs_Batch_Proceso!bprcfecha
        Hora = rs_Batch_Proceso!bprchora
        Parametros = rs_Batch_Proceso!bprcparam
        rs_Batch_Proceso.Close
        Set rs_Batch_Proceso = Nothing
        
        Call LevantarParamteros(NroProcesoBatch, Parametros)
    End If
    
    TiempoFinalProceso = GetTickCount
    Flog.writeline "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    
    'Actualizo el estado del proceso
    If Not HuboErrores Then
       StrSql = "UPDATE batch_proceso SET  bprcprogreso = 100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
       Flog.writeline "Proceso Finalizado Correctamente"
    Else
       StrSql = "UPDATE batch_proceso SET  bprcprogreso = 100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProceso
       Flog.writeline "Proceso Incompleto"
    End If
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Fin :" & Now
    Flog.Close
    
    
    objconnProgreso.Close
    objConn.Close
    If rs_Batch_Proceso.State = adStateOpen Then rs_Batch_Proceso.Close
    Set rs_Batch_Proceso = Nothing
    Set objConn = Nothing
    Set objconnProgreso = Nothing
    
    Exit Sub

M_Error:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 0) & " Error " & Err.Description
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "SQL Ejecutado: " & StrSql
        Flog.writeline
    End If
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline
        
    'Actualizo el progreso
    MyBeginTrans
        StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Error', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
    MyCommitTrans
    
    objconnProgreso.Close
    objConn.Close
    If rs_Batch_Proceso.State = adStateOpen Then rs_Batch_Proceso.Close
    Set rs_Batch_Proceso = Nothing
    Set objConn = Nothing
    Set objconnProgreso = Nothing
End Sub


Public Sub LevantarParamteros(ByVal bpronro As Long, ByVal Parametros As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento para levantar los parametros pasados en batch_proceso en bprcparam
' Autor      : FGZ
' Fecha      : 29/07/2005
' Modificado :
'       28/10/05 - Fapitalle N. - Agregado Puestos Agrupados (pagrup, pagrupdesc, pagrupcodext)
' --------------------------------------------------------------------------------------------
Dim ArrParametros
Dim Periodo As Long
Dim Proceso As String
Dim Todos As Boolean
Dim Aprobados As Integer
Dim Empresa As Long
Dim Sucursal As Long
Dim Sector As Long
Dim CCosto As Long
Dim Puesto As Long
Dim PAgrup As Long

'Orden de los parametros
    'pliqnro
    'pronro
    'todospro
    'proaprob
    'empresa
    'sucursal
    'sector
    'centro de costo
    'puesto
    'puestos agrupados
 
ArrParametros = Split(Parametros, "@")
Periodo = ArrParametros(0)
Proceso = IIf(EsNulo(ArrParametros(1)), 0, ArrParametros(1))
Todos = IIf(EsNulo(ArrParametros(2)), False, CBool(ArrParametros(2)))
Aprobados = ArrParametros(3)
Empresa = ArrParametros(4)
Sucursal = ArrParametros(5)
Sector = ArrParametros(6)
CCosto = ArrParametros(7)
Puesto = ArrParametros(8)
PAgrup = ArrParametros(9)
        
Call GenerarDatos(Periodo, Proceso, Todos, Aprobados, Empresa, Sucursal, Sector, CCosto, Puesto, PAgrup)

End Sub


Sub GenerarDatos(ByVal Periodo As Long, ByVal Proceso As String, ByVal Todos As Boolean, _
    ByVal Aprobados As Integer, ByVal Empresa As Long, ByVal Sucursal As Long, ByVal Sector As Long, ByVal CCosto As Long, ByVal Puesto As Long, ByVal PAgrup As Long)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que genera los datos para el reporte.
' Autor      : FGZ
' Fecha      : 29/07/2005
' Ult. Mod   : 20-10-2005 - Leticia A. - se agregaron los cod ext. de las extructuras
'               28/10/2005 - Fapitalle N. - se agrega la estructura Puestos Agrupados y sus campos
' --------------------------------------------------------------------------------------------
Dim I As Long
Dim Cantidad As Integer
Dim CantidadProcesada As Integer

Dim Columnas(21) As TColumnas
Dim Pliqdesc As String
Dim Legajo As Long
Dim ApeyNom As String
Dim Txt_Sucursal As String
Dim Txt_Sector As String
Dim Txt_CCosto As String
Dim Txt_Puesto As String
Dim Txt_PAgrup As String
Dim CodExt_Sucursal As String
Dim CodExt_Sector As String
Dim CodExt_CCosto As String
Dim CodExt_Puesto As String
Dim CodExt_PAgrup As String

Dim EmpNom As String
Dim EmpDirec As String
Dim EmpLogo As String
Dim EmpLogoalto As String
Dim EmpLogoancho As String

Dim Indice_Subtotal As Integer
Dim Indice_Total As Integer

Dim l_Porc_tkt As Double
Dim l_Bases As Double
Dim Subtotal As Double
Dim Total As Double
Dim Porcentaje As Double

Dim PeriodoDesde As Date
Dim PeriodoHasta As Date

Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset
Dim rsE As New ADODB.Recordset
Dim rs_Listado As New ADODB.Recordset
Dim v_redo As Double
Dim cantidadCeros

Dim hayNovedad As Boolean

On Error GoTo MError

'Período a considerar (para mostrar en el título)
Pliqdesc = ""
If Not EsNulo(Periodo) Then
    StrSql = "SELECT pliqdesc,pliqdesde, pliqhasta "
    StrSql = StrSql & "FROM periodo "
    StrSql = StrSql & "WHERE pliqnro = " & Periodo
    If rs.State = adStateOpen Then rs.Close
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
       Pliqdesc = rs!Pliqdesc
       PeriodoDesde = rs!pliqdesde
       PeriodoHasta = rs!pliqhasta
    Else
        Flog.writeline Espacios(Tabulador * 1) & "No se encontró el Periodo " & Periodo
        Exit Sub
    End If
End If


'Depuro los datos del proceso
Flog.writeline Espacios(Tabulador * 1) & "Depuro los datos del proceso"
StrSql = "DELETE FROM rep_listado_sldo WHERE bpronro = " & NroProceso
objConn.Execute StrSql, , adExecuteNoRecords

'consulta principal

StrSql = "SELECT DISTINCT empleg,terape,ternom,cabliq.cliqnro,e3.estrnro as suc_nro,"
StrSql = StrSql & "e3.estrdabr as suc_desc, e3.estrcodext as suc_codext, e2.estrnro as sec_nro,e2.estrdabr as sec_desc , e2.estrcodext as sec_codext,"
StrSql = StrSql & "e4.estrnro as ccos_nro,e4.estrdabr as ccos_desc ,e4.estrcodext as ccos_codext, e5.estrnro as pue_nro,"
StrSql = StrSql & "e5.estrdabr as pue_desc, e5.estrcodext as pue_codext,  empleado.ternro, "
StrSql = StrSql & "e6.estrnro as pag_nro, e6.estrdabr as pag_desc, e6.estrcodext as pag_codext "
StrSql = StrSql & " FROM empleado "
StrSql = StrSql & " INNER JOIN cabliq ON cabliq.empleado = empleado.ternro "
StrSql = StrSql & " inner join detliq on cabliq.cliqnro = detliq.cliqnro "
StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro "
'StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro "
'Empleados de la empresa
StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.ternro = empleado.ternro AND his_estructura.tenro = 10 "
'MAF - 20/09/2006 - Buscar estr a la fecha
'StrSql = StrSql & " AND his_estructura.htethasta IS NULL "
StrSql = StrSql & " AND (his_estructura.htetdesde<=" & ConvFecha(PeriodoHasta) & " AND (his_estructura.htethasta is null or his_estructura.htethasta>=" & ConvFecha(PeriodoHasta) & "))"
StrSql = StrSql & " AND his_estructura.estrnro = " & Empresa
'sector
StrSql = StrSql & " INNER JOIN his_estructura he2 ON he2.ternro = empleado.ternro AND he2.tenro = 2 "
'MAF - 20/09/2006 - Buscar estr a la fecha
'StrSql = StrSql & " and he2.htethasta IS NULL "
StrSql = StrSql & " AND (he2.htetdesde<=" & ConvFecha(PeriodoHasta) & " AND (he2.htethasta is null or he2.htethasta>=" & ConvFecha(PeriodoHasta) & "))"
'suc
StrSql = StrSql & " INNER JOIN his_estructura he3 ON he3.ternro = empleado.ternro AND he3.tenro = 1 "
'MAF - 20/09/2006 - Buscar estr a la fecha
'StrSql = StrSql & " and he3.htethasta IS NULL "
StrSql = StrSql & " AND (he3.htetdesde<=" & ConvFecha(PeriodoHasta) & " AND (he3.htethasta is null or he3.htethasta>=" & ConvFecha(PeriodoHasta) & "))"
'centro
StrSql = StrSql & " INNER JOIN his_estructura he4 ON he4.ternro = empleado.ternro AND he4.tenro = 5 "
'MAF - 20/09/2006 - Buscar estr a la fecha
'StrSql = StrSql & " and he4.htethasta IS NULL "
StrSql = StrSql & " AND (he4.htetdesde<=" & ConvFecha(PeriodoHasta) & " AND (he4.htethasta is null or he4.htethasta>=" & ConvFecha(PeriodoHasta) & "))"
'puesto
StrSql = StrSql & " INNER JOIN his_estructura he5 ON he5.ternro = empleado.ternro AND he5.tenro = 4 "
'MAF - 20/09/2006 - Buscar estr a la fecha
'StrSql = StrSql & " and he5.htethasta IS NULL "
StrSql = StrSql & " AND (he5.htetdesde<=" & ConvFecha(PeriodoHasta) & " AND (he5.htethasta is null or he5.htethasta>=" & ConvFecha(PeriodoHasta) & "))"
'PuestosAgrupados
StrSql = StrSql & " INNER JOIN his_estructura he6 ON he6.ternro = empleado.ternro AND he6.tenro = 52 "
'MAF - 20/09/2006 - Buscar estr a la fecha
'StrSql = StrSql & " and he6.htethasta IS NULL "
StrSql = StrSql & " AND (he6.htetdesde<=" & ConvFecha(PeriodoHasta) & " AND (he6.htethasta is null or he6.htethasta>=" & ConvFecha(PeriodoHasta) & "))"

StrSql = StrSql & " INNER JOIN estructura e2 ON he2.estrnro = e2.estrnro "
If Sector <> 0 Then
   StrSql = StrSql & " AND e2.estrnro = " & Sector
End If
StrSql = StrSql & " INNER JOIN estructura e3 ON he3.estrnro = e3.estrnro "
If Sucursal <> 0 Then
   StrSql = StrSql & " AND e3.estrnro = " & Sucursal
End If
StrSql = StrSql & " INNER JOIN estructura e4 ON he4.estrnro = e4.estrnro "
If CCosto <> 0 Then
   StrSql = StrSql & " AND e4.estrnro = " & CCosto
End If
StrSql = StrSql & " INNER JOIN estructura e5 ON he5.estrnro = e5.estrnro "
If Puesto <> 0 Then
   StrSql = StrSql & " AND e5.estrnro   = " & Puesto
End If
StrSql = StrSql & " INNER JOIN estructura e6 ON he6.estrnro = e6.estrnro "
If PAgrup <> 0 Then
   StrSql = StrSql & " AND e6.estrnro   = " & PAgrup
End If


'If Periodo <> "" Then
    StrSql = StrSql & " WHERE proceso.pliqnro   = " & Periodo
'End If
If Proceso <> "" And Proceso <> 0 Then
   StrSql = StrSql & " AND proceso.pronro   = " & Proceso
End If
StrSql = StrSql & " AND proceso.proaprob = " & Aprobados
StrSql = StrSql & " ORDER BY suc_desc,sec_desc,ccos_desc,pue_desc,pag_desc,empleg"

Flog.writeline "Consulta Empleados: " & StrSql

If rs.State = adStateOpen Then rs.Close
OpenRecordset StrSql, rs
If rs.EOF Then
    Flog.writeline Espacios(Tabulador * 1) & "No se encontraron Datos"
Else
    StrSql = "SELECT calle,nro,locdesc,empnom,codigopostal, ter_imag.terimnombre, "
    StrSql = StrSql & " tipoimag.tipimdire, tipoimag.tipimanchodef, tipoimag.tipimaltodef"
    StrSql = StrSql & " FROM empresa "
    StrSql = StrSql & " LEFT JOIN ter_imag ON ter_imag.ternro = empresa.ternro "
    StrSql = StrSql & " LEFT JOIN tipoimag ON tipoimag.tipimnro = ter_imag.tipimnro "
    StrSql = StrSql & " LEFT JOIN cabdom ON cabdom.ternro = empresa.ternro "
    StrSql = StrSql & " LEFT JOIN detdom ON detdom.domnro = cabdom.domnro "
    StrSql = StrSql & " LEFT JOIN localidad ON detdom.locnro = localidad.locnro "
    StrSql = StrSql & " WHERE empresa.estrnro = " & Empresa & " AND tipoimag.tipimnro = 1 "
    If rs1.State = adStateOpen Then rs1.Close
    OpenRecordset StrSql, rs1
    If Not rs1.EOF Then
        EmpNom = Left(rs1("empnom"), 60)
        EmpDirec = Left(rs1("calle") & " " & rs1("nro") & ", " & rs1("locdesc") & " " & rs1("codigopostal"), 60)
        EmpLogo = Left(rs1("tipimdire") & rs1("terimnombre"), 100)
        EmpLogoalto = CInt(rs1("tipimaltodef"))
        EmpLogoancho = CInt(rs1("tipimanchodef"))
    Else
        Flog.writeline Espacios(Tabulador * 1) & "No se encontraron Datos de la empresa " & Empresa
    End If
    rs1.Close

End If

CEmpleadosAProc = rs.RecordCount
If CEmpleadosAProc = 0 Then
    CEmpleadosAProc = 1
End If

IncPorc = 99 / CEmpleadosAProc
Progreso = 0

Flog.writeline "Cantidad de empleados a procesar: " & rs.RecordCount
   
Do Until rs.EOF
     For I = 1 To 20
        Columnas(I).Valor = 0
        Columnas(I).OrigenTxt = ""
        Columnas(I).Origen = 0
        Columnas(I).TipoOrigen = ""
     Next

    'busco los datos del empleado
    Legajo = rs!empleg
    ApeyNom = Left(rs!terape & ", " & rs!ternom, 100)
    l_Bases = 0
    Flog.writeline " "
    Flog.writeline " "
    Flog.writeline "Procesando empleado " & Legajo & ":" & ApeyNom
    'Busco las descripciones de las estructuras
    'Sucursal
    Sucursal = rs!suc_nro
    Txt_Sucursal = Left(rs!suc_desc, 60)
    CodExt_Sucursal = Left(rs!suc_codext, 20)
    
    'Sector
    Sector = rs!sec_nro
    Txt_Sector = Left(rs!sec_desc, 60)
    CodExt_Sector = Left(rs!sec_codext, 20)

    'Centro de Costo
    CCosto = rs!ccos_nro
    Txt_CCosto = Left(rs!ccos_desc, 60)
    CodExt_CCosto = IIf(EsNulo(rs!ccos_codext), "", Left(rs!ccos_codext, 20))
    
    'Puesto
    Puesto = rs!pue_nro
    Txt_Puesto = Left(rs!pue_desc, 60)
    CodExt_Puesto = Left(rs!pue_codext, 20)
    
    'Puestos Agrupados
    PAgrup = rs!pag_nro
    Txt_PAgrup = Left(rs!pag_desc, 60)
    CodExt_PAgrup = Left(rs!pag_codext, 20)
    
    'Busco las columnas configuradas en el Confrep
    StrSql = "SELECT confrep.confnrocol, confrep.confetiq, reporte.repnro, conftipo, confval, confval2"
    StrSql = StrSql & " FROM reporte"
    StrSql = StrSql & " INNER JOIN confrep ON reporte.repnro = confrep.repnro"
    StrSql = StrSql & " WHERE (reporte.repnro = 85)"
    StrSql = StrSql & " AND confnrocol <= 20"
    StrSql = StrSql & " order by confnrocol"
    If rs1.State = adStateOpen Then rs1.Close
    OpenRecordset StrSql, rs1
   
    Do While Not rs1.EOF
        
        Select Case UCase(rs1("conftipo"))
        Case "CO":
            Columnas(rs1("confnrocol")).TipoOrigen = "CO"
            If Not EsNulo(rs1("confval2")) Then
                Columnas(rs1("confnrocol")).OrigenTxt = rs1("confval2")
            Else
                Columnas(rs1("confnrocol")).OrigenTxt = rs1("confval")
            End If
            
            StrSql = "SELECT concepto.concnro, dlimonto"
            StrSql = StrSql & " FROM concepto"
            StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro =" & rs("cliqnro") & " AND detliq.concnro = concepto.concnro"
            StrSql = StrSql & " WHERE concepto.conccod = " & rs1("confval")
            StrSql = StrSql & " OR concepto.conccod = '" & rs1("confval2") & "'"
            
            Flog.writeline "Consulta Columna " & rs1("confnrocol") & ": " & StrSql
            
            If rs2.State = adStateOpen Then rs2.Close
            OpenRecordset StrSql, rs2
            If Not rs2.EOF Then
                Columnas(rs1("confnrocol")).Origen = rs2("concnro")
                Columnas(rs1("confnrocol")).Valor = Columnas(rs1("confnrocol")).Valor + CDbl(rs2("dlimonto"))
                Flog.writeline "Valor Columna " & rs1("confnrocol") & ": " & rs2("dlimonto")
            End If
             rs2.Close
        Case "SUB":
            Indice_Subtotal = rs1("confnrocol")
        Case "TOT":
            Indice_Total = rs1("confnrocol")
        Case Else
        
        End Select
        
        rs1.MoveNext
    Loop

    '------------------------------------------------------------------------------
    'Calculo de columnas especiales
    '------------------------------------------------------------------------------
    ' Columna 5 - Subtotal = Columnas (1 + 2 + 3 + 4)
    
    ' Columna 6 - Sueldo Basico
    'verificar si no hay ninguna novedad en el parámetro Coeficiente (1) del concepto 0005
    'si hubiera entonces habría que considerar esa novedad cargada como el coeficiente
    'sino el Coeficiente es el default 0.9091%
    'Busco las columnas configuradas en el Confrep
    StrSql = "SELECT confrep.confnrocol, confrep.confetiq, reporte.repnro, conftipo, confval, confval2"
    StrSql = StrSql & " FROM reporte"
    StrSql = StrSql & " INNER JOIN confrep ON reporte.repnro = confrep.repnro"
    StrSql = StrSql & " WHERE (reporte.repnro = 85)"
    StrSql = StrSql & " AND confnrocol = 6 AND conftipo = 'PAR'"
    StrSql = StrSql & " order by confnrocol"
    If rs1.State = adStateOpen Then rs1.Close
    OpenRecordset StrSql, rs1
    If Not rs1.EOF Then
        'el valor numerico tiene el nro de parametro y
        'el valor alfanumerico tiene el conccod del concepto
        Columnas(rs1("confnrocol")).TipoOrigen = "PAR"
        Columnas(rs1("confnrocol")).Origen = rs1("confval")
        
        StrSql = "SELECT concepto.concnro, nevalor, nedesde, nehasta "
        StrSql = StrSql & " FROM concepto"
        StrSql = StrSql & " INNER JOIN novemp ON concepto.concnro = novemp.concnro AND novemp.empleado = " & rs("ternro")
        StrSql = StrSql & " WHERE concepto.conccod = '" & rs1("confval2") & "'"
        StrSql = StrSql & " AND novemp.tpanro = " & rs1("confval")
        
        Flog.writeline "Consulta Columna 6: " & StrSql
        
        If rs2.State = adStateOpen Then rs2.Close
        OpenRecordset StrSql, rs2
        
        l_Porc_tkt = 0.9091
        
        Do Until rs2.EOF
            If vigenciaValida(rs2, PeriodoDesde, PeriodoHasta) Then
               Flog.writeline Espacios(Tabulador * 1) & "Coeficiente del Basico. Concepto " & rs1("confval2") & " Parametro " & rs1("confval") & ". "
               l_Porc_tkt = CDbl(rs2("nevalor"))
            End If
            
            rs2.MoveNext
        Loop
            
        Flog.writeline Espacios(Tabulador * 1) & " Se usa el valor " & l_Porc_tkt
                
        If rs2.State = adStateOpen Then rs2.Close
        'Columnas(rs1("confnrocol")).Valor = l_Porc_tkt
    Else
        Flog.writeline Espacios(Tabulador * 1) & "No esta configurado el Coeficiente del Basico. Columna 6 Tipo PAR. Se usa el Default 0.9091 "
        l_Porc_tkt = 0.9091
    End If
    Columnas(6).Valor = CDbl((Columnas(1).Valor * l_Porc_tkt))
     
     
    ' Columna 4 - Presentismo
    'verificar si no hay ninguna novedad en el parámetro Múltiplo (1024) del concepto 013
    'si hubiera entonces habría que considerar esa novedad cargada como el porcentaje
    'sino el porcentaje es el default 12%
    'Busco las columnas configuradas en el Confrep
    StrSql = "SELECT confrep.confnrocol, confrep.confetiq, reporte.repnro, conftipo, confval, confval2"
    StrSql = StrSql & " FROM reporte"
    StrSql = StrSql & " INNER JOIN confrep ON reporte.repnro = confrep.repnro"
    StrSql = StrSql & " WHERE (reporte.repnro = 85)"
    StrSql = StrSql & " AND confnrocol = 4 AND conftipo = 'PAR'"
    StrSql = StrSql & " order by confnrocol"
    If rs1.State = adStateOpen Then rs1.Close
    OpenRecordset StrSql, rs1
    If Not rs1.EOF Then
        'el valor numerico tiene el nro de parametro y
        'el valor alfanumerico tiene el conccod del concepto
        Columnas(rs1("confnrocol")).TipoOrigen = "PAR"
        Columnas(rs1("confnrocol")).Origen = rs1("confval")
        
        StrSql = "SELECT concepto.concnro, nevalor, nedesde, nehasta "
        StrSql = StrSql & " FROM concepto"
        StrSql = StrSql & " INNER JOIN novemp ON concepto.concnro = novemp.concnro AND novemp.empleado = " & rs("ternro")
        StrSql = StrSql & " WHERE concepto.conccod = '" & rs1("confval2") & "'"
        StrSql = StrSql & " AND novemp.tpanro = " & rs1("confval")
        
        Flog.writeline "Consulta Columna 4: " & StrSql
        
        If rs2.State = adStateOpen Then rs2.Close
        OpenRecordset StrSql, rs2
        
        Porcentaje = 12
        
        Do Until rs2.EOF
            If vigenciaValida(rs2, PeriodoDesde, PeriodoHasta) Then
               Flog.writeline Espacios(Tabulador * 1) & "Porcentaje de Presentismo. Concepto " & rs1("confval2") & " Parametro " & rs1("confval") & "."
               Porcentaje = CDbl(rs2("nevalor"))
            End If
            rs2.MoveNext
        Loop
            
        Flog.writeline Espacios(Tabulador * 1) & "Se usa el valor " & Porcentaje
                
        If rs2.State = adStateOpen Then rs2.Close
        'Columnas(rs1("confnrocol")).Valor = Porcentaje
    Else
        Flog.writeline Espacios(Tabulador * 1) & "No esta configurado el Porcentaje de Presentismo. Columna 4 Tipo PAR. Se usa el Default 12% "
        Porcentaje = 12
    End If
    'si el legajo es administrativo tiene un calculo distinto
    StrSql = " Select his_estructura.estrnro "
    StrSql = StrSql & " From his_estructura "
    StrSql = StrSql & " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro "
    StrSql = StrSql & " where his_estructura.tenro = 29 " 'ACTIVIDAD
    'MAF - 20/09/2006 - Buscar estr a la fecha
    'StrSql = StrSql & " AND his_estructura.htethasta IS NULL "
    StrSql = StrSql & " AND (his_estructura.htetdesde<=" & ConvFecha(PeriodoHasta) & " AND (his_estructura.htethasta is null or his_estructura.htethasta>=" & ConvFecha(PeriodoHasta) & "))"
    'StrSql = StrSql & " AND his_estructura.estrnro = 556 " 'Administrativo
    StrSql = StrSql & " AND estructura.estrcodext = '5'" 'Administrativo
    StrSql = StrSql & " AND his_estructura.ternro = " & rs("ternro")
    
    Flog.writeline "Consulta es Administrativo: " & StrSql
    
    If rs3.State = adStateOpen Then rs3.Close
    OpenRecordset StrSql, rs3
    If Not rs3.EOF Then 'la actividad es 5
        Columnas(4).Valor = CDbl((Columnas(6).Valor + Columnas(3).Valor)) * Porcentaje / 100
    Else
        Columnas(4).Valor = 0
    End If
    rs3.Close
     
    ' Columna 7 - Tickets
    'Primero preguntar si grupo de liquidación (estructura 32) = "1" (personal)
    'para que aparezcan solo los que tienen tickets.
    '   Luego si hay alguna novedad en el parámetro % (35) del concepto 449
    '       aplicarla sobre (Col. 3 + Col. 4 + Col. 6 ) x esta novedad
    '       sino sería (Col. 3 + Col. 4 + Col. 6) x 0,10  (es decir x 10%)
    l_Bases = l_Bases + Columnas(3).Valor + Columnas(4).Valor + Columnas(6).Valor
    'si el legajo es de Grupo de Liquidacion Personal tiene un calculo distinto
    StrSql = " Select his_estructura.estrnro "
    StrSql = StrSql & " From his_estructura "
    StrSql = StrSql & " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro "
    StrSql = StrSql & " where his_estructura.tenro = 32 " 'Grupo de Liquidacion
    'MAF - 20/09/2006 - Buscar estr a la fecha
    'StrSql = StrSql & " AND his_estructura.htethasta IS NULL "
    StrSql = StrSql & " AND (his_estructura.htetdesde<=" & ConvFecha(PeriodoHasta) & " AND (his_estructura.htethasta is null or his_estructura.htethasta>=" & ConvFecha(PeriodoHasta) & "))"
    StrSql = StrSql & " AND estructura.estrcodext = '1'" 'Personal
    StrSql = StrSql & " AND his_estructura.ternro = " & rs("ternro")
    If rs3.State = adStateOpen Then rs3.Close
    
    Flog.writeline "Consulta es Grupo de Liquidacion: " & StrSql
    
    OpenRecordset StrSql, rs3
    
    If Not rs3.EOF Then 'la actividad es 5
        'verificar si no hay ninguna novedad en el parámetro porcentaje (35) del concepto 0449
        'si hubiera entonces habría que considerar esa novedad cargada como el porcentaje
        'sino el porcentaje es el default 10%
        'Busco las columnas configuradas en el Confrep
        StrSql = "SELECT confrep.confnrocol, confrep.confetiq, reporte.repnro, conftipo, confval, confval2"
        StrSql = StrSql & " FROM reporte"
        StrSql = StrSql & " INNER JOIN confrep ON reporte.repnro = confrep.repnro"
        StrSql = StrSql & " WHERE (reporte.repnro = 85)"
        StrSql = StrSql & " AND confnrocol = 7 AND conftipo = 'PAR'"
        StrSql = StrSql & " order by confnrocol"
        
        If rs1.State = adStateOpen Then rs1.Close
        
        OpenRecordset StrSql, rs1
        
        'Aplico el porcentaje default
        l_Porc_tkt = 10
        hayNovedad = False
        cantidadCeros = 0
        
        Do Until rs1.EOF
            
            If Not hayNovedad Then
                
                'el valor numerico tiene el nro de parametro y
                'el valor alfanumerico tiene el conccod del concepto
                Columnas(rs1("confnrocol")).TipoOrigen = "PAR"
                Columnas(rs1("confnrocol")).Origen = rs1("confval")
                
                StrSql = "SELECT concepto.concnro, nevalor, nedesde, nehasta "
                StrSql = StrSql & " FROM concepto"
                StrSql = StrSql & " INNER JOIN novemp ON concepto.concnro = novemp.concnro AND novemp.empleado = " & rs("ternro")
                StrSql = StrSql & " WHERE concepto.conccod = '" & rs1("confval2") & "'"
                StrSql = StrSql & " AND novemp.tpanro = " & rs1("confval")
                
                Flog.writeline "Consulta Columna 7: " & StrSql
                
                If rs2.State = adStateOpen Then rs2.Close
                OpenRecordset StrSql, rs2
                Do Until rs2.EOF
                    If vigenciaValida(rs2, PeriodoDesde, PeriodoHasta) Then
                        If CDbl(rs2("nevalor")) <> 0 Then
                           l_Porc_tkt = CDbl(rs2("nevalor"))
                           Flog.writeline Espacios(Tabulador * 1) & "Porcentaje de Tickets. Concepto " & rs1("confval2") & " Parametro " & rs1("confval") & ". Se usa porcentaje " & l_Porc_tkt
                           hayNovedad = True
                        Else
                           cantidadCeros = cantidadCeros + 1
                        End If
                    End If
                    
                    rs2.MoveNext
                Loop
                If rs2.State = adStateOpen Then rs2.Close
                'Columnas(rs1("confnrocol")).Valor = l_Porc_tkt
            End If
        
            rs1.MoveNext
        Loop
        
        v_redo = Round((l_Bases * l_Porc_tkt / 100), 2) - Round((l_Bases * l_Porc_tkt / 100), 0)
        
        If cantidadCeros = 2 Then
            Columnas(7).Valor = Columnas(7).Valor + 0
        Else
            If v_redo > 0 Then
                   Columnas(7).Valor = Columnas(7).Valor + Round((l_Bases * l_Porc_tkt / 100), 0) + 1
            Else
                   Columnas(7).Valor = Columnas(7).Valor + Round((l_Bases * l_Porc_tkt / 100), 0)
            End If
        End If
    Else
        Columnas(7).Valor = Columnas(7).Valor + 0
    End If
    
    ' Columna 8 - Decreto 2005
    ' Columna 9 -
    ' Columna 10 -
    
    Subtotal = Columnas(1).Valor + Columnas(2).Valor + Columnas(3).Valor + Columnas(4).Valor
    'El total es la suma de todas las columnas excepto la 1 y la 5
    Total = Columnas(2).Valor + Columnas(3).Valor + Columnas(4).Valor + Columnas(6).Valor + Columnas(7).Valor + Columnas(8).Valor + Columnas(9).Valor + Columnas(10).Valor
    
    Columnas(Indice_Subtotal).Valor = Subtotal
    Columnas(Indice_Total).Valor = Total
    rs1.Close
    
    '------------------------------------------------------------------------------
    'inserto en la tabla
    '------------------------------------------------------------------------------
    StrSql = "SELECT * FROM rep_listado_sldo "
    StrSql = StrSql & " WHERE bpronro = " & NroProceso
    StrSql = StrSql & " AND ternro = " & rs!ternro
    StrSql = StrSql & " AND sucursal = " & Sucursal
    StrSql = StrSql & " AND sector = " & Sector
    StrSql = StrSql & " AND ccosto = " & CCosto
    StrSql = StrSql & " AND puesto = " & Puesto
    StrSql = StrSql & " AND pagrup = " & PAgrup
    If rs_Listado.State = adStateOpen Then rs_Listado.Close
    OpenRecordset StrSql, rs_Listado
    If rs_Listado.EOF Then
        Flog.writeline Espacios(Tabulador * 1) & "inserto en la tabla rep_listado_sldo "
        StrSql = "INSERT INTO rep_listado_sldo "
        StrSql = StrSql & "("
        StrSql = StrSql & "bpronro,Fecha,Hora,iduser"
        StrSql = StrSql & ",periodo,empleado,ternro,apeynom"
        StrSql = StrSql & ",empresa,empnom,empdirec,emplogo,emplogoalto,emplogoancho"
        StrSql = StrSql & ",sucursal,sucursaldesc, sucursalcodext"
        StrSql = StrSql & ",sector,sectordesc, sectorcodext"
        StrSql = StrSql & ",ccosto,ccostodesc, ccostocodext"
        StrSql = StrSql & ",puesto,puestodesc, puestocodext"
        StrSql = StrSql & ",pagrup,pagrupdesc, pagrupcodext"
        For I = 1 To 20
            StrSql = StrSql & ",columna" & I
        Next I
        StrSql = StrSql & ") VALUES ("
        
        StrSql = StrSql & NroProceso
        StrSql = StrSql & "," & ConvFecha(Fecha)
        StrSql = StrSql & ",'" & Hora & "'"
        StrSql = StrSql & ",'" & IdUser & "'"
        
        StrSql = StrSql & ",'" & Pliqdesc & "'"
        StrSql = StrSql & "," & Legajo
        StrSql = StrSql & "," & rs!ternro
        StrSql = StrSql & ",'" & ApeyNom & "'"
        
        StrSql = StrSql & "," & Empresa
        StrSql = StrSql & ",'" & EmpNom & "'"
        StrSql = StrSql & ",'" & EmpDirec & "'"
        StrSql = StrSql & ",'" & EmpLogo & "'"
        StrSql = StrSql & "," & EmpLogoalto
        StrSql = StrSql & "," & EmpLogoancho
        
        StrSql = StrSql & "," & Sucursal
        StrSql = StrSql & ",'" & Txt_Sucursal & "'"
        StrSql = StrSql & ",'" & CodExt_Sucursal & "'"
        
        StrSql = StrSql & "," & Sector
        StrSql = StrSql & ",'" & Txt_Sector & "'"
        StrSql = StrSql & ",'" & CodExt_Sector & "'"
        
        StrSql = StrSql & "," & CCosto
        StrSql = StrSql & ",'" & Txt_CCosto & "'"
        StrSql = StrSql & ",'" & CodExt_CCosto & "'"
        
        StrSql = StrSql & "," & Puesto
        StrSql = StrSql & ",'" & Txt_Puesto & "'"
        StrSql = StrSql & ",'" & CodExt_Puesto & "'"
        
        StrSql = StrSql & "," & PAgrup
        StrSql = StrSql & ",'" & Txt_PAgrup & "'"
        StrSql = StrSql & ",'" & CodExt_PAgrup & "'"
        
'        StrSql = StrSql & "," & Subtotal
'        StrSql = StrSql & "," & Total
        For I = 1 To 20
            StrSql = StrSql & "," & Replace(CStr(Columnas(I).Valor), ",", ".")
        Next I
        StrSql = StrSql & ")"
    Else
        'Actualizo
        Flog.writeline Espacios(Tabulador * 1) & "Actualizo en la tabla rep_listado_sldo "
        StrSql = "UPDATE rep_listado_sldo SET "
'        StrSql = StrSql & "columna1=" & Subtotal
'        StrSql = StrSql & ",columna2=" & Total
        For I = 1 To 20
            If I = 1 Then
               StrSql = StrSql & " columna" & I & " = " & " columna" & I & " + " & Replace(CStr(Columnas(I).Valor), ",", ".")
            Else
               StrSql = StrSql & ",columna" & I & " = " & " columna" & I & " + " & Replace(CStr(Columnas(I).Valor), ",", ".")
            End If
        Next I
        StrSql = StrSql & " WHERE bpronro = " & NroProceso
        StrSql = StrSql & " AND ternro = " & rs!ternro
        StrSql = StrSql & " AND sucursal = " & Sucursal
        StrSql = StrSql & " AND sector = " & Sector
        StrSql = StrSql & " AND ccosto = " & CCosto
        StrSql = StrSql & " AND puesto = " & Puesto
        StrSql = StrSql & " AND pagrup = " & PAgrup
    End If
    
    Flog.writeline "Consulta Actualizacion Datos: " & StrSql
    
    objConn.Execute StrSql, , adExecuteNoRecords


    'Actualizo el progreso
    TiempoAcumulado = GetTickCount
    Progreso = Progreso + IncPorc
    Flog.writeline Espacios(Tabulador * 1) & "Actualizo el progreso a " & FormatNumber(Progreso, 2) & " %"
    CantidadProcesada = CantidadProcesada - 1
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Replace(CStr(Progreso), ",", ".")
    StrSql = StrSql & ", bprctiempo ='" & Replace(CStr((TiempoAcumulado - TiempoInicialProceso)), ",", ".") & "'"
    StrSql = StrSql & ", bprcempleados ='" & CStr(CantidadProcesada) & "' WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
           
    rs.MoveNext
Loop

'Cierro y libero
If rs_Listado.State = adStateOpen Then rs_Listado.Close
If rs.State = adStateOpen Then rs.Close
If rs1.State = adStateOpen Then rs1.Close
If rs2.State = adStateOpen Then rs2.Close
If rs3.State = adStateOpen Then rs3.Close
If rsE.State = adStateOpen Then rsE.Close

Set rs_Listado = Nothing
Set rs = Nothing
Set rs1 = Nothing
Set rs2 = Nothing
Set rs3 = Nothing
Set rsE = Nothing

Exit Sub
            
MError:
    Flog.writeline " Error: " & Err.Description
    HuboErrores = True
    EmpErrores = True
    
    'Cierro y libero
    If rs_Listado.State = adStateOpen Then rs_Listado.Close
    If rs.State = adStateOpen Then rs.Close
    If rs1.State = adStateOpen Then rs1.Close
    If rs2.State = adStateOpen Then rs2.Close
    If rs3.State = adStateOpen Then rs3.Close
    If rsE.State = adStateOpen Then rsE.Close
    
    Set rs_Listado = Nothing
    Set rs = Nothing
    Set rs1 = Nothing
    Set rs2 = Nothing
    Set rs3 = Nothing
    Set rsE = Nothing
End Sub
'----------------------------------------------------------------
'Busca cual es el mapeo de un codigo RHPro a un codigo SAP
'----------------------------------------------------------------
Public Function CalcularMapeo(ByVal Parametro, ByVal Tabla, ByVal Default)

    Dim StrSql As String
    Dim rs_Consult As New ADODB.Recordset
    Dim correcto As Boolean
    Dim Salida
    
    If IsNull(Parametro) Then
       correcto = False
    Else
       correcto = Parametro <> ""
    End If
           
    Salida = Default

    If correcto Then
        StrSql = " SELECT * FROM infotipos_mapeo " & _
                 " WHERE tablaref = '" & Tabla & "' " & _
                 "   AND codinterno = '" & Parametro & "' "
        OpenRecordset StrSql, rs_Consult
        
        If Not rs_Consult.EOF Then
            Salida = CStr(IIf(Not IsNull(rs_Consult!codexterno), rs_Consult!codexterno, Default))
        Else
            Flog.writeline Espacios(Tabulador * 2) & "No se encontró el mapeo para la tabla " & Tabla & " con el codigo interno " & Parametro
        End If
        
        rs_Consult.Close
    Else
        Flog.writeline Espacios(Tabulador * 2) & "Parametro incorrecto al calcular el mapero de la tabla " & Tabla
    End If
    
    CalcularMapeo = Salida

End Function


Function vigenciaValida(ByRef rs, ByVal Desde As Date, ByVal Hasta As Date)
  Dim Salida As Boolean
  
  Salida = False
  
  If Not rs.EOF Then
     If IsNull(rs!nedesde) Then
        Salida = True
     Else
        If rs!nedesde <= Desde Then
           If IsNull(rs!nehasta) Then
              Salida = True
           Else
              If rs!nehasta >= Hasta Then
                 Salida = True
              End If
           End If
        End If
     End If
  
  End If
  
  vigenciaValida = Salida

End Function



