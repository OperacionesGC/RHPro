Attribute VB_Name = "MdlRHProPlanillaDescuentos"
Option Explicit

Const Version = "1.00"
Const FechaVersion = "27/07/2015"
''    Sebastian Stremel - Version Inicial - CAS-28352 - Salto Grande - Planillas de Descuento

Global TiempoInicialProceso As Long
Global objFeriado As New Feriado
Public Type conc_pais
     concnro As Integer
     conccod As String
     pais As Integer
     turno As Integer
End Type

Dim Progreso As Double
Dim IncPorc As Double

Global TiempoAcumulado As Long


Sub Main()

Dim strcmdLine  As String

Dim rs As New ADODB.Recordset

Dim Fecha As Date
Dim Hora As String
Dim NroProceso As Long
Dim nroReporte As Long
Dim StrParametros As String

Dim rs_Batch_Proceso As New ADODB.Recordset
Dim rs_His_Batch_Proceso As New ADODB.Recordset
Dim PID As String
Dim ArrParametros

Dim Archivo As String
Dim NroProcesoBatch As Long
Dim bprcparam As String
Dim HuboErrores

Progreso = 0

    strcmdLine = Command()
    ArrParametros = Split(strcmdLine, " ", -1)
    If UBound(ArrParametros) > 1 Then
        If IsNumeric(ArrParametros(0)) Then
            NroProceso = ArrParametros(0)
            Etiqueta = ArrParametros(1)
            EncriptStrconexion = CBool(ArrParametros(2))
            c_seed = ArrParametros(2)
        Else
            Exit Sub
        End If
    Else
        If UBound(ArrParametros) > 0 Then
            If IsNumeric(ArrParametros(0)) Then
                NroProceso = ArrParametros(0)
                Etiqueta = ArrParametros(1)
            Else
                Exit Sub
            End If
        Else
            If IsNumeric(strcmdLine) Then
                NroProceso = strcmdLine
            Else
                Exit Sub
            End If
        End If
    End If
    
    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas

    'Creo el archivo de texto del desglose
    Archivo = PathFLog & "RHProPlanillaDescuento-" & CStr(NroProceso) & Format(Now, "DD-MM-YYYY") & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Archivo, True)

    'Abro la conexion
    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    On Error Resume Next
    OpenConnection strconexion, objConnProgreso
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If

    On Error GoTo CE

    'Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaVersion
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-------------------------------------------------"
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "Inicio :" & Now
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprcpid = " & PID & ", bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords

    Flog.writeline Espacios(Tabulador * 0) & "Levanta Proceso y Setea Parámetros:  " & " " & Now
    
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 454 AND bpronro =" & NroProceso
    OpenRecordset StrSql, rs_Batch_Proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_Batch_Proceso.EOF Then
        bprcparam = rs_Batch_Proceso!bprcparam
        rs_Batch_Proceso.Close
        Set rs_Batch_Proceso = Nothing
        Call acumParcial(NroProceso, bprcparam)
    Else
        Flog.writeline "no encontró el proceso"
    End If

    
    
    'Actualizo el Btach_Proceso
    If Not HuboErrores Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProceso
    End If
    objConn.Execute StrSql, , adExecuteNoRecords
    
    ' -----------------------------------------------------------------------------------
    'Copio el proceso en el historico de batch_proceso y lo borro de batch_proceso
    If Not HuboErrores Then
        'Copio el proceso en el historico de batch_proceso y lo borro de batch_proceso
        StrSql = "SELECT * FROM batch_proceso WHERE bpronro =" & NroProceso
        OpenRecordset StrSql, rs_Batch_Proceso
        
        StrSql = "INSERT INTO His_Batch_Proceso (bpronro,btprcnro,bprcfecha,iduser"
        StrSqlDatos = rs_Batch_Proceso!bpronro & "," & rs_Batch_Proceso!btprcnro & "," & _
        ConvFecha(rs_Batch_Proceso!bprcfecha) & ",'" & rs_Batch_Proceso!iduser & "'"
        
        If Not IsNull(rs_Batch_Proceso!bprchora) Then
            StrSql = StrSql & ",bprchora"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprchora & "'"
        End If
        If Not IsNull(rs_Batch_Proceso!bprcempleados) Then
            StrSql = StrSql & ",bprcempleados"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprcempleados & "'"
        End If
        If Not IsNull(rs_Batch_Proceso!bprcfecdesde) Then
            StrSql = StrSql & ",bprcfecdesde"
            StrSqlDatos = StrSqlDatos & "," & ConvFecha(rs_Batch_Proceso!bprcfecdesde)
        End If
        If Not IsNull(rs_Batch_Proceso!bprcfechasta) Then
            StrSql = StrSql & ",bprcfechasta"
            StrSqlDatos = StrSqlDatos & "," & ConvFecha(rs_Batch_Proceso!bprcfechasta)
        End If
        If Not IsNull(rs_Batch_Proceso!bprcestado) Then
            StrSql = StrSql & ",bprcestado"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprcestado & "'"
        End If
        If Not IsNull(rs_Batch_Proceso!bprcparam) Then
            StrSql = StrSql & ",bprcparam"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprcparam & "'"
        End If
        If Not IsNull(rs_Batch_Proceso!bprcprogreso) Then
            StrSql = StrSql & ",bprcprogreso"
            StrSqlDatos = StrSqlDatos & "," & rs_Batch_Proceso!bprcprogreso
        End If
        If Not IsNull(rs_Batch_Proceso!bprcfecfin) Then
            StrSql = StrSql & ",bprcfecfin"
            StrSqlDatos = StrSqlDatos & "," & ConvFecha(rs_Batch_Proceso!bprcfecfin)
        End If
        If Not IsNull(rs_Batch_Proceso!bprchorafin) Then
            StrSql = StrSql & ",bprchorafin"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprchorafin & "'"
        End If
        If Not IsNull(rs_Batch_Proceso!bprctiempo) Then
            StrSql = StrSql & ",bprctiempo"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprctiempo & "'"
        End If
        If Not IsNull(rs_Batch_Proceso!empnro) Then
            StrSql = StrSql & ",empnro"
            StrSqlDatos = StrSqlDatos & "," & rs_Batch_Proceso!empnro
        End If
        If Not IsNull(rs_Batch_Proceso!bprcPid) Then
            StrSql = StrSql & ",bprcPid"
            StrSqlDatos = StrSqlDatos & "," & rs_Batch_Proceso!bprcPid
        End If
        If Not IsNull(rs_Batch_Proceso!bprcfecInicioEj) Then
            StrSql = StrSql & ",bprcfecInicioEj"
            StrSqlDatos = StrSqlDatos & "," & ConvFecha(rs_Batch_Proceso!bprcfecInicioEj)
        End If
        If Not IsNull(rs_Batch_Proceso!bprcfecFinEj) Then
            StrSql = StrSql & ",bprcfecFinEj"
            StrSqlDatos = StrSqlDatos & "," & ConvFecha(rs_Batch_Proceso!bprcfecFinEj)
        End If
        If Not IsNull(rs_Batch_Proceso!bprcUrgente) Then
            StrSql = StrSql & ",bprcUrgente"
            StrSqlDatos = StrSqlDatos & "," & rs_Batch_Proceso!bprcUrgente
        End If
        If Not IsNull(rs_Batch_Proceso!bprcHoraInicioEj) Then
            StrSql = StrSql & ",bprcHoraInicioEj"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprcHoraInicioEj & "'"
        End If
        If Not IsNull(rs_Batch_Proceso!bprcHoraFinEj) Then
            StrSql = StrSql & ",bprcHoraFinEj"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprcHoraFinEj & "'"
        End If

        StrSql = StrSql & ") VALUES (" & StrSqlDatos & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        
        'Reviso que haya copiado
        StrSql = "SELECT * FROM His_batch_proceso WHERE bpronro =" & NroProceso
        OpenRecordset StrSql, rs_His_Batch_Proceso
        
        If Not rs_His_Batch_Proceso.EOF Then
            ' Borro de Batch_proceso
            StrSql = "DELETE FROM Batch_Proceso WHERE bpronro = " & NroProceso
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
    
        If rs_Batch_Proceso.State = adStateOpen Then rs_Batch_Proceso.Close
        If rs_His_Batch_Proceso.State = adStateOpen Then rs_His_Batch_Proceso.Close
    End If
    Flog.Close
    
    'Cierro y libero todo
    If TransactionRunning Then MyRollbackTrans
    
    If objConn.State = adStateOpen Then objConn.Close
    If objConnProgreso.State = adStateOpen Then objConnProgreso.Close
    If CnTraza.State = adStateOpen Then CnTraza.Close
Exit Sub

CE:
    Flog.writeline "Reporte abortado por Error:" & " " & Now
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
        StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Error General', bprcpid = " & PID & " WHERE bpronro = " & NroProceso
        objConnProgreso.Execute StrSql, , adExecuteNoRecords
    MyCommitTrans
    
    
End Sub

Sub acumParcial(NroProcesoBatch, bprcparam)

'aca se deben generar cada uno de los 3 reportes Presentismo - Ausentismo - Extras
Dim legdesde As Long
Dim leghasta As Long
Dim tenro1 As Integer
Dim estrnro1 As Integer
Dim tenro2 As Integer
Dim estrnro2 As Integer
Dim tenro3 As Integer
Dim estrnro3 As Integer

Dim tipoFiltro As Integer  ' por proceso(0) o por fecha(1)
Dim fecdesde As String
Dim fechasta As String
Dim listaProc As String
Dim listaEmpleados As String
listaEmpleados = "0"
Dim parametros
Dim StrSql As String
Dim nroReporte As Long

Dim rs_consultas As New ADODB.Recordset


'Tipos de hora
Dim concPresentismo As String
concPresentismo = 0

Dim cantRep As Integer
Dim cantEmp As Integer
Dim emp



parametros = Split(bprcparam, "@")

'68@71@10@1504@0@-1@0@@0@0,29,30

legdesde = parametros(0)
leghasta = parametros(1)
If EsNulo(parametros(2)) Then
    tenro1 = 0
Else
    tenro1 = parametros(2)
End If

If EsNulo(parametros(3)) Then
    estrnro1 = 0
Else
    estrnro1 = parametros(3)
End If

If EsNulo(parametros(4)) Then
    tenro2 = 0
Else
    tenro2 = parametros(4)
End If

If EsNulo(parametros(5)) Then
    estrnro2 = 0
Else
    estrnro2 = parametros(5)
End If

If EsNulo(parametros(6)) Then
    tenro3 = 0
Else
    tenro3 = parametros(6)
End If

If EsNulo(parametros(7)) Then
    estrnro3 = 0
Else
    estrnro3 = parametros(7)
End If

tipoFiltro = parametros(8)

If tipoFiltro = 0 Then
    listaProc = parametros(9)
    'busco la minima fecha del proceso y la maxima fecha
    StrSql = " SELECT MIN(gpadesde) desde, MAX(gpahasta) hasta "
    StrSql = StrSql & " FROM gti_procacum  "
    StrSql = StrSql & " WHERE gpanro  IN(" & listaProc & ")"
    OpenRecordset StrSql, rs_consultas
    If Not rs_consultas.EOF Then
        fecdesde = rs_consultas!desde
        fechasta = rs_consultas!hasta
        Flog.writeline "Fecha Inicial:" & fecdesde & " Fecha Final: " & fechasta & " POR PROCESO "
    Else
        Flog.writeline "Ocurrio un error con las fechas del filtro POR PROCESO"
    End If
    rs_consultas.Close
Else
    fecdesde = parametros(9)
    fechasta = parametros(10)
End If

'Filtro los empleados
StrSql = " SELECT empleado.ternro  "
StrSql = StrSql & " From Empleado "
If tenro1 <> "0" Then
    StrSql = StrSql & " INNER JOIN his_estructura h1 ON empleado.ternro = h1.ternro AND h1.tenro = " & tenro1
    If estrnro1 <> "-1" Then
        StrSql = StrSql & " AND estrnro=" & estrnro1
    End If
    StrSql = StrSql & " AND ("
    StrSql = StrSql & " (h1.htetdesde <= " & ConvFecha(fecdesde) & " AND (h1.htethasta is null or h1.htethasta >= " & ConvFecha(fechasta) & ""
    StrSql = StrSql & " or h1.htethasta >= " & ConvFecha(fecdesde) & ")) OR"
    StrSql = StrSql & " (h1.htetdesde >= " & ConvFecha(fecdesde) & " AND (h1.htetdesde <= " & ConvFecha(fechasta) & "))"
    StrSql = StrSql & " )"
End If
If tenro2 <> "0" Then
    StrSql = StrSql & " INNER JOIN his_estructura h2 ON empleado.ternro = h2.ternro AND h2.tenro = " & tenro2
    If estrnro2 <> "-1" Then
        StrSql = StrSql & " AND estrnro=" & estrnro2
    End If
    StrSql = StrSql & " AND ("
    StrSql = StrSql & " (h2.htetdesde <= " & ConvFecha(fecdesde) & " AND (h2.htethasta is null or h2.htethasta >= " & ConvFecha(fechasta) & ""
    StrSql = StrSql & " or h2.htethasta >= " & ConvFecha(fecdesde) & ")) OR"
    StrSql = StrSql & " (h2.htetdesde >= " & ConvFecha(fecdesde) & " AND (h2.htetdesde <= " & ConvFecha(fechasta) & "))"
    StrSql = StrSql & " )"
End If
If tenro3 <> "0" Then
    StrSql = StrSql & " INNER JOIN his_estructura h3 ON empleado.ternro = h3.ternro AND h3.tenro = " & tenro3
    If estrnro3 <> "-1" Then
        StrSql = StrSql & " AND estrnro=" & estrnro3
    End If
    StrSql = StrSql & " AND ("
    StrSql = StrSql & " (h3.htetdesde <= " & ConvFecha(fecdesde) & " AND (h3.htethasta is null or h3.htethasta >= " & ConvFecha(fechasta) & ""
    StrSql = StrSql & " or h3.htethasta >= " & ConvFecha(fecdesde) & ")) OR"
    StrSql = StrSql & " (h3.htetdesde >= " & ConvFecha(fecdesde) & " AND (h3.htetdesde <= " & ConvFecha(fechasta) & "))"
    StrSql = StrSql & " )"
End If
StrSql = StrSql & " WHERE empleado.empleg >=" & legdesde & " AND empleado.empleg <=" & leghasta
Flog.writeline "Consulta Empleados: " & StrSql
OpenRecordset StrSql, rs_consultas
If Not rs_consultas.EOF Then
    Do While Not rs_consultas.EOF
        listaEmpleados = listaEmpleados & "," & rs_consultas!ternro
    rs_consultas.MoveNext
    Loop
Else
    Flog.writeline "No hay empleados para procesar."
    Exit Sub
End If
rs_consultas.Close
'Hasta aca

'----------INSERTO LA CABECERA DEL REPORTE----------
StrSql = " INSERT INTO rep_planilla_dtos"
StrSql = StrSql & "(repbpronro, replegdesde, repleghasta, reptenro1, "
StrSql = StrSql & " repestrnro1, reptenro2, repestrnro2, reptenro3, repestrnro3, "
StrSql = StrSql & " repfiltro, replistaproc, repfdesde, repfhasta)"
StrSql = StrSql & " VALUES "
StrSql = StrSql & " ( "
StrSql = StrSql & NroProcesoBatch & "," & legdesde & "," & leghasta & ","
StrSql = StrSql & tenro1 & "," & estrnro1 & "," & tenro2 & "," & estrnro2 & ","
StrSql = StrSql & tenro3 & "," & estrnro3 & "," & tipoFiltro & ","
If tipoFiltro = 0 Then
    StrSql = StrSql & "'" & listaProc & "',"
Else
    StrSql = StrSql & "NULL,"
End If
StrSql = StrSql & ConvFecha(fecdesde) & "," & ConvFecha(fechasta)
StrSql = StrSql & " ) "
objConn.Execute StrSql, , adExecuteNoRecords
'-------------------HASTA ACA-----------------------

nroReporte = getLastIdentity(objConn, "rep_planilla_dtos")

'BUSCO LA CANTIDAD DE REPORTES CONF
StrSql = " SELECT count(repnro) cant FROM confrepadv "
StrSql = StrSql & " WHERE UPPER(conftipo) ='TR'"
StrSql = StrSql & " AND repnro=489 "
OpenRecordset StrSql, rs_consultas
If Not rs_consultas.EOF Then
    If Not EsNulo(rs_consultas!Cant) Then
        cantRep = rs_consultas!Cant
        Flog.writeline "Cantidad de reportes: " & cantRep
    Else
        cantRep = 0
        Flog.writeline "Cantidad de reportes: " & cantRep
        Exit Sub
    End If
Else
    cantRep = 0
    Flog.writeline "Cantidad de reportes: " & cantRep
    Exit Sub
End If
rs_consultas.Close
'HASTA ACA

'BUSCO LA CANTIDAD DE EMPLEADOS--
emp = Split(listaEmpleados, ",")
cantEmp = UBound(emp)

If cantEmp = 0 Then
    Flog.writeline "No hay empleados para procesar."
    Exit Sub
End If
'--------------------------------
'CALCULO EL PORC DE AVANCE
IncPorc = 100 / cantRep
IncPorc = IncPorc / cantEmp
'HASTA ACA

StrSql = " SELECT * FROM confrepadv "
StrSql = StrSql & " WHERE UPPER(conftipo) ='TR' "
StrSql = StrSql & " AND repnro=489 "
StrSql = StrSql & " ORDER BY confnrocol ASC "
OpenRecordset StrSql, rs_consultas
If Not rs_consultas.EOF Then
    Do While Not rs_consultas.EOF
        Select Case rs_consultas!confval
            Case 1:
                Flog.writeline "-----------SE EMPIEZA A PROCESAR EL REPORTE " & rs_consultas!confetiq & "-----------"
                Call repPresentismo(NroProcesoBatch, nroReporte, listaEmpleados, tipoFiltro, fecdesde, fechasta)
            Case 2:
                Flog.writeline "-----------SE EMPIEZA A PROCESAR EL REPORTE " & rs_consultas!confetiq & "-----------"
                Call repAusentismo(NroProcesoBatch, nroReporte, listaEmpleados, tipoFiltro, fecdesde, fechasta)
            Case 3:
                Flog.writeline "-----------SE EMPIEZA A PROCESAR EL REPORTE " & rs_consultas!confetiq & "-----------"
                Call repExtras(NroProcesoBatch, nroReporte, listaEmpleados, tipoFiltro, fecdesde, fechasta)
        End Select
    rs_consultas.MoveNext
    Loop
Else
    Flog.writeline "No hay reportes configurados para generar."
End If
rs_consultas.Close


End Sub

Sub repPresentismo(ByVal NroProcesoBatch As Long, ByVal nroReporte As Integer, ByVal listaEmpleados As String, ByVal tipoFiltro As Integer, ByVal fecdesde As String, ByVal fechasta As String)

'Aca vamos a generar el reporte de presentismo
Dim Empleado
Dim j As Integer
Dim k As Integer

Dim aConcepto() As conc_pais


Dim th As String
Dim cantdias As Integer
Dim rs As New ADODB.Recordset
Dim fecdesdeaux As Date
Dim ternro As Long
Dim objempleado As New datosPersonales
Dim Nombre
Dim Apellido
Dim Legajo
Dim nrodoc
Dim empDocPais As Integer


'Variables del confrep
Dim teRegimenHorario As Integer
Dim estrTurnoFijo As Integer
Dim estrTurnoRot As Integer
Dim esFijo As Boolean
Dim esRot As Boolean
Dim concPresentismo As String

teRegimenHorario = 0
estrTurnoFijo = 0
estrTurnoRot = 0
concPresentismo = 0
j = 0
th = "0"
'Levanto del confrep los tipos de Horas
StrSql = " SELECT * FROM confrepadv "
StrSql = StrSql & " WHERE repnro=489"
'StrSql = StrSql & " AND UPPER(conftipo)='CO1'"
StrSql = StrSql & " ORDER BY confnrocol ASC "
OpenRecordset StrSql, rs
If Not rs.EOF Then
    Do While Not rs.EOF
        Select Case UCase(rs!conftipo)
            Case "CO1":
                j = j + 1
                ReDim Preserve aConcepto(j)
                aConcepto(j).concnro = IIf(EsNulo(rs!confval), 0, rs!confval)
                aConcepto(j).pais = IIf(EsNulo(rs!confval2), 0, rs!confval2)
            Case "TE":
                If (rs!confnrocol = 10) Then
                    teRegimenHorario = IIf(EsNulo(rs!confval), 0, rs!confval)
                    estrTurnoFijo = IIf(EsNulo(rs!confval2), 0, rs!confval2)
                    estrTurnoRot = IIf(EsNulo(rs!confval3), 0, rs!confval3)
                End If
        End Select
    rs.MoveNext
    Loop
End If
rs.Close
'hasta aca

j = 0
'Busco los tipo de hora de los conceptos
'th = buscarTipoHora(concPresentismo)

Empleado = Split(listaEmpleados, ",")

'calculo la diferecia de dias
cantdias = DateDiff("d", fecdesde, fechasta)
'hasta aca

Dim listaFeriado
Dim esFeriado
Dim descuento
For j = 1 To UBound(Empleado) 'arranco del 1 porque en el indice 0 tengo como legajo el 0
    ternro = Empleado(j)
    fecdesdeaux = fecdesde
    esFijo = False
    esRot = False
    descuento = 0
    
    'busco los datos del empleado
    objempleado.buscarDatosPersonales (ternro)
    Nombre = objempleado.obtenerNombreApellido("nombre") & " " & objempleado.obtenerNombreApellido("nombre2")
    Apellido = objempleado.obtenerNombreApellido("apellido") & " " & objempleado.obtenerNombreApellido("apellido2")
    Legajo = objempleado.obtenerLegajo
    objempleado.buscarNroDoc ternro, 0
    nrodoc = objempleado.obtenerNroDoc
    empDocPais = objempleado.obtenerDocPais
    'hasta aca
    
    th = buscarTipoHora(empDocPais, aConcepto, concPresentismo)
    
    'Busco el turno del empleado
    StrSql = "SELECT estrnro FROM his_estructura h"
    StrSql = StrSql & " WHERE ("
    StrSql = StrSql & " (h.htetdesde <= " & ConvFecha(fecdesde) & " AND (h.htethasta is null or h.htethasta >= " & ConvFecha(fechasta) & ""
    StrSql = StrSql & " or h.htethasta >= " & ConvFecha(fecdesde) & ")) OR"
    StrSql = StrSql & " (h.htetdesde >= " & ConvFecha(fecdesde) & " AND (h.htetdesde <= " & ConvFecha(fechasta) & "))"
    StrSql = StrSql & " )"
    StrSql = StrSql & " AND h.ternro=" & ternro & " AND h.tenro=" & teRegimenHorario
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        If Not EsNulo(rs!estrnro) Then
            Select Case CLng(rs!estrnro)
                Case estrTurnoFijo:
                    esFijo = True
                Case estrTurnoRot:
                    esRot = True
            End Select
        End If
    End If
    rs.Close
    'hasta aca
    
    
    If esFijo Then
        listaFeriado = ""
        Do While (CDate(fecdesdeaux) <= CDate(fechasta))
            esFeriado = objFeriado.Feriado(fecdesdeaux, ternro, False)
            If esFeriado Then
                listaFeriado = listaFeriado & ",'" & fecdesdeaux & "'"
            End If
        fecdesdeaux = DateAdd("d", 1, fecdesdeaux)
        Loop
        listaFeriado = Replace(listaFeriado, ",", "", 1, 1)
        'busco la suma de las horas
        StrSql = " SELECT sum(adcanthoras) horas,gti_acumdiario.thnro, conccod "
        StrSql = StrSql & " From gti_acumdiario "
        StrSql = StrSql & " INNER JOIN tiph_con ON tiph_con.thnro = gti_acumdiario.thnro "
        StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = tiph_con.concnro "
        StrSql = StrSql & " WHERE adfecha >= " & ConvFecha(fecdesde)
        StrSql = StrSql & " AND adfecha <= " & ConvFecha(fechasta)
        If listaFeriado <> "" Then
            StrSql = StrSql & " AND adfecha NOT IN(" & listaFeriado & ")"
        End If
        StrSql = StrSql & " AND ternro=" & ternro
        StrSql = StrSql & " AND gti_acumdiario.thnro IN (" & th & ")"
        StrSql = StrSql & " AND tiph_con.concnro IN (" & concPresentismo & ")"
        StrSql = StrSql & " GROUP BY gti_acumdiario.thnro , conccod "
        Flog.writeline "Horas Presentismo:" & StrSql
        OpenRecordset StrSql, rs
        If Not rs.EOF Then
            If Not EsNulo(rs!Horas) Then
                'Si Total de Ausencias = 1, Descuento = 100
                'Si Total de Ausencias > 1, Descuento = 200
                If CDbl(rs!Horas) = 1 Then
                    descuento = 100
                Else
                    If CDbl(rs!Horas) > 1 Then
                        descuento = 200
                    Else
                        descuento = 0
                    End If
                End If
                
                StrSql = " INSERT INTO rep_planilla_dtos_det "
                StrSql = StrSql & "("
                StrSql = StrSql & " repnro, repbpronro, reptiporeporte,  "
                StrSql = StrSql & " replegajo, repnombre, repapellido, "
                StrSql = StrSql & " repdoc, repconcepto, reptipohora, repcantidad "
                StrSql = StrSql & ")"
                StrSql = StrSql & " VALUES "
                StrSql = StrSql & "("
                StrSql = StrSql & nroReporte & "," & NroProcesoBatch & ",1,"
                StrSql = StrSql & Legajo & ",'" & Nombre & "','" & Apellido & "',"
                StrSql = StrSql & "'" & nrodoc & "'," & rs!conccod & "," & rs!thnro & "," & descuento
                StrSql = StrSql & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
        Else
            Flog.writeline "El empleado: " & Legajo & " no tiene cantidad de horas."
        End If
        rs.Close
    Else
        If esRot Then
            Flog.writeline "El empleado " & Legajo & " tiene turno rotativo."
            Do While (CDate(fecdesdeaux) <= CDate(fechasta))
                esFeriado = objFeriado.Feriado(fecdesdeaux, ternro, False)
                If (Weekday(fecdesdeaux) = 1) Or (Weekday(fecdesdeaux) = 7) Or esFeriado Then
                    listaFeriado = listaFeriado & ",'" & fecdesdeaux & "'"
                End If
            fecdesdeaux = DateAdd("d", 1, fecdesdeaux)
            Loop
            'busco la suma de las horas
            StrSql = " SELECT sum(adcanthoras) horas,gti_acumdiario.thnro, conccod "
            StrSql = StrSql & " From gti_acumdiario "
            StrSql = StrSql & " INNER JOIN tiph_con ON tiph_con.thnro = gti_acumdiario.thnro "
            StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = tiph_con.concnro "
            StrSql = StrSql & " WHERE adfecha >= " & ConvFecha(fecdesde)
            StrSql = StrSql & " AND adfecha <= " & ConvFecha(fechasta)
            StrSql = StrSql & " AND adfecha IN(" & listaFeriado & ")"
            StrSql = StrSql & " AND ternro=" & ternro
            StrSql = StrSql & " AND gti_acumdiario.thnro IN (" & th & ")"
            StrSql = StrSql & " AND tiph_con.concnro IN (" & concPresentismo & ")"
            StrSql = StrSql & " GROUP BY gti_acumdiario.thnro , conccod "
            OpenRecordset StrSql, rs
            If Not rs.EOF Then
                If Not EsNulo(rs!Horas) Then
                    'Si Total de Ausencias = 1, Descuento = 100
                    'Si Total de Ausencias > 1, Descuento = 200
                    If CDbl(rs!Horas) = 1 Then
                        descuento = 100
                    Else
                        If CDbl(rs!Horas) > 1 Then
                            descuento = 200
                        Else
                            descuento = 0
                        End If
                    End If
                    StrSql = " INSERT INTO rep_planilla_dtos_det "
                    StrSql = StrSql & "("
                    StrSql = StrSql & " repnro, repbpronro, reptiporeporte,  "
                    StrSql = StrSql & " replegajo, repnombre, repapellido, "
                    StrSql = StrSql & " repdoc, repconcepto, reptipohora, repcantidad "
                    StrSql = StrSql & ")"
                    StrSql = StrSql & " VALUES "
                    StrSql = StrSql & "("
                    StrSql = StrSql & nroReporte & "," & NroProcesoBatch & ",1,"
                    StrSql = StrSql & Legajo & ",'" & Nombre & "','" & Apellido & "',"
                    StrSql = StrSql & "'" & nrodoc & "'," & rs!conccod & "," & rs!thnro & "," & descuento
                    StrSql = StrSql & ")"
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
            Else
                Flog.writeline "El empleado: " & Legajo & " no tiene cantidad de horas."
            End If
            rs.Close
        End If
    End If
    TiempoAcumulado = GetTickCount
    Progreso = Progreso + IncPorc
    Flog.writeline "Actualizo el progreso: " & Progreso
    StrSql = "UPDATE batch_proceso SET bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "', bprcprogreso = " & CLng(Progreso) & " WHERE bpronro = " & NroProcesoBatch
    objConnProgreso.Execute StrSql, , adExecuteNoRecords
    Flog.writeline "Se actualizo el progreso"
    
    'Progreso = Progreso + IncPorc
    'Progreso = Replace(Progreso, ",", ".")
    'StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Replace(Progreso, ",", ".") & " WHERE bpronro = " & NroProcesoBatch
    'objConnProgreso.Execute StrSql, , adExecuteNoRecords
Next j

    
End Sub
Sub repAusentismo(ByVal NroProcesoBatch As Long, ByVal nroReporte As Integer, ByVal listaEmpleados As String, ByVal tipoFiltro As Integer, ByVal fecdesde As String, ByVal fechasta As String)

'Aca vamos a generar el reporte de presentismo
Dim Empleado
Dim j As Integer
Dim k As Integer

Dim aConcepto() As conc_pais


Dim th As String
Dim cantdias As Integer
Dim rs As New ADODB.Recordset
Dim fecdesdeaux As Date
Dim ternro As Long
Dim objempleado As New datosPersonales
Dim Nombre
Dim Apellido
Dim Legajo
Dim nrodoc
Dim empDocPais As Integer


'Variables del confrep
Dim teRegimenHorario As Integer
Dim estrTurnoFijo As Integer
Dim estrTurnoRot As Integer
Dim estrSemanaNoCal As Integer
Dim esFijo As Boolean
Dim esRot As Boolean
Dim esSem As Boolean
Dim concPresentismo As String

teRegimenHorario = 0
estrTurnoFijo = 0
estrTurnoRot = 0
estrSemanaNoCal = 0
concPresentismo = 0
th = "0"
j = 0

'Levanto del confrep los tipos de Horas
StrSql = " SELECT * FROM confrepadv "
StrSql = StrSql & " WHERE repnro=489"
'StrSql = StrSql & " AND UPPER(conftipo)='CO1'"
StrSql = StrSql & " ORDER BY confnrocol ASC "
OpenRecordset StrSql, rs
If Not rs.EOF Then
    Do While Not rs.EOF
        Select Case UCase(rs!conftipo)
            Case "CO2":
                j = j + 1
                ReDim Preserve aConcepto(j)
                aConcepto(j).concnro = IIf(EsNulo(rs!confval), 0, rs!confval)
                aConcepto(j).pais = IIf(EsNulo(rs!confval2), 0, rs!confval2)
                aConcepto(j).turno = IIf(EsNulo(rs!confval3), 0, rs!confval3)
            Case "TE":
                If (rs!confnrocol = 10) Then
                    teRegimenHorario = IIf(EsNulo(rs!confval), 0, rs!confval)
                    estrTurnoFijo = IIf(EsNulo(rs!confval2), 0, rs!confval2)
                    estrTurnoRot = IIf(EsNulo(rs!confval3), 0, rs!confval3)
                    estrSemanaNoCal = IIf(EsNulo(rs!confval4), 0, rs!confval4)
                End If
        End Select
    rs.MoveNext
    Loop
End If
rs.Close
'hasta aca

j = 0
'Busco los tipo de hora de los conceptos
'th = buscarTipoHora(concPresentismo)

Empleado = Split(listaEmpleados, ",")

'calculo la diferecia de dias
cantdias = DateDiff("d", fecdesde, fechasta)
'hasta aca

Dim listaFeriado
Dim esFeriado
Dim porc As Double
Dim descuento As Integer
For j = 1 To UBound(Empleado) 'arranco del 1 porque en el indice 0 tengo como legajo el 0
    ternro = Empleado(j)
    fecdesdeaux = fecdesde
    esFijo = False
    esRot = False
    esSem = False
    
    'busco los datos del empleado
    objempleado.buscarDatosPersonales (ternro)
    Nombre = objempleado.obtenerNombreApellido("nombre") & " " & objempleado.obtenerNombreApellido("nombre2")
    Apellido = objempleado.obtenerNombreApellido("apellido") & " " & objempleado.obtenerNombreApellido("apellido2")
    Legajo = objempleado.obtenerLegajo
    objempleado.buscarNroDoc ternro, 0
    nrodoc = objempleado.obtenerNroDoc
    empDocPais = objempleado.obtenerDocPais
    'hasta aca
    
    'th = buscarTipoHora(empDocPais, aConcepto, concPresentismo)
    
    'Busco el turno del empleado
    StrSql = "SELECT estrnro FROM his_estructura h"
    StrSql = StrSql & " WHERE ("
    StrSql = StrSql & " (h.htetdesde <= " & ConvFecha(fecdesde) & " AND (h.htethasta is null or h.htethasta >= " & ConvFecha(fechasta) & ""
    StrSql = StrSql & " or h.htethasta >= " & ConvFecha(fecdesde) & ")) OR"
    StrSql = StrSql & " (h.htetdesde >= " & ConvFecha(fecdesde) & " AND (h.htetdesde <= " & ConvFecha(fechasta) & "))"
    StrSql = StrSql & " )"
    StrSql = StrSql & " AND h.ternro=" & ternro & " AND h.tenro=" & teRegimenHorario
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        If Not EsNulo(rs!estrnro) Then
            Select Case CLng(rs!estrnro)
                Case estrTurnoFijo:
                    esFijo = True
                    th = buscarTipoHora(empDocPais, aConcepto, concPresentismo, estrTurnoFijo)
                Case estrTurnoRot:
                    esRot = True
                    th = buscarTipoHora(empDocPais, aConcepto, concPresentismo, estrTurnoRot)
                Case estrSemanaNoCal:
                    esSem = True
                    th = buscarTipoHora(empDocPais, aConcepto, concPresentismo, estrSemanaNoCal)
            End Select
        End If
    End If
    rs.Close
    'hasta aca
    If th <> "" Then
    
        If esFijo Then
            listaFeriado = ""
            Do While (CDate(fecdesdeaux) <= CDate(fechasta))
                esFeriado = objFeriado.Feriado(fecdesdeaux, ternro, False)
                If esFeriado Then
                    listaFeriado = listaFeriado & ",'" & fecdesdeaux & "'"
                End If
            fecdesdeaux = DateAdd("d", 1, fecdesdeaux)
            Loop
            listaFeriado = Replace(listaFeriado, ",", "", 1, 1)
            'busco la suma de las horas
            StrSql = " SELECT sum(adcanthoras) horas,gti_acumdiario.thnro, conccod "
            StrSql = StrSql & " From gti_acumdiario "
            StrSql = StrSql & " INNER JOIN tiph_con ON tiph_con.thnro = gti_acumdiario.thnro "
            StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = tiph_con.concnro "
            StrSql = StrSql & " WHERE adfecha >= " & ConvFecha(fecdesde)
            StrSql = StrSql & " AND adfecha <= " & ConvFecha(fechasta)
            If listaFeriado <> "" Then
                StrSql = StrSql & " AND adfecha NOT IN(" & listaFeriado & ")"
            End If
            StrSql = StrSql & " AND ternro=" & ternro
            StrSql = StrSql & " AND gti_acumdiario.thnro IN (" & th & ")"
            StrSql = StrSql & " AND tiph_con.concnro IN (" & concPresentismo & ")"
            StrSql = StrSql & " GROUP BY gti_acumdiario.thnro , conccod "
            OpenRecordset StrSql, rs
            If Not rs.EOF Then
                If Not EsNulo(rs!Horas) Then
                    StrSql = " INSERT INTO rep_planilla_dtos_det "
                    StrSql = StrSql & "("
                    StrSql = StrSql & " repnro, repbpronro, reptiporeporte,  "
                    StrSql = StrSql & " replegajo, repnombre, repapellido, "
                    StrSql = StrSql & " repdoc, repconcepto, reptipohora, repcantidad "
                    StrSql = StrSql & ")"
                    StrSql = StrSql & " VALUES "
                    StrSql = StrSql & "("
                    StrSql = StrSql & nroReporte & "," & NroProcesoBatch & ",1,"
                    StrSql = StrSql & Legajo & ",'" & Nombre & "','" & Apellido & "',"
                    StrSql = StrSql & "'" & nrodoc & "'," & rs!conccod & "," & rs!thnro & "," & Replace(rs!Horas, ",", ".")
                    StrSql = StrSql & ")"
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
            Else
                Flog.writeline "El empleado: " & Legajo & " no tiene cantidad de horas."
            End If
            rs.Close
        Else
            If esRot Or esSem Then
                Flog.writeline "El empleado " & Legajo & " tiene turno rotativo."
                Do While (CDate(fecdesdeaux) <= CDate(fechasta))
                    esFeriado = objFeriado.Feriado(fecdesdeaux, ternro, False)
                    If (Weekday(fecdesdeaux) = 1) Or (Weekday(fecdesdeaux) = 7) Or esFeriado Then
                        listaFeriado = listaFeriado & ",'" & fecdesdeaux & "'"
                    End If
                fecdesdeaux = DateAdd("d", 1, fecdesdeaux)
                Loop
                'busco la suma de las horas
                StrSql = " SELECT sum(adcanthoras) horas,gti_acumdiario.thnro, conccod "
                StrSql = StrSql & " From gti_acumdiario "
                StrSql = StrSql & " INNER JOIN tiph_con ON tiph_con.thnro = gti_acumdiario.thnro "
                StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = tiph_con.concnro "
                StrSql = StrSql & " WHERE adfecha >= " & ConvFecha(fecdesde)
                StrSql = StrSql & " AND adfecha <= " & ConvFecha(fechasta)
                StrSql = StrSql & " AND adfecha IN(" & Replace(listaFeriado, ",", "", 1, 1) & ")"
                StrSql = StrSql & " AND ternro=" & ternro
                StrSql = StrSql & " AND gti_acumdiario.thnro IN (" & th & ")"
                StrSql = StrSql & " AND tiph_con.concnro IN (" & concPresentismo & ")"
                StrSql = StrSql & " GROUP BY gti_acumdiario.thnro , conccod "
                OpenRecordset StrSql, rs
                If Not rs.EOF Then
                    If Not EsNulo(rs!Horas) Then
                        If esRot Then
                            'Total de Ausencias*3 = Porcentaje
                            'Si Porcentaje > 9, Porcentaje = 9
                            'Si Porcentaje = 3, descuento = 100
                            'Si Porcentaje = 6, descuento = 200
                            'Si Porcentaje = 9, descuento = 300
                            porc = CDbl(rs!Horas) * 3
                            If porc > 9 Then
                                porc = 9
                            End If
                            Select Case porc
                                Case 3:
                                    descuento = 100
                                Case 6:
                                    descuento = 200
                                Case 9:
                                    descuento = 300
                            End Select
                        Else
                            'semana no cal
                            'Si Total de Ausencias=1 Porcentaje=4, Descuento 100
                            'Si Total de Ausencias>1 Porcentaje=8, Descuento 200
                            If (CDbl(rs!Horas) > 1) Then
                                descuento = 100
                            Else
                                If CDbl(rs!Horas) = 1 Then
                                    descuento = 200
                                Else
                                    descuento = 0
                                End If
                            End If
                        End If
                        StrSql = " INSERT INTO rep_planilla_dtos_det "
                        StrSql = StrSql & "("
                        StrSql = StrSql & " repnro, repbpronro, reptiporeporte,  "
                        StrSql = StrSql & " replegajo, repnombre, repapellido, "
                        StrSql = StrSql & " repdoc, repconcepto, reptipohora, repcantidad "
                        StrSql = StrSql & ")"
                        StrSql = StrSql & " VALUES "
                        StrSql = StrSql & "("
                        StrSql = StrSql & nroReporte & "," & NroProcesoBatch & ",1,"
                        StrSql = StrSql & Legajo & ",'" & Nombre & "','" & Apellido & "',"
                        StrSql = StrSql & "'" & nrodoc & "'," & rs!conccod & "," & rs!thnro & "," & descuento
                        StrSql = StrSql & ")"
                        objConn.Execute StrSql, , adExecuteNoRecords
                    End If
                Else
                    Flog.writeline "El empleado: " & Legajo & " no tiene cantidad de horas."
                End If
                rs.Close
            End If
        End If
    Else
        Flog.writeline "No hay tipos de hora"
    End If
    TiempoAcumulado = GetTickCount
    Progreso = Progreso + IncPorc
    Flog.writeline "Actualizo el progreso: " & Progreso
    StrSql = "UPDATE batch_proceso SET bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "', bprcprogreso = " & CLng(Progreso) & " WHERE bpronro = " & NroProcesoBatch
    objConnProgreso.Execute StrSql, , adExecuteNoRecords
    Flog.writeline "Se actualizo el progreso"
    'Progreso = Progreso + IncPorc
    'Progreso = Replace(Progreso, ",", ".")
    'StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & " WHERE bpronro = " & NroProcesoBatch
    'objConnProgreso.Execute StrSql, , adExecuteNoRecords
Next j
End Sub
Sub repExtras(ByVal NroProcesoBatch As Long, ByVal nroReporte As Integer, ByVal listaEmpleados As String, ByVal tipoFiltro As Integer, ByVal fecdesde As String, ByVal fechasta As String)

Dim rs As New ADODB.Recordset

Dim concExtras As String
Dim teAmbienteContable As Integer
Dim teCentroCosto As Integer
Dim th As String
Dim Empleado
Dim cantdias As Integer
Dim j As Integer
Dim ternro As Long
Dim fecdesdeaux As String
Dim objempleado As New datosPersonales
Dim Nombre As String
Dim Apellido As String
Dim Legajo As String
Dim nrodoc As String
Dim estrAmbContable As String
Dim estrCentroCosto As String
Dim anio As Integer
Dim mes As Integer

th = "0"
anio = Year(fecdesde)
mes = Month(fecdesde)
concExtras = "0"
teAmbienteContable = 0
teCentroCosto = 0

'Levanto del confrep los tipos de Horas
StrSql = " SELECT * FROM confrepadv "
StrSql = StrSql & " WHERE repnro=489"
OpenRecordset StrSql, rs
If Not rs.EOF Then
    Do While Not rs.EOF
        Select Case rs!confnrocol
            Case 10:
                If UCase(rs!conftipo) = "CO3" Then
                    concExtras = IIf(EsNulo(rs!confval), 0, rs!confval)
                End If
            Case 10:
                If UCase(rs!conftipo) = "TE" Then
                    teAmbienteContable = IIf(EsNulo(rs!confval), 0, rs!confval)
                End If
            Case 11:
                If UCase(rs!conftipo) = "TE" Then
                    teCentroCosto = IIf(EsNulo(rs!confval), 0, rs!confval)
                End If
        End Select
    rs.MoveNext
    Loop
End If
rs.Close
'hasta aca

'Busco los tipo de hora de los conceptos
th = buscarTipoHora2(concExtras)
'th = buscarTipoHora(empDocPais, aConcepto, concPresentismo)

Empleado = Split(listaEmpleados, ",")

'calculo la diferecia de dias
cantdias = DateDiff("d", fecdesde, fechasta)
'hasta aca

Dim listaFeriado
Dim esFeriado

For j = 1 To UBound(Empleado) 'arranco del 1 porque en el indice 0 tengo como legajo el 0
    ternro = Empleado(j)
    fecdesdeaux = fecdesde
    
    'busco los datos del empleado
    objempleado.buscarDatosPersonales (ternro)
    Nombre = objempleado.obtenerNombreApellido("nombre") & " " & objempleado.obtenerNombreApellido("nombre2")
    Apellido = objempleado.obtenerNombreApellido("apellido") & " " & objempleado.obtenerNombreApellido("apellido2")
    Legajo = objempleado.obtenerLegajo
    objempleado.buscarNroDoc ternro, 0
    nrodoc = objempleado.obtenerNroDoc
    'hasta aca
    
   ' Estructura  ambiente contable
    StrSql = " SELECT estructura.estrcodext "
    StrSql = StrSql & " From Empleado "
    StrSql = StrSql & " INNER JOIN his_estructura ON empleado.ternro = his_estructura.ternro AND his_estructura.tenro = " & teAmbienteContable
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro"
    StrSql = StrSql & " AND ("
    StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(fecdesde) & " AND (his_estructura.htethasta is null or his_estructura.htethasta >= " & ConvFecha(fechasta) & ""
    StrSql = StrSql & " or his_estructura.htethasta >= " & ConvFecha(fecdesde) & ")) OR"
    StrSql = StrSql & " (his_estructura.htetdesde >= " & ConvFecha(fecdesde) & " AND (his_estructura.htetdesde <= " & ConvFecha(fechasta) & "))"
    StrSql = StrSql & " )"
    StrSql = StrSql & " Where Empleado.Ternro =" & ternro
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        estrAmbContable = rs!estrcodext
    Else
        estrAmbContable = ""
        Flog.writeline "No se encontro la estructura ambiente contable para el legajo: " & Legajo
    End If
    rs.Close

   ' Estructura centro de costo
    StrSql = " SELECT estructura.estrcodext "
    StrSql = StrSql & " From Empleado "
    StrSql = StrSql & " INNER JOIN his_estructura ON empleado.ternro = his_estructura.ternro AND his_estructura.tenro = " & teCentroCosto
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro"
    StrSql = StrSql & " AND ("
    StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(fecdesde) & " AND (his_estructura.htethasta is null or his_estructura.htethasta >= " & ConvFecha(fechasta) & ""
    StrSql = StrSql & " or his_estructura.htethasta >= " & ConvFecha(fecdesde) & ")) OR"
    StrSql = StrSql & " (his_estructura.htetdesde >= " & ConvFecha(fecdesde) & " AND (his_estructura.htetdesde <= " & ConvFecha(fechasta) & "))"
    StrSql = StrSql & " )"
    StrSql = StrSql & " Where Empleado.Ternro =" & ternro
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        estrCentroCosto = rs!estrcodext
    Else
        estrCentroCosto = ""
        Flog.writeline "No se encontro la estructura centro de costo para el legajo: " & Legajo
    End If
    rs.Close
    
    'busco la suma de las horas
    StrSql = " SELECT sum(adcanthoras) horas,gti_acumdiario.thnro, conccod "
    StrSql = StrSql & " From gti_acumdiario "
    StrSql = StrSql & " INNER JOIN tiph_con ON tiph_con.thnro = gti_acumdiario.thnro "
    StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = tiph_con.concnro "
    StrSql = StrSql & " WHERE adfecha >= " & ConvFecha(fecdesde)
    StrSql = StrSql & " AND adfecha <= " & ConvFecha(fechasta)
    StrSql = StrSql & " AND ternro=" & ternro
    StrSql = StrSql & " AND gti_acumdiario.thnro IN (" & th & ")"
    StrSql = StrSql & " AND tiph_con.concnro IN (" & concExtras & ")"
    StrSql = StrSql & " GROUP BY gti_acumdiario.thnro , conccod "
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        If Not EsNulo(rs!Horas) Then
            StrSql = " INSERT INTO rep_planilla_dtos_det "
            StrSql = StrSql & "("
            StrSql = StrSql & " repnro, repbpronro, reptiporeporte,  "
            StrSql = StrSql & " repanio, repmes,"
            StrSql = StrSql & " replegajo, repnombre, repapellido, "
            StrSql = StrSql & " repdoc, repconcepto, reptipohora, repcantidad, "
            StrSql = StrSql & " reptenro1, repestrnro1, reptenro2, repestrnro2 "
            StrSql = StrSql & ")"
            StrSql = StrSql & " VALUES "
            StrSql = StrSql & "("
            StrSql = StrSql & nroReporte & "," & NroProcesoBatch & ",3,"
            StrSql = StrSql & anio & ",'" & mes & "',"
            StrSql = StrSql & Legajo & ",'" & Nombre & "','" & Apellido & "',"
            StrSql = StrSql & "'" & nrodoc & "'," & rs!conccod & "," & rs!thnro & "," & rs!Horas & ","
            StrSql = StrSql & teAmbienteContable & ",'" & estrAmbContable & "',"
            StrSql = StrSql & teCentroCosto & ",'" & estrCentroCosto & "'"
            StrSql = StrSql & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
    Else
        Flog.writeline "El empleado: " & Legajo & " no tiene cantidad de horas."
    End If
    rs.Close
    TiempoAcumulado = GetTickCount
    Progreso = Progreso + IncPorc
    Flog.writeline "Actualizo el progreso: " & Progreso
    StrSql = "UPDATE batch_proceso SET bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "', bprcprogreso = " & CLng(Progreso) & " WHERE bpronro = " & NroProcesoBatch
    objConnProgreso.Execute StrSql, , adExecuteNoRecords
    Flog.writeline "Se actualizo el progreso"
    
    'Progreso = Progreso + IncPorc
    'Progreso = Replace(Progreso, ",", ".")
    'StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & " WHERE bpronro = " & NroProcesoBatch
    'objConnProgreso.Execute StrSql, , adExecuteNoRecords
Next j




End Sub
'Function buscarTipoHora(ByVal concPresentismo As String)
Function buscarTipoHora(ByVal paisEmpleado As Integer, ByRef conceptos() As conc_pais, ByRef concPresentismo As String, Optional ByVal turno As String)
Dim StrSql As String
Dim rs As New ADODB.Recordset
Dim th As String
Dim j As Integer

concPresentismo = "0"

For j = 1 To UBound(conceptos)
    If (CInt(paisEmpleado) = CInt(conceptos(j).pais)) Then
        If turno <> "0" And turno <> "" Then
            If turno = conceptos(j).turno Then
                concPresentismo = concPresentismo & "," & conceptos(j).concnro
            End If
        Else
            concPresentismo = concPresentismo & "," & conceptos(j).concnro
        End If
    End If
Next j

StrSql = " SELECT DISTINCT thnro FROM tiph_con "
StrSql = StrSql & " WHERE concnro IN (" & concPresentismo & ")"
OpenRecordset StrSql, rs
If Not rs.EOF Then
    Do While Not rs.EOF
        th = th & "," & rs!thnro
    rs.MoveNext
    Loop
End If
rs.Close
th = Replace(th, ",", "", 1, 1)
buscarTipoHora = th

End Function

Function buscarTipoHora2(ByVal concepto As String)
Dim StrSql As String
Dim rs As New ADODB.Recordset
Dim th As String


StrSql = " SELECT DISTINCT thnro FROM tiph_con "
StrSql = StrSql & " WHERE concnro IN (" & concepto & ")"
OpenRecordset StrSql, rs
If Not rs.EOF Then
    Do While Not rs.EOF
        th = th & "," & rs!thnro
    rs.MoveNext
    Loop
End If
rs.Close
th = Replace(th, ",", "", 1, 1)
buscarTipoHora2 = th

End Function
