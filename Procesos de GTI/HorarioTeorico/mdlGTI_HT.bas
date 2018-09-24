Attribute VB_Name = "mdlGTI_HT"
Option Explicit

'Const Version = "1.00"
'Const FechaVersion = "16/08/2013"

'Const Version = "1.01"
'Const FechaVersion = "14/01/2015" ' Gonzalez Nicolás - CAS-28429 - Villa Maria - Custom GTI Se permiten procesar registros pasados.

' ----------------------------------------------------------------------------------------------------------------
Const Version = "1.02"
Const FechaVersion = "30/01/2015" ' FGZ - CAS-28429 - Villa Maria - Custom GTI Se permiten procesar registros pasados.
'                               Mejoras en mensajes de logs

' ----------------------------------------------------------------------------------------------------------------


Global CEmpleadosAProc As Integer
Global CDiasAProc As Integer
Global IncPorc As Single
Global IncPorcEmpleado As Single
Global HuboErrores As Boolean
Global EmpleadoSinError As Boolean
Global Progreso As Single
Global ProgresoEmpleado As Single
Global fec_proc As Integer ' 1 - Política Primer Reg.
                           ' 2 - Política Reg. del Turno
                           ' 3 - Política Ultima Reg.
Global Usa_Conv As Boolean
Global objBTurno As New BuscarTurno
Global objBDia As New BuscarDia
Global objFeriado As New Feriado
Global diatipo As Byte
Global ok As Boolean
Global esFeriado As Boolean
Global hora_desde As String
Global fecha_desde As Date
Global fecha_hasta As Date
Global Hora_desde_aux As String
Global hora_hasta As String
Global Hora_Hasta_aux As String
Global No_Trabaja_just As Boolean
Global nro_jus_ent As Long
Global nro_jus_sal As Long
Global Total_horas As Single
Global Tdias As Integer
Global Thoras As Integer
Global Tmin As Integer
Global Cod_justificacion1 As Long
Global Cod_justificacion2 As Long

Global Horas_Oblig As Single

Global Forma_embudo  As Boolean

Global tiene_turno As Boolean
Global Nro_Turno As Long
Global Tipo_Turno As Integer

Global Tiene_Justif As Boolean
Global Justif_Completa As Boolean
Global Nro_Justif As Long
Global justif_turno As Boolean
Global p_turcomp As Boolean
Global Nro_Grupo As Long
Global Nro_fpgo As Integer
Global Fecha_Inicio As Date
Global P_Asignacion  As Boolean
Global Trabaja     As Boolean ' Indica si trabaja para ese dia
Global Orden_Dia As Integer
Global Nro_Dia As Integer
Global Nro_Subturno As Integer
Global Dia_Libre As Boolean
Global Dias_trabajados As Integer
Global Dias_laborables As Integer

Global Aux_Tipohora As Integer
Global aux_TipoDia As Integer

Global Hora_Tol As String
Global Fecha_Tol As Date
Global hora_toldto As String
Global fecha_toldto As Date



Global fv1 As Date
Global fv2 As Date
Global fv3 As Date
Global fv4 As Date
Global fv5 As Date
Global fv6 As Date
Global fv7 As Date

Global v1 As String
Global v2 As String
Global v3 As String
Global v4 As String
Global v5 As String
Global v6 As String
Global v7 As String

Global tol As String

Global Cant_emb As Integer
Global toltemp As String
Global toldto As String
Global acumula As Boolean
Global acumula_dto As Boolean
Global acumula_temp As Boolean
Global convenio As Long

Global tdias_oblig As Single



'---------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Sub Main()

Dim Archivo As String
Dim pos As Integer
Dim strcmdLine  As String

'Dim objconnMain As New ADODB.Connection
Dim rs As New ADODB.Recordset

Dim NroProceso As Long
Dim NroReporte As Long
Dim StrParametros As String

Dim rs_Batch_Proceso As New ADODB.Recordset
Dim rs_His_Batch_Proceso As New ADODB.Recordset
Dim myrs As New ADODB.Recordset

Dim PID As String
Dim ArrParametros
Dim Periodo As String

Dim Ternro As Long
Dim FDesde As Date
Dim FHasta As Date

ReDim arrEmpProcesadosOK(0)
Dim listEmpleadoOk As String
Dim nroMaxEmpBorrar As Long
Dim topeListaBorrar As Long
Dim cantProgreso As Double
Dim totalProgreso As Double
Dim horaUpdateBD As String

Dim tiempoInicial
Dim tiempoActual

Dim i As Long

Dim ListaPar

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
    'Archivo = PathFLog & "RepAusentismo-" & CStr(NroProceso) & Format(Now, "DD-MM-YYYY") & ".log"
    Archivo = PathFLog & "HorarioTeorico-" & CStr(NroProceso) & ".log"

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

    On Error GoTo ce


    'Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaVersion
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-------------------------------------------------"
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "Inicio :" & Now
    
    
    Version_Valida = ValidarV(Version, 1, TipoBD)
    If Not Version_Valida Then
        StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Error de Version', bprcpid = " & PID & " WHERE bpronro = " & NroProceso
        objConnProgreso.Execute StrSql, , adExecuteNoRecords
    
        Flog.writeline
        GoTo Final
    End If
    

    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprcpid = " & PID & ", bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords

    'FGZ - 16/01/2015 ---------------------------------------------------
    'Flog.writeline "Levanta Proceso y Setea Parámetros:  " & " " & Now
    Flog.writeline Espacios(Tabulador * 0) & "Levanta Proceso y Setea Parámetros: " & Now
    'FGZ - 16/01/2015 ---------------------------------------------------
    
    'levanto los parametros del proceso
    StrSql = "SELECT bprcfecdesde,bprcfechasta,bprcparam,iduser FROM batch_proceso WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, rs
    If rs.EOF Then
        Flog.writeline "No se encontro el proceso " & NroProceso
        Exit Sub
    End If
    
   
    'FGZ - 16/01/2015 ----------------------------------------------------------
'    'ListaPar = Split(objRs!bprcparam, ".", -1)
'    Periodo = ""
'    If Not EsNulo(rs!bprcparam) Then
'        'le estoy pasando el periodo a procesar
'        Periodo = "01/" & Right(rs!bprcparam, 2) & "/" & Left(rs!bprcparam, 4)
'        If IsDate(Periodo) Then
'
'            'nunca proceso dias pasados
''            If CDate(Periodo) > Date Then
''                FDesde = CDate(Periodo)
''            Else
''                FDesde = Date
''            End If
'
'            FDesde = CDate(Periodo) 'V 1.01
'            FHasta = DateAdd("d", -1, DateAdd("m", 1, CDate("01/" + Format(FDesde, "mm/yyyy"))))
'        Else
'            Flog.writeline "Error en Parametro. Debe ser YYYYMM."
'            Exit Sub
'        End If
'    Else
'        FDesde = Date
'        'si no le paso el periodo proceso dos meses para adelante
'        FHasta = DateAdd("d", -1, DateAdd("m", 3, CDate("01/" + Format(FDesde, "mm/yyyy"))))
'    End If
    
    Periodo = ""
    depurar = False
    If Not EsNulo(rs!bprcparam) Then
        If InStr(1, rs!bprcparam, ".") <> 0 Then
            ListaPar = Split(rs!bprcparam, ".", -1)
            'Periodo
            Periodo = "01/" & Right(ListaPar(0), 2) & "/" & Left(ListaPar(0), 4)
            
            'Detalle de log
            If UBound(ListaPar) > 0 Then
                depurar = IIf(IsNumeric(ListaPar(1)), CBool(ListaPar(1)), False)
            Else
                depurar = False
            End If
        Else
            Periodo = "01/" & Right(rs!bprcparam, 2) & "/" & Left(rs!bprcparam, 4)
            If IsDate(Periodo) Then
                FDesde = CDate(Periodo) 'V 1.01
                FHasta = DateAdd("d", -1, DateAdd("m", 1, CDate("01/" + Format(FDesde, "mm/yyyy"))))
            Else
                Flog.writeline "Error en Parametro Periodo. Debe ser YYYYMM."
                Exit Sub
            End If
        End If
    Else
        depurar = False
        FDesde = Date
        'si no le paso el periodo proceso dos meses para adelante
        FHasta = DateAdd("d", -1, DateAdd("m", 3, CDate("01/" + Format(FDesde, "mm/yyyy"))))
    End If
    
    
    Flog.writeline Espacios(Tabulador * 1) & "Periodo       : " & Periodo
    Flog.writeline Espacios(Tabulador * 1) & "  Desde       : " & FDesde
    Flog.writeline Espacios(Tabulador * 1) & "  Hasta       : " & FHasta
    Flog.writeline Espacios(Tabulador * 1) & "Log detallado : " & depurar
    Flog.writeline
    'FGZ - 16/01/2015 ----------------------------------------------------------
    
    
    StrSql = "SELECT empleado.ternro, empleado.empleg FROM batch_empleado "
    StrSql = StrSql & " INNER JOIN empleado ON batch_empleado.ternro = empleado.Ternro "
    StrSql = StrSql & " WHERE batch_empleado.bpronro = " & NroProceso
    If myrs.State = adStateOpen Then myrs.Close
    OpenRecordset StrSql, myrs
    If myrs.EOF Then
        Flog.writeline "No se encontraron legajos a procesar!"
        Exit Sub
    End If
    
    CEmpleadosAProc = myrs.RecordCount
    If CEmpleadosAProc = 0 Then CEmpleadosAProc = 1
    CDiasAProc = DateDiff("d", FDesde, FHasta) + 1
    IncPorc = (100 / CEmpleadosAProc)
    '* (100 / CDiasAProc))
    IncPorcEmpleado = (100 / CDiasAProc)
    Flog.writeline Espacios(Tabulador * 0) & "Empleados a procesar: " & CEmpleadosAProc
    Flog.writeline Espacios(Tabulador * 0) & "Fechas a procesar. Desde: " & FDesde & " Hasta: " & FHasta
    Flog.writeline
    
    Progreso = 0
    totalProgreso = cantProgreso
    
    horaUpdateBD = Time
    TiempoDeEsperaNoResponde = 5
    
    tiempoInicial = GetTickCount
    
    Do While Not myrs.EOF
        
        'EmpleadoSinError = CalculoHT(710, FDesde, FHasta, NroProceso)
        'Flog.writeline "Tercero: " & myrs!Ternro
        If depurar Then
            Flog.writeline
            Flog.writeline Espacios(Tabulador * 1) & "Empleado: " & myrs!empleg
        End If
        
        EmpleadoSinError = CalculoHT(myrs!Ternro, FDesde, FHasta, NroProceso)
        
        Progreso = Progreso + IncPorc
   
        'EAM- Si el progreso es el configurado o esta por cumplirse el horario de espera del appserver actualiza en bach_proceso (v6.00)
        If (Progreso > totalProgreso) Or (DateDiff("s", Format(horaUpdateBD, "HH:mm:ss"), Format(Time, "HH:mm:ss")) >= (TiempoDeEsperaNoResponde - 1)) Then
                   
            'Actualizo progreso
            tiempoActual = GetTickCount
            StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & ", bprctiempo=" & (tiempoActual - tiempoInicial) & " WHERE bpronro = " & NroProceso
            objConnProgreso.Execute StrSql, , adExecuteNoRecords
            
            'Elimino Empleado de batch_empleado 1.01
            StrSql = "DELETE FROM batch_empleado WHERE bpronro = " & NroProceso & " AND ternro = " & myrs!Ternro
            objConnProgreso.Execute StrSql, , adExecuteNoRecords

            
            totalProgreso = totalProgreso + cantProgreso
            horaUpdateBD = Time
        End If
            
            
        If EmpleadoSinError Then
            ReDim Preserve arrEmpProcesadosOK(UBound(arrEmpProcesadosOK) + 1)
            arrEmpProcesadosOK(UBound(arrEmpProcesadosOK)) = Ternro
        End If
        
        myrs.MoveNext
    Loop
    
    'EAM- Elimina todos los empleados procesados correctamente
    listEmpleadoOk = 0
    topeListaBorrar = 900
    nroMaxEmpBorrar = topeListaBorrar
    
    If (UBound(arrEmpProcesadosOK) > topeListaBorrar) Then
    
        For i = 1 To UBound(arrEmpProcesadosOK)
            listEmpleadoOk = listEmpleadoOk & "," & arrEmpProcesadosOK(i)
            
            'EAM- Borro los empleados armados en la lista
            If (nroMaxEmpBorrar <= i) Then
                StrSql = "DELETE FROM batch_empleado WHERE bpronro = " & NroProceso & " And Ternro IN (" & listEmpleadoOk & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
                            
                If depurar Then
                    Flog.writeline Espacios(Tabulador * 1) & "---> " & Now & " Se borraron los Empleados " & listEmpleadoOk
                End If
                listEmpleadoOk = ""
                nroMaxEmpBorrar = nroMaxEmpBorrar + topeListaBorrar
            End If
        Next
    Else
        For i = 1 To UBound(arrEmpProcesadosOK)
            listEmpleadoOk = listEmpleadoOk & "," & arrEmpProcesadosOK(i)
        Next
        
        StrSql = "DELETE FROM batch_empleado WHERE bpronro = " & NroProceso & " And Ternro IN (" & listEmpleadoOk & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
                    
        If depurar Then
            Flog.writeline Espacios(Tabulador * 1) & "---> " & Now & " Se borraron los Empleados " & listEmpleadoOk
        End If
        listEmpleadoOk = ""
    End If
    
    
    ' poner el bprcestado en procesado
    StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado', bprcprogreso = 100 WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords

    ' -----------------------------------------------------------------------------------
    'Copio el proceso en el historico de batch_proceso y lo borro de batch_proceso
        StrSql = "SELECT * FROM batch_proceso WHERE bpronro =" & NroProceso
        OpenRecordset StrSql, rs_Batch_Proceso

        StrSql = "INSERT INTO His_Batch_Proceso (bpronro,btprcnro,bprcfecha,iduser"
        StrSqlDatos = rs_Batch_Proceso!bpronro & "," & rs_Batch_Proceso!btprcnro & "," & _
                 ConvFecha(rs_Batch_Proceso!bprcfecha) & ",'" & rs_Batch_Proceso!IdUser & "'"
        
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
    ' -----------------------------------------------------------------------------------


    If objConn.State = adStateOpen Then objConn.Close
    If objConnProgreso.State = adStateOpen Then objConnProgreso.Close
    
    Set rs_Batch_Proceso = Nothing
    Set rs_His_Batch_Proceso = Nothing

Final:
    Flog.writeline Espacios(Tabulador * 0) & "Fin de Planificacion de Horario Teorico: " & " " & Now
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "---------------------------------------------------------------------------------"
    Flog.Close
Exit Sub

ce:
    Flog.writeline Espacios(Tabulador * 0) & "Proceso abortado por Error"
    Flog.writeline Espacios(Tabulador * 1) & "Error:" & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "Ultimo SQL: " & StrSql
    'MyRollbackTrans
    StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
End Sub

Public Function CalculoHT(Ternro As Long, FDesde As Date, FHasta As Date, NroProceso As Long) As Boolean
Dim f As Long
Dim Fecha As Date

Dim Emp As Templeado
Dim Reg_Dia As TDia
Dim esFeriado As Boolean
Dim rs_gti_Proc_Emp As New ADODB.Recordset
Dim Horas As Integer

Dim Reg_T As TipoTurno
Dim Justif As TJustif
Dim Turno As TTurno
    
    
    On Error GoTo ME_Local
    
    Debug.Print Ternro
    
    'borro todos los registros del empleado par el rango de fechas que voy a generar
    'Flog.writeline "Elimino todos los registros del empleado par el rango de fechas que voy a generar"
   
    StrSql = "DELETE FROM gti_proc_emp_plan WHERE ternro =" & Ternro
    StrSql = StrSql & " AND fecha >= " & ConvFecha(FDesde) & " AND fecha <= " & ConvFecha(FHasta)
    objConn.Execute StrSql, , adExecuteNoRecords
           
        
    For f = 0 To DateDiff("d", FDesde, FHasta)
        
        Fecha = DateAdd("d", f, FDesde)
        
        esFeriado = EsFeriado_Nuevo(Fecha, Ternro, False)
        
        'Flog.writeline "Día: " & Fecha
        

        Justif.Numero = 0
        Justif.Tiene_Justif = False
        Justif.justif_turno = False
               
                
        Call Buscar_Turno_Nuevo(Fecha, Ternro, False, Reg_T, Turno, Justif, Emp)
    
        If Reg_T.tiene_turno Then
            Call Buscar_Dia_Nuevo(Fecha, Reg_T.Fecha_Inicio, Reg_T.Nro_Turno, Ternro, Reg_T.P_Asignacion, depurar, Reg_Dia)
                
                'Flog.writeline "Emp.Grupo: " & Emp.Grupo
                
                Horas = Buscar_horas(Ternro, Reg_T.P_Asignacion, Fecha, Reg_Dia.Nro_Dia)
                
                StrSql = "INSERT INTO gti_proc_emp_plan (ternro,fecha,turnro,fpgnro,dianro,feriado,jusnro,pasig,dialibre,trabaja,estrnro, horast"
               
                StrSql = StrSql & " ) VALUES ("
                StrSql = StrSql & Ternro & ","
                StrSql = StrSql & ConvFecha(Fecha) & ","
                StrSql = StrSql & Reg_T.Nro_Turno & ","
                StrSql = StrSql & Reg_T.Nro_FPago & ","
                StrSql = StrSql & Reg_Dia.Nro_Dia & ","
                StrSql = StrSql & CInt(esFeriado) & ","
                StrSql = StrSql & Justif.Numero & ","
                StrSql = StrSql & CInt(Reg_T.P_Asignacion) & ","
                StrSql = StrSql & CInt(Reg_Dia.Dia_Libre) & ","
                StrSql = StrSql & CInt(Reg_Dia.Trabaja) & ","
                StrSql = StrSql & Emp.Grupo & ","
                StrSql = StrSql & Horas
                StrSql = StrSql & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            If rs_gti_Proc_Emp.State = adStateOpen Then rs_gti_Proc_Emp.Close

        End If
        
    Next
    
    Set rs_gti_Proc_Emp = Nothing
    
    CalculoHT = True
    
Exit Function
    
ME_Local:
    Set rs_gti_Proc_Emp = Nothing
    CalculoHT = False
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "***"
    Flog.writeline Espacios(Tabulador * 1) & " ---------------------------------------------------------------------------------------------------"
    Flog.writeline Espacios(Tabulador * 1) & "Error generando gti_proc_emp_plan. La informacion del horario teorico en el tablero no estará disponible."
    Flog.writeline Espacios(Tabulador * 2) & "Error: " & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 2) & "SQL: " & StrSql
    Flog.writeline Espacios(Tabulador * 1) & " ---------------------------------------------------------------------------------------------------"
    Flog.writeline Espacios(Tabulador * 1) & "***"
    Flog.writeline
End Function


Public Function Buscar_horas(ByVal Ternro As Long, ByVal pasig As Boolean, ByVal Fecha As Date, ByVal diaNro As Long) As Integer
        
   
    Dim a As String
    Dim objRs As New ADODB.Recordset

    If pasig Then
        StrSql = "SELECT diacanthoras,diamaxhoras,diaminhoras FROM gti_detturtemp WHERE (ternro =" & Ternro & ") AND " & _
                 "(gttempdesde <= " & ConvFecha(Fecha) & ") AND " & _
                 "(" & ConvFecha(Fecha) & " <= gttemphasta)"
        OpenRecordset StrSql, objRs
        If Not objRs.EOF Then
              Buscar_horas = objRs!diacanthoras
        Else
            Buscar_horas = 0
        End If
    Else
        StrSql = "SELECT diacanthoras,diamaxhoras,diaminhoras FROM gti_dias WHERE dianro = " & diaNro
        OpenRecordset StrSql, objRs
        If Not objRs.EOF Then
               Buscar_horas = objRs!diacanthoras
        Else
            Buscar_horas = 0
        End If
    End If
    

    If objRs.State = adStateOpen Then objRs.Close
    Set objRs = Nothing
    
    
End Function





Public Sub Insertar_GTI_Proc_Emp_plan(ByVal Ternro As Long, ByVal Fecha As Date)
' --------------------------------------------------------------
' Descripcion: Genera la informacion del dia procesado.
' Autor: FGZ - 27/10/2005
' Ultima modificacion: FGZ - 30/11/2006 modificaciones para Horario_Flexible_Rotativo
'                      Diego Rosso - 12/11/2007 - PARA HORARIO FLEXIBLE: Cuando ModificaHT es falso graba el horario que le corresponderia en el dia
'                                                   y cuando es verdadero graba el numero de dia calculado.
' --------------------------------------------------------------
Dim rs_gti_Proc_Emp As New ADODB.Recordset

        Dim Reg_T As TipoTurno
        Dim Justif As TJustif
        Dim Turno As TTurno
        Dim Emp As Templeado
        Dim Reg_Dia As TDia
        Dim esFeriado As Boolean
        
        
        On Error GoTo ME_Local
        esFeriado = EsFeriado_Nuevo(Fecha, Ternro, False)
        
        
        Call Buscar_Turno_Nuevo(Fecha, Ternro, False, Reg_T, Turno, Justif, Emp)
    
    
        If Reg_T.tiene_turno Then
            Call Buscar_Dia_Nuevo(Fecha, Reg_T.Fecha_Inicio, Reg_T.Nro_Turno, Ternro, Reg_T.P_Asignacion, depurar, Reg_Dia)


            StrSql = "SELECT ternro FROM gti_proc_emp_plan WHERE ternro =" & Ternro
            StrSql = StrSql & " AND fecha = " & ConvFecha(Fecha)
            OpenRecordset StrSql, rs_gti_Proc_Emp
            If rs_gti_Proc_Emp.EOF Then
                StrSql = "INSERT INTO gti_proc_emp_plan (ternro,fecha,turnro,fpgnro,dianro,feriado,jusnro,pasig,dialibre,trabaja,estrnro"
               
                StrSql = StrSql & " ) VALUES ("
                StrSql = StrSql & Ternro & ","
                StrSql = StrSql & ConvFecha(Fecha) & ","
                StrSql = StrSql & Reg_T.Nro_Turno & ","
                StrSql = StrSql & Reg_T.Nro_FPago & ","
                StrSql = StrSql & Reg_Dia.Nro_Dia & ","
                StrSql = StrSql & CInt(esFeriado) & ","
                StrSql = StrSql & Reg_T.Nro_Justif & ","
                StrSql = StrSql & CInt(Reg_T.P_Asignacion) & ","
                StrSql = StrSql & CInt(Reg_Dia.Dia_Libre) & ","
                StrSql = StrSql & CInt(Reg_Dia.Trabaja) & ","
                StrSql = StrSql & Emp.Grupo
                StrSql = StrSql & ")"
            Else
                StrSql = "UPDATE gti_proc_emp_plan SET "
                StrSql = StrSql & " turnro = " & Reg_T.Nro_Turno
                StrSql = StrSql & ",fpgnro = " & Reg_T.Nro_FPago
                StrSql = StrSql & ",dianro = " & Reg_Dia.Nro_Dia
                StrSql = StrSql & ",feriado = " & CInt(esFeriado)
                StrSql = StrSql & ",jusnro = " & Reg_T.Nro_Justif
                StrSql = StrSql & ",pasig = " & CInt(Reg_T.P_Asignacion)
                StrSql = StrSql & ",dialibre = " & CInt(Reg_Dia.Dia_Libre)
                StrSql = StrSql & ",trabaja = " & CInt(Reg_Dia.Trabaja)
                StrSql = StrSql & ",estrnro = " & Emp.Grupo
                StrSql = StrSql & " WHERE TERNRO =" & Ternro
                StrSql = StrSql & " AND fecha = " & ConvFecha(Fecha)
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
        

        End If
              
        
        
        
        If depurar Then
            Flog.writeline
            Flog.writeline Espacios(Tabulador * 1) & "Dia Procesado "
            Flog.writeline Espacios(Tabulador * 1) & "SQL " & StrSql
            Flog.writeline
        End If
        
'        Flog.writeline
'        Flog.writeline Espacios(Tabulador * 1) & "Dia Procesado "
'        Flog.writeline Espacios(Tabulador * 1) & "--------------------------------------"
        
    'Libero y dealoco
    If rs_gti_Proc_Emp.State = adStateOpen Then rs_gti_Proc_Emp.Close
    Set rs_gti_Proc_Emp = Nothing
Exit Sub

ME_Local:
    Flog.writeline
    Flog.writeline "***"
    Flog.writeline " ---------------------------------------------------------------------------------------------------"
    Flog.writeline "Error generando gti_proc_emp_plan. La informacion del horario teorico en el tablero no estará disponible."
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "SQL: " & StrSql
    Flog.writeline " ---------------------------------------------------------------------------------------------------"
    Flog.writeline "***"
    Flog.writeline
End Sub
