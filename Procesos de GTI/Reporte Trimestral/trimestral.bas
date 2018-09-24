Attribute VB_Name = "MdlTrimestral"
Option Explicit

' Variables globales necesarias para la integracion con
' el modulo de políticas

' "1.01 - 31/07/2009 - Martin Ferraro - Encriptacion de string connection"

Global G_traza As Boolean
Global fec_proc As Integer

Global diatipo As Byte
Global ok As Boolean
Global esFeriado As Boolean
Global hora_desde As String
Global fecha_desde As Date
Global fecha_hasta As Date
Global Hora_desde_aux As String
Global hora_hasta As String
Global Hora_Hasta_aux As String
Global no_trabaja_just As Boolean
Global nro_jus_ent As Long
Global nro_jus_sal As Long
Global Total_horas As Single
Global Tdias As Integer
Global Thoras As Integer
Global Tmin As Integer
Global Cod_justificacion1 As Long
Global Cod_justificacion2 As Long

Global horas_oblig As Single
Global Existe_Reg As Boolean
Global Forma_embudo  As Boolean

Global tiene_turno As Boolean
Global nro_turno As Long
Global Tipo_Turno As Integer

Global Tiene_Justif As Boolean
Global nro_justif As Long
Global justif_turno As Boolean
Global p_turcomp As Boolean
Global nro_grupo As Long
Global Nro_fpgo As Integer
Global Fecha_Inicio As Date
Global P_asignacion  As Boolean
Global Trabaja     As Boolean ' Indica si trabaja para ese dia
Global Orden_Dia As Integer
Global Nro_Dia As Integer
Global Nro_Subturno As Integer
Global Dia_Libre As Boolean
Global Dias_trabajados As Integer
Global Dias_laborables As Integer

Global aux_Tipohora As Integer
Global aux_TipoDia As Integer
Global Sigo_Generando As Boolean

Global hora_tol As String
Global fecha_tol As Date
Global hora_toldto As String
Global fecha_toldto As Date

Global Usa_Conv  As Boolean

Global tol As String

Global Cant_emb As Integer
Global toltemp As String
Global toldto As String
Global acumula As Boolean
Global acumula_dto As Boolean
Global acumula_temp As Boolean
Global convenio As Long

Global tdias_oblig As Single

Global NroProceso As Long

'Fin de globales

Global InicioTrimestre As Date
Global FinTrimestre As Date
Global Separador As String
Global Archivo As String
Global HorasSemestre As Single
Global MinHorasExtras As Single
Global THorasTrimestre As Integer
Global THorasMinimo As Integer
Global THorasSaldo As Integer
Global THorasAcumula As Integer
Global THorasPaga As Integer

Dim objBTurno As New BuscarTurno



Private Sub initVariablesTurno(ByRef T As BuscarTurno)
   p_turcomp = T.Compensa_Turno
   nro_grupo = T.Empleado_Grupo
   nro_justif = T.Justif_Numero
   justif_turno = T.justif_turno
   Tiene_Justif = T.Tiene_Justif
   Fecha_Inicio = T.FechaInicio
   Nro_fpgo = T.Numero_FPago
   nro_turno = T.Turno_Numero
   tiene_turno = T.tiene_turno
   Tipo_Turno = T.Turno_Tipo
   P_asignacion = T.Tiene_PAsignacion
   
End Sub

Public Sub Main()
Dim Nro_Cab As Long
Dim fdesde As Date
Dim fhasta As Date
Dim Nro_tpr As Long
Dim Nro_Pro As Long
Dim strCmdLine As String
Dim pos1 As Integer
Dim pos2 As Integer
Dim NroProceso As Long
Dim fechaDesde As Date
Dim fechaHasta As Date
Dim Fecha As Date
Dim objrsEmpleado As New ADODB.Recordset
Dim objRs As New ADODB.Recordset
Dim Progreso As Single
Dim CEmpleadosAProc As Integer
Dim IncPorc As Single

Dim rs_Batch_Proceso As New ADODB.Recordset
Dim rs_His_Batch_Proceso As New ADODB.Recordset

Dim PID As String
Dim ArrParametros


    
    'Activo el manejador de errores
    On Error GoTo CE
    
'    strCmdLine = Command()
'    'strCmdLine = "29"
'    If IsNumeric(strCmdLine) Then
'        NroProceso = strCmdLine
'    Else
'        Exit Sub
'    End If

    strCmdLine = Command()
    ArrParametros = Split(strCmdLine, " ", -1)
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
            If IsNumeric(strCmdLine) Then
                NroProceso = strCmdLine
            Else
                Exit Sub
            End If
        End If
    End If
    
    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(PathFLog & "\trimestral" & "-" & NroProceso & ".log", True)

    Flog.writeline "Inicio :" & Now
    
    'Abro la conexion
    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If
    
    On Error GoTo CE
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "PID = " & PID
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprcpid = " & PID & ", bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Set objFechasHoras.Conexion = objConn
    
    StrSql = " SELECT gti_procacum.gtprocnro,gti_procacum.gpadesde,gti_procacum.gpahasta,gti_cab.gpanro,gti_cab.cgtinro,batch_proceso.bpronro,gti_procacum.gtprocnro FROM batch_proceso " & _
             " INNER JOIN batch_procacum ON batch_procacum.bpronro = batch_proceso.bpronro " & _
             " INNER JOIN gti_cab ON  gti_cab.gpanro = batch_procacum.gpanro " & _
             " INNER JOIN gti_procacum ON batch_procacum.gpanro = gti_procacum.gpanro " & _
             " INNER JOIN batch_empleado ON batch_empleado.ternro = gti_cab.ternro AND batch_empleado.bpronro = batch_proceso.bpronro" & _
             " WHERE batch_proceso.bpronro = " & NroProceso
             
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        fechaDesde = objRs!gpadesde
        fechaHasta = objRs!gpahasta
        InicioTrimestre = fechaDesde
        FinTrimestre = fechaHasta
        'para el reporte trimestral
        Call ObtenerConfiguracionTrimestral
    End If
    
    Fecha = fechaDesde
    
    ' Seteo el incremento de progreso
    CEmpleadosAProc = objRs.RecordCount
    IncPorc = (100 / CEmpleadosAProc)
    Progreso = 0
    
    Do While Not objRs.EOF
            
            'MyBeginTrans
inicio:
            
        Flog.writeline "Inicio Cabecera:" & objRs!cgtinro & " " & Fecha

        SumarHoras objRs!cgtinro, fechaDesde, fechaHasta, objRs!gtprocnro, objRs!gpanro
        
            
Siguiente:
            ' Actualizo el progreso
            Progreso = Progreso + IncPorc
            StrSql = "UPDATE batch_proceso SET bprcprogreso = " & CInt(Progreso) & " WHERE bpronro = " & NroProceso
            objConn.Execute StrSql, , adExecuteNoRecords

            objRs.MoveNext
    Loop
        
    Flog.writeline "Fin :" & Now
    
    Flog.Close
    
    StrSql = "DELETE FROM batch_empleado WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
    StrSql = "DELETE FROM Batch_Procacum WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
    StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
    
    ' -----------------------------------------------------------------------------------
    'FGZ - 22/09/2003
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
    ' FGZ - 22/09/2003
    ' -----------------------------------------------------------------------------------
    
    
    If objConn.State = adStateOpen Then objConn.Close
    
    Exit Sub

CE:
    'MyRollbackTrans
    
    Flog.writeline "Error Cabecera" & " " & Fecha
    
    StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
    If objConn.State = adStateOpen Then objConn.Close
    
End Sub


Public Sub Desglose_ACParcial(Nro_Cab As Long, desde As Date, hasta As Date)
Dim StrSql As String
Dim E1 As Integer
Dim E2 As Integer
Dim E3 As Integer
Dim te1 As Integer
Dim te2 As Integer
Dim te3 As Integer
Dim canthoras As Single
Dim l_achpnro As Long
Dim auxi As String
Dim rs As New ADODB.Recordset

Dim fechaDesde As Date
Dim fechaHasta As Date
Dim Fecha As Date

Dim Ternro As Integer

StrSql = "delete from gti_achparc_estr where achpnro in(" & _
"select achpnro from gti_achparcial where cgtinro = " & Nro_Cab & _
")"
objConn.Execute StrSql, , adExecuteNoRecords

StrSql = "delete from gti_achparcial where cgtinro = " & Nro_Cab
objConn.Execute StrSql, , adExecuteNoRecords

StrSql = "select ternro from gti_cab where cgtinro = " & Nro_Cab
OpenRecordset StrSql, objRs
objConn.Execute StrSql, , adExecuteNoRecords

Ternro = objRs!Ternro
    
Fecha = desde
' Recorro el desglose del acum. diario por fecha
Do While Fecha <= hasta
    
    'Por cada desglose para la fecha
    StrSql = "SELECT thnro,achdcanthoras,achdnro FROM gti_achdiario where" & _
    " ternro = " & Ternro & " AND achdfecha = " & ConvFecha(Fecha)
    OpenRecordset StrSql, objRs
    
    Do While Not objRs.EOF
        If Not objRs.EOF Then
            StrSql = "SELECT * FROM gti_achdiario_estr WHERE achdnro = " & objRs!achdnro
            OpenRecordset StrSql, rs
            E1 = rs!estrnro
            te1 = rs!tenro
            rs.MoveNext
            E2 = rs!estrnro
            te2 = rs!tenro
            rs.MoveNext
            E3 = rs!estrnro
            te3 = rs!tenro
            rs.Close
        End If
        
    ' Busco en la tabla de desgloce de acumulado parcial uno para el empleado
        StrSql = "SELECT achpnro,achpcanthoras FROM gti_achparcial" & _
        " WHERE cgtinro = " & Nro_Cab & " AND thnro = " & objRs!thnro & _
        " AND EXISTS (SELECT achpnro FROM gti_achparc_estr WHERE " & _
        " gti_achparc_estr.achpnro = gti_achparcial.achpnro AND" & _
        " estrnro = " & E1 & ")" & _
        " AND EXISTS (SELECT achpnro FROM gti_achparc_estr WHERE " & _
        " gti_achparc_estr.achpnro = gti_achparcial.achpnro AND" & _
        " estrnro = " & E2 & ")" & _
        " AND EXISTS (SELECT achpnro FROM gti_achparc_estr WHERE " & _
        " gti_achparc_estr.achpnro = gti_achparcial.achpnro AND" & _
        " estrnro = " & E3 & ")"
                
        OpenRecordset StrSql, rs
        
        'Si no existe creo el registro en el desglose del acumulado parcial
        'y uno por cada estructura en el desgloce de
        'acumulado por estructuras
        If rs.EOF Then
            StrSql = "INSERT INTO gti_achparcial(achpcanthoras,cgtinro,thnro)" & _
            " VALUES(" & objRs!achdcanthoras & "," & Nro_Cab & "," & _
            objRs!thnro & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            StrSql = "select achpnro as next_id from gti_achparcial " & _
            " order by achpnro desc"
            OpenRecordset StrSql, rs
            l_achpnro = rs("next_id")
            
            StrSql = "INSERT INTO gti_achparc_estr(achpnro,tenro,estrnro)" & _
            " VALUES(" & l_achpnro & "," & te1 & "," & E1 & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
            StrSql = "INSERT INTO gti_achparc_estr(achpnro,tenro,estrnro)" & _
            " VALUES(" & l_achpnro & "," & te2 & "," & E2 & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
            StrSql = "INSERT INTO gti_achparc_estr(achpnro,tenro,estrnro)" & _
            " VALUES(" & l_achpnro & "," & te3 & "," & E3 & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
            
        Else
        'Si existe le sumo
            StrSql = "UPDATE gti_achparcial SET achpcanthoras = " & _
            rs!achpcanthoras + objRs!achdcanthoras & _
            " WHERE cgtinro = " & Nro_Cab & _
            " AND achpnro = " & rs!achpnro
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        
        ' Hay
        
        
        objRs.MoveNext
        
    Loop
    
    Fecha = DateAdd("d", 1, Fecha)

Loop


End Sub

Public Sub SumarHoras(Nro_Cab As Long, desde As Date, hasta As Date, Nro_tpr As Long, Nro_Pro As Long)
' Suma las horas para el reporte trimestral
' Escribe en gti_det

Dim Sumahoras As Single
Dim Ternro As Long
Dim Legajo As Long
Dim HorasTrimestre As Single
Dim DiasHabiles As Single
Dim MinHorasATrabajar As Single
Dim HorasSuma As Single
Dim HorasResta As Single
Dim acu As Single
Dim paga As Single
Dim PrimerTurno As Integer
Dim Nombre As String
      
      
    StrSql = "delete from gti_det where cgtinro = " & Nro_Cab
    objConn.Execute StrSql, , adExecuteNoRecords
     
    StrSql = "select ternro from gti_cab where cgtinro = " & Nro_Cab
    OpenRecordset StrSql, objRs
    Ternro = objRs!Ternro
    
    StrSql = "select empleg, terape, ternom from empleado where ternro = " & Ternro
    OpenRecordset StrSql, objRs
    Legajo = objRs!empleg
    Nombre = objRs!terape & " " & objRs!ternom
    
    StrSql = "select turnro from gti_config_ctacte"
    OpenRecordset StrSql, objRs
    PrimerTurno = objRs!turnro
       
    StrSql = "SELECT ccsigno, gti_acumdiario.thnro, gti_acumdiario.ternro,gti_cab.cgtinro, SUM(adcanthoras) as Sumahoras FROM gti_cab "
    StrSql = StrSql & " INNER JOIN gti_acumdiario ON gti_acumdiario.ternro = gti_cab.ternro "
    StrSql = StrSql & " INNER JOIN gti_tpro_th ON gti_tpro_th.thnro = gti_acumdiario.thnro "
    StrSql = StrSql & " INNER JOIN gti_config_ctacte ON gti_config_ctacte.thnro = gti_acumdiario.thnro "
    StrSql = StrSql & " AND gti_config_ctacte.turnro = " & PrimerTurno
    StrSql = StrSql & " WHERE (gti_cab.cgtinro =" & Nro_Cab & ") and (gti_cab.gpanro = " & Nro_Pro & ") "
    StrSql = StrSql & " AND ( " & ConvFecha(desde) & " <= gti_acumdiario.adfecha  AND gti_acumdiario.adfecha <= " & ConvFecha(hasta) & ") AND "
    StrSql = StrSql & "(gti_tpro_th.gtprocnro = " & Nro_tpr & ")"
    StrSql = StrSql & " GROUP BY  gti_acumdiario.thnro,ccsigno,gti_acumdiario.ternro, gti_cab.cgtinro "
    StrSql = StrSql & " ORDER BY  gti_acumdiario.thnro,ccsigno,gti_acumdiario.ternro, gti_cab.cgtinro "
    OpenRecordset StrSql, objRs
    
    'Totalizo las horas trabajadas de interés en el trimestre que suman
    HorasSuma = 0
    Do While Not objRs.EOF
        
        If objRs!ccsigno Then
            HorasSuma = HorasSuma + objRs!Sumahoras
        Else
            HorasSuma = HorasSuma - objRs!Sumahoras
        End If
        objRs.MoveNext
    Loop
   
    ' punto 1 y 2, Legajo, APYNOM
    StrSql = "select empleg, terape, ternom from empleado where ternro = " & Ternro
    OpenRecordset StrSql, objRs
    Legajo = objRs!empleg
    Nombre = objRs!terape & " " & objRs!ternom
        
    ' punto 3. Horas del Trimestre
        
    DiasHabiles = ContarDiasHabiles(InicioTrimestre, FinTrimestre)
    HorasTrimestre = (HorasSemestre / 2) * (DiasHabiles - _
    ContarDiasProporcionalesDelSemestre(Ternro, InicioTrimestre, FinTrimestre)) / _
    DiasHabiles
    Call InsertarEnGtiDet(Nro_Cab, THorasTrimestre, HorasTrimestre)
        
    ' punto 4, Minimo de Horas extras
    'Esto queda en el confrep
    
    ' punto 5, Minimo de Horas a trabajar
    MinHorasATrabajar = HorasTrimestre + MinHorasExtras
    Call InsertarEnGtiDet(Nro_Cab, THorasMinimo, MinHorasATrabajar)
        
    ' Punto 6, Horas saldo. Viene en recordset. Es la suma de horas.
    Call InsertarEnGtiDet(Nro_Cab, THorasSaldo, HorasSuma)
    
    ' punto 7, Acumula
    If HorasSuma <= MinHorasATrabajar Then
        acu = HorasSuma
    Else
        acu = HorasSuma - MinHorasExtras
    End If
    Call InsertarEnGtiDet(Nro_Cab, THorasAcumula, acu)
    
    ' punto 8, PAGA
    If HorasSuma <= MinHorasATrabajar Then
        paga = 0
    Else
        paga = MinHorasExtras
    End If
    Call InsertarEnGtiDet(Nro_Cab, THorasPaga, paga)
       
       
End Sub

Private Sub InsertarEnGtiDet(Cabecera, TipoHora, Cantidad)
        StrSql = "INSERT INTO gti_det(cgtinro,thnro,dgticant) VALUES (" & _
                 Cabecera & "," & TipoHora & "," & Cantidad & ")"
        objConn.Execute StrSql, , adExecuteNoRecords

End Sub

Private Function ContarDiasHabiles(ByVal FechaInicio As Date, ByVal FechaFin As Date) As Integer
Dim FechaActual As Date
Dim CantidadDias As Integer

CantidadDias = 0
FechaActual = FechaInicio
Do While FechaActual <= FechaFin
    If Not (Weekday(FechaActual) = 1 Or Weekday(FechaActual) = 7) Then
        ' ni Sabado, ni Domingo
        CantidadDias = CantidadDias + 1
    End If

    FechaActual = FechaActual + 1
Loop

ContarDiasHabiles = CantidadDias
End Function

Private Function ContarDiasProporcionalesDelSemestre(ByVal Aux_ternro As Long, ByVal FechaInicio As Date, ByVal FechaFin As Date) As Integer
' devuelve la cantidad de dias NO HABILES del empleado entre las fechas pasadas

Dim FechaAlta As Date
Dim DiasNoHabiles As Integer
Dim rs As New ADODB.Recordset

DiasNoHabiles = 0

StrSql = "SELECT empfaltagr FROM empleado WHERE ternro = " & Aux_ternro
OpenRecordset StrSql, rs

If Not rs.EOF Then
    If IsNull(rs!empfaltagr) Then
        FechaAlta = FechaInicio
    Else
        FechaAlta = rs!empfaltagr
        If FechaAlta >= FechaInicio And FechaAlta <= FechaFin Then
            ' entro en el trimestre, tengo que calcular la cantidad de dias que le corresponden
            DiasNoHabiles = DateDiff("D", FechaInicio, FechaAlta)
        End If
    End If
End If

ContarDiasProporcionalesDelSemestre = DiasNoHabiles
End Function

Private Function RecuperarTipoHoras(ByVal Aux_thnro) As Integer
' recupera el codigo del tipo de Hora definido en el
' mapeo entre Horas de Usuario y Horas del sistema
Dim rs As New ADODB.Recordset
Dim TipoHora As Integer

TipoHora = 0
StrSql = "SELECT thnro FROM Gti_config_turhor WHERE conhornro = " & Aux_thnro
OpenRecordset StrSql, rs

If Not rs.EOF Then
    TipoHora = rs!thnro
End If

RecuperarTipoHoras = TipoHora
End Function


Private Sub ObtenerConfiguracionTrimestral()

' Obtiene la configuracion para el reporte trimestral
' del confrep

    StrSql = "SELECT confval FROM confrep WHERE repnro = 55 AND confnrocol = 1"
    OpenRecordset StrSql, objRs
    HorasSemestre = objRs!confval

    StrSql = "SELECT confval FROM confrep WHERE repnro = 55 AND confnrocol = 2"
    OpenRecordset StrSql, objRs
    MinHorasExtras = objRs!confval
    
    StrSql = "SELECT confval FROM confrep WHERE repnro = 55 AND confnrocol = 3"
    OpenRecordset StrSql, objRs
    THorasTrimestre = objRs!confval
    
    StrSql = "SELECT confval FROM confrep WHERE repnro = 55 AND confnrocol = 4"
    OpenRecordset StrSql, objRs
    THorasMinimo = objRs!confval
    
    StrSql = "SELECT confval FROM confrep WHERE repnro = 55 AND confnrocol = 5"
    OpenRecordset StrSql, objRs
    THorasSaldo = objRs!confval
    
    StrSql = "SELECT confval FROM confrep WHERE repnro = 55 AND confnrocol = 6"
    OpenRecordset StrSql, objRs
    THorasAcumula = objRs!confval
    
    StrSql = "SELECT confval FROM confrep WHERE repnro = 55 AND confnrocol = 7"
    OpenRecordset StrSql, objRs
    THorasPaga = objRs!confval
    
End Sub

