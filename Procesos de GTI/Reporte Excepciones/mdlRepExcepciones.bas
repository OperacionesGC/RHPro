Attribute VB_Name = "mdlRepExcHorNor"
Option Explicit

'Const Version = "1.00"
'Const FechaVersion = "29/08/2012"
'Modificaciones: Lisandro Moro
'Version Inicial

'Const Version = "1.01"
'Const FechaVersion = "06/09/2012"
'Const Modificaciones = "Lisandro Moro - Se actualizan las fechas de los archivos y ya que estamos se mejoro el progreso del batch y se agrego esto al log"

'Const Version = "1.02"
'Const FechaVersion = "11/09/2012"
'Const Modificaciones = "Lisandro Moro - Realizaron varios cambios del orden de los reportaa y se agrego el usuario."

Const Version = "1.03"
Const FechaVersion = "14/09/2012"
Const Modificaciones = "Lisandro Moro - Se excluyen los valores en cero."

'---------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------

'Public Type TConfrep
'    Nrocol As Integer
'    Tipo As String
'    Descripcion As String
'    valAlfa As String
'    valNum As String
'End Type

Dim fs
Dim Flog
Dim FDesde As Date
Dim FHasta As Date

Dim sep As String
Dim objBTurno As New BuscarTurno

Global tiene_turno As Boolean 'licho
Global nro_turno As Long
Global Tipo_Turno As Integer
Global Tiene_Justif As Boolean
Global nro_justif As Long
Global justif_turno As Boolean
Global p_turcomp As Boolean
Global nro_grupo As Long
Global Nro_fpgo As Integer
Global Fecha_Inicio As Date
Global P_Asignacion  As Boolean
Global HuboErrores As Boolean
Global Usuario As String
Global FechaDesde As Date
Global FechaHasta As Date
Global confrepTH() As String
Global confrepCO() As String

Global pgtinro As Long
Global gpanro As Long
Global ternrodesde As Long
Global ternrohasta As Long
Global Tenro As Long
Global estrnro As Long
Global titulo As String
Global NroProceso As Long
Global tieneDestalle As Boolean
Global Aux_Fecha As String
Global progreso As Single
Global TiempoInicialProceso
Global TiempoAcumulado

Sub Main()
Dim Archivo As String
Dim pos As Integer
Dim strcmdLine  As String

'Dim objconnMain As New ADODB.Connection
Dim rs As New ADODB.Recordset

Dim Fecha As Date
Dim Hora As String

Dim NroReporte As Long
Dim StrParametros() As String

Dim rs_Batch_Proceso As New ADODB.Recordset
Dim rs_His_Batch_Proceso As New ADODB.Recordset
Dim PID As String
Dim ArrParametros

    strcmdLine = Command()
    ArrParametros = Split(strcmdLine, " ", -1)
    If UBound(ArrParametros) > 1 Then
        If IsNumeric(ArrParametros(0)) Then
            NroProceso = ArrParametros(0)
            Etiqueta = ArrParametros(1)
            EncriptStrconexion = CBool(ArrParametros(2))
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
    Archivo = PathFLog & "RepExcepcionesHorarioNormal-" & CStr(NroProceso) & " - " & Format(Now, "DD-MM-YYYY") & ".log"
    
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

    TiempoInicialProceso = GetTickCount
    TiempoAcumulado = GetTickCount

    'Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaVersion
    Flog.writeline "Modificaciones           : " & Modificaciones
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-------------------------------------------------"
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "Inicio :" & Now
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprcpid = " & PID & ", bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcprogreso = 0 "
    StrSql = StrSql & " , bprctiempo = " & TiempoInicialProceso
    StrSql = StrSql & " WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords

    Flog.writeline Espacios(Tabulador * 0) & "Levanta Proceso y Setea Parámetros:  " & " " & Now
    
    'levanto los parametros del proceso
    'StrParametros = ""
    sep = "@"
    StrSql = "SELECT bprcfecdesde,bprcfechasta,bprcparam,bprcfecha,bprchora,iduser  FROM batch_proceso WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        Fecha = rs!bprcfecha
        Hora = rs!bprchora
        Usuario = rs!iduser
        If Not IsNull(rs!bprcparam) Then
            If Len(rs!bprcparam) >= 1 Then
                StrParametros = Split(rs!bprcparam, sep)
                pgtinro = StrParametros(0)
                gpanro = StrParametros(1)
                ternrodesde = StrParametros(2)
                ternrohasta = StrParametros(3)
                Tenro = StrParametros(4)
                estrnro = StrParametros(5)
                titulo = StrParametros(6)

                'NroReporte = CLng(Left(rs!bprcparam, pos - 1))
                'StrParametros = Right(rs!bprcparam, Len(rs!bprcparam) - (pos))
            End If
        End If
    Else
        Exit Sub
    End If
    
    depurar = True
    
    If depurar Then
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "Inicio de Reporte de Excepciones al Horario Normal: " & " " & Now
    End If
    
    generarReporte
    
    If depurar Then
        Flog.writeline Espacios(Tabulador * 0) & "Fin de Reporte de Excepciones al Horario Normal: " & " " & Now
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "Copio al historico" & " " & Now
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
    If depurar Then
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 1) & "Copio el proceso en el historico de batch_proceso y lo borro de batch_proceso"
        Flog.writeline
    End If
    
    If Not HuboErrores Then
        If depurar Then
            Flog.writeline Espacios(Tabulador * 1) & "---> Proceso teminado, paso al historico ... " & Now
        End If

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
    
        If depurar Then
            Flog.writeline Espacios(Tabulador * 1) & "---> Historico Actualizado " & Now
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

Private Sub generarReporte()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que calcula las Excepciones al Horario Normal.
' Autor      : Lisandro Moro
' Fecha      : 29/08/2012
' Ultima Mod.: Null
' ---------------------------------------------------------------------------------------------
Dim pos1 As Integer
Dim pos2 As Integer
Dim i As Integer


'*- Cabecera -*
'Dim bpronro As Long
'Dim pgtinro As Long
'Dim gpanro As Long
'Dim legajo As Long
Dim Ternro As Long
Dim terape As String
Dim terape2 As String
Dim ternom As String
Dim ternom2 As String
'Dim tenro As Long
Dim tedesc As String
'Dim estrnro As Long
Dim estrdabr As String
Dim emprep1 As Long
Dim emprepdesc1 As String
Dim emprep2 As Long
Dim emprepdesc2 As String
Dim emprep3 As Long
Dim emprepdesc3 As String
Dim repdesc As String

'*- Detalle -*
'Dim bpronro As Long
'Dim ternro
Dim novtip As String 'Tipo de Novedad (LIC,NH,NOV)
Dim novnro As Long '/* Codigo INTERNO Novedad */
Dim novdesc As String '/* Descripcion de la Novedad */
Dim novdesde As Date '/* Fecha Desde */
Dim novhasta As Date '/* Fecha Hasta */
Dim novcant As Double  '/* HS base 19,4 */
    

'Dim l_nrocol As Long
'Dim l_tipo As String
'Dim l_val1 As Long
'Dim l_val2 As String
'Dim l_accion As String


'Dim l_val As Double

Dim CEmpleadosAProc As Integer
Dim CDiasAProc As Integer
Dim IncPorc As Single
'Dim progreso As Single

Dim rs As New ADODB.Recordset
'Dim rs2 As New ADODB.Recordset
Dim rsEmpleados As New ADODB.Recordset
'Dim rs3 As New ADODB.Recordset
'Dim rs_Doc As New ADODB.Recordset


On Error GoTo ME_Local

''Busco la fecha hasta del periodo de gti
'StrSql = "SELECT pgtidesde, pgtihasta FROM gti_per WHERE pgtinro = " & pgtinro
'OpenRecordset StrSql, rs
'If Not rs.EOF Then
'    'Aux_Fecha = rs!pgtihasta
'    'FechaDesde = rs!pgtidesde
'    'FechaHasta = rs!pgtihasta
'End If

'Busco la fecha hasta del Proceso de gti
StrSql = "SELECT gpanro, gpadesabr, pgtinro, gpadesde, gpahasta, gtprocnro "
StrSql = StrSql & " FROM gti_procacum "
StrSql = StrSql & " WHERE gpanro = " & gpanro
OpenRecordset StrSql, rs
If Not rs.EOF Then
    Aux_Fecha = rs!gpahasta
    FechaDesde = rs!gpadesde
    FechaHasta = rs!gpahasta
End If


'Busco las descripciones de las estructuras
StrSql = "SELECT tenro, tedabr FROM tipoestructura WHERE tenro = " & Tenro
OpenRecordset StrSql, rs
If Not rs.EOF Then
    tedesc = rs!tedabr
Else
    tedesc = ""
End If

' CONFREP
' Tipo Horas
StrSql = "SELECT * FROM confrep WHERE repnro = 382 AND conftipo = 'TH' "
OpenRecordset StrSql, rs
If Not rs.EOF Then
    'confrepTH = rs.GetRows
Else
    Flog.writeline Espacios(Tabulador * 0) & "ERROR: No existen TH configurados en el confrep"
End If

' Tipo Horas
StrSql = "SELECT * FROM confrep WHERE repnro = 382 AND conftipo = 'CO' "
OpenRecordset StrSql, rs
If Not rs.EOF Then
    'confrepCO = rs.GetRows
Else
    Flog.writeline Espacios(Tabulador * 0) & "ERROR: No existen CO configurados en el confrep"
End If


'confrepCo()

StrSql = " SELECT empleado.ternro, empleado.terape, empleado.terape2, empleado.ternom, empleado.ternom2, empleado.empleg, empleado.empreporta "
StrSql = StrSql & " ,estructura.estrnro, estructura.estrdabr, tipoestructura.tenro, tipoestructura.tedabr "
'StrSql = StrSql & " , jefe.empleg "
StrSql = StrSql & " From Empleado "
StrSql = StrSql & " INNER JOIN empleado jefe ON empleado.empreporta = jefe.ternro "
StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.ternro = empleado.ternro "
StrSql = StrSql & " INNER JOIN tipoestructura ON his_estructura.tenro = tipoestructura.tenro "
StrSql = StrSql & " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro "
StrSql = StrSql & " WHERE empleado.empreporta IN ( "
StrSql = StrSql & "    SELECT r.ternro FROM empleado r WHERE r.empleg>= (select d.empleg from empleado d where d.ternro = " & ternrodesde & ")"
StrSql = StrSql & "    AND empleg <= (select h.empleg from empleado h where h.ternro = " & ternrohasta & "))"
StrSql = StrSql & " AND ((his_estructura.htetdesde <= " & ConvFecha(Aux_Fecha) & ")"
StrSql = StrSql & " AND (" & ConvFecha(Aux_Fecha) & " <= his_estructura.htethasta OR his_estructura.htethasta is null))"

'StrSql = " SELECT empleado.ternro, terape, terape2, ternom, ternom2, empleg, empreporta "
'StrSql = StrSql & " ,estructura.estrnro, estructura.estrdabr, tipoestructura.tenro, tipoestructura.tedabr "
'StrSql = StrSql & " FROM empleado "
'StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.ternro = empleado.ternro "
'StrSql = StrSql & " INNER JOIN tipoestructura ON his_estructura.tenro = tipoestructura.tenro "
'StrSql = StrSql & " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro "
'StrSql = StrSql & " WHERE empreporta IN ("
'StrSql = StrSql & "      SELECT r.ternro FROM empleado r WHERE r.empleg>= (select d.empleg from empleado d where d.ternro = " & ternrodesde & ")"
'StrSql = StrSql & "      AND empleg <= (select h.empleg from empleado h where h.ternro = " & ternrohasta & ") "
'StrSql = StrSql & "      )"
'StrSql = StrSql & " AND ((his_estructura.htetdesde <= " & ConvFecha(Aux_Fecha) & ")"
'StrSql = StrSql & " AND (" & ConvFecha(Aux_Fecha) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
If CStr(estrnro) <> "0" Then
    StrSql = StrSql & " AND his_estructura.estrnro = " & estrnro
Else
    StrSql = StrSql & " AND his_estructura.tenro = " & Tenro
End If
StrSql = StrSql & " ORDER BY jefe.empleg " 'empreporta
'Flog.writeline StrSql
OpenRecordset StrSql, rsEmpleados
If Not rsEmpleados.EOF Then
    progreso = (100 / rsEmpleados.RecordCount)
    Do While Not rsEmpleados.EOF
        Call generarCabecera(rsEmpleados)
        rsEmpleados.MoveNext
    Loop
Else
    'No hay datos
    
End If


Fin:
'Cierro y libero
If rs.State = adStateOpen Then rs.Close
Set rs = Nothing
'If rs2.State = adStateOpen Then rs2.Close
'Set rs2 = Nothing
Exit Sub

ME_Local:
    HuboErrores = True
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 1) & " Error " & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline
    GoTo Fin
End Sub

Private Sub generarCabecera(rsEmpleados As Recordset)

Dim empter1 As Long
Dim empter2 As Long
Dim empter3 As Long

Dim empresa As String
Dim emprep1 As String
Dim emprepdesc1 As String
Dim empreppuesto1 As String
Dim emprepfirma1 As String
Dim emprep2 As String
Dim emprepdesc2 As String
Dim empreppuesto2 As String
Dim emprepfirma2 As String
Dim emprep3 As String
Dim emprepdesc3 As String
Dim empreppuesto3 As String
Dim emprepfirma3 As String
Dim rs2 As New ADODB.Recordset

On Error GoTo error_genera

tieneDestalle = False

'Busco los reporta a del reorta a del reporta a
emprep1 = 0
emprepdesc1 = ""
emprepfirma1 = ""
empreppuesto1 = ""
emprep2 = 0
emprepdesc2 = ""
emprepfirma1 = ""
empreppuesto1 = ""
emprep3 = 0
emprepdesc3 = ""
emprepfirma3 = ""
empreppuesto3 = ""

StrSql = " SELECT empleado.ternro, empleg, terape, terape2, ternom, ternom2, empreporta "
StrSql = StrSql & " FROM empleado "
StrSql = StrSql & " INNER JOIN user_ter ON empleado.ternro = user_ter.ternro "
StrSql = StrSql & " WHERE iduser = '" & Usuario & "'"
OpenRecordset StrSql, rs2
If Not rs2.EOF Then
    emprep1 = n2Long(rs2!empleg)
    emprepdesc1 = Left(Trim(n2String(rs2!terape) & " " & n2String(rs2!terape2)) & ", " & Trim(n2String(rs2!ternom) & " " & n2String(rs2!ternom2)), 100)
    emprepfirma1 = getFirma(rs2!Ternro, 2)
    empreppuesto1 = getEstrdabr(rs2!Ternro, 5, Aux_Fecha) 'CENTRO DE COSTOS
    empresa = getEstrdabr(rs2!Ternro, 10, Aux_Fecha)
End If

' EmpReporta N 1
If Not IsNull(rsEmpleados!empreporta) Then
    StrSql = " SELECT ternro, empleg, terape, terape2, ternom, ternom2, empreporta "
    StrSql = StrSql & " FROM empleado "
    StrSql = StrSql & " WHERE ternro = " & rsEmpleados!empreporta
    OpenRecordset StrSql, rs2
    If Not rs2.EOF Then
        emprep2 = n2Long(rs2!empleg)
        emprepdesc2 = Left(Trim(n2String(rs2!terape) & " " & n2String(rs2!terape2)) & ", " & Trim(n2String(rs2!ternom) & " " & n2String(rs2!ternom2)), 100)
        emprepfirma2 = getFirma(rs2!Ternro, 2)
        empreppuesto2 = getEstrdabr(rs2!Ternro, 5, Aux_Fecha) 'CENTRO DE COSTOS
        'empresa = getEstrdabr(rs2!Ternro, 10, Aux_Fecha)
        ' EmpReporta N 2
        If Not IsNull(rs2!empreporta) Then
            StrSql = " SELECT ternro, empleg, terape, terape2, ternom, ternom2, empreporta "
            StrSql = StrSql & " FROM empleado "
            StrSql = StrSql & " WHERE ternro = " & n2String(rs2!empreporta)
            OpenRecordset StrSql, rs2
            If Not rs2.EOF Then
                emprep3 = n2Long(rs2!empleg)
                emprepdesc3 = Left(Trim(n2String(rs2!terape) & " " & n2String(rs2!terape2)) & ", " & Trim(n2String(rs2!ternom) & " " & n2String(rs2!ternom2)), 100)
                emprepfirma3 = getFirma(rs2!Ternro, 2)
                empreppuesto3 = getEstrdabr(rs2!Ternro, 4, Aux_Fecha)
                ' EmpReporta N 3
                'If Not IsNull(rs2!empreporta) Then
                '    StrSql = " SELECT ternro, empleg, terape, terape2, ternom, ternom2, empreporta "
                '    StrSql = StrSql & " FROM empleado "
                '    StrSql = StrSql & " WHERE ternro = " & n2String(rs2!empreporta)
                '    OpenRecordset StrSql, rs2
                '    If Not rs2.EOF Then
                '        emprep3 = n2Long(rs2!empleg)
                '        emprepdesc3 = Left(Trim(n2String(rs2!terape) & " " & n2String(rs2!terape2)) & ", " & Trim(n2String(rs2!ternom) & " " & n2String(rs2!ternom2)), 100)
                '        emprepfirma3 = getFirma(rs2!Ternro, 2)
                '        empreppuesto3 = getEstrdabr(rs2!Ternro, 4, Aux_Fecha)
                '    End If
                'End If
            End If
        End If
    End If
End If


'' EmpReporta N 1
'If Not IsNull(rsEmpleados!empreporta) Then
'    StrSql = " SELECT ternro, empleg, terape, terape2, ternom, ternom2, empreporta "
'    StrSql = StrSql & " FROM empleado "
'    StrSql = StrSql & " WHERE ternro = " & rsEmpleados!empreporta
'    OpenRecordset StrSql, rs2
'    If Not rs2.EOF Then
'        emprep1 = n2Long(rs2!empleg)
'        emprepdesc1 = Left(Trim(n2String(rs2!terape) & " " & n2String(rs2!terape2)) & ", " & Trim(n2String(rs2!ternom) & " " & n2String(rs2!ternom2)), 100)
'        emprepfirma1 = getFirma(rs2!Ternro, 2)
'        empreppuesto1 = getEstrdabr(rs2!Ternro, 5, Aux_Fecha) 'CENTRO DE COSTOS
'        empresa = getEstrdabr(rs2!Ternro, 10, Aux_Fecha)
'        ' EmpReporta N 2
'        If Not IsNull(rs2!empreporta) Then
'            StrSql = " SELECT ternro, empleg, terape, terape2, ternom, ternom2, empreporta "
'            StrSql = StrSql & " FROM empleado "
'            StrSql = StrSql & " WHERE ternro = " & n2String(rs2!empreporta)
'            OpenRecordset StrSql, rs2
'            If Not rs2.EOF Then
'                emprep2 = n2Long(rs2!empleg)
'                emprepdesc2 = Left(Trim(n2String(rs2!terape) & " " & n2String(rs2!terape2)) & ", " & Trim(n2String(rs2!ternom) & " " & n2String(rs2!ternom2)), 100)
'                emprepfirma2 = getFirma(rs2!Ternro, 2)
'                empreppuesto2 = getEstrdabr(rs2!Ternro, 4, Aux_Fecha)
'                ' EmpReporta N 3
'                If Not IsNull(rs2!empreporta) Then
'                    StrSql = " SELECT ternro, empleg, terape, terape2, ternom, ternom2, empreporta "
'                    StrSql = StrSql & " FROM empleado "
'                    StrSql = StrSql & " WHERE ternro = " & n2String(rs2!empreporta)
'                    OpenRecordset StrSql, rs2
'                    If Not rs2.EOF Then
'                        emprep3 = n2Long(rs2!empleg)
'                        emprepdesc3 = Left(Trim(n2String(rs2!terape) & " " & n2String(rs2!terape2)) & ", " & Trim(n2String(rs2!ternom) & " " & n2String(rs2!ternom2)), 100)
'                        emprepfirma3 = getFirma(rs2!Ternro, 2)
'                        empreppuesto3 = getEstrdabr(rs2!Ternro, 4, Aux_Fecha)
'                    End If
'                End If
'            End If
'        End If
'    End If
'End If

Flog.writeline Espacios(Tabulador * 0) & "Generando Emplado : " & rsEmpleados!empleg & " - " & n2String(rsEmpleados!terape) & " " & n2String(rsEmpleados!ternom)

tieneDestalle = False
generarDetalle rsEmpleados

If tieneDestalle Then
    StrSql = " INSERT INTO rep_exc_hor_nor ( "
    StrSql = StrSql & " bpronro , pgtinro, gpanro, gpanroFecha, legajo, ternro, terape, terape2, ternom, ternom2 "
    StrSql = StrSql & " ,tenro,tedesc,estrnro,estrdabr,emprep1,emprepdesc1,emprepfirma1,empreppuesto1,emprep2,emprepdesc2,emprepfirma2,empreppuesto2,emprep3,emprepdesc3,emprepfirma3,empreppuesto3"
    StrSql = StrSql & " ,repdesc, repfec, empresa "
    StrSql = StrSql & ") VALUES ("
    StrSql = StrSql & NroProceso
    StrSql = StrSql & "," & pgtinro
    StrSql = StrSql & "," & gpanro
    StrSql = StrSql & ",'" & FechaDesde & " - " & FechaHasta & "'"
    StrSql = StrSql & "," & rsEmpleados!empleg
    StrSql = StrSql & "," & rsEmpleados!Ternro
    StrSql = StrSql & ",'" & n2String(rsEmpleados!terape) & "'"
    StrSql = StrSql & ",'" & n2String(rsEmpleados!terape2) & "'"
    StrSql = StrSql & ",'" & n2String(rsEmpleados!ternom) & "'"
    StrSql = StrSql & ",'" & n2String(rsEmpleados!ternom2) & "'"
    StrSql = StrSql & "," & rsEmpleados!Tenro
    StrSql = StrSql & ",'" & rsEmpleados!tedabr & "'"
    StrSql = StrSql & "," & rsEmpleados!estrnro
    StrSql = StrSql & ",'" & rsEmpleados!estrdabr & "'"
    If emprep1 = 0 Then
        StrSql = StrSql & ",NULL"
    Else
        StrSql = StrSql & "," & emprep1
    End If
    StrSql = StrSql & ",'" & emprepdesc1 & "'"
    StrSql = StrSql & ",'" & emprepfirma1 & "'"
    StrSql = StrSql & ",'" & empreppuesto1 & "'"
    If emprep2 = 0 Then
        StrSql = StrSql & ",NULL"
    Else
        StrSql = StrSql & "," & emprep2
    End If
    StrSql = StrSql & ",'" & emprepdesc2 & "'"
    StrSql = StrSql & ",'" & emprepfirma2 & "'"
    StrSql = StrSql & ",'" & empreppuesto2 & "'"
    If emprep1 = 0 Then
        StrSql = StrSql & ",NULL"
    Else
        StrSql = StrSql & "," & emprep3
    End If
    StrSql = StrSql & ",'" & emprepdesc3 & "'"
    StrSql = StrSql & ",'" & emprepfirma3 & "'"
    StrSql = StrSql & ",'" & empreppuesto3 & "'"
    StrSql = StrSql & ",'" & Left(titulo, 200) & "'"
    StrSql = StrSql & ",'" & Now & "'"
    StrSql = StrSql & ",'" & empresa & "'"
    StrSql = StrSql & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
Else
    Flog.writeline Espacios(Tabulador * 2) & "No hay detalle. No se genera el Empleado."
End If

'Progreso
TiempoAcumulado = GetTickCount
StrSql = "UPDATE batch_proceso SET bprcprogreso = bprcprogreso + " & progreso
StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'"
StrSql = StrSql & ", bprcestado = 'Procesando' "
StrSql = StrSql & " WHERE bpronro = " & NroProceso
'StrSql = "UPDATE batch_proceso SET bprcestado = 'Procesando', bprcprogreso=bprcprogreso+ " & progreso & " WHERE bpronro = " & NroProceso
objConn.Execute StrSql, , adExecuteNoRecords


Exit Sub
error_genera:
   HuboErrores = True
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 1) & " Error " & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline
    'GoTo Fin

End Sub

Private Sub generarDetalle(rsEmpleados As Recordset)
Dim rs2 As New ADODB.Recordset

Flog.writeline Espacios(Tabulador * 1) & "Buscando Licencias y Novedades de GTI desde Acumulados Diarios."
'Busco Licencias y Novedades de GTI desde Acumulados Diarios.
'StrSql = " SELECT adfecha, adcanthoras, admanual, adestado, advalido, tiphora.thdesc, tiphora.thnro, horas " 'full cuac
StrSql = " SELECT adfecha, adcanthoras, tiphora.thdesc, tiphora.thnro, horas, confetiq "
StrSql = StrSql & " FROM gti_acumdiario "
StrSql = StrSql & " INNER JOIN tiphora ON gti_acumdiario.thnro=tiphora.thnro "
StrSql = StrSql & " INNER JOIN confrep ON confrep.confval = gti_acumdiario.thnro AND repnro = 382 "
StrSql = StrSql & " WHERE ternro = " & rsEmpleados!Ternro
StrSql = StrSql & " AND adfecha >= " & ConvFecha(FechaDesde)
StrSql = StrSql & " AND adfecha <= " & ConvFecha(FechaHasta)
StrSql = StrSql & " AND adcanthoras <> 0 "
OpenRecordset StrSql, rs2
If Not rs2.EOF Then
    tieneDestalle = True
    Do While Not rs2.EOF
        insertarDetalle rsEmpleados!Ternro, "TH", rs2!thnro, rs2!confetiq, rs2!adfecha, "", rs2!adcanthoras
        rs2.MoveNext
    Loop
Else
    'Flog.writeline Espacios(Tabulador * 0) & "No se encontro detalle Acumulado Diario."
End If
    
Flog.writeline Espacios(Tabulador * 1) & "Buscando Novedades del Empleado."
'Busco Novedades del Empleado.
StrSql = " SELECT novemp.concnro, tpanro, empleado, nevalor, nevigencia, nedesde, nehasta, confetiq "
StrSql = StrSql & " FROM novemp "
StrSql = StrSql & " INNER JOIN concepto ON novemp.concnro = concepto.concnro "
StrSql = StrSql & " INNER JOIN confrep ON concepto.conccod = confrep.confval2 AND confrep.confval = novemp.tpanro AND repnro = 382 "
StrSql = StrSql & " WHERE empleado = " & rsEmpleados!Ternro
StrSql = StrSql & " AND ("
StrSql = StrSql & "     (nevigencia = -1 AND (nedesde >= " & ConvFecha(FechaDesde) & " AND nedesde <= " & ConvFecha(FechaDesde) & "))"
StrSql = StrSql & "     OR (nevigencia = 0)"
StrSql = StrSql & " )"
StrSql = StrSql & " AND nevalor <> 0 "
OpenRecordset StrSql, rs2
If Not rs2.EOF Then
    tieneDestalle = True
    Do While Not rs2.EOF
        insertarDetalle rsEmpleados!Ternro, "CO", rs2!concnro, rs2!confetiq, n2String(rs2!nedesde), n2String(rs2!nehasta), rs2!nevalor
        rs2.MoveNext
    Loop
Else
    'Flog.writeline Espacios(Tabulador * 0) & "No se encontro detalle Acumulado Diario."
End If
    
    
End Sub

Function insertarDetalle(Ternro As String, novtip As String, novnro As String, novdesc As String, novdesde As String, novhasta As String, novcant As String)

StrSql = " INSERT INTO rep_exc_hor_nor_det ("
StrSql = StrSql & " bpronro , ternro, novtip, novnro, novdesc, novdesde, novhasta, novcant "
StrSql = StrSql & ") VALUES ("
StrSql = StrSql & NroProceso
StrSql = StrSql & "," & Ternro
StrSql = StrSql & ",'" & novtip & "'"
If novnro = "" Then
    StrSql = StrSql & ",NULL"
Else
    StrSql = StrSql & "," & novnro
End If
StrSql = StrSql & ",'" & novdesc & "'"
If novdesde = "" Then
    StrSql = StrSql & ",NULL"
Else
    StrSql = StrSql & "," & ConvFecha(novdesde)
End If
If novhasta = "" Then
    StrSql = StrSql & ",NULL"
Else
    StrSql = StrSql & "," & ConvFecha(novhasta)
End If
StrSql = StrSql & "," & IIf(novcant = "", "0", novcant)
StrSql = StrSql & ")"
objConn.Execute StrSql, , adExecuteNoRecords

Flog.writeline Espacios(Tabulador * 1) & " Inserto DETALLE " & Ternro & " - " & novtip & " - " & novnro & " - " & novdesc & " - " & novdesde & " - " & novhasta & " - " & novcant

End Function

Public Function n2String(Algo As Variant) As String
    If IsNull(Algo) Then
        n2String = ""
    Else
        n2String = CStr(Algo)
    End If
End Function

Public Function n2Long(Algo As Variant) As Long
    If IsNull(Algo) Then
        n2Long = 0
    Else
        If Algo = "" Then
            n2Long = 0
        Else
            n2Long = CLng(Algo)
        End If
    End If
End Function
Function getFirma(Ternro As Long, tipoimagen As Long) As String
    Dim rs2 As New ADODB.Recordset
    StrSql = "SELECT tipimdire, terimnombre"
    StrSql = StrSql & " FROM ter_imag "
    StrSql = StrSql & " INNER JOIN tipoimag ON tipoimag.tipimnro = ter_imag.tipimnro "
    StrSql = StrSql & " WHERE ternro = " & Ternro
    StrSql = StrSql & " AND ter_imag.tipimnro = " & tipoimagen
    OpenRecordset StrSql, rs2
    If Not rs2.EOF Then
        getFirma = n2String(rs2!tipimdire) & n2String(rs2!terimnombre)
    Else
        getFirma = ""
    End If
    Set rs2 = Nothing
End Function

Function getEstrdabr(Ternro As Long, tipoestructura As Long, alafecha As String)
    Dim rs2 As New ADODB.Recordset
    StrSql = "SELECT estrdabr "
    StrSql = StrSql & " FROM empleado "
    StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.ternro = empleado.ternro "
    StrSql = StrSql & " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro "
    StrSql = StrSql & " WHERE empleado.ternro = " & Ternro
    StrSql = StrSql & " AND (his_estructura.htetdesde <= " & ConvFecha(alafecha) & ")"
    StrSql = StrSql & " AND ((" & ConvFecha(alafecha) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
    StrSql = StrSql & " AND his_estructura.tenro = " & tipoestructura
    OpenRecordset StrSql, rs2
    If Not rs2.EOF Then
        getEstrdabr = n2String(rs2!estrdabr)
    Else
        getEstrdabr = ""
    End If
    Set rs2 = Nothing
End Function
