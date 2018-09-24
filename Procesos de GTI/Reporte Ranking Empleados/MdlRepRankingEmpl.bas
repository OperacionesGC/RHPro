Attribute VB_Name = "MdlRepRankingEmpl"
Option Explicit

'Const Version = "1.00"
'Const FechaVersion = "28/08/2008"
'Modificaciones: CS
'    Version Inicial, con varios cambios del esatandar

Const Version = "1.01"
Const FechaVersion = "31/07/2009"
'Modificaciones: Martin Ferraro - Encriptacion de string connection

'---------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------

'Public Type TConfrep
'    Nrocol As Integer
'    Tipo As String
'    val1 As tring
'    val2 As tring
'    Accion As String
'End Type

Dim fs
Dim Flog
Dim FDesde As Date
Dim FHasta As Date

Dim sep As String
Dim objBTurno As New BuscarTurno

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
Global P_Asignacion  As Boolean
Global HuboErrores As Boolean
Global Usuario As String


Sub Main()
Dim Archivo As String
Dim pos As Integer
Dim strcmdLine As String

'Dim objconnMain As New ADODB.Connection
Dim rs As New ADODB.Recordset

Dim Fecha As Date
Dim Hora As String
Dim NroProceso As Long
Dim NroReporte As Long
Dim StrParametros As String

Dim rs_Batch_Proceso As New ADODB.Recordset
Dim rs_His_Batch_Proceso As New ADODB.Recordset
Dim PID As String
Dim ArrParametros

'    strcmdLine = Command()
'    ArrParametros = Split(strcmdLine, " ", -1)
'    If UBound(ArrParametros) > 0 Then
'        If IsNumeric(ArrParametros(0)) Then
'            NroProceso = ArrParametros(0)
'            Etiqueta = ArrParametros(1)
'        Else
'            Exit Sub
'        End If
'    Else
'        If IsNumeric(strcmdLine) Then
'            NroProceso = strcmdLine
'        Else
'            Exit Sub
'        End If
'    End If

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
    'Archivo = PathFLog & "RepRankingEmpl-" & CStr(NroProceso) & Format(Now, "DD-MM-YYYY") & ".log"
    Archivo = PathFLog & "RepRankingEmpl-" & CStr(NroProceso) & ".log"
    
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Archivo, True)

    'Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaVersion
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-------------------------------------------------"
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "Inicio :" & Now
    
    'Abro la conexion
    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    OpenConnection strconexion, objConnProgreso
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If

    On Error GoTo CE

    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprcpid = " & PID & ", bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords

    Flog.writeline Espacios(Tabulador * 0) & "Levanta Proceso y Setea Parámetros:  " & " " & Now
    
    'levanto los parametros del proceso
    StrParametros = ""
    sep = "@"
    StrSql = "SELECT bprcfecdesde,bprcfechasta,bprcparam,bprcfecha,bprchora,iduser  FROM batch_proceso WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        Fecha = rs!bprcfecha
        Hora = rs!bprchora
        Usuario = rs!iduser
        If Not IsNull(rs!bprcparam) Then
            If Len(rs!bprcparam) >= 1 Then
                pos = InStr(1, rs!bprcparam, sep)
                NroReporte = CLng(Left(rs!bprcparam, pos - 1))
                StrParametros = Right(rs!bprcparam, Len(rs!bprcparam) - (pos))
            End If
        End If
    Else
        Exit Sub
    End If
    
    depurar = True
    
    If depurar Then
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "Inicio de Reporte de Novedades: " & " " & Now
    End If
    Call Reporte_01(NroReporte, NroProceso, StrParametros, Fecha, Hora)
    If depurar Then
        Flog.writeline Espacios(Tabulador * 0) & "Fin de Reporte de Novedades: " & " " & Now
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "copio al historico" & " " & Now
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

Private Sub Reporte_01(NroReporte As Long, NroProceso As Long, Parametros As String, Fecha As Date, Hora As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que calcula las novedades.
' Autor      : CS
' Fecha      :
' Ultima Mod.:
' ---------------------------------------------------------------------------------------------
Dim pos1 As Integer
Dim pos2 As Integer
Dim i As Integer

Dim Por_Usuario As Boolean
Dim Aux_Fecha As Date

Dim l_nrocol As Long
Dim l_tipo As String
Dim l_val1 As Long
Dim l_val2 As String
Dim l_accion As String

Dim l_ternro As Long
Dim l_empleg As Long
Dim l_terape As String
Dim l_ternom As String
Dim l_tarjeta As Long
Dim l_tarjeta_ant As Long
Dim l_dife As Long
Dim l_ant_anios As Long
Dim l_ant_dias As Long
Dim l_porc_antig As Double
Dim l_situacion As String
Dim l_observacion As String
Dim l_cerrado As String
Dim l_dias_trabajados As Double
Dim l_val As Double
Dim l_modif As String

'se supone que estos son parametros de entrada y vienen en "parametros"
Dim l_anio As String
Dim l_procesos As String

' en esta variable almaceno los tipos de horas
Dim l_horas As String

Dim CEmpleadosAProc As Integer
Dim CDiasAProc As Integer
Dim IncPorc As Single
Dim Progreso As Single

Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset
Dim rs_Estructura As New ADODB.Recordset
Dim rs_Doc As New ADODB.Recordset

Dim Tope_Dias As Long
Dim Tope_Horas As Long


On Error GoTo ME_Local

If Not IsNull(Parametros) Then
    If Len(Parametros) >= 1 Then
        l_anio = Parametros
    End If
End If

'Inicializaciones
Tope_Dias = 213
Tope_Horas = 8

'Para el empleado actual ciclo entre las columnas configuradas en el rep y guardo los valores en el detalle
'Busco si el reporte se configura por usuario
StrSql = "SELECT repagr FROM reporte WHERE repnro = 237"
OpenRecordset StrSql, rs3
If Not rs3.EOF Then
    If CBool(rs3!repagr) Then
        Por_Usuario = True
    Else
        Por_Usuario = False
    End If
Else
    Por_Usuario = False
End If

'Busco los topes de dias por año configuradas en el confrep
StrSql = " SELECT confnrocol, conftipo, confval "
StrSql = StrSql & " FROM confrep "
StrSql = StrSql & " WHERE repnro = " & NroReporte
StrSql = StrSql & " AND upper(conftipo) = 'DIA'"
If Por_Usuario Then
    StrSql = StrSql & " AND upper(iduser) = '" & UCase(Usuario) & "'"
Else
    StrSql = StrSql & " AND (iduser = '' OR iduser IS NULL )"
End If
OpenRecordset StrSql, rs3
If Not rs3.EOF Then
    If rs3!confval <> 0 Then
        Tope_Dias = rs3!confval
    End If
End If
'Busco los topes de horas por dia configuradas en el confrep
StrSql = " SELECT confnrocol, conftipo, confval "
StrSql = StrSql & " FROM confrep "
StrSql = StrSql & " WHERE repnro = " & NroReporte
StrSql = StrSql & " AND upper(conftipo) = 'HOR'"
If Por_Usuario Then
    StrSql = StrSql & " AND upper(iduser) = '" & UCase(Usuario) & "'"
Else
    StrSql = StrSql & " AND (iduser = '' OR iduser IS NULL )"
End If
OpenRecordset StrSql, rs3
If Not rs3.EOF Then
    If rs3!confval <> 0 Then
        Tope_Horas = rs3!confval
    End If
End If

'Busco los tipos de horas configuradas en el confrep
StrSql = " SELECT confnrocol, conftipo, confval "
StrSql = StrSql & " FROM confrep "
StrSql = StrSql & " WHERE repnro = " & NroReporte
StrSql = StrSql & " AND conftipo = 'TH'"
If Por_Usuario Then
    StrSql = StrSql & " AND iduser = '" & Usuario & "'"
Else
    StrSql = StrSql & " AND (iduser = '' OR iduser IS NULL )"
End If
OpenRecordset StrSql, rs3
l_horas = 0
Do While Not rs3.EOF
    If rs3!confval <> 0 Then
        l_horas = l_horas & "," & rs3!confval
    End If
    rs3.MoveNext
Loop

If l_horas = 0 Then
    If depurar Then
        Flog.writeline Espacios(Tabulador * 1) & "No hay tipos de horas configuradas en la configuracion del reporte de ranking de empleados"
    End If
End If


'Levanto todos los empleados a Procesar
StrSql = "SELECT batch_empleado.ternro, empleado.empleg, tercero.terape, tercero.ternom, tercero.terape2, tercero.ternom2 "
StrSql = StrSql & " FROM batch_empleado "
StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = batch_empleado.ternro "
StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro = empleado.ternro "
StrSql = StrSql & " WHERE batch_empleado.bpronro = " & NroProceso
OpenRecordset StrSql, rs2

'Seteo las variables de progreso
Progreso = 0
If Not rs2.EOF Then
    CEmpleadosAProc = rs2.RecordCount
    IncPorc = ((100 / CEmpleadosAProc))
Else
    If depurar Then
        Flog.writeline Espacios(Tabulador * 1) & "No hay empleados asociados al proceso" & " " & Now
    End If
    IncPorc = 100
    Exit Sub
End If

' Por cada empleado
Do While Not rs2.EOF
    l_ternro = rs2!Ternro
    StrSql = "SELECT * FROM gti_his_rnk "
    StrSql = StrSql & " WHERE ternro = " & l_ternro
    StrSql = StrSql & " AND anio = " & l_anio
    StrSql = StrSql & " AND cerrado = -1"
    l_cerrado = "no"
    OpenRecordset StrSql, rs3
    If Not rs3.EOF Then
        l_cerrado = "si"
        If depurar Then
            Flog.writeline Espacios(Tabulador * 1) & "El empleado " & l_terape & ", " & l_ternom & " - Legajo: " & l_empleg & ", estab cerrado, no se recalcula"
        End If
    End If

    If l_cerrado = "no" Then
        StrSql = "SELECT * FROM gti_his_rnk "
        StrSql = StrSql & " WHERE ternro = " & l_ternro
        StrSql = StrSql & " AND anio = " & CLng(l_anio) - 1
        OpenRecordset StrSql, rs
        If Not rs.EOF Then
            l_ant_anios = rs!antig_anio
            l_ant_dias = rs!antig_dias
            'l_dias_trabajados = rs!dias_trabajados
            l_porc_antig = rs!porc_antig
            l_tarjeta_ant = rs!tarjeta
        Else
            l_ant_anios = 0
            l_ant_dias = 0
            l_porc_antig = 0
            l_tarjeta_ant = 0
        End If

        l_empleg = rs2!empleg
        l_terape = rs2!terape
        l_ternom = rs2!ternom
        l_tarjeta = 0
        l_dife = 0
        l_observacion = ""
        l_dias_trabajados = 0
        l_situacion = ""


        ' Llamo a la subrutina 'CalcularDias' para que calcule el valor de la columna actual
        'Call CalcularDias(l_ternro, l_horas, l_dias_trabajados, l_anio, l_ant_anios)
        Call CalcularDias(l_ternro, l_horas, Tope_Dias, Tope_Horas, l_anio, l_dias_trabajados, l_ant_anios)
        StrSql = "SELECT * FROM gti_his_rnk "
        StrSql = StrSql & " WHERE ternro = " & l_ternro
        StrSql = StrSql & " AND anio = '" & l_anio & "'"
        l_modif = "no"
        OpenRecordset StrSql, rs
        If Not rs.EOF Then
            l_modif = "si"
        End If
        
        l_ant_dias = l_ant_dias + l_dias_trabajados
        
        'Porcentaje de antiguedad
        If l_ant_anios <= 10 Then
            l_porc_antig = l_ant_anios
        Else
            If l_ant_anios >= 11 And l_ant_anios <= 20 Then
                l_porc_antig = 10 + ((l_ant_anios - 10) * 1.25)
            Else
                l_porc_antig = 22.5 + ((l_ant_anios - 20) * 1.5)
            End If
        End If
        
        If l_modif = "no" Then
            StrSql = "INSERT INTO gti_his_rnk (bpronro, iduser, Ternro, empleg, terape, ternom, tarjeta, anio, cerrado, dias_trabajados, antig_anio,antig_dias, situacion, observacion, porc_antig, tarjeta_ant, dif) VALUES ("
            StrSql = StrSql & NroProceso & ", "
            StrSql = StrSql & "'" & Usuario & "', "
            StrSql = StrSql & l_ternro & ", "
            StrSql = StrSql & l_empleg & ", "
            StrSql = StrSql & "'" & l_terape & "', "
            StrSql = StrSql & "'" & l_ternom & "', "
            StrSql = StrSql & l_tarjeta & ", "
            StrSql = StrSql & "'" & l_anio & "', "
            StrSql = StrSql & 0 & ", "
            StrSql = StrSql & Fix(l_dias_trabajados) & ", "
            StrSql = StrSql & Fix(l_ant_anios) & ", "
            StrSql = StrSql & Fix(l_ant_dias) & ", "
            StrSql = StrSql & "'" & l_situacion & "', "
            StrSql = StrSql & "'" & l_observacion & "', "
            StrSql = StrSql & FormatNumber(l_porc_antig, 2) & ", "
            StrSql = StrSql & l_tarjeta_ant & ", "
            StrSql = StrSql & l_dife & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            If depurar Then
                Flog.writeline Espacios(Tabulador * 1) & "Se inserta el empleado: " & l_terape & ", " & l_ternom & " - Legajo: " & l_empleg
            End If
        Else
            StrSql = "UPDATE gti_his_rnk SET "
            StrSql = StrSql & " bpronro = " & NroProceso & ", "
            StrSql = StrSql & " iduser = " & "'" & Usuario & "', "
            StrSql = StrSql & " Ternro = " & l_ternro & ", "
            StrSql = StrSql & " empleg = " & l_empleg & ", "
            StrSql = StrSql & " terape = " & "'" & l_terape & "', "
            StrSql = StrSql & " ternom = " & "'" & l_ternom & "', "
            StrSql = StrSql & " tarjeta = " & l_tarjeta & ", "
            StrSql = StrSql & " anio = " & "'" & l_anio & "', "
            StrSql = StrSql & " cerrado = " & 0 & ", "
            StrSql = StrSql & " dias_trabajados = " & Fix(l_dias_trabajados) & ", "
            StrSql = StrSql & " antig_anio = " & Fix(l_ant_anios) & ", "
            StrSql = StrSql & " antig_dias = " & Fix(l_ant_dias) & ", "
            StrSql = StrSql & " situacion = " & "'" & l_situacion & "', "
            StrSql = StrSql & " observacion = " & "'" & l_observacion & "', "
            StrSql = StrSql & " porc_antig = " & l_porc_antig & ", "
            StrSql = StrSql & " tarjeta_ant = " & l_tarjeta_ant & ", "
            StrSql = StrSql & " dif = " & l_dife
            StrSql = StrSql & " WHERE ternro = " & l_ternro
            StrSql = StrSql & " AND anio = '" & l_anio & "'"
            objConn.Execute StrSql, , adExecuteNoRecords

            If depurar Then
                Flog.writeline Espacios(Tabulador * 1) & "Se modifico el empleado: " & l_terape & ", " & l_ternom & " - Legajo: " & l_empleg
            End If
        End If
    End If
    
    Progreso = Progreso + IncPorc
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & CInt(Progreso) & " WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords

    rs2.MoveNext
Loop

If depurar Then
    Flog.writeline Espacios(Tabulador * 1) & "Reordeno los empleados asignando Nº de tarjetas segun antiguedad..."
End If

' Tarjetas
StrSql = "SELECT * FROM gti_his_rnk "
StrSql = StrSql & " WHERE anio = " & l_anio
StrSql = StrSql & " ORDER BY antig_anio DESC, antig_dias DESC "
OpenRecordset StrSql, rs2
l_tarjeta = 0
Do While Not rs2.EOF
    l_tarjeta = l_tarjeta + 1
    
    If rs2!tarjeta_ant = 0 Then
        l_dife = 0
        l_observacion = "Primero"
    Else
        l_dife = rs2!tarjeta_ant - l_tarjeta
        
        If l_dife = 0 Then
            l_observacion = "Igual"
        Else
            If l_dife > 0 Then
                l_observacion = "Subio"
            Else
                l_observacion = "Bajo"
            End If
        End If
    End If
    
    StrSql = "UPDATE gti_his_rnk SET "
    StrSql = StrSql & " tarjeta = " & l_tarjeta & ","
    StrSql = StrSql & " dif = " & l_dife & ","
    StrSql = StrSql & " observacion = '" & l_observacion & "'"
    StrSql = StrSql & " WHERE anio = '" & l_anio & "'"
    StrSql = StrSql & " AND Ternro = " & rs2!Ternro
    objConn.Execute StrSql, , adExecuteNoRecords
    
    rs2.MoveNext
Loop

If depurar Then
    Flog.writeline Espacios(Tabulador * 1) & "Orden Tarjetas 'OK'"
End If

Fin:
'Cierro y libero
If rs.State = adStateOpen Then rs.Close
Set rs = Nothing
If rs2.State = adStateOpen Then rs2.Close
Set rs2 = Nothing
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

Private Sub BuscarHorasAusencia(Fecha As Date, Ternro As Long, ByRef Conjunto As String)
Dim rs As New ADODB.Recordset
Dim i As Integer


    ' seteo el turno y el dia
    'Set objBTurno.Conexion = objConn
    'objBTurno.Buscar_Turno Fecha, Ternro, False
    'initVariablesTurno objBTurno

    If Not tiene_turno Then

        ' -----------------------------------------------------------------
        ' si no tiene turno no puedo buscar en GTI_Config_tur_hor el tipo de hora ausencia
        ' la pregunta es: ¿Qué deberia retornar en el conjunto de horas de Ausencia ????
        ' -----------------------------------------------------------------
        Conjunto = "0"
        Exit Sub
    End If

    If tiene_turno Then
        StrSql = "SELECT gti_config_tur_hor.thnro as TipoHora FROM gti_config_hora " & _
                " INNER JOIN gti_config_tur_hor ON gti_config_hora.conhornro = gti_config_tur_hor.conhornro " & _
                " WHERE gti_config_hora.conhornro = 2 AND gti_config_tur_hor.turnro =" & nro_turno
        OpenRecordset StrSql, rs
    End If

    Conjunto = ""
    i = 1
    Do While Not rs.EOF
        If i = 1 Then
            Conjunto = rs("TipoHora")
            i = i + 1
        Else
            Conjunto = Conjunto & "," & rs("TipoHora")
        End If
        rs.MoveNext
    Loop

If rs.State = adStateOpen Then rs.Close
Set rs = Nothing

End Sub

Private Sub BuscarTodasLasHorasAusencia(ByRef Conjunto As String)
Dim rs As New ADODB.Recordset
Dim i As Integer
Dim Fecha As Date

' StrSql = " SELECT DISTINCT gti_config_tur_hor.thnro as TipoHora FROM gti_config_hora " & _
'          " INNER JOIN gti_config_tur_hor ON gti_config_hora.conhornro = gti_config_tur_hor.conhornro " & _
'          " WHERE gti_config_hora.conhornro = 2 "
StrSql = " SELECT confrep.confval as TipoHora " & _
         " FROM   confrep " & _
         " WHERE  confrep.repnro = 54 "
  OpenRecordset StrSql, rs
' Cambiado para mostrar los tipos de horas convertidas a unidades 'jornadas'
' O.D.A. 29/03/2004

  If rs.EOF Then
    Conjunto = "0"
  End If

  i = 1
  Do While Not rs.EOF
    If i = 1 Then
      Conjunto = rs("TipoHora")
      i = i + 1
    Else
      Conjunto = Conjunto & "," & rs("TipoHora")
    End If
    rs.MoveNext
  Loop

  If rs.State = adStateOpen Then rs.Close
  Set rs = Nothing
End Sub

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
   P_Asignacion = T.Tiene_PAsignacion
End Sub

Private Sub CalcularDias(ByVal Tercero As Long, ByVal th As String, ByVal Tope_Dias As Long, ByVal Tope_Horas As Long, ByVal anio As String, ByRef totaldias As Double, ByRef anios As Long)
Dim StrSql2 As String
Dim tot As Long
Dim l_rs3 As New ADODB.Recordset

StrSql2 = "SELECT sum(adcanthoras) tot FROM gti_acumdiario"
StrSql2 = StrSql2 & " WHERE thnro IN (" & th & ")"
StrSql2 = StrSql2 & " AND ternro = " & Tercero
StrSql2 = StrSql2 & " AND Year(adfecha) = " & anio
OpenRecordset StrSql2, l_rs3
totaldias = 0
If Not l_rs3.EOF Then
    If Not EsNulo(l_rs3!tot) Then
        totaldias = l_rs3!tot
    End If
End If

totaldias = Fix(totaldias / Tope_Horas)
If totaldias > Tope_Dias Then
    totaldias = Tope_Dias
End If
tot = totaldias / Tope_Dias
anios = anios + tot

'cerrar y liberar
If l_rs3.State = adStateOpen Then l_rs3.Close
Set l_rs3 = Nothing
End Sub
