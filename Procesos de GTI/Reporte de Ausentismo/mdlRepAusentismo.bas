Attribute VB_Name = "mdlRepAusentismo"
Option Explicit

'Const Version = 1
'Const FechaVersion = "20/10/2006"
'Modificaciones: Mariano Capriz
'               Se adapto el inicio del main para que corra con el nuevo appserver
'               Se agrego la version y log inicial

Global Const Version = 1.01
Global Const FechaVersion = "14/10/2009"
Global Const UltimaModificacion = "Encriptacion del string de conexion"
Global Const UltimaModificacion1 = "Manuel Lopez"

'---------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------

Dim fs
Dim Flog
Dim FDesde As Date
Dim FHasta As Date

Global HuboErrores As Boolean
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

Dim objBTurno As New BuscarTurno


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

Dim TiempoInicialProceso
Dim tituloReporte

Dim PID As String
Dim ArrParametros

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
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    
    TiempoInicialProceso = GetTickCount
    tituloReporte = ""
    depurar = False
    HuboErrores = False

    'Creo el archivo de texto del desglose
    Archivo = PathFLog & "RepAusentismo-" & CStr(NroProceso) & Format(Now, "DD-MM-YYYY") & ".log"
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
    Flog.writeline "Modificacion             : " & UltimaModificacion
    Flog.writeline "                         : " & UltimaModificacion1
    Flog.writeline "-------------------------------------------------"
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "Inicio :" & Now
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprcpid = " & PID & ", bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords

    Flog.writeline "Levanta Proceso y Setea Parámetros:  " & " " & Now

    'levanto los parametros del proceso
    StrParametros = ""

    StrSql = "SELECT bprcfecdesde,bprcfechasta,bprcparam FROM batch_proceso WHERE bpronro = " & NroProceso
    'rs.Open StrSql, objConn
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        FDesde = rs!bprcfecdesde
        FHasta = rs!bprcfechasta
        If Not IsNull(rs!bprcparam) Then
            If Len(rs!bprcparam) >= 1 Then
                pos = InStr(1, rs!bprcparam, ",")
                NroReporte = CLng(Left(rs!bprcparam, pos - 1))
                StrParametros = Right(rs!bprcparam, Len(rs!bprcparam) - (pos))
            End If
        End If
    Else
        Exit Sub
    End If
    If rs.State = adStateOpen Then rs.Close
    
    Flog.writeline "Inicio de Reporte de Ausencias: " & " " & Now
    
    Call Reporte_01(NroReporte, NroProceso, FDesde, FHasta, StrParametros)
    
    ' poner el bprcestado en procesado
    StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado', bprcprogreso = 100 WHERE bpronro = " & NroProceso
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

    Flog.writeline "Fin de Reporte de Ausencias: " & " " & Now

    If objConn.State = adStateOpen Then objConn.Close
    If objConnProgreso.State = adStateOpen Then objConnProgreso.Close
    
    Set rs_Batch_Proceso = Nothing
    Set rs_His_Batch_Proceso = Nothing
    
    Flog.Close

Exit Sub

CE:
'    Flog.writeline "Reporte abortado por Error:" & " " & Now
'    Flog.writeline "Ultimo SQL " & StrSql
    HuboErrores = True
    Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
        Flog.writeline Espacios(Tabulador * 0) & "Reporte abortado por Error:" & " " & Now
        Flog.writeline Espacios(Tabulador * 0) & " Error General " & Err.Description
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "SQL Ejecutado: " & StrSql
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
        Flog.writeline
    StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
End Sub


Private Sub Reporte_01(NroReporte As Long, NroProceso As Long, l_desde As Date, l_hasta As Date, Parametros As String)
Dim rs As New ADODB.Recordset ' l_rs
Dim rs2 As New ADODB.Recordset ' l_rs2

' declaracion de variable locales
Dim pos1 As Integer
Dim pos2 As Integer
Dim i As Integer

Dim HorasAusencia As String
Dim HorasLic As String
Dim HorasNov As String
Dim HorasCur As String
Dim HorasAnor As String
Dim l_union As String
Dim l_fecha As Date


' se supone que estos son parametros de entrada y vienen en "parametros"
Dim l_tenro1 As String
Dim l_tenro2 As String
Dim l_tenro3 As String

Dim l_estrnro1 As String
Dim l_estrnro2 As String
Dim l_estrnro3 As String

Dim l_filtro As String
Dim l_CodJust As Integer
Dim l_tipo As Integer   ' 0  = todas
                        ' -2 = Ausencias
                        ' -1 = Anormalidades
                        ' 1  = Licencias
                        ' 2  = Novedad
                        ' 3  = Cursos

Dim CEmpleadosAProc As Integer
Dim CDiasAProc As Integer
Dim IncPorc As Single
Dim Progreso As Single
Dim Columna As Integer

' ------------------------------------

'levanto cada parametro por separado, el separador de parametros es ";"
If Not IsNull(Parametros) Then
    If Len(Parametros) >= 1 Then
        pos1 = 1
        pos2 = InStr(pos1, Parametros, ",") - 1
        l_tenro1 = Mid(Parametros, pos1, pos2)
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, Parametros, ",") - 1
        l_tenro2 = Mid(Parametros, pos1, pos2 - pos1 + 1)
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, Parametros, ",") - 1
        l_tenro3 = Mid(Parametros, pos1, pos2 - pos1 + 1)
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, Parametros, ",") - 1
        l_estrnro1 = Mid(Parametros, pos1, pos2 - pos1 + 1)
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, Parametros, ",") - 1
        l_estrnro2 = Mid(Parametros, pos1, pos2 - pos1 + 1)
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, Parametros, ",") - 1
        l_estrnro3 = Mid(Parametros, pos1, pos2 - pos1 + 1)
        
        ' NO lo USO. Ahora los empleados vienen en batch_empleados
        pos1 = pos2 + 2
        pos2 = InStr(pos1, Parametros, ",") - 1
        l_filtro = Mid(Parametros, pos1, pos2 - pos1 + 1)
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, Parametros, ",") - 1
        l_tipo = CInt(Mid(Parametros, pos1, pos2 - pos1 + 1))
        
        pos1 = pos2 + 2
        pos2 = Len(Parametros)
        l_CodJust = CInt(Mid(Parametros, pos1, pos2 - pos1 + 1))
        
    End If
End If

'OpenConnection strconexion, objConn

' Busca todos los tipos de Hora de Ausencia para todos los turnos configurados y los devuelve
' en el conjunto "HorasAusencia"

If l_tipo = 0 Or l_tipo = -2 Then ' Ausencia  o Todas
    HorasAusencia = ""
    Columna = 1
    Call BuscarTodasLasHorasAusencia(HorasAusencia, Columna)
End If

If l_tipo = 0 Or l_tipo = 1 Then ' todas o Licencias
    HorasLic = ""
    Columna = 2
    Call BuscarTodasLasHorasAusencia(HorasLic, Columna)
Else
    HorasLic = "0"
End If

If l_tipo = 0 Or l_tipo = 2 Then ' Todas o Novedad
    HorasNov = ""
    Columna = 3
    Call BuscarTodasLasHorasAusencia(HorasNov, Columna)
Else
    HorasNov = "0"
End If

If l_tipo = 0 Or l_tipo = 3 Then ' todas o cursos
    HorasCur = ""
    Columna = 4
    Call BuscarTodasLasHorasAusencia(HorasCur, Columna)
Else
    HorasCur = "0"
End If

If l_tipo = 0 Or l_tipo = -1 Then ' todas o Anormalidad
    HorasAnor = ""
    Columna = 5
    Call BuscarTodasLasHorasAusencia(HorasAnor, Columna)
Else
    HorasAnor = "0"
End If


' Levanto todos los empleados a Procesar
StrSql = "SELECT empleado.ternro FROM empleado INNER JOIN batch_empleado ON empleado.ternro = batch_empleado.ternro " & _
         " WHERE batch_empleado.bpronro = " & NroProceso
rs.Open StrSql, objConn


If l_tipo = 0 Then
    l_union = " UNION "
Else
    l_union = " "
End If
    
' -------------------
CDiasAProc = DateDiff("d", l_desde, l_hasta) + 1
Progreso = 0


CEmpleadosAProc = rs.RecordCount
IncPorc = ((100 / CEmpleadosAProc))

' -------------------


Do Until rs.EOF
    StrSql = ""
    If l_tipo = 0 Or l_tipo = -2 Then ' Ausencia  o Todas
        StrSql = "SELECT 'hola'  as descripcion "
        StrSql = StrSql & "FROM gti_acumdiario "
        StrSql = StrSql & "WHERE ternro=" & rs("ternro") & " AND gti_acumdiario.thnro in (" & HorasAusencia & ")"
        StrSql = StrSql & " AND adfecha>=" & ConvFecha(l_desde) & " AND adfecha<=" & ConvFecha(l_hasta)
    End If
    If l_tipo = 0 Or l_tipo = 1 Or l_tipo = 2 Or l_tipo = 3 Then ' Todas, Licencias, Novedades, Cursos
        StrSql = "SELECT 'hola'  as descripcion "
        StrSql = StrSql & "FROM gti_acumdiario "
        StrSql = StrSql & "WHERE ternro=" & rs("ternro") & " AND gti_acumdiario.thnro in (" & HorasLic & "," & HorasNov & "," & HorasCur & ")"
        StrSql = StrSql & " AND adfecha>=" & ConvFecha(l_desde) & " AND adfecha<=" & ConvFecha(l_hasta)
    End If
    If l_tipo = 0 Or l_tipo = -1 Then ' Todas o Anormalidades
        StrSql = "SELECT 'hola'  as descripcion "
        StrSql = StrSql & "FROM gti_acumdiario "
        StrSql = StrSql & "WHERE ternro=" & rs("ternro") & " AND gti_acumdiario.thnro in (" & HorasAusencia & ")"
        StrSql = StrSql & " AND adfecha>=" & ConvFecha(l_desde) & " AND adfecha<=" & ConvFecha(l_hasta)
    End If


'    If l_tipo = 0 Or l_tipo = 1 Or l_tipo = 2 Or l_tipo = 3 Then ' Todas, Licencias, Novedades, Cursos
'        StrSql = StrSql & l_union
'        StrSql = StrSql & " SELECT 'Justificacion'  as descripcion "
'        StrSql = StrSql & " from gti_justificacion "
'        StrSql = StrSql & " WHERE gti_justificacion.ternro=" & rs("ternro")
'        StrSql = StrSql & " and ((gti_justificacion.jushasta>=" & ConvFecha(l_desde) & " and gti_justificacion.jushasta<=" & ConvFecha(l_hasta)
'        StrSql = StrSql & ") or (gti_justificacion.jusdesde<=" & ConvFecha(l_hasta) & " and gti_justificacion.jushasta>=" & ConvFecha(l_hasta) & "))"
'    End If
'
'    If l_tipo = 0 Or l_tipo = -1 Then ' Todas o Anormalidades
'        StrSql = StrSql & l_union
'        StrSql = StrSql & " SELECT 'Anorm' descripcion "
'        StrSql = StrSql & " from gti_horcumplido "
'        StrSql = StrSql & " INNER JOIN gti_anormalidad ON gti_horcumplido.normnro= gti_anormalidad.normnro or gti_horcumplido.normnro2= gti_anormalidad.normnro "
'        StrSql = StrSql & " where gti_horcumplido.ternro=" & rs("ternro")
'        StrSql = StrSql & " and gti_horcumplido.horfecrep>=" & ConvFecha(l_desde)
'        StrSql = StrSql & " and gti_horcumplido.horfecrep<=" & ConvFecha(l_hasta)
'    End If
    rs2.Open StrSql, objConn

    If Not rs2.EOF Then
        rs2.Close
        l_fecha = CDate(l_desde)
        
        Do Until (l_fecha = CDate(l_hasta) + 1)
            StrSql = ""
            If l_tipo = 0 Or l_tipo = -2 Then ' Todas o Ausencias
                StrSql = "SELECT -2 as TipoJust, 0 as CodJust, thdesc as tipo, '                         '  as descripcion, adcanthoras as horas "
                StrSql = StrSql & " FROM gti_acumdiario INNER JOIN tiphora ON gti_acumdiario.thnro=tiphora.thnro "
                StrSql = StrSql & " WHERE ternro=" & rs("ternro") & " AND gti_acumdiario.thnro in (" & HorasAusencia & ") AND adfecha=" & ConvFecha(l_fecha)
            End If
            If l_tipo = 0 Or l_tipo = 1 Then ' todas o Licencias
                StrSql = StrSql & l_union
                StrSql = StrSql & " SELECT -2 as TipoJust, 0 as CodJust, thdesc as tipo, '                         '  as descripcion, adcanthoras as horas "
                StrSql = StrSql & " FROM gti_acumdiario INNER JOIN tiphora ON gti_acumdiario.thnro=tiphora.thnro "
                StrSql = StrSql & " WHERE ternro=" & rs("ternro") & " AND gti_acumdiario.thnro in (" & HorasLic & ") AND adfecha=" & ConvFecha(l_fecha)
            End If
            If l_tipo = 0 Or l_tipo = 2 Then ' Todas o Novedad
                StrSql = StrSql & l_union
                StrSql = StrSql & " SELECT -2 as TipoJust, 0 as CodJust, thdesc as tipo, '                         '  as descripcion, adcanthoras as horas "
                StrSql = StrSql & " FROM gti_acumdiario INNER JOIN tiphora ON gti_acumdiario.thnro=tiphora.thnro "
                StrSql = StrSql & " WHERE ternro=" & rs("ternro") & " AND gti_acumdiario.thnro in (" & HorasNov & ") AND adfecha=" & ConvFecha(l_fecha)
            End If
            If l_tipo = 0 Or l_tipo = 3 Then ' todas o cursos
                StrSql = StrSql & l_union
                StrSql = StrSql & " SELECT -2 as TipoJust, 0 as CodJust, thdesc as tipo, '                         '  as descripcion, adcanthoras as horas "
                StrSql = StrSql & " FROM gti_acumdiario INNER JOIN tiphora ON gti_acumdiario.thnro=tiphora.thnro "
                StrSql = StrSql & " WHERE ternro=" & rs("ternro") & " AND gti_acumdiario.thnro in (" & HorasCur & ") AND adfecha=" & ConvFecha(l_fecha)
            End If
            If l_tipo = 0 Or l_tipo = -1 Then ' todas o Anormalidad
                StrSql = StrSql & l_union
                StrSql = StrSql & " SELECT -2 as TipoJust, 0 as CodJust, thdesc as tipo, '                         '  as descripcion, adcanthoras as horas "
                StrSql = StrSql & " FROM gti_acumdiario INNER JOIN tiphora ON gti_acumdiario.thnro=tiphora.thnro "
                StrSql = StrSql & " WHERE ternro=" & rs("ternro") & " AND gti_acumdiario.thnro in (" & HorasAnor & ") AND adfecha=" & ConvFecha(l_fecha)
            End If
            
            
'            If l_tipo = 0 Or l_tipo = 1 Then ' todas o Licencias
'                StrSql = StrSql & l_union
'                StrSql = StrSql & " SELECT gti_justificacion.tjusnro as TipoJust,  tipdia.tdnro as CodJust, 'Licencia' as tipo,tipdia.tddesc as descripcion,adcanthoras as horas "
'                StrSql = StrSql & " from gti_justificacion "
'                StrSql = StrSql & " INNER JOIN emp_lic ON gti_justificacion.juscodext=emp_lic.emp_licnro "
'                StrSql = StrSql & " INNER JOIN tipdia ON emp_lic.tdnro=tipdia.tdnro "
'                StrSql = StrSql & " LEFT JOIN gti_acumdiario ON gti_acumdiario.ternro=" & rs("ternro") & " and gti_acumdiario.adfecha=" & ConvFecha(l_fecha)
'                StrSql = StrSql & " and gti_acumdiario.thnro= tipdia.thnro "
'                If l_tipo = 1 Then
'                    StrSql = StrSql & " WHERE tipdia.tdnro = " & l_CodJust & " and jussigla ='LIC' and gti_justificacion.ternro=" & rs("ternro")
'                Else
'                    StrSql = StrSql & " WHERE jussigla ='LIC' and gti_justificacion.ternro=" & rs("ternro")
'                End If
'                StrSql = StrSql & " and gti_justificacion.jusdesde<=" & ConvFecha(l_fecha) & " and gti_justificacion.jushasta>=" & ConvFecha(l_fecha)
'            End If
'            If l_tipo = 0 Or l_tipo = 2 Then ' Todas o Novedad
'                StrSql = StrSql & l_union
'                StrSql = StrSql & " SELECT gti_justificacion.tjusnro as TipoJust,gti_tiponovedad.gtnovnro as CodJust, 'Novedad' as tipo, gti_novedad.gnovdesabr  as descripcion, adcanthoras as horas "
'                StrSql = StrSql & " from gti_justificacion "
'                StrSql = StrSql & " INNER JOIN  gti_novedad ON gti_justificacion.juscodext= gti_novedad.gnovnro "
'                StrSql = StrSql & " INNER JOIN  gti_tiponovedad ON gti_tiponovedad.gtnovnro= gti_novedad.gtnovnro "
'                StrSql = StrSql & " LEFT JOIN gti_acumdiario ON gti_acumdiario.ternro=" & rs("ternro") & " and gti_acumdiario.adfecha=" & ConvFecha(l_fecha)
'                StrSql = StrSql & " and gti_acumdiario.thnro= gti_tiponovedad.thnro "
'                If l_tipo = 2 Then
'                    StrSql = StrSql & " where gti_tiponovedad.gtnovnro = " & l_CodJust & " and jussigla ='NOV' and gti_justificacion.ternro=" & rs("ternro")
'                Else
'                    StrSql = StrSql & " where jussigla ='NOV' and gti_justificacion.ternro=" & rs("ternro")
'                End If
'                StrSql = StrSql & " and gti_justificacion.jusdesde<=" & ConvFecha(l_fecha) & " and gti_justificacion.jushasta>=" & ConvFecha(l_fecha)
'            End If
'            If l_tipo = 0 Or l_tipo = 3 Then ' todas o cursos
'                StrSql = StrSql & l_union
'                StrSql = StrSql & " SELECT gti_justificacion.tjusnro as TipoJust, 0 as CodJust,'Curso' as tipo,'                         '  as descripcion, juscanths as horas "
'                StrSql = StrSql & " from gti_justificacion where jussigla ='CUR' and gti_justificacion.ternro=" & rs("ternro")
'                StrSql = StrSql & " and gti_justificacion.jusdesde<=" & ConvFecha(l_fecha) & " and gti_justificacion.jushasta>=" & ConvFecha(l_fecha)
'            End If
'            If l_tipo = 0 Or l_tipo = -1 Then ' todas o Anormalidad
'                StrSql = StrSql & l_union
'                StrSql = StrSql & " SELECT -1 as TipoJust,0 as CodJust,'Anormalidad' as tipo,  gti_anormalidad.normdesabr as descripcion, horcant as horas "
'                StrSql = StrSql & " from gti_horcumplido "
'                StrSql = StrSql & " INNER JOIN gti_anormalidad ON gti_horcumplido.normnro= gti_anormalidad.normnro "
'                StrSql = StrSql & " where gti_horcumplido.ternro=" & rs("ternro")
'                StrSql = StrSql & " and gti_horcumplido.horfecrep=" & ConvFecha(l_fecha)
'                StrSql = StrSql & " UNION "
'                StrSql = StrSql & " SELECT -1 as TipoJust,0 as CodJust,'Anormalidad' as tipo,  gti_anormalidad.normdesabr as descripcion, horcant as horas "
'                StrSql = StrSql & " from gti_horcumplido "
'                StrSql = StrSql & " INNER JOIN gti_anormalidad ON gti_horcumplido.normnro2= gti_anormalidad.normnro "
'                StrSql = StrSql & " WHERE gti_horcumplido.normnro2<> gti_horcumplido.normnro and gti_horcumplido.ternro=" & rs("ternro")
'                StrSql = StrSql & " and gti_horcumplido.horfecrep=" & ConvFecha(l_fecha)
'            End If
                
                rs2.Open StrSql, objConn
                
                Do Until rs2.EOF
                    ' insertar (codigodeproceso,nroreporte,l_rs("ternro"),l_fecha,l_rs2("tipo"),l_rs2("descripcion"),l_rs2("horas"))
                    
                    If IsNull(rs2("horas")) Then
                        StrSql = "INSERT INTO rep_asp_01 (bprcnro,repnro,ternro,Fecha,causa,descripcion, TipoJust, CodJust) VALUES (" & _
                        NroProceso & "," & NroReporte & "," & rs("ternro") & "," & ConvFecha(l_fecha) & ",'" & rs2("tipo") & "','" & Left(rs2("descripcion"), 25) & "'," & rs2("TipoJust") & "," & rs2("CodJust") & ")"
                    Else
                        StrSql = "INSERT INTO rep_asp_01 (bprcnro,repnro,ternro,Fecha,causa,descripcion,horas, TipoJust, CodJust) VALUES (" & _
                        NroProceso & "," & NroReporte & "," & rs("ternro") & "," & ConvFecha(l_fecha) & ",'" & rs2("tipo") & "','" & Left(rs2("descripcion"), 25) & "'," & rs2("horas") & "," & rs2("TipoJust") & "," & rs2("CodJust") & ")"
                    End If

                    objConn.Execute StrSql, , adExecuteNoRecords
                    
                    Flog.writeline "inserta hora de Ausencias: " & " " & rs("ternro") & "," & ConvFecha(l_fecha) & ",'" & rs2("tipo") & "','" & rs2("descripcion") & "','" & rs2("horas") & "'," & rs2("TipoJust") & "," & rs2("CodJust")
                    
                    rs2.MoveNext
                Loop
                rs2.Close
        
        l_fecha = l_fecha + 1
        
        ' actualizar progreso
        'StrSql = "UPDATE batch_proceso SET bprcprogreso = " & CInt(Progreso) & " WHERE bpronro = " & NroProceso
        'objConn.Execute StrSql, , adExecuteNoRecords
        
        Loop
    
    Else
        rs2.Close
    End If
    
    'aca deberia actualizar el progreso
    ' actualizo tambien el porcentaje del empleado
    Progreso = Progreso + IncPorc
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & " WHERE bpronro = " & NroProceso
    objConnProgreso.Execute StrSql, , adExecuteNoRecords


    rs.MoveNext
Loop


If rs.State = adStateOpen Then rs.Close
Set rs = Nothing
If rs2.State = adStateOpen Then rs2.Close
Set rs2 = Nothing

End Sub


Private Sub BuscarHorasAusencia(Fecha As Date, Ternro As Long, ByRef Conjunto As String)
Dim rs As New ADODB.Recordset
Dim i As Integer


    ' seteo el turno y el dia
    Set objBTurno.Conexion = objConn
    objBTurno.Buscar_Turno Fecha, Ternro, False
    initVariablesTurno objBTurno
    
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



Private Sub BuscarTodasLasHorasAusencia(ByRef Conjunto As String, ByVal Col As Integer)
Dim rs As New ADODB.Recordset
Dim i As Integer
Dim Fecha As Date

' StrSql = "SELECT DISTINCT gti_config_tur_hor.thnro as TipoHora FROM gti_config_hora " & _
'          " INNER JOIN gti_config_tur_hor ON gti_config_hora.conhornro = gti_config_tur_hor.conhornro " & _
'          " WHERE gti_config_hora.conhornro = 2"
  StrSql = "SELECT confrep.confval as TipoHora " & _
           "FROM   confrep " & _
           "WHERE  confrep.repnro = 54 AND confrep.confnrocol = " & Col
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

