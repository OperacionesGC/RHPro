Attribute VB_Name = "mdlRepAnormalidades"
Option Explicit

'--------------------------------------------------
'Const Version = "1.00"
'Const FechaVersion = "13/01/2012"
'Modificaciones: Sebastian Stremel

'Const Version = "1.01"
'Const FechaVersion = "12/03/2012"
'Modificaciones: Sebastian Stremel - Barra de proceso

'Const Version = "1.03"
'Const FechaVersion = "19/04/2012"
'Modificaciones: Sebastian Stremel - Se modifico nombre de campo en consulta anormalidades

'Const Version = "1.04"
'Const FechaVersion = "23/04/2012"
'Modificaciones: Sebastian Stremel - Se modifico  barra de proceso

'Const Version = "1.05"
'Const FechaVersion = "25/04/2012"
'Modificaciones: Sebastian Stremel - Se modifico  barra de proceso

'Const Version = "1.06"
'Const FechaVersion = "02/05/2012"
'Modificaciones: Sebastian Stremel - Se eliminaron transacciones y se cambio calculo del tiempo


Const Version = "1.07"
Const FechaVersion = "21/10/2015"
'Modificaciones: MDF- CAS-32781 - CARGILL - Bug en Reporte de Anormalidades - se referencian los campos que se van a insertar



'   ====================================================================================================
'   ====================================================================================================

Dim fs
Dim Flog
Dim FDesde As Date
Dim FHasta As Date


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
Global CEmpleadosAProc As Long
Global IncPorc As Double
Global Progreso As Double
Global TiempoInicialProceso As Long
Global totalEmpleados As Long
Global TiempoAcumulado As Long
Global cantRegistros As Long

Dim objBTurno As New BuscarTurno
Dim l_iduser
Dim l_estrnro1
Dim l_estrnro2
Dim l_estrnro3




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

    'Creo el archivo de texto del desglose
    'Archivo = PathFLog & "RepAusentismo-" & CStr(NroProceso) & Format(Now, "DD-MM-YYYY") & ".log"
    Archivo = PathFLog & "RepAnormalidades-" & CStr(NroProceso) & ".log"

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
    
    'FGZ - 05/08/2009 --------- Control de versiones ------
    Version_Valida = ValidarV(Version, 1, TipoBD)
    If Not Version_Valida Then
        'Actualizo el progreso
        'MyBeginTrans
            StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Error de Version', bprcpid = " & PID & " WHERE bpronro = " & NroProceso
            objConnProgreso.Execute StrSql, , adExecuteNoRecords
        'MyCommitTrans
        Flog.writeline
        GoTo Final
    End If
    'FGZ - 05/08/2009 --------- Control de versiones ------

    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprcpid = " & PID & ", bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords

    Flog.writeline "Levanta Proceso y Setea Parámetros:  " & " " & Now

    'levanto los parametros del proceso
    StrParametros = ""

    StrSql = "SELECT bprcfecdesde,bprcfechasta,bprcparam,iduser FROM batch_proceso WHERE bpronro = " & NroProceso
    'rs.Open StrSql, objConn
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        FDesde = rs!bprcfecdesde
        FHasta = rs!bprcfechasta
        l_iduser = rs!iduser
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
    
    Flog.writeline "Inicio de Reporte de Anormalidades: " & " " & Now
    
    Call Reporte_02(NroReporte, NroProceso, FDesde, FHasta, StrParametros)
    'Call Reporte_01(NroReporte, NroProceso, FDesde, FHasta, StrParametros)
    
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
    ' FGZ - 22/09/2003
    ' -----------------------------------------------------------------------------------


    If objConn.State = adStateOpen Then objConn.Close
    If objConnProgreso.State = adStateOpen Then objConnProgreso.Close
    
    Set rs_Batch_Proceso = Nothing
    Set rs_His_Batch_Proceso = Nothing

Final:
    Flog.writeline Espacios(Tabulador * 0) & "Fin de Reporte de Anormalidades: " & " " & Now
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "---------------------------------------------------------------------------------"
    'Flog.writeline "Cantidad de Lecturas en BD          : " & Cantidad_de_OpenRecordset
    'Flog.writeline "Cantidad de llamadas a politicas    : " & Cantidad_Call_Politicas
    Flog.writeline Espacios(Tabulador * 0) & "---------------------------------------------------------------------------------"
    Flog.Close
Exit Sub

ce:
    Flog.writeline Espacios(Tabulador * 0) & "Reporte abortado por Error:" & " " & Now
    Flog.writeline Espacios(Tabulador * 0) & "Reporte abortado por :" & Err.Description
    Flog.writeline Espacios(Tabulador * 0) & "Ultimo SQL " & StrSql
    'MyRollbackTrans
    StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
End Sub

Private Sub Reporte_02(NroReporte As Long, NroProceso As Long, l_desde As Date, l_hasta As Date, parametros As String)
Dim rs As New ADODB.Recordset ' l_rs
Dim rs2 As New ADODB.Recordset ' l_rs2
Dim l_rs As New ADODB.Recordset
Dim l_rs3 As New ADODB.Recordset
Dim sql2 As String
Dim sql3 As String

Dim cadena
Dim cadenalic
Dim cadenanov
Dim ArrParametros
Dim l_estrnro1
Dim l_estrnro2
Dim l_estrnro3
Dim l_tenro1
Dim l_tenro2
Dim l_tenro3
Dim totaliza1
Dim totaliza2
Dim totaliza3
Dim totaliza
Dim l_filtro
Dim desde
Dim hasta
Dim sql
Dim sql1
Dim l_orden
Dim ncadenanov
Dim ncadenalic
Dim ncadenanorm
Dim j
Dim l_ternro
Dim l_sql
Dim l_tipnov
Dim l_tiplic
Dim l_tipanor
Dim l_repnro
Dim l_tipo As Integer   ' 0  = todas
                        ' -2 = Ausencias
                        ' -1 = Anormalidades
                        ' 1  = Licencias
                        ' 2  = Novedad
                        ' 3  = Cursos
'Dim l_tipo
Dim l_tiponov() As String
Dim l_tipolic() As String
Dim l_tipoanor() As String
Dim terape2 As String
Dim ternom2 As String

'Dim CEmpleadosAProc As Long
'Dim IncPorc As Double
'Dim Progreso As Double
'Dim TiempoInicialProceso As Long





Dim legdesde
Dim leghasta
Dim estado
'Dim l_repnro

On Error GoTo Err1

If Not IsNull(parametros) Then
    ArrParametros = Split(parametros, ",")
    Flog.writeline " limite del array --> " & UBound(ArrParametros)
    
    l_tenro1 = ArrParametros(0)
    l_tenro2 = ArrParametros(1)
    l_tenro3 = ArrParametros(2)
    l_estrnro1 = ArrParametros(3)
    l_estrnro2 = ArrParametros(4)
    l_estrnro3 = ArrParametros(5)
    totaliza1 = ArrParametros(6)
    totaliza2 = ArrParametros(7)
    totaliza3 = ArrParametros(8)
    l_filtro = ArrParametros(9)
    
    
 '-----------aca levanto anormalidadess------------------------------------
    cadena = ArrParametros(10)
    l_tipanor = Left(cadena, 1) 'busco el codigo de las anormalidades
    cadena = Mid(cadena, 1, Len(cadena))
    
    l_tipoanor = Split(cadena, "-")
    For j = 0 To UBound(l_tipoanor) - 1
        l_tipoanor(j) = Right(l_tipoanor(j), Len(l_tipoanor(j)) - 2)
        ncadenanorm = ncadenanorm + l_tipoanor(j) + ","
    Next
    ncadenanorm = Left(ncadenanorm, Len(ncadenanorm) - 1)
    
 '-------------------------------------------------------------------------

 '----------aca levanto licencias------------------------------------------
  
    cadenalic = ArrParametros(11)
    l_tiplic = Left(cadenalic, 1) 'busco el codigo de las anormalidades
    cadenalic = Mid(cadenalic, 1, Len(cadenalic))
    
    l_tipolic = Split(cadenalic, "-")
    For j = 0 To UBound(l_tipolic) - 1
        l_tipolic(j) = Right(l_tipolic(j), Len(l_tipolic(j)) - 2)
        ncadenalic = ncadenalic + l_tipolic(j) + ","
    Next
    ncadenalic = Left(ncadenalic, Len(ncadenalic) - 1)
 '-------------------------------------------------------------------------

 '----------aca levanto novedades------------------------------------------
    
    cadenanov = ArrParametros(12)
    l_tipnov = Left(cadenanov, 1) 'busco el codigo de las novedades
    cadenanov = Mid(cadenanov, 1, Len(cadenanov))
    
    l_tiponov = Split(cadenanov, "-")
    For j = 0 To UBound(l_tiponov) - 1
        l_tiponov(j) = Right(l_tiponov(j), Len(l_tiponov(j)) - 2)
        ncadenanov = ncadenanov + l_tiponov(j) + ","
    Next
    ncadenanov = Left(ncadenanov, Len(ncadenanov) - 1)
    
 '-------------------------------------------------------------------------

    
    desde = ArrParametros(13)
    hasta = ArrParametros(14)
    
    totaliza = ArrParametros(15)
    l_orden = ArrParametros(16)
    
    legdesde = ArrParametros(17)
    leghasta = ArrParametros(18)
    estado = ArrParametros(19)
l_filtro = Replace(l_filtro, "v_empleado", "empleado")
l_orden = Replace(l_orden, "v_empleado", "empleado")
End If

'inserto en la tabla de cabecera gti_anormalidades
If l_estrnro1 = "" Then
    l_estrnro1 = 0
End If
If l_estrnro2 = "" Then
    l_estrnro2 = 0
End If
If l_estrnro3 = "" Then
    l_estrnro3 = 0
End If

'MyBeginTrans
l_sql = " INSERT INTO rep_anormalidades (bpronro,iduser,fgeneracion,legdesde,leghasta,estado, te1,e1,tte1,te2,e2,tte2,te3,e3,tte3,fdesde,fhasta,novedad,licencia, anormalidad,totaliza) "
l_sql = l_sql & " VALUES (" & NroProceso & ", " & "'" & l_iduser & "'" & ", " & ConvFecha(Now()) & ", " & legdesde & ", "
l_sql = l_sql & leghasta & ", " & estado & ", " & l_tenro1 & ", " & l_estrnro1 & ", " & "'" & totaliza1 & "'" & ", " & l_tenro2 & ", " & l_estrnro2 & ", " & "'" & totaliza2 & "'" & ", "
l_sql = l_sql & l_tenro3 & ", " & l_estrnro3 & ", " & "'" & totaliza3 & "'" & ", " & ConvFecha(desde) & ", " & ConvFecha(hasta) & ", " & "''" & ", " & "''" & ", " & "''" & ", " & "'" & totaliza & "'" & ")"
objConn.Execute l_sql, , adExecuteNoRecords
Flog.writeline l_sql
'obtengo el numero de reporte una vez que se inserto en la cabecera
l_sql = "SELECT repnro FROM rep_anormalidades "
l_sql = l_sql & " WHERE bpronro =" & NroProceso
rs.Open l_sql, objConn
If Not rs.EOF Then
    l_repnro = rs!repnro
End If
rs.Close
'------------------------------------------------------------------

If l_tenro3 <> "" And l_tenro3 <> "0" Then
    sql = "SELECT DISTINCT empleado.ternro,empleg, terape, ternom, estact1.tenro AS tenro1, estact1.estrnro AS estrnro1, "
    sql = sql & " estact2.tenro AS tenro2, estact2.estrnro AS estrnro2, estact3.tenro AS tenro3, estact3.estrnro AS estrnro3 "
    sql = sql & " FROM empleado INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro AND estact1.htethasta IS NULL AND estact1.tenro  = " & l_tenro1
    If l_estrnro1 <> "" And l_estrnro1 <> "0" Then
        sql = sql & " AND estact1.estrnro =" & l_estrnro1
    End If
    sql = sql & " INNER JOIN his_estructura estact2 ON empleado.ternro = estact2.ternro AND estact2.htethasta IS NULL AND estact2.tenro  = " & l_tenro2
    
    If l_estrnro2 <> "" And l_estrnro2 <> "0" Then
        sql = sql & " AND estact2.estrnro =" & l_estrnro2
    End If
    sql = sql & " INNER JOIN his_estructura estact3 ON empleado.ternro = estact3.ternro AND estact3.htethasta IS NULL AND estact3.tenro  = " & l_tenro3
    
    If l_estrnro3 <> "" And l_estrnro3 <> "0" Then
        sql = sql & " AND estact3.estrnro =" & l_estrnro3
    End If
    sql = sql & " WHERE " & l_filtro
    sql = sql & " ORDER BY tenro1,estrnro1,tenro2,estrnro2,tenro3,estrnro3," & l_orden
    
Else
    If l_tenro2 <> "" And l_tenro2 <> "0" Then
        sql = "SELECT DISTINCT empleado.ternro,empleg, terape, ternom, estact1.tenro AS tenro1, estact1.estrnro AS estrnro1, "
        sql = sql & " estact2.tenro AS tenro2, estact2.estrnro AS estrnro2 "
        sql = sql & " FROM empleado INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro AND estact1.htethasta IS NULL AND estact1.tenro  = " & l_tenro1
            If l_estrnro1 <> "" And l_estrnro1 <> "0" Then
                sql = sql & " AND estact1.estrnro =" & l_estrnro1
            End If
        sql = sql & " INNER JOIN his_estructura estact2 ON empleado.ternro = estact2.ternro AND estact2.htethasta IS NULL AND estact2.tenro  = " & l_tenro2
            If l_estrnro2 <> "" And l_estrnro2 <> "0" Then
                sql = sql & " AND estact2.estrnro =" & l_estrnro2
            End If
        sql = sql & " WHERE " & l_filtro
        sql = sql & " ORDER BY tenro1,estrnro1,tenro2,estrnro2," & l_orden
    
    Else
            If l_tenro1 <> "" And l_tenro1 <> "0" Then
                sql = "SELECT DISTINCT empleado.ternro,empleg, terape, ternom, estact1.tenro AS tenro1, estact1.estrnro AS estrnro1 "
                sql = sql & " FROM empleado INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro AND estact1.htethasta IS NULL AND estact1.tenro  = " & l_tenro1
            If l_estrnro1 <> "" And l_estrnro1 <> "0" Then
                sql = sql & " AND estact1.estrnro =" & l_estrnro1
            End If
            sql = sql & " WHERE " & l_filtro
            sql = sql & " ORDER BY tenro1,estrnro1," & l_orden

    Else
            sql = "SELECT DISTINCT empleado.ternro,empleg, terape, ternom "
            sql = sql & "FROM empleado "
            sql = sql & " WHERE " & l_filtro
            sql = sql & " ORDER BY " & l_orden
   End If
            
End If

End If
rs.Open sql, objConn


'Determino la proporcion de progreso
Progreso = 0
totalEmpleados = rs.RecordCount
CEmpleadosAProc = rs.RecordCount
If CEmpleadosAProc = 0 Then
    CEmpleadosAProc = 1
End If
IncPorc = (99 / CEmpleadosAProc)
TiempoInicialProceso = GetTickCount





Do While Not rs.EOF
Progreso = Progreso + IncPorc

l_ternro = rs!Ternro
If l_tenro3 <> "" And l_tenro3 <> "0" Then
    l_estrnro1 = rs!estrnro1
    l_estrnro2 = rs!estrnro2
    l_estrnro3 = rs!estrnro3
Else
    If l_tenro2 <> "" And l_tenro2 <> "0" Then
        l_estrnro1 = rs!estrnro1
        l_estrnro2 = rs!estrnro2
    Else
        If l_tenro1 <> "" And l_tenro1 <> "0" Then
            l_estrnro1 = rs!estrnro1
        End If
   
    End If
End If
   'busco si el empleado tuvo novedades de algun tipo en las fechas del filtro
    If l_tipnov = 2 Then 'Novedad
        If ncadenanov <> 0 Then
            sql1 = " SELECT * from gti_justificacion "
            sql1 = sql1 & "INNER JOIN  gti_novedad ON gti_justificacion.juscodext= gti_novedad.gnovnro "
            sql1 = sql1 & "inner join tercero on tercero.ternro = gti_justificacion.ternro "
            sql1 = sql1 & "inner join empleado on empleado.ternro = tercero.ternro "
            sql1 = sql1 & " where jussigla ='NOV' and gti_justificacion.ternro=" & l_ternro & ""
            sql1 = sql1 & " and gti_justificacion.jusdesde >= " & ConvFecha(desde) & "  and gti_justificacion.jushasta <= " & ConvFecha(hasta)
            sql1 = sql1 & " and gtnovnro in (" & ncadenanov & ")"
            
            'sql1 = " SELECT * from gti_horcumplido "
            'sql1 = sql1 & " inner join gti_justificacion on gti_justificacion.jusnro = gti_horcumplido.jusnro "
            'sql1 = sql1 & " inner join tiphora on tiphora.thnro= gti_horcumplido.thnro "
            'sql1 = sql1 & " inner join tercero on tercero.ternro = gti_horcumplido.ternro "
            'sql1 = sql1 & " Where gti_horcumplido.thnro "
            'sql1 = sql1 & " IN (SELECT thnro FROM gti_tiponovedad WHERE gtnovnro in (" & ncadenanov & "))"
            'sql1 = sql1 & " AND horfecrep >=" & ConvFecha(desde) & " AND horfecrep<=" & ConvFecha(hasta)
            'sql1 = sql1 & " AND gti_horcumplido.ternro =" & l_ternro & ""
           
            'sql1 = "SELECT * from gti_horcumplido "
            'sql1 = sql1 & " WHERE thnro IN(SELECT thnro FROM gti_tiponovedad WHERE gtnovnro in (" & ncadenanov & "))"
            'sql1 = sql1 & " AND ternro =" & l_ternro & ""
            'sql1 = sql1 & " AND horfecrep >=" & ConvFecha(desde) & " AND horfecrep<=" & ConvFecha(hasta)
            'l_rs.Open sql1, objConn
        
        Else
            sql1 = " SELECT * from gti_justificacion "
            sql1 = sql1 & "INNER JOIN  gti_novedad ON gti_justificacion.juscodext= gti_novedad.gnovnro "
            sql1 = sql1 & "INNER JOIN tercero on tercero.ternro = gti_justificacion.ternro "
            sql1 = sql1 & "inner join empleado on empleado.ternro = tercero.ternro "
            sql1 = sql1 & " where jussigla ='NOV' and gti_justificacion.ternro=" & l_ternro & ""
            sql1 = sql1 & " and gti_justificacion.jusdesde >= " & ConvFecha(desde) & "  and gti_justificacion.jushasta <= " & ConvFecha(hasta)
            'sql1 = sql1 & " and gtnovnro in (" & ncadenanov & "))"
            
            'sql1 = " SELECT * from gti_horcumplido "
            'sql1 = sql1 & " inner join gti_justificacion on gti_justificacion.jusnro = gti_horcumplido.jusnro "
            'sql1 = sql1 & " inner join tiphora on tiphora.thnro= gti_horcumplido.thnro "
            'sql1 = sql1 & " inner join tercero on tercero.ternro = gti_horcumplido.ternro "
            'sql1 = sql1 & " Where gti_horcumplido.thnro "
            'sql1 = sql1 & " IN (SELECT thnro FROM gti_tiponovedad)"
            'sql1 = sql1 & " AND horfecrep >=" & ConvFecha(desde) & " AND horfecrep<=" & ConvFecha(hasta)
            'sql1 = sql1 & " AND gti_horcumplido.ternro =" & l_ternro & ""
            
            
            'sql1 = "SELECT * from gti_horcumplido "
            'sql1 = sql1 & " WHERE thnro IN(SELECT thnro FROM gti_tiponovedad)"
            'sql1 = sql1 & " AND ternro =" & l_ternro & ""
            'sql1 = sql1 & " AND horfecrep >=" & ConvFecha(desde) & " AND horfecrep<=" & ConvFecha(hasta)
        End If
        l_rs.Open sql1, objConn 'ejecuto consulta de novedades
        
        Do While Not l_rs.EOF   'PRIMER INSERT - TIENE 19 PARAMETRO
           'si entra aca tiene novedades
            sql2 = "insert into rep_anormalidades_det(bpronro,repnro,leg,ape1,ape2,nom1,nom2,e1,e2,e3,fech,ht,registraciones,clase,tipo,descripcion,cdec,chhmm,sexo)"
            sql2 = sql2 & "  VALUES (" & NroProceso & ", " & l_repnro & ", " & l_rs!empleg & ", "
            sql2 = sql2 & "'" & l_rs!terape & "'" & ", "
            If l_rs!terape2 = "" Or EsNulo(l_rs!terape2) Then
                terape2 = ""
            Else
                terape2 = l_rs!terape2
            End If
            sql2 = sql2 & "'" & terape2 & "'" & ", " & ""
            
            sql2 = sql2 & "'" & l_rs!ternom & "'" & ", " & ""
            
            If ((l_rs!ternom2 = "") Or EsNulo(l_rs!ternom2)) Then
                ternom2 = ""
            Else
                ternom2 = l_rs!ternom2
            End If
            sql2 = sql2 & "'" & ternom2 & "'" & ", " & ""
            
            'inserto las e1,e2,e3
            sql2 = sql2 & "'" & l_estrnro1 & "'" & ", " & ""
            sql2 = sql2 & "'" & l_estrnro2 & "'" & ", " & ""
            sql2 = sql2 & "'" & l_estrnro3 & "'" & ", " & ""
            sql2 = sql2 & "'" & l_rs!jusdesde & "'" & ", " & ""
            sql2 = sql2 & "'" & obtenerTeorico(l_rs!jusdesde, l_rs!Ternro) & "'" & ", "
            sql2 = sql2 & "'" & mostrarRegistracion(l_rs!jusdesde, l_rs!Ternro) & "'" & ", " & ""
            sql2 = sql2 & "'novedad'" & ", " & ""
            sql2 = sql2 & "'" & l_rs!gnovdesabr & "'" & ", " & ""
            sql2 = sql2 & "'" & l_rs!gnovdesext & "'" & ", " & ""
            
            'busco si tienen una justificacion
            sql3 = " SELECT * FROM gti_horcumplido where jusnro=" & l_rs!jusnro
            l_rs3.Open sql3, objConn
            If Not l_rs3.EOF Then
                sql2 = sql2 & "'" & l_rs3!horcant & "'" & ", " & ""
                sql2 = sql2 & "'" & l_rs3!Horas & "'" & ", " & ""
                
            Else
                sql2 = sql2 & "'" & 0 & "'" & ", " & ""
                sql2 = sql2 & "'" & 0 & "'" & ", " & ""
            
            End If
            l_rs3.Close
            sql2 = sql2 & "'" & l_rs!tersex & "'" & ")"
            objConn.Execute sql2, , adExecuteNoRecords
            Flog.writeline "Inserto la Novedad"
            Flog.writeline sql2
        l_rs.MoveNext 'registro de novedades
        Loop 'loop de novedades
        l_rs.Close
    End If
    
    
    'anormalidades
    If l_tipanor = 0 Then
        If ncadenanorm <> 0 Then 'busco alguna anormalidad en particular
            sql = " SELECT 'Anorm' as tipo,'' as estado, gti_anormalidad.normdesabr as descripcion,gti_horcumplido.normnro,gti_horcumplido.horas,gti_horcumplido.horcant, "
            sql = sql & " empleado.empleg, empleado.terape,empleado.terape2,empleado.ternom,empleado.ternom2,tercero.tersex, gti_horcumplido.horfecrep,gti_anormalidad.normdesabr,gti_anormalidad.normdesext,empleado.ternro   "
            sql = sql & " FROM gti_horcumplido "
            sql = sql & " INNER JOIN gti_anormalidad ON gti_horcumplido.normnro= gti_anormalidad.normnro"
            sql = sql & " INNER JOIN tercero on tercero.ternro = gti_horcumplido.ternro "
            sql = sql & " inner join empleado on empleado.ternro = tercero.ternro "
            sql = sql & " Where gti_horcumplido.Ternro =" & l_ternro
            sql = sql & " AND horfecrep >= " & ConvFecha(desde) & ""
            sql = sql & " AND horfecrep <= " & ConvFecha(hasta) & ""
            sql = sql & " AND gti_horcumplido.normnro in(" & ncadenanorm & ")"
            
            'sql = sql & " Union  "
                        
            sql = sql & " Union SELECT 'Anorm' as tipo,'' as estado, gti_anormalidad.normdesabr as descripcion, gti_horcumplido.normnro,gti_horcumplido.horas,gti_horcumplido.horcant, "
            sql = sql & " empleado.empleg, empleado.terape,empleado.terape2,empleado.ternom,empleado.ternom2,tercero.tersex, gti_horcumplido.horfecrep,gti_anormalidad.normdesabr,gti_anormalidad.normdesext,empleado.ternro   "
            sql = sql & " From gti_horcumplido"
            sql = sql & " INNER JOIN gti_anormalidad ON gti_horcumplido.normnro2= gti_anormalidad.normnro"
            sql = sql & " INNER JOIN tercero on tercero.ternro = gti_horcumplido.ternro "
            sql = sql & " inner join empleado on empleado.ternro = tercero.ternro "
            sql = sql & " Where gti_horcumplido.normnro2 <> gti_horcumplido.normnro And gti_horcumplido.Ternro =" & l_ternro
            sql = sql & " AND horfecrep >= " & ConvFecha(desde) & ""
            sql = sql & " AND horfecrep <= " & ConvFecha(hasta) & ""
            sql = sql & " AND gti_horcumplido.normnro in(" & ncadenanorm & ")"
    
            
            'sql = "SELECT * FROM gti_horcumplido "
            'sql = sql & "INNER JOIN tercero on tercero.ternro = gti_horcumplido.ternro "
            'sql = sql & "inner join empleado on empleado.ternro = tercero.ternro "
            'sql = sql & "inner join gti_anormalidad on gti_anormalidad.normnro= gti_horcumplido.normnro "
            'sql = sql & " WHERE gti_horcumplido.normnro IS NOT NULL "
            'sql = sql & " AND horfecgen >= " & ConvFecha(desde) & ""
            'sql = sql & " AND horfecgen <= " & ConvFecha(hasta) & ""
            'sql = sql & " AND normnro in(" & ncadenanorm & ")"""
            'sql = sql & " AND gti_horcumplido.ternro=" & l_ternro
        Else 'busco todas las anormalidades
            sql = " SELECT 'Anorm' as tipo,'' as estado, gti_anormalidad.normdesabr as descripcion,gti_horcumplido.normnro,gti_horcumplido.horas,gti_horcumplido.horcant, "
            sql = sql & " empleado.empleg, empleado.terape,empleado.terape2,empleado.ternom,empleado.ternom2,tercero.tersex, gti_horcumplido.horfecrep,gti_anormalidad.normdesabr,gti_anormalidad.normdesext,empleado.ternro   "
            sql = sql & " FROM gti_horcumplido "
            sql = sql & " INNER JOIN gti_anormalidad ON gti_horcumplido.normnro= gti_anormalidad.normnro"
            sql = sql & " INNER JOIN tercero on tercero.ternro = gti_horcumplido.ternro "
            sql = sql & " inner join empleado on empleado.ternro = tercero.ternro "
            sql = sql & " Where gti_horcumplido.Ternro =" & l_ternro
            sql = sql & " AND horfecrep >= " & ConvFecha(desde) & ""
            sql = sql & " AND horfecrep <= " & ConvFecha(hasta) & ""

            
           ' sql = sql & " Union "
                        
            sql = sql & " Union SELECT 'Anorm' as tipo,'' as estado, gti_anormalidad.normdesabr as descripcion, gti_horcumplido.normnro,gti_horcumplido.horas,gti_horcumplido.horcant, "
            sql = sql & " empleado.empleg, empleado.terape,empleado.terape2,empleado.ternom,empleado.ternom2,tercero.tersex, gti_horcumplido.horfecrep,gti_anormalidad.normdesabr,gti_anormalidad.normdesext,empleado.ternro   "
            sql = sql & " From gti_horcumplido"
            sql = sql & " INNER JOIN gti_anormalidad ON gti_horcumplido.normnro2= gti_anormalidad.normnro "
            sql = sql & " INNER JOIN tercero on tercero.ternro = gti_horcumplido.ternro "
            sql = sql & " inner join empleado on empleado.ternro = tercero.ternro "
            sql = sql & " Where gti_horcumplido.normnro2 <> gti_horcumplido.normnro And gti_horcumplido.Ternro =" & l_ternro
            sql = sql & " AND horfecrep >= " & ConvFecha(desde) & ""
            sql = sql & " AND horfecrep <= " & ConvFecha(hasta) & ""
         
            
            'sql = "SELECT * FROM gti_horcumplido "
            'sql = sql & "INNER JOIN tercero on tercero.ternro = gti_horcumplido.ternro "
            'sql = sql & "inner join empleado on empleado.ternro = tercero.ternro "
            'sql = sql & "inner join gti_anormalidad on gti_anormalidad.normnro= gti_horcumplido.normnro "
            'sql = sql & " WHERE gti_horcumplido.normnro IS NOT NULL "
            'sql = sql & " AND horfecgen >= " & ConvFecha(desde) & ""
            'sql = sql & " AND horfecgen <= " & ConvFecha(hasta) & ""
            'sql = sql & " AND gti_horcumplido.ternro=" & l_ternro
        End If
        l_rs.Open sql, objConn
        Do While Not l_rs.EOF 'SEGUNDO INSERT - TIENE 19 PARAMETROS
            sql2 = " insert into rep_anormalidades_det(bpronro,repnro,leg,ape1,ape2,nom1,nom2,e1,e2,e3,fech,ht,registraciones,clase,tipo,descripcion,cdec,chhmm,sexo) "
            sql2 = sql2 & " VALUES (" & NroProceso & ", " & l_repnro & ", " & l_rs!empleg & ", "
            sql2 = sql2 & "'" & l_rs!terape & "'" & ", "
            If l_rs!terape2 = "" Or EsNulo(l_rs!terape2) Then
                terape2 = ""
            Else
                terape2 = l_rs!terape2
            End If
            sql2 = sql2 & "'" & terape2 & "'" & ", " & ""
            
            sql2 = sql2 & "'" & l_rs!ternom & "'" & ", " & ""
            
            If ((l_rs!ternom2 = "") Or EsNulo(l_rs!ternom2)) Then
                ternom2 = ""
            Else
                ternom2 = l_rs!ternom2
            End If
            sql2 = sql2 & "'" & ternom2 & "'" & ", " & ""
            
            'inserto las e1,e2,e3
            sql2 = sql2 & "'" & l_estrnro1 & "'" & ", " & ""
            sql2 = sql2 & "'" & l_estrnro2 & "'" & ", " & ""
            sql2 = sql2 & "'" & l_estrnro3 & "'" & ", " & ""
            sql2 = sql2 & "'" & l_rs!horfecrep & "'" & ", " & ""
            sql2 = sql2 & "'" & obtenerTeorico(l_rs!horfecrep, l_rs!Ternro) & "'" & ", "
            sql2 = sql2 & "'" & mostrarRegistracion(l_rs!horfecrep, l_rs!Ternro) & "'" & ", " & ""
            sql2 = sql2 & "'ANORMALIDAD'" & ", " & ""
            sql2 = sql2 & "'" & l_rs!normdesabr & "'" & ", " & ""
            sql2 = sql2 & "'" & l_rs!normdesext & "'" & ", " & ""
            sql2 = sql2 & "'" & l_rs!horcant & "'" & ", " & ""
            sql2 = sql2 & "'" & l_rs!Horas & "'" & ", " & ""
            sql2 = sql2 & "'" & l_rs!tersex & "'" & ")"
            objConn.Execute sql2, , adExecuteNoRecords
            Flog.writeline "Inserto la anormalidad"
            Flog.writeline sql2
         l_rs.MoveNext 'registro de anormalidad
        Loop 'loop de anormalidad
        l_rs.Close
        'hasta aca-----------------------
        
    End If
    
'nuevo
    'licencias
    If l_tiplic = 1 Then
        If ncadenalic <> 0 Then 'busco alguna licencia en particular
            sql = " SELECT * From gti_justificacion "
            sql = sql & " INNER JOIN emp_lic ON gti_justificacion.juscodext=emp_lic.emp_licnro "
            sql = sql & " INNER JOIN tipdia ON emp_lic.tdnro=tipdia.tdnro "
            sql = sql & " INNER JOIN gti_horcumplido on gti_horcumplido.jusnro=gti_justificacion.jusnro "
            sql = sql & " INNER JOIN tercero on gti_horcumplido.ternro= tercero.ternro "
            sql = sql & " INNER JOIN empleado on empleado.ternro= tercero.ternro "
            sql = sql & " WHERE jussigla ='LIC' and gti_justificacion.ternro=" & l_ternro
            sql = sql & " AND gti_justificacion.jusdesde >=" & ConvFecha(desde) & " and gti_justificacion.jushasta <= " & ConvFecha(hasta) & ""
            sql = sql & " AND emp_lic.tdnro in (" & cadenalic & ")"
        Else 'busco todas las licencias
            sql = " SELECT * From gti_justificacion "
            sql = sql & " INNER JOIN emp_lic ON gti_justificacion.juscodext=emp_lic.emp_licnro "
            sql = sql & " INNER JOIN tipdia ON emp_lic.tdnro=tipdia.tdnro "
            sql = sql & " INNER JOIN gti_horcumplido on gti_horcumplido.jusnro=gti_justificacion.jusnro "
            sql = sql & " INNER JOIN tercero on gti_horcumplido.ternro= tercero.ternro "
            sql = sql & " INNER JOIN empleado on empleado.ternro= tercero.ternro "
            sql = sql & " WHERE jussigla ='LIC' and gti_justificacion.ternro=" & l_ternro
            sql = sql & " AND gti_justificacion.jusdesde >=" & ConvFecha(desde) & " and gti_justificacion.jushasta <=" & ConvFecha(hasta) & ""
            
        End If
        l_rs.Open sql, objConn
        Do While Not l_rs.EOF 'TERCER INSERT - TIENE 19 PARAMETROS
            sql2 = " insert into rep_anormalidades_det(bpronro,repnro,leg,ape1,ape2,nom1,nom2,e1,e2,e3,fech,ht,registraciones,clase,tipo,descripcion,cdec,chhmm,sexo) "
            sql2 = sql2 & "  VALUES (" & NroProceso & ", " & l_repnro & ", " & l_rs!empleg & ", "
            sql2 = sql2 & "'" & l_rs!terape & "'" & ", "
            If l_rs!terape2 = "" Or EsNulo(l_rs!terape2) Then
                terape2 = ""
            Else
                terape2 = l_rs!terape2
            End If
            sql2 = sql2 & "'" & terape2 & "'" & ", " & ""
            
            sql2 = sql2 & "'" & l_rs!ternom & "'" & ", " & ""
            
            If ((l_rs!ternom2 = "") Or EsNulo(l_rs!ternom2)) Then
                ternom2 = ""
            Else
                ternom2 = l_rs!ternom2
            End If
            sql2 = sql2 & "'" & ternom2 & "'" & ", " & ""
            
            'inserto las e1,e2,e3
            sql2 = sql2 & "'" & l_estrnro1 & "'" & ", " & ""
            sql2 = sql2 & "'" & l_estrnro2 & "'" & ", " & ""
            sql2 = sql2 & "'" & l_estrnro3 & "'" & ", " & ""
            sql2 = sql2 & "'" & l_rs!horfecrep & "'" & ", " & ""
            sql2 = sql2 & "'" & obtenerTeorico(l_rs!horfecrep, l_rs!Ternro) & "'" & ", "
            sql2 = sql2 & "'" & mostrarRegistracion(l_rs!horfecrep, l_rs!Ternro) & "'" & ", " & ""
            sql2 = sql2 & "'LICENCIA'" & ", " & ""
            sql2 = sql2 & "'" & l_rs!tddesc & "'" & ", " & ""
            sql2 = sql2 & "'" & l_rs!tddesc & "'" & ", " & ""
            sql2 = sql2 & "'" & l_rs!horcant & "'" & ", " & ""
            sql2 = sql2 & "'" & l_rs!Horas & "'" & ", " & ""
            sql2 = sql2 & "'" & l_rs!tersex & "'" & ")"
            objConn.Execute sql2, , adExecuteNoRecords
            Flog.writeline "Inserto la licencia"
            Flog.writeline sql2
        l_rs.MoveNext 'registro de licencia
        Loop 'loop de licencia
        l_rs.Close
        'hasta aca-----------------------
        
    End If
'aca deberia actualizar el progreso
' actualizo tambien el porcentaje del empleado
'Progreso = Progreso + IncPorc
'StrSql = "UPDATE batch_proceso SET bprcprogreso = " & CInt(Progreso) & " WHERE bpronro = " & NroProceso
'objConn.Execute StrSql, , adExecuteNoRecords

'Actualizo el estado del proceso
TiempoAcumulado = GetTickCount
               
CEmpleadosAProc = CEmpleadosAProc - 1
                  
StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
", bprcempleados ='" & CStr(CEmpleadosAProc) & "' WHERE bpronro = " & NroProceso
             
objConn.Execute StrSql, , adExecuteNoRecords

'MyBeginTrans
'StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
'StrSql = StrSql & " , bprctiempo = " & TiempoInicialProceso
'StrSql = StrSql & " WHERE bpronro = " & NroProceso
'objConnProgreso.Execute StrSql, , adExecuteNoRecords
'MyCommitTrans



rs.MoveNext 'registro de empleados
Loop 'loop de empleados


If rs.EOF Then
    'MyCommitTrans
    Flog.writeline "No hay mas empleados"
    GoTo Final
End If

Final:
    Flog.writeline Espacios(Tabulador * 0) & "Fin de Reporte de Anormalidades: " & " " & Now
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "---------------------------------------------------------------------------------"
    'Flog.writeline "Cantidad de Lecturas en BD          : " & Cantidad_de_OpenRecordset
    'Flog.writeline "Cantidad de llamadas a politicas    : " & Cantidad_Call_Politicas
    Flog.writeline Espacios(Tabulador * 0) & "---------------------------------------------------------------------------------"
    'Flog.Close
Exit Sub

Err1:
    'MyRollbackTrans
    Flog.writeline Espacios(Tabulador * 0) & "Reporte abortado por Error:" & " " & Now
    Flog.writeline Espacios(Tabulador * 0) & "Reporte abortado por :" & Err.Description
    Flog.writeline Espacios(Tabulador * 0) & "Ultimo SQL " & StrSql

End Sub
'Rutina que muestra las registraciones de un empleado en un dia
Function mostrarRegistracion(ByVal Dia, ByVal Ternro)
 Dim l_str
 Dim l_rs2 As New ADODB.Recordset ' l_rs2
 Dim l_cantreg
 Dim l_sql As String
 Dim l_salida
     l_salida = ""

    l_sql = "SELECT ternro, regfecha, reghora, regentsal, reldabr, regestado, regllamada "
    l_sql = l_sql & " FROM gti_registracion INNER JOIN gti_reloj ON gti_registracion.relnro = gti_reloj.relnro "
    l_sql = l_sql & " WHERE gti_registracion.ternro = " & Ternro
    l_sql = l_sql & " AND ((fechaproc = " & ConvFecha(Dia) & " AND (regestado = 'P' OR regestado = 'H'))"
    l_sql = l_sql & " OR (regfecha = " & ConvFecha(Dia) & " AND fechaproc IS NULL AND (regestado = 'P' OR regestado = 'H')))"
    l_sql = l_sql & " ORDER BY regfecha, reghora ASC "
    l_rs2.Open l_sql, objConn
    'rsOpen l_rs2, cn, l_sql, 0

    If Not l_rs2.EOF Then
       l_cantreg = 1
       
       'Cargo los valores en las columnas
        Do While Not l_rs2.EOF And l_cantreg < 9
            If l_rs2!reghora <> "" Then
               l_str = Mid(l_rs2!reghora, 1, 2) & ":" & Mid(l_rs2!reghora, 3, 2) & "-" & l_rs2!regentsal
               If CStr(l_rs2!regestado) <> "P" Then
                  l_str = l_str
               Else
                  l_str = l_str
               End If
               If l_salida = "" Then
                  l_salida = l_str
               Else
                  If (l_cantreg Mod 2) = 0 Then
                     l_salida = l_salida & "&nbsp;" & l_str
                  Else
                     l_salida = l_salida & "<br>" & l_str
                  End If
               End If
            End If
            l_cantreg = l_cantreg + 1
            l_rs2.MoveNext
            
        Loop
    End If

    l_rs2.Close
    
    mostrarRegistracion = l_salida

End Function 'mostrarRegistracion(dia,ternro)

'Rutina que se encarga de buscar el horario teorico
Function obtenerTeorico(ByVal Dia, ByVal Ternro)
Dim l_sql
Dim l_rs2 As New ADODB.Recordset
Dim l_salida
Dim l_hay_datos
    l_salida = ""
    l_hay_datos = False
    l_sql = " SELECT * FROM gti_proc_emp "
    l_sql = l_sql & " LEFT JOIN gti_dias ON gti_dias.dianro = gti_proc_emp.dianro "
    l_sql = l_sql & " WHERE gti_proc_emp.ternro = " & Ternro
    l_sql = l_sql & "   AND gti_proc_emp.fecha  = " & ConvFecha(Dia)
    l_rs2.Open l_sql, objConn
           
    If Not l_rs2.EOF Then
        If l_rs2!feriado = -1 Then
            l_salida = "Feriado"
        Else
            If l_rs2!pasig = -1 Then
                'Busco el parte
                l_sql = "SELECT * FROM gti_detturtemp WHERE (ternro = " & Ternro & ") AND "
                l_sql = l_sql & " (gttempdesde <= " & ConvFecha(Dia) & ") AND "
                l_sql = l_sql & " (" & ConvFecha(Dia) & " <= gttemphasta)"
                   
                l_rs2.Close
                      
                l_rs2.Open l_sql, objConn
                       
        
                If Not l_rs2.EOF Then
                    If Not IsNull(l_rs2!ttemphdesde1) Then
                        If l_rs2!ttemphdesde1 <> "" Then
                            l_salida = l_salida & convHora(l_rs2!ttemphdesde1) & "-" & convHora(l_rs2!ttemphhasta1)
                            l_hay_datos = True
                        End If
                    End If
                        If Not IsNull(l_rs2!ttemphdesde2) Then
                              If l_rs2!ttemphdesde2 <> "" Then
                                 If l_hay_datos Then
                                    l_salida = l_salida & "<br>"
                                 End If
                                 l_hay_datos = True
                                 l_salida = l_salida & convHora(l_rs2!ttemphdesde2) & "-" & convHora(l_rs2!ttemphhasta2)
                              End If
                           End If
                           If Not IsNull(l_rs2!ttemphdesde3) Then
                              If l_rs2!ttemphdesde3 <> "" Then
                                 If l_hay_datos Then
                                    l_salida = l_salida & "<br>"
                                 End If
                                 l_hay_datos = True
                                 l_salida = l_salida & convHora(l_rs2!ttemphdesde3) & "-" & convHora(l_rs2!ttemphhasta3)
                              End If
                           End If
        
                       Else
                           l_salida = l_salida & "Parte Asig. Hor. Eliminado"
                       End If
        
         '              l_salida = l_salida & "</font>"
                  Else
                     If l_rs2!dialibre = -1 Then
                         l_salida = "Franco"
                     Else
                           If Not IsNull(l_rs2!diahoradesde1) Then
                              If Replace(l_rs2!diahoradesde1, "0", "") <> "" And Replace(l_rs2!diahorahasta1, "0", "") <> "" Then
                                 l_salida = l_salida & convHora(l_rs2!diahoradesde1) & "-" & convHora(l_rs2!diahorahasta1)
                                 l_hay_datos = True
                              End If
                           End If
                           If Not IsNull(l_rs2!diahoradesde2) Then
                              If Replace(l_rs2!diahoradesde2, "0", "") <> "" And Replace(l_rs2!diahorahasta2, "0", "") <> "" Then
                                 If l_hay_datos Then
                                    l_salida = l_salida & "<br>"
                                 End If
                                 l_hay_datos = True
                                 l_salida = l_salida & convHora(l_rs2!diahoradesde2) & "-" & convHora(l_rs2!diahorahasta2)
                              End If
                           End If
                           If Not IsNull(l_rs2!diahoradesde3) Then
                              If Replace(l_rs2!diahoradesde3, "0", "") <> "" And Replace(l_rs2!diahorahasta3, "0", "") <> "" Then
                                 If l_hay_datos Then
                                    l_salida = l_salida & "<br>"
                                 End If
                                 l_hay_datos = True
                                 l_salida = l_salida & convHora(l_rs2!diahoradesde3) & "-" & convHora(l_rs2!diahorahasta3)
                              End If
                           End If
                     End If
                  End If
               End If
            Else
               l_salida = "Sin Procesar"
            End If
            
            l_rs2.Close
            
            obtenerTeorico = l_salida
        End Function
        
Function convHora(str)
  convHora = Mid(str, 1, 2) & ":" & Mid(str, 3, 2)
End Function 'convHora(str)

Private Sub Reporte_01(NroReporte As Long, NroProceso As Long, l_desde As Date, l_hasta As Date, parametros As String)
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
'Dim Progreso As Single
Dim Columna As Integer
Dim TotHorHHMM As String

' ------------------------------------

'levanto cada parametro por separado, el separador de parametros es ";"
If Not IsNull(parametros) Then
    If Len(parametros) >= 1 Then
        pos1 = 1
        pos2 = InStr(pos1, parametros, ",") - 1
        l_tenro1 = Mid(parametros, pos1, pos2)
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, ",") - 1
        l_tenro2 = Mid(parametros, pos1, pos2 - pos1 + 1)
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, ",") - 1
        l_tenro3 = Mid(parametros, pos1, pos2 - pos1 + 1)
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, ",") - 1
        l_estrnro1 = Mid(parametros, pos1, pos2 - pos1 + 1)
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, ",") - 1
        l_estrnro2 = Mid(parametros, pos1, pos2 - pos1 + 1)
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, ",") - 1
        l_estrnro3 = Mid(parametros, pos1, pos2 - pos1 + 1)
        
        ' NO lo USO. Ahora los empleados vienen en batch_empleados
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, ",") - 1
        l_filtro = Mid(parametros, pos1, pos2 - pos1 + 1)
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, ",") - 1
        l_tipo = CInt(Mid(parametros, pos1, pos2 - pos1 + 1))
        
        pos1 = pos2 + 2
        pos2 = Len(parametros)
        l_CodJust = CInt(Mid(parametros, pos1, pos2 - pos1 + 1))
        
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
        StrSql = StrSql & l_union
        StrSql = StrSql & " SELECT 'hola'  as descripcion "
        'StrSql = "SELECT 'hola'  as descripcion "
        StrSql = StrSql & "FROM gti_acumdiario "
        StrSql = StrSql & "WHERE ternro=" & rs("ternro") & " AND gti_acumdiario.thnro in (" & HorasLic & "," & HorasNov & "," & HorasCur & ")"
        StrSql = StrSql & " AND adfecha>=" & ConvFecha(l_desde) & " AND adfecha<=" & ConvFecha(l_hasta)
    End If
    If l_tipo = 0 Or l_tipo = -1 Then ' Todas o Anormalidades
        StrSql = StrSql & l_union
        StrSql = StrSql & " SELECT 'hola'  as descripcion "
        'StrSql = "SELECT 'hola'  as descripcion "
        StrSql = StrSql & "FROM gti_acumdiario "
        'FGZ - 19/10/2006
        'HorasAusencia x HorasAnor
        'StrSql = StrSql & "WHERE ternro=" & rs("ternro") & " AND gti_acumdiario.thnro in (" & HorasAusencia & ")"
        StrSql = StrSql & "WHERE ternro=" & rs("ternro") & " AND gti_acumdiario.thnro in (" & HorasAnor & ")"
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
                        'FGZ - 17/05/2010 ----------------
                        'StrSql = "INSERT INTO rep_asp_01 (bprcnro,repnro,ternro,Fecha,causa,descripcion,horas, TipoJust, CodJust) VALUES (" & _
                        'NroProceso & "," & NroReporte & "," & rs("ternro") & "," & ConvFecha(l_fecha) & ",'" & rs2("tipo") & "','" & Left(rs2("descripcion"), 25) & "'," & rs2("horas") & "," & rs2("TipoJust") & "," & rs2("CodJust") & ")"
                        
                        
                        TotHorHHMM = CHoras(rs2("horas"), 60)
                        'FGZ - 17/05/2010 ----------------
                    
                        StrSql = "INSERT INTO rep_asp_01 (bprcnro,repnro,ternro,Fecha,causa,descripcion,horastr,horas, TipoJust, CodJust) VALUES (" & _
                        NroProceso & "," & NroReporte & "," & rs("ternro") & "," & ConvFecha(l_fecha) & ",'" & rs2("tipo") & "','" & Left(rs2("descripcion"), 25) & "'," & TotHorHHMM & "," & rs2("horas") & "," & rs2("TipoJust") & "," & rs2("CodJust") & ")"
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

