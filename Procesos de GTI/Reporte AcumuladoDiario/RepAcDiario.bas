Attribute VB_Name = "mdlRepAcDiario"
Option Explicit

'--------------------------------------------------
'Const Version = "1.00"
'Const FechaVersion = "10/05/2013"
'Gonzalez Nicolás

'Const Version = "1.01"
'Const FechaVersion = "01/07/2013"
'MODIFICACIONES :
'                 Mauricio Zwenger - CAS-19229 - MONRESA - BUG EN EL  REPORTE DE ACUMULADO DIARIO
'
Const Version = "1.02"
Const FechaVersion = "04/07/2013"
'MODIFICACIONES :
'                 Mauricio Zwenger - CAS-19229 - MONRESA - BUG EN EL  REPORTE DE ACUMULADO DIARIO
'                                                cambio a tipo long de variables integer

'   ====================================================================================================
'   ====================================================================================================

Dim fs
Dim Flog
Dim FDesde As Date
Dim FHasta As Date
Global Fecha_Inicio As Date
Global CEmpleadosAProc As Long
Global IncPorc As Double
Global Progreso As Double
Global TiempoInicialProceso As Long
Global totalEmpleados As Long
Global TiempoAcumulado As Long
Global cantRegistros As Long
Dim l_iduser
Dim l_estrnro1
Dim l_estrnro2
Dim l_estrnro3




Sub Main()

Dim Archivo As String
Dim pos As Long
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
    Archivo = PathFLog & "RepAcumuladoDiario-" & CStr(NroProceso) & ".log"

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
    
    '01/07/2013 - MDZ - CAS-19229 - MONRESA - BUG EN EL  REPORTE DE ACUMULADO DIARIO
    'rs.Open StrSql, objConn
    OpenRecordset StrSql, rs
    
    If Not rs.EOF Then
        'FDesde = rs!bprcfecdesde
        'FHasta = rs!bprcfechasta
        'l_iduser = rs!iduser
        If Not IsNull(rs!bprcparam) Then
            If Len(rs!bprcparam) >= 1 Then
                'pos = InStr(1, rs!bprcparam, ",")
                'NroReporte = CLng(Left(rs!bprcparam, pos - 1))
                'StrParametros = Right(rs!bprcparam, Len(rs!bprcparam) - (pos))
                'If rs.State = adStateOpen Then rs.Close
                Flog.writeline "Inicio de Reporte de Acumulado diario: " & " " & Now
                Call Rep_AcumD(rs!bprcparam, NroProceso)
            End If
        End If
    Else
        Exit Sub
    End If
    
   
    
    'Call Rep_AcumD(NroReporte, NroProceso, FDesde, FHasta, StrParametros)
    
    
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
    Flog.writeline Espacios(Tabulador * 0) & "Fin de Reporte de Acumlado Diario: " & " " & Now
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
Private Sub Rep_AcumD(parametros As String, NroProceso As Long)

Dim rs As New ADODB.Recordset ' l_rs

' declaracion de variable locales
' Dim CEmpleadosAProc As Integer
Dim CDiasAProc As Long
Dim IncPorc As Double
'------
Dim emplegDesde As String
Dim emplegHasta As String
Dim tenro1 As Long
Dim tenro2 As Long
Dim tenro3 As Long

Dim estrnro1 As Long
Dim estrnro2 As Long
Dim estrnro3 As Long

Dim agrupa1 As Long
Dim agrupa2 As Long
Dim agrupa3 As Long
Dim Orden As String
Dim filtro As String
Dim detallado As Long

Dim ArrParametros
Dim Continua As Boolean
'-------------------------------------
Dim Ternro
Dim DescNivel1 As String
Dim DescNivel2 As String
Dim DescNivel3 As String

Dim adhistoricoDesc As String
Dim i
Continua = False
' ------------------------------------

'======================================================================================================
'SE VALIDA Y LEVANTAN PARAMETROS.
'======================================================================================================
If Not IsNull(parametros) And Len(parametros) >= 1 Then
    'If Len(parametros) >= 1 Then
        
        '1@2147483647@0@0@false@0@0@false@0@0@false@07/05/2013@07/05/2013@empleg@-1
        ArrParametros = Split(parametros, "@")
        If UBound(ArrParametros) = 14 Then
            emplegDesde = ArrParametros(0)
            emplegHasta = ArrParametros(1)
            
            adhistoricoDesc = "Legajos " & emplegDesde & " hasta " & emplegHasta
            filtro = " empleado.empleg >=" & emplegDesde & " AND empleado.empleg <=" & emplegHasta
            
            'NIVEL 1
            tenro1 = ArrParametros(2)
            estrnro1 = ArrParametros(3)
            agrupa1 = 0
            If CBool(ArrParametros(4)) = True Then
                agrupa1 = -1
            End If
            'NIVEL 2
            tenro2 = ArrParametros(5)
            estrnro2 = ArrParametros(6)
            agrupa2 = 0
            If CBool(ArrParametros(7)) = True Then
                agrupa2 = -1
            End If
            
            'NIVEL 3
            tenro3 = ArrParametros(8)
            estrnro3 = ArrParametros(9)
            agrupa3 = 0
            If CBool(ArrParametros(10)) = True Then
                agrupa3 = -1
            End If
            
            If ArrParametros(11) <> "" Then
                FDesde = ArrParametros(11)
            Else
                FDesde = Date
            End If
            If ArrParametros(12) <> "" Then
                FHasta = ArrParametros(12)
            Else
                FHasta = Date
            End If
            
            filtro = filtro & " AND adfecha >= " & ConvFecha(FDesde) & " AND adfecha <=" & ConvFecha(FHasta)
            
            adhistoricoDesc = adhistoricoDesc & " Desde " & FDesde & " hasta " & FHasta
            
            Orden = ArrParametros(13)
            
            detallado = CBool(ArrParametros(14))
           
            Continua = True
            
            Flog.writeline ""
            Flog.writeline "Parametros:" & parametros
            Flog.writeline ""
        Else
            Flog.writeline "Error en parámetros al generar el proceso."
            Exit Sub
        End If
End If

'======================================================================================================
'ARMO QUERY PRINCIPAL
'======================================================================================================
If Continua = True Then
    If tenro3 <> 0 Then
        StrSql = "SELECT empleado.ternro,empleg, terape, ternom, estact1.tenro tenro1, estact1.estrnro estrnro1 "
        StrSql = StrSql & ",estact2.tenro tenro2, estact2.estrnro estrnro2, estact3.tenro tenro3, estact3.estrnro estrnro3 "
        StrSql = StrSql & ", gti_acumdiario.adfecha, gti_acumdiario.adcanthoras, tiphora.thdesc, gti_acumdiario.admanual, gti_acumdiario.advalido, gti_acumdiario.adestado, gti_acumdiario.horas"
        StrSql = StrSql & " FROM empleado "
        StrSql = StrSql & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro "
        StrSql = StrSql & " AND estact1.htethasta IS NULL AND estact1.tenro  = " & tenro1
        If estrnro1 <> 0 Then
            StrSql = StrSql & " AND estact1.estrnro =" & estrnro1
        End If
        StrSql = StrSql & " INNER JOIN his_estructura estact2 ON empleado.ternro = estact2.ternro AND estact2.htethasta IS NULL  AND estact2.tenro  = " & tenro2
        
        If estrnro2 <> 0 Then
            StrSql = StrSql & " AND estact2.estrnro =" & estrnro2
        End If
        
        StrSql = StrSql & " INNER JOIN his_estructura estact3 ON empleado.ternro = estact3.ternro AND estact3.htethasta IS NULL  AND estact3.tenro  = " & tenro3
        If estrnro3 <> 0 Then
            StrSql = StrSql & " AND estact3.estrnro =" & estrnro3
        End If
        StrSql = StrSql & " INNER JOIN gti_acumdiario ON gti_acumdiario.ternro = empleado.ternro"
        StrSql = StrSql & " INNER JOIN tiphora ON tiphora.thnro = gti_acumdiario.thnro"
        StrSql = StrSql & " WHERE " & filtro
        StrSql = StrSql & " ORDER BY tenro1,estrnro1,tenro2,estrnro2,tenro3,estrnro3," & Orden
        StrSql = StrSql & ", gti_acumdiario.adfecha"
        
    ElseIf tenro2 <> 0 Then
        StrSql = "SELECT empleado.ternro,empleg, terape, ternom, estact1.tenro  tenro1, estact1.estrnro  estrnro1 "
        StrSql = StrSql & ",estact2.tenro  tenro2, estact2.estrnro  estrnro2 "
        StrSql = StrSql & ", gti_acumdiario.adfecha, gti_acumdiario.adcanthoras, tiphora.thdesc, gti_acumdiario.admanual, gti_acumdiario.advalido, gti_acumdiario.adestado, gti_acumdiario.horas"
        StrSql = StrSql & " FROM empleado "
        StrSql = StrSql & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro AND estact1.htethasta IS NULL AND estact1.tenro  = " & tenro1
        If estrnro1 <> 0 Then
            StrSql = StrSql & " AND estact1.estrnro =" & estrnro1
        End If
        StrSql = StrSql & " INNER JOIN his_estructura estact2 ON empleado.ternro = estact2.ternro AND estact2.htethasta IS NULL AND estact2.tenro  = " & tenro2
        If estrnro2 <> 0 Then
            StrSql = StrSql & " AND estact2.estrnro =" & estrnro2
        End If
        StrSql = StrSql & " INNER JOIN gti_acumdiario ON gti_acumdiario.ternro = empleado.ternro"
        StrSql = StrSql & " INNER JOIN tiphora ON tiphora.thnro = gti_acumdiario.thnro"
        StrSql = StrSql & " WHERE " & filtro
        StrSql = StrSql & " ORDER BY tenro1,estrnro1,tenro2,estrnro2," & Orden
        StrSql = StrSql & ", gti_acumdiario.adfecha"
    
    ElseIf tenro1 <> 0 Then
        StrSql = "SELECT empleado.ternro,empleg, terape, ternom, estact1.tenro tenro1, estact1.estrnro estrnro1 "
        StrSql = StrSql & ", gti_acumdiario.adfecha, gti_acumdiario.adcanthoras, tiphora.thdesc, gti_acumdiario.admanual, gti_acumdiario.advalido, gti_acumdiario.adestado, gti_acumdiario.horas"
        StrSql = StrSql & " FROM empleado "
        StrSql = StrSql & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro AND estact1.htethasta IS NULL AND estact1.tenro  = " & tenro1
        
        StrSql = StrSql & " INNER JOIN gti_acumdiario ON gti_acumdiario.ternro = empleado.ternro"
        StrSql = StrSql & " INNER JOIN tiphora ON tiphora.thnro = gti_acumdiario.thnro"
        
        If estrnro1 <> 0 Then
            StrSql = StrSql & " AND estact1.estrnro =" & estrnro1
        End If
        StrSql = StrSql & " WHERE " & filtro
        StrSql = StrSql & " ORDER BY tenro1,estrnro1," & Orden
        StrSql = StrSql & ", gti_acumdiario.adfecha"

    Else
        StrSql = "SELECT empleado.ternro,empleg, terape, ternom "
        StrSql = StrSql & ", gti_acumdiario.adfecha, gti_acumdiario.adcanthoras, tiphora.thdesc, gti_acumdiario.admanual, gti_acumdiario.advalido, gti_acumdiario.adestado, gti_acumdiario.horas"
        StrSql = StrSql & " FROM Empleado"

        StrSql = StrSql & " INNER JOIN gti_acumdiario ON gti_acumdiario.ternro = empleado.ternro"
        StrSql = StrSql & " INNER JOIN tiphora ON tiphora.thnro = gti_acumdiario.thnro"
        
        StrSql = StrSql & " WHERE " & filtro
        StrSql = StrSql & " ORDER BY " & Orden
        StrSql = StrSql & ", gti_acumdiario.adfecha"
    End If
    Flog.writeline StrSql
    Flog.writeline ""
    
    '01/07/2013 - MDZ - CAS-19229 - MONRESA - BUG EN EL  REPORTE DE ACUMULADO DIARIO
    'rs.Open StrSql, objConn
    OpenRecordset StrSql, rs
    
    If Not rs.EOF Then
    
    
    Progreso = 0
    totalEmpleados = rs.RecordCount
    CEmpleadosAProc = rs.RecordCount
    If CEmpleadosAProc = 0 Then
        CEmpleadosAProc = 1
    End If
    IncPorc = CDbl((100 / CEmpleadosAProc))
    TiempoInicialProceso = GetTickCount
    i = 0
    
        Ternro = ""
        DescNivel1 = ""
        DescNivel2 = ""
        DescNivel3 = ""
        Do While Not rs.EOF
            i = i + 1
            Progreso = Progreso + IncPorc
            If Ternro <> rs!Ternro Then
            'CAMBIO DE EMPLEADO Y GUARDO REGISTRO CON LOS TOTALES DE HORAS.
                DescNivel1 = ""
                DescNivel2 = ""
                DescNivel3 = ""
                estrnro1 = 0
                estrnro2 = 0
                estrnro3 = 0
                'BUSCO NOMBRE DEL NIVEL 1
                If tenro1 <> 0 Then
                    DescNivel1 = descripcionEstructura(tenro1, rs!estrnro1)
                    estrnro1 = rs!estrnro1
                End If
                'BUSCO NOMBRE DEL NIVEL 2
                If tenro2 <> 0 Then
                    DescNivel2 = descripcionEstructura(tenro2, rs!estrnro2)
                    estrnro2 = rs!estrnro2
                End If
                'BUSCO NOMBRE DEL NIVEL 3
                If tenro3 <> 0 Then
                    DescNivel3 = descripcionEstructura(tenro3, rs!estrnro3)
                    estrnro3 = rs!estrnro3
                End If
                Ternro = rs!Ternro
            End If
            
            
            'TotalHsHs = TotalHsHs + rs!Horas
            'TotalHsDec = TotalHsDec + rs!adcanthoras
            
            
            StrSql = "INSERT INTO rep_Acu_diario"
            StrSql = StrSql & "(bpronro,adhistoricoDesc, Ternro, adLegajo, adapellido, adNombre, adFecha, adDia, adtenro1, adestrnro1,adDescNivel1,adtotaliza1, adtenro2, adestrnro2,adDescNivel2,adtotaliza2, adtenro3, adestrnro3,adDescNivel3,adtotaliza3, adcanths, adcanthsDec, adtipohs, admanual, advalido, adtotalhs,addetallado)"
            StrSql = StrSql & " VALUES ("
            StrSql = StrSql & NroProceso
            StrSql = StrSql & ",'" & adhistoricoDesc & "'"
            StrSql = StrSql & "," & rs!Ternro
            StrSql = StrSql & "," & rs!empleg
            StrSql = StrSql & ",'" & rs!terape & "'"
            StrSql = StrSql & ",'" & rs!ternom & "'"
            StrSql = StrSql & ",'" & rs!adfecha & "'"
            StrSql = StrSql & ",'" & Weekday(rs!adfecha) & "'"
            StrSql = StrSql & "," & tenro1
            StrSql = StrSql & "," & estrnro1
            StrSql = StrSql & ",'" & DescNivel1 & "'"
            StrSql = StrSql & "," & agrupa1
            StrSql = StrSql & "," & tenro2
            StrSql = StrSql & "," & estrnro2
            StrSql = StrSql & ",'" & DescNivel2 & "'"
            StrSql = StrSql & "," & agrupa2
            StrSql = StrSql & "," & tenro3
            StrSql = StrSql & "," & estrnro3
            StrSql = StrSql & ",'" & DescNivel3 & "'"
            StrSql = StrSql & "," & agrupa3
            StrSql = StrSql & ",'" & rs!Horas & "'"
            StrSql = StrSql & "," & CDbl(rs!adcanthoras)
            StrSql = StrSql & ",'" & rs!thdesc & "'"
            StrSql = StrSql & ",'" & rs!admanual & "'"
            StrSql = StrSql & ",'" & rs!advalido & "'"
            StrSql = StrSql & ",null"
            StrSql = StrSql & "," & detallado
            StrSql = StrSql & ")"
            
            objConn.Execute StrSql, , adExecuteNoRecords
            Flog.writeline "Inserta registro N°: " & i
            Flog.writeline ""
            
            
            'Actualizo el progreso
            TiempoAcumulado = GetTickCount
            CEmpleadosAProc = CEmpleadosAProc - 1
            StrSql = "UPDATE batch_proceso SET bprcprogreso = " & CDbl(Progreso)
            StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'"
            StrSql = StrSql & ", bprcempleados ='" & CStr(CEmpleadosAProc) & "' WHERE bpronro = " & NroProceso
            objConn.Execute StrSql, , adExecuteNoRecords
            
            rs.MoveNext
        Loop
    Else
            StrSql = "INSERT INTO rep_Acu_diario"
            StrSql = StrSql & "(bpronro,adhistoricoDesc, Ternro, adLegajo, adapellido, adNombre, adFecha, adDia, adtenro1, adestrnro1,adDescNivel1,adtotaliza1, adtenro2, adestrnro2,adDescNivel2,adtotaliza2, adtenro3, adestrnro3,adDescNivel3,adtotaliza3, adcanths, adcanthsDec, adtipohs, admanual, advalido, adtotalhs,addetallado)"
            StrSql = StrSql & " VALUES ("
            StrSql = StrSql & NroProceso
            StrSql = StrSql & ",'" & adhistoricoDesc & "'"
            StrSql = StrSql & ",0"
            StrSql = StrSql & ",0"
            StrSql = StrSql & ",null"
            StrSql = StrSql & ",null"
            StrSql = StrSql & ",null"
            StrSql = StrSql & ",null"
            StrSql = StrSql & ",0"
            StrSql = StrSql & ",0"
            StrSql = StrSql & ",null"
            StrSql = StrSql & ",0"
            StrSql = StrSql & ",0"
            StrSql = StrSql & ",0"
            StrSql = StrSql & ",''"
            StrSql = StrSql & ",0"
            StrSql = StrSql & ",0"
            StrSql = StrSql & ",0"
            StrSql = StrSql & ",null"
            StrSql = StrSql & ",0"
            StrSql = StrSql & ",null"
            StrSql = StrSql & ",0"
            StrSql = StrSql & ",null"
            StrSql = StrSql & ",null"
            StrSql = StrSql & ",null"
            StrSql = StrSql & ",null"
            StrSql = StrSql & ",0"
            StrSql = StrSql & ")"
            
            objConn.Execute StrSql, , adExecuteNoRecords
            Flog.writeline ""
            Flog.writeline "No existen empleados a procesar"
        Exit Sub
    End If
    
End If
Flog.writeline ""
Flog.writeline "Proceso finalizado."



If rs.State = adStateOpen Then rs.Close
Set rs = Nothing

End Sub




Function descripcionEstructura(ByVal Tenro As Long, ByVal estrnro As Long)
'BUSCA DESCRIPCION DE ESTRUCTURA
    'mdz
    Dim StrSql As String
    Dim rs2 As New ADODB.Recordset
    
    StrSql = "SELECT tedabr, estrdabr FROM tipoestructura INNER JOIN estructura ON tipoestructura.tenro = estructura.tenro "
    StrSql = StrSql & "WHERE tipoestructura.tenro = " & Tenro
    If estrnro <> 0 Then
        StrSql = StrSql & " AND estrnro = " & estrnro
    End If
    
    '01/07/2013 - MDZ - CAS-19229 - MONRESA - BUG EN EL  REPORTE DE ACUMULADO DIARIO
    'rs2.Open StrSql, objConn
    OpenRecordset StrSql, rs2
    
    If Not rs2.EOF Then
        descripcionEstructura = rs2!tedabr & ": " & rs2!estrdabr
    End If
    rs2.Close
    
    
End Function












