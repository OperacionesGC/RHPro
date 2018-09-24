Attribute VB_Name = "mdlFeriados"
Option Explicit

'Const Version = 2.01    'Version inicial
'Const FechaVersion = "15/05/2006"

'Const Version = 2.02    '
'Const FechaVersion = "16/05/2006"
'    'FGZ - No purgaba bien los feriados

'Const Version = 2.03    '
'Const FechaVersion = "26/05/2006"
'    'FGZ - mas log

'Const Version = 2.04
'Const FechaVersion = "29/05/2006"
    'FGZ - funcion SiguienteDiaHabil con ciclo infinito

Const Version = 2.05
Const FechaVersion = "31/07/2009"
    'MB - Encriptacion de string connection

'-------------------------------------------------------------------------------

Global NroProceso As Long
Dim CEmpleadosAProc As Integer
Dim CDiasAProc As Integer
Dim IncPorc As Double
Dim Progreso As Double

Global fec_proc As Integer ' 1 - Política Primer Reg.
                           ' 2 - Política Reg. del Turno
                           ' 3 - Política Ultima Reg.
Global Usa_Conv As Boolean

Dim objBTurno As New BuscarTurno
Dim objBDia As New BuscarDia
Dim objFeriado As New Feriado
Dim objFechasHoras As New FechasHoras

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
Global Nro_Turno As Long
Global Tipo_Turno As Integer

Global Tiene_Justif As Boolean
Global nro_justif As Long
Global justif_turno As Boolean
Global p_turcomp As Boolean
Global nro_grupo As Long
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

Global aux_Tipohora As Integer
Global aux_TipoDia As Integer
Global Sigo_Generando As Boolean

Global hora_tol As String
Global fecha_tol As Date
Global hora_toldto As String
Global fecha_toldto As Date

Global E1 As String
Global E2 As String
Global E3 As String
Global S1 As String
Global S2 As String
Global S3 As String
Global FE1 As Date
Global FE2 As Date
Global FE3 As Date
Global FS1 As Date
Global FS2 As Date
Global FS3 As Date

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
Global Tipo_Hora As Integer
Global HuboErrores As Boolean
Global SinError As Boolean

Public Sub Main()

Dim Fecha As Date
Dim Ternro As Long
Dim TipoFeriado As Integer
Dim DiaCompleto As Boolean
Dim FerHoraDesde As String
Dim FerHoraHasta As String

Dim Legajo As Long
Dim objReg As New ADODB.Recordset
Dim strCmdLine As String
Dim objconnMain As New ADODB.Connection
Dim Archivo As String

Dim rs_Batch_Proceso As New ADODB.Recordset
Dim rs_His_Batch_Proceso As New ADODB.Recordset
Dim PID As String
Dim ArrParametros

'    strCmdLine = Command()
'    ArrParametros = Split(strCmdLine, " ", -1)
'    If UBound(ArrParametros) > 0 Then
'        If IsNumeric(ArrParametros(0)) Then
'            NroProceso = ArrParametros(0)
'            Etiqueta = ArrParametros(1)
'        Else
'            Exit Sub
'        End If
'    Else
'        If IsNumeric(strCmdLine) Then
'            NroProceso = strCmdLine
'        Else
'            Exit Sub
'        End If
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
    
    ' Creo el archivo de texto del desglose
    Archivo = PathFLog & "Feriados" & "-" & NroProceso & ".log"
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Archivo, True)
    
    'Abro la conexion
    On Error Resume Next
'    'OpenConnection strconexion, objConn
'    If Err.Number <> 0 Then
'        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
'        Exit Sub
'    End If
'    On Error Resume Next
'    OpenConnection strconexion, objConnProgreso
'    If Err.Number <> 0 Then
'        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
'        Exit Sub
'    End If
    
    
    'Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaVersion
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-------------------------------------------------"
    Flog.writeline
    
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
    
    ' pongo el estado del proceso en PROCESANDO
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Levanta Proceso y Setea Parámetros:  " & " " & Now
       
    Set objFechasHoras.Conexion = objConn
        
    
    'Activo el manejador de errores
    On Error GoTo CE

    StrSql = "SELECT fer.tipferinro " & _
             "FROM   batch_feriado bf, feriado fer " & _
             "WHERE  bf.bpronro  = " & NroProceso & " " & _
             "AND    fer.ferinro = bf.ferinro"
    OpenRecordset StrSql, objReg
    If Not objReg.EOF Then
      TipoFeriado = objReg!tipferinro
    End If
    If objReg.State = adStateOpen Then objReg.Close
    Set objReg = Nothing
' Se recupera el tipo de feriado que se debe procesar

    If (TipoFeriado = 2) Then
      StrSql = "SELECT be.ternro," & _
                      "Fer.ferifecha," & _
                      "fer.tipferinro," & _
                      "fer.fericompleto," & _
                      "fer.ferihoradesde," & _
                      "fer.ferihorahasta " & _
               "FROM   batch_proceso bp " & _
               "INNER JOIN batch_feriado  bf ON  bf.bpronro = bp.bpronro " & _
               "INNER JOIN batch_empleado be ON  be.bpronro = bp.bpronro " & _
               "INNER JOIN Feriado       fer ON  bf.ferinro = Fer.ferinro " & _
               "INNER JOIN fer_estr       fe ON  fe.ferinro = fer.ferinro " & _
               "INNER JOIN his_estructura he ON  he.tenro   = fe.tenro " & _
                                           "AND  he.estrnro = fe.estrnro " & _
                                           "AND  he.ternro  = be.ternro " & _
                                           "AND (he.htethasta IS NULL OR " & _
                                                "he.htethasta <= fer.ferifecha) " & _
               "WHERE bp.bpronro = " & NroProceso & " " & _
               "GROUP BY be.ternro," & _
                        "Fer.ferifecha," & _
                        "fer.tipferinro," & _
                        "fer.fericompleto," & _
                        "fer.ferihoradesde," & _
                        "fer.ferihorahasta " & _
               "ORDER BY be.ternro, Fer.ferifecha "
    Else
      StrSql = " SELECT batch_empleado.ternro, Feriado.ferifecha, feriado.tipferinro, feriado.fericompleto, feriado.ferihoradesde, feriado.ferihorahasta FROM batch_proceso " & _
               " INNER JOIN batch_feriado ON batch_feriado.bpronro = batch_proceso.bpronro " & _
               " INNER JOIN batch_empleado ON batch_empleado.bpronro = batch_proceso.bpronro" & _
               " INNER JOIN Feriado ON batch_feriado.ferinro = Feriado.ferinro" & _
               " WHERE batch_proceso.bpronro = " & NroProceso & _
               " ORDER BY batch_empleado.ternro, Feriado.ferifecha"
    End If
         
    OpenRecordset StrSql, objReg
    
    
    CEmpleadosAProc = objReg.RecordCount
    IncPorc = (100 / CEmpleadosAProc)
    
    SinError = True
    HuboErrores = False
    Do While Not objReg.EOF
   
        Ternro = objReg!Ternro
        Fecha = objReg!ferifecha
        TipoFeriado = objReg!tipferinro
        DiaCompleto = CBool(objReg!fericompleto)
                
        FerHoraDesde = "0000"
        FerHoraHasta = "0000"
                
        If Not CBool(objReg!fericompleto) Then
            FerHoraDesde = objReg!ferihoradesde
            FerHoraHasta = objReg!ferihorahasta
        End If
        
        
        Flog.writeline "Inicio Empleado:" & Ternro & " Feriado del " & Fecha
        
        'por si tienen un tratamiento especial segun el tipo de feriado
        ' por como se definió la tarea hasta el momento, todos los tipos de feriados
        ' tienen el mismo tratamiento
'        Select Case TipoFeriado
'        Case 1: 'Feriado Nacional
            Call BuscarFeriados(Ternro, Fecha, TipoFeriado, DiaCompleto, FerHoraDesde, FerHoraHasta)
'        Case 2: 'Feriado por convenio
'            Call BuscarFeriadosPorConvenio(Ternro, Fecha, TipoFeriado, DiaCompleto, FerHoraDesde, FerHoraHasta)
'        Case Else
'
'        End Select

siguiente:
        Progreso = Progreso + IncPorc
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & CInt(Progreso) & " WHERE bpronro = " & NroProceso
        objConnProgreso.Execute StrSql, , adExecuteNoRecords
            
        objReg.MoveNext
        
        If Not objReg.EOF Then
            If Not Ternro = objReg!Ternro Then
                Flog.writeline "Fin Empleado:" & Ternro & " Feriado del " & Fecha
               ' cambio el empleado
               Flog.writeline "Sin Error: " & SinError
               If SinError Then
                    ' borro
                    StrSql = "DELETE FROM batch_empleado WHERE ternro = " & Ternro & " AND bpronro = " & NroProceso
                    objConn.Execute StrSql, , adExecuteNoRecords
               Else
                    StrSql = "UPDATE batch_empleado SET estado = 'Error' WHERE ternro = " & Ternro & " AND bpronro = " & NroProceso
                    objConn.Execute StrSql, , adExecuteNoRecords
               End If
               SinError = True
            End If
        Else
            Flog.writeline "Fin Empleado:" & Ternro & " Feriado del " & Fecha
            If SinError Then
                 ' borro
                 StrSql = "DELETE FROM batch_empleado WHERE ternro = " & Ternro & " AND bpronro = " & NroProceso
                 objConn.Execute StrSql, , adExecuteNoRecords
            Else
                 StrSql = "UPDATE batch_empleado SET estado = 'Error' WHERE ternro = " & Ternro & " AND bpronro = " & NroProceso
                 objConn.Execute StrSql, , adExecuteNoRecords
            End If
        End If
    Loop


'Deshabilito el manejador de errores
On Error GoTo 0

Flog.writeline "Fin :" & Now
Flog.Close
   
    If HuboErrores Then
        ' actualizo el estado del proceso a Error
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProceso
        objConn.Execute StrSql, , adExecuteNoRecords
        
        ' Actualizo el feriado como no procesado
        StrSql = "UPDATE Batch_feriado SET bfestado = 'Error' ,bfprocesado = -1 WHERE bpronro = " & NroProceso
        objConn.Execute StrSql, , adExecuteNoRecords
    Else
        'StrSql = "UPDATE Batch_feriado SET bfestado = 'Ok' ,bfprocesado = -1 WHERE bpronro = " & NroProceso
        'objConn.Execute StrSql, , adExecuteNoRecords
        ' FGZ - 23/09/2003
        '
        StrSql = "DELETE FROM Batch_feriado WHERE bpronro = " & NroProceso
        objConn.Execute StrSql, , adExecuteNoRecords
        
        
        ' poner el bprcestado en procesado
        StrSql = "UPDATE batch_proceso SET bprcprogreso = 100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
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
    End If
        
Fin:
    objConn.Close
    Set objConn = Nothing
    If objReg.State = adStateOpen Then objReg.Close
    Set objReg = Nothing
    
    Exit Sub
    
    
CE:
    'Error handler
    HuboErrores = True
    SinError = False
    
    Flog.writeline "Error procesando Empleado:" & Ternro & " " & Fecha
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl: " & StrSql
    GoTo siguiente
    
End Sub


Public Sub Main_old()

Dim Fecha As Date
Dim Ternro As Long
Dim TipoFeriado As Integer
Dim DiaCompleto As Boolean
Dim FerHoraDesde As String
Dim FerHoraHasta As String

Dim Legajo As Long
Dim objReg As New ADODB.Recordset
Dim strCmdLine As String
Dim objconnMain As New ADODB.Connection
Dim Archivo As String

Dim rs_Batch_Proceso As New ADODB.Recordset
Dim rs_His_Batch_Proceso As New ADODB.Recordset
Dim PID As String

' carga las configuraciones basicas, formato de fecha, string de conexion,
' tipo de BD y ubicacion del archivo de log
Call CargarConfiguracionesBasicas

    strCmdLine = Command()
    'strCmdLine = "78"
    
    If IsNumeric(strCmdLine) Then
        NroProceso = strCmdLine
    Else
        Exit Sub
    End If
    
    
    ' Creo el archivo de texto del desglose
    Archivo = PathFLog & "Feriados" & "-" & NroProceso & ".log"
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Archivo, True)
    
    OpenConnection strconexion, objConn
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "PID = " & PID
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprcpid = " & PID & ", bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Levanta Proceso y Setea Parámetros:  " & " " & Now
       
    Set objFechasHoras.Conexion = objConn
        
    
    'Activo el manejador de errores
    On Error GoTo CE
          
        
    StrSql = " SELECT batch_empleado.ternro, Feriado.ferifecha, feriado.tipferinro, feriado.fericompleto, feriado.ferihoradesde, feriado.ferihorahasta FROM batch_proceso " & _
             " INNER JOIN batch_feriado ON batch_feriado.bpronro = batch_proceso.bpronro " & _
             " INNER JOIN batch_empleado ON batch_empleado.bpronro = batch_proceso.bpronro" & _
             " INNER JOIN Feriado ON batch_feriado.ferinro = Feriado.ferinro" & _
             " WHERE batch_proceso.bpronro = " & NroProceso & _
             " ORDER BY batch_empleado.ternro, Feriado.ferifecha"
             
    OpenRecordset StrSql, objReg
    
    
    CEmpleadosAProc = objReg.RecordCount
    IncPorc = (100 / CEmpleadosAProc)
    
    SinError = True
    HuboErrores = False
    Do While Not objReg.EOF
   
        Ternro = objReg!Ternro
        Fecha = objReg!ferifecha
        TipoFeriado = objReg!tipferinro
        DiaCompleto = CBool(objReg!fericompleto)
                
        FerHoraDesde = "0000"
        FerHoraHasta = "0000"
                
        If Not CBool(objReg!fericompleto) Then
            FerHoraDesde = objReg!ferihoradesde
            FerHoraHasta = objReg!ferihorahasta
        End If
        
        
        Flog.writeline "Inicio Empleado:" & Ternro & " Feriado del " & Fecha
        
        'por si tienen un tratamiento especial segun el tipo de feriado
        ' por como se definió la tarea hasta el momento, todos los tipos de feriados
        ' tienen el mismo tratamiento
'        Select Case TipoFeriado
'        Case 1: 'Feriado Nacional
            Call BuscarFeriados(Ternro, Fecha, TipoFeriado, DiaCompleto, FerHoraDesde, FerHoraHasta)
'        Case 2: 'Feriado por convenio
'            Call BuscarFeriadosPorConvenio(Ternro, Fecha, TipoFeriado, DiaCompleto, FerHoraDesde, FerHoraHasta)
'        Case Else
'
'        End Select

siguiente:
        Progreso = Progreso + IncPorc
            
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & CInt(Progreso) & " WHERE bpronro = " & NroProceso
        objConn.Execute StrSql, , adExecuteNoRecords
            
        objReg.MoveNext
        
        If Not objReg.EOF Then
            If Not Ternro = objReg!Ternro Then
                Flog.writeline "Fin Empleado:" & Ternro & " Feriado del " & Fecha
               ' cambio el empleado
               Flog.writeline "Sin Error: " & SinError
               If SinError Then
                    ' borro
                    StrSql = "DELETE FROM batch_empleado WHERE ternro = " & Ternro & " AND bpronro = " & NroProceso
                    objConn.Execute StrSql, , adExecuteNoRecords
               Else
                    StrSql = "UPDATE batch_empleado SET estado = 'Error' WHERE ternro = " & Ternro & " AND bpronro = " & NroProceso
                    objConn.Execute StrSql, , adExecuteNoRecords
               End If
               SinError = True
            End If
        Else
            Flog.writeline "Fin Empleado:" & Ternro & " Feriado del " & Fecha
            If SinError Then
                 ' borro
                 StrSql = "DELETE FROM batch_empleado WHERE ternro = " & Ternro & " AND bpronro = " & NroProceso
                 objConn.Execute StrSql, , adExecuteNoRecords
            Else
                 StrSql = "UPDATE batch_empleado SET estado = 'Error' WHERE ternro = " & Ternro & " AND bpronro = " & NroProceso
                 objConn.Execute StrSql, , adExecuteNoRecords
            End If
        End If
    Loop


'Deshabilito el manejador de errores
On Error GoTo 0

Flog.writeline "Fin :" & Now
Flog.Close
   
    If HuboErrores Then
        ' actualizo el estado del proceso a Error
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProceso
        objConn.Execute StrSql, , adExecuteNoRecords
        
        ' Actualizo el feriado como no procesado
        StrSql = "UPDATE Batch_feriado SET bfestado = 'Error' ,bfprocesado = -1 WHERE bpronro = " & NroProceso
        objConn.Execute StrSql, , adExecuteNoRecords
    Else
        'StrSql = "UPDATE Batch_feriado SET bfestado = 'Ok' ,bfprocesado = -1 WHERE bpronro = " & NroProceso
        'objConn.Execute StrSql, , adExecuteNoRecords
        ' FGZ - 23/09/2003
        '
        StrSql = "DELETE FROM Batch_feriado WHERE bpronro = " & NroProceso
        objConn.Execute StrSql, , adExecuteNoRecords
        
        
        ' poner el bprcestado en procesado
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
    End If
        
Fin:
    objConn.Close
    Set objConn = Nothing
    If objReg.State = adStateOpen Then objReg.Close
    Set objReg = Nothing
    
    Exit Sub
    
    
CE:
    'Error handler
    HuboErrores = True
    SinError = False
    
    Flog.writeline "Error procesando Empleado:" & Ternro & " " & Fecha
    
    GoTo siguiente
    
End Sub

Private Sub PurgarFeriado(Ternro As Long, Fecha As Date)
Dim HoraFeriadoNacionalporConvenio As Long
Dim ThNroFranco As Long
Dim ThNroNOFranco As Long
Dim rs As New ADODB.Recordset


    Flog.writeline "Purga de Feriado Empleado:" & Ternro & " " & Fecha
    
    'busco el tipo de hora franco Nacional
    HoraFeriadoNacionalporConvenio = 41
    StrSql = "SELECT * FROM gti_config_tur_hor WHERE conhornro = " & HoraFeriadoNacionalporConvenio
    StrSql = StrSql & " AND turnro = " & Nro_Turno & " ORDER BY conhornro ASC, turnro ASC"
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
         ThNroFranco = rs!thnro
    Else
        Flog.writeline "Error, no esta configurado el tipo de hora " & HoraFeriadoNacionalporConvenio & " " & Ternro & " " & Fecha
        Flog.writeline "No se depura"
    End If

    'Tipo de Hora "NO Corresponde Feriado"
    HoraFeriadoNacionalporConvenio = 43
    'busco el tipo de hora franco Nacional
    StrSql = "SELECT * FROM gti_config_tur_hor WHERE conhornro = " & HoraFeriadoNacionalporConvenio
    StrSql = StrSql & " AND turnro = " & Nro_Turno & " ORDER BY conhornro ASC, turnro ASC"
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        ThNroNOFranco = rs!thnro
    Else
        Flog.writeline "Error, no esta configurado el tipo de hora " & HoraFeriadoNacionalporConvenio & " Tecero: " & Ternro & " " & Fecha
        Flog.writeline "No se depura"
    End If


    'borro el feriado a reprocesar
    'FGZ - Borro de Horario Cumplido
    StrSql = "DELETE FROM gti_horcumplido WHERE ternro = " & Ternro & " AND horfecrep = " & ConvFecha(Fecha) & " AND hormanual = 0 AND (thnro = " & ThNroFranco & " OR thnro = " & ThNroNOFranco & " )"
    objConn.Execute StrSql, , adExecuteNoRecords
        
'    StrSql = " DELETE FROM gti_acumdiario WHERE adfecha = " & ConvFecha(Fecha) & " AND ternro = " & Ternro & _
'    "AND thnro = " & ThNroFranco & "AND admanual = " & CInt(False)
'    objConn.Execute StrSql, , adExecuteNoRecords
    
    
    Flog.writeline "Fin Purga"

    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
End Sub


Private Sub BuscarFeriados(Ternro As Long, Fecha As Date, TipoFeriado As Integer, DiaCompleto As Boolean, FerHoraDesde As String, FerHoraHasta As String)
' a)El empleado ha trabajado al menos 48 horas ó 6 de los ultimos 10 dias habiles al feriado.
' b)El empleado ha trabajado el dia habil anterior al feriado y al menos 1 dia mas
'       dentro de los 5 dias habiles posteriores.

'OBS:
'hay que tener en cuenta que las licencias pagas por ley,
'   enfermedad, accidente, etc. se tienen que tomar como trabajado
'salvo el caso de que la licencia se superponga sobre el dia del feriado
' que ahi si no corresponde el feriado.

Dim FechaActual As Date
Dim Trabajo As Boolean
Dim CantDiasARevisar As Integer
Dim DiasTrabajados As Integer
Dim hora_desde As String
Dim hora_hasta As String

Dim esFeriado As Boolean
Dim DHabiles(1 To 7) As Boolean
Dim ExcluyeFeriados As Boolean

    'FGZ - 20/07/2005
    'hay que contemplar que los dias sean habiles
    DHabiles(1) = False
    DHabiles(2) = True
    DHabiles(3) = True
    DHabiles(4) = True
    DHabiles(5) = True
    DHabiles(6) = True
    DHabiles(7) = False
    ExcluyeFeriados = False


' seteo el turno y el dia
    Set objBTurno.Conexion = objConn
    'Set objBTurno.ConexionTraza = CnTraza
    objBTurno.Buscar_Turno Fecha, Ternro, depurar
    initVariablesTurno objBTurno
    
    If Not tiene_turno Then
        SinError = False
        HuboErrores = True
        Flog.writeline "Empleado sin Turno: " & Ternro & " " & Fecha
        Exit Sub
    Else
        Flog.writeline "Turno del Empleado: " & Nro_Turno
    End If
    
    If tiene_turno Then
         Set objBDia.Conexion = objConn
         Set objBDia.ConexionTraza = CnTraza
         objBDia.Buscar_Dia Fecha, Fecha_Inicio, Nro_Turno, Ternro, P_Asignacion, depurar
         initVariablesDia objBDia
    End If


' lo que primero tengo que hacer es limpiar los feriados que estoy procesando
' Busco si ya tiene generado ese feriado y si lo tiene, entonces lo borra.
Call PurgarFeriado(Ternro, Fecha)

' reviso que el feriado sea feriado para este empleado
Set objFeriado.Conexion = objConn
depurar = False
If Not objFeriado.Feriado(Fecha, Ternro, depurar) Then Exit Sub

' a) El empleado ha trabajado al menos 6 de los ultimos 10 dias habiles al feriado.
Flog.writeline "Evaluo condicion a) El empleado ha trabajado al menos 6 de los ultimos 10 dias habiles al feriado."
CantDiasARevisar = 11
FechaActual = AnteriorDiaHabil(Fecha, Ternro)
DiasTrabajados = 0
hora_desde = "0000"
hora_hasta = "2359"

Do While CantDiasARevisar > 0 And DiasTrabajados < 6
    esFeriado = objFeriado.Feriado(FechaActual, Ternro, depurar)
    If Not (esFeriado And Not ExcluyeFeriados) Then
        If DHabiles(Weekday(FechaActual)) Then
            ' si existen registraciones en ese dia, entonces trabajo
            If Existe_Registracion(Ternro, FechaActual, hora_desde, FechaActual, hora_hasta) Then
                DiasTrabajados = DiasTrabajados + 1
            End If
        
            CantDiasARevisar = CantDiasARevisar - 1
        End If
    End If
    FechaActual = AnteriorDiaHabil(FechaActual, Ternro)
Loop

Flog.writeline "Dias Trabajados: " & DiasTrabajados
' si se cumple la condicion a)
If DiasTrabajados >= 6 Then
    Call InsertarFranco(Ternro, Fecha)
    Flog.writeline "Inserto Feriado: " & Ternro & " " & Fecha
    Exit Sub
End If


Flog.writeline
Flog.writeline "Evaluo condicion b)El empleado ha trabajado el dia habil anterior al feriado y al menos 1 dia mas"
Flog.writeline "                     dentro de los 5 dias habiles posteriores."
' b)El empleado ha trabajado el dia habil anterior al feriado y al menos 1 dia mas
'       dentro de los 5 dias habiles posteriores.
CantDiasARevisar = 6
FechaActual = AnteriorDiaHabil(Fecha, Ternro)
Trabajo = False

' si trabajó el dia habil anterior
If Existe_Registracion(Ternro, FechaActual, hora_desde, FechaActual, hora_hasta) Then
    ' entonces reviso si trabajo al menos un dia mas en los cinco dias posteriores
    FechaActual = SiguienteDiaHabil(Fecha, Ternro)
    
    Do While CantDiasARevisar > 0 And Not Trabajo
        esFeriado = objFeriado.Feriado(FechaActual, Ternro, depurar)
        If Not (esFeriado And Not ExcluyeFeriados) Then
            If DHabiles(Weekday(FechaActual)) Then
                If Existe_Registracion(Ternro, FechaActual, hora_desde, FechaActual, hora_hasta) Then
                    Trabajo = True
                End If
            
                CantDiasARevisar = CantDiasARevisar - 1
            End If
        End If
        FechaActual = SiguienteDiaHabil(FechaActual, Ternro)
    Loop
End If

Flog.writeline
Flog.writeline "Trabajó: " & Trabajo
If Trabajo Then
    Call InsertarFranco(Ternro, Fecha)
    Flog.writeline "Inserto Feriado: " & Ternro & " " & Fecha
Else
    Flog.writeline "No cumplio con las condiciones ==> GENERA HORAS NO CORRESPONDE FERIADO: " & Ternro & " " & Fecha
    Call InsertarNOFranco(Ternro, Fecha)
    Flog.writeline "Inserto No Corresponde Feriado: " & Ternro & " " & Fecha
End If

'Inserto dia procesado
Flog.writeline "Inserto dia procesado"
Call Insertar_GTI_Proc_Emp(Ternro, Fecha)
Flog.writeline "Fin"

End Sub


Public Sub Insertar_GTI_Proc_Emp(ByVal Ternro As Long, ByVal Fecha As Date)
' --------------------------------------------------------------
' Descripcion: Genera la informacion del dia procesado.
' Autor: FGZ - 27/10/2005
' Ultima modificacion:
' --------------------------------------------------------------
Dim rs_gti_Proc_Emp As New ADODB.Recordset

esFeriado = True
Trabaja = False

        On Error GoTo ME_Local
        
        StrSql = "SELECT * FROM gti_proc_emp WHERE ternro =" & Ternro
        StrSql = StrSql & " AND fecha = " & ConvFecha(Fecha)
        OpenRecordset StrSql, rs_gti_Proc_Emp
        If rs_gti_Proc_Emp.EOF Then
            StrSql = "INSERT INTO gti_proc_emp (ternro,fecha,turnro,fpgnro,dianro,feriado,jusnro,pasig,dialibre,trabaja,estrnro"
            If Horario_Movil Then
                StrSql = StrSql & ",manual "
                If Not EsNulo(E1) Then
                    StrSql = StrSql & ",horadesde1, horahasta1 "
                End If
                If Not EsNulo(E2) Then
                    StrSql = StrSql & ",horadesde2, horahasta2 "
                End If
                If Not EsNulo(E3) Then
                    StrSql = StrSql & ",horadesde3, horahasta3 "
                End If
            End If
            StrSql = StrSql & " ) VALUES ("
            StrSql = StrSql & Ternro & ","
            StrSql = StrSql & ConvFecha(Fecha) & ","
            StrSql = StrSql & Nro_Turno & ","
            StrSql = StrSql & Nro_fpgo & ","
            StrSql = StrSql & Nro_Dia & ","
            StrSql = StrSql & CInt(esFeriado) & ","
            StrSql = StrSql & nro_justif & ","
            StrSql = StrSql & CInt(P_Asignacion) & ","
            StrSql = StrSql & CInt(Dia_Libre) & ","
            StrSql = StrSql & CInt(Trabaja) & ","
            StrSql = StrSql & nro_grupo
            If Horario_Movil Then
                StrSql = StrSql & ",-1 "
                If Not EsNulo(E1) Then
                    StrSql = StrSql & ",'" & E1 & "','" & S1 & "'"
                End If
                If Not EsNulo(E2) Then
                    StrSql = StrSql & ",'" & E2 & "','" & S2 & "'"
                End If
                If Not EsNulo(E3) Then
                    StrSql = StrSql & ",'" & E3 & "','" & S3 & "'"
                End If
            End If
            StrSql = StrSql & ")"
        Else
            StrSql = "UPDATE gti_proc_emp SET "
            StrSql = StrSql & " turnro = " & Nro_Turno
            StrSql = StrSql & ",fpgnro = " & Nro_fpgo
            StrSql = StrSql & ",dianro = " & Nro_Dia
            StrSql = StrSql & ",feriado = " & CInt(esFeriado)
            StrSql = StrSql & ",jusnro = " & nro_justif
            StrSql = StrSql & ",pasig = " & CInt(P_Asignacion)
            StrSql = StrSql & ",dialibre = " & CInt(Dia_Libre)
            StrSql = StrSql & ",trabaja = " & CInt(Trabaja)
            StrSql = StrSql & ",estrnro = " & nro_grupo
            If Horario_Movil Then
                StrSql = StrSql & ",manual = -1 "
                If Not EsNulo(E1) Then
                    StrSql = StrSql & ",horadesde1 ='" & E1 & "',horahasta1 ='" & S1 & "'"
                End If
                If Not EsNulo(E2) Then
                    StrSql = StrSql & ",horadesde2 ='" & E2 & "',horahasta2 ='" & S2 & "'"
                End If
                If Not EsNulo(E3) Then
                    StrSql = StrSql & ",horadesde3 ='" & E3 & "',horahasta3 ='" & S3 & "'"
                End If
            End If
            StrSql = StrSql & " WHERE TERNRO =" & Ternro
            StrSql = StrSql & " AND fecha = " & ConvFecha(Fecha)
        End If
        objConn.Execute StrSql, , adExecuteNoRecords
        
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
    Flog.writeline "Error generando gti_proc_emp. La informacion del horario teorico en el tablero no estará disponible."
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "SQL: " & StrSql
    Flog.writeline " ---------------------------------------------------------------------------------------------------"
    Flog.writeline "***"
    Flog.writeline
End Sub




Private Sub BuscarFeriadosPorConvenio(Ternro As Long, Fecha As Date, TipoFeriado As Integer, DiaCompleto As Boolean, FerHoraDesde As String, FerHoraHasta As String)
' a)El empleado ha trabajado al menos 6 de los ultimos 10 dias habiles al feriado.
' b)El empleado ha trabajado el dia habil anterior al feriado y al menos 1 dia mas
'       dentro de los 5 dias habiles posteriores.

Dim FechaActual As Date
Dim Trabajo As Boolean
Dim CantDiasARevisar As Integer
Dim DiasTrabajados As Integer
Dim hora_desde As String
Dim hora_hasta As String


' seteo el turno y el dia
    Set objBTurno.Conexion = objConn
    'Set objBTurno.ConexionTraza = CnTraza
    objBTurno.Buscar_Turno Fecha, Ternro, depurar
    initVariablesTurno objBTurno
    
    If Not tiene_turno Then
        Exit Sub
    Else
        Flog.writeline "Turno del Empleado: " & Nro_Turno
    End If
    
    If tiene_turno Then
        SinError = False
        HuboErrores = True
         Set objBDia.Conexion = objConn
         Set objBDia.ConexionTraza = CnTraza
         objBDia.Buscar_Dia Fecha, Fecha_Inicio, Nro_Turno, Ternro, P_Asignacion, depurar
         initVariablesDia objBDia
    End If

' lo que primero tengo que hacer es limpiar los feriados que estoy procesando
' Busco si ya tiene generado ese feriado y si lo tiene, entonces lo borra.
Call PurgarFeriado(Ternro, Fecha)


' reviso que el feriado sea feriado para este empleado
Set objFeriado.Conexion = objConn
depurar = False
If Not objFeriado.Feriado(Fecha, Ternro, depurar) Then Exit Sub

' a) El empleado ha trabajado al menos 6 de los ultimos 10 dias habiles al feriado.
CantDiasARevisar = 11
FechaActual = AnteriorDiaHabil(Fecha, Ternro)
DiasTrabajados = 0
hora_desde = "0000"
hora_hasta = "2359"

Do While CantDiasARevisar > 0 And DiasTrabajados < 6
    ' si existen registraciones en ese dia, entonces trabajo
    If Existe_Registracion(Ternro, FechaActual, hora_desde, FechaActual, hora_hasta) Then
        DiasTrabajados = DiasTrabajados + 1
    End If

    FechaActual = AnteriorDiaHabil(FechaActual, Ternro)
    CantDiasARevisar = CantDiasARevisar - 1
Loop

' si se cumple la condicion a)
If DiasTrabajados >= 6 Then
    Call InsertarFranco(Ternro, Fecha)
    Exit Sub
End If


' b)El empleado ha trabajado el dia habil anterior al feriado y al menos 1 dia mas
'       dentro de los 5 dias habiles posteriores.
CantDiasARevisar = 6
FechaActual = AnteriorDiaHabil(Fecha, Ternro)
Trabajo = False

' si trabajó el dia habil anterior
If Existe_Registracion(Ternro, FechaActual, hora_desde, FechaActual, hora_hasta) Then
    ' entonces reviso si trabajo al menos un dia mas en los cinco dias posteriores
    FechaActual = SiguienteDiaHabil(Fecha, Ternro)
    
    Do While CantDiasARevisar > 0 And Not Trabajo
        If Existe_Registracion(Ternro, FechaActual, hora_desde, FechaActual, hora_hasta) Then
            Trabajo = True
        End If
    
        FechaActual = SiguienteDiaHabil(FechaActual, Ternro)
        CantDiasARevisar = CantDiasARevisar - 1
    Loop

    If Trabajo Then
        Call InsertarFranco(Ternro, Fecha)
    End If

End If

End Sub




Private Function Existe_Registracion(Ternro As Long, fecha_desde As Date, hora_desde As String, fecha_hasta As Date, hora_hasta As String) As Boolean
Dim result As Boolean
Dim Continuar As Boolean
Dim salir As Boolean
Dim objRs As New ADODB.Recordset
Dim rs_ExisteLic As New ADODB.Recordset
Dim Lista As String

    Flog.writeline "Existe_Registracion Desde: " & fecha_desde & " " & hora_desde & " hasta: " & fecha_hasta & " " & hora_hasta
    salir = False
    Existe_Registracion = False
    result = False
    StrSql = "SELECT * FROM gti_registracion WHERE ternro = " & Ternro & " AND  regfecha >= " & ConvFecha(fecha_desde) & " AND " & _
             "regfecha <=" & ConvFecha(fecha_hasta)
    OpenRecordset StrSql, objRs
    Do While Not salir And Not objRs.EOF
        Continuar = True
        If objRs!regfecha = fecha_desde Then
            If objRs!reghora < hora_desde Then
                objRs.MoveNext
                Continuar = False
            End If
        End If
        
        If Continuar Then
            If (objRs!regfecha = fecha_hasta) And (Continuar) Then
                If objRs!reghora > hora_hasta Then
                    Continuar = False
                    salir = True
                End If
            End If
        End If
        
        If (Continuar) Then result = True
        If Not objRs.EOF Then objRs.MoveNext
    Loop
    
    Lista = "2,3,4,5,7,8,9,19,22,23,28,30,31,32,34"
    If Not result Then
        StrSql = " SELECT * FROM emp_lic WHERE (empleado = " & Ternro
        StrSql = StrSql & " ) AND (tdnro IN (" & Lista & "))"
        StrSql = StrSql & " AND eltipo = 1"
        StrSql = StrSql & " AND elfechadesde <= " & ConvFecha(fecha_desde)
        StrSql = StrSql & " AND elfechahasta >= " & ConvFecha(fecha_hasta)
        If rs_ExisteLic.State = adStateOpen Then rs_ExisteLic.Close
        OpenRecordset StrSql, rs_ExisteLic
        If Not rs_ExisteLic.EOF Then
            result = True
        End If
    End If
    Existe_Registracion = result
    
    Flog.writeline "Fin Existe_Registracion"
    'cierro y libero
    If objRs.State = adStateOpen Then objRs.Close
    Set objRs = Nothing
    
    If rs_ExisteLic.State = adStateOpen Then rs_ExisteLic.Close
    Set rs_ExisteLic = Nothing
End Function

Private Sub InsertarFranco(Ternro As Long, Fecha As Date)
    ' obtengo el turno del empleado, busco el tipo de hora franco nacional e inserto en
    ' el acumulado diario por 8 horas (en principio)
Dim rs As New ADODB.Recordset
Dim rsEsta As New ADODB.Recordset
Dim ThNroFranco As Long
Dim canthoras As Single
Dim HoraFeriadoNacionalporConvenio As Long
Dim Rs_HorasTurno As New ADODB.Recordset

    Flog.writeline
    Flog.writeline "Inserto el franco"
    ' de las configuraciones basicas
    HoraFeriadoNacionalporConvenio = 41
    'ThNroFranco = HoraFeriadoNacionalporConvenio
    
'     ' busco el tipo de hora franco Nacional
     StrSql = "SELECT * FROM gti_config_tur_hor WHERE conhornro = " & HoraFeriadoNacionalporConvenio & _
                " AND turnro = " & Nro_Turno & " ORDER BY conhornro ASC, turnro ASC"
    
     rs.Open StrSql, objConn

     If Not rs.EOF Then
         ThNroFranco = rs!thnro
     Else
        Flog.writeline "Error, no esta configurado el tipo de hora" & Ternro & " " & Fecha
         ' Error, no esta configurado el tipo de hora
         Exit Sub
     End If
    
    'seteo la cantidad de horas
    'FGZ - 27/07/2005
    'Busco la cantidad de horas del turno
    StrSql = "SELECT * FROM gti_dias WHERE subturnro = " & Nro_Subturno
    StrSql = StrSql & " ORDER BY diaorden"
    If Rs_HorasTurno.State = adStateOpen Then Rs_HorasTurno.Close
    OpenRecordset StrSql, Rs_HorasTurno
    If Not Rs_HorasTurno.EOF Then
        canthoras = Rs_HorasTurno!diacanthoras
    Else
        canthoras = 8
    End If
    
    
'    'seteo la cantidad de horas
'    If Weekday(Fecha) = 7 Then ' sabado
'        canthoras = 8
'    Else
'       canthoras = 8
'    End If
    
    'FGZ - 02/09/2003 Debo insertar en el Horario Cumplido
    StrSql = "INSERT INTO gti_horcumplido(horcant,hordesde,horhasta,horestado,hormanual,horvalido,Ternro,thnro,horfecrep) " & _
            " VALUES (" & _
            canthoras & "," & _
            ConvFecha(Fecha) & "," & _
            ConvFecha(Fecha) & ",' '," & _
            CInt(False) & "," & _
            CInt(True) & "," & _
            Ternro & "," & _
            ThNroFranco & "," & _
            ConvFecha(Fecha) & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
    
'    ' si existe, lo actualizo
'    StrSql = " SELECT * FROM gti_acumdiario WHERE adfecha = " & ConvFecha(Fecha) & " AND ternro = " & Ternro & _
'    "AND thnro = " & ThNroFranco & "AND admanual = " & CInt(False)
'    rsEsta.Open StrSql, objConn
'
'    If rsEsta.EOF Then ' no existe
'        StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,adcanthoras,admanual,advalido) " & _
'        " VALUES (" & ConvFecha(Fecha) & "," & Ternro & "," & ThNroFranco & "," & canthoras & "," & _
'        CInt(False) & "," & CInt(True) & ")"
'    Else
'        StrSql = " UPDATE gti_acumdiario SET adcanthoras = " & canthoras & " WHERE adfecha = " & ConvFecha(Fecha) & " AND ternro = " & Ternro & _
'        "AND thnro = " & ThNroFranco & "AND admanual = " & CInt(False)
'    End If
'
'    objConn.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Fin Inserto el franco"
    ' cierro los rs
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    If rsEsta.State = adStateOpen Then rsEsta.Close
    Set rsEsta = Nothing
    If Rs_HorasTurno.State = adStateOpen Then Rs_HorasTurno.Close
    Set Rs_HorasTurno = Nothing
    
End Sub

Private Sub InsertarFranco_OLD(Ternro As Long, Fecha As Date)
    ' obtengo el turno del empleado, busco el tipo de hora franco nacional e inserto en
    ' el acumulado diario por 8 horas (en principio)
Dim rs As New ADODB.Recordset
Dim rsEsta As New ADODB.Recordset
Dim ThNroFranco As Long
Dim canthoras As Single
Dim HoraFeriadoNacionalporConvenio As Long

    ' de las configuraciones basicas
    HoraFeriadoNacionalporConvenio = 12
    ThNroFranco = HoraFeriadoNacionalporConvenio
    
'     ' busco el tipo de hora franco Nacional
'     StrSql = "SELECT * FROM gti_config_tur_hor WHERE conhornro = " & HoraFeriadoNacionalporConvenio & _
'     " AND turnro = " & nro_turno & " ORDER BY conhornro ASC, turnro ASC"
     
'     StrSql = "SELECT * FROM gti_config_tur_hor WHERE conhornro = " & HoraFeriadoNacionalporConvenio & _
'     " ORDER BY conhornro ASC, turnro ASC"
'     rs.Open StrSql, objConn
'
'     If Not rs.EOF Then
'         ThNroFranco = rs!thnro
'     Else
'        Flog.Writeline "Error, no esta configurado el tipo de hora" & Ternro & " " & Fecha
'         ' Error, no esta configurado el tipo de hora
'         Exit Sub
'     End If
    
    
    
    'seteo la cantidad de horas
    If Weekday(Fecha) = 7 Then ' sabado
        canthoras = 8
    Else
       canthoras = 8
    End If
    
    'FGZ - 02/09/2003 Debo insertar en el Horario Cumplido
    StrSql = "INSERT INTO gti_horcumplido(horcant,horestado,hormanual,horvalido,Ternro,thnro,horfecrep) " & _
            " VALUES (" & _
            canthoras & ",' '," & _
            CInt(False) & "," & _
            CInt(True) & "," & _
            Ternro & "," & _
            ThNroFranco & "," & _
            ConvFecha(Fecha) & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
    
'    ' si existe, lo actualizo
'    StrSql = " SELECT * FROM gti_acumdiario WHERE adfecha = " & ConvFecha(Fecha) & " AND ternro = " & Ternro & _
'    "AND thnro = " & ThNroFranco & "AND admanual = " & CInt(False)
'    rsEsta.Open StrSql, objConn
'
'    If rsEsta.EOF Then ' no existe
'        StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,adcanthoras,admanual,advalido) " & _
'        " VALUES (" & ConvFecha(Fecha) & "," & Ternro & "," & ThNroFranco & "," & canthoras & "," & _
'        CInt(False) & "," & CInt(True) & ")"
'    Else
'        StrSql = " UPDATE gti_acumdiario SET adcanthoras = " & canthoras & " WHERE adfecha = " & ConvFecha(Fecha) & " AND ternro = " & Ternro & _
'        "AND thnro = " & ThNroFranco & "AND admanual = " & CInt(False)
'    End If
'
'    objConn.Execute StrSql, , adExecuteNoRecords
    
    ' cierro los rs
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    If rsEsta.State = adStateOpen Then rsEsta.Close
    Set rsEsta = Nothing

    
End Sub

Private Sub InsertarNOFranco(Ternro As Long, Fecha As Date)
    ' obtengo el turno del empleado, busco el tipo de hora franco nacional e inserto en
    ' el acumulado diario por 8 horas (en principio)
Dim rs As New ADODB.Recordset
Dim rsEsta As New ADODB.Recordset
Dim ThNroNOFranco As Long
Dim canthoras As Single
Dim HoraFeriadoNacionalporConvenio As Long
Dim Rs_HorasTurno As New ADODB.Recordset

    ' de las configuraciones basicas
    ' En realidad este tipo deberia sacarse o de una politica (lo mas adecuado) o sino del confrep
    ' pero por razones de tiempo (debe estar para ahora, lo hago así). FGZ 03-09-2003
    
    ' Tipo de Hora "NO Corresponde Feriado"
    HoraFeriadoNacionalporConvenio = 43
    'ThNroNOFranco = HoraFeriadoNacionalporConvenio
    
'     ' busco el tipo de hora franco Nacional
     StrSql = "SELECT * FROM gti_config_tur_hor WHERE conhornro = " & HoraFeriadoNacionalporConvenio & _
                " AND turnro = " & Nro_Turno & " ORDER BY conhornro ASC, turnro ASC"
     rs.Open StrSql, objConn

     If Not rs.EOF Then
         ThNroNOFranco = rs!thnro
     Else
        Flog.writeline "Error, no esta configurado el tipo de hora" & Ternro & " " & Fecha
         ' Error, no esta configurado el tipo de hora
         Exit Sub
     End If
    
    
    'seteo la cantidad de horas
    'FGZ - 27/07/2005
    'Busco la cantidad de horas del turno
    StrSql = "SELECT * FROM gti_dias WHERE subturnro = " & Nro_Subturno
    StrSql = StrSql & " ORDER BY diaorden"
    If Rs_HorasTurno.State = adStateOpen Then Rs_HorasTurno.Close
    OpenRecordset StrSql, Rs_HorasTurno
    If Not Rs_HorasTurno.EOF Then
        canthoras = Rs_HorasTurno!diacanthoras
    Else
        canthoras = 8
    End If
    
'    'seteo la cantidad de horas
'    If Weekday(Fecha) = 7 Then ' sabado
'        canthoras = 8
'    Else
'       canthoras = 8
'    End If
    
    'FGZ - 02/09/2003 Debo insertar en el Horario Cumplido
    StrSql = "INSERT INTO gti_horcumplido(horcant,hordesde,horhasta,horestado,hormanual,horvalido,Ternro,thnro,horfecrep) " & _
            " VALUES (" & _
            canthoras & "," & _
            ConvFecha(Fecha) & "," & _
            ConvFecha(Fecha) & ",' ',0,-1," & _
            Ternro & "," & _
            ThNroNOFranco & "," & _
            ConvFecha(Fecha) & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
    
'    ' si existe, lo actualizo
'    StrSql = " SELECT * FROM gti_acumdiario WHERE adfecha = " & ConvFecha(Fecha) & " AND ternro = " & Ternro & _
'    "AND thnro = " & ThNroFranco & "AND admanual = " & CInt(False)
'    rsEsta.Open StrSql, objConn
'
'    If rsEsta.EOF Then ' no existe
'        StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,adcanthoras,admanual,advalido) " & _
'        " VALUES (" & ConvFecha(Fecha) & "," & Ternro & "," & ThNroFranco & "," & canthoras & "," & _
'        CInt(False) & "," & CInt(True) & ")"
'    Else
'        StrSql = " UPDATE gti_acumdiario SET adcanthoras = " & canthoras & " WHERE adfecha = " & ConvFecha(Fecha) & " AND ternro = " & Ternro & _
'        "AND thnro = " & ThNroFranco & "AND admanual = " & CInt(False)
'    End If
'
'    objConn.Execute StrSql, , adExecuteNoRecords
    
    ' cierro los rs
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    If rsEsta.State = adStateOpen Then rsEsta.Close
    Set rsEsta = Nothing
    If Rs_HorasTurno.State = adStateOpen Then Rs_HorasTurno.Close
    Set Rs_HorasTurno = Nothing
    
End Sub


Private Sub InsertarNOFranco_OLD(Ternro As Long, Fecha As Date)
    ' obtengo el turno del empleado, busco el tipo de hora franco nacional e inserto en
    ' el acumulado diario por 8 horas (en principio)
Dim rs As New ADODB.Recordset
Dim rsEsta As New ADODB.Recordset
Dim ThNroNOFranco As Long
Dim canthoras As Single
Dim HoraFeriadoNacionalporConvenio As Long

    ' de las configuraciones basicas
    ' En realidad este tipo deberia sacarse o de una politica (lo mas adecuado) o sino del confrep
    ' pero por razones de tiempo (debe estar para ahora, lo hago así). FGZ 03-09-2003
    
    HoraFeriadoNacionalporConvenio = 82
    ThNroNOFranco = HoraFeriadoNacionalporConvenio
    
'     ' busco el tipo de hora franco Nacional
'     StrSql = "SELECT * FROM gti_config_tur_hor WHERE conhornro = " & HoraFeriadoNacionalporConvenio & _
'     " AND turnro = " & nro_turno & " ORDER BY conhornro ASC, turnro ASC"
     
'     StrSql = "SELECT * FROM gti_config_tur_hor WHERE conhornro = " & HoraFeriadoNacionalporConvenio & _
'     " ORDER BY conhornro ASC, turnro ASC"
'     rs.Open StrSql, objConn
'
'     If Not rs.EOF Then
'         ThNroFranco = rs!thnro
'     Else
'        Flog.Writeline "Error, no esta configurado el tipo de hora" & Ternro & " " & Fecha
'         ' Error, no esta configurado el tipo de hora
'         Exit Sub
'     End If
    
    
    
    'seteo la cantidad de horas
    If Weekday(Fecha) = 7 Then ' sabado
        canthoras = 8
    Else
       canthoras = 8
    End If
    
    'FGZ - 02/09/2003 Debo insertar en el Horario Cumplido
    StrSql = "INSERT INTO gti_horcumplido(horcant,horestado,hormanual,horvalido,Ternro,thnro,horfecrep) " & _
            " VALUES (" & _
            canthoras & ",' '," & _
            CInt(False) & "," & _
            CInt(True) & "," & _
            Ternro & "," & _
            ThNroNOFranco & "," & _
            ConvFecha(Fecha) & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
    
'    ' si existe, lo actualizo
'    StrSql = " SELECT * FROM gti_acumdiario WHERE adfecha = " & ConvFecha(Fecha) & " AND ternro = " & Ternro & _
'    "AND thnro = " & ThNroFranco & "AND admanual = " & CInt(False)
'    rsEsta.Open StrSql, objConn
'
'    If rsEsta.EOF Then ' no existe
'        StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,adcanthoras,admanual,advalido) " & _
'        " VALUES (" & ConvFecha(Fecha) & "," & Ternro & "," & ThNroFranco & "," & canthoras & "," & _
'        CInt(False) & "," & CInt(True) & ")"
'    Else
'        StrSql = " UPDATE gti_acumdiario SET adcanthoras = " & canthoras & " WHERE adfecha = " & ConvFecha(Fecha) & " AND ternro = " & Ternro & _
'        "AND thnro = " & ThNroFranco & "AND admanual = " & CInt(False)
'    End If
'
'    objConn.Execute StrSql, , adExecuteNoRecords
    
    ' cierro los rs
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    If rsEsta.State = adStateOpen Then rsEsta.Close
    Set rsEsta = Nothing

    
End Sub



Private Sub initVariablesTurno(ByRef T As BuscarTurno)
   p_turcomp = T.Compensa_Turno
   nro_grupo = T.Empleado_Grupo
   nro_justif = T.Justif_Numero
   justif_turno = T.justif_turno
   Tiene_Justif = T.Tiene_Justif
   Fecha_Inicio = T.FechaInicio
   Nro_fpgo = T.Numero_FPago
   Nro_Turno = T.Turno_Numero
   tiene_turno = T.tiene_turno
   Tipo_Turno = T.Turno_Tipo
   P_Asignacion = T.Tiene_PAsignacion
End Sub

Private Sub initVariablesDia(ByRef D As BuscarDia)
   Dia_Libre = D.Dia_Libre
   Nro_Dia = D.Numero_Dia
   Nro_Subturno = D.SubTurno_Numero
   Orden_Dia = D.Orden_Dia
   Trabaja = D.Trabaja
End Sub

Private Function SiguienteDiaHabil(Fecha As Date, Ternro As Long) As Date
Dim Aux_fecha As Date
Dim DHabiles(1 To 7) As Boolean
Dim ExcluyeFeriados As Boolean

    Flog.writeline "Dia hábil siguiente a " & Fecha & " para el tercero " & Ternro
    'FGZ - 20/07/2005
    'hay que contemplar que los dias sean habiles
    DHabiles(1) = False
    DHabiles(2) = True
    DHabiles(3) = True
    DHabiles(4) = True
    DHabiles(5) = True
    DHabiles(6) = True
    DHabiles(7) = False
    ExcluyeFeriados = False

    Set objFeriado.Conexion = objConn
    'Set objBTurno.ConexionTraza = CnTraza
    depurar = False
    Aux_fecha = Fecha + 1
    
    'Do While (objFeriado.Feriado(Aux_fecha, Ternro, depurar) = True Or Not DHabiles(Weekday(Fecha)))
    Do While (objFeriado.Feriado(Aux_fecha, Ternro, depurar) = True Or Not DHabiles(Weekday(Aux_fecha)))
        Aux_fecha = Aux_fecha + 1
    Loop
    
    Flog.writeline "Fin Dia hábil siguiente " & Aux_fecha
    SiguienteDiaHabil = Aux_fecha
End Function


Private Function AnteriorDiaHabil(Fecha As Date, Ternro As Long) As Date
Dim Aux_fecha As Date
Dim DHabiles(1 To 7) As Boolean
Dim ExcluyeFeriados As Boolean

    Flog.writeline "Dia hábil anterior a " & Fecha & " para el tercero " & Ternro
    'FGZ - 20/07/2005
    'hay que contemplar que los dias sean habiles
    DHabiles(1) = False
    DHabiles(2) = True
    DHabiles(3) = True
    DHabiles(4) = True
    DHabiles(5) = True
    DHabiles(6) = True
    DHabiles(7) = False
    ExcluyeFeriados = False

    Set objFeriado.Conexion = objConn
    'Set objBTurno.ConexionTraza = CnTraza
    depurar = False
    Aux_fecha = Fecha - 1
    
    Do While (objFeriado.Feriado(Aux_fecha, Ternro, depurar) = True Or Not DHabiles(Weekday(Aux_fecha)))
        Aux_fecha = Aux_fecha - 1
    Loop
    
    Flog.writeline "Fin dia hábil anterior " & Aux_fecha
    AnteriorDiaHabil = Aux_fecha

End Function



