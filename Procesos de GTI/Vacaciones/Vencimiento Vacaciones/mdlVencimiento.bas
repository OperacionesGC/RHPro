Attribute VB_Name = "mdlVencimiento"
Option Explicit

Global Ternro As Long
Global NroProceso As Long
Dim CEmpleadosAProc As Integer
Dim CDiasAProc As Integer
Dim IncPorc As Single
Dim Progreso As Single

Global fec_proc As Integer ' 1 - Política Primer Reg.
                           ' 2 - Política Reg. del Turno
                           ' 3 - Política Ultima Reg.
Global Usa_Conv As Boolean
Global diatipo As Byte
Global ok As Boolean

Global Tdias As Integer
Global Thoras As Integer
Global Tmin As Integer
Global Cod_justificacion1 As Long
Global Cod_justificacion2 As Long

Global Existe_Reg As Boolean

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
Global dias_trabajados As Long
Global Dias_laborables As Long

Global aux_Tipohora As Integer
Global aux_TipoDia As Integer

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
Global modeloPais As Integer




Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Proceso de calculo de vencimiento de dias de vacaciones para un periodo.
' Autor      : FGZ
' Fecha      : 21/10/2009
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Fecha As Date
Dim pos1 As Integer
Dim pos2 As Integer

Dim FechaProc As Date
Dim TodosEmpleados As Boolean
Dim NroVacSiguiente As Long
Dim NroVacVence As Long

Dim objReg As New ADODB.Recordset
Dim strCmdLine As String
Dim Archivo As String

Dim rs As New ADODB.Recordset
Dim rs_Batch_Proceso As New ADODB.Recordset
Dim rs_His_Batch_Proceso As New ADODB.Recordset
Dim PID As String
Dim ArrParametros
Dim ArrPar
Dim Periodo_Anio As Long

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
    Archivo = PathFLog & "Vac_Vencimiento" & "-" & NroProceso & ".log"
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
    
    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline "Problemas con la conexión "
        Exit Sub
    End If
    
    OpenConnection strconexion, objConnProgreso
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline "Problemas con la conexión "
        Exit Sub
    End If
    
    On Error GoTo 0

    'Activo el manejador de errores
    On Error GoTo CE
    
    'FGZ - 05/08/2009 --------- Control de versiones ------
    Version_Valida = ValidarVersion(Version, 256, TipoBD)
    If Not Version_Valida Then
        'Actualizo el progreso
        MyBeginTrans
            StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Error de Version', bprcpid = " & PID & " WHERE bpronro = " & NroProceso
            objConnProgreso.Execute StrSql, , adExecuteNoRecords
        MyCommitTrans
        Flog.writeline
        GoTo Final
    End If
    'FGZ - 05/08/2009 --------- Control de versiones ------
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprcpid = " & PID & ", bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objConnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Levanta Proceso y Setea Parámetros:  " & " " & Now
       
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro =" & NroProceso
    OpenRecordset StrSql, rs_Batch_Proceso
    If rs_Batch_Proceso.EOF Then Exit Sub
    parametros = rs_Batch_Proceso!bprcparam
    
    If Not IsNull(parametros) Then
        parametros = rs_Batch_Proceso!bprcparam
        ArrPar = Split(parametros, ".")
        
        NroVac = CLng(ArrPar(0))
        Reproceso = CBool(ArrPar(1))
        
        'LM - en el proceso de colombia llega en nrovac = 0
        If NroVac <> 0 Then
            StrSql = " SELECT * FROM vacacion WHERE vacacion.vacnro = " & NroVac
            OpenRecordset StrSql, objRs
            If Not objRs.EOF Then
                Periodo_Anio = objRs!vacanio
                fecha_desde = objRs!vacfecdesde
                fecha_hasta = objRs!vacfechasta
            End If
        End If
        
        If UBound(ArrPar) > 1 Then
            FechaProc = IIf(EsNulo(ArrPar(2)), 0, ArrPar(2))
            
            If UBound(ArrPar) > 2 Then
                If EsNulo(ArrPar(3)) Then
                    TodosEmpleados = False
                Else
                    TodosEmpleados = CBool(ArrPar(3))
                End If
            Else
                FechaProc = fecha_hasta
                TodosEmpleados = False
            End If
        Else
            FechaProc = fecha_hasta
            TodosEmpleados = False
        End If
    End If
    
    'Se agrego un parametro para determinar el país con el cual se quiere calcular el modelo de vacacion
    'El parametro 7 hace referencia al modelo de vacaciones configurado en la tabla confper
    modeloPais = Pais_Modelo(7)

    Set objFechasHoras.Conexion = objConn
        
    StrSql = " SELECT empleado.ternro, empleado.empleg FROM batch_empleado "
    StrSql = StrSql & " INNER JOIN empleado ON batch_empleado.ternro = empleado.ternro"
    StrSql = StrSql & " WHERE batch_empleado.bpronro = " & NroProceso
    OpenRecordset StrSql, objReg
    
    CEmpleadosAProc = objReg.RecordCount
    If CEmpleadosAProc = 0 Then
        CEmpleadosAProc = 1
        Flog.writeline "No hay empleados para procesar."
    End If
    IncPorc = (100 / CEmpleadosAProc)
    
    SinError = True
    HuboErrores = False
    Do While Not objReg.EOF
        '----------------------------------------------------------
        MyBeginTrans
        
         Ternro = objReg!Ternro
         Flog.writeline "Inicio Empleado:" & objReg!empleg
        
        
         CalculaVencimientos = False
         Call Politica(1512)
        
         If CalculaVencimientos Then
            Select Case modeloPais
                Case 3: 'Colombia
                    
                    Call DiasVencidos_Col(Ternro, FechaProc)
                    
                Case Else '0: Argentina - 1: chile - 2:uruguay
                    
                    'EAM- Busca el periodo siguiente para el empleado
                    NroVacSiguiente = PeriodoCorrespondiente(Ternro, Periodo_Anio + 1)
                
                    'EAM- Si tiene periodo siguiente calculo el vencimiento y tranferencia
                    If NroVacSiguiente <> 0 Then
                        Call DiasVencidos(Ternro, NroVacSiguiente, NroVac)
                    Else
                        Flog.writeline "El Empleado: " & objReg!empleg & " no posee un perido para tranferir dias "
                    End If
            End Select
         End If
        
        MyCommitTrans
' ----------------------------------------------------------
siguiente:
        Progreso = Progreso + IncPorc
            
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & " WHERE bpronro = " & NroProceso
        objConnProgreso.Execute StrSql, , adExecuteNoRecords
            
        If SinError Then
             ' borro
             StrSql = "DELETE FROM batch_empleado WHERE ternro = " & Ternro & " AND bpronro = " & NroProceso
             objConnProgreso.Execute StrSql, , adExecuteNoRecords
        Else
             StrSql = "UPDATE batch_empleado SET estado = 'Error' WHERE ternro = " & Ternro & " AND bpronro = " & NroProceso
             objConnProgreso.Execute StrSql, , adExecuteNoRecords
        End If
        
        
        objReg.MoveNext
    Loop


'Deshabilito el manejador de errores
On Error GoTo 0

Final:
Flog.writeline "Fin :" & Now
Flog.Close
   
    If HuboErrores Then
        ' actualizo el estado del proceso a Error
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProceso
        objConnProgreso.Execute StrSql, , adExecuteNoRecords
    Else
        ' poner el bprcestado en procesado
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
        objConnProgreso.Execute StrSql, , adExecuteNoRecords
        
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
        If Not IsNull(rs_Batch_Proceso!Empnro) Then
            StrSql = StrSql & ",empnro"
            StrSqlDatos = StrSqlDatos & "," & rs_Batch_Proceso!Empnro
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
        objConnProgreso.Execute StrSql, , adExecuteNoRecords
        
        
        'Reviso que haya copiado
        StrSql = "SELECT * FROM His_batch_proceso WHERE bpronro =" & NroProceso
        OpenRecordset StrSql, rs_His_Batch_Proceso
        
        If Not rs_His_Batch_Proceso.EOF Then
            ' Borro de Batch_proceso
            StrSql = "DELETE FROM Batch_Proceso WHERE bpronro = " & NroProceso
            objConnProgreso.Execute StrSql, , adExecuteNoRecords
        End If
    
        If rs_Batch_Proceso.State = adStateOpen Then rs_Batch_Proceso.Close
        If rs_His_Batch_Proceso.State = adStateOpen Then rs_His_Batch_Proceso.Close
        ' FGZ - 22/09/2003
        ' -----------------------------------------------------------------------------------
    End If
        
Fin:
    objConn.Close
    objConnProgreso.Close
    Set objConn = Nothing
    Set objConnProgreso = Nothing
    
    If objReg.State = adStateOpen Then objReg.Close
    Set objReg = Nothing
    
    Exit Sub
    
    
CE:
    MyRollbackTrans
    HuboErrores = True
    SinError = False
    
    Flog.writeline " ------------------------------------------------------------"
    Flog.writeline "Error procesando Empleado:" & Ternro & " " & Fecha
    Flog.writeline Err.Description
    Flog.writeline "SQL: " & StrSql
    Flog.writeline " ------------------------------------------------------------"
    GoTo siguiente
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

'Public Function FechaAltaEmpleado(Ternro) As Date
'    Dim StrSql As String
'    Dim rsFases As New ADODB.Recordset
'    Dim i_dia As Integer
'    Dim i_mes As Integer
'    Dim i_anio As Integer
'
'    StrSql = "SELECT * FROM fases where fases.empleado = " & Ternro & " AND fases.fasrecofec = -1 "
'    OpenRecordset StrSql, rsFases
'    If rsFases.EOF Then
'        FechaAltaEmpleado = ""
'    Else
'        If IsNull(rsFases("altfec")) Then
'            FechaAltaEmpleado = ""
'        Else
'            FechaAltaEmpleado = CDate(rsFases("altfec"))
'        End If
'    End If
'
'    If rsFases.State = adStateOpen Then rsFases.Close
'    Set rsFases = Nothing
'
'End Function
