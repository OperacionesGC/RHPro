Attribute VB_Name = "mdlNotificaciones"
Option Explicit
'------------------------------------------------------------------------------------
' 11/10/2013 - Gonzalez Nicolás se movieron comentarios a mdlValidarBD
'------------------------------------------------------------------------------------

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

'Dim objBTurno As New BuscarTurno
'Dim objBDia As New BuscarDia
'Dim objFeriado As New Feriado
'Dim objFechasHoras As New FechasHoras

Global diatipo As Byte
Global ok As Boolean

Global FechaAcept As Date
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
Global dias_trabajados As Integer
Global Dias_laborables As Integer

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
Global Version_Valida As Boolean



Public Sub Main()

Dim Fecha As Date
Dim NroVac As Long
Dim Reproceso As Boolean
Dim parametros As String
Dim cantdias As Integer
Dim Columna As Integer
Dim Mensaje As String
Dim Genera As Boolean
Dim NroTPV As String

Dim pos1 As Integer
Dim pos2 As Integer

Dim strCmdLine As String
Dim objconnMain As New ADODB.Connection
Dim Archivo As String

Dim objReg As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim rsNotif As New ADODB.Recordset
Dim rs_Batch_Proceso As New ADODB.Recordset
Dim rs_His_Batch_Proceso As New ADODB.Recordset
Dim rsConsult As New ADODB.Recordset
Dim PID As String
Dim ArrParametros
Dim DiasAnticipacion
'NG
Dim usuario As String
Dim Texto As String
Dim modeloPais
Dim VersionPais
Dim Continua As Boolean
Dim usa1515 As Boolean


Dim TipoNotif As Integer


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
    Archivo = PathFLog & "Vac_Notificaciones" & "-" & NroProceso & ".log"
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Archivo, True)
    
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
    On Error GoTo 0
    
    'Activo el manejador de errores
    On Error GoTo CE
    
    
    '*******************************************************************************************************
    '--------------- VALIDO MODELOS SEGUN POLITICA 1515 | PUEDE TENER ALCANCE POR ESTRUCTURAS --------------
    '*******************************************************************************************************
    '____________________________________________________________________
    'VALIDO QUE LA POLITICA 1515 ESTE ACTIVA Y CONFIGURADA
    Version_Valida = ValidaModeloyVersiones(Version, 13)
    If (Version_Valida = False) Then
        'SI NO ESTA ACTIVA LA 1515 O NO EXISTE CONFIGURACIÓN, TOMA DEFAULT
        modeloPais = Pais_Modelo(7)
        Version_Valida = ValidarVBD(Version, 13, TipoBD, modeloPais)
        usa1515 = False
    Else
        usa1515 = True
    End If
       
    '--------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------

    
    If Not Version_Valida Then
        'Actualizo el progreso
        MyBeginTrans
            StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Error de Version', bprcpid = " & PID & " WHERE bpronro = " & NroProceso
            objConnProgreso.Execute StrSql, , adExecuteNoRecords
        MyCommitTrans
        Flog.writeline
        'GoTo Final
        GoTo CE
    End If
    
    'Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaVersion
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-------------------------------------------------"
    Flog.writeline
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprcpid = " & PID & ", bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objConnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Levanta Proceso y Setea Parámetros:  " & " " & Now
       
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro =" & NroProceso
    OpenRecordset StrSql, rs_Batch_Proceso
       
    If rs_Batch_Proceso.EOF Then Exit Sub
    parametros = rs_Batch_Proceso!bprcparam
    
    '____________________________________________________________
    'NG - VALIDA QUE ESTE ACTIVO LA TRADUCCION A MULTI IDIOMA
    usuario = rs_Batch_Proceso!iduser
    Call Valida_MultiIdiomaActivo(usuario)
    '------------------------------------------------------------
    '------------------------------------------------------------

    
    If Not IsNull(parametros) Then
        If Len(parametros) >= 1 Then
'            pos1 = 1
'            pos2 = InStr(pos1, Parametros, ".") - 1
'            NroVac = CLng(Mid(Parametros, pos1, pos2))
            
            pos1 = 1
            pos2 = InStr(pos1, parametros, ".") - 1
            Reproceso = CBool(Mid(parametros, pos1, pos2 - pos1 + 1))
            
            pos1 = pos2 + 2
            pos2 = InStr(pos1, parametros, ".") - 1
            fecha_desde = CDate(Mid(parametros, pos1, pos2 - pos1 + 1))
            
            pos1 = pos2 + 2
            pos2 = Len(parametros)
            fecha_hasta = CDate(Mid(parametros, pos1, pos2 - pos1 + 1))
    
            If Not IsNull(fecha_desde) Then
                FechaAcept = CDate(fecha_hasta)
            Else
                Flog.writeline "No se pasó correctamente la fecha de aceptación (fecha hasta)"
                Exit Sub
            End If
        End If
    End If
        
        
        'EAM- (v3.07) - Se comento las lineas del IF de la política 1515
        'If usa1515 = False Then
           StrSql = "SELECT * from batch_empleado " & _
                 " INNER JOIN emp_lic ON emp_lic.empleado = batch_empleado.ternro " & _
                 " INNER JOIN lic_vacacion ON lic_vacacion.emp_licnro = emp_lic.emp_licnro " & _
                 " WHERE bpronro = " & NroProceso & " AND emp_lic.tdnro = 2 " & _
                 " AND emp_lic.elfechadesde >= " & ConvFecha(fecha_desde) & _
                 " AND emp_lic.elfechahasta <= " & ConvFecha(fecha_hasta)
        
        'Else
            'EAM- (v3.07) - Se comento las lineas que escribio NG
            'BUSCO LOS EMPLEADOS EN BATCH_PROCESO
            'StrSql = "SELECT bpronro,ternro FROM batch_empleado WHERE bpronro = " & NroProceso
            
        'End If
        OpenRecordset StrSql, objReg
        
        Flog.writeline StrSql
        
        CEmpleadosAProc = objReg.RecordCount
        If CEmpleadosAProc = 0 Then
            CEmpleadosAProc = 1
        End If
        IncPorc = (100 / CEmpleadosAProc)
        
        
        SinError = True
        HuboErrores = False
        
        DiasAnticipacion = 45
        StrSql = "SELECT * FROM confrep WHERE repnro = 168"
        OpenRecordset StrSql, rsConsult
        If Not rsConsult.EOF Then
            If rsConsult!confnrocol = 1 Then
                DiasAnticipacion = rsConsult!confval
                Flog.writeline "Los Dias de Aviso de anticipacion son " & DiasAnticipacion & " dias."
            Else
                Flog.writeline "Los Dias de Aviso de anticipacion se configuran en la columna 1. Sino se consideran 45 dias."
            End If
        End If
        
        Do While Not objReg.EOF
            MyBeginTrans
        

            'NroVac = objReg!vacnro
            Ternro = objReg!Ternro
                       
            Flog.writeline ""
            Flog.writeline "========================================================================"
            Flog.writeline EscribeLogMI("Inicio Empleado") & ": " & Ternro
            
            If usa1515 = True Then
                Call Politica(1515)
                If Not PoliticaOK Then
                    Flog.writeline ""
                    Texto = EscribeLogMI("No se puede procesar al empleado.") & " "
                    Texto = Texto & Replace(EscribeLogMI("Revisar configuración de Política @@NUM@@"), "@@NUM@@", "1515")
                    Flog.writeline Texto
                    
                    Flog.writeline "************************************************************************"
                    Flog.writeline "************************************************************************"
                    GoTo siguiente
                End If
                'modeloPais = st_Opcion
                modeloPais = st_ModeloPais
                VersionPais = st_Opcion
            End If
     
            
            'Flog.writeline "Inicio Empleado:" & Ternro
            Texto = EscribeLogMI("Modelo de vacaciones de @@TXT@@. Nro.")
            Select Case modeloPais
                Case 6: 'Paraguay
                     '************************* PARAGUAY *****************************
                     Select Case VersionPais
                         Case 0: 'Standard Paraguay
                            Texto = Replace(Texto, "@@TXT@@", EscribeLogMI("Paraguay"))
                            Flog.writeline Texto & " " & modeloPais
                            Flog.writeline ""
                            StrSql = "SELECT vdiapednro,vdiapeddesde elfechadesde,vdiapedhasta elfechahasta,vdiapedcant,ternro,vdiaspedestado,vacnro,vdiaspedferiados,vdiaspedhabiles,vdiaspednohabiles,LiquidaVac"
                            StrSql = StrSql & " FROM  vacdiasped"
                            StrSql = StrSql & " WHERE vacdiasped.vdiapeddesde >= " & ConvFecha(fecha_desde)
                            StrSql = StrSql & " AND vacdiasped.vdiapedhasta <= " & ConvFecha(fecha_hasta)
                            StrSql = StrSql & " AND vacdiasped.Ternro =" & Ternro
                            'Tipo notificacion de pedidos
                            TipoNotif = 2
                         End Select

                Case Else
                    StrSql = "SELECT * FROM vacnotif WHERE (emp_licnro = " & objReg!emp_licnro & ") AND (ternro = " & Ternro & ")"
                    TipoNotif = 1
            End Select
            
            Flog.writeline StrSql
            'Ejecuto segun modelo
            OpenRecordset StrSql, rs
            
            If TipoNotif = 2 Then
                If Not rs.EOF = True Then
                    Do While Not rs.EOF
                        '=================================================================
                        ' NOTIFICACIONES
                        '=================================================================
                        '__________________________________
                        'Si son notificaciones de licencias
                        StrSql = "SELECT * FROM vacnotif "
                        StrSql = StrSql & " WHERE "
                        
                        Select Case TipoNotif
                            Case 1: 'BUSCA POR LICENCIAS
                                StrSql = StrSql & " (emp_licnro = " & rs!emp_licnro & ") "
                            Case 2: 'BUSCA POR PEDIDOS
                                StrSql = StrSql & " (vdiapednro = " & rs!vdiapednro & ") "
                        End Select
                        
                        StrSql = StrSql & " AND (ternro = " & Ternro & ")"
                        
                        OpenRecordset StrSql, rsNotif
        
                        If rsNotif.EOF Then
                            '17/10/2014 Carmen Quintero
'                            StrSql = " INSERT INTO vacnotif (vacnotiffecha,vacnotifestado,vacnotiffecacep,vacnotifmanual,ternro,emp_licnro,vdiapednro)"
'                            StrSql = StrSql & " VALUES ("
'                            StrSql = StrSql & ConvFecha(FechaAcept)
'                            StrSql = StrSql & ",0"
'                            StrSql = StrSql & "," & ConvFecha(Date)
'                            StrSql = StrSql & ",0"
'                            StrSql = StrSql & "," & Ternro
                            
                            StrSql = " INSERT INTO vacnotif (vacnotiffecha,vacnotifestado,vacnotiffecacep,vacnotifmanual,ternro,emp_licnro,vdiapednro)"
                            StrSql = StrSql & " VALUES ("
                            StrSql = StrSql & "," & ConvFecha(Date)
                            StrSql = StrSql & ",0"
                            StrSql = StrSql & ConvFecha(FechaAcept)
                            StrSql = StrSql & ",0"
                            StrSql = StrSql & "," & Ternro
                            'fin
                            If TipoNotif = 1 Then
                                StrSql = StrSql & "," & rs!emp_licnro 'emp_lic
                                StrSql = StrSql & ",NULL"
                            ElseIf TipoNotif = 2 Then
                                StrSql = StrSql & ",NULL"
                                StrSql = StrSql & "," & rs!vdiapednro 'vdiapednro
                            End If
                            StrSql = StrSql & ")"
                            
                            objConn.Execute StrSql, , adExecuteNoRecords
                            
                            Flog.writeline EscribeLogMI("Notificación Generada")
                        Else
                            If Reproceso Then
                                StrSql = "UPDATE vacnotif SET vacnotifestado = 0"
                                StrSql = StrSql & ",vacnotiffecacep = " & ConvFecha(FechaAcept)
                                StrSql = StrSql & ",vacnotiffecha = " & ConvFecha(Date)
                                StrSql = StrSql & " WHERE "
                                                                
                                If TipoNotif = 1 Then
                                    StrSql = StrSql & " (emp_licnro = " & rs!emp_licnro & ")"
                                ElseIf TipoNotif = 2 Then
                                    StrSql = StrSql & " (vdiapednro = " & rs!vdiapednro & ")"
                                End If
                                
                                StrSql = StrSql & " AND (ternro = " & Ternro & ")"
                                
                                objConn.Execute StrSql, , adExecuteNoRecords
                                Flog.writeline EscribeLogMI("Notificación Generada")
                            End If
                        End If
                        '=================================================================
                        '=================================================================
                  
                        rs.MoveNext
                    Loop
                
                Else
                    Flog.writeline "Sin nada para notificar"
                End If
            Else
                'FGZ - 17/12/2004
                ' Por Ley la fecha de aceptacion debe ser 45 dias antes de la fecha de la licencia
                FechaAcept = CDate(objReg!elfechadesde) - DiasAnticipacion
                Do While Not Es_Dia_Habil(FechaAcept, objReg!Ternro)
                    FechaAcept = FechaAcept - 1
                Loop
                'Flog.writeline FechaAcept
                
               ' If TipoNotif = 1 And Not rs.EOF Then MDF
                If TipoNotif = 1 And rs.EOF Then
                
                        StrSql = "SELECT * FROM vacnotif "
                        StrSql = StrSql & " WHERE "
                        
                        Select Case TipoNotif
                            Case 1: 'BUSCA POR LICENCIAS
                                StrSql = StrSql & " (emp_licnro = " & objReg!emp_licnro & ") "
                        End Select
                        
                        StrSql = StrSql & " AND (ternro = " & Ternro & ")"
                        
                        OpenRecordset StrSql, rsNotif
        
                        If rsNotif.EOF Then
                            '21/10/2014 Carmen Quintero
                            'StrSql = " INSERT INTO vacnotif (vacnotiffecha,vacnotifestado,vacnotiffecacep,vacnotifmanual,ternro,emp_licnro,vdiapednro)"
                            'StrSql = StrSql & " VALUES ("
                            'StrSql = StrSql & ConvFecha(FechaAcept)
                            'StrSql = StrSql & ",0"
                            'StrSql = StrSql & "," & ConvFecha(Date)
                            'StrSql = StrSql & ",0"
                            'StrSql = StrSql & "," & Ternro
                            
                            StrSql = " INSERT INTO vacnotif (vacnotiffecha,vacnotifestado,vacnotiffecacep,vacnotifmanual,ternro,emp_licnro,vdiapednro)"
                            StrSql = StrSql & " VALUES ("
                            'StrSql = StrSql & "," & ConvFecha(Date) 'mdf
                            StrSql = StrSql & ConvFecha(Date)
                            'StrSql = StrSql & ",0" 'mdf
                            StrSql = StrSql & ",0,"
                            StrSql = StrSql & ConvFecha(FechaAcept)
                            StrSql = StrSql & ",0"
                            StrSql = StrSql & "," & Ternro
                            
                            If TipoNotif = 1 Then
                                'StrSql = StrSql & "," & rs!emp_licnro 'emp_lic
                                StrSql = StrSql & "," & objReg!emp_licnro
                                StrSql = StrSql & ",NULL"
                            ElseIf TipoNotif = 2 Then
                                StrSql = StrSql & ",NULL"
                                StrSql = StrSql & "," & rs!vdiapednro 'vdiapednro
                            End If
                            StrSql = StrSql & ")"
                            
                            objConn.Execute StrSql, , adExecuteNoRecords
                            
                            Flog.writeline EscribeLogMI("Notificación Generada")
                        Else
                            If Reproceso Then
                                StrSql = "UPDATE vacnotif SET vacnotifestado = 0"
                                StrSql = StrSql & ",vacnotiffecacep = " & ConvFecha(FechaAcept)
                                StrSql = StrSql & ",vacnotiffecha = " & ConvFecha(Date)
                                StrSql = StrSql & " WHERE "
                                                                
                                If TipoNotif = 1 Then
                                    StrSql = StrSql & " (emp_licnro = " & rs!emp_licnro & ")"
                                ElseIf TipoNotif = 2 Then
                                    StrSql = StrSql & " (vdiapednro = " & rs!vdiapednro & ")"
                                End If
                                
                                StrSql = StrSql & " AND (ternro = " & Ternro & ")"
                                
                                objConn.Execute StrSql, , adExecuteNoRecords
                                Flog.writeline EscribeLogMI("Notificación Generada")
                            End If
                        End If
                        '=================================================================
                        '=================================================================
                Else
                    Flog.writeline "Sin nada para notificar"
                End If
                
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


Public Function Es_Dia_Habil(ByVal Fecha As Date, ByVal Ternro As Long) As Boolean
' Devuelve true si es dia habil para el empleado o falso en caso contrario
Dim objFeriado As New Feriado
Dim Habil As Boolean
Dim EsFeriado As Boolean

    Set objFeriado.Conexion = objConn
    Set objFeriado.ConexionTraza = objConn
    
    Habil = True
    EsFeriado = objFeriado.Feriado(Fecha, Ternro, False)

    If Not EsFeriado Then
        If Weekday(Fecha) = 1 Or Weekday(Fecha) = 7 Then
            Habil = False
        End If
    Else
        Habil = False
    End If

    Es_Dia_Habil = Habil
End Function
