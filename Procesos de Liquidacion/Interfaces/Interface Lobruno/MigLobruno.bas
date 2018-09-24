Attribute VB_Name = "MigracionLobruno"
Option Explicit

'Global Const Version = "1.01"
'Global Const FechaModificacion = "23/03/2009"
'Global Const UltimaModificacion = "Inicial"

'Global Const Version = "1.02"
'Global Const FechaModificacion = "27/05/2009"
'Global Const UltimaModificacion = "Se modificaron los nombres de las tablas en la base lobruno"
' FAF - 27/05/2009 - Se modificaron los nombres de las tablas en la base lobruno

'Global Const Version = "1.03"
'Global Const FechaModificacion = "29/06/2009"
'Global Const UltimaModificacion = ""
' FAF - 29/06/2009 - No se realiza la conexion a la base de datos Lobruno, ya que las tablas son vistas desde la base rhpro

'Global Const Version = "1.04"
'Global Const FechaModificacion = "16/07/2009"
'Global Const UltimaModificacion = ""
' FAF - 16/07/2009 - Se modifico el nombre de la tabla legajos_familiares_rhpro por legajo_familiares_rhpro a pedido del cliente.

Global Const Version = "1.05"
Global Const FechaModificacion = "03/09/2009"
Global Const UltimaModificacion = ""
' FAF - 03/09/2009 - Si el registro de auditoria hace referencia a un registro que no existe, da error. Caso de historico estructuras.

'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
Dim fs, f

Dim NroProceso As Long

Global Path As String
Global NArchivo As String
Global Rta
Global HuboErrores As Boolean
Global EmpErrores As Boolean

Global NroColumna As Long
Global Tabulador As Long
Global Tabs As Long

Global EncontroAlguno As Boolean

Global IdUser As String
Global Fecha As Date
Global Hora As String

Global objconnBaseLobruno As New ADODB.Connection

Dim conexion_base_LBR As String
Dim formatofecha As String
Dim C_puesto As Integer
Dim C_grado As Integer
Dim C_empresa As Integer
Dim C_ccosto As Integer
Dim C_estruct_area As Integer
Dim C_centro_trabajo As Integer
Dim C_edificio As Integer
Dim C_canal_costo As Integer
Dim C_dgi_exento As Integer
Dim C_rel_laboral As Integer
Dim C_convenio As Integer
Dim C_cat_conv As Integer
Dim C_fun_conv As Integer
Dim C_horas_mesh As Integer
Dim C_horas_mesd As Integer
Dim C_osocial As Integer
Dim C_caracter_serv As Integer
Dim C_grupo_liq As Integer
Dim C_instrum_pago As Integer



Private Sub Main()

Dim Directorio As String
Dim CArchivos
Dim archivo
Dim Folder
Dim strCmdLine As String
Dim Nombre_Arch As String

Dim StrSql As String
Dim objRs As New ADODB.Recordset
Dim tipoDepuracion
Dim historico As Boolean
Dim param
Dim I
Dim cantRegistros
Dim PID As String
Dim parametros
Dim ArrParametros
Dim fechadesde As Date
Dim fechahasta As Date
Dim horadesde As String
Dim horahasta As String
Dim sinprocesar As Integer
Dim hayfiltro As Boolean
Dim fechainiej As String
Dim horainiej As String

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

    Nombre_Arch = PathFLog & "InterfaceLobruno" & "-" & NroProceso & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    TiempoInicialProceso = GetTickCount
    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If

    On Error Resume Next
    OpenConnection strconexion, objconnProgreso
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    
    On Error GoTo CE:
    
    HuboErrores = False
    Tabulador = 5
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Version = " & Version
    Flog.writeline "Modificacion = " & UltimaModificacion
    Flog.writeline "Fecha = " & FechaModificacion
    Flog.writeline "-----------------------------------------------------------------"
    
    Flog.writeline
    Flog.writeline "PID = " & PID
    
    Flog.writeline "Inicio Proceso de Migracion de Lobruno : " & Now
    Flog.writeline "Cambio el estado del proceso a Procesando"
    
    'Cambio el estado del proceso a Procesando
    fechainiej = ConvFecha(Date)
    horainiej = Format(Now, "hh:mm:ss ")
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & horainiej & "', bprcfecinicioej = " & fechainiej & ", bprcestado = 'Procesando', bprcpid = " & PID & ", bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Obtengo los parametros del proceso"
    
    TiempoAcumulado = GetTickCount
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, objRs
    
    If Not objRs.EOF Then
       IdUser = objRs!IdUser
       Fecha = objRs!bprcfecha
       Hora = objRs!bprchora
       
       'Obtengo los parametros del proceso
       parametros = objRs!bprcparam
       
       hayfiltro = False
       sinprocesar = 0
       
       ' Armar una alerta que avice que una auditoria no se modifico.
       If parametros <> "" Then
            ArrParametros = Split(parametros, "@")
            
            fechadesde = ArrParametros(0)
            horadesde = ArrParametros(1)
            fechahasta = ArrParametros(2)
            horahasta = ArrParametros(3)
            sinprocesar = ArrParametros(4)
            
            hayfiltro = True
       Else
            ' Si no hay parametros, esta programado automaticamente. Sacar la fecha desde de ultimo
            ' proceso del mismo tipo
            
            StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 235 AND bpronro <> " & NroProceso & " ORDER BY bpronro DESC"
            OpenRecordset StrSql, objRs
            If Not objRs.EOF Then
                fechadesde = objRs!bprcfecInicioEj
                horadesde = objRs!bprcHoraInicioEj
                fechahasta = Format(Mid(fechainiej, 2, Len(fechainiej) - 2), FormatoInternoFecha)
                horahasta = horainiej
                sinprocesar = -1
            Else
                fechadesde = Format(Mid(fechainiej, 2, Len(fechainiej) - 2), FormatoInternoFecha)
                horadesde = Format("00:00:00", "hh:mm:ss ")
                fechahasta = Format(Mid(fechainiej, 2, Len(fechainiej) - 2), FormatoInternoFecha)
                horahasta = horainiej
                sinprocesar = -1
            End If
       End If
       
       Flog.writeline Espacios(Tabulador * 1) & "FECHA DESDE    :" & fechadesde
       Flog.writeline Espacios(Tabulador * 1) & "HORA DESDE     :" & horadesde
       Flog.writeline Espacios(Tabulador * 1) & "FECHA HASTA    :" & fechahasta
       Flog.writeline Espacios(Tabulador * 1) & "HORA HASTA     :" & horahasta
       Flog.writeline Espacios(Tabulador * 1) & "AUD. SIN PROC. :" & sinprocesar
       Flog.writeline
       
       ' Proceso que migra (exporta) los datos
       Call ComenzarTransferencia(hayfiltro, fechadesde, fechahasta, horadesde, horahasta, sinprocesar)
       
       ' Hacer el pasaje a la base del cliente
       'Call Comenzarmigracion
       
    Else
        Exit Sub
    End If
   
    'Actualizo el estado del proceso
    If Not HuboErrores Then
       StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
       Flog.writeline "Proceso Finalizado Correctamente"
    Else
       StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProceso
       Flog.writeline "Proceso Incompleto"
    End If
    
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Fin :" & Now
    Flog.Close

    Exit Sub
    
CE:
'Resume Next
    HuboErrores = True
    Flog.writeline " Error: " & Err.Description & Now

End Sub

Public Sub ComenzarTransferencia(ByVal hayfiltro As Boolean, ByVal fechadesde As Date, ByVal fechahasta As Date, ByVal horadesde As String, ByVal horahasta As String, ByVal sinprocesar As Integer)
Dim objAuditoria As New ADODB.Recordset
Dim objrS1 As New ADODB.Recordset
Dim objRs2 As New ADODB.Recordset
Dim v_legajo As String
Dim v_apellido_nombre As String
Dim v_sexo As String
Dim v_estcivil As String
Dim v_apellido_casada As String
Dim v_fech_nac
Dim v_nacion_nac As String
Dim v_lugar_nac As String
Dim v_nacionalidad As String
Dim v_doc_tipo As String
Dim v_doc_nro As String
Dim v_empresa As String
Dim sin_error As Boolean
Dim conexion_base_LBR As String
Dim formatofecha As String
Dim v_fech_ini_vinc
Dim v_parentesco As String
Dim v_fecha
Dim v_estado As String
Dim v_puesto As String
Dim v_grado As String
Dim v_ccosto As String
Dim v_estruct_area As String
Dim v_centro_trabajo As String
Dim v_edificio As String
Dim v_canal_costo As String
Dim v_fecha_ingreso
Dim v_Fecha_Antiguedad
Dim v_Ingreso_Grupo
Dim v_Fecha_Promo
Dim v_Fecha_Baja
Dim v_Motivo_Baja As String
Dim v_rel_laboral As String
Dim v_dgi_exento As String
Dim v_convenio As String
Dim v_cat_conv As String
Dim v_fun_conv As String
Dim v_horas_mesh As String
Dim v_horas_mesd As String
Dim v_osocial As String
Dim v_caracter_serv As String
Dim v_grupo_liq As String
Dim v_est_Nivel As String
Dim v_tipo_cuenta As String
Dim v_entid_pago As String
Dim v_nro_cuenta As String
Dim v_cbu As String
Dim v_suc_pago As String
Dim v_instrum_pago As String
Dim v_nro_cuil As String
Dim v_prov_trabaja As String
Dim v_estrnro
Dim v_inicial
Dim v_final
Dim analizar As Boolean
Dim v_tipo
'Dim v_cambio_fec_antig As Boolean
'Dim v_cambio_ing_grupo As Boolean
'Dim v_cambio_fec_promo As Boolean


    Call ConfReporte
    
    If HuboErrores Then
        GoTo Fin:
    End If
    
    On Error GoTo CE:
    
    StrSql = "SELECT * FROM  auditoria "
    If sinprocesar Then
        StrSql = StrSql & " WHERE procesado = 0 "
    Else
        StrSql = StrSql & " WHERE procesado = -1 "
    End If
    
    StrSql = StrSql & " AND (auditoria.aud_fec >= " & ConvFecha(fechadesde) & " ) AND "
    StrSql = StrSql & " (auditoria.aud_fec <= " & ConvFecha(fechahasta) & " ) "
    StrSql = StrSql & " ORDER BY audnro ASC"
    
    OpenRecordset StrSql, objAuditoria
    
    Flog.writeline Espacios(Tabulador * 0) & "SQL de auditorias a procesar --> " & StrSql
    
    'Determino la proporcion de progreso
    Progreso = 0
    CEmpleadosAProc = objAuditoria.RecordCount
    If CEmpleadosAProc = 0 Then
        CEmpleadosAProc = 1
    End If
    IncPorc = (95 / CEmpleadosAProc)
    
    'Actualizo el progreso
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & IncPorc & " WHERE bpronro = " & NroProceso
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline Espacios(Tabulador * 0) & " "
    Flog.writeline Espacios(Tabulador * 0) & "Actualización Primera vez del avance del proceso"
    
    Do Until objAuditoria.EOF
        If Not ((Format(fechadesde, FormatoInternoFecha) = Format(objAuditoria!aud_fec, FormatoInternoFecha) And Format(horadesde, FormatoInternoHora) > Format(objAuditoria!aud_hor, FormatoInternoHora)) Or (Format(fechahasta, FormatoInternoFecha) = Format(objAuditoria!aud_fec, FormatoInternoFecha) And Format(horahasta, FormatoInternoHora) < Format(objAuditoria!aud_hor, FormatoInternoHora))) Then
            
            sin_error = True
            v_tipo = ""
'            v_cambio_fec_antig = False
'            v_cambio_ing_grupo = False
'            v_cambio_fec_promo = False
            
            Select Case objAuditoria!caudnro
'--------'1, 'Alta Empleado'
                Case 1:
                    Flog.writeline Espacios(Tabulador * 0) & "  Tipo " & objAuditoria!caudnro & " - Auditoria Nro --> " & objAuditoria!audnro & " - " & objAuditoria!aud_des
                    StrSql = "SELECT tercero.terape, tercero.ternom, tercero.tersex, tercero.estcivnro, tercero.tercasape "
                    StrSql = StrSql & " ,tercero.terfecnac,tercero.paisnro,lugar_nac.lugardesc,tercero.nacionalnro,empleado.empleg,empleado.empfaltagr,empleado.nivnro "
                    StrSql = StrSql & " FROM tercero INNER JOIN empleado ON tercero.ternro=empleado.ternro "
                    StrSql = StrSql & " LEFT JOIN lugar_nac ON tercero.lugarnro=lugar_nac.lugarnro "
                    StrSql = StrSql & " WHERE tercero.ternro = " & objAuditoria!aud_ternro
                    If objAuditoria!aud_ternro = "" Or IsNull(objAuditoria!aud_ternro) Then
                        Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", objAuditoria!audnro, True, StrSql)
                    Else
                        OpenRecordset StrSql, objrS1
                        If Not objrS1.EOF Then
                            v_legajo = objrS1!empleg
                            v_apellido_nombre = objrS1!terape & ", " & objrS1!ternom
                            If objrS1!tersex = -1 Then
                                v_sexo = "M"
                            Else
                                v_sexo = "F"
                            End If
                            v_estcivil = Mapeo(v_legajo, "LEGAJOS", "ESTADO_CIVIL", objAuditoria!audnro, "[ESTCIVIL]", objrS1!estcivnro, sin_error)
                            If sin_error Then
                                v_apellido_casada = IIf(EsNulo(objrS1!tercasape), "", objrS1!tercasape)
                                v_fech_nac = objrS1!terfecnac
                                v_nacion_nac = Mapeo(v_legajo, "LEGAJOS", "NACION_NACIMIENTO", objAuditoria!audnro, "[PAIS]", objrS1!paisnro, sin_error)
                                If sin_error Then
                                    v_lugar_nac = IIf(EsNulo(objrS1!lugardesc), "", objrS1!lugardesc)
                                    v_nacionalidad = Mapeo(v_legajo, "LEGAJOS", "NACIONALIDAD", objAuditoria!audnro, "[NACIONAL]", objrS1!nacionalnro, sin_error)
                                    If sin_error Then
                                        v_tipo = "ALTA_EMP,"
'                                        Call InsertarItemLegajo("ALTA_EMP", v_legajo, v_apellido_nombre, "", v_sexo, v_estcivil, v_apellido_casada, CDate(v_fech_nac), v_nacion_nac, v_lugar_nac, v_nacionalidad, "", "", "", sin_error, objAuditoria!audnro)
                                    End If
                                End If
                            End If
                            If sin_error And Not IsNull(objrS1!empfaltagr) And objrS1!empfaltagr <> "" Then
                                v_fecha_ingreso = objrS1!empfaltagr
                                v_tipo = v_tipo & "FECHA_INGRESO,"
'                                Call InsertarItemLegajoLiquidar("FECHA_INGRESO", v_legajo, "", CDate(v_fecha_ingreso), Empty, Empty, "", "", "", "", "", "", "", 0, 0, "", "", "", "", "", "", Empty, "", "", "", "", "", "", Empty, "", sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
                            End If
                            If sin_error And Not EsNulo(objrS1!nivnro) And objrS1!nivnro <> "0" Then
                                v_est_Nivel = Mapeo(v_legajo, "LEGAJO_LIQUIDAR", "ESTUDIOS_NIVEL", objAuditoria!audnro, "[EST_NIVE]", objrS1!nivnro, sin_error)
                                v_tipo = v_tipo & "ESTUDIOS_NIVEL,"
'                                Call InsertarItemLegajoLiquidar("ESTUDIOS_NIVEL", v_legajo, "", Empty, Empty, Empty, "", "", "", v_est_Nivel, "", "", "", 0, 0, "", "", "", "", "", "", Empty, "", "", "", "", "", "", Empty, "", sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
                            End If
                            If Not EsNulo(v_tipo) Then
                                v_tipo = Left(v_tipo, Len(v_tipo) - 1)
                                Call InsertarItemLegajoItemLiquidar(v_tipo, v_legajo, v_apellido_nombre, "", v_sexo, v_estcivil, v_apellido_casada, CDate(v_fech_nac), v_nacion_nac, v_lugar_nac, v_nacionalidad, "", "", "", CDate(v_fecha_ingreso), v_est_Nivel, sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
                            End If
                        Else
                            Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", objAuditoria!audnro, True, StrSql)
                        End If
                    End If
                    
'--------'3, 'Modificación Empleado'
                Case 3:
                    Select Case objAuditoria!aud_campnro
                        Case 55, 57:
                            '55, 'Apellido del empleado', 'terape', 'empleado'
                            '57, 'Nombre  del empleado', 'ternom', 'empleado'
                            Flog.writeline Espacios(Tabulador * 0) & "  Tipo " & objAuditoria!caudnro & " - Campo 55 - Auditoria Nro --> " & objAuditoria!audnro & " - " & objAuditoria!aud_des
                            StrSql = "SELECT tercero.terape, tercero.ternom, empleado.empleg "
                            StrSql = StrSql & " FROM tercero INNER JOIN empleado ON tercero.ternro=empleado.ternro "
                            StrSql = StrSql & " WHERE tercero.ternro = " & objAuditoria!aud_ternro
                            If objAuditoria!aud_ternro = "" Or IsNull(objAuditoria!aud_ternro) Then
                                Call InsertarError("0", "GENERAL", "GENERAL", "Error. El campo aud_ternro es nulo. No se puede procesar.", objAuditoria!audnro, True, StrSql)
                            Else
                                OpenRecordset StrSql, objrS1
                                If Not objrS1.EOF Then
                                    v_legajo = objrS1!empleg
                                    v_apellido_nombre = objrS1!terape & " / " & objrS1!ternom
                                    Call InsertarItemLegajo("AP_NOM", v_legajo, v_apellido_nombre, "", "", "", "", Empty, "", "", "", "", "", "", sin_error, objAuditoria!audnro)
                                Else
                                    Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", objAuditoria!audnro, True, StrSql)
                                End If
                            End If
                        Case 60:
                            '60, 'Fecha de Alta al grupo del empleado', 'empfaltagr', 'empleado'
                            Flog.writeline Espacios(Tabulador * 0) & "  Auditoria Nro --> " & objAuditoria!audnro & " - " & objAuditoria!aud_des
                            StrSql = "SELECT empleado.empfaltagr, empleado.empleg "
                            StrSql = StrSql & " FROM empleado "
                            StrSql = StrSql & " WHERE empleado.ternro = " & objAuditoria!aud_ternro
                            If objAuditoria!aud_ternro = "" Or IsNull(objAuditoria!aud_ternro) Then
                                Call InsertarError("0", "GENERAL", "GENERAL", "Error. El campo aud_ternro es nulo. No se puede procesar.", objAuditoria!audnro, True, StrSql)
                            Else
                                OpenRecordset StrSql, objrS1
                                If Not objrS1.EOF Then
                                    v_legajo = objrS1!empleg
                                    v_fecha_ingreso = objrS1!empfaltagr
                                    Call InsertarItemLegajoLiquidar("FECHA_INGRESO", v_legajo, "", CDate(v_fecha_ingreso), Empty, Empty, "", "", "", "", "", "", "", 0, 0, "", "", "", "", "", "", Empty, "", "", "", "", "", "", Empty, "", sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
                                Else
                                    Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", objAuditoria!audnro, True, StrSql)
                                End If
                            End If
                        Case 58:
                            '58, 'Estado del empleado', 'empest', 'empleado'
                            Flog.writeline Espacios(Tabulador * 0) & "  Tipo " & objAuditoria!caudnro & " - Campo 58 - Auditoria Nro --> " & objAuditoria!audnro & " - " & objAuditoria!aud_des
                            StrSql = "SELECT empleado.empest, empleado.empleg "
                            StrSql = StrSql & " FROM empleado "
                            StrSql = StrSql & " WHERE empleado.ternro = " & objAuditoria!aud_ternro
                            If objAuditoria!aud_ternro = "" Or IsNull(objAuditoria!aud_ternro) Then
                                Call InsertarError("0", "GENERAL", "GENERAL", "Error. El campo aud_ternro es nulo. No se puede procesar.", objAuditoria!audnro, True, StrSql)
                            Else
                                OpenRecordset StrSql, objrS1
                                If Not objrS1.EOF Then
                                    If CBool(objrS1!empest) Then
                                        v_estado = "A"
                                    Else
                                        v_estado = "T"
                                    End If
                                    Call InsertarItemLegajoLiquidar("ESTADO", v_legajo, "", Empty, Empty, Empty, v_estado, "", "", "", "", "", "", 0, 0, "", "", "", "", "", "", Empty, "", "", "", "", "", "", Empty, "", sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
                                Else
                                    Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", objAuditoria!audnro, True, StrSql)
                                End If
                            End If
                            
                        Case 262:
                            '262, 'Nivel Estudio Actual Empleado', 'nivnro', 'empleado'
                            Flog.writeline Espacios(Tabulador * 0) & "  Tipo " & objAuditoria!caudnro & " - Campo 262 - Auditoria Nro --> " & objAuditoria!audnro & " - " & objAuditoria!aud_des
                            StrSql = "SELECT empleado.nivnro, empleado.empleg FROM empleado "
                            StrSql = StrSql & " WHERE empleado.ternro = " & objAuditoria!aud_ternro
                            If objAuditoria!aud_ternro = "" Or IsNull(objAuditoria!aud_ternro) Then
                                Call InsertarError("0", "GENERAL", "GENERAL", "Error. El campo aud_ternro es nulo. No se puede procesar.", objAuditoria!audnro, True, StrSql)
                            Else
                                OpenRecordset StrSql, objrS1
                                If Not objrS1.EOF Then
                                    v_legajo = objrS1!empleg
                                    v_est_Nivel = Mapeo(v_legajo, "LEGAJO_LIQUIDAR", "ESTUDIOS_NIVEL", objAuditoria!audnro, "[EST_NIVE]", objrS1!nivnro, sin_error)
                                    Call InsertarItemLegajoLiquidar("ESTUDIOS_NIVEL", v_legajo, "", Empty, Empty, Empty, "", "", "", v_est_Nivel, "", "", "", 0, 0, "", "", "", "", "", "", Empty, "", "", "", "", "", "", Empty, "", sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
                                Else
                                    Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", objAuditoria!audnro, True, StrSql)
                                End If
                            End If
                    End Select
                        
'--------'33, 'Alta Histórico Estruc Empleado
'--------'34, 'Modif. Hist. Estruc. Empleado'
                Case 33, 34:
                    Flog.writeline Espacios(Tabulador * 0) & "  Tipo " & objAuditoria!caudnro & " - Auditoria Nro --> " & objAuditoria!audnro & " - " & objAuditoria!aud_des
                    If objAuditoria!aud_campnro = 91 Or objAuditoria!aud_campnro = 94 Then
                        v_inicial = InStr(1, objAuditoria!aud_des, ":")
                        v_final = InStr(1, objAuditoria!aud_des, "-")
                        If v_inicial > 0 And v_final > 0 Then
                            v_inicial = v_inicial + 2
                            v_estrnro = Mid(objAuditoria!aud_des, v_inicial, v_final - v_inicial)
                        Else
                            v_estrnro = 0
                        End If
                    Else
                        v_estrnro = objAuditoria!aud_actual
                    End If
                    StrSql = "SELECT empleado.empleg, his_estructura.tenro, his_estructura.estrnro, his_estructura.htetdesde, his_estructura.htethasta "
                    StrSql = StrSql & " FROM empleado INNER JOIN his_estructura ON his_estructura.ternro=empleado.ternro "
                    StrSql = StrSql & " WHERE empleado.ternro = " & objAuditoria!aud_ternro & " AND his_estructura.estrnro = " & v_estrnro
'                    If objAuditoria!aud_campnro = 89 Then
'                        StrSql = "SELECT tenro FROM estructura WHERE estrnro=" & objAuditoria!aud_actual
                    If IsNull(objAuditoria!aud_ternro) Or objAuditoria!aud_ternro = "" Then
                        Call InsertarError("0", "GENERAL", "GENERAL", "Error. El campo aud_ternro es nulo. No se puede procesar.", objAuditoria!audnro, True, StrSql)
                    Else
                        OpenRecordset StrSql, objrS1
                        If objrS1.EOF Then
                            Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se encontro a que Tipo de Estructura pertenece la estructura con código " & objAuditoria!aud_actual & ". No se puede procesar.", objAuditoria!audnro, True, StrSql)
                        Else
'                            StrSql = "SELECT empleado.empleg, his_estructura.tenro, his_estructura.estrnro, his_estructura.htetdesde, his_estructura.htethasta "
'                            StrSql = StrSql & " FROM empleado INNER JOIN his_estructura ON his_estructura.ternro=empleado.ternro "
'                            StrSql = StrSql & " WHERE empleado.ternro = " & objAuditoria!aud_ternro
                            analizar = False
                            Select Case CInt(objrS1!tenro)
                                Case C_puesto:
                                    StrSql = StrSql & " AND his_estructura.tenro=" & CStr(C_puesto)
                                    analizar = True
                                Case C_grado:
                                    StrSql = StrSql & " AND his_estructura.tenro=" & CStr(C_grado)
                                    analizar = True
                                Case C_empresa:
                                    StrSql = StrSql & " AND his_estructura.tenro=" & CStr(C_empresa)
                                    analizar = True
                                Case C_ccosto:
                                    StrSql = StrSql & " AND his_estructura.tenro=" & CStr(C_ccosto)
                                    analizar = True
                                Case C_estruct_area:
                                    StrSql = StrSql & " AND his_estructura.tenro=" & CStr(C_estruct_area)
                                    analizar = True
                                Case C_centro_trabajo:
                                    StrSql = StrSql & " AND his_estructura.tenro=" & CStr(C_centro_trabajo)
                                    analizar = True
                                Case C_edificio:
                                    StrSql = StrSql & " AND his_estructura.tenro=" & CStr(C_edificio)
                                    analizar = True
                                Case C_canal_costo:
                                    StrSql = StrSql & " AND his_estructura.tenro=" & CStr(C_canal_costo)
                                    analizar = True
                                Case C_dgi_exento:
                                    StrSql = StrSql & " AND his_estructura.tenro=" & CStr(C_dgi_exento)
                                    analizar = True
                                Case C_rel_laboral:
                                    StrSql = StrSql & " AND his_estructura.tenro=" & CStr(C_rel_laboral)
                                    analizar = True
                                Case C_convenio:
                                    StrSql = StrSql & " AND his_estructura.tenro=" & CStr(C_convenio)
                                    analizar = True
                                Case C_cat_conv:
                                    StrSql = StrSql & " AND his_estructura.tenro=" & CStr(C_cat_conv)
                                    analizar = True
                                Case C_fun_conv:
                                    StrSql = StrSql & " AND his_estructura.tenro=" & CStr(C_fun_conv)
                                    analizar = True
                                Case C_horas_mesh:
                                    StrSql = StrSql & " AND his_estructura.tenro=" & CStr(C_horas_mesh)
                                    analizar = True
                                Case C_horas_mesd:
                                    StrSql = StrSql & " AND his_estructura.tenro=" & CStr(C_horas_mesd)
                                    analizar = True
                                Case C_osocial:
                                    StrSql = StrSql & " AND his_estructura.tenro=" & CStr(C_osocial)
                                    analizar = True
                                Case C_caracter_serv:
                                    StrSql = StrSql & " AND his_estructura.tenro=" & CStr(C_caracter_serv)
                                    analizar = True
                                Case C_grupo_liq:
                                    StrSql = StrSql & " AND his_estructura.tenro=" & CStr(C_grupo_liq)
                                    analizar = True
                                Case C_instrum_pago:
                                    StrSql = StrSql & " AND his_estructura.tenro=" & CStr(C_instrum_pago)
                                    analizar = True
                            End Select
                            StrSql = StrSql & " ORDER BY his_estructura.htetdesde DESC"
                            If analizar Then
                                OpenRecordset StrSql, objrS1
                                If Not objrS1.EOF Then
                                    v_legajo = objrS1!empleg
                                    If IsNull(objrS1!htethasta) Or objrS1!htethasta = "" Then
                                        v_fecha = CDate(objrS1!htetdesde)
                                        v_estado = "S"
                                    Else
                                        v_fecha = CDate(objrS1!htethasta)
                                        v_estado = "N"
                                    End If
                                    
                                    Select Case CInt(objrS1!tenro)
                                        Case C_puesto:
                                            v_puesto = Mapeo(v_legajo, "LEGAJO_POSICION", "PUESTO", objAuditoria!audnro, "[PUESTO]", objrS1!estrnro, sin_error)
                                            If sin_error Then
                                                Call InsertarItemLegajoPosicion("PUESTO", v_legajo, CDate(v_fecha), v_puesto, v_estado, "", "", "", "", "", "", "", "", Empty, "", "", "", "", "", "", sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
                                            End If
                                        Case C_grado:
                                            v_grado = Mapeo(v_legajo, "LEGAJO_POSICION", "GRADO", objAuditoria!audnro, "[GRADO]", objrS1!estrnro, sin_error)
                                            If sin_error Then
                                                Call InsertarItemLegajoPosicion("GRADO", v_legajo, CDate(v_fecha), "", "", v_grado, v_estado, "", "", "", "", "", "", Empty, "", "", "", "", "", "", sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
                                            End If
                                        Case C_empresa:
                                            v_empresa = Mapeo(v_legajo, "LEGAJOS", "EMPRESA", objAuditoria!audnro, "[EMPRESA]", objrS1!estrnro, sin_error)
                                            If sin_error Then
                                                v_tipo = "EMPRESA,"
    '                                            Call InsertarItemLegajo("EMPRESA", v_legajo, "", "", "", "", "", Empty, "", "", "", "", "", v_empresa, sin_error, objAuditoria!audnro)
    '                                            Call InsertarItemLegajoPosicion("EMPRESA", v_legajo, CDate(v_fecha), "", "", "", "", v_empresa, v_estado, "", "", "", "", Empty, "", "", "", "", "", "", sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
                                                If objAuditoria!caudnro = 33 Then
                                                    v_tipo = v_tipo & "LIQUIDAR,"
    '                                                Call InsertarItemLegajoLiquidar("EMPRESA", v_legajo, v_empresa, Empty, Empty, Empty, "", "", "", "", "", "", "", 0, 0, "", "", "", "", "", "", Empty, "", "", "", "", "", "", Empty, "", sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
                                                End If
                                                v_tipo = Left(v_tipo, Len(v_tipo) - 1)
                                                Call InsertarItemEmpresa(v_tipo, v_legajo, CDate(v_fecha), v_empresa, v_estado, sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
                                            End If
                                        Case C_ccosto:
                                            v_ccosto = Mapeo(v_legajo, "LEGAJO_POSICION", "CENTRO_COSTO", objAuditoria!audnro, "[CCOSTO]", objrS1!estrnro, sin_error)
                                            If sin_error Then
                                                Call InsertarItemLegajoPosicion("CCOSTO", v_legajo, CDate(v_fecha), "", "", "", "", "", "", v_ccosto, v_estado, "", "", Empty, "", "", "", "", "", "", sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
                                            End If
                                        Case C_estruct_area:
                                            v_estruct_area = Mapeo(v_legajo, "LEGAJO_POSICION", "ESTRUC_AREA", objAuditoria!audnro, "[ESTRUC_A]", objrS1!estrnro, sin_error)
                                            If sin_error Then
                                                Call InsertarItemLegajoPosicion("EST_AREA", v_legajo, CDate(v_fecha), "", "", "", "", "", "", "", "", v_estruct_area, "", Empty, "", "", "", v_estado, "", "", sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
                                            End If
                                        Case C_centro_trabajo:
                                            v_centro_trabajo = Mapeo(v_legajo, "LEGAJO_POSICION", "CENTRO_TRABAJO", objAuditoria!audnro, "[C_TRABAJ]", objrS1!estrnro, sin_error)
                                            If sin_error Then
                                                Call InsertarItemLegajoPosicion("CEN_TRAB", v_legajo, CDate(v_fecha), "", "", "", "", "", "", "", "", "", "", Empty, v_centro_trabajo, v_estado, "", "", "", "", sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
                                            End If
                                        Case C_edificio:
                                            v_edificio = Mapeo(v_legajo, "LEGAJO_POSICION", "EDIFICIO", objAuditoria!audnro, "[EDIFICIO]", objrS1!estrnro, sin_error)
                                            If sin_error Then
                                                Call InsertarItemLegajoPosicion("EDIFICIO", v_legajo, CDate(v_fecha), "", "", "", "", "", "", "", "", "", "", Empty, "", "", v_edificio, "", "", "", sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
                                            End If
                                            StrSql = "SELECT provnro FROM sucursal INNER JOIN cabdom ON sucursal.ternro=cabdom.ternro "
                                            StrSql = StrSql & " INNER JOIN detdom ON cabdom.domnro = detdom.domnro "
                                            StrSql = StrSql & " WHERE estrnro = " & objrS1!estrnro & " AND domdefault=-1"
                                            OpenRecordset StrSql, objrS1
                                            If Not objrS1.EOF Then
                                                v_prov_trabaja = Mapeo(v_legajo, "LEGAJO_POSICION", "PROVINCIA_TRABAJA", objAuditoria!audnro, "[PROV_TRA]", objrS1!provnro, sin_error)
                                                If sin_error Then
                                                    Call InsertarItemLegajoLiquidar("PROVINCIA_TRABAJA", v_legajo, "", Empty, Empty, Empty, "", "", "", "", "", "", "", 0, 0, "", "", "", "", "", "", Empty, "", v_prov_trabaja, "", "", "", "", Empty, "", sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
                                                End If
                                            End If
                                        Case C_canal_costo:
                                            v_canal_costo = Mapeo(v_legajo, "LEGAJO_POSICION", "CANAL_COSTO", objAuditoria!audnro, "[CANAL_CO]", objrS1!estrnro, sin_error)
                                            If sin_error Then
                                                Call InsertarItemLegajoPosicion("CANAL_C", v_legajo, CDate(v_fecha), "", "", "", "", "", "", "", "", "", "", Empty, "", "", "", "", v_canal_costo, v_estado, sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
                                            End If
                                        Case C_rel_laboral:
                                            v_rel_laboral = Mapeo(v_legajo, "LEGAJO_LIQUIDAR", "RELACION_LABORAL", objAuditoria!audnro, "[REL_LABO]", objrS1!estrnro, sin_error)
                                            If sin_error Then
                                                Call InsertarItemLegajoLiquidar("RELACION_LABORAL", v_legajo, "", Empty, Empty, Empty, "", v_rel_laboral, "", "", "", "", "", 0, 0, "", "", "", "", "", "", Empty, "", "", "", "", "", "", Empty, "", sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
                                            End If
                                        Case C_dgi_exento:
                                            v_dgi_exento = Mapeo(v_legajo, "LEGAJO_LIQUIDAR", "DGI_EXENTO", objAuditoria!audnro, "[DGI_EXEN]", objrS1!estrnro, sin_error)
                                            If sin_error Then
                                                Call InsertarItemLegajoLiquidar("DGI_EXENTO", v_legajo, "", Empty, Empty, Empty, "", "", v_dgi_exento, "", "", "", "", 0, 0, "", "", "", "", "", "", Empty, "", "", "", "", "", "", Empty, "", sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
                                            End If
                                        Case C_convenio:
                                            v_convenio = Mapeo(v_legajo, "LEGAJO_LIQUIDAR", "CONVENIO", objAuditoria!audnro, "[CONVENIO]", objrS1!estrnro, sin_error)
                                            If sin_error Then
                                                Call InsertarItemLegajoLiquidar("CONVENIO", v_legajo, "", Empty, Empty, Empty, "", "", "", "", v_convenio, "", "", 0, 0, "", "", "", "", "", "", Empty, "", "", "", "", "", "", Empty, "", sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
                                            End If
                                        Case C_cat_conv:
                                            v_cat_conv = Mapeo(v_legajo, "LEGAJO_LIQUIDAR", "CATEGORIA_CONVENIO", objAuditoria!audnro, "[CAT_CONV]", objrS1!estrnro, sin_error)
                                            If sin_error Then
                                                Call InsertarItemLegajoLiquidar("CATEGORIA_CONVENIO", v_legajo, "", Empty, Empty, Empty, "", "", "", "", "", v_cat_conv, "", 0, 0, "", "", "", "", "", "", Empty, "", "", "", "", "", "", Empty, "", sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
                                            End If
                                        Case C_fun_conv:
                                            v_fun_conv = Mapeo(v_legajo, "LEGAJO_LIQUIDAR", "FUNCION_CONVENIO", objAuditoria!audnro, "[FUN_CONV]", objrS1!estrnro, sin_error)
                                            If sin_error Then
                                                Call InsertarItemLegajoLiquidar("FUNCION_CONVENIO", v_legajo, "", Empty, Empty, Empty, "", "", "", "", "", "", v_fun_conv, 0, 0, "", "", "", "", "", "", Empty, "", "", "", "", "", "", Empty, "", sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
                                            End If
                                        Case C_horas_mesh:
                                            v_horas_mesh = Mapeo(v_legajo, "LEGAJO_LIQUIDAR", "HORAS_MESH", objAuditoria!audnro, "[HOR_MESH]", objrS1!estrnro, sin_error)
                                            If sin_error Then
                                                Call InsertarItemLegajoLiquidar("HORAS_MESH", v_legajo, "", Empty, Empty, Empty, "", "", "", "", "", "", "", CDbl(v_horas_mesh), CDbl(v_horas_mesh), "", "", "", "", "", "", Empty, "", "", "", "", "", "", Empty, "", sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
                                            End If
    '                                    Case C_horas_mesd:
    '                                        v_horas_mesd = Mapeo(v_legajo, "LEGAJO_LIQUIDAR", "HORAS_MESD", objAuditoria!audnro, "[HOR_MESD]", objrS1!estrnro, sin_error)
    '                                        If sin_error Then
    '                                            Call InsertarItemLegajoLiquidar("HORAS_MESD", v_legajo, "", Empty, Empty, Empty, "", "", "", "", "", "", "", 0, CDbl(v_horas_mesd), "", "", "", "", "", "", Empty, "", "", "", "", "", "", Empty, "", sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
    '                                        End If
                                        Case C_osocial:
                                            v_osocial = Mapeo(v_legajo, "LEGAJO_LIQUIDAR", "OSOCIAL", objAuditoria!audnro, "[OSOCIAL]", objrS1!estrnro, sin_error)
                                            If sin_error Then
                                                Call InsertarItemLegajoLiquidar("OSOCIAL", v_legajo, "", Empty, Empty, Empty, "", "", "", "", "", "", "", 0, 0, "", "", "", "", "", v_osocial, Empty, "", "", "", "", "", "", Empty, "", sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
                                            End If
                                        Case C_caracter_serv:
                                            v_caracter_serv = Mapeo(v_legajo, "LEGAJO_LIQUIDAR", "CARACTER_SERVICIO", objAuditoria!audnro, "[CAR_SERV]", objrS1!estrnro, sin_error)
                                            If sin_error Then
                                                Call InsertarItemLegajoLiquidar("CARACTER_SERVICIO", v_legajo, "", Empty, Empty, Empty, "", "", "", "", "", "", "", 0, 0, "", "", "", "", "", "", Empty, "", "", v_caracter_serv, "", "", "", Empty, "", sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
                                            End If
                                        Case C_grupo_liq:
                                            v_grupo_liq = Mapeo(v_legajo, "LEGAJO_LIQUIDAR", "GRUPO_LIQUIDACION", objAuditoria!audnro, "[GRUP_LIQ]", objrS1!estrnro, sin_error)
                                            If sin_error Then
                                                Call InsertarItemLegajoLiquidar("GRUPO_LIQUIDACION", v_legajo, "", Empty, Empty, Empty, "", "", "", "", "", "", "", 0, 0, "", "", "", "", "", "", Empty, "", "", "", "", "", v_grupo_liq, Empty, "", sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
                                            End If
                                        Case C_instrum_pago:
                                            v_instrum_pago = Mapeo(v_legajo, "LEGAJO_LIQUIDAR", "INSTRUM_PAGO", objAuditoria!audnro, "[INS_PAGO]", objrS1!estrnro, sin_error)
                                            If sin_error Then
                                                Call InsertarItemLegajoLiquidar("INSTRUM_PAGO", v_legajo, "", Empty, Empty, Empty, "", "", "", "", "", "", "", 0, 0, v_instrum_pago, "", "", "", "", "", Empty, "", "", "", "", "", "", Empty, "", sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
                                            End If
                                    End Select
                                Else
                                    Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", objAuditoria!audnro, True, StrSql)
                                End If
                            End If
                        End If
                    End If
                
'--------'62, 'Alta Familiar'
                Case 62:
                    Flog.writeline Espacios(Tabulador * 0) & "  Tipo " & objAuditoria!caudnro & " - Campo 262 - Auditoria Nro --> " & objAuditoria!audnro & " - " & objAuditoria!aud_des
                    StrSql = "SELECT tercero.terape, tercero.ternom, tercero.tersex, tercero.estcivnro, familiar.famfec, familiar.parenro "
                    StrSql = StrSql & " ,tercero.terfecnac,tercero.paisnro,lugar_nac.lugardesc,tercero.nacionalnro,empleado.empleg "
                    StrSql = StrSql & " FROM tercero INNER JOIN familiar ON tercero.ternro=familiar.ternro "
                    StrSql = StrSql & " INNER JOIN empleado ON familiar.empleado=empleado.ternro "
                    StrSql = StrSql & " LEFT JOIN lugar_nac ON tercero.lugarnro=lugar_nac.lugarnro "
                    StrSql = StrSql & " WHERE tercero.ternro = " & objAuditoria!aud_ternro
                    If objAuditoria!aud_ternro = "" Or IsNull(objAuditoria!aud_ternro) Then
                        Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", objAuditoria!audnro, True, StrSql)
                    Else
                        OpenRecordset StrSql, objrS1
                        If Not objrS1.EOF Then
                            v_legajo = objrS1!empleg
                            v_apellido_nombre = objrS1!terape & " / " & objrS1!ternom
                            If objrS1!tersex = -1 Then
                                v_sexo = "M"
                            Else
                                v_sexo = "F"
                            End If
                            v_estcivil = Mapeo(v_legajo, "LEGAJO_FAMILIARES", "ESTADO_CIVIL", objAuditoria!audnro, "[ESTCIVIL]", objrS1!estcivnro, sin_error)
                            If sin_error Then
                                v_fech_nac = objrS1!terfecnac
                                v_nacion_nac = Mapeo(v_legajo, "LEGAJO_FAMILIARES", "NACION_NACIMIENTO", objAuditoria!audnro, "[PAIS]", objrS1!paisnro, sin_error)
                                If sin_error Then
                                    v_lugar_nac = IIf(EsNulo(objrS1!lugardesc), "", objrS1!lugardesc)
                                    v_nacionalidad = Mapeo(v_legajo, "LEGAJO_FAMILIARES", "NACIONALIDAD", objAuditoria!audnro, "[NACIONAL]", objrS1!nacionalnro, sin_error)
                                    If sin_error Then
                                        v_fech_ini_vinc = objrS1!famfec
                                        v_parentesco = Mapeo(v_legajo, "LEGAJO_FAMILIARES", "PARENTESCO", objAuditoria!audnro, "[PARENTES]", objrS1!parenro, sin_error)
                                        If sin_error Then
                                            Call InsertarItemLegajoFamiliares("ALTA_FAM", v_legajo, v_parentesco, v_apellido_nombre, CDate(v_fech_nac), v_nacionalidad, v_nacion_nac, v_lugar_nac, "", "", v_sexo, v_estcivil, "", CDate(v_fech_ini_vinc), sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
                                        End If
                                    End If
                                End If
                            End If
                        Else
                            Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", objAuditoria!audnro, True, StrSql)
                        End If
                    End If
                
'--------63, 'Modificación Familiar'
                Case 63:
                    Select Case objAuditoria!aud_campnro
                        Case 259:
                            '259, 'Familiar - Fecha Inicio (famfec)'
                            Flog.writeline Espacios(Tabulador * 0) & "  Tipo " & objAuditoria!caudnro & " - Campo 259 - Auditoria Nro --> " & objAuditoria!audnro & " - " & objAuditoria!aud_des
                            StrSql = "SELECT familiar.famfec, empleado.empleg "
                            StrSql = StrSql & " FROM familiar "
                            StrSql = StrSql & " INNER JOIN empleado ON familiar.empleado=empleado.ternro "
                            StrSql = StrSql & " WHERE familiar.ternro = " & objAuditoria!aud_ternro
                            If objAuditoria!aud_ternro = "" Or IsNull(objAuditoria!aud_ternro) Then
                                 Call InsertarError("0", "GENERAL", "GENERAL", "Error. El campo aud_ternro es nulo. No se puede procesar.", objAuditoria!audnro, True, StrSql)
                            Else
                                 OpenRecordset StrSql, objrS1
                                 If Not objrS1.EOF Then
                                     v_legajo = objrS1!empleg
                                     v_fech_ini_vinc = objrS1!famfec
                                     Call InsertarItemLegajoFamiliares("F_INI_VINC_F", v_legajo, "", "", Empty, "", "", "", "", "", "", "", "", CDate(v_fech_ini_vinc), sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
                                 Else
                                     Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", objAuditoria!audnro, True, StrSql)
                                 End If
                            End If
                        
                        Case 68:
                            '68, 'Parentesco del familiar', 'parenro', 'familiar'
                            Flog.writeline Espacios(Tabulador * 0) & "  Tipo " & objAuditoria!caudnro & " - Campo 68 - Auditoria Nro --> " & objAuditoria!audnro & " - " & objAuditoria!aud_des
                            StrSql = "SELECT familiar.parenro, empleado.empleg "
                            StrSql = StrSql & " FROM familiar "
                            StrSql = StrSql & " INNER JOIN empleado ON familiar.empleado=empleado.ternro "
                            StrSql = StrSql & " WHERE familiar.ternro = " & objAuditoria!aud_ternro
                            If objAuditoria!aud_ternro = "" Or IsNull(objAuditoria!aud_ternro) Then
                                Call InsertarError("0", "GENERAL", "GENERAL", "Error. El campo aud_ternro es nulo. No se puede procesar.", objAuditoria!audnro, True, StrSql)
                            Else
                                OpenRecordset StrSql, objrS1
                                If Not objrS1.EOF Then
                                    v_legajo = objrS1!empleg
                                    v_parentesco = Mapeo(v_legajo, "LEGAJO_FAMILIARES", "PARENTESCO", objAuditoria!audnro, "[PARENT]", objrS1!parenro, sin_error)
                                    If sin_error Then
                                        Call InsertarItemLegajoFamiliares("PARENT_F", v_legajo, v_parentesco, "", Empty, "", "", "", "", "", "", "", "", Empty, sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
                                    End If
                                Else
                                    Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", objAuditoria!audnro, True, StrSql)
                                End If
                            End If
                    End Select
                    
'--------'79, 'Alta Fase (altas y Bajas)'
                Case 79:
                    Flog.writeline Espacios(Tabulador * 0) & "  Tipo " & objAuditoria!caudnro & " - Auditoria Nro --> " & objAuditoria!audnro & " - " & objAuditoria!aud_des
                    StrSql = "SELECT empleado.empleg, fases.altfec, fases.bajfec, fases.fasrecofec, fases.real, fases.vacaciones"
                    StrSql = StrSql & " FROM fases INNER JOIN empleado ON fases.empleado=empleado.ternro "
                    StrSql = StrSql & " WHERE empleado.ternro = " & objAuditoria!aud_ternro & " AND fases.fasnro = " & objAuditoria!aud_actual
                    If objAuditoria!aud_ternro = "" Or IsNull(objAuditoria!aud_ternro) Or EsNulo(objAuditoria!aud_actual) Then
                        Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", objAuditoria!audnro, True, StrSql)
                    Else
                        OpenRecordset StrSql, objrS1
                        If Not objrS1.EOF Then
                            v_legajo = objrS1!empleg
                            If CBool(objrS1!fasrecofec) Then
                                v_Fecha_Antiguedad = objrS1!altfec
                                v_tipo = "FECHA_ANTIGUEDAD,"
                            End If
                            If CBool(objrS1!Real) Then
                                v_Ingreso_Grupo = objrS1!altfec
                                v_tipo = v_tipo & "INGRESO_GRUPO,"
                            End If
                            If CBool(objrS1!vacaciones) Then
                                v_Fecha_Promo = objrS1!altfec
                                v_tipo = v_tipo & "FECHA_PROMO_VAC,"
                            End If
'                            informar_baja = False
                            If Not EsNulo(objrS1!bajfec) Then
                                StrSql = "SELECT empleado.empleg, fases.bajfec, fases.caunro "
                                StrSql = StrSql & " FROM fases INNER JOIN empleado ON fases.empleado=empleado.ternro "
                                StrSql = StrSql & " WHERE empleado.ternro = " & objAuditoria!aud_ternro & " ORDER BY fases.altfec DESC "
                                OpenRecordset StrSql, objrS1
                                If Not objrS1.EOF Then
                                   If Not EsNulo(objrS1!bajfec) Then
                                        v_Fecha_Baja = objrS1!bajfec
                                        If Not EsNulo(objrS1!caunro) Then
                                            v_Motivo_Baja = Mapeo(v_legajo, "LEGAJO_LIQUIDAR", "MOTIVO_BAJA", objAuditoria!audnro, "[MOT_BAJA]", objrS1!caunro, sin_error) ' Hacer mapeo
                                        End If
                                        If sin_error Then
                                            v_tipo = v_tipo & "FECHA_BAJA,"
'                                            informar_baja = True
                                        End If
                                    End If
                                End If
                            End If
                            
                            If sin_error And Not EsNulo(v_tipo) Then
                                v_tipo = Left(v_tipo, Len(v_tipo) - 1)
                                Call InsertarItemLegajoLiquidar(v_tipo, v_legajo, "", Empty, CDate(v_Fecha_Antiguedad), CDate(v_Ingreso_Grupo), "", "", "", "", "", "", "", 0, 0, "", "", "", "", "", "", CDate(v_Fecha_Baja), v_Motivo_Baja, "", "", "", "", "", CDate(v_Fecha_Promo), "", sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
                            End If
                            
'                            If informar_baja Then
'                                Call InsertarItemLegajoLiquidar("BAJ_MOD_FASE", v_legajo, "", Empty, v_Fecha_Antiguedad, v_Ingreso_Grupo, "", "", "", "", "", "", "", 0, 0, "", "", "", "", "", "", v_Fecha_Baja, v_Motivo_Baja, "", "", "", "", "", v_Fecha_Promo, "", sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
'                            Else
'                                Call InsertarItemLegajoLiquidar("ALT_MOD_FASE", v_legajo, "", Empty, v_Fecha_Antiguedad, v_Ingreso_Grupo, "", "", "", "", "", "", "", 0, 0, "", "", "", "", "", "", Empty, "", "", "", "", "", "", v_Fecha_Promo, "", sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
'                            End If
                        Else
                            Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", objAuditoria!audnro, True, StrSql)
                        End If
                    End If
                            
'--------'80, 'Baja Fase (altas y Bajas)'
                Case 80:
'                    StrSql = "SELECT empleado.empleg, fases.altfec, fases.bajfec, fases.fasrecofec, fases.real, fases.vacaciones"
'                    StrSql = StrSql & " FROM fases INNER JOIN empleado ON fases.empleado=empleado.ternro "
'                    StrSql = StrSql & " WHERE empleado.ternro = " & objAuditoria!aud_ternro & " AND fases.fasnro = " & objAuditoria!aud_ant
'                    If objAuditoria!aud_ternro = "" Or IsNull(objAuditoria!aud_ternro) Then
'                        Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", objAuditoria!audnro, True, StrSql)
'                    Else
'                        OpenRecordset StrSql, objrS1
'                        If Not objrS1.EOF Then
'                            If CBool(objrS1!fasrecofec) Then
'                                v_Fecha_Antiguedad = Empty
''                                Call InsertarItemLegajoLiquidar("FECHA_ANTIGUEDAD", v_legajo, "", Empty, v_Fecha_Antiguedad, Empty, "", "", "", "", "", "", "", 0, 0, "", "", "", "", "", "", Empty, "", "", "", "", "", "", Empty, "", sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
'                            End If
'                            If CBool(objrS1!Real) Then
'                                v_Ingreso_Grupo = Empty
''                                Call InsertarItemLegajoLiquidar("INGRESO_GRUPO", v_legajo, "", Empty, Empty, v_Ingreso_Grupo, "", "", "", "", "", "", "", 0, 0, "", "", "", "", "", "", Empty, "", "", "", "", "", "", Empty, "", sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
'                            End If
'                            If CBool(objrS1!vacaciones) Then
'                                v_Fecha_Promo = Empty
''                                Call InsertarItemLegajoLiquidar("FECHA_PROMO_VAC", v_legajo, "", Empty, Empty, Empty, "", "", "", "", "", "", "", 0, 0, "", "", "", "", "", "", Empty, "", "", "", "", "", "", v_Fecha_Promo, "", sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
'                            End If
'                            v_Fecha_Baja = Empty
'                            v_Motivo_Baja = ""
'                            If Not EsNulo(objrS1!bajfec) Then
'                                StrSql = "SELECT empleado.empleg, fases.bajfec, fases.caunro "
'                                StrSql = StrSql & " FROM fases INNER JOIN empleado ON fases.empleado=empleado.ternro "
'                                StrSql = StrSql & " WHERE empleado.ternro = " & objAuditoria!aud_ternro & " ORDER BY fases.altfec DESC "
'                                OpenRecordset StrSql, objrS1
'                                If Not objrS1.EOF Then
'                                   If Not EsNulo(objrS1!bajfec) Then
'                                        v_Fecha_Baja = objrS1!bajfec
'                                        v_Motivo_Baja = Mapeo(v_legajo, "LEGAJO_LIQUIDAR", "MOTIVO_BAJA", objAuditoria!audnro, "[MOT_BAJA]", objrS1!caunro, sin_error) ' Hacer mapeo
''                                        If sin_error Then
''                                            Call InsertarItemLegajoLiquidar("FECHA_BAJA", v_legajo, "", "", Empty, Empty, "", "", "", "", "", "", "", 0, 0, "", "", "", "", "", "", v_Fecha_Baja, v_Motivo_Baja, "", "", "", "", "", "", "", sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
''                                        End If
'                                    End If
'                                Else
''                                    v_Fecha_Baja = Empty
''                                    v_Motivo_Baja = ""
''                                    Call InsertarItemLegajoLiquidar("FECHA_BAJA", v_legajo, "", "", Empty, Empty, "", "", "", "", "", "", "", 0, 0, "", "", "", "", "", "", v_Fecha_Baja, v_Motivo_Baja, "", "", "", "", "", v_Fecha_Promo, "", sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
'                                End If
'                            End If
'                            If sin_error Then
'                                Call InsertarItemLegajoLiquidar("FECHA_BAJA", v_legajo, "", "", Empty, Empty, "", "", "", "", "", "", "", 0, 0, "", "", "", "", "", "", v_Fecha_Baja, v_Motivo_Baja, "", "", "", "", "", "", "", sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
'                            End If
'                        Else
'                            Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", objAuditoria!audnro, True, StrSql)
'                        End If
'                    End If
'
'--------'81, 'Modif. Fase (altas y Bajas)'
                Case 81:
                    Select Case objAuditoria!aud_campnro
                        Case 80:
                            '80, 'Número de Causa', 'caunro'
                            Flog.writeline Espacios(Tabulador * 0) & "  Tipo " & objAuditoria!caudnro & " - Campo 80 - Auditoria Nro --> " & objAuditoria!audnro & " - " & objAuditoria!aud_des
                            StrSql = "SELECT empleado.empleg, fases.fasnro, fases.caunro "
                            StrSql = StrSql & " FROM fases INNER JOIN empleado ON fases.empleado=empleado.ternro "
                            StrSql = StrSql & " WHERE empleado.ternro = " & objAuditoria!aud_ternro & " ORDER BY altfec DESC "
                            If objAuditoria!aud_ternro = "" Or IsNull(objAuditoria!aud_ternro) Then
                                Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", objAuditoria!audnro, True, StrSql)
                            Else
                                OpenRecordset StrSql, objrS1
                                If Not objrS1.EOF Then
                                    v_legajo = objrS1!empleg
                                    If objAuditoria!aud_codigo = objrS1!fasnro Then
                                        v_Motivo_Baja = Mapeo(v_legajo, "LEGAJO_LIQUIDAR", "MOTIVO_BAJA", objAuditoria!audnro, "[MOT_BAJA]", objrS1!caunro, sin_error) ' Hacer mapeo
                                        If sin_error Then
                                            Call InsertarItemLegajoLiquidar("MOTIVO_BAJA", v_legajo, "", Empty, Empty, Empty, "", "", "", "", "", "", "", 0, 0, "", "", "", "", "", "", Empty, v_Motivo_Baja, "", "", "", "", "", Empty, "", sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
                                        End If
                                    End If
                                Else
                                    Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", objAuditoria!audnro, True, StrSql)
                                End If
                            End If
                        Case 82:
                            '82, 'Fecha de alta de la fase'
                            Flog.writeline Espacios(Tabulador * 0) & "  Tipo " & objAuditoria!caudnro & " - Campo 82 - Auditoria Nro --> " & objAuditoria!audnro & " - " & objAuditoria!aud_des
                            StrSql = "SELECT empleado.empleg, fases.altfec, fases.fasrecofec, fases.real, fases.vacaciones "
                            StrSql = StrSql & " FROM fases INNER JOIN empleado ON fases.empleado=empleado.ternro "
                            StrSql = StrSql & " WHERE empleado.ternro = " & objAuditoria!aud_ternro & " AND fases.altfec = " & ConvFecha(objAuditoria!AUD_ANT)
                            If objAuditoria!aud_ternro = "" Or IsNull(objAuditoria!aud_ternro) Then
                                Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", objAuditoria!audnro, True, StrSql)
                            Else
                                OpenRecordset StrSql, objrS1
                                If Not objrS1.EOF Then
                                    v_legajo = objrS1!empleg
                                    v_Fecha_Antiguedad = Empty
                                    v_Ingreso_Grupo = Empty
                                    v_Fecha_Promo = Empty
                                    If CBool(objrS1!fasrecofec) Then
                                        v_Fecha_Antiguedad = objrS1!altfec
                                        v_tipo = "FECHA_ANTIGUEDAD,"
'                                        Call InsertarItemLegajoLiquidar("FECHA_ANTIGUEDAD", v_legajo, "", Empty, v_Fecha_Antiguedad, Empty, "", "", "", "", "", "", "", 0, 0, "", "", "", "", "", "", Empty, "", "", "", "", "", "", Empty, "", sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
                                    End If
                                    If CBool(objrS1!Real) Then
                                        v_Ingreso_Grupo = objrS1!altfec
                                        v_tipo = v_tipo & "INGRESO_GRUPO,"
'                                        Call InsertarItemLegajoLiquidar("INGRESO_GRUPO", v_legajo, "", Empty, Empty, v_Ingreso_Grupo, "", "", "", "", "", "", "", 0, 0, "", "", "", "", "", "", Empty, "", "", "", "", "", "", Empty, "", sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
                                    End If
                                    If CBool(objrS1!vacaciones) Then
                                        v_Fecha_Promo = objrS1!altfec
                                        v_tipo = v_tipo & "FECHA_PROMO_VAC,"
'                                        Call InsertarItemLegajoLiquidar("FECHA_PROMO_VAC", v_legajo, "", Empty, Empty, Empty, "", "", "", "", "", "", "", 0, 0, "", "", "", "", "", "", Empty, "", "", "", "", "", "", v_Fecha_Promo, "", sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
                                    End If
                                    If Not EsNulo(v_tipo) Then
                                        v_tipo = Left(v_tipo, Len(v_tipo) - 1)
                                        Call InsertarItemLegajoLiquidar(v_tipo, v_legajo, "", Empty, CDate(v_Fecha_Antiguedad), CDate(v_Ingreso_Grupo), "", "", "", "", "", "", "", 0, 0, "", "", "", "", "", "", Empty, "", "", "", "", "", "", CDate(v_Fecha_Promo), "", sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
                                    End If
                                Else
                                    Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", objAuditoria!audnro, True, StrSql)
                                End If
                            End If
                        Case 83:
                            '83, 'Fecha de baja de la fase'
                            Flog.writeline Espacios(Tabulador * 0) & "  Tipo " & objAuditoria!caudnro & " - Campo 83 - Auditoria Nro --> " & objAuditoria!audnro & " - " & objAuditoria!aud_des
                            StrSql = "SELECT empleado.empleg, fases.bajfec, fases.caunro "
                            StrSql = StrSql & " FROM fases INNER JOIN empleado ON fases.empleado=empleado.ternro "
                            StrSql = StrSql & " WHERE empleado.ternro = " & objAuditoria!aud_ternro & " ORDER BY fases.altfec DESC"
                            If objAuditoria!aud_ternro = "" Or IsNull(objAuditoria!aud_ternro) Then
                                Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", objAuditoria!audnro, True, StrSql)
                            Else
                                OpenRecordset StrSql, objrS1
                                If Not objrS1.EOF Then
                                    v_legajo = objrS1!empleg
                                    If Not EsNulo(objrS1!bajfec) Then
                                        v_Fecha_Baja = objrS1!bajfec
                                        v_Motivo_Baja = Mapeo(v_legajo, "LEGAJO_LIQUIDAR", "MOTIVO_BAJA", objAuditoria!audnro, "[MOT_BAJA]", objrS1!caunro, sin_error) ' Hacer mapeo
                                        If sin_error Then
                                            Call InsertarItemLegajoLiquidar("FECHA_BAJA", v_legajo, "", Empty, Empty, Empty, "", "", "", "", "", "", "", 0, 0, "", "", "", "", "", "", CDate(v_Fecha_Baja), v_Motivo_Baja, "", "", "", "", "", Empty, "", sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
                                        End If
                                    End If
                                Else
                                    Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", objAuditoria!audnro, True, StrSql)
                                End If
                            End If
                        Case 260:
                            '260, 'Marca Fecha Alta Rec. en fases (fasrecofec)'
                            Flog.writeline Espacios(Tabulador * 0) & "  Tipo " & objAuditoria!caudnro & " - Campo 260 - Auditoria Nro --> " & objAuditoria!audnro & " - " & objAuditoria!aud_des
                            StrSql = "SELECT empleado.empleg, fases.altfec "
                            StrSql = StrSql & " FROM fases INNER JOIN empleado ON fases.empleado=empleado.ternro "
                            StrSql = StrSql & " WHERE empleado.ternro = " & objAuditoria!aud_ternro & " AND fases.fasrecofec = -1"
                            If objAuditoria!aud_ternro = "" Or IsNull(objAuditoria!aud_ternro) Then
                                Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", objAuditoria!audnro, True, StrSql)
                            Else
                                OpenRecordset StrSql, objrS1
                                If Not objrS1.EOF Then
                                    v_Fecha_Antiguedad = objrS1!altfec
                                    v_legajo = objrS1!empleg
                                    Call InsertarItemLegajoLiquidar("FECHA_ANTIGUEDAD", v_legajo, "", Empty, CDate(v_Fecha_Antiguedad), Empty, "", "", "", "", "", "", "", 0, 0, "", "", "", "", "", "", Empty, "", "", "", "", "", "", Empty, "", sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
                                End If
                            End If
                        Case 261:
                            '261, 'Marca Fecha Alta Real en fases (real)'
                            Flog.writeline Espacios(Tabulador * 0) & "  Tipo " & objAuditoria!caudnro & " - Campo 261 - Auditoria Nro --> " & objAuditoria!audnro & " - " & objAuditoria!aud_des
                            StrSql = "SELECT empleado.empleg, fases.altfec, fases.vacaciones"
                            StrSql = StrSql & " FROM fases INNER JOIN empleado ON fases.empleado=empleado.ternro "
                            StrSql = StrSql & " WHERE empleado.ternro = " & objAuditoria!aud_ternro & " AND fases.real = -1"
                            If objAuditoria!aud_ternro = "" Or IsNull(objAuditoria!aud_ternro) Then
                                Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", objAuditoria!audnro, True, StrSql)
                            Else
                                OpenRecordset StrSql, objrS1
                                If Not objrS1.EOF Then
                                    v_legajo = objrS1!empleg
                                    v_Ingreso_Grupo = objrS1!altfec
                                    Call InsertarItemLegajoLiquidar("INGRESO_GRUPO", v_legajo, "", Empty, Empty, CDate(v_Ingreso_Grupo), "", "", "", "", "", "", "", 0, 0, "", "", "", "", "", "", Empty, "", "", "", "", "", "", Empty, "", sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
                                End If
                            End If
                        Case 265:
                            '265, 'Marca Vacaciones en fases', 'Vacaciones'
                            Flog.writeline Espacios(Tabulador * 0) & "  Tipo " & objAuditoria!caudnro & " - Campo 265 - Auditoria Nro --> " & objAuditoria!audnro & " - " & objAuditoria!aud_des
                            StrSql = "SELECT empleado.empleg, fases.altfec "
                            StrSql = StrSql & " FROM fases INNER JOIN empleado ON fases.empleado=empleado.ternro "
                            StrSql = StrSql & " WHERE empleado.ternro = " & objAuditoria!aud_ternro & " AND fases.vacaciones = -1"
                            If objAuditoria!aud_ternro = "" Or IsNull(objAuditoria!aud_ternro) Then
                                Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", objAuditoria!audnro, True, StrSql)
                            Else
                                OpenRecordset StrSql, objrS1
                                If Not objrS1.EOF Then
                                    v_legajo = objrS1!empleg
                                    v_Fecha_Promo = objrS1!altfec
                                    Call InsertarItemLegajoLiquidar("FECHA_PROMO_VAC", v_legajo, "", Empty, Empty, Empty, "", "", "", "", "", "", "", 0, 0, "", "", "", "", "", "", Empty, "", "", "", "", "", "", CDate(v_Fecha_Promo), "", sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
                                End If
                            End If
                        
                    End Select
                
'--------'99, 'Alta Cuenta Bancaria'
                Case 99:
                    Flog.writeline Espacios(Tabulador * 0) & "  Tipo " & objAuditoria!caudnro & " - Auditoria Nro --> " & objAuditoria!audnro & " - " & objAuditoria!aud_des
                    StrSql = "SELECT tipnro FROM ter_tip WHERE ternro=" & objAuditoria!aud_ternro
                    If objAuditoria!aud_ternro = "" Or IsNull(objAuditoria!aud_ternro) Then
                        Call InsertarError("0", "GENERAL", "GENERAL", "Error. El campo aud_ternro es nulo. No se puede procesar.", objAuditoria!audnro, True, StrSql)
                    Else
                        OpenRecordset StrSql, objrS1
                        Select Case objrS1!tipnro
                            Case 1, 2:
                                'EMPLEADO
                                StrSql = "SELECT empleg, fpagnro, banco, ctabnro, ctabcbu, ctabsuc FROM ctabancaria "
                                StrSql = StrSql & "INNER JOIN empleado ON empleado.ternro = ctabancaria.ternro "
                                StrSql = StrSql & " WHERE ctabancaria.ternro=" & objAuditoria!aud_ternro & " AND ctabnro='" & objAuditoria!AUD_ANT & "'"
                                OpenRecordset StrSql, objrS1
                                If Not objrS1.EOF Then
                                    v_legajo = objrS1!empleg
                                    v_tipo_cuenta = Mapeo(v_legajo, "LEGAJO_LIQUIDAR", "TIPO_CUENTA", objAuditoria!audnro, "[TIPO_CTA]", objrS1!fpagnro, sin_error)
                                    If sin_error Then
                                        v_entid_pago = Mapeo(v_legajo, "LEGAJO_LIQUIDAR", "ENTID_PAGO", objAuditoria!audnro, "[ENT_PAGO]", objrS1!banco, sin_error)
                                        If sin_error Then
                                            v_nro_cuenta = objrS1!ctabnro
                                            v_cbu = objrS1!ctabcbu
                                            v_suc_pago = objrS1!ctabsuc
                                            Call InsertarItemLegajoLiquidar("ALTA_CTABANCARIA", v_legajo, "", Empty, Empty, Empty, "", "", "", "", "", "", "", 0, 0, "", v_tipo_cuenta, v_entid_pago, v_nro_cuenta, "", "", Empty, "", "", "", "", v_suc_pago, "", Empty, v_cbu, sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
                                        End If
                                    End If
                                Else
                                    Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", objAuditoria!audnro, True, StrSql)
                                End If
                        End Select
                    End If
                
'--------100, 'Modificación Cuenta Bancaria'
                Case 100:
                    Flog.writeline Espacios(Tabulador * 0) & "  Tipo " & objAuditoria!caudnro & " - Auditoria Nro --> " & objAuditoria!audnro & " - " & objAuditoria!aud_des
                    StrSql = "SELECT tipnro FROM ter_tip WHERE ternro=" & objAuditoria!aud_ternro
                    If objAuditoria!aud_ternro = "" Or IsNull(objAuditoria!aud_ternro) Then
                        Call InsertarError("0", "GENERAL", "GENERAL", "Error. El campo aud_ternro es nulo. No se puede procesar.", objAuditoria!audnro, True, StrSql)
                    Else
                        OpenRecordset StrSql, objrS1
                        Select Case objrS1!tipnro
                            Case 1, 2:
                                'EMPLEADO
                                Select Case objAuditoria!aud_campnro
                                    Case 180:
                                        '180, 'Código de la Forma de Pago  de la Cuenta Bancaria', 'fpagnro', 'ctabancaria'
                                        StrSql = "SELECT ctabancaria.fpagnro, empleado.empleg "
                                        StrSql = StrSql & " FROM empleado INNER JOIN ctabancaria ON empleado.ternro = ctabancaria.ternro "
                                        StrSql = StrSql & " WHERE empleado.ternro = " & objAuditoria!aud_ternro
                                        OpenRecordset StrSql, objrS1
                                        If Not objrS1.EOF Then
                                            v_legajo = objrS1!empleg
                                            v_tipo_cuenta = Mapeo(v_legajo, "LEGAJO_LIQUIDAR", "TIPO_CUENTA", objAuditoria!audnro, "[TIPO_CTA]", objrS1!fpagnro, sin_error)
                                            If sin_error Then
                                                Call InsertarItemLegajoLiquidar("TIPO_CUENTA", v_legajo, "", Empty, Empty, Empty, "", "", "", "", "", "", "", 0, 0, "", v_tipo_cuenta, "", "", "", "", Empty, "", "", "", "", "", "", Empty, "", sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
                                            End If
                                        Else
                                            Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", objAuditoria!audnro, True, StrSql)
                                        End If
                                    Case 179:
                                        '179, 'Código del Banco de la Cuenta Bancaria', 'banco', 'ctabancaria'
                                        StrSql = "SELECT ctabancaria.banco, empleado.empleg "
                                        StrSql = StrSql & " FROM empleado INNER JOIN ctabancaria ON empleado.ternro = ctabancaria.ternro "
                                        StrSql = StrSql & " WHERE empleado.ternro = " & objAuditoria!aud_ternro
                                        OpenRecordset StrSql, objrS1
                                        If Not objrS1.EOF Then
                                            v_legajo = objrS1!empleg
                                            v_entid_pago = Mapeo(v_legajo, "LEGAJO_LIQUIDAR", "ENTID_PAGO", objAuditoria!audnro, "[ENT_PAGO]", objrS1!banco, sin_error)
                                            If sin_error Then
                                                Call InsertarItemLegajoLiquidar("ENTID_PAGO", v_legajo, "", Empty, Empty, Empty, "", "", "", "", "", "", "", 0, 0, "", "", v_entid_pago, "", "", "", Empty, "", "", "", "", "", "", Empty, "", sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
                                            End If
                                        Else
                                            Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", objAuditoria!audnro, True, StrSql)
                                        End If
                                    Case 185:
                                        '185, 'Cuenta Bancaria', 'ctabnro', 'ctabancaria'
                                        StrSql = "SELECT ctabancaria.ctabnro, empleado.empleg "
                                        StrSql = StrSql & " FROM empleado INNER JOIN ctabancaria ON empleado.ternro = ctabancaria.ternro "
                                        StrSql = StrSql & " WHERE empleado.ternro = " & objAuditoria!aud_ternro
                                        OpenRecordset StrSql, objrS1
                                        If Not objrS1.EOF Then
                                            v_legajo = objrS1!empleg
                                            v_nro_cuenta = objrS1!ctabnro
                                            Call InsertarItemLegajoLiquidar("NRO_CUENTA", v_legajo, "", Empty, Empty, Empty, "", "", "", "", "", "", "", 0, 0, "", "", "", v_nro_cuenta, "", "", Empty, "", "", "", "", "", "", Empty, "", sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
                                        Else
                                            Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", objAuditoria!audnro, True, StrSql)
                                        End If
                                    Case 182:
                                        '182, 'CBU de la Cuenta Bancaria', 'ctabcbu', 'ctabancaria'
                                        StrSql = "SELECT ctabancaria.ctabcbu, empleado.empleg "
                                        StrSql = StrSql & " FROM empleado INNER JOIN ctabancaria ON empleado.ternro = ctabancaria.ternro "
                                        StrSql = StrSql & " WHERE empleado.ternro = " & objAuditoria!aud_ternro
                                        OpenRecordset StrSql, objrS1
                                        If Not objrS1.EOF Then
                                            v_legajo = objrS1!empleg
                                            v_cbu = objrS1!ctabcbu
                                            Call InsertarItemLegajoLiquidar("CBU", v_legajo, "", Empty, Empty, Empty, "", "", "", "", "", "", "", 0, 0, "", "", "", "", "", "", Empty, "", "", "", "", "", "", Empty, v_cbu, sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
                                        Else
                                            Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", objAuditoria!audnro, True, StrSql)
                                        End If
                                    Case 264:
                                        '264, 'Cód. Sucursal de Cuenta Bancaria', 'ctabsuc', 'ctabancaria'
                                        StrSql = "SELECT ctabancaria.ctabcbu, empleado.empleg "
                                        StrSql = StrSql & " FROM empleado INNER JOIN ctabancaria ON empleado.ternro = ctabancaria.ternro "
                                        StrSql = StrSql & " WHERE empleado.ternro = " & objAuditoria!aud_ternro
                                        OpenRecordset StrSql, objrS1
                                        If Not objrS1.EOF Then
                                            v_legajo = objrS1!empleg
                                            v_suc_pago = objrS1!ctabsuc
                                            Call InsertarItemLegajoLiquidar("SUC_PAGO", v_legajo, "", Empty, Empty, Empty, "", "", "", "", "", "", "", 0, 0, "", "", "", "", "", "", Empty, "", "", "", "", v_suc_pago, "", Empty, "", sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
                                        Else
                                            Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", objAuditoria!audnro, True, StrSql)
                                        End If
                                End Select
                        End Select
                    End If
                
'--------110, 'Mod. Tercero'
                Case 110:
                    Select Case objAuditoria!aud_campnro
                        Case 194:
                            '194, 'Sexo', 'tersex', 'tercero'
                            Flog.writeline Espacios(Tabulador * 0) & "  Tipo " & objAuditoria!caudnro & " - Campo 194 - Auditoria Nro --> " & objAuditoria!audnro & " - " & objAuditoria!aud_des
                            StrSql = "SELECT tipnro FROM ter_tip WHERE ternro=" & objAuditoria!aud_ternro
                            If objAuditoria!aud_ternro = "" Or IsNull(objAuditoria!aud_ternro) Then
                                Call InsertarError("0", "GENERAL", "GENERAL", "Error. El campo aud_ternro es nulo. No se puede procesar.", objAuditoria!audnro, True, StrSql)
                            Else
                                OpenRecordset StrSql, objrS1
                                Select Case objrS1!tipnro
                                    Case 1, 2:
                                        'EMPLEADO
                                        StrSql = "SELECT tercero.tersex, empleado.empleg "
                                        StrSql = StrSql & " FROM tercero INNER JOIN empleado ON tercero.ternro=empleado.ternro "
                                        StrSql = StrSql & " WHERE tercero.ternro = " & objAuditoria!aud_ternro
                                        OpenRecordset StrSql, objrS1
                                        If Not objrS1.EOF Then
                                            v_legajo = objrS1!empleg
                                            If objrS1!tersex = -1 Then
                                                v_sexo = "M"
                                            Else
                                                v_sexo = "F"
                                            End If
                                            Call InsertarItemLegajo("SEXO", v_legajo, "", "", v_sexo, "", "", Empty, "", "", "", "", "", "", sin_error, objAuditoria!audnro)
                                        Else
                                            Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", objAuditoria!audnro, True, StrSql)
                                        End If
                                    Case 3:
                                        'FAMILIAR
                                        StrSql = "SELECT tercero.tersex, empleado.empleg "
                                        StrSql = StrSql & " FROM tercero INNER JOIN familiar ON tercero.ternro = familiar.ternro "
                                        StrSql = StrSql & " INNER JOIN empleado ON familiar.empleado=empleado.ternro "
                                        StrSql = StrSql & " WHERE familiar.ternro = " & objAuditoria!aud_ternro
                                        OpenRecordset StrSql, objrS1
                                        If Not objrS1.EOF Then
                                            v_legajo = objrS1!empleg
                                            If objrS1!tersex = -1 Then
                                                v_sexo = "M"
                                            Else
                                                v_sexo = "F"
                                            End If
                                            Call InsertarItemLegajoFamiliares("SEXO_F", v_legajo, "", "", Empty, "", "", "", "", "", v_sexo, "", "", Empty, sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
                                        Else
                                            Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", objAuditoria!audnro, True, StrSql)
                                        End If
                                End Select
                            End If
                        
                        Case 257:
                            '257, 'Estado Civil - Tercero', 'estcivnro', 'tercero'
                            Flog.writeline Espacios(Tabulador * 0) & "  Tipo " & objAuditoria!caudnro & " - Campo 257 - Auditoria Nro --> " & objAuditoria!audnro & " - " & objAuditoria!aud_des
                            StrSql = "SELECT tipnro FROM ter_tip WHERE ternro=" & objAuditoria!aud_ternro
                            If objAuditoria!aud_ternro = "" Or IsNull(objAuditoria!aud_ternro) Then
                                Call InsertarError("0", "GENERAL", "GENERAL", "Error. El campo aud_ternro es nulo. No se puede procesar.", objAuditoria!audnro, True, StrSql)
                            Else
                                OpenRecordset StrSql, objrS1
                                Select Case objrS1!tipnro
                                    Case 1, 2:
                                        'EMPLEADO
                                        StrSql = "SELECT tercero.estcivnro,empleado.empleg "
                                        StrSql = StrSql & " FROM tercero INNER JOIN empleado ON tercero.ternro=empleado.ternro "
                                        StrSql = StrSql & " WHERE tercero.ternro = " & objAuditoria!aud_ternro
                                        OpenRecordset StrSql, objrS1
                                        If Not objrS1.EOF Then
                                            v_legajo = objrS1!empleg
                                            v_estcivil = Mapeo(v_legajo, "LEGAJOS", "ESTADO_CIVIL", objAuditoria!audnro, "[ESTCIVIL]", objrS1!estcivnro, sin_error)
                                            If sin_error Then
                                                Call InsertarItemLegajo("EST_CIV", v_legajo, "", "", "", v_estcivil, "", Empty, "", "", "", "", "", "", sin_error, objAuditoria!audnro)
                                            End If
                                        Else
                                            Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", objAuditoria!audnro, True, StrSql)
                                        End If
                                    Case 3:
                                        'FAMILIAR
                                        StrSql = "SELECT tercero.estcivnro, empleado.empleg "
                                        StrSql = StrSql & " FROM tercero INNER JOIN familiar ON tercero.ternro = familiar.ternro "
                                        StrSql = StrSql & " INNER JOIN empleado ON familiar.empleado=empleado.ternro "
                                        StrSql = StrSql & " WHERE familiar.ternro = " & objAuditoria!aud_ternro
                                        OpenRecordset StrSql, objrS1
                                        If Not objrS1.EOF Then
                                            v_legajo = objrS1!empleg
                                            v_estcivil = Mapeo(v_legajo, "LEGAJO_FAMILIARES", "ESTADO_CIVIL", objAuditoria!audnro, "[ESTCIVIL]", objrS1!estcivnro, sin_error)
                                            If sin_error Then
                                                Call InsertarItemLegajoFamiliares("EST_CIV_F", v_legajo, "", "", Empty, "", "", "", "", "", "", v_estcivil, "", Empty, sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
                                            End If
                                        Else
                                            Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", objAuditoria!audnro, True, StrSql)
                                        End If
                                End Select
                            End If
                        
                        Case 207:
                            '207, 'Apellido Casada', 'tercasape', 'tercero'
                            Flog.writeline Espacios(Tabulador * 0) & "  Tipo " & objAuditoria!caudnro & " - Campo 207 - Auditoria Nro --> " & objAuditoria!audnro & " - " & objAuditoria!aud_des
                            StrSql = "SELECT tercero.tercasape,empleado.empleg "
                            StrSql = StrSql & " FROM tercero INNER JOIN empleado ON tercero.ternro=empleado.ternro "
                            StrSql = StrSql & " WHERE tercero.ternro = " & objAuditoria!aud_ternro
                            If objAuditoria!aud_ternro = "" Or IsNull(objAuditoria!aud_ternro) Then
                                Call InsertarError("0", "GENERAL", "GENERAL", "Error. El campo aud_ternro es nulo. No se puede procesar.", objAuditoria!audnro, True, StrSql)
                            Else
                                OpenRecordset StrSql, objrS1
                                If Not objrS1.EOF Then
                                    v_legajo = objrS1!empleg
                                    v_apellido_casada = objrS1!tercasape
                                    Call InsertarItemLegajo("AP_CAS", v_legajo, "", "", "", "", v_apellido_casada, Empty, "", "", "", "", "", "", sin_error, objAuditoria!audnro)
                                Else
                                    Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", objAuditoria!audnro, True, StrSql)
                                End If
                            End If
                        
                        Case 193:
                            '193, 'Fecha Nacimiento', 'terfecnac', 'tercero'
                            Flog.writeline Espacios(Tabulador * 0) & "  Tipo " & objAuditoria!caudnro & " - Campo 193 - Auditoria Nro --> " & objAuditoria!audnro & " - " & objAuditoria!aud_des
                            StrSql = "SELECT tipnro FROM ter_tip WHERE ternro=" & objAuditoria!aud_ternro
                            If objAuditoria!aud_ternro = "" Or IsNull(objAuditoria!aud_ternro) Then
                                Call InsertarError("0", "GENERAL", "GENERAL", "Error. El campo aud_ternro es nulo. No se puede procesar.", objAuditoria!audnro, True, StrSql)
                            Else
                                OpenRecordset StrSql, objrS1
                                Select Case objrS1!tipnro
                                    Case 1, 2:
                                        'EMPLEADO
                                        StrSql = "SELECT tercero.terfecnac,empleado.empleg "
                                        StrSql = StrSql & " FROM tercero INNER JOIN empleado ON tercero.ternro=empleado.ternro "
                                        StrSql = StrSql & " WHERE tercero.ternro = " & objAuditoria!aud_ternro
                                        OpenRecordset StrSql, objrS1
                                        If Not objrS1.EOF Then
                                            v_legajo = objrS1!empleg
                                            v_fech_nac = objrS1!terfecnac
                                            Call InsertarItemLegajo("F_NAC", v_legajo, "", "", "", "", "", v_fech_nac, "", "", "", "", "", "", sin_error, objAuditoria!audnro)
                                        Else
                                            Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", objAuditoria!audnro, True, StrSql)
                                        End If
                                    Case 3:
                                        'FAMILIAR
                                        StrSql = "SELECT tercero.terfecnac, empleado.empleg "
                                        StrSql = StrSql & " FROM tercero INNER JOIN familiar ON tercero.ternro = familiar.ternro "
                                        StrSql = StrSql & " INNER JOIN empleado ON familiar.empleado=empleado.ternro "
                                        StrSql = StrSql & " WHERE familiar.ternro = " & objAuditoria!aud_ternro
                                        OpenRecordset StrSql, objrS1
                                        If Not objrS1.EOF Then
                                            v_legajo = objrS1!empleg
                                            v_fech_nac = objrS1!terfecnac
                                            Call InsertarItemLegajoFamiliares("F_NAC_F", v_legajo, "", "", CDate(v_fech_nac), "", "", "", "", "", "", "", "", Empty, sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
                                        Else
                                            Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", objAuditoria!audnro, True, StrSql)
                                        End If
                                End Select
                            End If
                        
                        Case 250:
                            '250, 'Pais Nac. Tercero', 'paisnro', 'tercero'
                            Flog.writeline Espacios(Tabulador * 0) & "  Tipo " & objAuditoria!caudnro & " - Campo 250 - Auditoria Nro --> " & objAuditoria!audnro & " - " & objAuditoria!aud_des
                            StrSql = "SELECT tipnro FROM ter_tip WHERE ternro=" & objAuditoria!aud_ternro
                            If objAuditoria!aud_ternro = "" Or IsNull(objAuditoria!aud_ternro) Then
                                Call InsertarError("0", "GENERAL", "GENERAL", "Error. El campo aud_ternro es nulo. No se puede procesar.", objAuditoria!audnro, True, StrSql)
                            Else
                                OpenRecordset StrSql, objrS1
                                Select Case objrS1!tipnro
                                    Case 1, 2:
                                        'EMPLEADO
                                        StrSql = "SELECT tercero.paisnro,empleado.empleg "
                                        StrSql = StrSql & " FROM tercero INNER JOIN empleado ON tercero.ternro=empleado.ternro "
                                        StrSql = StrSql & " WHERE tercero.ternro = " & objAuditoria!aud_ternro
                                        OpenRecordset StrSql, objrS1
                                        If Not objrS1.EOF Then
                                            v_legajo = objrS1!empleg
                                            v_nacion_nac = Mapeo(v_legajo, "LEGAJOS", "NACION_NACIMIENTO", objAuditoria!audnro, "[PAIS]", objrS1!paisnro, sin_error)
                                            If sin_error Then
                                                Call InsertarItemLegajo("PAIS", v_legajo, "", "", "", "", "", Empty, v_nacion_nac, "", "", "", "", "", sin_error, objAuditoria!audnro)
                                            End If
                                        Else
                                            Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", objAuditoria!audnro, True, StrSql)
                                        End If
                                    Case 3:
                                        'FAMILIAR
                                        StrSql = "SELECT tercero.paisnro, empleado.empleg "
                                        StrSql = StrSql & " FROM tercero INNER JOIN familiar ON tercero.ternro = familiar.ternro "
                                        StrSql = StrSql & " INNER JOIN empleado ON familiar.empleado=empleado.ternro "
                                        StrSql = StrSql & " WHERE familiar.ternro = " & objAuditoria!aud_ternro
                                        OpenRecordset StrSql, objrS1
                                        If Not objrS1.EOF Then
                                            v_legajo = objrS1!empleg
                                            v_nacion_nac = Mapeo(v_legajo, "LEGAJO_FAMILIARES", "NACION_NACIMIENTO", objAuditoria!audnro, "[PAIS]", objrS1!paisnro, sin_error)
                                            If sin_error Then
                                                Call InsertarItemLegajoFamiliares("PAIS_F", v_legajo, "", "", Empty, "", v_nacion_nac, "", "", "", "", "", "", Empty, sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
                                            End If
                                        Else
                                            Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", objAuditoria!audnro, True, StrSql)
                                        End If
                                End Select
                            End If
                        
                        Case 258:
                            '258, 'Lugar Nacimiento - Tercero', 'lugarnro', 'tercero'
                            Flog.writeline Espacios(Tabulador * 0) & "  Tipo " & objAuditoria!caudnro & " - Campo 258 - Auditoria Nro --> " & objAuditoria!audnro & " - " & objAuditoria!aud_des
                            StrSql = "SELECT tipnro FROM ter_tip WHERE ternro=" & objAuditoria!aud_ternro
                            If objAuditoria!aud_ternro = "" Or IsNull(objAuditoria!aud_ternro) Then
                                Call InsertarError("0", "GENERAL", "GENERAL", "Error. El campo aud_ternro es nulo. No se puede procesar.", objAuditoria!audnro, True, StrSql)
                            Else
                                OpenRecordset StrSql, objrS1
                                Select Case objrS1!tipnro
                                    Case 1, 2:
                                        'EMPLEADO
                                        StrSql = "SELECT lugar_nac.lugardesc,empleado.empleg "
                                        StrSql = StrSql & " FROM tercero INNER JOIN empleado ON tercero.ternro=empleado.ternro "
                                        StrSql = StrSql & " LEFT JOIN lugar_nac ON tercero.lugarnro=lugar_nac.lugarnro "
                                        StrSql = StrSql & " WHERE tercero.ternro = " & objAuditoria!aud_ternro
                                        OpenRecordset StrSql, objrS1
                                        If Not objrS1.EOF Then
                                            v_legajo = objrS1!empleg
                                            v_lugar_nac = objrS1!lugardesc
                                            Call InsertarItemLegajo("LUG_NAC", v_legajo, "", "", "", "", "", Empty, "", v_lugar_nac, "", "", "", "", sin_error, objAuditoria!audnro)
                                        Else
                                            Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", objAuditoria!audnro, True, StrSql)
                                        End If
                                    Case 3:
                                        'FAMILIAR
                                        StrSql = "SELECT lugar_nac.lugardesc, empleado.empleg "
                                        StrSql = StrSql & " FROM tercero INNER JOIN familiar ON tercero.ternro = familiar.ternro "
                                        StrSql = StrSql & " INNER JOIN empleado ON familiar.empleado=empleado.ternro "
                                        StrSql = StrSql & " LEFT JOIN lugar_nac ON tercero.lugarnro=lugar_nac.lugarnro "
                                        StrSql = StrSql & " WHERE familiar.ternro = " & objAuditoria!aud_ternro
                                        OpenRecordset StrSql, objrS1
                                        If Not objrS1.EOF Then
                                            v_legajo = objrS1!empleg
                                            v_lugar_nac = objrS1!lugardesc
                                            Call InsertarItemLegajoFamiliares("LUG_NAC_F", v_legajo, "", "", Empty, "", "", v_lugar_nac, "", "", "", "", "", Empty, sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
                                        Else
                                            Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", objAuditoria!audnro, True, StrSql)
                                        End If
                                End Select
                            End If
                        
                        Case 206:
                            '206, 'Nacionalidad Tercero', 'nacionalnro', 'tercero'
                            Flog.writeline Espacios(Tabulador * 0) & "  Tipo " & objAuditoria!caudnro & " - Campo 206 - Auditoria Nro --> " & objAuditoria!audnro & " - " & objAuditoria!aud_des
                            StrSql = "SELECT tipnro FROM ter_tip WHERE ternro=" & objAuditoria!aud_ternro
                            If objAuditoria!aud_ternro = "" Or IsNull(objAuditoria!aud_ternro) Then
                                Call InsertarError("0", "GENERAL", "GENERAL", "Error. El campo aud_ternro es nulo. No se puede procesar.", objAuditoria!audnro, True, StrSql)
                            Else
                                OpenRecordset StrSql, objrS1
                                Select Case objrS1!tipnro
                                    Case 1, 2:
                                        'EMPLEADO
                                        StrSql = "SELECT tercero.nacionalnro,empleado.empleg "
                                        StrSql = StrSql & " FROM tercero INNER JOIN empleado ON tercero.ternro=empleado.ternro "
                                        StrSql = StrSql & " WHERE tercero.ternro = " & objAuditoria!aud_ternro
                                        OpenRecordset StrSql, objrS1
                                        If Not objrS1.EOF Then
                                            v_legajo = objrS1!empleg
                                            v_nacionalidad = Mapeo(v_legajo, "LEGAJOS", "NACIONALIDAD", objAuditoria!audnro, "[NACIONAL]", objrS1!nacionalnro, sin_error)
                                            If sin_error Then
                                                Call InsertarItemLegajo("NACION", v_legajo, "", "", "", "", "", Empty, "", "", v_nacionalidad, "", "", "", sin_error, objAuditoria!audnro)
                                            End If
                                        Else
                                            Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", objAuditoria!audnro, True, StrSql)
                                        End If
                                    Case 3:
                                        'FAMILIAR
                                        StrSql = "SELECT tercero.nacionalnro, empleado.empleg "
                                        StrSql = StrSql & " FROM tercero INNER JOIN familiar ON tercero.ternro = familiar.ternro "
                                        StrSql = StrSql & " INNER JOIN empleado ON familiar.empleado=empleado.ternro "
                                        StrSql = StrSql & " WHERE familiar.ternro = " & objAuditoria!aud_ternro
                                        OpenRecordset StrSql, objrS1
                                        If Not objrS1.EOF Then
                                            v_legajo = objrS1!empleg
                                            v_nacionalidad = Mapeo(v_legajo, "LEGAJO_FAMILIARES", "NACIONALIDAD", objAuditoria!audnro, "[NACIONAL]", objrS1!nacionalnro, sin_error)
                                            If sin_error Then
                                                Call InsertarItemLegajoFamiliares("NACION_F", v_legajo, "", "", Empty, v_nacionalidad, "", "", "", "", "", "", "", Empty, sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
                                            End If
                                        Else
                                            Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", objAuditoria!audnro, True, StrSql)
                                        End If
                                End Select
                            End If
                            
                        Case 266, 267:
                            '266, 'Nombre Familiar', 'ternom', 'tercero'
                            '267, 'Apellido Familiar', 'terape', 'tercero'
                            Flog.writeline Espacios(Tabulador * 0) & "  Tipo " & objAuditoria!caudnro & " - Campo 266/7 - Auditoria Nro --> " & objAuditoria!audnro & " - " & objAuditoria!aud_des
                            StrSql = "SELECT tercero.terape, tercero.ternom, empleado.empleg "
                            StrSql = StrSql & " FROM tercero INNER JOIN familiar ON tercero.ternro = familiar.ternro "
                            StrSql = StrSql & " INNER JOIN empleado ON familiar.empleado=empleado.ternro "
                            StrSql = StrSql & " WHERE familiar.ternro = " & objAuditoria!aud_ternro
                            If objAuditoria!aud_ternro = "" Or IsNull(objAuditoria!aud_ternro) Then
                                Call InsertarError("0", "GENERAL", "GENERAL", "Error. El campo aud_ternro es nulo. No se puede procesar.", objAuditoria!audnro, True, StrSql)
                            Else
                                OpenRecordset StrSql, objrS1
                                If Not objrS1.EOF Then
                                    v_legajo = objrS1!empleg
                                    v_apellido_nombre = objrS1!terape & " / " & objrS1!ternom
                                    Call InsertarItemLegajoFamiliares("AP_NOM_F", v_legajo, "", v_apellido_nombre, Empty, "", "", "", "", "", "", "", "", Empty, sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
                                Else
                                    Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", objAuditoria!audnro, True, StrSql)
                                End If
                            End If
                    End Select
                
'--------111, 'Mod. Documento'
                Case 111:
                    Select Case objAuditoria!aud_campnro
                        Case 197:
                            '197, 'Tipo Documento', 'tidnro', 'ter_doc';
                            Flog.writeline Espacios(Tabulador * 0) & "  Tipo " & objAuditoria!caudnro & " - Campo 197 - Auditoria Nro --> " & objAuditoria!audnro & " - " & objAuditoria!aud_des
                            StrSql = "SELECT tipnro FROM ter_tip WHERE ternro=" & objAuditoria!aud_ternro
                            If objAuditoria!aud_ternro = "" Or IsNull(objAuditoria!aud_ternro) Then
                                Call InsertarError("0", "GENERAL", "GENERAL", "Error. El campo aud_ternro es nulo. No se puede procesar.", objAuditoria!audnro, True, StrSql)
                            Else
                                OpenRecordset StrSql, objrS1
                                Select Case objrS1!tipnro
                                    Case 1, 2:
                                        'EMPLEADO
                                        StrSql = "SELECT empleado.empleg, ter_doc.tidnro "
                                        StrSql = StrSql & " FROM empleado INNER JOIN ter_doc ON ter_doc.ternro=empleado.ternro "
                                        StrSql = StrSql & " WHERE empleado.ternro = " & objAuditoria!aud_ternro & " AND ter_doc.tidnro = " & objAuditoria!aud_actual
                                        OpenRecordset StrSql, objrS1
                                        If Not objrS1.EOF Then
                                            If objrS1!tidnro <= 5 Then
                                                v_legajo = objrS1!empleg
                                                v_doc_tipo = Mapeo(v_legajo, "LEGAJOS", "DOCUMENTO_TIPO", objAuditoria!audnro, "[TIPODOCU]", objrS1!tidnro, sin_error)
                                                If sin_error Then
                                                    Call InsertarItemLegajo("DOC_TPO", v_legajo, "", "RNP", "", "", "", Empty, "", "", "", v_doc_tipo, "", "", sin_error, objAuditoria!audnro)
                                                End If
                                            ElseIf objrS1!tidnro = 10 Then
                                                'CUIL
                                                v_nro_cuil = objrS1!nrodoc
                                                Call InsertarItemLegajoLiquidar("NRO_CUIL", v_legajo, "", Empty, Empty, Empty, "", "", "", "", "", "", "", 0, 0, "", "", "", "", v_nro_cuil, "", Empty, "", "", "", "", "", "", Empty, "", sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
                                            End If
                                        Else
                                            Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", objAuditoria!audnro, True, StrSql)
                                        End If
                                    Case 3:
                                        'Familiar
                                        StrSql = "SELECT empleado.empleg, ter_doc.tidnro "
                                        StrSql = StrSql & " FROM empleado INNER JOIN familiar ON empleado.ternro=familiar.empleado "
                                        StrSql = StrSql & " INNER JOIN ter_doc ON ter_doc.ternro=familiar.ternro "
                                        StrSql = StrSql & " WHERE familiar.ternro = " & objAuditoria!aud_ternro & " AND ter_doc.tidnro = " & objAuditoria!aud_actual
                                        OpenRecordset StrSql, objrS1
                                        If Not objrS1.EOF Then
                                            If objrS1!tidnro <= 5 Then
                                                'Nro documento
                                                v_legajo = objrS1!empleg
                                                v_doc_tipo = Mapeo(v_legajo, "LEGAJO_FAMILIARES", "DOCUMENTO_TIPO", objAuditoria!audnro, "[TIPODOCU]", objrS1!tidnro, sin_error)
                                                If sin_error Then
                                                    Call InsertarItemLegajoFamiliares("DOC_TPO_F", v_legajo, "", "", Empty, "", "", "", v_doc_tipo, v_doc_nro, "", "", "", Empty, sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
                                                End If
                                            End If
                                        Else
                                            Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", objAuditoria!audnro, True, StrSql)
                                        End If
                                    End Select
                            End If
                        
                        Case 198:
                            '198, 'Número Documento', 'nrodoc', 'ter_doc'
                            Flog.writeline Espacios(Tabulador * 0) & "  Tipo " & objAuditoria!caudnro & " - Campo 198 - Auditoria Nro --> " & objAuditoria!audnro & " - " & objAuditoria!aud_des
                            StrSql = "SELECT tipnro FROM ter_tip WHERE ternro=" & objAuditoria!aud_ternro
                            If objAuditoria!aud_ternro = "" Or IsNull(objAuditoria!aud_ternro) Then
                                Call InsertarError("0", "GENERAL", "GENERAL", "Error. El campo aud_ternro es nulo. No se puede procesar.", objAuditoria!audnro, True, StrSql)
                            Else
                                OpenRecordset StrSql, objrS1
                                Select Case objrS1!tipnro
                                    Case 1, 2:
                                    'EMPLEADO
                                        StrSql = "SELECT empleado.empleg, ter_doc.tidnro, ter_doc.nrodoc "
                                        StrSql = StrSql & " FROM empleado INNER JOIN ter_doc ON ter_doc.ternro=empleado.ternro "
                                        StrSql = StrSql & " WHERE empleado.ternro = " & objAuditoria!aud_ternro & " AND ter_doc.nrodoc = '" & objAuditoria!aud_actual & "'"
                                        OpenRecordset StrSql, objrS1
                                        If Not objrS1.EOF Then
                                            If objrS1!tidnro <= 5 Then
                                                v_legajo = objrS1!empleg
                                                v_doc_nro = objrS1!nrodoc
                                                Call InsertarItemLegajo("DOC_NRO", v_legajo, "", "RNP", "", "", "", Empty, "", "", "", "", v_doc_nro, "", sin_error, objAuditoria!audnro)
                                            End If
                                        Else
                                            Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", objAuditoria!audnro, True, StrSql)
                                        End If
                                    Case 3:
                                        'Familiar
                                        StrSql = "SELECT empleado.empleg, ter_doc.tidnro, ter_doc.nrodoc "
                                        StrSql = StrSql & " FROM empleado INNER JOIN familiar ON empleado.ternro=familiar.empleado "
                                        StrSql = StrSql & " INNER JOIN ter_doc ON ter_doc.ternro=familiar.ternro "
                                        StrSql = StrSql & " WHERE familiar.ternro = " & objAuditoria!aud_ternro & " AND ter_doc.nrodoc = '" & objAuditoria!aud_actual & "'"
                                        OpenRecordset StrSql, objrS1
                                        If Not objrS1.EOF Then
                                            If objrS1!tidnro <= 5 Then
                                                'Nro documento
                                                v_legajo = objrS1!empleg
                                                v_doc_nro = objrS1!nrodoc
                                                Call InsertarItemLegajoFamiliares("DOC_NRO_F", v_legajo, "", "", Empty, "", "", "", "", v_doc_nro, "", "", "", Empty, sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
                                            ElseIf objrS1!tidnro = 10 Then
                                                'CUIL
                                                v_legajo = objrS1!empleg
                                                v_doc_nro = objrS1!nrodoc
                                                Call InsertarItemLegajoFamiliares("CUIL_F", v_legajo, "", "", Empty, "", "", "", "", "", "", "", v_doc_nro, Empty, sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
                                            End If
                                        Else
                                            Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", objAuditoria!audnro, True, StrSql)
                                        End If
                                End Select
                            End If
                    End Select
                    
'--------'112, 'Alta Documento'
                Case 112:
                    Flog.writeline Espacios(Tabulador * 0) & "  Tipo " & objAuditoria!caudnro & " - Auditoria Nro --> " & objAuditoria!audnro & " - " & objAuditoria!aud_des
                    StrSql = "SELECT tipnro FROM ter_tip WHERE ternro=" & objAuditoria!aud_ternro
                    If objAuditoria!aud_ternro = "" Or IsNull(objAuditoria!aud_ternro) Then
                        Call InsertarError("0", "GENERAL", "GENERAL", "Error. El campo aud_ternro es nulo. No se puede procesar.", objAuditoria!audnro, True, StrSql)
                    Else
                        OpenRecordset StrSql, objrS1
                        Select Case objrS1!tipnro
                            Case 1, 2:
                                'EMPLEADO
                                StrSql = "SELECT empleado.empleg, ter_doc.tidnro, ter_doc.nrodoc "
                                StrSql = StrSql & " FROM empleado INNER JOIN ter_doc ON ter_doc.ternro=empleado.ternro "
                                StrSql = StrSql & " WHERE empleado.ternro = " & objAuditoria!aud_ternro & " AND ter_doc.tidnro = " & objAuditoria!AUD_ANT
                                OpenRecordset StrSql, objrS1
                                If Not objrS1.EOF Then
                                    v_legajo = objrS1!empleg
                                    v_doc_tipo = Mapeo(v_legajo, "LEGAJOS", "DOCUMENTO_TIPO", objAuditoria!audnro, "[TIPODOCU]", objrS1!tidnro, sin_error)
                                    If objrS1!tidnro <= 5 Then
                                        v_doc_nro = objrS1!nrodoc
                                        If sin_error Then
                                            Call InsertarItemLegajo("ALTA_DOC", v_legajo, "", "RNP", "", "", "", Empty, "", "", "", v_doc_tipo, v_doc_nro, "", sin_error, objAuditoria!audnro)
                                        End If
                                    ElseIf objrS1!tidnro = 10 Then
                                        'CUIL
                                        v_nro_cuil = objrS1!nrodoc
                                        Call InsertarItemLegajoLiquidar("NRO_CUIL", v_legajo, "", Empty, Empty, Empty, "", "", "", "", "", "", "", 0, 0, "", "", "", "", v_nro_cuil, "", Empty, "", "", "", "", "", "", Empty, "", sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
                                    End If
                                Else
                                    Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", objAuditoria!audnro, True, StrSql)
                                End If
                            Case 3:
                                'FAMILIAR
                                StrSql = "SELECT empleado.empleg, ter_doc.tidnro, ter_doc.nrodoc "
                                StrSql = StrSql & " FROM empleado INNER JOIN familiar ON empleado.ternro = familiar.empleado "
                                StrSql = StrSql & " INNER JOIN ter_doc ON ter_doc.ternro=familiar.ternro "
                                StrSql = StrSql & " WHERE familiar.ternro = " & objAuditoria!aud_ternro & " AND ter_doc.tidnro = " & objAuditoria!AUD_ANT
                                OpenRecordset StrSql, objrS1
                                If Not objrS1.EOF Then
                                    If objrS1!tidnro <= 5 Then
                                        v_legajo = objrS1!empleg
                                        v_doc_tipo = Mapeo(v_legajo, "LEGAJO_FAMILIARES", "DOCUMENTO_TIPO", objAuditoria!audnro, "[TIPODOCU]", objrS1!tidnro, sin_error)
                                        v_doc_nro = objrS1!nrodoc
                                        If sin_error Then
                                            'Nro documento
                                            Call InsertarItemLegajoFamiliares("ALTA_DOC_F", v_legajo, "", "", Empty, "", "", "", v_doc_tipo, v_doc_nro, "", "", "", Empty, sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
                                        End If
                                    ElseIf objrS1!tidnro = 10 Then
                                        'CUIL
                                        v_legajo = objrS1!empleg
                                        v_doc_nro = objrS1!nrodoc
                                        Call InsertarItemLegajoFamiliares("CUIL_F", v_legajo, "", "", Empty, "", "", "", "", "", "", "", v_doc_nro, Empty, sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
                                    End If
                                Else
                                    Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", objAuditoria!audnro, True, StrSql)
                                End If
                        End Select
                    End If
                    
'--------'114, 'Alta Domicilio'
                Case 114:
                    Flog.writeline Espacios(Tabulador * 0) & "  Tipo " & objAuditoria!caudnro & " - Campo 114 - Auditoria Nro --> " & objAuditoria!audnro & " - " & objAuditoria!aud_des
                    StrSql = "SELECT sucursal.estrnro, provnro FROM sucursal INNER JOIN cabdom ON sucursal.ternro=cabdom.ternro "
                    StrSql = StrSql & " INNER JOIN detdom ON cabdom.domnro = detdom.domnro "
                    StrSql = StrSql & " WHERE cabdom.domnro = " & objAuditoria!aud_actual & " AND domdefault=-1"
                    OpenRecordset StrSql, objrS1
                    If Not objrS1.EOF Then
                        v_prov_trabaja = Mapeo(v_legajo, "LEGAJO_POSICION", "PROVINCIA_TRABAJA", objAuditoria!audnro, "[PROV_TRA]", objrS1!estrnro, sin_error)
                        If sin_error Then
                            StrSql = "SELECT DISTINCT empleado.empleg, empleado.ternro "
                            StrSql = StrSql & " FROM empleado INNER JOIN his_estructura ON his_estructura.ternro=empleado.ternro "
                            'StrSql = StrSql & " WHERE empleado.ternro = " & objAuditoria!aud_ternro
                            StrSql = StrSql & " AND his_estructura.tenro=1 AND his_estructura.estrnro = " & objrS1!estrnro
                            'StrSql = StrSql & " ORDER BY his_estructura.htetdesde DESC"
                            OpenRecordset StrSql, objrS1
                            Do Until objrS1.EOF
                                v_legajo = objrS1!empleg
                                Call InsertarItemLegajoLiquidar("PROVINCIA_TRABAJA", v_legajo, "", Empty, Empty, Empty, "", "", "", "", "", "", "", 0, 0, "", "", "", "", "", "", Empty, "", v_prov_trabaja, "", "", "", "", Empty, "", sin_error, objAuditoria!audnro, objrS1!ternro)
                                objrS1.MoveNext
                            Loop
                        End If
                    End If
            
'--------'116, 'Mod. Domicilio - Provincia'
                Case 116:
                    If objAuditoria!aud_campnro = 263 Then
                        Flog.writeline Espacios(Tabulador * 0) & "  Tipo " & objAuditoria!caudnro & " - Campo 263 - Auditoria Nro --> " & objAuditoria!audnro & " - " & objAuditoria!aud_des
                        StrSql = "SELECT provnro FROM sucursal INNER JOIN cabdom ON sucursal.ternro=cabdom.ternro "
                        StrSql = StrSql & " INNER JOIN detdom ON cabdom.domnro = detdom.domnro "
                        StrSql = StrSql & " WHERE cabdom.domnro = " & objAuditoria!aud_actual & " AND domdefault=-1"
                        OpenRecordset StrSql, objrS1
                        If Not objrS1.EOF Then
                            v_prov_trabaja = Mapeo(v_legajo, "LEGAJO_POSICION", "PROVINCIA_TRABAJA", objAuditoria!audnro, "[PROV_TRA]", objrS1!estrnro, sin_error)
                            If sin_error Then
                                StrSql = "SELECT DISTINCT empleado.empleg "
                                StrSql = StrSql & " FROM empleado INNER JOIN his_estructura ON his_estructura.ternro=empleado.ternro "
                                StrSql = StrSql & " WHERE empleado.ternro = " & objAuditoria!aud_ternro
                                StrSql = StrSql & " AND his_estructura.tenro=" & CStr(C_edificio) & " AND his_estructura.estrnro = " & objrS1!estrnro
                                StrSql = StrSql & " ORDER BY his_estructura.htetdesde DESC"
                                OpenRecordset StrSql, objrS1
                                Do Until objrS1.EOF
                                    v_legajo = objrS1!empleg
                                    Call InsertarItemLegajoLiquidar("PROVINCIA_TRABAJA", v_legajo, "", Empty, Empty, Empty, "", "", "", "", "", "", "", 0, 0, "", "", "", "", "", "", Empty, "", v_prov_trabaja, "", "", "", "", Empty, "", sin_error, objAuditoria!audnro, objAuditoria!aud_ternro)
                                    objrS1.MoveNext
                                Loop
                            End If
                        End If
                    End If
            End Select
        End If
    
        'Actualizo el progreso
        Progreso = Progreso + IncPorc
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & " WHERE bpronro = " & NroProceso
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
        
        objAuditoria.MoveNext
        
    Loop
    
    
Fin:
    Exit Sub
    
CE:
    HuboErrores = True
 Resume Next
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 0) & "Error Function ComenzarTransferencia. " & Format(Now, "dd/mm/yyyy HH:mm:ss")
    Flog.writeline Espacios(Tabulador * 0) & "Error: " & Err.Number
    Flog.writeline Espacios(Tabulador * 0) & "Descripcion: " & Err.Description
    Flog.writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "SQL Ejecutado: " & StrSql
        Flog.writeline
    End If
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline
    Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", objAuditoria!audnro, False, "")
    GoTo Fin
End Sub

Private Function Mapeo(ByVal legajo As String, ByVal tabla As String, ByVal Campo As String, ByVal audnro As Long, ByVal tipo As String, ByVal Valor As String, ByRef no_error As Boolean) As String
Dim objMapeo As New ADODB.Recordset
Dim error As String

no_error = True

On Error GoTo CE:
    
    Select Case tipo
        Case "[ESTCIVIL]":
            error = "No se encontro mapeo INFOTIPO: [ESTCIVIL], TABLA: estcivil, COD. INTERNO: " & Valor
        Case "[PAIS]":
            error = "No se encontro mapeo INFOTIPO: [PAIS], TABLA: pais, COD. INTERNO: " & Valor
'        Case "[LUGAR]"
'            StrSql = "SELECT codexterno FROM mapeo_sap "
'            StrSql = StrSql & " WHERE UPPER(infotipo) = '[LUGAR]'"
'            StrSql = StrSql & " AND codinterno = '" & Valor & "'"
'            error = "No se encontro mapeo INFOTIPO: [LUGAR], TABLA: lugar_nac, COD. INTERNO: " & Valor
        Case "[NACIONAL]":
            error = "No se encontro mapeo INFOTIPO: [NACIONAL], TABLA: nacionalidad, COD. INTERNO: " & Valor
        Case "[TIPODOCU]":
            error = "No se encontro mapeo INFOTIPO: [TIPODOCU], TABLA: tipodocu, COD. INTERNO: " & Valor
        Case "[EMPRESA]":
            error = "No se encontro mapeo INFOTIPO: [EMPRESA], TABLA: empresa, COD. INTERNO: " & Valor
        Case "[PARENTES]":
            error = "No se encontro mapeo INFOTIPO: [PARENTES], TABLA: parentesco, COD. INTERNO: " & Valor
        Case "[EDIFICIO]":
            error = "No se encontro mapeo INFOTIPO: [EDIFICIO], TABLA: estructura, COD. INTERNO: " & Valor
        Case "[PUESTO]":
            error = "No se encontro mapeo INFOTIPO: [PUESTO], TABLA: estructura, COD. INTERNO: " & Valor
        Case "[GRADO]":
            error = "No se encontro mapeo INFOTIPO: [GRADO], TABLA: estructura, COD. INTERNO: " & Valor
        Case "[CCOSTO]":
            error = "No se encontro mapeo INFOTIPO: [CCOSTO], TABLA: estructura, COD. INTERNO: " & Valor
        Case "[ESTRUC_A]":
            error = "No se encontro mapeo INFOTIPO: [ESTRUC_A], TABLA: estructura, COD. INTERNO: " & Valor
        Case "[C_TRABAJ]":
            error = "No se encontro mapeo INFOTIPO: [C_TRABAJ], TABLA: estructura, COD. INTERNO: " & Valor
        Case "[EDIFICIO]":
            error = "No se encontro mapeo INFOTIPO: [EDIFICIO], TABLA: estructura, COD. INTERNO: " & Valor
        Case "[CANAL_CO]":
            error = "No se encontro mapeo INFOTIPO: [CANAL_CO], TABLA: estructura, COD. INTERNO: " & Valor
        Case "[MOT_BAJA]":
            error = "No se encontro mapeo INFOTIPO: [MOT_BAJA], TABLA: causa, COD. INTERNO: " & Valor
        Case "[REL_LABO]":
            error = "No se encontro mapeo INFOTIPO: [REL_LABO], TABLA: estructura, COD. INTERNO: " & Valor
        Case "[DGI_EXEN]":
            error = "No se encontro mapeo INFOTIPO: [DGI_EXEN], TABLA: estructura, COD. INTERNO: " & Valor
        Case "[CONVENIO]":
            error = "No se encontro mapeo INFOTIPO: [CONVENIO], TABLA: estructura, COD. INTERNO: " & Valor
        Case "[CAT_CONV]":
            error = "No se encontro mapeo INFOTIPO: [CAT_CONV], TABLA: estructura, COD. INTERNO: " & Valor
        Case "[FUN_CONV]":
            error = "No se encontro mapeo INFOTIPO: [FUN_CONV], TABLA: estructura, COD. INTERNO: " & Valor
        Case "[HOR_MESH]":
            error = "No se encontro mapeo INFOTIPO: [HOR_MESH], TABLA: estructura, COD. INTERNO: " & Valor
        Case "[HOR_MESD]":
            error = "No se encontro mapeo INFOTIPO: [HOR_MESD], TABLA: estructura, COD. INTERNO: " & Valor
        Case "[OSOCIAL]":
            error = "No se encontro mapeo INFOTIPO: [OSOCIAL], TABLA: estructura, COD. INTERNO: " & Valor
        Case "[CAR_SERV]":
            error = "No se encontro mapeo INFOTIPO: [CAR_SERV], TABLA: estructura, COD. INTERNO: " & Valor
        Case "[GRUP_LIQ]":
            error = "No se encontro mapeo INFOTIPO: [GRUP_LIQ], TABLA: estructura, COD. INTERNO: " & Valor
        Case "[EST_NIVE]":
            error = "No se encontro mapeo INFOTIPO: [EST_NIVE], TABLA: nivest, COD. INTERNO: " & Valor
        Case "[INS_PAGO]":
            error = "No se encontro mapeo INFOTIPO: [INS_PAGO], TABLA: estructura, COD. INTERNO: " & Valor
        Case "[TIPO_CTA]":
            error = "No se encontro mapeo INFOTIPO: [TIPO_CTA], TABLA: formapago, COD. INTERNO: " & Valor
        Case "[ENT_PAGO]":
            error = "No se encontro mapeo INFOTIPO: [ENT_PAGO], TABLA: banco, COD. INTERNO: " & Valor
        Case "[PROV_TRA]":
            error = "No se encontro mapeo INFOTIPO: [PROV_TRA], TABLA: provincia, COD. INTERNO: " & Valor
    End Select
    
    StrSql = "SELECT codexterno FROM mapeo_sap "
    StrSql = StrSql & " WHERE UPPER(infotipo) = '" & tipo & "'"
    StrSql = StrSql & " AND codinterno = '" & Valor & "'"
    OpenRecordset StrSql, objMapeo
    
    If Not objMapeo.EOF Then
        Mapeo = objMapeo!codexterno
    Else
        Call InsertarError(legajo, tabla, Campo, error, CLng(audnro), False, "")
        no_error = False
    End If
  
Fin:
    If objMapeo.State = adStateOpen Then objMapeo.Close
    Set objMapeo = Nothing
    Exit Function
    
CE:
    HuboErrores = True
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 0) & "Error Function Mapeo. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline Espacios(Tabulador * 0) & "Error: " & Err.Number
    Flog.writeline Espacios(Tabulador * 0) & "Decripcion: " & Err.Description
    Flog.writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "SQL Ejecutado: " & StrSql
        Flog.writeline
    End If
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline
    Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", audnro, False, "")
    GoTo Fin
End Function
Private Sub InsertarItemLegajoItemLiquidar(ByVal tipo As String, ByVal legajo As String, ByVal apellido_nombre As String, ByVal documento_emisor As String, ByVal sexo As String, ByVal estcivil As String, ByVal apellido_casada As String, ByVal fech_nac As Date, ByVal nacion_nac As String, ByVal lugar_nac As String, ByVal nacionalidad As String, ByVal doc_tipo As String, ByVal doc_nro As String, ByVal empresa As String, ByVal Fecha_Ingreso As Date, ByVal estudios_nivel As String, ByVal no_error As Boolean, ByVal audnro As Long, ByVal ternro As Long)
Dim objInsert As New ADODB.Recordset
Dim s_error As Boolean
Dim strsql2
Dim strsql3
Dim aux_empresa As String
Dim aux_fecha_ingreso
Dim v
Dim I
Dim fecha_nac_aux
Dim estado_aux

On Error GoTo CE:

MyBeginTrans

    fecha_nac_aux = IIf(CStr(fech_nac) = "12:00:00 a.m.", "Null", ConvFecha(fech_nac))

    StrSql = "SELECT * FROM legajos_rhpro WHERE LEGAJO = '" & legajo & "'"
    
    OpenRecordset StrSql, objInsert
    
    If objInsert.EOF Then
        strsql2 = "INSERT INTO legajos_rhpro (LEGAJO,APELLIDO_NOMBRE,DOCUMENTO_EMISOR,SEXO,ESTADO_CIVIL,APELLIDO_CASADA,FECHA_NACIMIENTO,NACION_NACIMIENTO,LUGAR_NACIMIENTO,NACIONALIDAD,DOCUMENTO_TIPO,DOCUMENTO_NUMERO,EMPRESA)"
        strsql2 = strsql2 & " VALUES ('" & legajo & "','"
        strsql2 = strsql2 & apellido_nombre & "','"
        strsql2 = strsql2 & "RNP','"
        strsql2 = strsql2 & sexo & "','"
        strsql2 = strsql2 & estcivil & "','"
        strsql2 = strsql2 & apellido_casada & "',"
        strsql2 = strsql2 & fecha_nac_aux & ",'"
        strsql2 = strsql2 & nacion_nac & "','"
        strsql2 = strsql2 & lugar_nac & "','"
        strsql2 = strsql2 & nacionalidad & "','"
        strsql2 = strsql2 & doc_tipo & "','"
        strsql2 = strsql2 & doc_nro & "','"
        strsql2 = strsql2 & empresa & "')"
    Else
        strsql2 = "UPDATE legajos_rhpro SET "
        v = Split(tipo, ",")
        For I = 0 To UBound(v)
            Select Case v(I)
                Case "ALTA_EMP":
                    strsql2 = strsql2 & "APELLIDO_NOMBRE='" & apellido_nombre & "',"
                    strsql2 = strsql2 & "SEXO = '" & sexo & "',"
                    strsql2 = strsql2 & "ESTADO_CIVIL = '" & estcivil & "',"
                    strsql2 = strsql2 & "APELLIDO_CASADA = '" & apellido_casada & "',"
                    strsql2 = strsql2 & "FECHA_NACIMIENTO = " & fecha_nac_aux & ","
                    strsql2 = strsql2 & "NACION_NACIMIENTO = '" & nacion_nac & "',"
                    strsql2 = strsql2 & "LUGAR_NACIMIENTO = '" & lugar_nac & "',"
                    strsql2 = strsql2 & "NACIONALIDAD = '" & nacionalidad & "',"
                Case "ALTA_DOC":
                    strsql2 = strsql2 & "DOCUMENTO_TIPO = '" & doc_tipo & "',"
                    strsql2 = strsql2 & "DOCUMENTO_NUMERO = '" & doc_nro & "',"
                Case "AP_NOM":
                    strsql2 = strsql2 & "APELLIDO_NOMBRE='" & apellido_nombre & "',"
                Case "SEXO":
                    strsql2 = strsql2 & "SEXO = '" & sexo & "',"
                Case "EST_CIV":
                    strsql2 = strsql2 & "ESTADO_CIVIL = '" & estcivil & "',"
                Case "AP_CAS":
                    strsql2 = strsql2 & "APELLIDO_CASADA = '" & apellido_casada & "',"
                Case "F_NAC":
                    strsql2 = strsql2 & "FECHA_NACIMIENTO = " & fecha_nac_aux & ","
                Case "PAIS":
                    strsql2 = strsql2 & "NACION_NACIMIENTO = '" & nacion_nac & "',"
                Case "LUG_NAC":
                    strsql2 = strsql2 & "LUGAR_NACIMIENTO = '" & lugar_nac & "',"
                Case "NACION":
                    strsql2 = strsql2 & "NACIONALIDAD = '" & nacionalidad & "',"
                Case "DOC_TPO":
                    strsql2 = strsql2 & "DOCUMENTO_TIPO = '" & doc_tipo & "',"
                Case "DOC_NRO":
                    strsql2 = strsql2 & "DOCUMENTO_NUMERO = '" & doc_nro & "',"
                Case "EMPRESA":
                    strsql2 = strsql2 & "EMPRESA = '" & empresa & "',"
            End Select
        Next
        strsql2 = Left(strsql2, Len(strsql2) - 1)
        strsql2 = strsql2 & " WHERE LEGAJO = '" & legajo & "'"
    End If
    
    If no_error Then
        aux_fecha_ingreso = IIf(Not EsNulo(Fecha_Ingreso), CDate(Fecha_Ingreso), Empty)
        If InStr(1, tipo, "EMPRESA") = 0 Then
            'Buscar tipo estructura EMPRESA
            StrSql = "SELECT his_estructura.estrnro "
            StrSql = StrSql & " FROM his_estructura "
            StrSql = StrSql & " WHERE his_estructura.ternro = " & ternro & " AND his_estructura.tenro= " & C_empresa
            StrSql = StrSql & " ORDER BY htetdesde DESC"
            OpenRecordset StrSql, objInsert
            If Not objInsert.EOF Then
'                Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", audnro, True, StrSql)
'            Else
                aux_empresa = Mapeo(legajo, "LEGAJO_LIQUIDAR", "EMPRESA", audnro, "[EMPRESA]", objInsert!estrnro, s_error)
                If Not s_error Then
                    GoTo Fin
                End If
            End If
        End If
        
        ' Es un operacion de ALTA, con lo cual generalmente se informa la FECHA_INGRESO.
        If InStr(1, tipo, "FECHA_INGRESO") = 0 Then
            'Buscar FECHA_INGRESO
            StrSql = "SELECT empleado.empfaltagr, empleado.empleg "
            StrSql = StrSql & " FROM empleado "
            StrSql = StrSql & " WHERE empleado.ternro = " & ternro
            OpenRecordset StrSql, objInsert
'            If objInsert.EOF Then
'                Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", audnro, True, StrSql)
'            Else
            If Not objInsert.EOF Then
                aux_fecha_ingreso = objInsert!empfaltagr
            End If
        End If
        
        If no_error And Not (EsNulo(legajo) Or EsNulo(aux_empresa) Or EsNulo(aux_fecha_ingreso)) Then
            StrSql = "SELECT * FROM legajo_liquidar_rhpro WHERE LEGAJO = '" & legajo & "' AND EMPRESA = '" & aux_empresa & "' AND FECHA_INGRESO = " & ConvFecha(aux_fecha_ingreso)
            OpenRecordset StrSql, objInsert
            
            'Buscar ESTADO (A si es activo,  T si es baja)
            StrSql = "SELECT empleado.empest "
            StrSql = StrSql & " FROM empleado "
            StrSql = StrSql & " WHERE empleado.ternro = " & ternro
            OpenRecordset StrSql, objInsert
            estado_aux = "T"
            If objInsert.EOF Then
                If CInt(objInsert!empest) = -1 Then
                    estado_aux = "A"
                End If
            End If
            
            If objInsert.EOF Then
                strsql3 = "INSERT INTO legajo_liquidar_rhpro(legajo,EMPRESA,ESTADO,FECHA_INGRESO,ESTUDIOS_NIVEL)"
                strsql3 = strsql3 & " VALUES ('" & legajo & "','"
                strsql3 = strsql3 & aux_empresa & "','"
                strsql3 = strsql3 & estado_aux & "',"
                strsql3 = strsql3 & ConvFecha(aux_fecha_ingreso) & ","
                strsql3 = strsql3 & estudios_nivel & "','"
                strsql3 = strsql3 & "S',')"
            End If
        End If
    End If
        
    
    If no_error Then
        If strsql2 <> "" Then
            objConn.Execute strsql2, , adExecuteNoRecords
        End If
        If strsql3 <> "" Then
            objConn.Execute strsql3, , adExecuteNoRecords
        End If
        
        StrSql = "UPDATE auditoria SET procesado = -1 "
        StrSql = StrSql & " WHERE audnro = " & audnro
        objConn.Execute StrSql, , adExecuteNoRecords
        
        StrSql = "DELETE FROM error_LBR "
        StrSql = StrSql & " WHERE LEGAJO = '" & legajo & "' AND AUDNRO = " & audnro
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
MyCommitTrans

Fin:
    If objInsert.State = adStateOpen Then objInsert.Close
    Set objInsert = Nothing
    Exit Sub
    
CE:
'Resume Next
    HuboErrores = True
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 0) & "Error Function InsertarItemLegajoItemLiquidar. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline Espacios(Tabulador * 0) & "Error: " & Err.Number
    Flog.writeline Espacios(Tabulador * 0) & "Decripcion: " & Err.Description
    Flog.writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "SQL Ejecutado: " & StrSql
        Flog.writeline
    End If
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline
    Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", audnro, False, "")
    GoTo Fin
End Sub
Private Sub InsertarItemLegajo(ByVal tipo As String, ByVal legajo As String, ByVal apellido_nombre As String, ByVal documento_emisor As String, ByVal sexo As String, ByVal estcivil As String, ByVal apellido_casada As String, ByVal fech_nac As Date, ByVal nacion_nac As String, ByVal lugar_nac As String, ByVal nacionalidad As String, ByVal doc_tipo As String, ByVal doc_nro As String, ByVal empresa As String, ByVal no_error As Boolean, ByVal audnro As Long)
Dim objInsert As New ADODB.Recordset
Dim fecha_nac_aux
On Error GoTo CE:

MyBeginTrans

    StrSql = "SELECT * FROM legajos_rhpro WHERE LEGAJO = '" & legajo & "'"
    
    OpenRecordset StrSql, objInsert
    
    fecha_nac_aux = IIf(CStr(fech_nac) = "12:00:00 a.m.", "Null", ConvFecha(fech_nac))
    
    If objInsert.EOF Then
        StrSql = "INSERT INTO legajos_rhpro (LEGAJO,APELLIDO_NOMBRE,DOCUMENTO_EMISOR,SEXO,ESTADO_CIVIL,APELLIDO_CASADA,FECHA_NACIMIENTO,NACION_NACIMIENTO,LUGAR_NACIMIENTO,NACIONALIDAD,DOCUMENTO_TIPO,DOCUMENTO_NUMERO,EMPRESA)"
        StrSql = StrSql & " VALUES ('" & legajo & "','"
        StrSql = StrSql & apellido_nombre & "','"
        StrSql = StrSql & "RNP','"
        StrSql = StrSql & sexo & "','"
        StrSql = StrSql & estcivil & "','"
        StrSql = StrSql & apellido_casada & "',"
        StrSql = StrSql & fecha_nac_aux & ",'"
        StrSql = StrSql & nacion_nac & "','"
        StrSql = StrSql & lugar_nac & "','"
        StrSql = StrSql & nacionalidad & "','"
        StrSql = StrSql & doc_tipo & "','"
        StrSql = StrSql & doc_nro & "','"
        StrSql = StrSql & empresa & "')"
    Else
        StrSql = "UPDATE legajos_rhpro SET "
        Select Case tipo
            Case "ALTA_EMP":
                StrSql = StrSql & "APELLIDO_NOMBRE='" & apellido_nombre & "',"
                StrSql = StrSql & "SEXO = '" & sexo & "',"
                StrSql = StrSql & "ESTADO_CIVIL = '" & estcivil & "',"
                StrSql = StrSql & "APELLIDO_CASADA = '" & apellido_casada & "',"
                StrSql = StrSql & "FECHA_NACIMIENTO = " & ConvFecha(fecha_nac_aux) & ","
                StrSql = StrSql & "NACION_NACIMIENTO = '" & nacion_nac & "',"
                StrSql = StrSql & "LUGAR_NACIMIENTO = '" & lugar_nac & "',"
                StrSql = StrSql & "NACIONALIDAD = '" & nacionalidad & "'"
            Case "ALTA_DOC":
                StrSql = StrSql & "DOCUMENTO_TIPO = '" & doc_tipo & "',"
                StrSql = StrSql & "DOCUMENTO_NUMERO = '" & doc_nro & "'"
            Case "AP_NOM":
                StrSql = StrSql & "APELLIDO_NOMBRE='" & apellido_nombre & "'"
            Case "SEXO":
                StrSql = StrSql & "SEXO = '" & sexo & "'"
            Case "EST_CIV":
                StrSql = StrSql & "ESTADO_CIVIL = '" & estcivil & "'"
            Case "AP_CAS":
                StrSql = StrSql & "APELLIDO_CASADA = '" & apellido_casada & "'"
            Case "F_NAC":
                StrSql = StrSql & "FECHA_NACIMIENTO = " & fecha_nac_aux
            Case "PAIS":
                StrSql = StrSql & "NACION_NACIMIENTO = '" & nacion_nac & "'"
            Case "LUG_NAC":
                StrSql = StrSql & "LUGAR_NACIMIENTO = '" & lugar_nac & "'"
            Case "NACION":
                StrSql = StrSql & "NACIONALIDAD = '" & nacionalidad & "'"
            Case "DOC_TPO":
                StrSql = StrSql & "DOCUMENTO_TIPO = '" & doc_tipo & "'"
            Case "DOC_NRO":
                StrSql = StrSql & "DOCUMENTO_NUMERO = '" & doc_nro & "'"
            Case "EMPRESA":
                StrSql = StrSql & "EMPRESA = '" & empresa & "'"
        End Select
        StrSql = StrSql & " WHERE LEGAJO = '" & legajo & "'"
    End If
    
    objConn.Execute StrSql, , adExecuteNoRecords
    
    If no_error Then
        StrSql = "UPDATE auditoria SET procesado = -1 "
        StrSql = StrSql & " WHERE audnro = " & audnro
        objConn.Execute StrSql, , adExecuteNoRecords
        
        StrSql = "DELETE FROM error_LBR "
        StrSql = StrSql & " WHERE LEGAJO = '" & legajo & "' AND AUDNRO = " & audnro
        objConn.Execute StrSql, , adExecuteNoRecords
    End If

MyCommitTrans

Fin:
    If objInsert.State = adStateOpen Then objInsert.Close
    Set objInsert = Nothing
    Exit Sub
    
CE:
'Resume Next
    HuboErrores = True
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 0) & "Error Function InsertarItemLegajo. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline Espacios(Tabulador * 0) & "Error: " & Err.Number
    Flog.writeline Espacios(Tabulador * 0) & "Decripcion: " & Err.Description
    Flog.writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "SQL Ejecutado: " & StrSql
        Flog.writeline
    End If
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline
    Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", audnro, False, "")
    GoTo Fin
End Sub

Private Sub InsertarItemEmpresa(ByVal tipo As String, ByVal legajo As String, ByVal Fecha As Date, ByVal empresa As String, ByVal empresa_cambio As String, ByVal no_error As Boolean, ByVal audnro As Long, ByVal ternro As Long)
Dim objInsert As New ADODB.Recordset
Dim v_edificio As String
Dim aux_fecha_ingreso
Dim s_error As Boolean
Dim strsql2
Dim strsql3
Dim strsql4
On Error GoTo CE:

MyBeginTrans
    
'--------------------------------
    ' LEGAJOS_RHPRO
    StrSql = "SELECT * FROM legajos_rhpro WHERE LEGAJO = '" & legajo & "'"
    
    OpenRecordset StrSql, objInsert
    
    If objInsert.EOF Then
        strsql2 = "INSERT INTO legajos_rhpro (LEGAJO,EMPRESA)"
        strsql2 = strsql2 & " VALUES ('" & legajo & "','"
        strsql2 = strsql2 & empresa & "')"
    Else
        strsql2 = "UPDATE legajos_rhpro SET "
        strsql2 = strsql2 & "EMPRESA = '" & empresa & "'"
        strsql2 = strsql2 & " WHERE LEGAJO = '" & legajo & "'"
    End If
      
'--------------------------------
    ' LEGAJO_POSICION
    'Buscar tipo estructura EDIFICIO
    StrSql = "SELECT empleado.empleg, his_estructura.estrnro "
    StrSql = StrSql & " FROM empleado INNER JOIN his_estructura ON his_estructura.ternro=empleado.ternro "
    StrSql = StrSql & " WHERE empleado.ternro = " & ternro & " AND his_estructura.tenro= " & C_edificio
    StrSql = StrSql & " ORDER BY htetdesde DESC"
    OpenRecordset StrSql, objInsert
    If Not objInsert.EOF Then
'        Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", audnro, True, StrSql)
'    Else
        v_edificio = Mapeo(legajo, "LEGAJO_POSICION", "EDIFICIO", audnro, "[EDIFICIO]", objInsert!estrnro, s_error)
        If Not s_error Then
            GoTo Fin
        End If
    End If
    
    StrSql = "SELECT * FROM legajo_posicion_rhpro WHERE LEGAJO = '" & legajo & "' AND FECHA = " & ConvFecha(Fecha)
    OpenRecordset StrSql, objInsert
    
    If objInsert.EOF Then
        strsql3 = "INSERT INTO legajo_posicion_rhpro(legajo,fecha,empresa,empresa_cambio,activo,grupo_tablas,edificio)"
        strsql3 = strsql3 & " VALUES ('" & legajo & "',"
        strsql3 = strsql3 & ConvFecha(Fecha) & ",'"
        strsql3 = strsql3 & empresa & "','"
        strsql3 = strsql3 & empresa_cambio & "','"
        strsql3 = strsql3 & "S','"
        strsql3 = strsql3 & "01','"
        strsql3 = strsql3 & v_edificio & "')"
    Else
        strsql3 = "UPDATE legajo_posicion_rhpro SET "
        strsql3 = strsql3 & "empresa = '" & empresa & "',"
        strsql3 = strsql3 & "empresa_cambio = '" & empresa_cambio & "'"
        strsql3 = strsql3 & " WHERE legajo = '" & legajo & "'"
        strsql3 = strsql3 & " AND fecha = " & ConvFecha(Fecha)
    End If
    
    
'--------------------------------
    ' LEGAJO_LIQUIDAR
    'Buscar FECHA_INGRESO
    If InStr(1, tipo, "LIQUIDAR") <> 0 Then
        StrSql = "SELECT empleado.empfaltagr, empleado.empleg "
        StrSql = StrSql & " FROM empleado "
        StrSql = StrSql & " WHERE empleado.ternro = " & ternro
        OpenRecordset StrSql, objInsert
        If Not objInsert.EOF Then
'            Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", audnro, True, StrSql)
'            no_error = False
'        Else
            aux_fecha_ingreso = objInsert!empfaltagr
        End If
    
        If Not (EsNulo(legajo) Or EsNulo(empresa) Or EsNulo(aux_fecha_ingreso)) Then
            StrSql = "SELECT * FROM legajo_liquidar_rhpro WHERE LEGAJO = '" & legajo & "' AND EMPRESA = '" & empresa & "' AND FECHA_INGRESO = " & ConvFecha(aux_fecha_ingreso)
            OpenRecordset StrSql, objInsert
    
            If objInsert.EOF Then
                strsql4 = "INSERT INTO legajo_liquidar_rhpro(legajo,EMPRESA,FECHA_INGRESO,ACTIVO)"
                strsql4 = strsql4 & " VALUES ('" & legajo & "','"
                strsql4 = strsql4 & empresa & "',"
                strsql4 = strsql4 & ConvFecha(aux_fecha_ingreso) & ",'"
                strsql4 = strsql4 & "S')"
            End If
        End If
    End If
     
    If no_error Then
        If strsql2 <> "" Then
            objConn.Execute strsql2, , adExecuteNoRecords
        End If
        If strsql3 <> "" Then
            objConn.Execute strsql3, , adExecuteNoRecords
        End If
        If strsql4 <> "" Then
            objConn.Execute strsql4, , adExecuteNoRecords
        End If
        
        StrSql = "UPDATE auditoria SET procesado = -1 "
        StrSql = StrSql & " WHERE audnro = " & audnro
        objConn.Execute StrSql, , adExecuteNoRecords
        
        StrSql = "DELETE FROM error_LBR "
        StrSql = StrSql & " WHERE LEGAJO = '" & legajo & "' AND AUDNRO = " & audnro
        objConn.Execute StrSql, , adExecuteNoRecords
    End If

MyCommitTrans

Fin:
    If objInsert.State = adStateOpen Then objInsert.Close
    Set objInsert = Nothing
    Exit Sub
    
CE:
'Resume Next
    HuboErrores = True
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 0) & "Error Function InsertarItemEmpresa. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline Espacios(Tabulador * 0) & "Error: " & Err.Number
    Flog.writeline Espacios(Tabulador * 0) & "Decripcion: " & Err.Description
    Flog.writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "SQL Ejecutado: " & StrSql
        Flog.writeline
    End If
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline
    Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", audnro, False, "")
    GoTo Fin
End Sub
Private Sub InsertarItemLegajoFamiliares(ByVal tipo As String, ByVal legajo As String, ByVal parentesco As String, ByVal apellido_nombre As String, ByVal fech_nac As Date, ByVal nacionalidad As String, ByVal nacion_nac As String, ByVal lugar_nac As String, ByVal doc_tipo As String, ByVal doc_nro As String, ByVal sexo As String, ByVal estcivil As String, ByVal Cuil As String, ByVal famfec As Date, ByVal no_error As Boolean, ByVal audnro As Long, ByVal ternro As Long)
Dim objInsert As New ADODB.Recordset
Dim fliar_nro As Integer
Dim fecha_nac_aux
Dim famfec_aux
On Error GoTo CE:

MyBeginTrans

    fecha_nac_aux = IIf(CStr(fech_nac) = "12:00:00 a.m.", "Null", ConvFecha(fech_nac))
    famfec_aux = IIf(CStr(famfec) = "12:00:00 a.m.", "Null", ConvFecha(famfec))
    
    StrSql = "SELECT COUNT(*) cant FROM familiar "
    StrSql = StrSql & " INNER JOIN empleado ON familiar.empleado = empleado.ternro "
    StrSql = StrSql & " WHERE empleg = '" & legajo & "' AND familiar.ternro < " & ternro
    OpenRecordset StrSql, objInsert
    fliar_nro = 1
    If Not objInsert.EOF Then
        fliar_nro = CInt(objInsert!cant) + 1
    End If
    
    StrSql = "SELECT * FROM legajo_familiares_rhpro WHERE LEGAJO = '" & legajo & "' AND FAMILIAR_NUMERO = '" & CStr(fliar_nro) & "'"
    OpenRecordset StrSql, objInsert
    
    If objInsert.EOF Then
        StrSql = "INSERT INTO legajo_familiares_rhpro (LEGAJO,FAMILIAR_NUMERO,PARENTESCO,APELLIDO_NOMBRE,FECHA_NACIMIENTO,NACIONALIDAD,NACION_NACIMIENTO,LUGAR_NACIMIENTO,DOCUMENTO_TIPO,DOCUMENTO_NUMERO,SEXO,ESTADO_CIVIL,CUIL_FAM,FECHA_INICIO_VINCULO) "
        StrSql = StrSql & " VALUES ('" & legajo & "','"
        StrSql = StrSql & CStr(fliar_nro) & "','"
        StrSql = StrSql & parentesco & "','"
        StrSql = StrSql & apellido_nombre & "',"
        StrSql = StrSql & fecha_nac_aux & ",'"
        StrSql = StrSql & nacionalidad & "','"
        StrSql = StrSql & nacion_nac & "','"
        StrSql = StrSql & lugar_nac & "','"
        StrSql = StrSql & doc_tipo & "','"
        StrSql = StrSql & doc_nro & "','"
        StrSql = StrSql & sexo & "','"
        StrSql = StrSql & estcivil & "','"
        StrSql = StrSql & Cuil & "',"
        StrSql = StrSql & famfec_aux & ")"
    Else
        StrSql = "UPDATE legajo_familiares_rhpro SET "
        Select Case tipo
            Case "ALTA_FAM":
                StrSql = StrSql & "PARENTESCO='" & parentesco & "',"
                StrSql = StrSql & "APELLIDO_NOMBRE='" & apellido_nombre & "',"
                StrSql = StrSql & "FECHA_NACIMIENTO = " & fecha_nac_aux & ","
                StrSql = StrSql & "NACIONALIDAD = '" & nacionalidad & "',"
                StrSql = StrSql & "NACION_NACIMIENTO = '" & nacion_nac & "',"
                StrSql = StrSql & "LUGAR_NACIMIENTO = '" & lugar_nac & "',"
                StrSql = StrSql & "SEXO = '" & sexo & "',"
                StrSql = StrSql & "ESTADO_CIVIL = '" & estcivil & "',"
                StrSql = StrSql & "FECHA_INICIO_VINCULO = " & famfec_aux
            Case "ALTA_DOC_F":
                StrSql = StrSql & "DOCUMENTO_TIPO = '" & doc_tipo & "',"
                StrSql = StrSql & "DOCUMENTO_NUMERO = '" & doc_nro & "'"
            Case "PARENT_F":
                StrSql = StrSql & "PARENTESCO='" & parentesco & "'"
            Case "AP_NOM_F":
                StrSql = StrSql & "APELLIDO_NOMBRE='" & apellido_nombre & "'"
            Case "F_NAC_F":
                StrSql = StrSql & "FECHA_NACIMIENTO = " & fecha_nac_aux
            Case "NACION_F":
                StrSql = StrSql & "NACIONALIDAD = '" & nacionalidad & "'"
            Case "PAIS_F":
                StrSql = StrSql & "NACION_NACIMIENTO = '" & nacion_nac & "'"
            Case "LUG_NAC_F":
                StrSql = StrSql & "LUGAR_NACIMIENTO = '" & lugar_nac & "'"
            Case "DOC_TPO_F":
                StrSql = StrSql & "DOCUMENTO_TIPO = '" & doc_tipo & "'"
            Case "DOC_NRO_F":
                StrSql = StrSql & "DOCUMENTO_NUMERO = '" & doc_nro & "'"
            Case "SEXO_F":
                StrSql = StrSql & "SEXO = '" & sexo & "'"
            Case "EST_CIV_F":
                StrSql = StrSql & "ESTADO_CIVIL = '" & estcivil & "'"
            Case "CUIL_F":
                StrSql = StrSql & "CUIL_FAM = '" & Cuil & "'"
            Case "F_INI_VINC_F":
                StrSql = StrSql & "FECHA_INICIO_VINCULO = " & famfec_aux
        End Select
        StrSql = StrSql & " WHERE LEGAJO = '" & legajo & "'"
        StrSql = StrSql & " AND FAMILIAR_NUMERO = '" & CStr(fliar_nro) & "'"
    End If
    
    objConn.Execute StrSql, , adExecuteNoRecords
    
    If no_error Then
        StrSql = "UPDATE auditoria SET procesado = -1 "
        StrSql = StrSql & " WHERE audnro = " & audnro
        objConn.Execute StrSql, , adExecuteNoRecords
        
        StrSql = "DELETE FROM error_LBR "
        StrSql = StrSql & " WHERE LEGAJO = '" & legajo & "' AND AUDNRO = " & audnro
        objConn.Execute StrSql, , adExecuteNoRecords
    End If

MyCommitTrans

Fin:
    If objInsert.State = adStateOpen Then objInsert.Close
    Set objInsert = Nothing
    Exit Sub
    
CE:
'Resume Next
    HuboErrores = True
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 0) & "Error Function InsertarItemLegajoFamiliares. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline Espacios(Tabulador * 0) & "Error: " & Err.Number
    Flog.writeline Espacios(Tabulador * 0) & "Decripcion: " & Err.Description
    Flog.writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "SQL Ejecutado: " & StrSql
        Flog.writeline
    End If
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline
    Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", audnro, False, "")
    GoTo Fin
End Sub

Private Sub InsertarItemLegajoPosicion(ByVal tipo As String, ByVal legajo As String, ByVal Fecha As Date, ByVal puesto As String, ByVal puesto_cambio As String, ByVal grado As String, ByVal grado_cambio As String, ByVal empresa As String, ByVal empresa_cambio As String, ByVal centro_costo As String, ByVal c_costo_cambio As String, ByVal estruc_area As String, ByVal activo As String, ByVal grupo_tablas As Date, ByVal centro_trabajo As String, ByVal c_trabajo_cambio As String, ByVal edificio As String, ByVal area_cambio As String, ByVal canal_costo As String, ByVal cnl_costo_cambio As String, ByVal no_error As Boolean, ByVal audnro As Long, ByVal ternro As Long)
Dim objInsert As New ADODB.Recordset
Dim v_edificio As String
Dim s_error As Boolean

On Error GoTo CE:

MyBeginTrans
    
    v_edificio = edificio
    If tipo <> "EDIFICIO" Then
        'Buscar tipo estructura EDIFICIO
        StrSql = "SELECT empleado.empleg, his_estructura.estrnro "
        StrSql = StrSql & " FROM empleado INNER JOIN his_estructura ON his_estructura.ternro=empleado.ternro "
        StrSql = StrSql & " WHERE empleado.ternro = " & ternro & " AND his_estructura.tenro= " & C_edificio
        StrSql = StrSql & " ORDER BY htetdesde DESC"
        OpenRecordset StrSql, objInsert
        If Not objInsert.EOF Then
'            Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", audnro, True, StrSql)
'        Else
            v_edificio = Mapeo(legajo, "LEGAJO_POSICION", "EDIFICIO", audnro, "[EDIFICIO]", objInsert!estrnro, s_error)
            If Not s_error Then
                GoTo Fin
            End If
        End If
    End If
    
    
    StrSql = "SELECT * FROM legajo_posicion_rhpro WHERE LEGAJO = '" & legajo & "' AND FECHA = " & ConvFecha(Fecha)
    OpenRecordset StrSql, objInsert
    
    If objInsert.EOF Then
        StrSql = "INSERT INTO legajo_posicion_rhpro(legajo,fecha,puesto,puesto_cambio,grado,grado_cambio,empresa,"
        StrSql = StrSql & "empresa_cambio,centro_costo,c_costo_cambio,estruc_area,activo,grupo_tablas,centro_trabajo,"
        StrSql = StrSql & "c_trabajo_cambio,edificio,area_cambio,canal_costo,cnl_costo_cambio)"
        StrSql = StrSql & " VALUES ('" & legajo & "',"
        StrSql = StrSql & ConvFecha(Fecha) & ",'"
        StrSql = StrSql & puesto & "','"
        StrSql = StrSql & puesto_cambio & "','"
        StrSql = StrSql & grado & "','"
        StrSql = StrSql & grado_cambio & "','"
        StrSql = StrSql & empresa & "','"
        StrSql = StrSql & empresa_cambio & "','"
        StrSql = StrSql & centro_costo & "','"
        StrSql = StrSql & c_costo_cambio & "','"
        StrSql = StrSql & estruc_area & "','"
        StrSql = StrSql & "S','"
        StrSql = StrSql & "01','"
        StrSql = StrSql & centro_trabajo & "','"
        StrSql = StrSql & c_trabajo_cambio & "','"
        StrSql = StrSql & v_edificio & "','"
        StrSql = StrSql & area_cambio & "','"
        StrSql = StrSql & canal_costo & "','"
        StrSql = StrSql & cnl_costo_cambio & "')"
    Else
        StrSql = "UPDATE legajo_posicion_rhpro SET "
        Select Case tipo
            Case "PUESTO":
                StrSql = StrSql & "puesto='" & puesto & "',"
                StrSql = StrSql & "puesto_cambio='" & puesto_cambio & "'"
            Case "GRADO":
                StrSql = StrSql & "grado = '" & grado & "',"
                StrSql = StrSql & "grado_cambio = '" & grado_cambio & "'"
            Case "EMPRESA":
                StrSql = StrSql & "empresa = '" & empresa & "',"
                StrSql = StrSql & "empresa_cambio = '" & empresa_cambio & "'"
            Case "CCOSTO":
                StrSql = StrSql & "centro_costo = '" & centro_costo & "',"
                StrSql = StrSql & "c_costo_cambio= '" & c_costo_cambio & "'"
            Case "EST_AREA":
                StrSql = StrSql & "estruc_area = '" & estruc_area & "',"
                StrSql = StrSql & "area_cambio = '" & area_cambio & "'"
            Case "CEN_TRAB":
                StrSql = StrSql & "centro_trabajo = '" & centro_trabajo & "',"
                StrSql = StrSql & "c_trabajo_cambio= '" & c_trabajo_cambio & "'"
            Case "EDIFICIO":
                StrSql = StrSql & "edificio = '" & edificio & "'"
            Case "CANAL_C":
                StrSql = StrSql & "canal_costo = '" & canal_costo & "',"
                StrSql = StrSql & "cnl_costo_cambio = '" & cnl_costo_cambio & "'"
        End Select
        StrSql = StrSql & " WHERE legajo = '" & legajo & "'"
        StrSql = StrSql & " AND fecha = " & ConvFecha(Fecha)
    End If
    
    objConn.Execute StrSql, , adExecuteNoRecords
    
    If no_error Then
        StrSql = "UPDATE auditoria SET procesado = -1 "
        StrSql = StrSql & " WHERE audnro = " & audnro
        objConn.Execute StrSql, , adExecuteNoRecords
        
        StrSql = "DELETE FROM error_LBR "
        StrSql = StrSql & " WHERE LEGAJO = '" & legajo & "' AND AUDNRO = " & audnro
        objConn.Execute StrSql, , adExecuteNoRecords
    End If

MyCommitTrans

Fin:
    If objInsert.State = adStateOpen Then objInsert.Close
    Set objInsert = Nothing
    Exit Sub
    
CE:
    HuboErrores = True
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 0) & "Error Function InsertarItemLegajoPosicion. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline Espacios(Tabulador * 0) & "Error: " & Err.Number
    Flog.writeline Espacios(Tabulador * 0) & "Decripcion: " & Err.Description
    Flog.writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "SQL Ejecutado: " & StrSql
        Flog.writeline
    End If
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline
    Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", audnro, False, "")
    GoTo Fin
End Sub

Private Sub InsertarItemLegajoLiquidar(ByVal tipo As String, ByVal legajo As String, ByVal empresa As String, ByVal Fecha_Ingreso As Date, ByVal Fecha_Antiguedad As Date, ByVal Ingreso_Grupo As Date, ByVal estado As String, ByVal relacion_laboral As String, ByVal DGI_exento As String, ByVal estudios_nivel As String, ByVal convenio As String, ByVal categoria_convenio As String, ByVal funcion_convenio As String, ByVal horas_mesh As Double, horas_mesd As Double, instrum_pago As String, ByVal tipo_cuenta As String, ByVal entid_pago As String, ByVal nro_cuenta As String, ByVal nro_cuil As String, ByVal osocial As String, ByVal Fecha_baja As Date, ByVal motivo_baja As String, ByVal provincia_trabaja As String, ByVal caracter_servicio As String, ByVal activo As String, ByVal suc_pago As String, ByVal grupo_liquidacion As String, ByVal Fecha_Promo As Date, ByVal CBU As String, ByVal no_error As Boolean, ByVal audnro As Long, ByVal ternro As Long)
'                                                           LEGAJO ,                EMPRESA ,               FECHA_INGRESO DATE ,       FECHA_ANTIGUEDAD DATE,           INGRESO_GRUPO DATE,             ESTADO ,                RELACION_LABORAL,               DGI_EXENTO,                 ESTUDIOS_NIVEL,                 CONVENIO,                   CATEGORIA_CONVENIO,                 FUNCION_CONVENIO,               HORAS_MESH NUMBER,          HORAS_MESD NUMBER,      INSTRUM_PAGO,           TIPO_CUENTA,                ENTID_PAGO,                 NRO_CUENTA,                 NRO_CUIL,               OSOCIAL,                FECHA_BAJA DATE,            MOTIVO_BAJA,                PROVINCIA_TRABAJA,                  CARACTER_SERVICIO,                  ACTIVO,                 SUC_PAGO,               GRUPO_LIQUIDACION,                  FECHA_PROMO_VAC DATE,       CBU
Dim objInsert As New ADODB.Recordset
Dim s_error As Boolean
Dim aux_empresa As String
Dim aux_fecha_ingreso
Dim strsql2
Dim v
Dim I
Dim Fecha_Antiguedad_aux
Dim Ingreso_Grupo_aux
Dim Fecha_baja_aux
Dim Fecha_Promo_aux
Dim horas_mesh_aux
Dim horas_mesd_aux
Dim estado_aux

On Error GoTo CE:

MyBeginTrans
    
    Fecha_Antiguedad_aux = IIf(CStr(Fecha_Antiguedad) = "12:00:00 a.m.", "Null", ConvFecha(Fecha_Antiguedad))
    Ingreso_Grupo_aux = IIf(CStr(Ingreso_Grupo) = "12:00:00 a.m.", "Null", ConvFecha(Ingreso_Grupo))
    Fecha_baja_aux = IIf(CStr(Fecha_baja) = "12:00:00 a.m.", "Null", ConvFecha(Fecha_baja))
    Fecha_Promo_aux = IIf(CStr(Fecha_Promo) = "12:00:00 a.m.", "Null", ConvFecha(Fecha_Promo))
    horas_mesh_aux = IIf(horas_mesh = 0, "Null", CDbl(horas_mesh))
    horas_mesd_aux = IIf(horas_mesd = 0, "Null", CDbl(horas_mesd))
    
    aux_empresa = empresa
    aux_fecha_ingreso = IIf(Not EsNulo(Fecha_Ingreso), CDate(Fecha_Ingreso), Empty)
    If InStr(1, tipo, "EMPRESA") = 0 Then
        'Buscar tipo estructura EMPRESA
        StrSql = "SELECT his_estructura.estrnro "
        StrSql = StrSql & " FROM his_estructura "
        StrSql = StrSql & " WHERE his_estructura.ternro = " & ternro & " AND his_estructura.tenro= " & C_empresa
        StrSql = StrSql & " ORDER BY htetdesde DESC"
        OpenRecordset StrSql, objInsert
        If objInsert.EOF Then
'            Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", audnro, True, StrSql)
            Call InsertarError(legajo, "LEGAJO_LIQUIDAR", "EMPRESA", "No tiene asignado una empresa", CLng(audnro), False, "")
        Else
            aux_empresa = Mapeo(legajo, "LEGAJO_LIQUIDAR", "EMPRESA", audnro, "[EMPRESA]", objInsert!estrnro, s_error)
            If Not s_error Then
                GoTo Fin
            End If
        End If
    End If
    
    If InStr(1, tipo, "FECHA_INGRESO") = 0 Then
        'Buscar FECHA_INGRESO
        StrSql = "SELECT empleado.empfaltagr, empleado.empleg "
        StrSql = StrSql & " FROM empleado "
        StrSql = StrSql & " WHERE empleado.ternro = " & ternro
        OpenRecordset StrSql, objInsert
        If objInsert.EOF Then
'            Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", audnro, True, StrSql)
            Call InsertarError(legajo, "LEGAJO_LIQUIDAR", "FECHA_INGRESO", "No tiene asignado FEuna empresa", CLng(audnro), False, "")
        Else
            aux_fecha_ingreso = objInsert!empfaltagr
        End If
    End If
    
    'Buscar ESTADO (A si es activo,  T si es baja)
    StrSql = "SELECT empleado.empest "
    StrSql = StrSql & " FROM empleado "
    StrSql = StrSql & " WHERE empleado.ternro = " & ternro
    OpenRecordset StrSql, objInsert
    estado_aux = "T"
    If objInsert.EOF Then
        If CInt(objInsert!empest) = -1 Then
            estado_aux = "A"
        End If
    End If

    If Not (EsNulo(legajo) Or EsNulo(aux_empresa) Or EsNulo(aux_fecha_ingreso)) Then
        StrSql = "SELECT * FROM legajo_liquidar_rhpro WHERE LEGAJO = '" & legajo & "' AND EMPRESA = '" & aux_empresa & "' AND FECHA_INGRESO = " & ConvFecha(aux_fecha_ingreso)
        OpenRecordset StrSql, objInsert
        
        If objInsert.EOF Then
            StrSql = "INSERT INTO legajo_liquidar_rhpro(legajo,EMPRESA,FECHA_INGRESO,FECHA_ANTIGUEDAD,INGRESO_GRUPO,"
            StrSql = StrSql & "ESTADO,RELACION_LABORAL,DGI_EXENTO,ESTUDIOS_NIVEL,CONVENIO,CATEGORIA_CONVENIO,"
            StrSql = StrSql & "FUNCION_CONVENIO,HORAS_MESH,HORAS_MESD,INSTRUM_PAGO,TIPO_CUENTA,ENTID_PAGO,NRO_CUENTA,"
            StrSql = StrSql & "NRO_CUIL,OSOCIAL,FECHA_BAJA,MOTIVO_BAJA,PROVINCIA_TRABAJA,CARACTER_SERVICIO,ACTIVO,"
            StrSql = StrSql & "SUC_PAGO,GRUPO_LIQUIDACION,FECHA_PROMO_VAC,CBU)"
            StrSql = StrSql & " VALUES ('" & legajo & "','"
            StrSql = StrSql & aux_empresa & "',"
            StrSql = StrSql & ConvFecha(aux_fecha_ingreso) & ","
            StrSql = StrSql & Fecha_Antiguedad_aux & ","
            StrSql = StrSql & Ingreso_Grupo_aux & ",'"
            StrSql = StrSql & estado_aux & "','"
            StrSql = StrSql & relacion_laboral & "','"
            StrSql = StrSql & DGI_exento & "','"
            StrSql = StrSql & estudios_nivel & "','"
            StrSql = StrSql & convenio & "','"
            StrSql = StrSql & categoria_convenio & "','"
            StrSql = StrSql & funcion_convenio & "',"
            StrSql = StrSql & horas_mesh_aux & ","
            StrSql = StrSql & horas_mesd_aux & ",'"
            StrSql = StrSql & instrum_pago & "','"
            StrSql = StrSql & tipo_cuenta & "','"
            StrSql = StrSql & entid_pago & "','"
            StrSql = StrSql & nro_cuenta & "','"
            StrSql = StrSql & nro_cuil & "','"
            StrSql = StrSql & osocial & "',"
            StrSql = StrSql & Fecha_baja_aux & ",'"
            StrSql = StrSql & motivo_baja & "','"
            StrSql = StrSql & provincia_trabaja & "','"
            StrSql = StrSql & caracter_servicio & "','"
            StrSql = StrSql & "S','"
            StrSql = StrSql & suc_pago & "','"
            StrSql = StrSql & grupo_liquidacion & "',"
            StrSql = StrSql & Fecha_Promo_aux & ",'"
            StrSql = StrSql & CBU & "')"
        Else
            StrSql = "UPDATE legajo_liquidar_rhpro SET "
            v = Split(tipo, ",")
            For I = 0 To UBound(v)
                Select Case v(I)
'                    Case "ALT_MOD_FASE":
'                        StrSql = StrSql & "FECHA_ANTIGUEDAD = " & ConvFecha(Fecha_Antiguedad)
'                        StrSql = StrSql & ",INGRESO_GRUPO=" & ConvFecha(Ingreso_Grupo)
'                        StrSql = StrSql & ",FECHA_PROMO_VAC = " & ConvFecha(Fecha_Promo) & ""
                    Case "FECHA_ANTIGUEDAD":
                        StrSql = StrSql & "FECHA_ANTIGUEDAD = " & Fecha_Antiguedad_aux
                    Case "INGRESO_GRUPO":
                        StrSql = StrSql & "INGRESO_GRUPO=" & Ingreso_Grupo_aux
                    Case "ESTADO":
                        StrSql = StrSql & "ESTADO = '" & estado_aux & "'"
                    Case "RELACION_LABORAL":
                        StrSql = StrSql & "RELACION_LABORAL = '" & relacion_laboral & "'"
                    Case "DGI_EXENTO":
                        StrSql = StrSql & "DGI_EXENTO = '" & DGI_exento & "'"
                    Case "ESTUDIOS_NIVEL":
                        StrSql = StrSql & "ESTUDIOS_NIVEL = '" & estudios_nivel & "'"
                    Case "CONVENIO":
                        StrSql = StrSql & "CONVENIO = '" & convenio & "'"
                    Case "CATEGORIA_CONVENIO":
                        StrSql = StrSql & "CATEGORIA_CONVENIO = '" & categoria_convenio & "'"
                    Case "FUNCION_CONVENIO":
                        StrSql = StrSql & "FUNCION_CONVENIO = '" & funcion_convenio & "'"
                    Case "HORAS_MESH":
                        StrSql = StrSql & "HORAS_MESH = " & horas_mesh_aux
    '                Case "HORAS_MESD":
                        StrSql = StrSql & ",HORAS_MESD = " & horas_mesd_aux
                    Case "INSTRUM_PAGO":
                        StrSql = StrSql & "INSTRUM_PAGO = '" & instrum_pago & "'"
                    Case "TIPO_CUENTA":
                        StrSql = StrSql & "TIPO_CUENTA = '" & tipo_cuenta & "'"
                    Case "ENTID_PAGO":
                        StrSql = StrSql & "ENTID_PAGO = '" & entid_pago & "'"
                    Case "NRO_CUENTA":
                        StrSql = StrSql & "NRO_CUENTA = '" & nro_cuenta & "'"
                    Case "NRO_CUIL":
                        StrSql = StrSql & "NRO_CUIL = '" & nro_cuil & "'"
                    Case "OSOCIAL":
                        StrSql = StrSql & "OSOCIAL = '" & osocial & "'"
                    Case "FECHA_BAJA":
                        StrSql = StrSql & "FECHA_BAJA = " & Fecha_baja_aux
                        StrSql = StrSql & ",MOTIVO_BAJA = '" & motivo_baja & "'"
                    Case "MOTIVO_BAJA":
                        StrSql = StrSql & "MOTIVO_BAJA = '" & motivo_baja & "'"
                    Case "PROVINCIA_TRABAJA":
                        StrSql = StrSql & "PROVINCIA_TRABAJA = '" & provincia_trabaja & "'"
                    Case "CARACTER_SERVICIO":
                        StrSql = StrSql & "CARACTER_SERVICIO = '" & caracter_servicio & "'"
        '            Case "ACTIVO":
        '                StrSql = StrSql & "ACTIVO = '" & activo & "',"
                    Case "SUC_PAGO":
                        StrSql = StrSql & "SUC_PAGO = '" & suc_pago & "'"
                    Case "GRUPO_LIQUIDACION":
                        StrSql = StrSql & "GRUPO_LIQUIDACION = '" & grupo_liquidacion & "'"
                    Case "FECHA_PROMO_VAC":
                        StrSql = StrSql & "FECHA_PROMO_VAC = " & Fecha_Promo_aux & ""
                    Case "CBU":
                        StrSql = StrSql & "CBU = '" & CBU & "'"
                    Case "ALTA_CTABANCARIA"
                        StrSql = StrSql & "TIPO_CUENTA = '" & tipo_cuenta & "'"
                        StrSql = StrSql & ",ENTID_PAGO = '" & entid_pago & "'"
                        StrSql = StrSql & ",NRO_CUENTA = '" & nro_cuenta & "'"
                        StrSql = StrSql & ",SUC_PAGO = '" & suc_pago & "'"
                        StrSql = StrSql & ",CBU = '" & CBU & "'"
                End Select
                StrSql = StrSql & ","
            Next
            StrSql = Left(StrSql, Len(StrSql) - 1)
            StrSql = StrSql & " WHERE legajo = '" & legajo & "'"
            StrSql = StrSql & " AND EMPRESA = '" & aux_empresa & "'"
            StrSql = StrSql & " AND FECHA_INGRESO = " & ConvFecha(aux_fecha_ingreso)
        End If
        
        If no_error Then
            objConn.Execute StrSql, , adExecuteNoRecords
            
            StrSql = "UPDATE auditoria SET procesado = -1 "
            StrSql = StrSql & " WHERE audnro = " & audnro
            objConn.Execute StrSql, , adExecuteNoRecords
            
            StrSql = "DELETE FROM error_LBR "
            StrSql = StrSql & " WHERE LEGAJO = '" & legajo & "' AND AUDNRO = " & audnro
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
    End If

MyCommitTrans

Fin:
    If objInsert.State = adStateOpen Then objInsert.Close
    Set objInsert = Nothing
    Exit Sub
    
CE:
'Resume Next
    HuboErrores = True
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 0) & "Error Function InsertarItemLegajoLiquidar. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline Espacios(Tabulador * 0) & "Error: " & Err.Number
    Flog.writeline Espacios(Tabulador * 0) & "Decripcion: " & Err.Description
    Flog.writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "SQL Ejecutado: " & StrSql
        Flog.writeline
    End If
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline
    Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", audnro, False, "")
    GoTo Fin
End Sub

Private Sub InsertarError(ByVal p_legajo As String, ByVal p_tabla As String, ByVal p_campo As String, ByVal p_error As String, ByVal p_audnro As Long, ByVal p_informar_log As Boolean, ByVal p_StrSql As String)
Dim objInsert As New ADODB.Recordset

    StrSql = "SELECT * FROM error_LBR WHERE LEGAJO = '" & p_legajo & "' AND AUDNRO =" & p_audnro & " AND ERROR ='" & p_error & "'"
    OpenRecordset StrSql, objInsert
        
    If objInsert.EOF Then
        StrSql = "INSERT INTO error_LBR (LEGAJO, TABLA, CAMPO, ERROR, AUDNRO) "
        StrSql = StrSql & " VALUES ('" & p_legajo & "', '" & p_tabla & "', '" & p_campo & "', '" & p_error & "', " & p_audnro & ")"
        
        objConn.Execute StrSql, , adExecuteNoRecords
            
        If p_informar_log Then
            Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
            Flog.writeline Espacios(Tabulador * 1) & "WARNING. No se proceso la AUDITORIA (audnro = " & p_audnro & ")"
            Flog.writeline Espacios(Tabulador * 1) & "         SQL ==> " & p_StrSql
            Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
        End If
    End If

End Sub

Private Sub ConfReporte()
Dim objRs As New ADODB.Recordset
Dim objrS1 As New ADODB.Recordset

On Error GoTo CE:

    Flog.writeline Espacios(Tabulador * 0) & "Comienza la lectura de configuración del reporte (confrep)."
    
    StrSql = "SELECT * FROM confrep "
    StrSql = StrSql & " WHERE repnro = 252"
   
    OpenRecordset StrSql, objRs
    
    Do Until objRs.EOF
        Select Case objRs!confnrocol
            Case 1:
                If Not (IsNull(objRs!confval) Or objRs!confval = "") Then
                    StrSql = "SELECT cnstring FROM conexion WHERE cnnro = " & objRs!confval
                    OpenRecordset StrSql, objrS1
                    If Not objrS1.EOF Then
                        conexion_base_LBR = objrS1!cnstring
                    End If
                    objrS1.Close
                End If
            Case 2:
                If Not (IsNull(objRs!confval2) Or objRs!confval2 = "") Then
                    formatofecha = objRs!confval2
                End If
            Case 3:
                If Not (IsNull(objRs!confval) Or objRs!confval = "") Then
                    C_puesto = objRs!confval
                End If
            Case 4:
                If Not (IsNull(objRs!confval) Or objRs!confval = "") Then
                    C_grado = objRs!confval
                End If
            Case 5:
                If Not (IsNull(objRs!confval) Or objRs!confval = "") Then
                    C_empresa = objRs!confval
                End If
            Case 6:
                If Not (IsNull(objRs!confval) Or objRs!confval = "") Then
                    C_ccosto = objRs!confval
                End If
            Case 7:
                If Not (IsNull(objRs!confval) Or objRs!confval = "") Then
                    C_estruct_area = objRs!confval
                End If
            Case 8:
                If Not (IsNull(objRs!confval) Or objRs!confval = "") Then
                    C_centro_trabajo = objRs!confval
                End If
            Case 9:
                If Not (IsNull(objRs!confval) Or objRs!confval = "") Then
                    C_edificio = objRs!confval
                End If
            Case 10:
                If Not (IsNull(objRs!confval) Or objRs!confval = "") Then
                    C_canal_costo = objRs!confval
                End If
            Case 11:
                If Not (IsNull(objRs!confval) Or objRs!confval = "") Then
                    C_rel_laboral = objRs!confval
                End If
            Case 12:
                If Not (IsNull(objRs!confval) Or objRs!confval = "") Then
                    C_dgi_exento = objRs!confval
                End If
            Case 13:
                If Not (IsNull(objRs!confval) Or objRs!confval = "") Then
                    C_convenio = objRs!confval
                End If
            Case 14:
                If Not (IsNull(objRs!confval) Or objRs!confval = "") Then
                    C_cat_conv = objRs!confval
                End If
            Case 15:
                If Not (IsNull(objRs!confval) Or objRs!confval = "") Then
                    C_fun_conv = objRs!confval
                End If
            Case 16:
                If Not (IsNull(objRs!confval) Or objRs!confval = "") Then
                    C_horas_mesh = objRs!confval
                End If
            Case 17:
                If Not (IsNull(objRs!confval) Or objRs!confval = "") Then
                    C_horas_mesd = objRs!confval
                End If
            Case 18:
                If Not (IsNull(objRs!confval) Or objRs!confval = "") Then
                    C_osocial = objRs!confval
                End If
            Case 19:
                If Not (IsNull(objRs!confval) Or objRs!confval = "") Then
                    C_caracter_serv = objRs!confval
                End If
            Case 20:
                If Not (IsNull(objRs!confval) Or objRs!confval = "") Then
                    C_grupo_liq = objRs!confval
                End If
            Case 21:
                If Not (IsNull(objRs!confval) Or objRs!confval = "") Then
                    C_instrum_pago = objRs!confval
                End If
        End Select
        objRs.MoveNext
    Loop
    
    objRs.Close
    
    If IsNull(conexion_base_LBR) Or conexion_base_LBR = "" Then
        Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontró el número de conexión definido en el Valor Numérico de la Columna 1 (Nro Columna)"
        Flog.writeline Espacios(Tabulador * 1) & "       Dicha conexión se debe definir en 'Supervisor/Instalación/Herramientas/Conexión de Bases de Datos'."
        Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
        HuboErrores = True
        GoTo Fin:
    End If
    
    If IsNull(formatofecha) Or formatofecha = "" Then
        Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
        Flog.writeline Espacios(Tabulador * 1) & "WARNING. No se encontró el formato de fecha, definido en el Valor Alfanumérico de la Columna 2 (Nro Columna)"
        Flog.writeline Espacios(Tabulador * 1) & "         Por defecto se considera: dd/mm/yyyy"
        Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
        formatofecha = "dd/mm/yyyy"
    End If
    
    If IsNull(C_puesto) Or C_puesto = 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontró código del Tipo de Estructura para el PUESTO, definido en el Valor Numérico de la Columna 3 (Nro Columna)"
        Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
        HuboErrores = True
        GoTo Fin:
    End If
    
    If IsNull(C_grado) Or C_grado = 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontró código del Tipo de Estructura para el GRADO, definido en el Valor Numérico de la Columna 4 (Nro Columna)"
        Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
        HuboErrores = True
        GoTo Fin:
    End If
    
    If IsNull(C_empresa) Or C_empresa = 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontró código del Tipo de Estructura para el EMPRESA, definido en el Valor Numérico de la Columna 5 (Nro Columna)"
        Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
        HuboErrores = True
        GoTo Fin:
    End If
    
    If IsNull(C_ccosto) Or C_ccosto = 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontró código del Tipo de Estructura para el CENTRO_COSTO, definido en el Valor Numérico de la Columna 6 (Nro Columna)"
        Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
        HuboErrores = True
        GoTo Fin:
    End If
    
    If IsNull(C_estruct_area) Or C_estruct_area = 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontró código del Tipo de Estructura para el ESTRUC_AREA, definido en el Valor Numérico de la Columna 7 (Nro Columna)"
        Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
        HuboErrores = True
        GoTo Fin:
    End If
    
    If IsNull(C_centro_trabajo) Or C_centro_trabajo = 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontró código del Tipo de Estructura para el CENTRO_TRABAJO, definido en el Valor Numérico de la Columna 8 (Nro Columna)"
        Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
        HuboErrores = True
        GoTo Fin:
    End If
    
    If IsNull(C_edificio) Or C_edificio = 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontró código del Tipo de Estructura para el EDIFICIO, definido en el Valor Numérico de la Columna 9 (Nro Columna)"
        Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
        HuboErrores = True
        GoTo Fin:
    End If
    
    If IsNull(C_canal_costo) Or C_canal_costo = 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontró código del Tipo de Estructura para el CANAL_COSTO, definido en el Valor Numérico de la Columna 10 (Nro Columna)"
        Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
        HuboErrores = True
        GoTo Fin:
    End If
    
    If IsNull(C_rel_laboral) Or C_rel_laboral = 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontró código del Tipo de Estructura para la RELACION_LABORAL, definido en el Valor Numérico de la Columna 11 (Nro Columna)"
        Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
        HuboErrores = True
        GoTo Fin:
    End If

    If IsNull(C_dgi_exento) Or C_dgi_exento = 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontró código del Tipo de Estructura para DGI_EXENTO, definido en el Valor Numérico de la Columna 12 (Nro Columna)"
        Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
        HuboErrores = True
        GoTo Fin:
    End If

    If IsNull(C_convenio) Or C_convenio = 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontró código del Tipo de Estructura para CONVENIO, definido en el Valor Numérico de la Columna 13 (Nro Columna)"
        Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
        HuboErrores = True
        GoTo Fin:
    End If
    
    If IsNull(C_cat_conv) Or C_cat_conv = 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontró código del Tipo de Estructura para CATEGORIA_CONVENIO, definido en el Valor Numérico de la Columna 14 (Nro Columna)"
        Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
        HuboErrores = True
        GoTo Fin:
    End If
    
    If IsNull(C_fun_conv) Or C_fun_conv = 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontró código del Tipo de Estructura para FUNCION_CONVENIO, definido en el Valor Numérico de la Columna 15 (Nro Columna)"
        Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
        HuboErrores = True
        GoTo Fin:
    End If
    
    If IsNull(C_horas_mesh) Or C_horas_mesh = 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontró código del Tipo de Estructura para HORAS_MESH, definido en el Valor Numérico de la Columna 16 (Nro Columna)"
        Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
        HuboErrores = True
        GoTo Fin:
    End If
    
    If IsNull(C_horas_mesd) Or C_horas_mesd = 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontró código del Tipo de Estructura para HORAS_MESD, definido en el Valor Numérico de la Columna 17 (Nro Columna)"
        Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
        HuboErrores = True
        GoTo Fin:
    End If
    
    If IsNull(C_osocial) Or C_osocial = 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontró código del Tipo de Estructura para OSOCIAL, definido en el Valor Numérico de la Columna 18 (Nro Columna)"
        Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
        HuboErrores = True
        GoTo Fin:
    End If
    
    If IsNull(C_caracter_serv) Or C_caracter_serv = 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontró código del Tipo de Estructura para CARACTER_SERVICIO, definido en el Valor Numérico de la Columna 19 (Nro Columna)"
        Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
        HuboErrores = True
        GoTo Fin:
    End If
    
    If IsNull(C_grupo_liq) Or C_grupo_liq = 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontró código del Tipo de Estructura para GRUPO_LIQUIDACION, definido en el Valor Numérico de la Columna 20 (Nro Columna)"
        Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
        HuboErrores = True
        GoTo Fin:
    End If

    If IsNull(C_instrum_pago) Or C_instrum_pago = 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontró código del Tipo de Estructura para INSTRUM_PAGO, definido en el Valor Numérico de la Columna 21 (Nro Columna)"
        Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
        HuboErrores = True
        GoTo Fin:
    End If
    
    Flog.writeline Espacios(Tabulador * 0) & "Finaliza la lectura de configuración del reporte (confrep)."

Fin:
    If HuboErrores Then
        Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", 0, False, "")
    End If
    Exit Sub

CE:
    HuboErrores = True
'    Resume Next
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 0) & "Error Function ConfReporte. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline Espacios(Tabulador * 0) & "Error: " & Err.Number
    Flog.writeline Espacios(Tabulador * 0) & "Decripcion: " & Err.Description
    Flog.writeline
    If InStr(1, Err.Description, "ODBC") > 0 Then
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "SQL Ejecutado: " & StrSql
        Flog.writeline
    End If
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline
    Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede procesar. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", 0, False, "")
    GoTo Fin
End Sub


Private Sub Comenzarmigracion()
Dim objRs As New ADODB.Recordset
Dim objRsLob As New ADODB.Recordset

'    On Error Resume Next
'
'    objconnBaseLobruno.CursorLocation = adUseClient
'    objconnBaseLobruno.IsolationLevel = adXactCursorStability
'    objconnBaseLobruno.CommandTimeout = 60
'    objconnBaseLobruno.ConnectionTimeout = 60
'    objconnBaseLobruno.Open conexion_base_LBR
'
'    If Err.Number <> 0 Then
'        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion Base Lobruno"
'        Call InsertarError("0", "GENERAL", "GENERAL", "Error. No se puede conectar Base Lobruno. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", 0, False, "")
'        Exit Sub
'    End If
'
'    On Error GoTo CE_Comenzarmigracion:
'
'MyBeginTrans
'
''-- LEGAJO -----------------
'    StrSql = "select * from legajo "
'
'    OpenRecordset StrSql, objRs
'
'    CEmpleadosAProc = objRs.RecordCount
'    If CEmpleadosAProc = 0 Then
'        CEmpleadosAProc = 1
'    End If
'    IncPorc = (1 / CEmpleadosAProc)
'
'    Do Until objRs.EOF
'        StrSql = "select * from legajos_rhpro where legajo='" & objRs!legajo & "'"
'
'        If objRsLob.State <> adStateClosed Then
'            If objRsLob.lockType <> adLockReadOnly Then objRsLob.UpdateBatch
'            objRsLob.Close
'        End If
'        objRsLob.CacheSize = 500
'        objRsLob.Open StrSql, objconnBaseLobruno, adOpenDynamic, adLockReadOnly, adCmdText
'
'        If objRsLob.EOF Then
'            'Insert
'            StrSql = "INSERT INTO legajos_rhpro (LEGAJO,APELLIDO_NOMBRE,DOCUMENTO_EMISOR,SEXO,ESTADO_CIVIL,APELLIDO_CASADA,FECHA_NACIMIENTO,NACION_NACIMIENTO,LUGAR_NACIMIENTO,NACIONALIDAD,DOCUMENTO_TIPO,DOCUMENTO_NUMERO,EMPRESA)"
'            StrSql = StrSql & " VALUES ('" & objRs!legajo & "','"
'            StrSql = StrSql & objRs!apellido_nombre & "','"
'            StrSql = StrSql & "RNP','"
'            StrSql = StrSql & objRs!sexo & "','"
'            StrSql = StrSql & objRs!ESTADO_CIVIL & "','"
'            StrSql = StrSql & objRs!apellido_casada & "','"
'            StrSql = StrSql & Format(objRs!FECHA_NACIMIENTO, formatofecha) & "','"
'            StrSql = StrSql & objRs!NACION_NACIMIENTO & "','"
'            StrSql = StrSql & objRs!LUGAR_NACIMIENTO & "','"
'            StrSql = StrSql & objRs!nacionalidad & "','"
'            StrSql = StrSql & objRs!DOCUMENTO_TIPO & "','"
'            StrSql = StrSql & objRs!DOCUMENTO_NUMERO & "','"
'            StrSql = StrSql & objRs!empresa & "')"
'        Else
'            'Update
'            StrSql = "UPDATE legajos_rhpro SET "
'            If Not EsNulo(objRs!apellido_nombre) Then
'                StrSql = StrSql & "APELLIDO_NOMBRE = '" & objRs!apellido_nombre & "',"
'            End If
'            If Not EsNulo(objRs!documento_emisor) Then
'                StrSql = StrSql & "DOCUMENTO_EMISOR = '" & objRs!documento_emisor & "',"
'            End If
'            If Not EsNulo(objRs!sexo) Then
'                StrSql = StrSql & "SEXO = '" & objRs!sexo & "',"
'            End If
'            If Not EsNulo(objRs!ESTADO_CIVIL) Then
'                StrSql = StrSql & "ESTADO_CIVIL = '" & objRs!ESTADO_CIVIL & "',"
'            End If
'            If Not EsNulo(objRs!apellido_casada) Then
'                StrSql = StrSql & "APELLIDO_CASADA = '" & objRs!apellido_casada & "',"
'            End If
'            If Not EsNulo(objRs!FECHA_NACIMIENTO) Then
'                StrSql = StrSql & "FECHA_NACIMIENTO = '" & Format(objRs!FECHA_NACIMIENTO, formatofecha) & "',"
'            End If
'            If Not EsNulo(objRs!NACION_NACIMIENTO) Then
'                StrSql = StrSql & "NACION_NACIMIENTO = '" & objRs!NACION_NACIMIENTO & "',"
'            End If
'            If Not EsNulo(objRs!LUGAR_NACIMIENTO) Then
'                StrSql = StrSql & "LUGAR_NACIMIENTO = '" & objRs!LUGAR_NACIMIENTO & "',"
'            End If
'            If Not EsNulo(objRs!nacionalidad) Then
'                StrSql = StrSql & "NACIONALIDAD = '" & objRs!nacionalidad & "',"
'            End If
'            If Not EsNulo(objRs!DOCUMENTO_TIPO) Then
'                StrSql = StrSql & "DOCUMENTO_TIPO = '" & objRs!DOCUMENTO_TIPO & "',"
'            End If
'            If Not EsNulo(objRs!empresa) Then
'                StrSql = StrSql & "EMPRESA= '" & objRs!empresa & "',"
'            End If
'            StrSql = Left(StrSql, Len(StrSql) - 1)
'            StrSql = StrSql & " WHERE LEGAJO = '" & objRs!legajo & "'"
'        End If
'        objconnBaseLobruno.Execute StrSql, , adExecuteNoRecords
'
'        'Actualizo el progreso
'        Progreso = Progreso + IncPorc
'        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & " WHERE bpronro = " & NroProceso
'        objconnProgreso.Execute StrSql, , adExecuteNoRecords
'
'        objRs.MoveNext
'    Loop
'
'    ' Borro los datos en la tabla temporal
'    StrSql = "DELETE FROM legajo"
'    objConn.Execute StrSql, , adExecuteNoRecords
'
''-- LEGAJO FAMILIARES -----------------
'    StrSql = "select * from legajo_familiares "
'
'    OpenRecordset StrSql, objRs
'
'    CEmpleadosAProc = objRs.RecordCount
'    If CEmpleadosAProc = 0 Then
'        CEmpleadosAProc = 1
'    End If
'    IncPorc = (1 / CEmpleadosAProc)
'
'    Do Until objRs.EOF
'        StrSql = "select * from legajo_familiares_rhpro where legajo='" & objRs!legajo & "' and FAMILIAR_NUMERO='" & objRs!FAMILIAR_NUMERO & "'"
'
'        If objRsLob.State <> adStateClosed Then
'            If objRsLob.lockType <> adLockReadOnly Then objRsLob.UpdateBatch
'            objRsLob.Close
'        End If
'        objRsLob.CacheSize = 500
'        objRsLob.Open StrSql, objconnBaseLobruno, adOpenDynamic, adLockReadOnly, adCmdText
'
'        If objRsLob.EOF Then
'            'INSERT
'            StrSql = "INSERT INTO legajo_familiares_rhpro (LEGAJO,FAMILIAR_NUMERO,PARENTESCO,APELLIDO_NOMBRE,FECHA_NACIMIENTO,NACIONALIDAD,NACION_NACIMIENTO,LUGAR_NACIMIENTO,DOCUMENTO_TIPO,DOCUMENTO_NUMERO,SEXO,ESTADO_CIVIL,CUIL_FAM,FECHA_INICIO_VINCULO) "
'            StrSql = StrSql & " VALUES ('" & objRs!legajo & "','"
'            StrSql = StrSql & objRs!FAMILIAR_NUMERO & "','"
'            StrSql = StrSql & objRs!parentesco & "','"
'            StrSql = StrSql & objRs!apellido_nombre & "','"
'            StrSql = StrSql & Format(objRs!FECHA_NACIMIENTO, formatofecha) & "','"
'            StrSql = StrSql & objRs!nacionalidad & "','"
'            StrSql = StrSql & objRs!NACION_NACIMIENTO & "','"
'            StrSql = StrSql & objRs!LUGAR_NACIMIENTO & "','"
'            StrSql = StrSql & objRs!DOCUMENTO_TIPO & "','"
'            StrSql = StrSql & objRs!DOCUMENTO_NUMERO & "','"
'            StrSql = StrSql & objRs!sexo & "','"
'            StrSql = StrSql & objRs!ESTADO_CIVIL & "','"
'            StrSql = StrSql & objRs!CUIL_FAM & "','"
'            StrSql = StrSql & Format(objRs!FECHA_INICIO_VINCULO, formatofecha) & "')"
'        Else
'            StrSql = "UPDATE legajo_familiares_rhpro SET "
'            If Not EsNulo(objRs!parentesco) Then
'                StrSql = StrSql & "PARENTESCO= '" & objRs!parentesco & "',"
'            End If
'            If Not EsNulo(objRs!apellido_nombre) Then
'                StrSql = StrSql & "APELLIDO_NOMBRE= '" & objRs!apellido_nombre & "',"
'            End If
'            If Not EsNulo(objRs!FECHA_NACIMIENTO) Then
'                StrSql = StrSql & "FECHA_NACIMIENTO= '" & Format(objRs!FECHA_NACIMIENTO, formatofecha) & "',"
'            End If
'            If Not EsNulo(objRs!nacionalidad) Then
'                StrSql = StrSql & "NACIONALIDAD= '" & objRs!nacionalidad & "',"
'            End If
'            If Not EsNulo(objRs!NACION_NACIMIENTO) Then
'                StrSql = StrSql & "NACION_NACIMIENTO= '" & objRs!NACION_NACIMIENTO & "',"
'            End If
'            If Not EsNulo(objRs!LUGAR_NACIMIENTO) Then
'                StrSql = StrSql & "LUGAR_NACIMIENTO= '" & objRs!LUGAR_NACIMIENTO & "',"
'            End If
'            If Not EsNulo(objRs!DOCUMENTO_TIPO) Then
'                StrSql = StrSql & "DOCUMENTO_TIPO= '" & objRs!DOCUMENTO_TIPO & "',"
'            End If
'            If Not EsNulo(objRs!DOCUMENTO_NUMERO) Then
'                StrSql = StrSql & "DOCUMENTO_NUMERO= '" & objRs!DOCUMENTO_NUMERO & "',"
'            End If
'            If Not EsNulo(objRs!sexo) Then
'                StrSql = StrSql & "SEXO= '" & objRs!sexo & "',"
'            End If
'            If Not EsNulo(objRs!ESTADO_CIVIL) Then
'                StrSql = StrSql & "ESTADO_CIVIL= '" & objRs!ESTADO_CIVIL & "',"
'            End If
'            If Not EsNulo(objRs!CUIL_FAM) Then
'                StrSql = StrSql & "CUIL_FAM= '" & objRs!CUIL_FAM & "',"
'            End If
'            If Not EsNulo(objRs!FECHA_INICIO_VINCULO) Then
'                StrSql = StrSql & "FECHA_INICIO_VINCULO= '" & Format(objRs!FECHA_INICIO_VINCULO, formatofecha) & "',"
'            End If
'            StrSql = Left(StrSql, Len(StrSql) - 1)
'            StrSql = StrSql & " where legajo = '" & objRs!legajo & "' "
'            StrSql = StrSql & " and FAMILIAR_NUMERO= '" & objRs!FAMILIAR_NUMERO & "'"
'        End If
'        objconnBaseLobruno.Execute StrSql, , adExecuteNoRecords
'
'        'Actualizo el progreso
'        Progreso = Progreso + IncPorc
'        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & " WHERE bpronro = " & NroProceso
'        objconnProgreso.Execute StrSql, , adExecuteNoRecords
'
'        objRs.MoveNext
'    Loop
'
'    ' Borro los datos en la tabla temporal
'    StrSql = "DELETE FROM legajo_familiares"
'    objConn.Execute StrSql, , adExecuteNoRecords
'
''-- LEGAJO POSICION -----------------
'    StrSql = "select * from legajo_posicion "
'
'    OpenRecordset StrSql, objRs
'
'    CEmpleadosAProc = objRs.RecordCount
'    If CEmpleadosAProc = 0 Then
'        CEmpleadosAProc = 1
'    End If
'    IncPorc = (1 / CEmpleadosAProc)
'
'    Do Until objRs.EOF
'        StrSql = "SELECT * FROM legajo_posicion_rhpro WHERE LEGAJO = '" & objRs!legajo & "' AND FECHA = '" & Format(objRs!Fecha, formatofecha) & "'"
'
'        If objRsLob.State <> adStateClosed Then
'            If objRsLob.lockType <> adLockReadOnly Then objRsLob.UpdateBatch
'            objRsLob.Close
'        End If
'        objRsLob.CacheSize = 500
'        objRsLob.Open StrSql, objconnBaseLobruno, adOpenDynamic, adLockReadOnly, adCmdText
'
'        If objRsLob.EOF Then
'            'INSERT
'            StrSql = "INSERT INTO legajo_posicion_rhpro (legajo,fecha,puesto,puesto_cambio,grado,grado_cambio,empresa,"
'            StrSql = StrSql & "empresa_cambio,centro_costo,c_costo_cambio,estruc_area,activo,grupo_tablas,centro_trabajo,"
'            StrSql = StrSql & "c_trabajo_cambio,edificio,area_cambio,canal_costo,cnl_costo_cambio)"
'            StrSql = StrSql & " VALUES ('" & objRs!legajo & "','"
'            StrSql = StrSql & Format(objRs!Fecha, formatofecha) & "','"
'            StrSql = StrSql & objRs!puesto & "','"
'            StrSql = StrSql & objRs!puesto_cambio & "','"
'            StrSql = StrSql & objRs!grado & "','"
'            StrSql = StrSql & objRs!grado_cambio & "','"
'            StrSql = StrSql & objRs!empresa & "','"
'            StrSql = StrSql & objRs!empresa_cambio & "','"
'            StrSql = StrSql & objRs!centro_costo & "','"
'            StrSql = StrSql & objRs!c_costo_cambio & "','"
'            StrSql = StrSql & objRs!estruc_area & "','"
'            StrSql = StrSql & "S','"
'            StrSql = StrSql & "01','"
'            StrSql = StrSql & objRs!centro_trabajo & "','"
'            StrSql = StrSql & objRs!c_trabajo_cambio & "','"
'            StrSql = StrSql & objRs!edificio & "','"
'            StrSql = StrSql & objRs!area_cambio & "','"
'            StrSql = StrSql & objRs!canal_costo & "','"
'            StrSql = StrSql & objRs!cnl_costo_cambio & "')"
'        Else
'            StrSql = "UPDATE legajo_posicion_rhpro SET "
'            If Not EsNulo(objRs!puesto) Then
'                StrSql = StrSql & "puesto='" & objRs!puesto & "',"
'                StrSql = StrSql & "puesto_cambio='" & objRs!puesto_cambio & "',"
'            End If
'            If Not EsNulo(objRs!grado) Then
'                StrSql = StrSql & "grado = '" & objRs!grado & "',"
'                StrSql = StrSql & "grado_cambio = '" & objRs!grado_cambio & "',"
'            End If
'            If Not EsNulo(objRs!empresa) Then
'                StrSql = StrSql & "empresa = '" & objRs!empresa & "',"
'                StrSql = StrSql & "empresa_cambio = '" & objRs!empresa_cambio & "',"
'            End If
'            If Not EsNulo(objRs!centro_costo) Then
'                StrSql = StrSql & "centro_costo = '" & objRs!centro_costo & "',"
'                StrSql = StrSql & "c_costo_cambio= '" & objRs!c_costo_cambio & "',"
'            End If
'            If Not EsNulo(objRs!estruc_area) Then
'                StrSql = StrSql & "estruc_area = '" & objRs!estruc_area & "',"
'                StrSql = StrSql & "area_cambio = '" & objRs!area_cambio & "',"
'            End If
'            If Not EsNulo(objRs!centro_trabajo) Then
'                StrSql = StrSql & "centro_trabajo = '" & objRs!centro_trabajo & "',"
'                StrSql = StrSql & "c_trabajo_cambio= '" & objRs!c_trabajo_cambio & "',"
'            End If
'            If Not EsNulo(objRs!edificio) Then
'                StrSql = StrSql & "edificio = '" & objRs!edificio & "',"
'            End If
'            If Not EsNulo(objRs!canal_costo) Then
'                StrSql = StrSql & "canal_costo = '" & objRs!canal_costo & "',"
'                StrSql = StrSql & "cnl_costo_cambio = '" & objRs!cnl_costo_cambio & "',"
'            End If
'            StrSql = Left(StrSql, Len(StrSql) - 1)
'            StrSql = StrSql & " WHERE legajo = '" & objRs!legajo & "'"
'            StrSql = StrSql & " AND fecha = '" & Format(objRs!Fecha, formatofecha) & "'"
'        End If
'
'        objconnBaseLobruno.Execute StrSql, , adExecuteNoRecords
'
'        'Actualizo el progreso
'        Progreso = Progreso + IncPorc
'        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & " WHERE bpronro = " & NroProceso
'        objconnProgreso.Execute StrSql, , adExecuteNoRecords
'
'        objRs.MoveNext
'    Loop
'
'    ' Borro los datos en la tabla temporal
'    StrSql = "DELETE FROM legajo_posicion"
'    objConn.Execute StrSql, , adExecuteNoRecords
'
''-- LEGAJO LIQUIDAR -----------------
'    StrSql = "select * from legajo_liquidar "
'
'    OpenRecordset StrSql, objRs
'
'    CEmpleadosAProc = objRs.RecordCount
'    If CEmpleadosAProc = 0 Then
'        CEmpleadosAProc = 1
'    End If
'    IncPorc = (1 / CEmpleadosAProc)
'
'    Do Until objRs.EOF
'        StrSql = "SELECT * FROM legajo_liquidar_rhpro WHERE LEGAJO = '" & objRs!legajo & "' AND EMPRESA = '" & objRs!empresa & "' AND FECHA_INGRESO = '" & Format(objRs!Fecha_Ingreso, formatofecha) & "'"
'
'        If objRsLob.State <> adStateClosed Then
'            If objRsLob.lockType <> adLockReadOnly Then objRsLob.UpdateBatch
'            objRsLob.Close
'        End If
'        objRsLob.CacheSize = 500
'        objRsLob.Open StrSql, objconnBaseLobruno, adOpenDynamic, adLockReadOnly, adCmdText
'
'        If objRsLob.EOF Then
'            'INSERT
'            StrSql = "INSERT INTO legajo_liquidar_rhpro(legajo,EMPRESA,FECHA_INGRESO,FECHA_ANTIGUEDAD,INGRESO_GRUPO,"
'            StrSql = StrSql & "ESTADO,RELACION_LABORAL,DGI_EXENTO,ESTUDIOS_NIVEL,CONVENIO,CATEGORIA_CONVENIO,"
'            StrSql = StrSql & "FUNCION_CONVENIO,HORAS_MESH,HORAS_MESD,INSTRUM_PAGO,TIPO_CUENTA,ENTID_PAGO,NRO_CUENTA,"
'            StrSql = StrSql & "NRO_CUIL,OSOCIAL,FECHA_BAJA,MOTIVO_BAJA,PROVINCIA_TRABAJA,CARACTER_SERVICIO,ACTIVO,"
'            StrSql = StrSql & "SUC_PAGO,GRUPO_LIQUIDACION,FECHA_PROMO_VAC,CBU)"
'            StrSql = StrSql & " VALUES ('" & objRs!legajo & "','"
'            StrSql = StrSql & objRs!empresa & "','"
'            StrSql = StrSql & Format(objRs!Fecha_Ingreso, formatofecha) & "','"
'            StrSql = StrSql & Format(objRs!Fecha_Antiguedad, formatofecha) & "','"
'            StrSql = StrSql & Format(objRs!Ingreso_Grupo, formatofecha) & "','"
'            StrSql = StrSql & objRs!estado & "','"
'            StrSql = StrSql & objRs!relacion_laboral & "','"
'            StrSql = StrSql & objRs!DGI_exento & "','"
'            StrSql = StrSql & objRs!estudios_nivel & "','"
'            StrSql = StrSql & objRs!convenio & "','"
'            StrSql = StrSql & objRs!categoria_convenio & "','"
'            StrSql = StrSql & objRs!funcion_convenio & "',"
'            StrSql = StrSql & objRs!horas_mesh & ","
'            StrSql = StrSql & objRs!horas_mesd & ",'"
'            StrSql = StrSql & objRs!instrum_pago & "','"
'            StrSql = StrSql & objRs!tipo_cuenta & "','"
'            StrSql = StrSql & objRs!entid_pago & "','"
'            StrSql = StrSql & objRs!nro_cuenta & "','"
'            StrSql = StrSql & objRs!nro_cuil & "','"
'            StrSql = StrSql & objRs!osocial & "','"
'            StrSql = StrSql & Format(objRs!Fecha_baja, formatofecha) & "','"
'            StrSql = StrSql & objRs!motivo_baja & "','"
'            StrSql = StrSql & objRs!provincia_trabaja & "','"
'            StrSql = StrSql & objRs!caracter_servicio & "','"
'            StrSql = StrSql & "S','"
'            StrSql = StrSql & objRs!suc_pago & "','"
'            StrSql = StrSql & objRs!grupo_liquidacion & "','"
'            StrSql = StrSql & Format(objRs!Fecha_Promo_vac, formatofecha) & "','"
'            StrSql = StrSql & objRs!CBU & "')"
'        Else
'            StrSql = "UPDATE legajo_liquidar_rhpro SET "
'            If Not EsNulo(objRs!Fecha_Antiguedad) Then
'                StrSql = StrSql & "FECHA_ANTIGUEDAD = '" & Format(objRs!Fecha_Antiguedad, formatofecha) & "',"
'            End If
'            If Not EsNulo(objRs!Ingreso_Grupo) Then
'                StrSql = StrSql & "INGRESO_GRUPO='" & Format(objRs!Ingreso_Grupo, formatofecha) & "',"
'            End If
'            If Not EsNulo(objRs!estado) Then
'                StrSql = StrSql & "ESTADO = '" & objRs!estado & "',"
'            End If
'            If Not EsNulo(objRs!relacion_laboral) Then
'                StrSql = StrSql & "RELACION_LABORAL = '" & objRs!relacion_laboral & "',"
'            End If
'            If Not EsNulo(objRs!DGI_exento) Then
'                StrSql = StrSql & "DGI_EXENTO = '" & objRs!DGI_exento & "',"
'            End If
'            If Not EsNulo(objRs!estudios_nivel) Then
'                StrSql = StrSql & "ESTUDIOS_NIVEL = '" & objRs!estudios_nivel & "',"
'            End If
'            If Not EsNulo(objRs!convenio) Then
'                StrSql = StrSql & "CONVENIO = '" & objRs!convenio & "',"
'            End If
'            If Not EsNulo(objRs!categoria_convenio) Then
'                StrSql = StrSql & "CATEGORIA_CONVENIO = '" & objRs!categoria_convenio & "',"
'            End If
'            If Not EsNulo(objRs!funcion_convenio) Then
'                StrSql = StrSql & "FUNCION_CONVENIO = '" & objRs!funcion_convenio & "',"
'            End If
'            If Not EsNulo(objRs!horas_mesh) Then
'                StrSql = StrSql & "HORAS_MESH = " & objRs!horas_mesh & ","
'            End If
'            If Not EsNulo(objRs!horas_mesd) Then
'                StrSql = StrSql & "HORAS_MESD = " & objRs!horas_mesd & ","
'            End If
'            If Not EsNulo(objRs!instrum_pago) Then
'                StrSql = StrSql & "INSTRUM_PAGO = '" & objRs!instrum_pago & "',"
'            End If
'            If Not EsNulo(objRs!tipo_cuenta) Then
'                StrSql = StrSql & "TIPO_CUENTA = '" & objRs!tipo_cuenta & "',"
'            End If
'            If Not EsNulo(objRs!entid_pago) Then
'                StrSql = StrSql & "ENTID_PAGO = '" & objRs!entid_pago & "',"
'            End If
'            If Not EsNulo(objRs!nro_cuenta) Then
'                StrSql = StrSql & "NRO_CUENTA = '" & objRs!nro_cuenta & "',"
'            End If
'            If Not EsNulo(objRs!nro_cuil) Then
'                StrSql = StrSql & "NRO_CUIL = '" & objRs!nro_cuil & "',"
'            End If
'            If Not EsNulo(objRs!osocial) Then
'                StrSql = StrSql & "OSOCIAL = '" & objRs!osocial & "',"
'            End If
'            If Not EsNulo(objRs!Fecha_baja) Then
'                StrSql = StrSql & "FECHA_BAJA = '" & Format(objRs!Fecha_baja, formatofecha) & "',"
'            End If
'            If Not EsNulo(objRs!motivo_baja) Then
'                StrSql = StrSql & ",MOTIVO_BAJA = '" & objRs!motivo_baja & "',"
'            End If
'            If Not EsNulo(objRs!provincia_trabaja) Then
'                StrSql = StrSql & "PROVINCIA_TRABAJA = '" & objRs!provincia_trabaja & "',"
'            End If
'            If Not EsNulo(objRs!caracter_servicio) Then
'                StrSql = StrSql & "CARACTER_SERVICIO = '" & objRs!caracter_servicio & "',"
'            End If
'            If Not EsNulo(objRs!suc_pago) Then
'                StrSql = StrSql & "SUC_PAGO = '" & objRs!suc_pago & "',"
'            End If
'            If Not EsNulo(objRs!grupo_liquidacion) Then
'                StrSql = StrSql & "GRUPO_LIQUIDACION = '" & objRs!grupo_liquidacion & "',"
'            End If
'            If Not EsNulo(objRs!Fecha_Promo_vac) Then
'                StrSql = StrSql & "FECHA_PROMO_VAC = '" & Format(objRs!Fecha_Promo_vac, formatofecha) & "',"
'            End If
'            If Not EsNulo(objRs!CBU) Then
'                StrSql = StrSql & "CBU = '" & objRs!CBU & "',"
'            End If
'            StrSql = Left(StrSql, Len(StrSql) - 1)
'            StrSql = StrSql & " WHERE legajo = '" & objRs!legajo & "'"
'            StrSql = StrSql & " AND EMPRESA = '" & objRs!empresa & "'"
'            StrSql = StrSql & " AND FECHA_INGRESO = '" & Format(objRs!Fecha_Ingreso, formatofecha) & "'"
'        End If
'
'        objconnBaseLobruno.Execute StrSql, , adExecuteNoRecords
'
'        'Actualizo el progreso
'        Progreso = Progreso + IncPorc
'        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & " WHERE bpronro = " & NroProceso
'        objconnProgreso.Execute StrSql, , adExecuteNoRecords
'
'        objRs.MoveNext
'    Loop
'
'    ' Borro los datos en la tabla temporal
'    StrSql = "DELETE FROM legajo_liquidar"
'    objConn.Execute StrSql, , adExecuteNoRecords
'
'
'MyCommitTrans
'
'Fin:
'    Exit Sub
'
'CE_Comenzarmigracion:
'    HuboErrores = True
''    Resume Next
'    Flog.writeline
'    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
'    Flog.writeline Espacios(Tabulador * 0) & "Error Function ComenzarMigracion. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
'    Flog.writeline Espacios(Tabulador * 0) & "Error: " & Err.Number
'    Flog.writeline Espacios(Tabulador * 0) & "Decripcion: " & Err.Description
'    Flog.writeline
'    If InStr(1, Err.Description, "ODBC") > 0 Then
'        Flog.writeline
'        Flog.writeline Espacios(Tabulador * 0) & "SQL Ejecutado: " & StrSql
'        Flog.writeline
'    End If
'    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
'    Flog.writeline
'    Call InsertarError("0", "GENERAL", "GENERAL", "Error en la copia de datos a la base Lobruno. Ver archivo log InterfaceLobruno-" & NroProceso & ".log", 0, False, "")
'    GoTo Fin
End Sub
