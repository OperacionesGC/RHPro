Attribute VB_Name = "mdlDiasPedidos"
Option Explicit
'------------------------------------------------------------------------------------
'11/11/2013" ' Gonzalez Nicolás - Se movieron comentarios a mdlValidarBD
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

Dim objBTurno As New BuscarTurno
Dim objBDia As New BuscarDia
Dim objFeriado As New Feriado
Dim objFechasHoras As New FechasHoras

Global diatipo As Byte
Global ok As Boolean
'Global fecha_desde As Date 'mdf - estan declarada en las politicas de vac
'Global fecha_hasta As Date 'mdf - estan declarada en las politicas de vac
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
Global Todos_Posibles As Boolean
'---------------------------
'---------------------------
Global diascoract As Integer
Global DiasTom As Integer
Global diascorant As Integer
Global diasdebe As Integer
Global diastot As Integer
Global diasyaped As Integer
Global diaspend As Integer

Global nroTipvac As Long
Global hasta As Date
Global totferiados As Integer
Global tothabiles As Integer
Global totNohabiles As Integer
Global Aux_Fecha_Desde As Date
Global Aux_Fecha_Hasta As Date
Global Vac_Fecha_Desde As Date
Global Aux_Cant_dias As Integer
Global Vac_Cant_dias As Integer

'---------------------------
'---------------------------


Public Sub Main()

Dim Fecha As Date
'Dim Ternro As Long
'Dim fecha_desde As Date
'Dim fecha_hasta As Date
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

Dim objReg As New ADODB.Recordset
Dim strcmdLine As String
Dim objconnMain As New ADODB.Connection
Dim Archivo As String

Dim rs As New ADODB.Recordset
Dim rs_Batch_Proceso As New ADODB.Recordset
Dim rs_His_Batch_Proceso As New ADODB.Recordset

Dim rs_tipovacac As New ADODB.Recordset
Dim rs_vacdiascor As New ADODB.Recordset

Dim rsDias As New ADODB.Recordset
Dim rsVac As New ADODB.Recordset
Dim rs_Periodos_Vac As New ADODB.Recordset
Dim l_TienePolAlcance As Boolean
'---------------------------------
'Dim diascoract As Integer
'Dim DiasTom As Integer
'Dim diascorant As Integer
'Dim diasdebe As Integer
'Dim diastot As Integer
'Dim diasyaped As Integer
'Dim diaspend As Integer
'
'Dim nroTipvac As Long
'Dim hasta As Date
'Dim totferiados As Integer
'Dim tothabiles As Integer
'Dim totNohabiles As Integer
'Dim Aux_Fecha_Desde As Date
'Dim Vac_Fecha_Desde As Date
'Dim Aux_Cant_dias As Integer
'Dim Vac_Cant_dias As Integer
'EAM
Dim dias_tranf_PAct As Integer
Dim estadoPeriodo As Integer

Dim PID As String
Dim ArrParametros

'NG
Dim usuario As String
Dim Texto As String
Dim modeloPais
Dim VersionPais
Dim Continua As Boolean
Dim Usa1515 As Boolean 'NG - (Vacas con pol 1515)


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
    
    ' Creo el archivo de texto del desglose
    Archivo = PathFLog & "Vac_DiasPedidos" & "-" & NroProceso & ".log"
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
    On Error GoTo ce
    
    'FGZ - 05/08/2009 --------- Control de versiones ------
    'Version_Valida = ValidarV(Version, 11, TipoBD)
    
    '*******************************************************************************************************
    '--------------- VALIDO MODELOS SEGUN POLITICA 1515 | PUEDE TENER ALCANCE POR ESTRUCTURAS --------------
    '*******************************************************************************************************
    Version_Valida = ValidaModeloyVersiones(Version, 10)
    If (Version_Valida = False) Then
        'SI NO ESTA ACTIVA LA 1515 O NO EXISTE CONFIGURACIÓN, TOMA DEFAULT
        modeloPais = Pais_Modelo(7)
        Version_Valida = ValidarVBD(Version, 10, TipoBD, modeloPais)
        Usa1515 = False
    Else
        Usa1515 = True
    End If

    
    
    
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
    
    '____________________________________________________________
    'NG - VALIDA QUE ESTE ACTIVO LA TRADUCCION A MULTI IDIOMA
    usuario = rs_Batch_Proceso!iduser
    Call Valida_MultiIdiomaActivo(usuario)
    '------------------------------------------------------------
    '------------------------------------------------------------
    
    Flog.writeline "parametros: " & parametros

    
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
            pos2 = InStr(pos1, parametros, ".") - 1
            fecha_hasta = CDate(Mid(parametros, pos1, pos2 - pos1 + 1))
            
            pos1 = pos2 + 2
            pos2 = InStr(pos1, parametros, ".") - 1
            Aux_Fecha_Desde = CDate(Mid(parametros, pos1, pos2 - pos1 + 1))
            Vac_Fecha_Desde = Aux_Fecha_Desde
            
            pos1 = pos2 + 2
            pos2 = Len(parametros) - 1
            If pos2 >= pos1 Then
                Aux_Cant_dias = Mid(parametros, pos1, pos2 - pos1 + 1)
                Vac_Cant_dias = Aux_Cant_dias
                If Aux_Cant_dias = 0 Then
                    Todos_Posibles = True
                Else
                    Todos_Posibles = False
                End If
            Else
                Todos_Posibles = True
            End If
        End If
    End If
       
    
    Set objFechasHoras.Conexion = objConn
        
    StrSql = " SELECT * FROM batch_empleado " & _
             " WHERE batch_empleado.bpronro = " & NroProceso
    OpenRecordset StrSql, objReg
    
    SinError = True
    HuboErrores = False

    'SI NO SE UTILIZA LA POLITICA 1515
    If Usa1515 = False Then
        StrSql = "SELECT * FROM alcance_testr WHERE tanro= 21"
        OpenRecordset StrSql, rs
        If Not rs.EOF Then
            l_TienePolAlcance = True
        Else
            l_TienePolAlcance = False
        End If
    End If

    
    Do While Not objReg.EOF
        
        Aux_Fecha_Desde = Vac_Fecha_Desde
        Aux_Cant_dias = Vac_Cant_dias
        
        Ternro = objReg!Ternro
        
        'Flog.writeline "Inicio Empleado:" & Ternro
        Flog.writeline ""
        Flog.writeline "========================================================================"
        Flog.writeline EscribeLogMI("Inicio Empleado") & ": " & Ternro
        
        'SI NO UTILIZA LA POLITICA 1515
        If Usa1515 = False Then
            'Obtiene los períodos de vacaciones entre el rango de fechas y dependiendo si tiene configurado
            'el alcance por esturctura (21)
            If l_TienePolAlcance Then
                StrSql = " SELECT DISTINCT vacacion.vacnro, vacdesc, vacfecdesde, vacfechasta, vacacion.vacanio"
                StrSql = StrSql & " FROM  vacacion "
                StrSql = StrSql & " INNER JOIN vac_estr ON vacacion.vacnro= vac_estr.vacnro "
                StrSql = StrSql & " INNER JOIN his_estructura ON vac_estr.estrnro = his_estructura.estrnro "
                StrSql = StrSql & " WHERE  his_estructura.ternro= " & Ternro
                StrSql = StrSql & " AND vacfecdesde <= " & ConvFecha(fecha_hasta)
                StrSql = StrSql & " AND  (vacfechasta >= " & ConvFecha(fecha_desde) & " OR vacfechasta IS NULL)"
                StrSql = StrSql & " ORDER BY vacfecdesde "
            Else
                StrSql = "SELECT vacacion.vacdesc, vacacion.vacnro, vacacion.vacfecdesde, vacacion.vacfechasta, vacacion.vacanio "
                StrSql = StrSql & "FROM  vacacion "
                StrSql = StrSql & " WHERE vacfecdesde <= " & ConvFecha(fecha_hasta)
                StrSql = StrSql & " AND  (vacfechasta >= " & ConvFecha(fecha_desde) & " OR vacfechasta IS NULL)"
                StrSql = StrSql & " ORDER BY vacfecdesde "
            End If
            
        Else  'SI SE UTILIZA LA POLITICA 1515
            Call Politica(1515)
            If Not PoliticaOK Then
                Flog.writeline ""
                Texto = EscribeLogMI("No se puede procesar al empleado.") & " "
                Texto = Texto & Replace(EscribeLogMI("Revisar configuración de Política @@NUM@@"), "@@NUM@@", "1515")
                Flog.writeline Texto
                
                Flog.writeline "************************************************************************"
                Flog.writeline "************************************************************************"
                StrSql = "Select * from vacacion where 1=3" 'mdf
                OpenRecordset StrSql, rs_Periodos_Vac    'mdf
                GoTo siguiente
            End If
            'modeloPais = st_Opcion
            modeloPais = st_ModeloPais
            VersionPais = st_Opcion
            'Flog.writeline " modeloPais: " & modeloPais
            'Flog.writeline " VersionPais: " & VersionPais
            
            If alcannivel = 1 Then 'Individual
                Select Case modeloPais
                    '************************* PARAGUAY *****************************
                    Case 6: 'Paraguay
                        Select Case VersionPais
                            Case 0: 'Standard Paraguay
                                'Texto = Replace(Texto, "@@TXT@@", EscribeLogMI("Argentina"))
                                'Flog.writeline Texto & " " & modeloPais
                                '---------------------------------------------------------
                                StrSql = "SELECT vacacion.vacdesc, vacacion.vacnro, vac_alcan.vacfecdesde, vac_alcan.vacfechasta, vacacion.vacanio,vacacion.alcannivel "
                                StrSql = StrSql & " FROM vacacion "
                                StrSql = StrSql & " INNER JOIN vac_alcan ON vac_alcan.vacnro = vacacion.vacnro"
                                StrSql = StrSql & " WHERE vac_alcan.vacfecdesde < " & ConvFecha(fecha_hasta)
                                StrSql = StrSql & " AND  (vac_alcan.vacfechasta < " & ConvFecha(fecha_desde) & " OR vac_alcan.vacfechasta IS NULL)"
                        End Select
                End Select
                StrSql = StrSql & " AND vacacion.alcannivel=1 AND vac_alcan.origen = " & Ternro
                StrSql = StrSql & " ORDER BY vac_alcan.vacfecdesde"
            ElseIf alcannivel = 2 Then 'Estructura
                'BUSCA PERIODOS INDIVIDUALES (Configurados por Estructura)
                StrSql = "SELECT vacacion.vacdesc, vacacion.vacnro, vacacion.vacfecdesde, vacacion.vacfechasta, vacacion.vacanio,vacacion.alcannivel "
                StrSql = StrSql & " FROM vacacion "
                StrSql = StrSql & " INNER JOIN vac_alcan ON vac_alcan.vacnro = vacacion.vacnro"
                StrSql = StrSql & " INNER JOIN his_estructura H1 ON H1.estrnro = vac_alcan.origen and H1.ternro = " & Ternro
                StrSql = StrSql & " WHERE "
                StrSql = StrSql & " vacacion.alcannivel = 2 And vac_alcan.alcannivel = 2"
            Else 'Global
                StrSql = "SELECT vacacion.vacdesc, vacacion.vacnro, vacacion.vacfecdesde, vacacion.vacfechasta, vacacion.vacanio,vacacion.alcannivel "
                StrSql = StrSql & " FROM vacacion "
                StrSql = StrSql & " WHERE vacfecdesde <= " & ConvFecha(fecha_hasta)
                StrSql = StrSql & " AND  (vacfechasta >= " & ConvFecha(fecha_desde) & " OR vacfechasta IS NULL)"
                StrSql = StrSql & " AND vacacion.alcannivel=" & alcannivel
                StrSql = StrSql & " ORDER BY vacfecdesde "
            End If
            '---------------------------------------------------------------------------------------------------------
        End If
       
        Flog.writeline
        Flog.writeline "------------------------------------------------"
        Flog.writeline StrSql
        Flog.writeline "------------------------------------------------"
        Flog.writeline

       
        'FGZ - 14/04/2010 ----------------------------------------------------------------------
        'StrSql = "SELECT * FROM vacacion "
        'StrSql = StrSql & " WHERE vacfecdesde <= " & ConvFecha(fecha_hasta)
        'StrSql = StrSql & " AND  vacfechasta >= " & ConvFecha(fecha_desde)
        'StrSql = StrSql & " ORDER BY vacnro"
        
        'FGZ - 14/04/2010 ----------------------------------------------------------------------
        OpenRecordset StrSql, rs_Periodos_Vac
        'Seteo el incremento del progreso
        CEmpleadosAProc = objReg.RecordCount
        If CEmpleadosAProc = 0 Then
            CEmpleadosAProc = 1
        End If
        CDiasAProc = rs_Periodos_Vac.RecordCount
        If CDiasAProc = 0 Then
            CDiasAProc = 1
        End If
        IncPorc = ((100 / CEmpleadosAProc) * (100 / CDiasAProc)) / 100
        
       'FGZ - 24/06/2009 -------
        Diashabiles_LV = False
        PoliticaOK = False
        Call Politica(1510)
        'FGZ - 24/06/2009 -------

        
        
        'Aux_Fecha_Desde = fecha_desde
        Do While (Not rs_Periodos_Vac.EOF)
            Continua = False
            MyBeginTrans
            Flog.writeline EscribeLogMI("Periodo") & ": " & rs_Periodos_Vac!vacfecdesde & " - " & rs_Periodos_Vac!vacfechasta
                        
            'EAM- Obtiene el estado del periodo 05-10-2010
            StrSql = "SELECT * FROM vacdiascor WHERE ternro = " & Ternro & " AND vacnro = " & rs_Periodos_Vac!vacnro & " AND (venc = 1)"
            OpenRecordset StrSql, rs_vacdiascor
            
            'Estados (0 Cerrado | -1 abierto)
            estadoPeriodo = 0
            If rs_vacdiascor.EOF Then
                estadoPeriodo = -1
            End If
            rs_vacdiascor.Close
            
            'Toma el primer período que se encuentre abierto
            If (estadoPeriodo = -1) Then
            
                   Select Case modeloPais
                       Case 6:
                         '************************* PARAGUAY *****************************
                         Select Case VersionPais
                             Case 0: 'Standard Paraguay
                                 'Texto = Replace(Texto, "@@TXT@@", EscribeLogMI("Paraguay"))
                                 'Flog.Writeline Texto & " " & modeloPais
                                 Flog.writeline "PARAGUAY"
                                 '---------------------------------------------------------
                                 'Valido que la fecha este dentro del período de vacaciones
                                 If (Aux_Fecha_Desde >= rs_Periodos_Vac!vacfechasta) And (Aux_Fecha_Desde <= DateAdd("m", 6, rs_Periodos_Vac!vacfechasta)) Then
                                    Continua = True
                                    Aux_Fecha_Hasta = DateAdd("m", 6, rs_Periodos_Vac!vacfechasta)
                                    Call GeneraPedido_PY(Aux_Fecha_Desde, rs_Periodos_Vac!vacnro, rs_Periodos_Vac!vacdesc, alcannivel, Reproceso)
                                 End If
                                 
                             End Select
                         Case Else 'Configuración estándard
                            'Texto = Replace(Texto, "@@TXT@@", EscribeLogMI("Argentina"))
                            'Flog.Writeline Texto & " " & modeloPais
                            '---------------------------------------------------------
                            'Valido que la fecha este dentro del período de vacaciones
                            If (Aux_Fecha_Desde >= rs_Periodos_Vac!vacfecdesde) And (Aux_Fecha_Desde <= rs_Periodos_Vac!vacfechasta) Then
                                Continua = True
                                Call GeneraPedido_ARG(Aux_Fecha_Desde, rs_Periodos_Vac!vacnro, rs_Periodos_Vac!vacdesc, alcannivel, Reproceso)
                            End If

                   End Select
                   If Continua = False Then
                    Flog.writeline "La fecha en la que se va a generar los dias pedidos estan fuera del rango de fechas del periodo " & Aux_Fecha_Desde
                   End If
                Else
                    Flog.writeline "El perido " & rs_Periodos_Vac!vacdesc & " (" & rs_Periodos_Vac!vacnro & ") se encuentra cerrado"
                End If
               'si la fecha en la que se va a generar los dias pedidos estan fuera del rengo de fechas del periodo
               'no se procesan
'                If (Aux_Fecha_Desde >= rs_Periodos_Vac!vacfecdesde) And (Aux_Fecha_Desde <= rs_Periodos_Vac!vacfechasta) Then
'                    diascoract = 0
'                    DiasTom = 0
'                    diascorant = 0
'                    diasdebe = 0
'                    diastot = 0
'                    diasyaped = 0
'                    diaspend = 0
'
'                    Flog.Writeline "Periodo de Vacaciones:" & rs_Periodos_Vac!vacnro & " " & rs_Periodos_Vac!vacdesc
'
'                    NroVac = rs_Periodos_Vac!vacnro
'
'                    'FGZ - 23/10/2009 - Se cambió esto ---------------------------------------------------------
'                    '               Ahora debe buscar los dias correspondientes - vencidos + transferidos
'
'                    'StrSql = "SELECT * FROM vacdiascor WHERE ternro = " & Ternro & " AND vacnro = " & NroVac
'                    'OpenRecordset StrSql, rs_vacdiascor
'                    'If Not rs_vacdiascor.EOF Then
'                    '
'                    '    StrSql = "SELECT * FROM tipovacac WHERE tipvacnro = " & rs_vacdiascor!tipvacnro
'                    '    OpenRecordset StrSql, rs_tipovacac
'                    '    If Not rs_tipovacac.EOF Then
'                    '        nroTipvac = rs_tipovacac!tipvacnro
'                    '    End If
'                    '    diascoract = rs_vacdiascor!vdiascorcant ' dias corresp al periodo actual
'                    '
'                    '    StrSql = "SELECT * FROM vacacion WHERE vacnro <> " & NroVac & " and vacfechasta < " & ConvFecha(fecha_desde)
'                    '    OpenRecordset StrSql, rsVac
'                    '    Do While Not rsVac.EOF
'                    '        DiasTom = 0
'                    '
'                    '        StrSql = "SELECT * FROM lic_vacacion " & _
'                    '                 " INNER JOIN emp_lic ON emp_lic.emp_licnro = lic_vacacion.emp_licnro " & _
'                    '                 " WHERE lic_vacacion.vacnro = " & rsVac!vacnro & " AND emp_lic.empleado = " & Ternro
'                    '        OpenRecordset StrSql, rsDias
'                    '        Do While Not rsDias.EOF
'                    '            DiasTom = DiasTom + rsDias!elcantdias
'                    '            rsDias.MoveNext
'                    '        Loop
'                    '        diascorant = 0
'                    '        StrSql = "SELECT * FROM vacdiascor WHERE ternro = " & Ternro & " AND vacnro = " & rsVac!vacnro
'                    '        OpenRecordset StrSql, rs
'                    '        If Not rs.EOF Then diascorant = rs!vdiascorcant
'                    '        diasdebe = diasdebe + (diascorant - DiasTom)
'                    '
'                    '        rsVac.MoveNext
'                    '    Loop
'                    '    diastot = diascoract + diasdebe
'                    'End If
'
'                    'EAM- Obtiene los días correspondientes
'                    StrSql = "SELECT * FROM vacdiascor WHERE ternro = " & Ternro & " AND vacnro = " & NroVac
'                    StrSql = StrSql & " AND (venc = 0 OR venc IS NULL)"
'                    OpenRecordset StrSql, rs_vacdiascor
'                    If Not rs_vacdiascor.EOF Then
'                        diascoract = rs_vacdiascor!vdiascorcant ' dias corresp al periodo actual
'                        nroTipvac = rs_vacdiascor!tipvacnro
'                    Else
'                        diascoract = 0
'                    End If
'
'                    'Resto los vencidos
'                    StrSql = "SELECT * FROM vacdiascor WHERE ternro = " & Ternro & " AND vacnro = " & NroVac
'                    StrSql = StrSql & " AND (venc = 1)"
'                    OpenRecordset StrSql, rs_vacdiascor
'                    If Not rs_vacdiascor.EOF Then
'                        diascoract = diascoract - rs_vacdiascor!vdiascorcant
'                    End If
'
'                    'Sumo los transferidos
'                    StrSql = "SELECT * FROM vacdiascor WHERE ternro = " & Ternro & " AND vacnro = " & NroVac
'                    StrSql = StrSql & " AND (venc = 2)"
'                    OpenRecordset StrSql, rs_vacdiascor
'                    'EAM- Dias tranferidos al periodo actual
'                    dias_tranf_PAct = 0
'                    DiasTom = 0
'
'                    If Not rs_vacdiascor.EOF Then
'                        diascoract = diascoract + rs_vacdiascor!vdiascorcant
'                        dias_tranf_PAct = rs_vacdiascor!vdiascorcant
'                    End If
'
'
'                    If diascoract > 0 Then
'                        'StrSql = "SELECT * FROM vacacion WHERE vacnro <> " & NroVac & " and vacfechasta < " & ConvFecha(fecha_desde)
'                        'EAM- Obtiene todos los periodos abiertos para el empleado en orden desc.
'                        StrSql = "SELECT DISTINCT vacacion.vacnro, vacdesc, vacfecdesde, vacfechasta,vacanio " & _
'                                " FROM vacacion " & _
'                                " INNER JOIN vac_estr ON vacacion.vacnro= vac_estr.vacnro" & _
'                                " INNER JOIN vacdiascor ON vac_estr.vacnro = vacdiascor.vacnro" & _
'                                " INNER JOIN his_estructura ON vac_estr.estrnro = his_estructura.estrnro " & _
'                                " WHERE his_estructura.ternro= " & Ternro & " AND vacacion.vacestado= -1 AND " & _
'                                " vacacion.vacnro <> " & NroVac & " and vacfechasta < " & ConvFecha(fecha_desde) & _
'                                " AND (venc = 1) " & _
'                                " ORDER BY vacanio DESC "
'                        OpenRecordset StrSql, rsVac
'                        Do While Not rsVac.EOF
'                            DiasTom = 0
'
'                            StrSql = "SELECT * FROM lic_vacacion " & _
'                                     " INNER JOIN emp_lic ON emp_lic.emp_licnro = lic_vacacion.emp_licnro " & _
'                                     " WHERE lic_vacacion.vacnro = " & rsVac!vacnro & " AND emp_lic.empleado = " & Ternro
'                            OpenRecordset StrSql, rsDias
'                            Do While Not rsDias.EOF
'                                DiasTom = DiasTom + rsDias!elcantdias
'                                rsDias.MoveNext
'                            Loop
'
'                            'Busco los correspondientes al periodo
'                            If dias_tranf_PAct <> 0 Then
'                                diascorant = (dias_tranf_PAct * (-1))
'                                dias_tranf_PAct = 0
'                            Else
'                                diascorant = 0
'                            End If
'
'                            'EAM- Obtine los días correspondientes del periodo
'                            StrSql = "SELECT * FROM vacdiascor WHERE ternro = " & Ternro & " AND vacnro = " & rsVac!vacnro
'                            StrSql = StrSql & " AND (venc = 0 OR venc IS NULL)"
'                            OpenRecordset StrSql, rs
'                            If Not rs.EOF Then
'                                diascorant = diascorant + rs!vdiascorcant
'
'                                'resto los vencidos
'                                StrSql = "SELECT * FROM vacdiascor WHERE ternro = " & Ternro & " AND vacnro = " & rsVac!vacnro
'                                StrSql = StrSql & " AND (venc = 1)"
'                                OpenRecordset StrSql, rs
'                                If Not rs.EOF Then
'                                    diascorant = diascorant - rs!vdiascorcant
'                                End If
'
'                                'sumo los transferidos
'                                StrSql = "SELECT * FROM vacdiascor WHERE ternro = " & Ternro & " AND vacnro = " & rsVac!vacnro
'                                StrSql = StrSql & " AND (venc = 2)"
'                                OpenRecordset StrSql, rs
'                                If Not rs.EOF Then
'                                    'diascorant = diascorant + rs!vdiascorcant
'                                    dias_tranf_PAct = rs!vdiascorcant
'                                End If
'                            Else
'                                diascorant = 0
'                            End If
'
'
'                            diasdebe = diasdebe + (diascorant - DiasTom)
'
'                            rsVac.MoveNext
'                        Loop
'                        diastot = diascoract + diasdebe
'                    End If
'
'
'                    If Not Reproceso Then
'                        'Busco los pedidos de ese periodo
'                        StrSql = "SELECT * FROM vacdiasped WHERE ternro = " & Ternro & " AND vacnro = " & NroVac
'                        OpenRecordset StrSql, objRs
'                        Do While Not objRs.EOF
'                            'diasyaped = diasyaped + objRs!vdiapedcant
'                            diasyaped = diasyaped + objRs!vdiaspedhabiles
'                            Aux_Fecha_Desde = IIf(Aux_Fecha_Desde < (objRs!vdiapedhasta + 1), objRs!vdiapedhasta + 1, Aux_Fecha_Desde)
'                            objRs.MoveNext
'                        Loop
'                    Else
'                        'borro los que estan en el rango de fechas
'                        StrSql = "DELETE FROM vacdiasped WHERE ternro = " & Ternro & " AND vacnro = " & NroVac
'                        StrSql = StrSql & " AND vdiapeddesde >= " & ConvFecha(fecha_desde)
'                        objConn.Execute StrSql, , adExecuteNoRecords
'                        Flog.Writeline "Se Borraron por reprocesamiento los días pedidos del período " & NroVac & " >= a la fecha " & fecha_desde
'
'                        ' Busco los pedidos de ese periodo que quedaron afuera del rango de fechas
'                        StrSql = "SELECT * FROM vacdiasped WHERE ternro = " & Ternro & " AND vacnro = " & NroVac
'                        OpenRecordset StrSql, objRs
'                        Do While Not objRs.EOF
'                            'diasyaped = diasyaped + objRs!vdiapedcant
'                            diasyaped = diasyaped + objRs!vdiaspedhabiles
'                            'Aux_Fecha_Desde = objRs!vdiapedhasta + 1
'                            objRs.MoveNext
'                        Loop
'
'                    End If
'
'                    diaspend = diastot - diasyaped
'                    If diaspend > 0 Then
'                        If Todos_Posibles Then
'                            Call DiasPedidos(nroTipvac, Aux_Fecha_Desde, hasta, Ternro, diaspend, tothabiles, totNohabiles, totferiados)
'
'                            'Verificar Fase
'                            If activo(Ternro, Aux_Fecha_Desde, hasta) Then
'                                StrSql = "INSERT INTO vacdiasped (vdiapedhasta,vdiaspedferiados,vdiaspedhabiles,vdiaspednohabiles,ternro,vacnro,vdiapedcant,vdiapeddesde,vdiaspedestado) VALUES (" & _
'                                          ConvFecha(hasta) & "," & totferiados & "," & tothabiles & "," & totNohabiles & "," & Ternro & "," & NroVac & "," & (totferiados + tothabiles + totNohabiles) & "," & ConvFecha(Aux_Fecha_Desde) & ",-1)"
'                                objConn.Execute StrSql, , adExecuteNoRecords
'                            Else
'                                Flog.Writeline "No se insertaron los días " & Aux_Fecha_Desde & " a " & hasta & " porque se superpone con un período inactivo del empleado."
'                            End If
'
'                            Aux_Fecha_Desde = hasta + 1
'                        Else
'                             If Aux_Cant_dias > 0 Then
'                                If diaspend >= Aux_Cant_dias Then
'                                    diaspend = Aux_Cant_dias
'                                    Aux_Cant_dias = 0
'                                Else
'                                    Aux_Cant_dias = Aux_Cant_dias - diaspend
'                                End If
'                                'Call DiasPedidos(nroTipvac, fecha_desde, hasta, Ternro, diaspend, tothabiles, totNohabiles, totferiados)
'                                Call DiasPedidos(nroTipvac, Aux_Fecha_Desde, hasta, Ternro, diaspend, tothabiles, totNohabiles, totferiados)
'
'                                'Verificar Fase
'                                If activo(Ternro, Aux_Fecha_Desde, hasta) Then
'                                    StrSql = "INSERT INTO vacdiasped (vdiapedhasta,vdiaspedferiados,vdiaspedhabiles,vdiaspednohabiles,ternro,vacnro,vdiapedcant,vdiapeddesde,vdiaspedestado) VALUES (" & _
'                                              ConvFecha(hasta) & "," & totferiados & "," & tothabiles & "," & totNohabiles & "," & Ternro & "," & NroVac & "," & (totferiados + tothabiles + totNohabiles) & "," & ConvFecha(Aux_Fecha_Desde) & ",-1)"
'                                    objConn.Execute StrSql, , adExecuteNoRecords
'                                Else
'                                    Flog.Writeline "No se insertaron los días " & Aux_Fecha_Desde & " a " & hasta & " porque se superpone con un período inactivo del empleado."
'                                End If
'
'                                Aux_Fecha_Desde = hasta + 1
'                            End If
'                        End If
'                    End If
'                If Continua = False Then
'                    Flog.Writeline "La fecha en la que se va a generar los dias pedidos estan fuera del rango de fechas del periodo " & Aux_Fecha_Desde
'                End If
'            Else
'                Flog.Writeline "El perido " & rs_Periodos_Vac!vacdesc & " (" & rs_Periodos_Vac!vacnro & ") se encuentra cerrado"
'            End If
                
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
        
        
        'FGZ - 27/09
            If Not rs_Periodos_Vac.EOF Then
              rs_Periodos_Vac.MoveNext
            End If
        Loop
         If Not objReg.EOF Then
          objReg.MoveNext
         End If
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
    
    
ce:
    MyRollbackTrans
    HuboErrores = True
    SinError = False
    
    Flog.writeline " ------------------------------------------------------------"
    Flog.writeline "Error procesando Empleado:" & Ternro & " " & Fecha
    Flog.writeline Err.Description
    Flog.writeline "SQL: " & StrSql
    Flog.writeline " ------------------------------------------------------------"
    StrSql = "Select * from vacacion where 1=3" 'mdf
    OpenRecordset StrSql, rs_Periodos_Vac    'mdf
    StrSql = "Select * from vacacion where 1=3" 'mdf
    OpenRecordset StrSql, objReg
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

Function activo(ByVal Tercero As Long, ByVal desde As Date, ByVal hasta As Date) As Boolean
Dim sql As String
Dim salida As Boolean
Dim rs_actv As New ADODB.Recordset

    'Busco una fase que contenga completamente el rango y este activa
    sql = " SELECT * "
    sql = sql & " FROM fases "
    sql = sql & " WHERE fases.empleado = " & Tercero
    sql = sql & " AND ( (altfec <= " & ConvFecha(desde) & ") AND ((bajfec >= " & ConvFecha(hasta) & ") OR (bajfec is null)) ) "
    sql = sql & " AND fases.estado = -1  "
    OpenRecordset sql, rs_actv
    
    If rs_actv.EOF Then
        salida = False
    Else
        salida = True
    End If
    rs_actv.Close
    
    activo = salida

End Function

'Private Sub DiasPedidos2(TipoVac As Long, FechaInicial As Date, FechaFinal As Date, Ternro As Long, ByRef CDias As Integer, ByRef cHabiles As Integer, ByRef cFeriados As Integer)
''Descripcion: Calcula la cantidad de días (LU-VI) entre dos fechas.
'
'Dim Fecha As Date
'Dim objFeriado As New Feriado
'Dim DHabiles(1 To 7) As Boolean
'Dim esFeriado As Boolean
'
'    CDias = 0
'    cHabiles = 0
'    cFeriados = 0
'
'    Fecha = FechaInicial
'
'    Set objFeriado.Conexion = objConn
'    Set objFeriado.ConexionTraza = objConn
'
'    StrSql = "SELECT * FROM tipovacac WHERE tpvnrocol = " & TipoVac
'    OpenRecordset StrSql, objRs
'    If objRs.EOF Then Exit Sub
'
'    DHabiles(1) = objRs!tpvhabiles__1
'    DHabiles(2) = objRs!tpvhabiles__2
'    DHabiles(3) = objRs!tpvhabiles__3
'    DHabiles(4) = objRs!tpvhabiles__4
'    DHabiles(5) = objRs!tpvhabiles__5
'    DHabiles(6) = objRs!tpvhabiles__6
'    DHabiles(7) = objRs!tpvhabiles__7
'
'
'    Do While Fecha <= FechaFinal
'
'        esFeriado = objFeriado.Feriado(Fecha, Ternro, False)
'
'        If (esFeriado) And (objRs!tpvferiado = 0) Then
'            cFeriados = cFeriados + 1
'        Else
'            If DHabiles(Weekday(Fecha)) Or ((esFeriado) And (objRs!tpvferiado = -1)) Then
'                CDias = CDias + 1
'            Else
'                cHabiles = cHabiles + 1
'            End If
'        End If
'
'        Fecha = DateAdd("d", 1, Fecha)
'    Loop
'
'    Set objFeriado = Nothing
'
'End Sub

'Private Sub DiasPedidos1(TipoVac As Long, FechaInicial As Date, ByRef Fecha As Date, Ternro As Long, cant As Integer, ByRef cHabiles As Integer, ByRef cFeriados As Integer)
''Calcula la fecha hasta a partir de la fecha desde, la cantidad de dias pedidos y el tipo
''de vacacion asociado a los dias correspòndientes, para el período
'
'Dim i As Integer
'Dim j As Integer
'Dim objFeriado As New Feriado
'Dim DHabiles(1 To 7) As Boolean
'Dim esFeriado As Boolean
'Dim objRs As New ADODB.Recordset
'
'    i = 0
'    j = 0
'    cHabiles = 0
'    cFeriados = 0
'
'    Fecha = FechaInicial
'
'    Set objFeriado.Conexion = objConn
'    Set objFeriado.ConexionTraza = objConn
'
'    StrSql = "SELECT * FROM tipovacac WHERE tpvnrocol = " & TipoVac
'    OpenRecordset StrSql, objRs
'    If objRs.EOF Then Exit Sub
'
'    DHabiles(1) = objRs!tpvhabiles__1
'    DHabiles(2) = objRs!tpvhabiles__2
'    DHabiles(3) = objRs!tpvhabiles__3
'    DHabiles(4) = objRs!tpvhabiles__4
'    DHabiles(5) = objRs!tpvhabiles__5
'    DHabiles(6) = objRs!tpvhabiles__6
'    DHabiles(7) = objRs!tpvhabiles__7
'
'
'    Do While i <= cant
'
'        esFeriado = objFeriado.Feriado(Fecha, Ternro, False)
'
'        If (esFeriado) And (objRs!tpvferiado = 0) Then
'            cFeriados = cFeriados + 1
'        Else
'            If DHabiles(Weekday(Fecha)) Or ((esFeriado) And (objRs!tpvferiado = -1)) Then
'                i = i + 1
'            Else
'                cHabiles = cHabiles + 1
'            End If
'        End If
'        If i < cant Then Fecha = DateAdd("d", 1, Fecha)
'    Loop
'
'    Set objFeriado = Nothing
'
'End Sub
'
'
'Private Sub DiasPedidos3(TipoVac As Long, FechaInicial As Date, FechaFinal As Date, Ternro As Long, ByRef CDias As Integer, ByRef cHabiles As Integer, ByRef cFeriados As Integer)
''Descripcion: Calcula la cantidad de días (LU-Sab) entre dos fechas.
'
'Dim Fecha As Date
'Dim objFeriado As New Feriado
'Dim DHabiles(1 To 7) As Boolean
'Dim esFeriado As Boolean
'
'    CDias = 0
'    cHabiles = 0
'    cFeriados = 0
'
'    Fecha = FechaInicial
'
'    Set objFeriado.Conexion = objConn
'    Set objFeriado.ConexionTraza = objConn
'
'    StrSql = "SELECT * FROM tipovacac WHERE tpvnrocol = " & TipoVac
'    OpenRecordset StrSql, objRs
'    If objRs.EOF Then Exit Sub
'
'    DHabiles(1) = objRs!tpvhabiles__1
'    DHabiles(2) = objRs!tpvhabiles__2
'    DHabiles(3) = objRs!tpvhabiles__3
'    DHabiles(4) = objRs!tpvhabiles__4
'    DHabiles(5) = objRs!tpvhabiles__5
'    DHabiles(6) = objRs!tpvhabiles__6
'    DHabiles(7) = objRs!tpvhabiles__7
'
'
'    Do While Fecha <= FechaFinal
'
'        esFeriado = objFeriado.Feriado(Fecha, Ternro, False)
'
'        If (esFeriado) And (objRs!tpvferiado = 0) Then
'            cFeriados = cFeriados + 1
'        Else
'            If (Weekday(Fecha) = 7) Or DHabiles(Weekday(Fecha)) Or ((esFeriado) And (objRs!tpvferiado = -1)) Then
'                CDias = CDias + 1
'            Else
'                cHabiles = cHabiles + 1
'            End If
'        End If
'
'        Fecha = DateAdd("d", 1, Fecha)
'    Loop
'
'    Set objFeriado = Nothing
'
'End Sub


Public Sub DiasPedidos_STD(ByVal tipoVac As Long, ByVal FechaInicial As Date, ByRef Fecha As Date, ByVal Ternro As Long, ByRef Cant As Integer, ByRef CHabiles As Integer, ByRef cNoHabiles As Integer, ByRef cFeriados As Integer)
'Calcula la fecha hasta a partir de la fecha desde, la cantidad de dias pedidos y el tipo
'de vacacion asociado a los dias correspòndientes, para el período
'se cambio de private a public y se cambio el nombre de DiasPedidos a DiasPedidos_STD
Dim i As Integer
Dim j As Integer
Dim objFeriado As New Feriado
Dim DHabiles(1 To 7) As Boolean
Dim EsFeriado As Boolean
Dim objRs As New ADODB.Recordset
Dim ExcluyeFeriados As Boolean


    StrSql = "SELECT * FROM tipovacac WHERE tipvacnro = " & tipoVac
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        DHabiles(1) = objRs!tpvhabiles__1
        DHabiles(2) = objRs!tpvhabiles__2
        DHabiles(3) = objRs!tpvhabiles__3
        DHabiles(4) = objRs!tpvhabiles__4
        DHabiles(5) = objRs!tpvhabiles__5
        DHabiles(6) = objRs!tpvhabiles__6
        DHabiles(7) = objRs!tpvhabiles__7
    
        ExcluyeFeriados = CBool(objRs!tpvferiado)
    Else
        Flog.writeline "No se encontro el tipo de Vacacion " & tipoVac
        Exit Sub
    End If

    
    Set objFeriado.Conexion = objConn
    Set objFeriado.ConexionTraza = objConn
    
    i = 0
    j = 0
    CHabiles = 0
    cNoHabiles = 0
    cFeriados = 0
    
    Fecha = FechaInicial
    
    Do While i <= Cant
    
        EsFeriado = objFeriado.Feriado(Fecha, Ternro, False)
        
        If (EsFeriado) And Not ExcluyeFeriados Then
            cFeriados = cFeriados + 1
            'FGZ - 24/06/2009 -----------------------
            If PoliticaOK And Diashabiles_LV Then
                i = i + 1
            End If
            'FGZ - 24/06/2009 -----------------------
        Else
            If DHabiles(Weekday(Fecha)) Or (EsFeriado And ExcluyeFeriados) Then
                i = i + 1
                If DHabiles(Weekday(Fecha)) Then
                    CHabiles = CHabiles + 1
                End If
            Else
                cNoHabiles = cNoHabiles + 1
                'FGZ - 24/06/2009 -----------------------
                If PoliticaOK And Diashabiles_LV Then
                    i = i + 1
                End If
                'FGZ - 24/06/2009 -----------------------
            End If
'            If DHabiles(Weekday(Fecha)) Or ((esFeriado) And ExcluyeFeriados) Then
'                i = i + 1
'            Else
'                cHabiles = cHabiles + 1
'            End If
        End If
        
        If i < Cant Then
            Fecha = DateAdd("d", 1, Fecha)
        Else
            i = i + 1
        End If
    Loop
    
    Set objFeriado = Nothing

End Sub
