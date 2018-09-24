Attribute VB_Name = "mdlLicencias"
Option Explicit
'------------------------------------------------------------------------------------
' 11/10/2013 - Gonzalez Nicolás - Se movieron versiones a mdlValidarBD
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
Global fecha_desde As Date
Global fecha_hasta As Date
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

Public Sub Main()

Dim Fecha As Date
'Dim Ternro As Long
'Dim fecha_desde As Date
'Dim fecha_hasta As Date
Dim NroVac As Long
'Dim Reproceso As Boolean
Dim parametros As String
Dim cantdias As Integer
Dim Columna As Integer
Dim Mensaje As String
Dim Genera As Boolean
Dim NroTPV As String
Dim Autorizadas As Boolean
Dim vacanio As String
Dim l_TienePolAlcance As Boolean

Dim pos1 As Integer
Dim pos2 As Integer

Dim objReg As New ADODB.Recordset
Dim strCmdLine As String
Dim objconnMain As New ADODB.Connection
Dim Archivo As String

Dim rs As New ADODB.Recordset
Dim rs_Vac As New ADODB.Recordset
Dim rs_Periodos_Vac As New ADODB.Recordset
Dim rs_Batch_Proceso As New ADODB.Recordset
Dim rs_His_Batch_Proceso As New ADODB.Recordset
Dim PID As String
Dim ArrParametros
Dim modeloPais As Integer

'NG
Dim usuario As String
Dim Texto As String
Dim VersionPais
Dim Continua As Boolean
Dim Usa1515 As Boolean 'NG - (Vacas con pol 1515)


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
    Archivo = PathFLog & "Vac_DiasLicencias" & "-" & NroProceso & ".log"
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
    
    'Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaVersion
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-------------------------------------------------"
    Flog.writeline
    
    'EAM - 08/09/2010 --------- Control de versiones ------
    ' 05/07/2011 - Se agrego un parametro para determinar el país con el que va a trabajar el modelo de vacacion
    'El parametro 7 hace referencia al modelo de vacaciones configurado en la tabla confper
    
    '*******************************************************************************************************
    '--------------- VALIDO MODELOS SEGUN POLITICA 1515 | PUEDE TENER ALCANCE POR ESTRUCTURAS --------------
    '*******************************************************************************************************
    Version_Valida = ValidaModeloyVersiones(Version, 12)
    If (Version_Valida = False) Then
        'SI NO ESTA ACTIVA LA 1515 O NO EXISTE CONFIGURACIÓN, TOMA DEFAULT
        modeloPais = Pais_Modelo(7)
        Version_Valida = ValidarVBD(Version, 12, TipoBD, modeloPais)
        Usa1515 = False
    Else
        Usa1515 = True
    End If
    
    'Version_Valida = ValidarV(Version, 12, TipoBD)
    If Not Version_Valida Then
        'Actualizo el progreso
        MyBeginTrans
            StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Error de Version', bprcpid = " & PID & " WHERE bpronro = " & NroProceso
            objConnProgreso.Execute StrSql, , adExecuteNoRecords
        MyCommitTrans
        Flog.writeline
        GoTo Fin
    End If
    'EAM - 08/09/2010 --------- Control de versiones ------
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprcpid = " & PID & ", bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objConnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Levanta Proceso y Setea Parámetros:  " & " " & Now
       
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro =" & NroProceso
    OpenRecordset StrSql, rs_Batch_Proceso
       
    If rs_Batch_Proceso.EOF Then
        Flog.writeline "Error obteniendo los datos del proceso. Posible problema de conexion."
        Exit Sub
    End If
   
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
            
            '------------------------------------------------------
            If Usa1515 = True Then
                pos1 = pos2 + 2
                pos2 = InStr(pos1, parametros, ".") - 1
                vacanio = Mid(parametros, pos1, pos2 - pos1 + 1)
            End If
            '------------------------------------------------------
            
            pos1 = pos2 + 2
            pos2 = InStr(pos1, parametros, ".")
            
            If pos2 = 0 Then
                'pos1 = pos2 + 2
                pos2 = Len(parametros)
                Autorizadas = CBool(Mid(parametros, pos1, pos2))
            
            Else
                Autorizadas = CBool(Mid(parametros, pos1, pos2 - pos1))
                
                'Usuario que dispara el proceso
                User_Proceso = IIf(Not EsNulo(rs_Batch_Proceso!iduser), rs_Batch_Proceso!iduser, "NULL")
                
                pos1 = pos2 + 2
                pos2 = Len(parametros)
                Firma_User_Destino = Mid(parametros, pos1, pos2 - pos1 + 1)
            End If
            'If (Len(parametros) > (pos2 + 1)) Then
                
                
            'Else
            'End If
        End If
    End If
    Set objFechasHoras.Conexion = objConn
        
    StrSql = " SELECT * FROM batch_empleado " & _
             " WHERE batch_empleado.bpronro = " & NroProceso
    OpenRecordset StrSql, objReg
    
    If objReg.EOF Then
        Flog.writeline "No se encontró ningun empleado para procesar "
    End If
    
    SinError = True
    HuboErrores = False
    
    If Usa1515 = True Then
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
    
    Select Case modeloPais
            Case 0: 'Argentina '--------------------------------------------------------------------------------------------------------------------
                Flog.writeline "Modelo de vacaciones de Argentina nro." & modeloPais
                
                StrSql = "SELECT * FROM alcance_testr WHERE tanro= 21"
                OpenRecordset StrSql, rs
                If Not rs.EOF Then
                    l_TienePolAlcance = True
                Else
                    l_TienePolAlcance = False
                End If
                    
               Do While Not objReg.EOF
                    Ternro = objReg!Ternro
                    Flog.writeline "--------------------------------------"
                    Flog.writeline "Inicio Empleado:" & Ternro
                    Flog.writeline "--------------------------------------"

                    'FGZ - 14/04/2010 ----------------------------------------------------------------------
                    'StrSql = "SELECT * FROM vacacion "
                    'StrSql = StrSql & " WHERE vacfecdesde <= " & ConvFecha(fecha_hasta)
                    'StrSql = StrSql & " AND  vacfechasta >= " & ConvFecha(fecha_desde)
                    'StrSql = StrSql & " ORDER BY vacnro"
                    
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
                    
                    
                    'FGZ - 24/06/2009 - le agregue la llamada a la politica
                    GenerarSituacionRevista = False
                    Call Politica(1509)
                    If GenerarSituacionRevista Then
                        'Tiene la logica inversa, es decir, si esta activa ==> quiere decir que no se debe generar la situacion de revista
                        GenerarSituacionRevista = False
                    Else
                        'quiere decir que no esta configurada o no tiene alcance, en cualquier caso hay que generar la situacion de revista
                        GenerarSituacionRevista = True
                    End If
                    
                   'FGZ - 24/06/2009 -------
                    Diashabiles_LV = False
                    PoliticaOK = False
                    Call Politica(1510)
                    'FGZ - 24/06/2009 -------
                    
                    
                    Do While Not rs_Periodos_Vac.EOF
                    
                        MyBeginTrans
                        
                        Flog.writeline "Periodo de Vacaciones:" & rs_Periodos_Vac!vacnro & " " & rs_Periodos_Vac!vacdesc
                        NroVac = rs_Periodos_Vac!vacnro
                    
                        'Genero las licencias para el periodo
                        Call Licencias3(NroVac, Ternro, Autorizadas)
            
                        MyCommitTrans
                        
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
                    
                        rs_Periodos_Vac.MoveNext
                    Loop
                                        
                    objReg.MoveNext
                Loop
                
            Case 1: '--------------------------------------------------------------------------------------------------------------------
                Flog.writeline "Modelo de vacaciones de Uruguay nro." & modeloPais
                
            Case 2:
                Flog.writeline "Modelo de vacaciones de Chile nro." & modeloPais
                
            Case 3: 'Colombia
                Flog.writeline "Modelo de vacaciones de Colombia nro." & modeloPais
                
            Case 4: 'Costa Rica
                Flog.writeline "Modelo de vacaciones de Costa Rica Nro." & modeloPais
                
                
                'Seteo el incremento del progreso
                CEmpleadosAProc = objReg.RecordCount
                If CEmpleadosAProc = 0 Then
                    CEmpleadosAProc = 1
                End If
'                'CDiasAProc = rs_Periodos_Vac.RecordCount
                If CDiasAProc = 0 Then
                    CDiasAProc = 1
                End If
                IncPorc = ((100 / CEmpleadosAProc) * (100 / CDiasAProc)) / 100
                
                'EAM- Carga los parametros Globales
                Call ParametrosGlobales
                
                Do While Not objReg.EOF
                    Ternro = objReg!Ternro
                    Flog.writeline "Inicio Empleado:" & Ternro
                    
                    'FGZ - 24/06/2009 -------
                    Diashabiles_LV = False
                    PoliticaOK = False
                    Call Politica(1510)
                    'FGZ - 24/06/2009 -------
                    
                    
'                Do While Not rs_Periodos_Vac.EOF
                    
                    MyBeginTrans
                        
'                    Flog.writeline "Periodo de Vacaciones:" & rs_Periodos_Vac!vacnro & " " & rs_Periodos_Vac!vacdesc
'                    NroVac = rs_Periodos_Vac!vacnro
                    
                    'Genero las licencias para el periodo
                    Call Licencias_CR(NroVac, Ternro, Autorizadas)
            
                    MyCommitTrans
                        
'siguiente:
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
                    
                       ' rs_Periodos_Vac.MoveNext
                    'Loop
                                        
                    objReg.MoveNext
                Loop
            Case 6: 'Paraguay
               Flog.writeline "Modelo de vacaciones de Argentina nro." & modeloPais
               l_TienePolAlcance = False
               Do While Not objReg.EOF
                    Ternro = objReg!Ternro
                    Flog.writeline "--------------------------------------"
                    Flog.writeline "Inicio Empleado:" & Ternro
                    Flog.writeline "--------------------------------------"
                    l_TienePolAlcance = False
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
                    
                        If Usa1515 = False Then
                            StrSql = "SELECT vacacion.vacdesc, vacacion.vacnro, vacacion.vacfecdesde, vacacion.vacfechasta, vacacion.vacanio "
                            StrSql = StrSql & "FROM  vacacion "
                            StrSql = StrSql & " WHERE vacfecdesde <= " & ConvFecha(fecha_hasta)
                            StrSql = StrSql & " AND  (vacfechasta >= " & ConvFecha(fecha_desde) & " OR vacfechasta IS NULL)"
                            StrSql = StrSql & " ORDER BY vacfecdesde "
                        Else
                            'USA POL. 1515
                            StrSql = "SELECT vacacion.vacdesc, vacacion.vacnro, vacacion.vacfecdesde, vacacion.vacfechasta, vacacion.vacanio "
                            StrSql = StrSql & "FROM  vacacion "
                            StrSql = StrSql & " INNER JOIN vac_alcan ON vac_alcan.vacnro = vacacion.vacnro"
                            StrSql = StrSql & " WHERE vacacion.vacanio = " & vacanio
                            StrSql = StrSql & " AND vac_alcan.alcannivel =1"
                            StrSql = StrSql & " AND vac_alcan.origen = " & Ternro
                            StrSql = StrSql & " ORDER BY vac_alcan.vacfecdesde "
                        End If
                    End If
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
                    
                    
                    'FGZ - 24/06/2009 - le agregue la llamada a la politica
                    GenerarSituacionRevista = False
                    Call Politica(1509)
                    If GenerarSituacionRevista Then
                        'Tiene la logica inversa, es decir, si esta activa ==> quiere decir que no se debe generar la situacion de revista
                        GenerarSituacionRevista = False
                    Else
                        'quiere decir que no esta configurada o no tiene alcance, en cualquier caso hay que generar la situacion de revista
                        GenerarSituacionRevista = True
                    End If
                    
                   'FGZ - 24/06/2009 -------
                    Diashabiles_LV = False
                    PoliticaOK = False
                    Call Politica(1510)
                    'FGZ - 24/06/2009 -------
                    
                    
                    Do While Not rs_Periodos_Vac.EOF
                    
                        MyBeginTrans
                        
                        Flog.writeline "Periodo de Vacaciones:" & rs_Periodos_Vac!vacnro & " " & rs_Periodos_Vac!vacdesc
                        NroVac = rs_Periodos_Vac!vacnro
                    
                        'Genero las licencias para el periodo
                        Call Licencias3(NroVac, Ternro, Autorizadas)
            
                        MyCommitTrans
                        
'siguiente:
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
                    
                        rs_Periodos_Vac.MoveNext
                    Loop
                                        
                    objReg.MoveNext
                Loop
            Case 7 'El salvador
                Flog.writeline "Modelo de vacaciones de El Salvador Nro." & modeloPais

               StrSql = "SELECT * FROM alcance_testr WHERE tanro= 21"
                OpenRecordset StrSql, rs
                If Not rs.EOF Then
                    l_TienePolAlcance = True
                Else
                    l_TienePolAlcance = False
                End If
                    
               Do While Not objReg.EOF
                    Ternro = objReg!Ternro
                    Flog.writeline "--------------------------------------"
                    Flog.writeline "Inicio Empleado:" & Ternro
                    Flog.writeline "--------------------------------------"

                    'FGZ - 14/04/2010 ----------------------------------------------------------------------
                    'StrSql = "SELECT * FROM vacacion "
                    'StrSql = StrSql & " WHERE vacfecdesde <= " & ConvFecha(fecha_hasta)
                    'StrSql = StrSql & " AND  vacfechasta >= " & ConvFecha(fecha_desde)
                    'StrSql = StrSql & " ORDER BY vacnro"
                    
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
                        StrSql = StrSql & " WHERE ternro= " & Ternro & " and (vacfecdesde <= " & ConvFecha(fecha_hasta) 'mdf
                        StrSql = StrSql & " AND  (vacfechasta >= " & ConvFecha(fecha_desde) & " OR vacfechasta IS NULL))"
                        StrSql = StrSql & " ORDER BY vacfecdesde "
                    End If
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
                    
                    
                    'FGZ - 24/06/2009 - le agregue la llamada a la politica
                    GenerarSituacionRevista = False
                    Call Politica(1509)
                    If GenerarSituacionRevista Then
                        'Tiene la logica inversa, es decir, si esta activa ==> quiere decir que no se debe generar la situacion de revista
                        GenerarSituacionRevista = False
                    Else
                        'quiere decir que no esta configurada o no tiene alcance, en cualquier caso hay que generar la situacion de revista
                        GenerarSituacionRevista = True
                    End If
                    
                   'FGZ - 24/06/2009 -------
                    Diashabiles_LV = False
                    PoliticaOK = False
                    Call Politica(1510)
                    'FGZ - 24/06/2009 -------
                    
                    
                    Do While Not rs_Periodos_Vac.EOF
                    
                        MyBeginTrans
                        
                        Flog.writeline "Periodo de Vacaciones:" & rs_Periodos_Vac!vacnro & " " & rs_Periodos_Vac!vacdesc
                        NroVac = rs_Periodos_Vac!vacnro
                    
                        'Genero las licencias para el periodo
                        Call Licencias3(NroVac, Ternro, Autorizadas)
            
                        MyCommitTrans
                        
'siguiente7:
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
                    
                        rs_Periodos_Vac.MoveNext
                    Loop
                                        
                    objReg.MoveNext
                Loop


'
'                'Seteo el incremento del progreso
'                CEmpleadosAProc = objReg.RecordCount
'                If CEmpleadosAProc = 0 Then
'                    CEmpleadosAProc = 1
'                End If
''                'CDiasAProc = rs_Periodos_Vac.RecordCount
'                If CDiasAProc = 0 Then
'                    CDiasAProc = 1
'                End If
'                IncPorc = ((100 / CEmpleadosAProc) * (100 / CDiasAProc)) / 100
'
'                'EAM- Carga los parametros Globales
'                Call ParametrosGlobales
'
'                Do While Not objReg.EOF
'                    Ternro = objReg!Ternro
'                    Flog.writeline "Inicio Empleado:" & Ternro
'
'                    'FGZ - 24/06/2009 -------
'                    Diashabiles_LV = False
'                    PoliticaOK = False
'                    Call Politica(1510)
'                    'FGZ - 24/06/2009 -------
'
'
'                    MyBeginTrans
'
'                    Call Licencias3(NroVac, Ternro, Autorizadas)
'                    'Call Licencias_CR(NroVac, Ternro, Autorizadas)
'
'                    MyCommitTrans
'
'                    'siguiente:
'                    Progreso = Progreso + IncPorc
'
'                    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & " WHERE bpronro = " & NroProceso
'                    objConnProgreso.Execute StrSql, , adExecuteNoRecords
'
'                        If SinError Then
'                             ' borro
'                             StrSql = "DELETE FROM batch_empleado WHERE ternro = " & Ternro & " AND bpronro = " & NroProceso
'                             objConnProgreso.Execute StrSql, , adExecuteNoRecords
'                        Else
'                             StrSql = "UPDATE batch_empleado SET estado = 'Error' WHERE ternro = " & Ternro & " AND bpronro = " & NroProceso
'                             objConnProgreso.Execute StrSql, , adExecuteNoRecords
'                        End If
'
'                       ' rs_Periodos_Vac.MoveNext
'                    'Loop
'
'                    objReg.MoveNext
'                Loop
            
        End Select
    



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
            objConn.Execute StrSql, , adExecuteNoRecords
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



Private Sub Licencias(NroVac As Long, Ternro As Long, ByVal Autorizadas As Boolean)
Dim hay_licencia As Boolean
Dim NroTipVac As Long
Dim nroEmpLic

Dim RsLic As New ADODB.Recordset
Dim rs_vacdiascor As New ADODB.Recordset
Dim rs_tipovacac As New ADODB.Recordset
Dim rs_existeLic As New ADODB.Recordset

'    StrSql = "SELECT * FROM vacdiascor WHERE ternro = " & Ternro & " AND vacnro = " & NroVac
'    OpenRecordset StrSql, rs_vacdiascor
'    If Not rs_vacdiascor.EOF Then
'
'        StrSql = "SELECT * FROM tipovacac WHERE tipvacnro = " & rs_vacdiascor!tipvacnro
'        OpenRecordset StrSql, rs_tipovacac
'        If Not rs_tipovacac.EOF Then
'            NroTipVac = rs_tipovacac!tipvacnro
'        Else
'            NroTipVac = 1
'        End If
'    Else
'        NroTipVac = 1
'    End If


    StrSql = "SELECT * FROM vacdiasped WHERE " & _
             " vacnro = " & NroVac & " AND vdiaspedestado = -1 And Ternro = " & Ternro & _
             " AND vdiapedcant <> 0 " & _
             " AND vdiapeddesde >= " & ConvFecha(fecha_desde) & _
             " AND vdiapeddesde <=" & ConvFecha(fecha_hasta)
    OpenRecordset StrSql, objRs
    
    Do While Not objRs.EOF
        
        hay_licencia = False
        
        StrSql = " SELECT * FROM emp_lic WHERE (empleado = " & Ternro & " ) AND (tdnro = 2) "
        OpenRecordset StrSql, RsLic
        Do While Not RsLic.EOF
            If (RsLic!elfechadesde >= objRs!vdiapeddesde And RsLic!elfechadesde <= objRs!vdiapedhasta) Or _
               (RsLic!elfechahasta >= objRs!vdiapeddesde And RsLic!elfechahasta <= objRs!vdiapedhasta) Or _
               (RsLic!elfechadesde <= objRs!vdiapeddesde And RsLic!elfechahasta >= objRs!vdiapedhasta) Or _
               (objRs!vdiapeddesde <= RsLic!elfechadesde And objRs!vdiapedhasta >= RsLic!elfechahasta) _
            Then
                hay_licencia = True
                If Reproceso Then
                    Call BorrarLicencia(RsLic!emp_licnro)
                End If
            End If
            RsLic.MoveNext
        Loop
        
        
        'reviso si ya existe esa Licencia
        StrSql = " SELECT * FROM emp_lic WHERE (empleado = " & Ternro & _
                 " ) AND (tdnro = 2) " & _
                 " AND eltipo = 1" & _
                 " AND elfechadesde = " & ConvFecha(objRs!vdiapeddesde) & _
                 " AND elfechahasta = " & ConvFecha(objRs!vdiapedhasta)
                 OpenRecordset StrSql, rs_existeLic

        If rs_existeLic.EOF Then
            StrSql = "INSERT INTO emp_lic (elcantdias,elcantdiasfer,elcantdiashab,eldiacompleto,eltipo,elfechadesde,elfechahasta,elhoradesde,elhorahasta,tdnro,empleado) VALUES ("
            
    '        If NroTipVac = 1 Then ' dias corridos
    '            StrSql = StrSql & objRs!vdiapedcant & ","
    '        Else
                StrSql = StrSql & (CInt(DateDiff("d", objRs!vdiapeddesde, objRs!vdiapedhasta)) + 1) & ","
    '        End If
            StrSql = StrSql & objRs!vdiaspedferiados & ","
            StrSql = StrSql & objRs!vdiaspedhabiles & ","
            StrSql = StrSql & "-1,"
            StrSql = StrSql & "1,"
            StrSql = StrSql & ConvFecha(objRs!vdiapeddesde) & ","
            StrSql = StrSql & ConvFecha(objRs!vdiapedhasta) & ","
            StrSql = StrSql & "null,"
            StrSql = StrSql & "null,"
            StrSql = StrSql & "2,"
            StrSql = StrSql & Ternro & ")"
            
            objConn.Execute StrSql, , adExecuteNoRecords
            
            nroEmpLic = getLastIdentity(objConn, "emp_lic")
            
            StrSql = "INSERT INTO lic_vacacion (emp_licnro,licvacmanual,vacnro) VALUES ("
            StrSql = StrSql & nroEmpLic & ","
            StrSql = StrSql & "0,"
            StrSql = StrSql & NroVac & ")"
            
            objConn.Execute StrSql, , adExecuteNoRecords
        Else
            ' La licencia ya existe
        End If
        objRs.MoveNext
    Loop
        
        
End Sub

Private Sub Licencias2(NroVac As Long, Ternro As Long, ByVal Autorizadas As Boolean)
Dim hay_licencia As Boolean
Dim NroTipVac As Long

Dim RsLic As New ADODB.Recordset
Dim rs_vacdiascor As New ADODB.Recordset
Dim rs_tipovacac As New ADODB.Recordset
Dim rs_existeLic As New ADODB.Recordset

    ' Busco el tipo de vacacion que le corresponde
    StrSql = "SELECT * FROM vacdiascor WHERE ternro = " & Ternro & " AND vacnro = " & NroVac
    OpenRecordset StrSql, rs_vacdiascor
    If Not rs_vacdiascor.EOF Then
        
        StrSql = "SELECT * FROM tipovacac WHERE tipvacnro = " & rs_vacdiascor!tipvacnro
        OpenRecordset StrSql, rs_tipovacac
        If Not rs_tipovacac.EOF Then
            NroTipVac = rs_tipovacac!tipvacnro
        End If
    End If


    StrSql = "SELECT * FROM vacdiasped WHERE " & _
             " vacnro = " & NroVac & " AND vdiaspedestado = -1 And Ternro = " & Ternro & _
             " AND vdiapedcant <> 0 " & _
             " AND vdiapeddesde >= " & ConvFecha(fecha_desde) & _
             " AND vdiapeddesde <=" & ConvFecha(fecha_hasta)
    OpenRecordset StrSql, objRs
    
    Do While Not objRs.EOF
        
        hay_licencia = False
        
        StrSql = " SELECT * FROM emp_lic WHERE (empleado = " & Ternro & " ) AND (tdnro = 2) "
        OpenRecordset StrSql, RsLic
        Do While Not RsLic.EOF
            If (RsLic!elfechadesde >= objRs!vdiapeddesde And RsLic!elfechadesde <= objRs!vdiapedhasta) Or _
               (RsLic!elfechahasta >= objRs!vdiapeddesde And RsLic!elfechahasta <= objRs!vdiapedhasta) Or _
               (RsLic!elfechadesde <= objRs!vdiapeddesde And RsLic!elfechahasta >= objRs!vdiapedhasta) Or _
               (objRs!vdiapeddesde <= RsLic!elfechadesde And objRs!vdiapedhasta >= RsLic!elfechahasta) _
            Then
                hay_licencia = True
                If Reproceso Then
                    Call BorrarLicencia(RsLic!emp_licnro)
                End If
            End If
            RsLic.MoveNext
        Loop
        Call SepararLicencias(NroTipVac, objRs!vdiapeddesde, objRs!vdiapedcant, Ternro, NroVac, Autorizadas)
        
        objRs.MoveNext
    Loop
        
        
End Sub

Private Sub Licencias_CR(NroVac As Long, Ternro As Long, ByVal Autorizadas As Boolean)
Dim hay_licencia As Boolean
Dim NroTipVac As Long

Dim Genero As Boolean
Dim Hay_Pedidos As Boolean

Dim RsLic As New ADODB.Recordset
Dim rs_vacdiascor As New ADODB.Recordset
Dim rs_tipovacac As New ADODB.Recordset
Dim rs_existeLic As New ADODB.Recordset
Dim rs_Vacdiasped As New ADODB.Recordset

    ' Busco el tipo de vacacion que le corresponde
'    StrSql = "SELECT * FROM vacdiascor WHERE ternro = " & Ternro & " AND vacnro = " & NroVac
'    OpenRecordset StrSql, rs_vacdiascor
'    If Not rs_vacdiascor.EOF Then
'
'        StrSql = "SELECT * FROM tipovacac WHERE tipvacnro = " & rs_vacdiascor!tipvacnro
'        OpenRecordset StrSql, rs_tipovacac
'        If Not rs_tipovacac.EOF Then
'            NroTipVac = rs_tipovacac!tipvacnro
'        End If
'    End If


    StrSql = "SELECT * FROM vacdiasped WHERE Ternro = " & Ternro & " AND vdiapedcant <> 0 " & _
             " AND vdiapeddesde >= " & ConvFecha(fecha_desde) & " AND vdiapeddesde <=" & ConvFecha(fecha_hasta)
    OpenRecordset StrSql, rs_Vacdiasped
    
    If rs_Vacdiasped.EOF Then
        Flog.writeline "No Se encontraron Pedidos de Vacaciones"
    End If
    Hay_Pedidos = False
    
    Do While Not rs_Vacdiasped.EOF
        If CBool(rs_Vacdiasped!vdiaspedestado) Then
            Hay_Pedidos = True
            hay_licencia = False
            Genero = True
            StrSql = " SELECT * FROM emp_lic "
            'FGZ - 18/04/2008 - se cambió esta linea sino no va a encontrar otro tipo de lic que no sea de vac
            'StrSql = StrSql & " INNER JOIN lic_vacacion ON emp_lic.emp_licnro = lic_vacacion.emp_licnro AND lic_vacacion.vacnro = " & NroVac
            'StrSql = StrSql & " LEFT JOIN lic_vacacion ON emp_lic.emp_licnro = lic_vacacion.emp_licnro AND lic_vacacion.vacnro = " & NroVac
            StrSql = StrSql & " WHERE empleado = " & Ternro & " and tdnro= 2"
            If RsLic.State = adStateOpen Then RsLic.Close
            OpenRecordset StrSql, RsLic
            Do While Not RsLic.EOF
                If (RsLic!elfechadesde >= rs_Vacdiasped!vdiapeddesde And RsLic!elfechadesde <= rs_Vacdiasped!vdiapedhasta) Or _
                   (RsLic!elfechahasta >= rs_Vacdiasped!vdiapeddesde And RsLic!elfechahasta <= rs_Vacdiasped!vdiapedhasta) Or _
                   (RsLic!elfechadesde <= rs_Vacdiasped!vdiapeddesde And RsLic!elfechahasta >= rs_Vacdiasped!vdiapedhasta) Or _
                   (rs_Vacdiasped!vdiapeddesde <= RsLic!elfechadesde And rs_Vacdiasped!vdiapedhasta >= RsLic!elfechahasta) _
                Then
                    Flog.writeline "Se encontró una Lic.de tipo " & RsLic!tdnro & " que se superpone con otra licencia"
                    hay_licencia = True
                    If Reproceso And RsLic!tdnro = 2 Then
                        Flog.writeline "Borro la La Lic. porque es de Vacaciones y reproceso"
                        Call BorrarLicencia(RsLic!emp_licnro)
                        If GenerarSituacionRevista = True Then
                            Call BorrarSituacionRevista(Ternro, RsLic!elfechadesde, RsLic!elfechahasta)
                        Else
                            Flog.writeline "Las Licencias tienen Situacion de Revista configurado para no generarse. Politica 1509."
                        End If
                    Else
                        Genero = False
                        Flog.writeline "Licencia de Vac. No Generada porque se superpone con otra licencia de tipo (" & RsLic!tdnro & ")"
                    End If
                End If
                RsLic.MoveNext
            Loop
            
            If Genero Then
'                'FGZ - 26/06/2009 -----------------------
'                If PoliticaOK And Diashabiles_LV Then
'                    Call SepararLicencias(NroTipVac, rs_Vacdiasped!vdiapeddesde, rs_Vacdiasped!vdiapedcant, Ternro, NroVac, Autorizadas)
'                Else
'                    Call SepararLicencias(NroTipVac, rs_Vacdiasped!vdiapeddesde, rs_Vacdiasped!vdiaspedhabiles, Ternro, NroVac, Autorizadas)
'                End If
'                'Call SepararLicencias(NroTipVac, rs_Vacdiasped!vdiapeddesde, rs_Vacdiasped!vdiaspedhabiles, Ternro, NroVac, Autorizadas)
'                'FGZ - 26/06/2009 -----------------------
                'Call SepararLicencias(NroTipVac, rs_Vacdiasped!vdiapeddesde, rs_Vacdiasped!vdiaspedhabiles, Ternro, NroVac, Autorizadas)
                Call InsertarLicencia_CR(rs_Vacdiasped!vdiapeddesde, rs_Vacdiasped!vdiapedhasta, rs_Vacdiasped!vdiaspedferiados, rs_Vacdiasped!vdiaspedhabiles, Ternro, NroVac, Autorizadas)
            Else
                Flog.writeline "Licencia de Vac. No Generada porque se superpone con otra licencia"
            End If
        Else
            Flog.writeline "El pedido esta en estado pendiente. " & rs_Vacdiasped!vdiapeddesde & " " & rs_Vacdiasped!vdiapedhasta & ". No se generará la licencia."
        End If
        rs_Vacdiasped.MoveNext
    Loop
        
        
End Sub

Private Sub Licencias_SV(NroVac As Long, Ternro As Long, ByVal Autorizadas As Boolean)
Dim hay_licencia As Boolean
Dim NroTipVac As Long

Dim Genero As Boolean
Dim Hay_Pedidos As Boolean

Dim RsLic As New ADODB.Recordset
Dim rs_vacdiascor As New ADODB.Recordset
Dim rs_tipovacac As New ADODB.Recordset
Dim rs_existeLic As New ADODB.Recordset
Dim rs_Vacdiasped As New ADODB.Recordset


    StrSql = "SELECT * FROM vacdiasped WHERE Ternro = " & Ternro & " AND vdiapedcant <> 0 " & _
             " AND vdiapeddesde >= " & ConvFecha(fecha_desde) & " AND vdiapeddesde <=" & ConvFecha(fecha_hasta)
    OpenRecordset StrSql, rs_Vacdiasped
    
    If rs_Vacdiasped.EOF Then
        Flog.writeline "No Se encontraron Pedidos de Vacaciones"
    End If
    Hay_Pedidos = False
    
    Do While Not rs_Vacdiasped.EOF
        If CBool(rs_Vacdiasped!vdiaspedestado) Then
            Hay_Pedidos = True
            hay_licencia = False
            Genero = True
            StrSql = " SELECT * FROM emp_lic WHERE empleado = " & Ternro & " and tdnro= 2"
            If RsLic.State = adStateOpen Then RsLic.Close
            OpenRecordset StrSql, RsLic
            Do While Not RsLic.EOF
                If (RsLic!elfechadesde >= rs_Vacdiasped!vdiapeddesde And RsLic!elfechadesde <= rs_Vacdiasped!vdiapedhasta) Or _
                   (RsLic!elfechahasta >= rs_Vacdiasped!vdiapeddesde And RsLic!elfechahasta <= rs_Vacdiasped!vdiapedhasta) Or _
                   (RsLic!elfechadesde <= rs_Vacdiasped!vdiapeddesde And RsLic!elfechahasta >= rs_Vacdiasped!vdiapedhasta) Or _
                   (rs_Vacdiasped!vdiapeddesde <= RsLic!elfechadesde And rs_Vacdiasped!vdiapedhasta >= RsLic!elfechahasta) _
                Then
                    Flog.writeline "Se encontró una Lic.de tipo " & RsLic!tdnro & " que se superpone con otra licencia"
                    hay_licencia = True
                    If Reproceso And RsLic!tdnro = 2 Then
                        Flog.writeline "Borro la La Lic. porque es de Vacaciones y reproceso"
                        Call BorrarLicencia(RsLic!emp_licnro)
                        If GenerarSituacionRevista = True Then
                            Call BorrarSituacionRevista(Ternro, RsLic!elfechadesde, RsLic!elfechahasta)
                        Else
                            Flog.writeline "Las Licencias tienen Situacion de Revista configurado para no generarse. Politica 1509."
                        End If
                    Else
                        Genero = False
                        Flog.writeline "Licencia de Vac. No Generada porque se superpone con otra licencia de tipo (" & RsLic!tdnro & ")"
                    End If
                End If
                RsLic.MoveNext
            Loop
            
            If Genero Then

                Call InsertarLicencia_CR(rs_Vacdiasped!vdiapeddesde, rs_Vacdiasped!vdiapedhasta, rs_Vacdiasped!vdiaspedferiados, rs_Vacdiasped!vdiaspedhabiles, Ternro, NroVac, Autorizadas)
            Else
                Flog.writeline "Licencia de Vac. No Generada porque se superpone con otra licencia"
            End If
        Else
            Flog.writeline "El pedido esta en estado pendiente. " & rs_Vacdiasped!vdiapeddesde & " " & rs_Vacdiasped!vdiapedhasta & ". No se generará la licencia."
        End If
        rs_Vacdiasped.MoveNext
    Loop
        
        
End Sub

Private Sub Licencias3(NroVac As Long, Ternro As Long, ByVal Autorizadas As Boolean)
Dim hay_licencia As Boolean
Dim NroTipVac As Long

Dim Genero As Boolean
Dim Hay_Pedidos As Boolean

Dim RsLic As New ADODB.Recordset
Dim rs_vacdiascor As New ADODB.Recordset
Dim rs_tipovacac As New ADODB.Recordset
Dim rs_existeLic As New ADODB.Recordset
Dim rs_Vacdiasped As New ADODB.Recordset

    ' Busco el tipo de vacacion que le corresponde
    StrSql = "SELECT * FROM vacdiascor WHERE ternro = " & Ternro & " AND vacnro = " & NroVac
    OpenRecordset StrSql, rs_vacdiascor
    If Not rs_vacdiascor.EOF Then
        
        StrSql = "SELECT * FROM tipovacac WHERE tipvacnro = " & rs_vacdiascor!tipvacnro
        OpenRecordset StrSql, rs_tipovacac
        If Not rs_tipovacac.EOF Then
            NroTipVac = rs_tipovacac!tipvacnro
        End If
    End If

'    StrSql = "SELECT * FROM vacdiasped WHERE " & _
'             " vacnro = " & NroVac & " AND vdiaspedestado = -1 And Ternro = " & Ternro & _
'             " AND vdiapedcant <> 0 " & _
'             " AND vdiapeddesde >= " & ConvFecha(fecha_desde) & _
'             " AND vdiapeddesde <=" & ConvFecha(fecha_hasta)
'    OpenRecordset StrSql, rs_Vacdiasped

    StrSql = "SELECT * FROM vacdiasped WHERE " & _
             " vacnro = " & NroVac & " And Ternro = " & Ternro & _
             " AND vdiapedcant <> 0 " & _
             " AND vdiapeddesde >= " & ConvFecha(fecha_desde) & _
             " AND vdiapeddesde <=" & ConvFecha(fecha_hasta)
    OpenRecordset StrSql, rs_Vacdiasped
    If rs_Vacdiasped.EOF Then
        Flog.writeline "No Se encontraron Pedidos de Vacaciones"
    End If
    Hay_Pedidos = False
    Do While Not rs_Vacdiasped.EOF
        If CBool(rs_Vacdiasped!vdiaspedestado) Then
            Hay_Pedidos = True
            hay_licencia = False
            Genero = True
                       StrSql = " SELECT * FROM emp_lic "
            'FGZ - 18/04/2008 - se cambió esta linea sino no va a encontrar otro tipo de lic que no sea de vac
            'StrSql = StrSql & " INNER JOIN lic_vacacion ON emp_lic.emp_licnro = lic_vacacion.emp_licnro AND lic_vacacion.vacnro = " & NroVac
            StrSql = StrSql & " LEFT JOIN lic_vacacion ON emp_lic.emp_licnro = lic_vacacion.emp_licnro AND lic_vacacion.vacnro = " & NroVac
            StrSql = StrSql & " WHERE empleado = " & Ternro
            If RsLic.State = adStateOpen Then RsLic.Close
            OpenRecordset StrSql, RsLic
            Do While Not RsLic.EOF
                If (RsLic!elfechadesde >= rs_Vacdiasped!vdiapeddesde And RsLic!elfechadesde <= rs_Vacdiasped!vdiapedhasta) Or _
                   (RsLic!elfechahasta >= rs_Vacdiasped!vdiapeddesde And RsLic!elfechahasta <= rs_Vacdiasped!vdiapedhasta) Or _
                   (RsLic!elfechadesde <= rs_Vacdiasped!vdiapeddesde And RsLic!elfechahasta >= rs_Vacdiasped!vdiapedhasta) Or _
                   (rs_Vacdiasped!vdiapeddesde <= RsLic!elfechadesde And rs_Vacdiasped!vdiapedhasta >= RsLic!elfechahasta) _
                Then
                    Flog.writeline "Se encontró una Lic.de tipo " & RsLic!tdnro & " que se superpone con otra licencia"
                    hay_licencia = True
                    If Reproceso And RsLic!tdnro = 2 Then
                        Flog.writeline "Borro la La Lic. porque es de Vacaciones y reproceso"
                        Call BorrarLicencia(RsLic!emp_licnro)
                        If GenerarSituacionRevista = True Then
                            Call BorrarSituacionRevista(Ternro, RsLic!elfechadesde, RsLic!elfechahasta)
                        Else
                            Flog.writeline "Las Licencias tienen Situacion de Revista configurado para no generarse. Politica 1509."
                        End If
                    Else
                        Genero = False
                        Flog.writeline "Licencia de Vac. No Generada porque se superpone con otra licencia de tipo (" & RsLic!tdnro & ")"
                    End If
                End If
                RsLic.MoveNext
            Loop
            
            If Genero Then
                'FGZ - 26/06/2009 -----------------------
                If PoliticaOK And Diashabiles_LV Then
                    Call SepararLicencias(NroTipVac, rs_Vacdiasped!vdiapeddesde, rs_Vacdiasped!vdiapedcant, Ternro, NroVac, Autorizadas)
                Else
                    Call SepararLicencias(NroTipVac, rs_Vacdiasped!vdiapeddesde, rs_Vacdiasped!vdiaspedhabiles, Ternro, NroVac, Autorizadas)
                End If
                'Call SepararLicencias(NroTipVac, rs_Vacdiasped!vdiapeddesde, rs_Vacdiasped!vdiaspedhabiles, Ternro, NroVac, Autorizadas)
                'FGZ - 26/06/2009 -----------------------
                
            Else
                'Flog.writeline "Licencia de Vac. No Generada porque se superpone con otra licencia"
            End If
        Else
            Flog.writeline "El pedido esta en estado pendiente. " & rs_Vacdiasped!vdiapeddesde & " " & rs_Vacdiasped!vdiapedhasta & ". No se generará la licencia."
        End If
        rs_Vacdiasped.MoveNext
    Loop
        
        
End Sub

Private Sub InsertarLicencia_CR(ByVal fechadesde As Date, ByVal FechaHasta As Date, ByVal cFeriados As Integer, ByVal CHabiles As Integer, ByVal Ternro As Long, ByVal NroVac As Long, ByVal Autorizadas As Boolean)
Dim rs_existeLic As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim rs_Est As New ADODB.Recordset
Dim rsFirmas As New ADODB.Recordset
Dim nroEmpLic
Dim Inserto As Boolean

Dim Estrnro_SitRev As String


    'reviso si ya existe esa Licencia
    StrSql = " SELECT * FROM emp_lic WHERE (empleado = " & Ternro & " ) AND (tdnro = 2) " & _
             " AND eltipo = 1 AND elfechadesde = " & ConvFecha(fechadesde) & " AND elfechahasta = " & ConvFecha(FechaHasta)
    OpenRecordset StrSql, rs_existeLic

    If rs_existeLic.EOF Then
        
        'Reviso que tenga fase activa que contenga rango de fechas
        If activo(Ternro, fechadesde, FechaHasta) Then
        
            StrSql = "INSERT INTO emp_lic (elcantdias,elcantdiasfer,elcantdiashab,eldiacompleto,eltipo,elfechadesde,elfechahasta,elhoradesde,elhorahasta,tdnro,licestnro,empleado) VALUES ("
            'FGZ - 23/09/2004
            'StrSql = StrSql & (CInt(DateDiff("d", FechaDesde, FechaHasta)) + 1) & ","
            StrSql = StrSql & CHabiles & ","
            
            StrSql = StrSql & cFeriados & ","
            StrSql = StrSql & CHabiles & ","
            StrSql = StrSql & "-1,"
            StrSql = StrSql & "1,"
            StrSql = StrSql & ConvFecha(fechadesde) & ","
            StrSql = StrSql & ConvFecha(FechaHasta) & ","
            StrSql = StrSql & "null,"
            StrSql = StrSql & "null,"
            'StrSql = StrSql & "2,1,"
            If Autorizadas Then
                StrSql = StrSql & "2,2,"
            Else
                StrSql = StrSql & "2,1,"
            End If
            StrSql = StrSql & Ternro & ")"
            Flog.writeline "Inserto Licencia. SQL " & StrSql
            objConn.Execute StrSql, , adExecuteNoRecords
                
            nroEmpLic = getLastIdentity(objConn, "emp_lic")
            Flog.writeline "Licencia insertada nro " & nroEmpLic
            Flog.writeline "Inserto complemento"
                    
            
            StrSql = "INSERT INTO gti_justificacion ( jusanterior,juscodext,jusdesde,jusdiacompleto,jushasta,jussigla,jussistema,ternro,tjusnro,turnro,jushoradesde,jushorahasta,juseltipo,juselorden,juselmaxhoras )" & _
            " VALUES( -1," & nroEmpLic & "," & ConvFecha(fechadesde) & ",-1," & ConvFecha(FechaHasta) & ",'LIC',-1," & Ternro & ",1,0,NULL,NULL, 1 ,NULL,0)"
            objConn.Execute StrSql, , adExecuteNoRecords
            Flog.writeline "Complemento Justificacion insertado"
            
            
            Inserto = True
            'Valido que este activo el circuito
            If Firma_Licencias Then
                Call generarFirma(6, nroEmpLic, True)
            Else
                Flog.writeline "Circuito de Autorización Inactivo"
            End If
            
            
        Else
            Flog.writeline "El rango de fechas de la licencia con fechas " & fechadesde & " " & FechaHasta & " se superpone con un período inactivo del empleado " & Ternro
        End If
    Else
        Flog.writeline "La licencia ya existe"
    End If
    Flog.writeline "Licencia insertada OK"
End Sub


Private Sub InsertarLicencia(ByVal fechadesde As Date, ByVal FechaHasta As Date, ByVal cFeriados As Integer, ByVal CHabiles As Integer, ByVal Ternro As Long, ByVal NroVac As Long, ByVal Autorizadas As Boolean)
Dim rs_existeLic As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim rs_Est As New ADODB.Recordset
Dim nroEmpLic

Dim Estrnro_SitRev As String


    'reviso si ya existe esa Licencia
    StrSql = " SELECT * FROM emp_lic WHERE (empleado = " & Ternro & _
             " ) AND (tdnro = 2) " & _
             " AND eltipo = 1" & _
             " AND elfechadesde = " & ConvFecha(fechadesde) & _
             " AND elfechahasta = " & ConvFecha(FechaHasta)
    OpenRecordset StrSql, rs_existeLic

    If rs_existeLic.EOF Then
        
        'Reviso que tenga fase activa que contenga rango de fechas
        If activo(Ternro, fechadesde, FechaHasta) Then
        
            StrSql = "INSERT INTO emp_lic (elcantdias,elcantdiasfer,elcantdiashab,eldiacompleto,eltipo,elfechadesde,elfechahasta,elhoradesde,elhorahasta,tdnro,licestnro,empleado) VALUES ("
            'FGZ - 23/09/2004
            'StrSql = StrSql & (CInt(DateDiff("d", FechaDesde, FechaHasta)) + 1) & ","
            StrSql = StrSql & CHabiles & ","
            
            StrSql = StrSql & cFeriados & ","
            StrSql = StrSql & CHabiles & ","
            StrSql = StrSql & "-1,"
            StrSql = StrSql & "1,"
            StrSql = StrSql & ConvFecha(fechadesde) & ","
            StrSql = StrSql & ConvFecha(FechaHasta) & ","
            StrSql = StrSql & "null,"
            StrSql = StrSql & "null,"
            'StrSql = StrSql & "2,1,"
            If Autorizadas Then
                StrSql = StrSql & "2,2,"
            Else
                StrSql = StrSql & "2,1,"
            End If
            StrSql = StrSql & Ternro & ")"
            Flog.writeline "Inserto Licencia. SQL " & StrSql
            objConn.Execute StrSql, , adExecuteNoRecords
                
            nroEmpLic = getLastIdentity(objConn, "emp_lic")
            Flog.writeline "Licencia insertada nro " & nroEmpLic
            Flog.writeline "Inserto complemento"
            
            StrSql = "INSERT INTO lic_vacacion (emp_licnro,licvacmanual,vacnro) VALUES ("
            StrSql = StrSql & nroEmpLic & ","
            StrSql = StrSql & "0,"
            StrSql = StrSql & NroVac & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
            Flog.writeline "Complemento insertado"
            
            StrSql = "INSERT INTO gti_justificacion ( jusanterior,juscodext,jusdesde,jusdiacompleto,jushasta,jussigla,jussistema,ternro,tjusnro,turnro,jushoradesde,jushorahasta,juseltipo,juselorden,juselmaxhoras )" & _
            " VALUES( -1," & nroEmpLic & "," & ConvFecha(fechadesde) & ",-1," & ConvFecha(FechaHasta) & ",'LIC',-1," & Ternro & ",1,0,NULL,NULL, 1 ,NULL,0)"
            objConn.Execute StrSql, , adExecuteNoRecords
            Flog.writeline "Complemento Justificacion insertado"
                        
            Flog.writeline "Situacion de revista"
            'CODIGO DE SIT.REVISTA ===========================================================
            
            'Modificacion AGD - No se genera la situacion de revista si se configura.
            If GenerarSituacionRevista = True Then
                
                StrSql = "SELECT estrnro, tdnro FROM csijp_srtd "
                StrSql = StrSql & " WHERE tdnro = 2"
                If rs.State = adStateOpen Then rs.Close
                OpenRecordset StrSql, rs
                If Not rs.EOF Then
                    Estrnro_SitRev = rs!estrnro
                End If
                If rs.State = adStateOpen Then rs.Close
            
                If Trim(Estrnro_SitRev) <> "" Then
                
                    'Busco el tipo de la situacion de revista anterior
                    StrSql = "SELECT * FROM his_estructura "
                    StrSql = StrSql & " WHERE tenro   = 30 "
                    StrSql = StrSql & " AND   ternro  = " & Ternro
                    StrSql = StrSql & " AND   htetdesde <= " & ConvFecha(fechadesde)
                    StrSql = StrSql & " AND   (htethasta >= " & ConvFecha(fechadesde)
                    StrSql = StrSql & " OR   htethasta  is null) "
                    If rs_Est.State = adStateOpen Then rs_Est.Close
                    OpenRecordset StrSql, rs_Est
                    If Not rs_Est.EOF Then
                        'la cierro un dia antes
                        If EsNulo(rs_Est!htethasta) Then
                            If Not (rs_Est!htetdesde = fechadesde) Then
                                StrSql = " UPDATE his_estructura SET "
                                StrSql = StrSql & " htethasta = " & ConvFecha(CDate(fechadesde - 1))
                                StrSql = StrSql & " WHERE tenro   = 30 "
                                StrSql = StrSql & " AND   estrnro  = " & rs_Est!estrnro
                                StrSql = StrSql & " AND   ternro  = " & Ternro
                                StrSql = StrSql & " AND   htetdesde = " & ConvFecha(rs_Est!htetdesde)
                                StrSql = StrSql & " AND   htethasta  is null "
                                objConn.Execute StrSql, , adExecuteNoRecords
                            Else
                                'la borro porque se va superponer con la licencia
                                StrSql = " DELETE his_estructura "
                                StrSql = StrSql & " WHERE tenro   = 30 "
                                StrSql = StrSql & " AND   estrnro  = " & rs_Est!estrnro
                                StrSql = StrSql & " AND   ternro  = " & Ternro
                                StrSql = StrSql & " AND   htetdesde = " & ConvFecha(rs_Est!htetdesde)
                                StrSql = StrSql & " AND   htethasta  is null "
                                objConn.Execute StrSql, , adExecuteNoRecords
                            End If
                            'Inserto la misma situacion despues de la nueva situacion (la de la licencia)
                            StrSql = "INSERT INTO his_estructura "
                            StrSql = StrSql & " (tenro, ternro, estrnro, htetdesde) "
                            StrSql = StrSql & " VALUES (30, " & Ternro & ", "
                            StrSql = StrSql & rs_Est!estrnro & ", "
                            StrSql = StrSql & ConvFecha(CDate(FechaHasta + 1)) & ")"
                            objConn.Execute StrSql, , adExecuteNoRecords
                        Else
                            If rs_Est!htethasta > FechaHasta Then
                                If rs_Est!htetdesde > fechadesde Then
                                    StrSql = " UPDATE his_estructura SET "
                                    StrSql = StrSql & " htethasta = " & ConvFecha(CDate(fechadesde - 1))
                                    StrSql = StrSql & " WHERE tenro   = 30 "
                                    StrSql = StrSql & " AND   ternro  = " & Ternro
                                    StrSql = StrSql & " AND   htetdesde = " & ConvFecha(rs_Est!htetdesde)
                                    StrSql = StrSql & " AND   htethasta  = " & ConvFecha(rs_Est!htethasta)
                                    objConn.Execute StrSql, , adExecuteNoRecords
                                Else
                                    'la borro porque se va superponer con la licencia
                                    StrSql = " DELETE his_estructura "
                                    StrSql = StrSql & " WHERE tenro   = 30 "
                                    StrSql = StrSql & " AND   estrnro  = " & rs_Est!estrnro
                                    StrSql = StrSql & " AND   ternro  = " & Ternro
                                    StrSql = StrSql & " AND   htetdesde = " & ConvFecha(rs_Est!htetdesde)
                                    StrSql = StrSql & " AND   htethasta  = " & ConvFecha(rs_Est!htethasta)
                                    objConn.Execute StrSql, , adExecuteNoRecords
                                End If
                                'Inserto la misma situacion despues de la nueva situacion (la de la licencia)
                                StrSql = "INSERT INTO his_estructura "
                                StrSql = StrSql & " (tenro, ternro, estrnro, htetdesde,htethasta) "
                                StrSql = StrSql & " VALUES (30, " & Ternro & ", "
                                StrSql = StrSql & rs_Est!estrnro & ", "
                                StrSql = StrSql & ConvFecha(CDate(FechaHasta + 1)) & ", "
                                StrSql = StrSql & ConvFecha(rs_Est!htethasta) & ")"
                                objConn.Execute StrSql, , adExecuteNoRecords
                            Else
                                If rs_Est!htetdesde > fechadesde Then
                                    StrSql = " UPDATE his_estructura SET "
                                    StrSql = StrSql & " htethasta = " & ConvFecha(CDate(fechadesde - 1))
                                    StrSql = StrSql & " WHERE tenro   = 30 "
                                    StrSql = StrSql & " AND   ternro  = " & Ternro
                                    StrSql = StrSql & " AND   htetdesde = " & ConvFecha(rs_Est!htetdesde)
                                    StrSql = StrSql & " AND   htethasta  is null "
                                    objConn.Execute StrSql, , adExecuteNoRecords
                                Else
                                    'la borro porque se va superponer con la licencia
                                    StrSql = " DELETE his_estructura "
                                    StrSql = StrSql & " WHERE tenro   = 30 "
                                    StrSql = StrSql & " AND   estrnro  = " & rs_Est!estrnro
                                    StrSql = StrSql & " AND   ternro  = " & Ternro
                                    StrSql = StrSql & " AND   htetdesde = " & ConvFecha(rs_Est!htetdesde)
                                    StrSql = StrSql & " AND   htethasta  = " & ConvFecha(rs_Est!htethasta)
                                    objConn.Execute StrSql, , adExecuteNoRecords
                                End If
                            End If
                        End If
                    End If
                
                    StrSql = "INSERT INTO his_estructura "
                    StrSql = StrSql & " (tenro, ternro, estrnro, htetdesde,htethasta) "
                    StrSql = StrSql & " VALUES (30, " & Ternro & ", "
                    StrSql = StrSql & Estrnro_SitRev & ", "
                    StrSql = StrSql & ConvFecha(fechadesde) & ", "
                    StrSql = StrSql & ConvFecha(FechaHasta) & ")"
                    objConn.Execute StrSql, , adExecuteNoRecords
                Else
                    Flog.writeline "Las Licencias por vacaciones no tienen Situacion de Revista asociado"
                End If
            Else
                Flog.writeline "Las Licencias tienen Situacion de Revista configurado para no generarse. Politica 1509."
            End If
        Else
            Flog.writeline "El rango de fechas de la licencia con fechas " & fechadesde & " " & FechaHasta & " se superpone con un período inactivo del empleado " & Ternro
        End If
    Else
        Flog.writeline "La licencia ya existe"
    End If
    Flog.writeline "Licencia insertada OK"
End Sub

Private Sub InsertarLicencia_old(ByVal fechadesde As Date, ByVal FechaHasta As Date, ByVal cFeriados As Integer, ByVal CHabiles As Integer, ByVal Ternro As Long, ByVal NroVac As Long)
Dim rs_existeLic As New ADODB.Recordset
Dim nroEmpLic

    'reviso si ya existe esa Licencia
    StrSql = " SELECT * FROM emp_lic WHERE (empleado = " & Ternro & _
             " ) AND (tdnro = 2) " & _
             " AND eltipo = 1" & _
             " AND elfechadesde = " & ConvFecha(fechadesde) & _
             " AND elfechahasta = " & ConvFecha(FechaHasta)
    OpenRecordset StrSql, rs_existeLic

    If rs_existeLic.EOF Then
        StrSql = "INSERT INTO emp_lic (elcantdias,elcantdiasfer,elcantdiashab,eldiacompleto,eltipo,elfechadesde,elfechahasta,elhoradesde,elhorahasta,tdnro,empleado) VALUES ("
        StrSql = StrSql & (CInt(DateDiff("d", fechadesde, FechaHasta)) + 1) & ","
        StrSql = StrSql & cFeriados & ","
        StrSql = StrSql & CHabiles & ","
        StrSql = StrSql & "-1,"
        StrSql = StrSql & "1,"
        StrSql = StrSql & ConvFecha(fechadesde) & ","
        StrSql = StrSql & ConvFecha(FechaHasta) & ","
        StrSql = StrSql & "null,"
        StrSql = StrSql & "null,"
        StrSql = StrSql & "2,"
        StrSql = StrSql & Ternro & ")"
            
        objConn.Execute StrSql, , adExecuteNoRecords
        
        nroEmpLic = getLastIdentity(objConn, "emp_lic")
            
        StrSql = "INSERT INTO lic_vacacion (emp_licnro,licvacmanual,vacnro) VALUES ("
        StrSql = StrSql & nroEmpLic & ","
        StrSql = StrSql & "0,"
        StrSql = StrSql & NroVac & ")"
            
        objConn.Execute StrSql, , adExecuteNoRecords
    Else
        ' La licencia ya existe
    End If

End Sub

Private Sub BorrarLicencia(ByVal NroLic As Long)
      
    ' Borra las Vacaciones
    StrSql = "DELETE FROM lic_vacacion WHERE emp_licnro = " & NroLic
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.writeline "Borrado de lic_vacacion por Reproceso"
    
    ' Borra las Notificaciones
    StrSql = "DELETE FROM vacnotif WHERE emp_licnro = " & NroLic
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.writeline "Borrado de vacnotif por Reproceso"
    
    ' Borra la Licencia
    StrSql = "DELETE FROM emp_lic WHERE emp_licnro = " & NroLic
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.writeline "Borrado de emp_lic por Reproceso"

    ' Borra la Justificaciones
    StrSql = "DELETE FROM gti_justificacion WHERE jussigla = 'LIC' and juscodext = " & NroLic
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.writeline "Borrado de gti_justificacion por Reproceso"

      
End Sub

Private Sub BorrarSituacionRevista(ByVal Ternro As Long, ByVal desde As Date, ByVal hasta As Date)
'-------------------------------------------------------------------------------------------
'
'
'-------------------------------------------------------------------------------------------
Dim rs_His As New ADODB.Recordset
Dim rs_Est As New ADODB.Recordset
Dim rs As New ADODB.Recordset

Dim SitRevista As Long
Dim f_Desde As Date
Dim f_Hasta As Date
Dim f_Anterior As Date
Dim f_Siguiente As Date

Dim f_Sit_Ant
Dim f_Sit_Sig
Dim f_Sit_Sig_hasta
Dim f_Estr_Activo
    
    f_Desde = desde
    f_Hasta = hasta
    f_Anterior = DateAdd("d", -1, CDate(desde))
    f_Siguiente = DateAdd("d", 1, CDate(hasta))
    
    'Busco cual es la estructura de situacion de revista Activo
    StrSql = "SELECT * FROM estructura "
    StrSql = StrSql & " WHERE tenro   = 30 "
    StrSql = StrSql & " AND   estrcodext = '1' "
    If rs_Est.State = adStateOpen Then rs_Est.Close
    OpenRecordset StrSql, rs_Est
    f_Estr_Activo = -1
    If Not rs_Est.EOF Then
        f_Estr_Activo = rs_Est!estrnro
    End If
    
    'Busco la sit de revista del tipo de dia
    StrSql = "SELECT estrnro, tdnro FROM csijp_srtd "
    StrSql = StrSql & " WHERE tdnro = 2"
    If rs.State = adStateOpen Then rs.Close
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        SitRevista = rs!estrnro
    End If
    If rs.State = adStateOpen Then rs.Close
    
    If SitRevista <> 0 Then
        StrSql = "SELECT * FROM his_estructura "
        StrSql = StrSql & " WHERE estrnro = " & SitRevista
        StrSql = StrSql & " AND   tenro   = 30 "
        StrSql = StrSql & " AND   ternro  = " & Ternro
        StrSql = StrSql & " AND   htetdesde  = " & ConvFecha(f_Desde)
        StrSql = StrSql & " AND   htethasta  = " & ConvFecha(f_Hasta)
        If rs_His.State = adStateOpen Then rs_His.Close
        OpenRecordset StrSql, rs_His
    
        'Me fijo si existe la situacion de revista
        If Not rs_His.EOF Then
            'vuelvo a la situacion de revista de activo
            StrSql = "UPDATE his_estructura SET "
            StrSql = StrSql & " estrnro = " & f_Estr_Activo
            StrSql = StrSql & " WHERE estrnro = " & SitRevista
            StrSql = StrSql & " AND   tenro   = 30 "
            StrSql = StrSql & " AND   ternro  = " & Ternro
            StrSql = StrSql & " AND   htetdesde  = " & ConvFecha(f_Desde)
            StrSql = StrSql & " AND   htethasta  = " & ConvFecha(f_Hasta)
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
    End If
    
If rs_His.State = adStateOpen Then rs_His.Close
If rs_Est.State = adStateOpen Then rs_Est.Close
If rs.State = adStateOpen Then rs.Close

Set rs_His = Nothing
Set rs_Est = Nothing
Set rs = Nothing
    
End Sub


Private Sub BorrarSituacionRevista_old(ByVal Ternro As Long, ByVal desde As Date, ByVal hasta As Date)
'-------------------------------------------------------------------------------------------
'
'
'-------------------------------------------------------------------------------------------
Dim rs_His As New ADODB.Recordset
Dim rs_Est As New ADODB.Recordset
Dim rs As New ADODB.Recordset

Dim SitRevista As Long
Dim f_Desde As Date
Dim f_Hasta As Date
Dim f_Anterior As Date
Dim f_Siguiente As Date

Dim f_Sit_Ant
Dim f_Sit_Sig
Dim f_Sit_Sig_hasta
Dim f_Estr_Activo
    
    f_Desde = desde
    f_Hasta = hasta
    f_Anterior = DateAdd("d", -1, CDate(desde))
    f_Siguiente = DateAdd("d", 1, CDate(hasta))
    
    'Busco la sit de revista del tipo de dia
    StrSql = "SELECT estrnro, tdnro FROM csijp_srtd "
    StrSql = StrSql & " WHERE tdnro = 2"
    If rs.State = adStateOpen Then rs.Close
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        SitRevista = rs!estrnro
    End If
    If rs.State = adStateOpen Then rs.Close
    
    If SitRevista <> 0 Then
        StrSql = "SELECT * FROM his_estructura "
        StrSql = StrSql & " WHERE estrnro = " & SitRevista
        StrSql = StrSql & " AND   tenro   = 30 "
        StrSql = StrSql & " AND   ternro  = " & Ternro
        StrSql = StrSql & " AND   htetdesde  = " & ConvFecha(f_Desde)
        StrSql = StrSql & " AND   htethasta  = " & ConvFecha(f_Hasta)
        If rs_His.State = adStateOpen Then rs_His.Close
        OpenRecordset StrSql, rs_His
    
        'Me fijo si existe la situacion de revista
        If Not rs_His.EOF Then
            'Borro la situacion de revista
            StrSql = "DELETE FROM his_estructura "
            StrSql = StrSql & " WHERE estrnro = " & SitRevista
            StrSql = StrSql & " AND   tenro   = 30 "
            StrSql = StrSql & " AND   ternro  = " & Ternro
            StrSql = StrSql & " AND   htetdesde  = " & ConvFecha(f_Desde)
            StrSql = StrSql & " AND   htethasta  = " & ConvFecha(f_Hasta)
            objConn.Execute StrSql, , adExecuteNoRecords
            
            'Busco cual es la estructura de situacion de revista Activo
            StrSql = "SELECT * FROM estructura "
            StrSql = StrSql & " WHERE tenro   = 30 "
            StrSql = StrSql & " AND   estrcodext = '1' "
            If rs_Est.State = adStateOpen Then rs_Est.Close
            OpenRecordset StrSql, rs_Est
            f_Estr_Activo = -1
            If Not rs_Est.EOF Then
                f_Estr_Activo = rs_Est!estrnro
            End If
            
             
            f_Sit_Ant = 0
            If Not rs_Est.EOF Then
                f_Sit_Ant = rs_Est!estrnro
            End If
           
            'Busco el tipo de la situacion de revista siguiente
            StrSql = "SELECT * FROM his_estructura "
            StrSql = StrSql & " WHERE tenro   = 30 "
            StrSql = StrSql & " AND   ternro  = " & Ternro
            StrSql = StrSql & " AND   htetdesde  = " & ConvFecha(f_Siguiente)
            If rs_Est.State = adStateOpen Then rs_Est.Close
            OpenRecordset StrSql, rs_Est
            
            f_Sit_Sig = 0
            If Not rs_Est.EOF Then
                f_Sit_Sig = rs_Est!estrnro
                f_Sit_Sig_hasta = rs_Est!htethasta
            End If
            
            'Re-organizamos la situacion de revista de acuerdo al caso
            'Los casos son 4 y son disjuntos
            '**********************************************************************************
            'Caso 1: la sit. de rev. anterior y siguiente son activas
            If CLng(f_Estr_Activo) = CLng(f_Sit_Ant) And CLng(f_Estr_Activo) = CLng(f_Sit_Sig) Then
                'Borro la situacion de revista siguiente
                StrSql = "DELETE FROM his_estructura "
                StrSql = StrSql & " WHERE tenro   = 30 "
                StrSql = StrSql & " AND   ternro  = " & Ternro
                StrSql = StrSql & " AND   htetdesde  = " & ConvFecha(f_Siguiente)
                objConn.Execute StrSql, , adExecuteNoRecords
    
                If IsNull(f_Sit_Sig_hasta) Then
                    'Modifico la situacion de revista anterior
                    StrSql = "UPDATE his_estructura SET "
                    StrSql = StrSql & "       htethasta = null "
                    StrSql = StrSql & " WHERE tenro   = 30 "
                    StrSql = StrSql & " AND   ternro  = " & Ternro
                    StrSql = StrSql & " AND   htethasta  = " & ConvFecha(f_Anterior)
                Else
                    'Modifico la situacion de revista anterior
                    StrSql = "UPDATE his_estructura SET "
                    StrSql = StrSql & "       htethasta = " & ConvFecha(f_Sit_Sig_hasta)
                    StrSql = StrSql & " WHERE tenro   = 30 "
                    StrSql = StrSql & " AND   ternro  = " & Ternro
                    StrSql = StrSql & " AND   htethasta  = " & ConvFecha(f_Anterior)
                End If
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
            
            '**********************************************************************************
            'Caso 2: la sit. de rev. anterior es activa y la siguiente no
            If CLng(f_Estr_Activo) = CLng(f_Sit_Ant) And CLng(f_Estr_Activo) <> CLng(f_Sit_Sig) Then
                'Modifico la situacion de revista anterior
                StrSql = "UPDATE his_estructura SET "
                StrSql = StrSql & "       htethasta = " & ConvFecha(hasta)
                StrSql = StrSql & " WHERE tenro   = 30 "
                StrSql = StrSql & " AND   ternro  = " & Ternro
                StrSql = StrSql & " AND   htethasta  = " & ConvFecha(f_Anterior)
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
            
            '**********************************************************************************
            'Caso 3: la sit. de rev. anterior no es activa y la siguiente si
            If CLng(f_Estr_Activo) <> CLng(f_Sit_Ant) And CLng(f_Estr_Activo) = CLng(f_Sit_Sig) Then
                'Modifico la situacion de revista siguiente
                StrSql = "UPDATE his_estructura SET "
                StrSql = StrSql & "       htetdesde = " & ConvFecha(desde)
                StrSql = StrSql & " WHERE tenro   = 30 "
                StrSql = StrSql & " AND   ternro  = " & Ternro
                StrSql = StrSql & " AND   htetdesde  = " & ConvFecha(f_Siguiente)
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
            
            '**********************************************************************************
            'Caso 4: la sit. de rev. anterior y siguiente no son activas
            If CLng(f_Estr_Activo) <> CLng(f_Sit_Ant) And CLng(f_Estr_Activo) <> CLng(f_Sit_Sig) Then
                'Modifico la situacion de revista siguiente
                StrSql = "INSERT INTO his_estructura "
                StrSql = StrSql & " (tenro, ternro, estrnro, htetdesde,htethasta) "
                StrSql = StrSql & " VALUES (30, " & Ternro & ", "
                StrSql = StrSql & "" & f_Estr_Activo & ", "
                StrSql = StrSql & ConvFecha(f_Desde) & ", "
                StrSql = StrSql & ConvFecha(f_Hasta) & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
        End If
    End If
    
If rs_His.State = adStateOpen Then rs_His.Close
If rs_Est.State = adStateOpen Then rs_Est.Close
If rs.State = adStateOpen Then rs.Close

Set rs_His = Nothing
Set rs_Est = Nothing
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


Private Sub SepararLicencias_OLD(ByVal tipoVac As Long, ByVal FechaInicial As Date, ByVal Cant As Integer, ByVal Ternro As Long, ByVal NroVac As Long, ByVal Autorizadas As Boolean)
'-----------------------------------------------------------------------
' Procedimiento
'       Divide las licencias en caso de que se incluyan los feriados
'           y en caso de que caiga un feriado en medio
'Autor: FGZ
'Ultima Mod: FGZ - 17/04/2007
'       Se toma un parametro nuevo para configurar si corto las licencias cuando hay un feriado

'-----------------------------------------------------------------------
Dim i As Integer
Dim j As Integer
Dim objFeriado As New Feriado
Dim DHabiles(1 To 7) As Boolean
Dim EsFeriado As Boolean
Dim objRs As New ADODB.Recordset
Dim ExcluyeFeriados As Boolean
Dim Fecha As Date

Dim AuxFechaDesde As Date
Dim AuxFechaHasta As Date
Dim Quedandias As Boolean

Dim CHabiles As Integer
Dim cNoHabiles As Integer
Dim cFeriados As Integer
Dim CortarLicencias As Boolean

    'FGZ - 17/04/2007
    'Por default deberia cortar las licencias
    CortarLicencias = True
    
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
        
        'FGZ - 17/04/2007
        If Not EsNulo(objRs!tpvprog3) Then
            CortarLicencias = CBool(objRs!tpvprog3)
        End If
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
   
   AuxFechaDesde = FechaInicial
   
   Fecha = FechaInicial
    
    Do While i <= Cant
        Quedandias = True
        EsFeriado = objFeriado.Feriado(Fecha, Ternro, False)
        If (EsFeriado) And Not ExcluyeFeriados Then
            'Call InsertarLicencia(AuxFechaDesde, Fecha - 1, 1, CHabiles, Ternro, NroVac)
            'Call InsertarLicencia(AuxFechaDesde, Fecha - 1, 1, (CHabiles + cNoHabiles + cFeriados), Ternro, NroVac)
            'FGZ - 17/04/2007
            If CortarLicencias Then
'                'FGZ - 26/06/2009 -----------------------
'                If PoliticaOK And Diashabiles_LV Then
'                    Call InsertarLicencia(AuxFechaDesde, Fecha - 1, 1, CHabiles + cNoHabiles, Ternro, NroVac, Autorizadas)
'                Else
'                    Call InsertarLicencia(AuxFechaDesde, Fecha - 1, 1, CHabiles, Ternro, NroVac, Autorizadas)
'                End If
                Call InsertarLicencia(AuxFechaDesde, Fecha - 1, 1, CHabiles, Ternro, NroVac, Autorizadas)
                'FGZ - 26/06/2009 -----------------------
                
                CHabiles = 0
                cNoHabiles = 0
                cFeriados = 0
            
                AuxFechaDesde = DateAdd("d", 1, Fecha)
                Quedandias = False
            Else
                cFeriados = cFeriados + 1
                'FGZ - 24/06/2009 -----------------------
                If PoliticaOK And Diashabiles_LV Then
                    i = i + 1
                End If
                'FGZ - 24/06/2009 -----------------------
            End If
        Else
            If DHabiles(Weekday(Fecha)) Or (EsFeriado And ExcluyeFeriados) Then
                i = i + 1
                If DHabiles(Weekday(Fecha)) Then
                    CHabiles = CHabiles + 1
                End If
            Else
                cNoHabiles = cNoHabiles + 1
                cFeriados = cFeriados + 1
                'FGZ - 24/06/2009 -----------------------
                If PoliticaOK And Diashabiles_LV Then
                    i = i + 1
                End If
                'FGZ - 24/06/2009 -----------------------
            End If
        End If
        
        If i < Cant Then
            Fecha = DateAdd("d", 1, Fecha)
        Else
            i = i + 1
        End If
    Loop
    
    If Quedandias Then
'        'FGZ - 26/06/2009 -----------------------
'        If PoliticaOK And Diashabiles_LV Then
'            Call InsertarLicencia(AuxFechaDesde, Fecha, 0, CHabiles + cNoHabiles, Ternro, NroVac, Autorizadas)
'        Else
'            Call InsertarLicencia(AuxFechaDesde, Fecha, 0, CHabiles, Ternro, NroVac, Autorizadas)
'        End If
        Call InsertarLicencia(AuxFechaDesde, Fecha, 0, CHabiles, Ternro, NroVac, Autorizadas)
        'FGZ - 26/06/2009 -----------------------
    End If
    
    Set objFeriado = Nothing
    

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

'Function GenerarSituacionRevista() As Boolean
'    ' Lisandro Moro
'    'AGD - Si se configura la politica 1509 para que no se genere la situacion de revista.
'
'    Dim sql As String
'    Dim salida As Boolean
'    Dim rs_actv As New ADODB.Recordset
'
'    StrSql = " SELECT alcpolestado, cabpolestado "
'    StrSql = StrSql & " FROM gti_cabpolitica "
'    StrSql = StrSql & " INNER JOIN gti_alcanpolitica ON gti_cabpolitica.cabpolnro = gti_alcanpolitica.cabpolnro "
'    StrSql = StrSql & " INNER JOIN gti_detpolitica ON gti_alcanpolitica.detpolnro = gti_detpolitica.detpolnro "
'    StrSql = StrSql & " WHERE gti_cabpolitica.cabpolnivel = 1509 "
'    StrSql = StrSql & " AND alcpolestado = -1 "
'    StrSql = StrSql & " AND cabpolestado = -1 "
'    OpenRecordset StrSql, rs_actv
'    If rs_actv.EOF Then
'        salida = True
'    Else
'        salida = False
'    End If
'    rs_actv.Close
'    GenerarSituacionRevista = salida
'
'End Function

Private Sub SepararLicencias(ByVal tipoVac As Long, ByVal FechaInicial As Date, ByVal Cant As Integer, ByVal Ternro As Long, ByVal NroVac As Long, ByVal Autorizadas As Boolean)
'-----------------------------------------------------------------------
' Procedimiento
'       Divide las licencias en caso de que se incluyan los feriados
'           y en caso de que caiga un feriado en medio
'Autor: FGZ
'Ultima Mod: FGZ - 17/04/2007
'       Se toma un parametro nuevo para configurar si corto las licencias cuando hay un feriado
'Ultima Mod: FGZ - 10/08/2010
'       Se agregó una nueva marca de feriado laborable
'-----------------------------------------------------------------------
Dim i As Integer
Dim j As Integer
Dim objFeriado As New Feriado
Dim DHabiles(1 To 7) As Boolean
Dim EsFeriado As Boolean
Dim objRs As New ADODB.Recordset
Dim ExcluyeFeriadosNoLab As Boolean
Dim ExcluyeFeriadosLab As Boolean
Dim Fecha As Date

Dim AuxFechaDesde As Date
Dim AuxFechaHasta As Date
Dim Quedandias As Boolean

Dim CHabiles As Integer
Dim cNoHabiles As Integer
Dim cFeriados As Integer
Dim CortarLicencias As Boolean
Dim primercorte As Boolean
    'FGZ - 17/04/2007
    'Por default deberia cortar las licencias
    CortarLicencias = True
    primercorte = True
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
    
        ExcluyeFeriadosNoLab = CBool(objRs!tpvferiado)
        
        If Not EsNulo(objRs!excferilab) Then
            ExcluyeFeriadosLab = CBool(objRs!excferilab)
        Else
            ExcluyeFeriadosLab = False
        End If
        
        'FGZ - 17/04/2007
        If Not EsNulo(objRs!tpvprog3) Then
            CortarLicencias = CBool(objRs!tpvprog3)
        End If
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
   
   AuxFechaDesde = FechaInicial
   
   Fecha = FechaInicial
    
    Do While i <= Cant
        Quedandias = True
        EsFeriado = objFeriado.Feriado(Fecha, Ternro, False)
        
        'FGZ - 10/08/2010 -----------------------
        If Not EsFeriado Then
            If DHabiles(Weekday(Fecha)) Then
                i = i + 1
                CHabiles = CHabiles + 1
            Else
                cNoHabiles = cNoHabiles + 1
                If PoliticaOK And Diashabiles_LV Then
                    i = i + 1
                End If
            End If
        Else    'Es feriado
            If Feriado_Laborable Then
                'inicio FAF - 07/10/2013 - Se agrego el NOT en la siguiente linea
                If Not ExcluyeFeriadosLab Then
                'fin FAF - 07/10/2013 - Se agrego el NOT en la siguiente linea
                    If DHabiles(Weekday(Fecha)) Then
                        i = i + 1
                        CHabiles = CHabiles + 1
                    Else
                        cNoHabiles = cNoHabiles + 1
                        If PoliticaOK And Diashabiles_LV Then
                            i = i + 1
                        End If
                    End If
                Else
                    cFeriados = cFeriados + 1
                    If PoliticaOK And Diashabiles_LV Then
                        i = i + 1
                    End If
                End If
            Else 'Feriado no laborable
                'inicio FAF - 07/10/2013 - Se agrego el NOT en la siguiente linea
                If ExcluyeFeriadosNoLab Then 'mdf se le saco el not (cuando la opcion esta tildada en el asp quiere decir que perdi el feriado)
                'fin FAF - 07/10/2013 - Se agrego el NOT en la siguiente linea
                    If DHabiles(Weekday(Fecha)) Then
                        i = i + 1
                        CHabiles = CHabiles + 1
                    Else
                        cNoHabiles = cNoHabiles + 1
                        If PoliticaOK And Diashabiles_LV Then
                            i = i + 1
                        End If
                    End If
                Else
                    ' 04/11/2013 - Se agrego condicion i>0 y se le suma 1 a la variable AuxFechaDesde
                    'mdf 09/01/2014 - la primera vez no debe sumar 1
                    If CortarLicencias And i > 0 Then
                        
                       If (primercorte) Then 'mdf
                        Call InsertarLicencia(AuxFechaDesde, Fecha - 1, 1, CHabiles, Ternro, NroVac, Autorizadas)
                        primercorte = False
                       Else   'mdf
                        Call InsertarLicencia(AuxFechaDesde + 1, Fecha - 1, 1, CHabiles, Ternro, NroVac, Autorizadas)
                       End If
                        
                        CHabiles = 0
                        cNoHabiles = 0
                        cFeriados = 0
                    
                        AuxFechaDesde = DateAdd("d", 1, Fecha)
                        Quedandias = False
                    Else
                        cFeriados = cFeriados + 1
                        If PoliticaOK And Diashabiles_LV Then
                            i = i + 1
                        End If
                    End If
                End If
            End If
        End If
        'FGZ - 10/08/2010 -----------------------
        
        
'        If (EsFeriado) And (Not ExcluyeFeriadosNoLab) Then
'            If CortarLicencias Then
'                Call InsertarLicencia(AuxFechaDesde, Fecha - 1, 1, CHabiles, Ternro, NroVac, Autorizadas)
'
'                CHabiles = 0
'                cNoHabiles = 0
'                cFeriados = 0
'
'                AuxFechaDesde = DateAdd("d", 1, Fecha)
'                Quedandias = False
'            Else
'                cFeriados = cFeriados + 1
'                'FGZ - 24/06/2009 -----------------------
'                If PoliticaOK And Diashabiles_LV Then
'                    i = i + 1
'                End If
'                'FGZ - 24/06/2009 -----------------------
'            End If
'        Else
'            If DHabiles(Weekday(Fecha)) Or (EsFeriado And ExcluyeFeriadosNoLab) Then
'                i = i + 1
'                If DHabiles(Weekday(Fecha)) Then
'                    CHabiles = CHabiles + 1
'                End If
'            Else
'                cNoHabiles = cNoHabiles + 1
'                cFeriados = cFeriados + 1
'                'FGZ - 24/06/2009 -----------------------
'                If PoliticaOK And Diashabiles_LV Then
'                    i = i + 1
'                End If
'                'FGZ - 24/06/2009 -----------------------
'            End If
'        End If
        
        If i < Cant Then
            Fecha = DateAdd("d", 1, Fecha)
        Else
            i = i + 1
        End If
    Loop
    
    If Quedandias Then
        Call InsertarLicencia(AuxFechaDesde, Fecha, 0, CHabiles, Ternro, NroVac, Autorizadas)
    End If
    
    Set objFeriado = Nothing
End Sub


Function generarFirma(ByVal cystipnro As Integer, ByVal emp_licnro As Long, ByVal Inserto As Boolean)
 Dim rs_proc As New ADODB.Recordset
 Dim rsFir As New ADODB.Recordset
 Dim rs_Empleado As New ADODB.Recordset
 Dim Firmas As Integer
 Dim rs As New ADODB.Recordset
    
    Flog.writeline "Genera la Firma para la licencia: " & emp_licnro
    
    'Setea las variables segun el estado de circuito
    If FIN_Firma_Licencias Then
        cysfirusuario = User_Proceso
        cysfirautoriza = User_Proceso
        cysfirdestino = Firma_User_Destino
        cysfirfin = -1
        cysfiryaaut = -1
        cysfirrecha = 0
    Else
        cysfirusuario = User_Proceso
        cysfirautoriza = User_Proceso
        cysfirdestino = Firma_User_Destino
        cysfirfin = 0
        cysfiryaaut = 0
        cysfirrecha = 0
    End If
    
    
    
    If Firma_Licencias Then
        If FIN_Firma_Licencias Then
            If Inserto Then
                Flog.writeline "Inserto una nueva Firma. Fin de Firmas. Usuario" & cysfirautoriza
                
                'Inserto firma autorizado final
                StrSql = "INSERT INTO cysfirmas ("
                StrSql = StrSql & "cysfirautoriza,cysfirfecaut,cysfirmhora,cysfirusuario,"
                StrSql = StrSql & "cysfirdestino,cystipnro,cysfircodext,cysfirsecuencia,cysfirdes"
                StrSql = StrSql & ",cysfirfin,cysfiryaaut,cysfirrecha"
                StrSql = StrSql & ")"
                StrSql = StrSql & "VALUES("
                StrSql = StrSql & "'" & cysfirautoriza & "'," & ConvFecha(Date) & ",'" & FormatDateTime(Now(), vbShortTime) & "'"
                StrSql = StrSql & ",'" & cysfirusuario & "','" & cysfirdestino & "'," & cystipnro & "," & emp_licnro & ",1,'Licencias'"
                StrSql = StrSql & "," & cysfirfin & "," & cysfiryaaut & "," & cysfirrecha
                StrSql = StrSql & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
            Else
                Flog.writeline "Actualizo la Firma. Fin de Firmas. Usuario" & cysfirautoriza
                'Busco la ultima y actualizo
                StrSql = "SELECT cysfirautoriza, cysfirsecuencia, cysfirdestino FROM cysfirmas "
                StrSql = StrSql & " WHERE cysfirmas.cystipnro = " & cystipnro & " AND cysfirmas.cysfircodext = '" & emp_licnro & "' "
                StrSql = StrSql & " ORDER BY cysfirsecuencia DESC"
                OpenRecordset StrSql, rs
                    If Not rs.EOF Then
                        StrSql = "UPDATE cysfirmas "
                        StrSql = StrSql & "SET cysfirdestino = '" & cysfirdestino & "'"
                        StrSql = StrSql & ", cysfirautoriza = '" & cysfirusuario & "'"
                        StrSql = StrSql & ", cysfirfecaut = " & ConvFecha(Date)
                        StrSql = StrSql & ", cysfirmhora = '" & FormatDateTime(Now(), vbShortTime) & "'"
                        StrSql = StrSql & ", cysfirfin = -1,cysfiryaaut = -1, cysfirrecha = 0"
                        StrSql = StrSql & " where cystipnro = " & cystipnro
                        StrSql = StrSql & " and cysfircodext = '" & emp_licnro & "' "
                        StrSql = StrSql & " and cysfirsecuencia = " & rs!cysfirsecuencia
                        objConn.Execute StrSql, , adExecuteNoRecords
                    Else
                        Flog.writeline "Inserto una nueva Firma. Fin de Firmas. Usuario" & cysfirautoriza
                        StrSql = "INSERT INTO cysfirmas ("
                        StrSql = StrSql & "cysfirautoriza,cysfirfecaut,cysfirmhora,cysfirusuario,"
                        StrSql = StrSql & "cysfirdestino,cystipnro,cysfircodext,cysfirsecuencia,cysfirdes"
                        StrSql = StrSql & ",cysfirfin,cysfiryaaut,cysfirrecha"
                        StrSql = StrSql & ")"
                        StrSql = StrSql & "VALUES("
                        StrSql = StrSql & "'" & cysfirautoriza & "'," & ConvFecha(Date) & ",'" & FormatDateTime(Now(), vbShortTime) & "'"
                        StrSql = StrSql & ",'" & cysfirusuario & "','" & cysfirdestino & "'," & cystipnro & "," & emp_licnro & ",1,'Licencias'"
                        StrSql = StrSql & "," & cysfirfin & "," & cysfiryaaut & "," & cysfirrecha
                        StrSql = StrSql & ")"
                        objConn.Execute StrSql, , adExecuteNoRecords
                    End If
                End If
            Else
                If Not EsNulo(Firma_User_Destino) Then
                    If Inserto Then
                        Flog.writeline "Inserto una nueva Firma. De " & cysfirautoriza & " para " & cysfirdestino
                        
                        'Inserto firma autorizado pendiente
                        StrSql = "INSERT INTO cysfirmas ("
                        StrSql = StrSql & "cysfirautoriza,cysfirfecaut,cysfirmhora,cysfirusuario,"
                        StrSql = StrSql & "cysfirdestino,cystipnro,cysfircodext,cysfirsecuencia,cysfirdes"
                        StrSql = StrSql & ",cysfirfin,cysfiryaaut,cysfirrecha"
                        StrSql = StrSql & ")"
                        StrSql = StrSql & "VALUES("
                        StrSql = StrSql & "'" & cysfirautoriza & "'," & ConvFecha(Date) & ",'" & FormatDateTime(Now(), vbShortTime) & "'"
                        StrSql = StrSql & ",'" & cysfirusuario & "','" & cysfirdestino & "'," & cystipnro & "," & emp_licnro & ",1,'Licencias'"
                        StrSql = StrSql & "," & cysfirfin & "," & cysfiryaaut & "," & cysfirrecha
                        StrSql = StrSql & ")"
                        objConn.Execute StrSql, , adExecuteNoRecords
                    Else
                        'Busco la ultima y actualizo
                        Flog.writeline "Actualiza una Firma. De " & cysfirautoriza & " para " & cysfirdestino
                        
                        StrSql = "SELECT cysfirautoriza, cysfirsecuencia, cysfirdestino FROM cysfirmas "
                        StrSql = StrSql & " WHERE cysfirmas.cystipnro = " & cystipnro & " AND cysfirmas.cysfircodext = '" & emp_licnro & "' "
                        StrSql = StrSql & " ORDER BY cysfirsecuencia DESC"
                        OpenRecordset StrSql, rs
                        If Not rs.EOF Then
                            StrSql = "UPDATE cysfirmas "
                            StrSql = StrSql & "SET cysfirdestino = '" & cysfirdestino & "'"
                            StrSql = StrSql & ", cysfirautoriza = '" & cysfirusuario & "'"
                            StrSql = StrSql & ", cysfirfecaut = " & ConvFecha(Date)
                            StrSql = StrSql & ", cysfirmhora = '" & FormatDateTime(Now(), vbShortTime) & "'"
                            StrSql = StrSql & ", cysfirfin = 0,cysfiryaaut = 0, cysfirrecha = 0"
                            StrSql = StrSql & " where cystipnro = " & cystipnro
                            StrSql = StrSql & " and cysfircodext = '" & emp_licnro & "' "
                            StrSql = StrSql & " and cysfirsecuencia = " & rs!cysfirsecuencia
                            objConn.Execute StrSql, , adExecuteNoRecords
                        Else
                            Flog.writeline "Inserto una Firma. De " & cysfirautoriza & " para " & cysfirdestino
                            
                            StrSql = "INSERT INTO cysfirmas ("
                            StrSql = StrSql & "cysfirautoriza,cysfirfecaut,cysfirmhora,cysfirusuario,"
                            StrSql = StrSql & "cysfirdestino,cystipnro,cysfircodext,cysfirsecuencia,cysfirdes"
                            StrSql = StrSql & ",cysfirfin,cysfiryaaut,cysfirrecha"
                            StrSql = StrSql & ")"
                            StrSql = StrSql & "VALUES("
                            StrSql = StrSql & "'" & cysfirautoriza & "'," & ConvFecha(Date) & ",'" & FormatDateTime(Now(), vbShortTime) & "'"
                            StrSql = StrSql & ",'" & cysfirusuario & "','" & cysfirdestino & "'," & cystipnro & "," & emp_licnro & ",1,'Licencias'"
                            StrSql = StrSql & "," & cysfirfin & "," & cysfiryaaut & "," & cysfirrecha
                            StrSql = StrSql & ")"
                            objConn.Execute StrSql, , adExecuteNoRecords
                        End If
                    End If
                Else
                    'Error no se puede insertar
                    Flog.writeline Espacios(Tabulador * 2) & "No se puede pasar la novedad, el usuario destino de firma es nulo. "
                End If
            End If
        End If
    
    
    
    
    
    
    
    
    
'
'
'    'SI EL CIRCUITO ESTA ACTIVO
'    If Firmas = "OK" Then
''        'Seteo todo en 0
''        cysfirusuario = ""
''        cysfirautoriza = ""
''        cysfirdestino = ""
''        cysfirfin = 0
''        cysfiryaaut = 0
''        cysfirrecha = 0
''        tipoorigen = ""
''        l_listperfnro = ""
'
'
'
'        'Verifica si es fin de firma
'        StrSql = "SELECT * FROM cysfincirc WHERE userid = '" & usuario & "' and cystipnro = " & cystipnro
'        OpenRecordset StrSql, rs_Empleado
'
'        If Not rs_Empleado.EOF Then
'    '        Esfin = True
'    '        cysfirusuario = usuario
'    '        cysfirautoriza = usuario
'    '        cysfirdestino = ""
'    '        cysfirfin = -1
'    '        cysfiryaaut = -1
'    '        cysfirrecha = 0
'        Else
'            Esfin = False
'        End If
'        rs_Empleado.Close
'
'    If Esfin = False Then
'        '=====================================
'        'QUE TENGA DELEGADO UN PERMISO
'        '=====================================
'        StrSql = "SELECT bk_cab.iduser, bkcystipnro "
'        StrSql = StrSql & " From bk_cab "
'        StrSql = StrSql & " INNER JOIN bk_firmas on bk_firmas.bkcabnro = bk_cab.bkcabnro "
'        StrSql = StrSql & " Where fdesde <= " & ConvFecha(Date)
'        StrSql = StrSql & " AND (fhasta >= " & ConvFecha(Date) & " OR fhasta IS NULL)"
'        StrSql = StrSql & " AND bk_firmas.iduser = '" & usuario & "'"
'        StrSql = StrSql & " AND bkcystipnro = " & cystipnro
'        StrSql = StrSql & " AND bk_cab.iduser <> '" & usuario & "'"
'        OpenRecordset StrSql, rs_Empleado
'
'        If Not rs_Empleado.EOF Then
'            Esfin = True
'            cysfirusuario = rs_Empleado!IdUser
'            cysfirautoriza = usuario
'            cysfirdestino = ""
'
'            cysfirfin = -1
'            cysfiryaaut = -1
'            cysfirrecha = 0
'        Else
'            '-----------------------------
'            'QUE EXISTA EL USUARIO
'            '-----------------------------
'            StrSql = "SELECT iduser FROM user_per "
'            StrSql = StrSql & "WHERE iduser='" & id_autoriz & "'"
'            OpenRecordset StrSql, rs_Estado
'            If rs_Estado.EOF Or UCase(usuario) = UCase(id_autoriz) Then
'                Texto = ": " & "No se encontro el usuario " & id_autoriz
'                Nrocolumna = 1
'                Call Escribir_Log("floge", NroLinea, Nrocolumna, Texto, Tabs, strLinea)
'                Call InsertaError(1, 128)
'                HuboError = True
'                Exit Function
'           End If
'            rs_Estado.Close
'            '---------------------------
'            cysfirusuario = usuario
'            cysfirautoriza = usuario
'            cysfirdestino = id_autoriz
'
'            cysfirfin = 0
'            cysfiryaaut = 0
'            cysfirrecha = 0
'        End If
'        rs_Empleado.Close
'
'
'    End If
'    '-----
'End If
'
'           ' ====================================================================
'           'INSERTO EN cysfirmas
'           ' ====================================================================
'            If Firmas = "OK" And emp_licnro <> 0 Then
'                StrSql = "INSERT INTO cysfirmas ("
'                StrSql = StrSql & "cysfirautoriza,cysfirfecaut,cysfirmhora,cysfirusuario,"
'                StrSql = StrSql & "cysfirdestino,cystipnro,cysfircodext,cysfirsecuencia,cysfirdes"
'                StrSql = StrSql & ",cysfirfin,cysfiryaaut,cysfirrecha"
'                StrSql = StrSql & ")"
'                StrSql = StrSql & "VALUES("
'                StrSql = StrSql & "'" & cysfirautoriza & "'," & ConvFecha(Date) & ",'" & FormatDateTime(Now(), vbShortTime) & "'"
'                StrSql = StrSql & ",'" & cysfirusuario & "','" & cysfirdestino & "'," & cystipnro & "," & emp_licnro & ",1,'Licencia desde interfaz'"
'                StrSql = StrSql & "," & cysfirfin & "," & cysfiryaaut & "," & cysfirrecha
'                StrSql = StrSql & ")"
'                objConn.Execute StrSql, , adExecuteNoRecords
'                Texto = "Autorización insertada "
'                Call Escribir_Log("flogp", NroLinea, 1, Texto, Tabs, strLinea)
'            End If

End Function

