Attribute VB_Name = "mdlDiasCorresp"
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
'Global fecha_desde As Date
'Global fecha_hasta As Date
'Global Periodo_Anio As Long
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
Global generarperiodossv  As Boolean  'MDF



Public Sub Main()
Dim Fecha As Date
Dim cantdias As Integer
Dim cantdiasCorr As Integer
Dim CantdiasCR As Double
Dim DiasCorraGen As Double
Dim dias_maternidad As Integer
Dim Columna As Integer
Dim columna2 As Integer
Dim Mensaje As String
Dim Genera As Boolean
Dim NroTPV As String
Dim NroTPVCorr As String
Dim AnioaProc As Integer
Dim TodosEmpleados As Boolean
Dim auxNroVac As Long
Dim NroVacAnterior As Long
Dim strparametros As String
Dim ArrPar

Dim pos1 As Integer
Dim pos2 As Integer

Dim objReg As New ADODB.Recordset
Dim strCmdLine As String
'Dim objconnMain As New ADODB.Connection
Dim Archivo As String

Dim rs As New ADODB.Recordset
Dim rs_Batch_Proceso As New ADODB.Recordset
Dim rs_His_Batch_Proceso As New ADODB.Recordset
Dim PID As String
Dim ArrParametros

Dim cantDiasTrabajado As Integer
Dim regHorarioActual As Integer
Dim fechaAlta As Date
Dim bprcfechasta As Date   'EAM- Fecha Hasta donde se calcula los días correspondientes
Dim bprcfechastaAux As Date   'EAM- Guarda la fecha hasta del proceso, se usa para los procesos planificados syke
Dim fechaBaja As Date
Dim vac_alcannivel As Integer 'NG - Alcance de vacaciones. (Vacas con pol 1515)
Dim Usa1515 As Boolean 'NG - (Vacas con pol 1515)
Dim generarRegVacDiasCor As Boolean
generarRegVacDiasCor = False



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
    Archivo = PathFLog & "Vac_DiasCorresp" & "-" & NroProceso & ".log"
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Archivo, True)
    
    'Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "------------------------------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaVersion
    Flog.writeline "PID                      : " & PID
    Flog.writeline "------------------------------------------------------------------------"
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
    
    'Activo el manejador de errores
    On Error GoTo CE
    
    'EAM - 18/01/2012 --------- Control de versiones ------
    'Version_Valida = ValidarV(Version, 10, TipoBD) 'NG - Ahora se valida desde la funcion ValidarVBD
    
    
    
    '*******************************************************************************************************
    '--------------- VALIDO MODELOS SEGUN POLITICA 1515 | PUEDE TENER ALCANCE POR ESTRUCTURAS --------------
    '*******************************************************************************************************
    Version_Valida = ValidaModeloyVersiones(Version, 10)
    If (Version_Valida = False) Then
        'SI NO ESTA ACTIVA LA 1515 O NO EXISTE CONFIGURACIÓN, TOMA DEFAULT
        modeloPais = Pais_Modelo(7)
        Version_Valida = ValidarV(Version, 10, TipoBD)
        Usa1515 = False
    Else
        Usa1515 = True
    End If
    
    '___________________________________________________________________________________________________
    'Se agrego un parametro para determinar el país con el cual se quiere calcular el modelo de vacacion
    'El parametro 7 hace referencia al modelo de vacaciones configurado en la tabla confper
   ' modeloPais = Pais_Modelo(7)
    'NG - 04/05/2012 --------- Control de versiones ------
     'Version_Valida = ValidarVBD(Version, 10, TipoBD, modeloPais)
    
    
        
    If Not Version_Valida Then
        'Actualizo el progreso
        MyBeginTrans
            StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Error de Version', bprcpid = " & PID & " WHERE bpronro = " & NroProceso
            objConnProgreso.Execute StrSql, , adExecuteNoRecords
        MyCommitTrans
        Flog.writeline
        GoTo Final
    End If
    
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprcpid = " & PID & ", bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objConnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Levanta Proceso y Setea Parámetros:  " & " " & Now
       
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro =" & NroProceso
    OpenRecordset StrSql, rs_Batch_Proceso
       
    If rs_Batch_Proceso.EOF Then Exit Sub
    
    '____________________________________________________________
    'NG - VALIDA QUE ESTE ACTIVO LA TRADUCCION A MULTI IDIOMA
    usuario = rs_Batch_Proceso!iduser
    Call Valida_MultiIdiomaActivo(usuario)
    
    '------------------------------------------------------------
'    Parametros = rs_Batch_Proceso!bprcparam
'
'    If Not IsNull(Parametros) Then
'        If Len(Parametros) >= 1 Then
'            pos1 = 1
'            pos2 = InStr(pos1, Parametros, ".") - 1
'            NroVac = CLng(Mid(Parametros, pos1, pos2))
'
'            pos1 = pos2 + 2
'            pos2 = InStr(pos1, Parametros, ".") - 1
'            Reproceso = CBool(Mid(Parametros, pos1, pos2 - pos1 + 1))
'
'
'            'FGZ - 25/03/2010 - se le agregaron 2 parametros
'            Flog.writeline "Año del periodo"
'
'            pos1 = pos2 + 2
'            pos2 = InStr(pos1, Parametros, ".") - 1
'            AnioaProc = Mid(Parametros, pos1, pos2 - pos1 + 1)
'
'
'            pos1 = pos2 + 2
'            pos2 = Len(Parametros)
'            TodosEmpleados = CBool(Mid(Parametros, pos1, pos2 - pos1 + 1))
'
'            StrSql = " SELECT * FROM vacacion WHERE vacacion.vacnro = " & NroVac
'            OpenRecordset StrSql, objRs
'            If Not objRs.EOF Then
'                fecha_desde = objRs!vacfecdesde
'                fecha_hasta = objRs!vacfechasta
'                Periodo_Anio = objRs!vacanio
'            Else
'                Exit Sub
'            End If
'
'        End If
'    End If
       
    'FGZ - 25/03/2010 - se le agregaron 2 parametros -------------------
    parametros = rs_Batch_Proceso!bprcparam
    Flog.writeline EscribeLogMI("Parametros") & " " & parametros
    
    ArrPar = Split(parametros, ".")
    
    'NroVac = CLng(ArrPar(0))
    NroVac = IIf(EsNulo(ArrPar(0)), 0, CLng(ArrPar(0)))

    'Reproceso = CBool(ArrPar(1))
    Reproceso = IIf(EsNulo(ArrPar(1)), 0, CBool(ArrPar(1)))
    
    If UBound(ArrPar) > 1 Then
        AnioaProc = IIf(EsNulo(ArrPar(2)), 0, ArrPar(2))
        
        If UBound(ArrPar) > 2 Then
            If EsNulo(ArrPar(3)) Then
                TodosEmpleados = False
            Else
                'TodosEmpleados = CBool(ArrPar(3))
                TodosEmpleados = IIf(EsNulo(ArrPar(3)), False, CBool(ArrPar(3)))
            End If
        Else
            AnioaProc = 0
            TodosEmpleados = False
        End If
    Else
        AnioaProc = 0
        TodosEmpleados = False
    End If
    
    'Obtiene la fecha Hasta donde se va a calcular los días correspondientes - CR
    If EsNulo(rs_Batch_Proceso!bprcfechasta) Then
        bprcfechasta = Date
        bprcfechastaAux = Date
    Else
        bprcfechasta = rs_Batch_Proceso!bprcfechasta
        bprcfechastaAux = rs_Batch_Proceso!bprcfechasta
    End If
   
    'FGZ - 25/03/2010 - se le agregaron 2 parametros -------------------
    
    'EAM- 11/08/2010 - Se agrego un parametro para determinar el país con el cual se quiere calcular el modelo de vacacion
    'El parametro 7 hace referencia al modelo de vacaciones configurado en la tabla confper
    'modeloPais = Pais_Modelo(7)
    

    
    Set objFechasHoras.Conexion = objConn
    
    If TodosEmpleados Then
        'EAM (v3.28)- Se modifico la sql para que además de los activos tome los inactivos con la fecha de generacion de días corresp. menores a la fecha de cierre de la fase.
        StrSql = "SELECT distinct empleado.ternro FROM empleado WHERE empest= -1" & _
            " UNION " & _
            " SELECT distinct e.ternro FROM empleado e " & _
            " INNER JOIN fases f ON f.empleado = e.ternro AND f.bajfec IS NOT NULL " & _
            " LEFT JOIN vacdiascor v ON v.ternro = e.ternro " & _
            " WHERE empest=0 AND (( " & _
            " f.altfec <= (SELECT MAX(vdiasfechasta) from vacdiascor WHERE ternro= e.ternro AND venc=0) " & _
            " AND (bajfec> (SELECT MAX(vdiasfechasta) from vacdiascor WHERE ternro= e.ternro AND venc=0) OR f.bajfec is null)) " & _
            " OR v.vdiasfechasta is NULL) "
        OpenRecordset StrSql, objReg
        
'        StrSql = " SELECT * FROM batch_empleado WHERE batch_empleado.bpronro = " & NroProceso
        
'
'        If objReg.EOF Then
'            objReg.Close
'            StrSql = "SELECT distinct empleado.ternro FROM empleado WHERE empest= -1"
'            OpenRecordset StrSql, objReg
'        End If
    Else
        StrSql = " SELECT * FROM batch_empleado WHERE batch_empleado.bpronro = " & NroProceso
        OpenRecordset StrSql, objReg
    End If

    
    CEmpleadosAProc = objReg.RecordCount
    If CEmpleadosAProc = 0 Then
        Flog.writeline "No hay empleados seleccionados"
        Exit Sub
    End If
    IncPorc = (100 / CEmpleadosAProc)
    
    SinError = True
    HuboErrores = False
    Do While Not objReg.EOF
   
        Ternro = objReg!Ternro
        Empleado.Ternro = objReg!Ternro
        
        Flog.writeline ""
        Flog.writeline "========================================================================"
        Flog.writeline EscribeLogMI("Inicio Empleado") & ": " & Ternro
        
        '------------------------------------------------------------
        ' SI USA CONFIGURACION DE LA POLITICA 1515
        '------------------------------------------------------------
        If Usa1515 = True Then
             If NroVac <> 0 Then
                 StrSql = " SELECT * FROM vacacion WHERE vacacion.vacnro = " & NroVac
                 OpenRecordset StrSql, objRs
                 If Not objRs.EOF Then
                     fecha_desde = objRs!vacfecdesde
                     fecha_hasta = objRs!vacfechasta
                 End If
                 objRs.Close
             Else
                 'levanta la fecha del proceso.
                 fecha_desde = bprcfechasta
                 fecha_hasta = bprcfechastaAux
             End If
             
             
            
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
        '---------------------------------------------------------------------------------------------------------
        ' FIN 1515
        '---------------------------------------------------------------------------------------------------------

   
        MyBeginTrans
        
        'lisandro moro - si no tengo el nrovac lo genero, se movio el codigo a esta zona para obtener el ternro
        Select Case modeloPais
            Case 3: 'Colombia
                Flog.writeline "Busco el periodo para colombia."
                StrSql = " SELECT * FROM vacacion WHERE vacacion.vacanio = " & AnioaProc & " AND ternro = " & Ternro
                OpenRecordset StrSql, objRs
                If Not objRs.EOF Then
                    fecha_desde = objRs!vacfecdesde
                    fecha_hasta = objRs!vacfechasta
                    Periodo_Anio = objRs!vacanio
                    NroVac = objRs!vacnro
                Else
                    Flog.writeline "Genero el periodo para colombia."
                    Call generarPeriodoVacacion(Ternro, AnioaProc, modeloPais)
                    
                    Flog.writeline "Vuelvo a buscar el periodo para colombia."
                    StrSql = " SELECT * FROM vacacion WHERE vacacion.vacanio = " & AnioaProc & " AND ternro = " & Ternro
                    OpenRecordset StrSql, objRs
                    If Not objRs.EOF Then
                        fecha_desde = objRs!vacfecdesde
                        fecha_hasta = objRs!vacfechasta
                        Periodo_Anio = objRs!vacanio
                        NroVac = objRs!vacnro
                    End If
                End If
                auxNroVac = NroVac
            Case 4: 'Costa Rica
            
                'EAM (v5.26) - Se modifico para que siempre se procese al domingo anterior a la fecha de procesamiento
                bprcfechasta = DateAdd("d", 1 - Weekday(bprcfechastaAux), bprcfechastaAux)
                
                
                Flog.writeline "Busco el periodo para Costa Rica."
                fechaAlta = FechaAltaEmpleado(Ternro)
                Flog.writeline "Fecha alta empleado:" & fechaAlta
                'EAM- Si el empleado no tiene fase, pasa al siguiente empleado
                If (fechaAlta = Empty) Then
                    GoTo siguiente
                End If
                                
                Flog.writeline "Fecha Procesamiento: " & bprcfechasta & "."
                                                     
                'Agregado el 18/10/2013
                'EAM- Obtiene el año que se va a procesar a partir de la fase del empleado.
                If bprcfechasta < CDate(Day(fechaAlta) & "/" & Month(fechaAlta) & "/" & Year(bprcfechasta)) Then
                    AnioaProc = Year(bprcfechasta) - 1
                Else
                    AnioaProc = Year(bprcfechasta)
                End If
                Flog.writeline "Año de Procesar: " & AnioaProc & "."
                
                
                StrSql = " SELECT * FROM vacacion WHERE vacacion.vacanio = " & AnioaProc & " AND ternro = " & Ternro & " ORDER BY vacfecdesde DESC"
                Flog.writeline "busco período: " & StrSql & "."
                OpenRecordset StrSql, objRs
                If Not objRs.EOF Then
                    If (CDate(objRs!vacfecdesde) <> CDate(Day(fechaAlta) & "/" & Month(fechaAlta) & "/" & AnioaProc)) Then
                        Flog.writeline "Encontro período- fase nueva"
                        l_vacnro = objRs!vacnro
                        'busco la fase anterior a la activa y recupero la fecha baja
                        StrSql = " SELECT bajfec "
                        StrSql = StrSql & " FROM fases "
                        StrSql = StrSql & " LEFT JOIN empant ON fases.empantnro=empant.empantnro "
                        StrSql = StrSql & " LEFT JOIN causa ON fases.caunro=causa.caunro "
                        StrSql = StrSql & " WHERE fases.Empleado = " & Ternro
                        StrSql = StrSql & " AND estado =0 "
                        StrSql = StrSql & " ORDER by bajfec DESC "
                        OpenRecordset StrSql, objRs
                        If Not objRs.EOF Then
                            fechaBaja = objRs!bajfec
                        End If
                        
                        If Len(fechaBaja) > 0 Then
                            Flog.writeline "Se encontro Baja para el ternro: " & Ternro
                            StrSql = " SELECT altfec "
                            StrSql = StrSql & " FROM fases "
                            StrSql = StrSql & " LEFT JOIN empant ON fases.empantnro=empant.empantnro "
                            StrSql = StrSql & " LEFT JOIN causa ON fases.caunro=causa.caunro "
                            StrSql = StrSql & " WHERE fases.Empleado = " & Ternro
                            StrSql = StrSql & " AND estado =-1 "
                            StrSql = StrSql & " ORDER by altfec DESC "
                            OpenRecordset StrSql, objRs
                            If Not objRs.EOF Then
                                fechaAlta = objRs!altfec
                            End If
                            
                            'update vacacion del periodo que recupere en la sql (StrSql = " SELECT * FROM vacacion WHERE vacacion.vacanio = " & AnioaProc & " AND ternro = " & Ternro) con la fecha de baja que recupere de la fase cerrada
                            StrSql = "UPDATE vacacion SET vacfechasta = " & ConvFecha(fechaBaja)
                            'StrSql = "UPDATE vacacion SET vacfechasta = " & ConvFecha(l_fechaalta)
                            StrSql = StrSql & " WHERE vacnro = " & l_vacnro
                            StrSql = StrSql & " AND ternro =" & Ternro
                            objConn.Execute StrSql, , adExecuteNoRecords
                            Flog.writeline "Actualizo la fecha Hasta del Último período de vacacion a :" & fechaBaja
                         
                            'NG  21/01/2014 - SE BUSCA ULTIMO REGISTROO DE VACDIASCOR PARA LUEGO ACTUALIZAR FECHA HASTA.
'                            StrSql = "SELECT vacnro FROM vacdiascor WHERE ternro =" & Ternro & " ORDER BY vdiasfechasta DESC"
'                            OpenRecordset StrSql, objRs
'                            If Not objRs.EOF Then
'                                'Actualizo a la misma fecha de la tabla vacación.
'                                 'StrSql = "UPDATE vacdiascor SET vdiasfechasta = '" & l_fechabaja & "'"
'                                 StrSql = "UPDATE vacdiascor SET vdiasfechasta =  " & ConvFecha(l_fechaalta)
'                                 StrSql = StrSql & " Where vacnro = " & objRs!vacnro
'                                 'StrSql = StrSql & " Where vacnro = " & l_vacnro
'                                 StrSql = StrSql & " AND ternro =" & Ternro
'                                 objConn.Execute StrSql, , adExecuteNoRecords
'                                 Flog.writeline "Actualizo la fecha de las tablas vacacion y vacdiascor a :" & l_fechaalta
'                            End If
                          
                          'Flog.writeline "Genero el periodo para Costa Rica. Cuando la fecha desde del periodo no coincide con la fase"
                          Flog.writeline "Genera periodo de vacaciones si Len(fechaBaja) > 0.Ternro=" & Ternro
                          Call generarPeriodoVacacion(Ternro, AnioaProc, modeloPais)
                          generarRegVacDiasCor = True
                          
                          StrSql = " SELECT * FROM vacacion WHERE vacacion.vacanio = " & AnioaProc & " AND ternro = " & Ternro & " ORDER BY vacfecdesde DESC"
                          Flog.writeline "Busco si existe periodo." & StrSql
                          OpenRecordset StrSql, objRs
                          '-------------------------mdf inicio
                          If Not objRs.EOF Then
                            'fecha_desde = objRs!vacfecdesde
                            'fecha_hasta = objRs!vacfechasta
                            'Periodo_Anio = objRs!vacanio
                             NroVac = objRs!vacnro
                        
                             StrSql = "SELECT * FROM vacdiascor " & _
                                " INNER JOIN vacacion on vacacion.vacnro= vacdiascor.vacnro " & _
                                " WHERE vacacion.vacanio= " & AnioaProc & " AND vacacion.ternro= " & Ternro
                                OpenRecordset StrSql, objRs
                                Flog.writeline "Si existe periodo.Nro vac=" & NroVac
                                Flog.writeline "Si existe periodo.busco:" & StrSql
                          '-------------------------mdf fin
                            If objRs.EOF Then 'mdf  inserta si no existe
                              StrSql = "INSERT INTO vacdiascor(vacnro,vdiascorcant,vdiascormanual,ternro,tipvacnro,vdiasfechasta) VALUES ("
                              StrSql = StrSql & NroVac & ",0,0," & Ternro & ",1," & ConvFecha(fechaAlta) & ")"
                              Flog.writeline "Query  si no existe" & StrSql
                              objConn.Execute StrSql, , adExecuteNoRecords
                              Flog.writeline "Genero registro en vacdiascor con fecha de procesamiento :" & fechaAlta
                            End If
                          End If 'mdf
                        End If
                    End If
                End If
                'Fin
                
                'Comentado el 18/10/2013
                
'                'EAM- Obtiene el año que se va a procesar a partir de la fase del empleado.
'                If bprcfechasta < CDate(Day(fechaAlta) & "/" & Month(fechaAlta) & "/" & Year(bprcfechasta)) Then
'                    AnioaProc = Year(bprcfechasta) - 1
'                Else
'                    AnioaProc = Year(bprcfechasta)
'                End If
'                Flog.writeline "Año de Procesar: " & AnioaProc & "."
                'fin
                
                StrSql = " SELECT * FROM vacacion WHERE vacacion.vacanio = " & AnioaProc & " AND ternro = " & Ternro & " ORDER BY vacfecdesde DESC"
                OpenRecordset StrSql, objRs
                Flog.writeline "Busco nuevamente periodo :" & StrSql
                If Not objRs.EOF Then
                    fecha_desde = CDate(objRs!vacfecdesde) 'mdf
                    fecha_hasta = CDate(objRs!vacfechasta) 'mdf
                    Periodo_Anio = objRs!vacanio
                    NroVac = objRs!vacnro
                     Flog.writeline "si existe, busco fase . nro vac:" & NroVac
                    StrSql = " SELECT altfec,bajfec FROM fases " & _
                        " LEFT JOIN empant ON fases.empantnro=empant.empantnro " & _
                        " LEFT JOIN causa ON fases.caunro=causa.caunro " & _
                        " WHERE fases.Empleado = " & Ternro & _
                        " ORDER by altfec DESC "
                    OpenRecordset StrSql, objRs
                    'Flog.writeline "Query. " & StrSql
                    If Not objRs.EOF Then
                        If (Not EsNulo(objRs!bajfec)) Then
                            bprcfechasta = objRs!bajfec
                            Flog.writeline "Seteo "
                        End If
                    End If
                Else
                
                        StrSql = " SELECT bajfec "
                        StrSql = StrSql & " FROM fases "
                        StrSql = StrSql & " LEFT JOIN empant ON fases.empantnro=empant.empantnro "
                        StrSql = StrSql & " LEFT JOIN causa ON fases.caunro=causa.caunro "
                        StrSql = StrSql & " WHERE fases.Empleado = " & Ternro
                        StrSql = StrSql & " AND estado =0 "
                        StrSql = StrSql & " ORDER by bajfec DESC "
                        Flog.writeline "No existe.Query. " & StrSql
                        OpenRecordset StrSql, objRs
                        If Not objRs.EOF Then
                            fechaBaja = objRs!bajfec
                        
                            Flog.writeline "Se encontro Baja para el ternro: " & Ternro
                            StrSql = " SELECT altfec "
                            StrSql = StrSql & " FROM fases "
                            StrSql = StrSql & " LEFT JOIN empant ON fases.empantnro=empant.empantnro "
                            StrSql = StrSql & " LEFT JOIN causa ON fases.caunro=causa.caunro "
                            StrSql = StrSql & " WHERE fases.Empleado = " & Ternro
                            StrSql = StrSql & " AND estado =-1 "
                            StrSql = StrSql & " ORDER by altfec DESC "
                            'Flog.writeline "Query. " & StrSql
                            OpenRecordset StrSql, objRs
                            If Not objRs.EOF Then
                                fechaAlta = objRs!altfec
                            End If
                            
                            StrSql = " SELECT * FROM vacacion WHERE ternro = " & Ternro & " ORDER BY vacfecdesde DESC"
                            OpenRecordset StrSql, objRs
                                                        
                            If Not objRs.EOF Then
                                l_vacnro = objRs!vacnro
                                'update vacacion del periodo que recupere en la sql (StrSql = " SELECT * FROM vacacion WHERE vacacion.vacanio = " & AnioaProc & " AND ternro = " & Ternro) con la fecha de baja que recupere de la fase cerrada
                                StrSql = "UPDATE vacacion SET vacfechasta = " & ConvFecha(fechaBaja)
                                'StrSql = "UPDATE vacacion SET vacfechasta = " & ConvFecha(l_fechaalta)
                                StrSql = StrSql & " WHERE vacnro = " & l_vacnro
                                StrSql = StrSql & " AND ternro =" & Ternro
                                'Flog.writeline "Query. " & StrSql
                                objConn.Execute StrSql, , adExecuteNoRecords
                                Flog.writeline "Actualizo la fecha Hasta del Último período de vacacion a :" & fechaBaja
                            End If
                
                            Flog.writeline "Genero el periodo para Costa Rica.Ternro:" & Ternro
                            Call generarPeriodoVacacion(Ternro, AnioaProc, modeloPais, -1)
                        End If
                
                End If   ' MDF movi el de abajo a aca, ojo q capaz hace cualquiera......!! (1)
                
                
                   
                    
                    Flog.writeline "Vuelvo a buscar el periodo para Costa Rica."
                    StrSql = " SELECT * FROM vacacion WHERE vacacion.vacanio = " & AnioaProc & " AND ternro = " & Ternro
                    StrSql = StrSql & " order by vacfecdesde desc" 'mdf puede quedar una vacacion sin vacdiascor
                    OpenRecordset StrSql, objRs
                    
                     Flog.writeline "Vuelvo a buscar el periodo para Costa Rica.SQL:" & StrSql
                    If Not objRs.EOF Then
                        fecha_desde = CDate(objRs!vacfecdesde)
                        fecha_hasta = CDate(objRs!vacfechasta)
                        Periodo_Anio = objRs!vacanio
                        NroVac = objRs!vacnro
                        
                        StrSql = "SELECT * FROM vacdiascor " & _
                                " INNER JOIN vacacion on vacacion.vacnro= vacdiascor.vacnro " & _
                                " WHERE vacacion.vacanio= " & AnioaProc & " AND vacacion.ternro= " & Ternro & _
                                " and vacdiascor.vacnro=" & NroVac  'mdf
                                
                        OpenRecordset StrSql, objRs
                         Flog.writeline "Si el periodo CR existe:" & StrSql
                        If objRs.EOF Then
                            StrSql = "INSERT INTO vacdiascor(vacnro,vdiascorcant,vdiascormanual,ternro,tipvacnro,vdiasfechasta) VALUES ("
                            StrSql = StrSql & NroVac & ",0,0," & Ternro & ",1,"
                            StrSql = StrSql & ConvFecha(fechaAlta) & ")"
                            Flog.writeline "Query : " & StrSql
                            objConn.Execute StrSql, , adExecuteNoRecords
                            Flog.writeline "Genero registro en vacdiascor con fecha de procesamiento :" & fechaAlta
                        End If
                    Else
                        Flog.writeline "Genero el periodo para Costa Rica."
                        Call generarPeriodoVacacion(Ternro, AnioaProc, modeloPais, 0)
                        
                        Flog.writeline "Vuelvo a buscar el periodo para Costa Rica."
                        
                        StrSql = " SELECT * FROM vacacion WHERE vacacion.vacanio = " & AnioaProc & " AND ternro = " & Ternro
                        OpenRecordset StrSql, objRs
                        NroVac = objRs!vacnro
                    End If
               ' End If     MDF ojo capaz q hace cualquiera (1)
                auxNroVac = NroVac
            
            Case 6: 'PARAGUAY
                '************************* PARAGUAY *****************************
                 Select Case VersionPais
                        Case 0: 'Standard Paraguay
                            bprcfechasta = bprcfechastaAux
            
                            Flog.writeline Replace(EscribeLogMI("Genero el período para @@TXT@@."), "@@TXT@@", EscribeLogMI("Paraguay"))
             
                            fechaAlta = FechaAltaEmpleado(Ternro)
                            
                            'Si el empleado no tiene fase, pasa al siguiente empleado
                            If (fechaAlta = Empty) Then
                                GoTo siguiente
                            End If
                                            
                            Flog.writeline EscribeLogMI("Fecha de Procesamiento") & ": " & bprcfechasta & "."
                            'Obtiene el año que se va a procesar a partir de la fase del empleado.
                            If bprcfechasta < CDate(Day(fechaAlta) & "/" & Month(fechaAlta) & "/" & Year(bprcfechasta)) Then
                                AnioaProc = Year(bprcfechasta) - 1
                            Else
                                AnioaProc = Year(bprcfechasta)
                            End If
                            Flog.writeline EscribeLogMI("Año a Procesar") & ": " & AnioaProc & "."
                            Flog.writeline Replace(EscribeLogMI("Busco el periodo para @@TXT@@."), "@@TXT@@", EscribeLogMI("Paraguay"))
                            '----------------
                            StrSql = " SELECT * FROM vacacion WHERE vacacion.vacanio = " & AnioaProc & " AND vacacion.alcannivel = 1"
                            OpenRecordset StrSql, objRs
                            If Not objRs.EOF Then
                                fecha_desde = objRs!vacfecdesde
                                fecha_hasta = objRs!vacfechasta
                                Periodo_Anio = objRs!vacanio
                                NroVac = objRs!vacnro
                                '-------------------
                                'Valido periodos.
                                auxNroVac = Valida_Periodo_STD(NroVac)
                                '-------------------
                                'Llamo función que inserta en vac_alcan
                                Call generarPeriodoVacacionAlance(Ternro, AnioaProc, modeloPais)
                            Else
                                Flog.writeline Replace(EscribeLogMI("Genero el período para @@TXT@@."), "@@TXT@@", EscribeLogMI("Paraguay"))
                                Call generarPeriodoVacacionAlance(Ternro, AnioaProc, modeloPais)
                            End If
                            auxNroVac = NroVac
                        End Select
            Case 7: 'El Salvador
                'EAM (v5.26) - Se modifico para que siempre se procese al domingo anterior a la fecha de procesamiento
                
                'bprcfechasta = DateAdd("d", 1 - Weekday(bprcfechastaAux), bprcfechastaAux) MDF- dijeron q no es asi!!
                
                Flog.writeline "Busco el periodo para El Salvador."
                fechaAlta = FechaAltaEmpleado(Ternro)
                
                'EAM- Si el empleado no tiene fase, pasa al siguiente empleado
                If (fechaAlta = Empty) Then
                    GoTo siguiente
                End If
                                
                Flog.writeline "Fecha Procesamiento: " & bprcfechasta & "."
                                                
                'EAM- Obtiene el año que se va a procesar a partir de la fase del empleado.
                
                
                If bprcfechasta <= CDate(Day(fechaAlta) & "/" & Month(fechaAlta) & "/" & Year(bprcfechasta)) Then
                    AnioaProc = Year(bprcfechasta) - 1
                Else
                    AnioaProc = Year(bprcfechasta)
                End If
                Flog.writeline "Año de Procesar: " & AnioaProc & "."
                
                StrSql = " SELECT * FROM vacacion WHERE vacacion.vacanio = " & AnioaProc & " AND ternro = " & Ternro
                OpenRecordset StrSql, objRs
                If Not objRs.EOF Then
                    fecha_desde = objRs!vacfecdesde
                    fecha_hasta = objRs!vacfechasta
                    Periodo_Anio = objRs!vacanio
                    NroVac = objRs!vacnro
                Else
                    Flog.writeline "Genero el periodo para El Salvador."
                    Call generarPeriodoVacacion(Ternro, AnioaProc, modeloPais)
                    
                    Flog.writeline "Vuelvo a buscar el periodo para Costa Rica."
                    StrSql = " SELECT * FROM vacacion WHERE vacacion.vacanio = " & AnioaProc & " AND ternro = " & Ternro
                    OpenRecordset StrSql, objRs
                    If Not objRs.EOF Then
                        fecha_desde = objRs!vacfecdesde
                        fecha_hasta = objRs!vacfechasta
                        Periodo_Anio = objRs!vacanio
                        NroVac = objRs!vacnro
                    Else
                        Flog.writeline "No se encontró el periodo de vacaciones para el ternro:" & Ternro
                        MyRollbackTrans
                        GoTo siguiente
                    End If
                End If
                auxNroVac = NroVac
            Case Else 'el resto (Agerntina - Uruguay)
                StrSql = " SELECT * FROM vacacion WHERE vacacion.vacnro = " & NroVac
                OpenRecordset StrSql, objRs
                If Not objRs.EOF Then
                    fecha_desde = objRs!vacfecdesde
                    fecha_hasta = objRs!vacfechasta
                    Periodo_Anio = objRs!vacanio
                    
                    AnioaProc = Periodo_Anio '14/05/2012
                    auxNroVac = NroVac       '14/05/2012
                Else
                    Exit Sub
                End If
                
                
                ' ---------------------------------------------------------------------
                'FGZ - 25/03/2010 --------------------------
                'Call bus_DiasVac(Ternro, NroVac, cantdias, Columna, Mensaje, Genera)
                If TodosEmpleados Then
                    'buscar el periodo correspondiente al empleado de acuerdo al alcance
                    auxNroVac = PeriodoCorrespondiente(Ternro, AnioaProc)
                Else
                    auxNroVac = NroVac
                End If
                StrSql = " SELECT * FROM vacacion WHERE vacacion.vacnro = " & auxNroVac
                OpenRecordset StrSql, objRs
                If Not objRs.EOF Then
                    fecha_desde = objRs!vacfecdesde
                    fecha_hasta = objRs!vacfechasta
                    Periodo_Anio = objRs!vacanio
                Else
                    Flog.writeline "No existe un periodo de vacaciones para el año " & AnioaProc & " para el Empleado " & Ternro
                    MyRollbackTrans
                    GoTo siguiente
                End If
        End Select
        
      
        
        
        'EAM- Calcula según el modelo de vacaciones configurado
        Select Case modeloPais
            Case 0: 'Argentina
                Flog.writeline "Modelo de vacaciones de Argentina nro." & modeloPais
                Call bus_DiasVac(Ternro, auxNroVac, cantdias, Columna, Mensaje, Genera, cantdiasCorr, columna2)
                DiasCorraGen = cantdias
            Case 1:
                Flog.writeline "Modelo de vacaciones de Uruguay nro." & modeloPais
                Call bus_DiasVac_uy(Ternro, auxNroVac, cantdias, Columna, Mensaje, Genera)
                DiasCorraGen = cantdias
            Case 2:
                Flog.writeline "Modelo de vacaciones de Chile nro." & modeloPais
                DiasCorraGen = cantdias
            Case 3: 'Colombia
                Flog.writeline "Modelo de vacaciones de Colombia nro." & modeloPais
                Call bus_DiasVac_Col(Ternro, auxNroVac, cantdias, Columna, Mensaje, Genera)
                DiasCorraGen = cantdias
            Case 4: 'Costa Rica
                Flog.writeline "Modelo de vacaciones de Costa Rica Nro." & modeloPais
                Flog.writeline "Se procesará hasta la fecha " & bprcfechasta
                Flog.writeline ""
                Call bus_DiasVac_CR(Ternro, auxNroVac, fechaAlta, bprcfechasta, CantdiasCR, Columna, Mensaje, Genera, Periodo_Anio, generarRegVacDiasCor)
                DiasCorraGen = CantdiasCR
            Case 5: 'Portugal
                Flog.writeline "Modelo de vacaciones de Portugal Nro." & modeloPais
                Call bus_DiasVac_PT(Ternro, auxNroVac, cantdias, Columna, Mensaje, Genera, cantdiasCorr, columna2)
                DiasCorraGen = cantdias
            Case 6: 'Paraguay
                '************************* PARAGUAY *****************************
                Select Case VersionPais
                    Case 0: 'Standard Paraguay
                        Texto = Replace(Texto, "@@TXT@@", EscribeLogMI("Paraguay"))
                        Flog.writeline Texto & " " & modeloPais
                        Flog.writeline ""
                        Call bus_DiasVac_PY(Ternro, auxNroVac, fechaAlta, bprcfechasta, CantdiasCR, Columna, Mensaje, Genera, Periodo_Anio, bprcfechasta, Reproceso)
                        DiasCorraGen = CantdiasCR
                    End Select
            Case 7:
                Flog.writeline "Modelo de vacaciones de El Salvador nro." & modeloPais
                generarperiodossv = True
                Call bus_DiasVac_SA(Ternro, auxNroVac, cantdias, Columna, Mensaje, Genera, cantdiasCorr, columna2, bprcfechasta)
                If Not generarperiodossv Then
                  Exit Sub 'mdf
                End If
                DiasCorraGen = cantdias
            
        End Select
       
       'FGZ - 25/03/2010 --------------------------
        
        ''Flog.writeline ""
        'Flog.writeline "genera: " & Genera
        'Flog.writeline ""
        If Not Genera Then GoTo siguiente
        
        If st_TipoDia2 = 0 Then
            NroTPVCorr = columna2
            StrSql = "SELECT * FROM tipovacac WHERE tpvnrocol = " & Columna
            OpenRecordset StrSql, rs
            
            If Not rs.EOF Then
                NroTPV = rs!tipvacnro
            Else
                'EAM- Verifica si tiene el tipo de días de vacaciones configurado Pol(1501)
                'sino pone el Primero de la tabla po Default
                If (st_TipoDia1 > 0) Then
                    NroTPV = st_TipoDia1
                Else
                    NroTPV = 1 ' por default
                End If
            End If
        Else
            NroTPV = Columna
            NroTPVCorr = columna2
        End If
        
        If modeloPais <> 4 Then
            StrSql = "SELECT * FROM vacdiascor WHERE vacnro = " & auxNroVac & " AND Ternro = " & Ternro
            StrSql = StrSql & " AND (venc = 0 OR venc IS NULL)"
            OpenRecordset StrSql, rs
            If Not rs.EOF Then
                If Reproceso Then
                    If Not IsNull(NroTPV) And Not NroTPV = "" Then
                        'StrSql = "UPDATE vacdiascor SET vdiascormanual = 0, vdiascorcant = " & Cantdias & ", tipvacnro = " & NroTPV & " WHERE vacnro = " & auxNroVac & " AND Ternro = " & Ternro
                        StrSql = "UPDATE vacdiascor SET vdiascormanual = 0, vdiascorcant = " & DiasCorraGen & ", tipvacnro = " & NroTPV & _
                                 " ,vdiascorcantcorr= " & cantdiasCorr & ", tipvacnrocorr= " & NroTPVCorr
                        
                        'EAM- Se agrego para CR el campo de la ultima generacion de Dias Correspondiente
                        If Not IsNull(bprcfechasta) Then
                            StrSql = StrSql & ",vdiasfechasta = " & ConvFecha(bprcfechasta)
                        End If
                        StrSql = StrSql & " WHERE vacnro = " & auxNroVac & " AND Ternro = " & Ternro & " AND (venc = 0 OR venc IS NULL)"
                        
                    Else
                        'StrSql = "UPDATE vacdiascor SET vdiascormanual = 0, vdiascorcant = " & Cantdias & ", tipvacnro = 1 WHERE vacnro = " & auxNroVac & " AND Ternro = " & Ternro
                        StrSql = "UPDATE vacdiascor SET vdiascormanual = 0, vdiascorcant = " & DiasCorraGen & ", tipvacnro = 1 " & _
                                 " vdiascorcantcorr= " & cantdiasCorr & " tipvacnrocorr= " & NroTPVCorr
                        'EAM- Se agrego para CR el campo de la ultima generacion de Dias Correspondiente
                        If Not IsNull(bprcfechasta) Then
                            StrSql = StrSql & ",vdiasfechasta = " & ConvFecha(bprcfechasta)
                        End If
                        
                        StrSql = StrSql & " WHERE vacnro = " & auxNroVac & " AND Ternro = " & Ternro & " AND (venc = 0 OR venc IS NULL)"
                    End If
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
            Else
                If Not IsNull(NroTPV) And Not NroTPV = "" Then
                    'StrSql = "INSERT INTO vacdiascor(vacnro,vdiascorcant,vdiascormanual,ternro,tipvacnro) VALUES (" & _
                    '         auxNroVac & "," & Cantdias & ",0," & Ternro & "," & NroTPV & ")"
                    StrSql = "INSERT INTO vacdiascor(vacnro,vdiascorcant,vdiascormanual,ternro,tipvacnro,vdiasfechasta,vdiascorcantcorr,tipvacnrocorr) VALUES (" & _
                             auxNroVac & "," & DiasCorraGen & ",0," & Ternro & "," & NroTPV & "," & ConvFecha(bprcfechasta) & "," & cantdiasCorr & "," & _
                             NroTPVCorr & ")"
                Else
                    'StrSql = "INSERT INTO vacdiascor(vacnro,vdiascorcant,vdiascormanual,ternro,tipvacnro) VALUES (" & _
                    '         auxNroVac & "," & Cantdias & ",0," & Ternro & ",1)"
                    StrSql = "INSERT INTO vacdiascor(vacnro,vdiascorcant,vdiascormanual,ternro,tipvacnro,vdiasfechasta,vdiascorcantcorr,tipvacnrocorr) VALUES (" & _
                             auxNroVac & "," & DiasCorraGen & ",0," & Ternro & ",1," & ConvFecha(bprcfechasta) & "," & cantdiasCorr & "," & _
                             NroTPVCorr & ")"
                End If
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
        End If

        'FGZ - 21/10/2009 - le agregué esta politica para manejar el vencimiento de dias de vacaciones
        CalculaVencimientos = False
        Call Politica(1512)
            If CalculaVencimientos Then
                'Busco el periodo del año anterior
                NroVacAnterior = PeriodoCorrespondiente(Ternro, Periodo_Anio - 1)
                If NroVacAnterior <> 0 Then
                    'EAM- 16/11/2010 Se saco el calculo de Vencimiento de vacaciones para la caja
                    'Call DiasVencidos(Ternro, auxNroVac, NroVacAnterior)
                End If
            End If
        

        'Customizacion TTI
        PoliticaOK = False
        Call Politica(1504)
        If PoliticaOK Then
           Flog.writeline "Politica 1504 activa. Actualizando cartera de vacaciones (Customizacion TTI)."
           Call Actulizar_Cartera_Vac(Ternro, auxNroVac, cantdias, Reproceso)
        End If
        
        
        'EAM- Politica de Beneficio de días de Vacaciones
        PoliticaOK = False
        Call Politica(1514)
        
        'EAM (v3.38) -  CAS-21472
        If PoliticaOK Then
            Flog.writeline "Politica 1514 activa."
            Flog.writeline "Versión de Política 1514: " & st_Opcion
        End If
        
        If PoliticaOK Then
            Select Case st_Opcion
                Case 1: 'Costa Rica
                    Flog.writeline "Politica 1514 Activa. Cálculo de Beneficio Adicional de vacaciones. COSTA RICA"
                    Flog.writeline "call CalcularBeneficioVac(" & Ternro & "," & auxNroVac & "," & NroTPV & "," & Reproceso & "," & fecha_desde & ")"
                    Call CalcularBeneficioVac(Ternro, auxNroVac, NroTPV, Reproceso, fecha_desde) ' NG - 24/05/2012
                    'Call CalcularBeneficioVac(Ternro, auxNroVac, NroTPV, Reproceso, fecha_hasta) ' NG - 23/05/2012
                    'Call CalcularBeneficioVac(Ternro, auxNroVac, NroTPV, Reproceso, bprcfechasta)
                    
                Case 2: 'Portugal
                    If cantdias = 22 Then
                        Flog.writeline "Politica 1514 Activa. Cálculo de PLUS Adicional de vacaciones. PORTUGAL"
                        Call CalcularBeneficioVac_PT(Ternro, auxNroVac, NroTPV, fecha_desde, fecha_hasta, Lic_Descuento, Reproceso, bprcfechasta)
                    End If
                Case 3: 'beneficio de 25 años uruguay
                    Call CalcularBeneficioVac_UY(Ternro, auxNroVac, NroTPV, fecha_desde, fecha_hasta, Lic_Descuento, Reproceso, bprcfechasta)
            End Select
        End If
        
        MyCommitTrans
' ----------------------------------------------------------
siguiente:
        MyCommitTrans
          
        Progreso = Round(Progreso + IncPorc, 4)
            
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
        
        Flog.writeline " "
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
        
fin:
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
   Fecha_Inicio = T.fechaInicio
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

Public Sub bus_DiasVac(ByVal Ternro As Long, ByVal NroVac As Long, ByRef cantdias As Integer, ByRef Columna As Integer, ByRef Mensaje As String, ByRef Genera As Boolean _
    , ByRef cantdiasCorr As Integer, ByRef columna2 As Integer)
' ---------------------------------------------------------------------------------------------
' Descripcion: Obtencion del valor de la escala para vacaciones.
' Autor      : FGZ
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Valor_Grilla(10) As Boolean ' Elemento de una coordenada de una grilla
Dim tipoBus As Long
Dim concnro As Long
Dim prog As Long

Dim tdinteger3 As Integer

Dim ValAnt As Single
Dim Busq As Integer
Dim dias_maternidad As Integer

Dim j As Integer
Dim antig As Boolean
Dim pvariable As Boolean
Dim pvar As Integer
Dim ant As Integer
Dim continuar As Boolean
Dim parametros(5) As Integer
Dim grilla_val(10) As Boolean     ' para alojar los valores de:  valgrilla.val(i)

Dim vgrcoor_ant As Long
Dim vgrcoor_pvar As Long

Dim rs_valgrilla As New ADODB.Recordset
Dim rs_cabgrilla As New ADODB.Recordset
Dim rs_tbase As New ADODB.Recordset
Dim TipoBase As Long

Dim NroBusqueda As Long

Dim antdia As Long
Dim antmes As Long
Dim antanio As Long
Dim q As Integer

Dim Aux_Dias_trab As Double
Dim aux_redondeo As Double
Dim ValorCoord As Single
Dim Encontro As Boolean
Dim VersionBaseAntig As Integer
Dim habiles, habilesCorr As Integer
Dim ExcluyeFeriados As Boolean
Dim ExcluyeFeriadosCorr  As Boolean
Dim rs As New ADODB.Recordset

'EAM- 08-07-2010
Dim dias_efect_trabajado As Long
Dim regHorarioActual As Integer

Dim Version As String

Version = ""
'Dim arrEscala()
'ReDim Preserve arrEscala(5, 0)  'la escala la carga al (total de registros y )

    Genera = False
    Encontro = False
    
    Call Politica(1502)
    If Not PoliticaOK Then
        Flog.writeline "Error cargando configuracion de la Politica 1502"
        Exit Sub
    End If
    

    StrSql = "SELECT * FROM cabgrilla " & _
             " WHERE cabgrilla.cgrnro = " & NroGrilla
    OpenRecordset StrSql, rs_cabgrilla

    If rs_cabgrilla.EOF Then
        'La escala de Vacaciones no esta configurada en el tipo de dia para vacaciones
        Flog.writeline "La escala de Vacaciones no esta configurada o el nro de grilla no esta bien configurado bien en la Politica 1502. Grilla " & NroGrilla
        Exit Sub
    End If
    
    Call Politica(1505)
    If Not PoliticaOK Then
        Flog.writeline "Error cargando configuracion de la Politica 1505. Tipo Base antiguedad estandar."
        VersionBaseAntig = 0
    Else
        VersionBaseAntig = st_BaseAntiguedad
    End If
    
    
    'El tipo Base de la antiguedad
    TipoBase = 4
    
    continuar = True
    ant = 1
    Do While (ant <= rs_cabgrilla!cgrdimension) And continuar
        Select Case ant
        Case 1:
            StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_1
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    continuar = False
                Else
                    ant = ant + 1
                End If
            End If
            
        Case 2:
            StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_2
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    continuar = False
                Else
                    ant = ant + 1
                End If
            End If
        Case 3:
            StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_3
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    continuar = False
                Else
                    ant = ant + 1
                End If
            End If
        Case 4:
            StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_4
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    continuar = False
                Else
                    ant = ant + 1
                End If
            End If
        Case 5:
            StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_5
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    continuar = False
                Else
                    ant = ant + 1
                End If
            End If
        End Select
    Loop

    'setea la proporcion de dias
    Call Politica(1501)
    Version = st_Opcion
    For j = 1 To rs_cabgrilla!cgrdimension
        Select Case j
        Case ant:
            'la busqueda es la de antiguedad
            Select Case VersionBaseAntig
            Case 0:
                Flog.writeline "Antiguedad estandar "
                Flog.writeline "Antiguedad En el ultimo año " ' Se computa al año actual
                Call bus_Antiguedad_R("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
                If (((antmes * 30) + antdia >= st_Dias) Or antanio <> 0) Then
                       antdia = 0
                       antmes = 0
                       antanio = 0
                       Call bus_Antiguedad("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
                 End If
            
            Case 1:
                Flog.writeline "Antiguedad Sin redondeo "
                Call bus_Antiguedad_R("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
                 If (((antmes * 30) + antdia) >= st_Dias) Or antanio <> 0 Then
                       antdia = 0
                       antmes = 0
                       antanio = 0
                      Call bus_Antiguedad_G("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
                 End If
            Case 2:
                Flog.writeline "Antiguedad Uruguay " ' Se computa al año anterior
                'Call bus_Antiguedad_G("VACACIONES", CDate("31/12/" & Year(fecha_desde)), antdia, antmes, antanio, q)
                Call bus_Antiguedad("VACACIONES", CDate("31/12/" & Periodo_Anio - 1), antdia, antmes, antanio, q)
            Case 3:
                 Flog.writeline "Antiguedad Standard " ' Se computa al año actual
                 Call bus_Antiguedad_R("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
                 If (((antmes * 30) + antdia) >= st_Dias) Or antanio <> 0 Then
                       antdia = 0
                       antmes = 0
                       antanio = 0
                       Call bus_Antiguedad("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
                 End If
            Case 4: ' Anguedad a una fecha dada por dia y mes del año
                Flog.writeline "Antiguedad a una fecha dada año siguiente"
                If Not (st_Dia = 0 Or st_Mes = 0) Then
                     Call bus_Antiguedad_R("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
                     If (((antmes * 30) + antdia >= st_Dias) Or antanio <> 0) Then
                       antdia = 0
                       antmes = 0
                       antanio = 0
                       Call bus_Antiguedad_G("VACACIONES", CDate(st_Dia & "/" & st_Mes & "/" & Periodo_Anio + 1), antdia, antmes, antanio, q)
                    End If
                 End If
            Case 5: ' Anguedad a una fecha dada por dia y mes del año
                Flog.writeline "Antiguedad a una fecha dada año actual"
                Call bus_Antiguedad_R("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
                If (((antmes * 30) + antdia >= st_Dias) Or antanio <> 0) Then
                       antdia = 0
                       antmes = 0
                       antanio = 0
                 Call bus_Antiguedad("VACACIONES", CDate(st_Dia & "/" & st_Mes & "/" & Periodo_Anio), antdia, antmes, antanio, q)
                 End If
            
            Case Else
                Flog.writeline "Antiguedad Mal configurada. Estandar "
                'Call bus_Antiguedad("VACACIONES", CDate("31/12/" & Year(fecha_desde)), antdia, antmes, antanio, q)
                Call bus_Antiguedad("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
            End Select
            parametros(j) = (antanio * 12) + antmes
            'Call bus_Antiguedad("VACACIONES", CDate("31/12/" & Year(fecha_desde)), antdia, antmes, antanio, q)
            Flog.writeline "Años " & antanio
            Flog.writeline "Meses " & antmes
            Flog.writeline "Dias " & antdia
            
        Case Else:
            Select Case j
            Case 1:
                Call bus_Estructura(rs_cabgrilla!grparnro_1)
            Case 2:
                Call bus_Estructura(rs_cabgrilla!grparnro_2)
            Case 3:
                Call bus_Estructura(rs_cabgrilla!grparnro_3)
            Case 4:
                Call bus_Estructura(rs_cabgrilla!grparnro_4)
            Case 5:
                Call bus_Estructura(rs_cabgrilla!grparnro_5)
            End Select
            parametros(j) = valor
        End Select
    Next j

'--------------------------------------------------------------------------------------------------
'EAM 18/01/2012- Se comento porque ahora esto se resuelve en la funcion buscarDiasVacEscala
'    'Busco la primera antiguedad de la escala menor a la del empleado de abajo hacia arriba
'    StrSql = "SELECT * FROM valgrilla WHERE cgrnro = " & NroGrilla
'    For j = 1 To rs_cabgrilla!cgrdimension
'        If j <> ant Then
'            StrSql = StrSql & " AND vgrcoor_" & j & "= " & parametros(j)
'            'ReDim Preserve arrEscala(0, UBound(arrEscala, 2) + 1)
'            ReDim Preserve arrEscala(UBound(arrEscala, 2) + 1, UBound(arrEscala, 2) + 1)
'        End If
'    Next j
'    StrSql = StrSql & " ORDER BY vgrcoor_" & ant & " DESC "
'    OpenRecordset StrSql, rs_valgrilla
'
'
'    Do While Not rs_valgrilla.EOF
'        ReDim Preserve arrEscala(1, UBound(arrEscala, 2))
'        'Dim arrEscala(0, 0)
'    Loop



'
'    Encontro = False
'    Do While Not rs_valgrilla.EOF And Not Encontro
'        Select Case ant
'        Case 1:
'            If parametros(ant) >= rs_valgrilla!vgrcoor_1 Then
'                 If rs_valgrilla!vgrvalor <> 0 Then
'                    cantdias = rs_valgrilla!vgrvalor
'                    Encontro = True
'                    Columna = rs_valgrilla!vgrorden
'                 End If
'            End If
'        Case 2:
'            If parametros(ant) >= rs_valgrilla!vgrcoor_2 Then
'                 If rs_valgrilla!vgrvalor <> 0 Then
'                    cantdias = rs_valgrilla!vgrvalor
'                    Encontro = True
'                    Columna = rs_valgrilla!vgrorden
'                 End If
'            End If
'        Case 3:
'            If parametros(ant) >= rs_valgrilla!vgrcoor_3 Then
'                 If rs_valgrilla!vgrvalor <> 0 Then
'                    cantdias = rs_valgrilla!vgrvalor
'                    Encontro = True
'                    Columna = rs_valgrilla!vgrorden
'                 End If
'            End If
'        Case 4:
'            If parametros(ant) >= rs_valgrilla!vgrcoor_4 Then
'                 If rs_valgrilla!vgrvalor <> 0 Then
'                    cantdias = rs_valgrilla!vgrvalor
'                    Encontro = True
'                    Columna = rs_valgrilla!vgrorden
'                 End If
'            End If
'        Case 5:
'            If parametros(ant) >= rs_valgrilla!vgrcoor_5 Then
'                 If rs_valgrilla!vgrvalor <> 0 Then
'                    cantdias = rs_valgrilla!vgrvalor
'                    Encontro = True
'                    Columna = rs_valgrilla!vgrorden
'                 End If
'            End If
'        End Select
'
'        rs_valgrilla.MoveNext
'    Loop
'--------------------------------------------------------------------------------------------------
    cantdias = buscarDiasVacEscala(ant, rs_cabgrilla!cgrdimension, parametros, TipoVacacionProporcion, Encontro)
    Columna = TipoVacacionProporcion
    cantdiasCorr = buscarDiasVacEscala(ant, rs_cabgrilla!cgrdimension, parametros, TipoVacacionProporcionCorr)
    columna2 = TipoVacacionProporcionCorr

    '------------------------------
    'llamada politica 1513
    '------------------------------
    
    'EAM- Tiene en cuenta los dias trabajados en el ultimo año
    Call Politica(1513)
    
    
    If Dias_efect_trab_anio Then
        Flog.writeline "Tiene en cuenta el ultimo año. Politica 1513."
            
        'Obtiene la proporcion de dias_trabajados  -->  (dias trabajados / 7) * regimen horio
        dias_efect_trabajado = DiasHabilTrabajado(Ternro, CDate("01/01/" & Periodo_Anio), CDate("31/12/" & Periodo_Anio))
        regHorarioActual = BuscarRegHorarioActual(Ternro)
        Aux_Dias_trab = ((180 / 7) * regHorarioActual)
        Aux_Dias_trab = RedondearNumero(Int(Aux_Dias_trab), (Aux_Dias_trab - Int(Aux_Dias_trab)))
        
        If dias_efect_trabajado <= Aux_Dias_trab Then
            Encontro = True
            cantdias = CalcularProporcionDiasVac(dias_efect_trabajado)
            
            Flog.writeline "Empleado " & Ternro & " con dias trabajado menor a mitad de año: " & dias_efect_trabajado
            Flog.writeline "Días Correspondientes: " & cantdias
            Flog.writeline "Tipo de redondeo: " & st_redondeo
            Flog.writeline "Parte decimal de los días correspondientes: " & aux_redondeo
            Flog.writeline
        End If
        
    End If
            
            
                                
    If Not Encontro Then
                
        'EAM- Si la columna1 es = a vacion no tiene el tipo de vacacion sino ya se configuro por la politica 1501 (columna,columna2) 10/02/2011
        If (Columna = 0) Then
            'Busco si existe algun valor para la estructura y ...
            'si hay que carga la columna correspondiente
            StrSql = "SELECT * FROM valgrilla WHERE cgrnro = " & NroGrilla
            StrSql = StrSql & " AND vgrvalor is not null"
            For j = 1 To rs_cabgrilla!cgrdimension
                If j <> ant Then
                    StrSql = StrSql & " AND vgrcoor_" & j & "= " & parametros(j)
                End If
            Next j
            OpenRecordset StrSql, rs_valgrilla
            If Not rs_valgrilla.EOF Then
                Columna = rs_valgrilla!vgrorden
            Else
                Columna = 1
            End If
        End If
        
        If Version = "4" Then
              habiles = cantDiasLaborable(TipoVacacionProporcion, ExcluyeFeriados)
              dias_trabajados = cantDiasProporcion(CDate("31/12/" & Periodo_Anio), habiles)
        Else
              dias_trabajados = ((antanio * 365) + (antmes * 30) + antdia)
        End If
        Flog.writeline "Dias trabajados " & dias_trabajados
        
        Flog.writeline "ANT " & ant
        
        If parametros(ant) <= BaseAntiguedad Then
            
            habiles = cantDiasLaborable(TipoVacacionProporcion, ExcluyeFeriados)
            habilesCorr = cantDiasLaborable(TipoVacacionProporcionCorr, ExcluyeFeriadosCorr)
            
            'EAM- (13972)Esto lo comente porque ahora lo resuelve en la función cantDiasLaborable() para no repetir codigo ya que hay que calcularlo tambien para dias corridos
            '            'FGZ - 16/02/2006
            '            habiles = 0
            '            StrSql = "SELECT * FROM tipovacac WHERE tipvacnro = " & TipoVacacionProporcion
            '            OpenRecordset StrSql, rs
            '            If Not rs.EOF Then
            '                If rs!tpvhabiles__1 Then habiles = habiles + 1
            '                If rs!tpvhabiles__2 Then habiles = habiles + 1
            '                If rs!tpvhabiles__3 Then habiles = habiles + 1
            '                If rs!tpvhabiles__4 Then habiles = habiles + 1
            '                If rs!tpvhabiles__5 Then habiles = habiles + 1
            '                If rs!tpvhabiles__6 Then habiles = habiles + 1
            '                If rs!tpvhabiles__7 Then habiles = habiles + 1
            '
            '                ExcluyeFeriados = CBool(rs!tpvferiado)
            '            Else
            '                'por default tomo 7
            '                habiles = 7
            '            End If
                        'Para que la proporcion sea lo mas exacto posible tengo que
                        'restar a los dias trabajados (que caen dentro de una fase) los dias feiados que son habiles
                        'antes de proporcionar
            If ExcluyeFeriados Then
                'deberia revisar dia por dia de los dias contemplados para la antiguedad revisando si son feriados y dia habil
                
            End If
            
            Flog.writeline "Empleado " & Ternro & " con menos de 6 meses de trabajo."
            Flog.writeline "Dias Proporcion " & DiasProporcion
            Flog.writeline "Factor de Division " & FactorDivision
            Flog.writeline "Tipo Base Antiguedad " & BaseAntiguedad
            Flog.writeline "Dias habiles " & habiles
            Flog.writeline "Dias habiles Corrido" & habilesCorr
            
            
            
'            If DiasProporcion = 20 Then
'                If (dias_trabajados / DiasProporcion) / 7 * 5 > Fix((dias_trabajados / DiasProporcion) / 7 * 5) Then
'                    cantdias = Fix((dias_trabajados / DiasProporcion) / 7 * 5) + 1<d
'                Else
'                    cantdias = Fix((dias_trabajados / DiasProporcion) / 7 * 5)
'                End If
'            Else
'                cantdias = Round((dias_trabajados / DiasProporcion) / FactorDivision, 0)
'            End If
            
'            Agregue el control del parámetro redondeo. Gustavo
                          
             
             If dias_trabajados < 20 Then
                cantdias = 0
             Else
                If DiasProporcion = 20 Then
                    If Version = "4" Then
                        cantdias = Fix(dias_trabajados / DiasProporcion)
                          aux_redondeo = ((dias_trabajados / DiasProporcion)) - Fix((dias_trabajados / DiasProporcion))
                    Else
                        cantdias = Fix((dias_trabajados / DiasProporcion) / 7 * habiles)
                       
                    End If
                Else
                        cantdias = Fix(20 * (dias_trabajados / DiasProporcion) / FactorDivision)
                End If
                If Version <> "4" Then
                    aux_redondeo = ((dias_trabajados / DiasProporcion) / 7 * habiles) - Fix((dias_trabajados / DiasProporcion) / 7 * habiles)
                End If
                cantdias = RedondearNumero(cantdias, aux_redondeo)
'EAM(13972)- Esto se resuelve en la funcin RedondearNumero
'                Select Case st_redondeo
'
'                    Case 0 ' Redondea hacia abajo - Ya se realizo el cálculo
'
'                    Case 1 ' Redondea hacia arriba
'                        If aux_redondeo <> 0 Then
'                            cantdias = cantdias + 1
'                        End If
'
'                    Case Else ' redondea hacia abajo si la parte decimal <.5 sino hacia arriba
'                        If aux_redondeo >= 0.5 Then
'                            cantdias = cantdias + 1
'                        End If
'                End Select
               
                
                'EAM(13972)- Obtiene los dias corridos de vacaciones a partir de los dias correspondientes
                cantdiasCorr = (cantdias * habilesCorr) / habiles
                aux_redondeo = ((cantdias * habilesCorr) / habiles) - Fix(((cantdias * habilesCorr) / habiles))
                cantdiasCorr = RedondearNumero(cantdiasCorr, aux_redondeo)
                
            End If
            Flog.writeline "Días Correspondientes:" & cantdias
            Flog.writeline "Días Correspondientes Corridos:" & cantdiasCorr
            Flog.writeline "Tipo de redondeo:" & st_redondeo
            Flog.writeline "Parte decimal de los días correspondientes:" & aux_redondeo
            Flog.writeline
            
            'FGZ - 25/06/2009 ------------- Vacaciones Acordadas ------------------------------
            PoliticaOK = False
            DiasAcordados = False
            Call Politica(1511)
            If PoliticaOK And DiasAcordados Then
                 StrSql = "SELECT tipvacnro, diasacord, tipvacnrocorr, diasacordcorr FROM vacdiasacord "
                 StrSql = StrSql & " WHERE ternro = " & Ternro
                 OpenRecordset StrSql, rs
                 If Not rs.EOF Then
                     If rs!diasacord > cantdias Then
                        Flog.writeline "La cantidad de dias correspondientes es menor a la cantidad de dias acordados. " & rs!diasacord
                        Flog.writeline "Se utilizará la cantidad de dias acordados"
                        cantdias = rs!diasacord
                        '23/09/2013 - MDZ - CAS-21183 - se setea cantdiasCorr
                        If Not IsNull(rs!diasacordcorr) Then
                            cantdiasCorr = rs!diasacordcorr
                        Else
                            cantdiasCorr = 0
                        End If
                     End If
                 End If
            End If
            'FGZ - 25/06/2009 ------------- Vacaciones Acordadas ------------------------------
            Flog.writeline
            
            ' NF - 03/07/06
            PoliticaOK = False
            Call Politica(1508)
            If PoliticaOK Then
                Flog.writeline "Politica 1508 activa. Analizando Licencias por Maternidad (" & Tipo_Dia_Maternidad & ")."
                dias_maternidad = 0
                'StrSql = "SELECT * FROM emplic "
                StrSql = "SELECT SUM(elcantdias) total FROM emp_lic "
                StrSql = StrSql & " WHERE tdnro = " & Tipo_Dia_Maternidad
                StrSql = StrSql & " AND empleado = " & Ternro
                StrSql = StrSql & " AND elfechadesde >= " & ConvFecha("01/01/" & (Periodo_Anio - 1))
                StrSql = StrSql & " AND elfechahasta <= " & ConvFecha("31/12/" & (Periodo_Anio - 1))
                OpenRecordset StrSql, rs
                If Not rs.EOF And (Not IsNull(rs!total)) Then
                    dias_maternidad = rs!total
                    If dias_maternidad <> 0 Then
                        Flog.writeline "  Dias por maternidad: " & dias_maternidad
                        Flog.writeline "  Dias = " & cantdias & " - (" & dias_maternidad & " x " & Factor & ")"
                        cantdias = cantdias - CInt(dias_maternidad * Factor)
                    End If
                Else
                    Flog.writeline "  No se encontraron dias por maternidad."
                End If
                rs.Close
            End If
        Else
            Flog.writeline "No se encontro la escala para el convenio"
            Genera = False
        End If
    Else
        'FGZ - 25/06/2009 ------------- Vacaciones Acordadas ------------------------------
        PoliticaOK = False
        DiasAcordados = False
        Call Politica(1511)
        If PoliticaOK And DiasAcordados Then
             StrSql = "SELECT tipvacnro, diasacord, tipvacnrocorr, diasacordcorr FROM vacdiasacord "
             StrSql = StrSql & " WHERE ternro = " & Ternro
             OpenRecordset StrSql, rs
             If Not rs.EOF Then
                 If rs!diasacord > cantdias Then
                    Flog.writeline "La cantidad de dias correspondientes es menor a la cantidad de dias acordados. " & rs!diasacord
                    Flog.writeline "Se utilizará la cantidad de dias acordados"
                    cantdias = rs!diasacord
                    '23/09/2013 - MDZ - CAS-21183 - se setea cantdiasCorr
                    If Not IsNull(rs!diasacordcorr) Then
                        cantdiasCorr = rs!diasacordcorr
                    Else
                        cantdiasCorr = 0
                    End If
                 End If
             End If
        End If
        'FGZ - 25/06/2009 ------------- Vacaciones Acordadas ------------------------------
    End If
    
    'EAM(18891)- Descunto a los días de vacaciones (acordados o por escala) la proporción de días por licencias
    Call Politica(1516)
    
    If PoliticaOK And (cantdias > 0) Then
        Select Case st_Opcion
            Case 1:
                Aux_Dias_trab = LicenciaGozadas(Ternro, CDate("01/12/" & (Periodo_Anio - 1)), CDate("30/11/" & (Periodo_Anio)))
                Aux_Dias_trab = (Aux_Dias_trab / DiasProporcion)
                Flog.writeline "Cantidad de días de descuento por Licencias: " & Aux_Dias_trab
                
                Aux_Dias_trab = RedondearNumero(Int(Aux_Dias_trab), (Aux_Dias_trab - Int(Aux_Dias_trab)))
                Flog.writeline "Cantidad de días de descuento por Licencias con redondeo: " & Aux_Dias_trab
                
                cantdias = (cantdias - (st_Dias * Aux_Dias_trab))
                
                If (cantdias < 0) Then
                    cantidas = 0
                End If
                Flog.writeline "Cantidad de días correspondientes: " & cantdias
            Case Else:
                Flog.writeline "No se aplica el descuento de licencia. Versión incorrecta"
        End Select
    End If
    
   
Genera = True
    
' Cierro todo y libero
If rs_cabgrilla.State = adStateOpen Then rs_cabgrilla.Close
If rs_valgrilla.State = adStateOpen Then rs_valgrilla.Close

Set rs_cabgrilla = Nothing
Set rs_valgrilla = Nothing
End Sub
Public Sub bus_DiasVac_SA(ByVal Ternro As Long, ByVal NroVac As Long, ByRef cantdias As Integer, ByRef Columna As Integer, ByRef Mensaje As String, ByRef Genera As Boolean _
    , ByRef cantdiasCorr As Integer, ByRef columna2 As Integer, Optional ByVal fechaGeneracion As Date)
' ---------------------------------------------------------------------------------------------
' Descripcion: Obtencion del valor de la escala para vacaciones.
' Autor      : FGZ
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Valor_Grilla(10) As Boolean ' Elemento de una coordenada de una grilla
Dim tipoBus As Long
Dim concnro As Long
Dim prog As Long

Dim tdinteger3 As Integer

Dim ValAnt As Single
Dim Busq As Integer
Dim dias_maternidad As Integer

Dim j As Integer
Dim antig As Boolean
Dim pvariable As Boolean
Dim pvar As Integer
Dim ant As Integer
Dim continuar As Boolean
Dim parametros(5) As Integer
Dim grilla_val(10) As Boolean     ' para alojar los valores de:  valgrilla.val(i)

Dim vgrcoor_ant As Long
Dim vgrcoor_pvar As Long

Dim rs_valgrilla As New ADODB.Recordset
Dim rs_cabgrilla As New ADODB.Recordset
Dim rs_tbase As New ADODB.Recordset
Dim TipoBase As Long

Dim NroBusqueda As Long

Dim antdia As Long
Dim antmes As Long
Dim antanio As Long
Dim q As Integer

Dim Aux_Dias_trab As Double
Dim aux_redondeo As Double
Dim ValorCoord As Single
Dim Encontro As Boolean
Dim VersionBaseAntig As Integer
Dim habiles, habilesCorr As Integer
Dim ExcluyeFeriados As Boolean
Dim ExcluyeFeriadosCorr  As Boolean
Dim rs As New ADODB.Recordset
'EAM- 08-07-2010
Dim dias_efect_trabajado As Long
Dim regHorarioActual As Integer
Dim antiguedadUltimoPeriodo As Long

'Dim arrEscala()
'ReDim Preserve arrEscala(5, 0)  'la escala la carga al (total de registros y )

    Genera = False
    Encontro = False
    
    Call Politica(1502)
    If Not PoliticaOK Then
        Flog.writeline "Error cargando configuracion de la Politica 1502"
        Exit Sub
    End If
    

    StrSql = "SELECT * FROM cabgrilla " & _
             " WHERE cabgrilla.cgrnro = " & NroGrilla
    OpenRecordset StrSql, rs_cabgrilla

    If rs_cabgrilla.EOF Then
        'La escala de Vacaciones no esta configurada en el tipo de dia para vacaciones
        Flog.writeline "La escala de Vacaciones no esta configurada o el nro de grilla no esta bien configurado bien en la Politica 1502. Grilla " & NroGrilla
        Exit Sub
    End If
    
    Call Politica(1505)
    If Not PoliticaOK Then
        Flog.writeline "Error cargando configuracion de la Politica 1505. Tipo Base antiguedad estandar."
        VersionBaseAntig = 0
    Else
        VersionBaseAntig = st_BaseAntiguedad
    End If
    
    
    'El tipo Base de la antiguedad
    TipoBase = 4
    
    continuar = True
    ant = 1
    Do While (ant <= rs_cabgrilla!cgrdimension) And continuar
        Select Case ant
        Case 1:
            StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_1
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    continuar = False
                Else
                    ant = ant + 1
                End If
            End If
            
        Case 2:
            StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_2
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    continuar = False
                Else
                    ant = ant + 1
                End If
            End If
        Case 3:
            StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_3
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    continuar = False
                Else
                    ant = ant + 1
                End If
            End If
        Case 4:
            StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_4
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    continuar = False
                Else
                    ant = ant + 1
                End If
            End If
        Case 5:
            StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_5
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    continuar = False
                Else
                    ant = ant + 1
                End If
            End If
        End Select
    Loop

    'setea la proporcion de dias
    Call Politica(1501)
        
    For j = 1 To rs_cabgrilla!cgrdimension
        Select Case j
        Case ant:
            'la busqueda es la de antiguedad
            Select Case VersionBaseAntig
            Case 0:
                Flog.writeline "Antiguedad estandar "
                Flog.writeline "Antiguedad En el ultimo año " ' Se computa al año actual
                Call bus_Antiguedad_R("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
                If (((antmes * 30) + antdia >= st_Dias) Or antanio <> 0) Then
                       antdia = 0
                       antmes = 0
                       antanio = 0
                       Call bus_Antiguedad("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
                 End If
            
            Case 1:
                Flog.writeline "Antiguedad Sin redondeo "
                Call bus_Antiguedad_R("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
                 If (((antmes * 30) + antdia) >= st_Dias) Or antanio <> 0 Then
                       antdia = 0
                       antmes = 0
                       antanio = 0
                      Call bus_Antiguedad_G("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
                 End If
            Case 2:
                Flog.writeline "Antiguedad Uruguay " ' Se computa al año anterior
                'Call bus_Antiguedad_G("VACACIONES", CDate("31/12/" & Year(fecha_desde)), antdia, antmes, antanio, q)
                Call bus_Antiguedad("VACACIONES", CDate("31/12/" & Periodo_Anio - 1), antdia, antmes, antanio, q)
            Case 3:
                 Flog.writeline "Antiguedad Standard " ' Se computa al año actual
                 Call bus_Antiguedad_R("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
                 If (((antmes * 30) + antdia) >= st_Dias) Or antanio <> 0 Then
                       antdia = 0
                       antmes = 0
                       antanio = 0
                       Call bus_Antiguedad("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
                 End If
            Case 4: ' Anguedad a una fecha dada por dia y mes del año
                Flog.writeline "Antiguedad a una fecha dada año siguiente"
                If Not (st_Dia = 0 Or st_Mes = 0) Then
                     Call bus_Antiguedad_R("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
                     If (((antmes * 30) + antdia >= st_Dias) Or antanio <> 0) Then
                       antdia = 0
                       antmes = 0
                       antanio = 0
                       Call bus_Antiguedad_G("VACACIONES", CDate(st_Dia & "/" & st_Mes & "/" & Periodo_Anio + 1), antdia, antmes, antanio, q)
                    End If
                 End If
            Case 5: ' Anguedad a una fecha dada por dia y mes del año
                Flog.writeline "Antiguedad a una fecha dada año actual"
                Call bus_Antiguedad_R("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
                If (((antmes * 30) + antdia >= st_Dias) Or antanio <> 0) Then
                       antdia = 0
                       antmes = 0
                       antanio = 0
                 Call bus_Antiguedad("VACACIONES", CDate(st_Dia & "/" & st_Mes & "/" & Periodo_Anio), antdia, antmes, antanio, q)
                 End If
            
            Case 6: 'Antiguedad El salvador. Calcula a la fecha que se está generando los días
                Flog.writeline "Antiguedad a la fecha de generación de días Correspondientes"
                antdia = 0
                antmes = 0
                antanio = 0
                Call bus_Antiguedad("VACACIONES", CDate(fechaGeneracion), antdia, antmes, antanio, q)
                 If antmes = 0 And antdia = 0 Then
                   antmes = 12
                 End If
                 If antmes < 6 Then      'mdf
                  Flog.writeline "No pasaron 6 meses aun despues del aniversario. No se generan dias."
                  generarperiodossv = False 'MDF
                  Exit Sub ' menos de 6 meses no debe generar nada
                End If
            Case Else
                Flog.writeline "Antiguedad Mal configurada. Estandar "
                'Call bus_Antiguedad("VACACIONES", CDate("31/12/" & Year(fecha_desde)), antdia, antmes, antanio, q)
                Call bus_Antiguedad("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
            End Select
            parametros(j) = (antanio * 12) + antmes
            'Call bus_Antiguedad("VACACIONES", CDate("31/12/" & Year(fecha_desde)), antdia, antmes, antanio, q)
            Flog.writeline "Años " & antanio
            Flog.writeline "Meses " & antmes
            Flog.writeline "Dias " & antdia
            
        Case Else:
            Select Case j
            Case 1:
                Call bus_Estructura(rs_cabgrilla!grparnro_1)
            Case 2:
                Call bus_Estructura(rs_cabgrilla!grparnro_2)
            Case 3:
                Call bus_Estructura(rs_cabgrilla!grparnro_3)
            Case 4:
                Call bus_Estructura(rs_cabgrilla!grparnro_4)
            Case 5:
                Call bus_Estructura(rs_cabgrilla!grparnro_5)
            End Select
            parametros(j) = valor
        End Select
    Next j


    'EAM (v3.38) - Busca los valores por escala
    cantdias = buscarDiasVacEscala(ant, rs_cabgrilla!cgrdimension, parametros, TipoVacacionProporcion, Encontro)
    Columna = TipoVacacionProporcion
    cantdiasCorr = buscarDiasVacEscala(ant, rs_cabgrilla!cgrdimension, parametros, TipoVacacionProporcionCorr)
    columna2 = TipoVacacionProporcionCorr

    

    

    '------------------------------
    'llamada politica 1513
    '------------------------------
    
    'EAM- Tiene en cuenta los dias trabajados en el ultimo año
    Call Politica(1513)
    
    
    If Dias_efect_trab_anio Then
        Flog.writeline "Tiene en cuenta el ultimo año. Politica 1513."
            
        'Obtiene la proporcion de dias_trabajados  -->  (dias trabajados / 7) * regimen horio
        dias_efect_trabajado = DiasHabilTrabajado(Ternro, CDate("01/01/" & Periodo_Anio), CDate("31/12/" & Periodo_Anio))
        regHorarioActual = BuscarRegHorarioActual(Ternro)
        Aux_Dias_trab = ((180 / 7) * regHorarioActual)
        Aux_Dias_trab = RedondearNumero(Int(Aux_Dias_trab), (Aux_Dias_trab - Int(Aux_Dias_trab)))
        
        If dias_efect_trabajado <= Aux_Dias_trab Then
            Encontro = True
            cantdias = CalcularProporcionDiasVac(dias_efect_trabajado)
            
            Flog.writeline "Empleado " & Ternro & " con dias trabajado menor a mitad de año: " & dias_efect_trabajado
            Flog.writeline "Días Correspondientes: " & cantdias
            Flog.writeline "Tipo de redondeo: " & st_redondeo
            Flog.writeline "Parte decimal de los días correspondientes: " & aux_redondeo
            Flog.writeline
        End If
        
    End If
     

    
            
    'EAM (v3.38) - Si tiene antiguedad menor a 6 meses le asigno el valor fijo de 7 días habilies y 9 corridos
    
  '----------------------------------------MDF --- Comento todo el codigo que sigue 14/04/2015-------------
    
   ' If ((Not Encontro) And (antmes >= 6)) Then
        'EAM- Si la columna1 es = a vacion no tiene el tipo de vacacion sino ya se configuro por la politica 1501 (columna,columna2) 10/02/2011
   '     If (Columna <> 0) Then
            
   '        dias_trabajados = ((antmes * 30) + antdia)
   '         Flog.writeline "Dias trabajados el último año: " & dias_trabajados
   '
   '         cantdias = 5
   '         Columna = TipoVacacionProporcion
   '         cantdiasCorr = 7
   '         columna2 = TipoVacacionProporcionCorr
   '
   '
   '         Flog.writeline "Empleado " & Ternro & " con menos de 6 meses de trabajo en el último año."
   '         Flog.writeline "Días Correspondientes:" & cantdias
   '         Flog.writeline "Días Correspondientes Corridos:" & cantdiasCorr
   '         Flog.writeline "Tipo de redondeo:" & st_redondeo
   '         Flog.writeline "Parte decimal de los días correspondientes:" & aux_redondeo
   '         Flog.writeline
   '
   '
   '
   '
   '
   '         ' NF - 03/07/06
   '         PoliticaOK = False
   '         Call Politica(1508)
   '         If PoliticaOK Then
   '             Flog.writeline "Politica 1508 activa. Analizando Licencias por Maternidad (" & Tipo_Dia_Maternidad & ")."
   '             dias_maternidad = 0
   '             'StrSql = "SELECT * FROM emplic "
   '             StrSql = "SELECT SUM(elcantdias) total FROM emp_lic "
   '             StrSql = StrSql & " WHERE tdnro = " & Tipo_Dia_Maternidad
   '             StrSql = StrSql & " AND empleado = " & Ternro
   '             StrSql = StrSql & " AND elfechadesde >= " & ConvFecha("01/01/" & (Periodo_Anio - 1))
   '             StrSql = StrSql & " AND elfechahasta <= " & ConvFecha("31/12/" & (Periodo_Anio - 1))
   '             OpenRecordset StrSql, rs
   '             If Not rs.EOF And (Not IsNull(rs!total)) Then
   '                 dias_maternidad = rs!total
   '                 If dias_maternidad <> 0 Then
   '                     Flog.writeline "  Dias por maternidad: " & dias_maternidad
   '                     Flog.writeline "  Dias = " & cantdias & " - (" & dias_maternidad & " x " & Factor & ")"
   '                     cantdias = cantdias - CInt(dias_maternidad * Factor)
   '                 End If
   '             Else
   '                 Flog.writeline "  No se encontraron dias por maternidad."
   '             End If
   '             rs.Close
   '         End If
   '     Else
   '         Flog.writeline "No se encontro Configurado el tipo de vacacion"
   '         Genera = False
   '     End If
   ' Else
   '     'FGZ - 25/06/2009 ------------- Vacaciones Acordadas ------------------------------
   '     PoliticaOK = False
   '     DiasAcordados = False
   '     Call Politica(1511)
   '     If PoliticaOK And DiasAcordados Then
   '          StrSql = "SELECT tipvacnro, diasacord, tipvacnrocorr, diasacordcorr FROM vacdiasacord "
   '          StrSql = StrSql & " WHERE ternro = " & Ternro
   '          OpenRecordset StrSql, rs
   '          If Not rs.EOF Then
   '              If rs!diasacord > cantdias Then
   '                 Flog.writeline "La cantidad de dias correspondientes es menor a la cantidad de dias acordados. " & rs!diasacord
   '                 Flog.writeline "Se utilizará la cantidad de dias acordados"
   '                 cantdias = rs!diasacord
   '                 '23/09/2013 - MDZ - CAS-21183 - se setea cantdiasCorr
   '                 If Not IsNull(rs!diasacordcorr) Then
   '                     Flog.writeline "dias corridos =" & rs!diasacordcorr
   '                     cantdiasCorr = rs!diasacordcorr
   '                 Else
   '                     cantdiasCorr = 0
   '                 End If
   '              End If
   '          End If
   '     End If
   '     'FGZ - 25/06/2009 ------------- Vacaciones Acordadas ------------------------------
   ' End If
    
  '--------------------------------------------MDF HASTA ACA COMENTO EL CODIGO, ESO ES REEMPLAZADO POR LO Q VIENE 14/04/2015
    
  '-------------------------------------------MDF NUEVO CALCULO  14/04/2015
     If Encontro Then
       If (Columna <> 0) Then
          If (antmes >= 6 And antmes < 12) Then
              dias_trabajados = ((antmes * 30) + antdia)
              Flog.writeline "Dias trabajados el último año: " & dias_trabajados
              cantdias = buscarDiasVacEscala(ant, rs_cabgrilla!cgrdimension, parametros, TipoVacacionProporcion, Encontro) / 2
              Columna = TipoVacacionProporcion
              cantdiasCorr = buscarDiasVacEscala(ant, rs_cabgrilla!cgrdimension, parametros, TipoVacacionProporcionCorr) / 2
              columna2 = TipoVacacionProporcionCorr
            
           Else
               If antmes = 12 Then
                 'los dias quedan como esta....
               End If
           End If
          '--------------
             PoliticaOK = False
             Call Politica(1508)
             If PoliticaOK Then
                 Flog.writeline "Politica 1508 activa. Analizando Licencias por Maternidad (" & Tipo_Dia_Maternidad & ")."
                 dias_maternidad = 0
                 'StrSql = "SELECT * FROM emplic "
                 StrSql = "SELECT SUM(elcantdias) total FROM emp_lic "
                 StrSql = StrSql & " WHERE tdnro = " & Tipo_Dia_Maternidad
                 StrSql = StrSql & " AND empleado = " & Ternro
                 StrSql = StrSql & " AND elfechadesde >= " & ConvFecha("01/01/" & (Periodo_Anio - 1))
                 StrSql = StrSql & " AND elfechahasta <= " & ConvFecha("31/12/" & (Periodo_Anio - 1))
                 OpenRecordset StrSql, rs
                 If Not rs.EOF And (Not IsNull(rs!total)) Then
                     dias_maternidad = rs!total
                     If dias_maternidad <> 0 Then
                         Flog.writeline "  Dias por maternidad: " & dias_maternidad
                         Flog.writeline "  Dias = " & cantdias & " - (" & dias_maternidad & " x " & Factor & ")"
                         cantdias = cantdias - CInt(dias_maternidad * Factor)
                     End If
                 Else
                     Flog.writeline "  No se encontraron dias por maternidad."
                 End If
                 rs.Close
             End If
          '--------------
       Else
         Flog.writeline "No se encontro Configurado el tipo de vacacion"
         Genera = False
         
       End If
    Else
        If antmes >= 6 Then
            cantdias = 5
            Columna = TipoVacacionProporcion
            cantdiasCorr = 7
            columna2 = TipoVacacionProporcionCorr
        Else
            'dias acordados
                  'FGZ - 25/06/2009 ------------- Vacaciones Acordadas ------------------------------
            PoliticaOK = False
            DiasAcordados = False
            Call Politica(1511)
            If PoliticaOK And DiasAcordados Then
                 StrSql = "SELECT tipvacnro, diasacord, tipvacnrocorr, diasacordcorr FROM vacdiasacord "
                 StrSql = StrSql & " WHERE ternro = " & Ternro
                 OpenRecordset StrSql, rs
                 If Not rs.EOF Then
                     If rs!diasacord > cantdias Then
                        Flog.writeline "La cantidad de dias correspondientes es menor a la cantidad de dias acordados. " & rs!diasacord
                        Flog.writeline "Se utilizará la cantidad de dias acordados"
                        cantdias = rs!diasacord
                        '23/09/2013 - MDZ - CAS-21183 - se setea cantdiasCorr
                        If Not IsNull(rs!diasacordcorr) Then
                            Flog.writeline "dias corridos =" & rs!diasacordcorr
                            cantdiasCorr = rs!diasacordcorr
                        Else
                            cantdiasCorr = 0
                        End If
                     End If
                 End If
            End If
            'FGZ - 25/06/2009 ------------- Vacaciones Acordadas ------------------------------
        End If
    End If
  
  '-------------------------------------------FIN MDF NUEVO CALCULO  14/04/2015
    
    'EAM(18891)- Descunto a los días de vacaciones (acordados o por escala) la proporción de días por licencias
    Call Politica(1516)
    
    If PoliticaOK And (cantdias > 0) Then
        Select Case st_Opcion
            Case 1:
                Aux_Dias_trab = LicenciaGozadas(Ternro, CDate("01/12/" & (Periodo_Anio - 1)), CDate("30/11/" & (Periodo_Anio)))
                If DiasProporcion <> 0 Then
                 Aux_Dias_trab = (Aux_Dias_trab / DiasProporcion)
                Else
                 Aux_Dias_trab = 0
                End If
                Flog.writeline "Cantidad de días de descuento por Licencias: " & Aux_Dias_trab
                
                Aux_Dias_trab = RedondearNumero(Int(Aux_Dias_trab), (Aux_Dias_trab - Int(Aux_Dias_trab)))
                Flog.writeline "Cantidad de días de descuento por Licencias con redondeo: " & Aux_Dias_trab
              
                cantdias = (cantdias - (st_Dias * Aux_Dias_trab))
                
                If (cantdias < 0) Then
                    cantidas = 0
                End If
                Flog.writeline "Cantidad de días correspondientes: " & cantdias
                
            Case Else:
                Flog.writeline "No se aplica el descuento de licencia. Versión incorrecta"
        End Select
    End If
    
   
Genera = True
    
' Cierro todo y libero
If rs_cabgrilla.State = adStateOpen Then rs_cabgrilla.Close
If rs_valgrilla.State = adStateOpen Then rs_valgrilla.Close

Set rs_cabgrilla = Nothing
Set rs_valgrilla = Nothing
End Sub

Public Sub bus_DiasVac_uy(ByVal Ternro As Long, ByVal NroVac As Long, ByRef cantdias As Integer, ByRef Columna As Integer, ByRef Mensaje As String, ByRef Genera As Boolean)
' ---------------------------------------------------------------------------------------------
' Descripcion: Obtencion del valor de la escala para vacaciones Uruguay.
' Autor      : Margiotta, Emanuel
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Valor_Grilla(10) As Boolean ' Elemento de una coordenada de una grilla
Dim tipoBus As Long
Dim concnro As Long
Dim prog As Long

Dim tdinteger3 As Integer

Dim ValAnt As Single
Dim Busq As Integer
Dim dias_maternidad As Integer

Dim j As Integer
Dim antig As Boolean
Dim pvariable As Boolean
Dim pvar As Integer
Dim ant As Integer
Dim continuar As Boolean
Dim parametros(5) As Integer
Dim grilla_val(10) As Boolean     ' para alojar los valores de:  valgrilla.val(i)

Dim vgrcoor_ant As Long
Dim vgrcoor_pvar As Long

Dim rs_valgrilla As New ADODB.Recordset
Dim rs_cabgrilla As New ADODB.Recordset
Dim rs_tbase As New ADODB.Recordset
Dim TipoBase As Long

Dim NroBusqueda As Long

Dim antdia As Long
Dim antmes As Long
Dim antanio As Long
Dim q As Integer

Dim Aux_Dias_trab As Double
Dim aux_redondeo As Double
Dim ValorCoord As Single
Dim Encontro As Boolean
Dim VersionBaseAntig As Integer
Dim habiles As Integer
Dim ExcluyeFeriados As Boolean
Dim rs As New ADODB.Recordset
'EAM- 08-07-2010
Dim dias_efect_trabajado As Long
Dim regHorarioActual As Integer
Dim aux_antmes As Long


    Genera = False
    
    Call Politica(1502)
    If Not PoliticaOK Then
        Flog.writeline "Error cargando configuracion de la Politica 1502"
        Exit Sub
    End If
    

    StrSql = "SELECT * FROM cabgrilla " & _
             " WHERE cabgrilla.cgrnro = " & NroGrilla
    OpenRecordset StrSql, rs_cabgrilla

    If rs_cabgrilla.EOF Then
        'La escala de Vacaciones no esta configurada en el tipo de dia para vacaciones
        Flog.writeline "La escala de Vacaciones no esta configurada o el nro de grilla no esta bien configurado bien en la Politica 1502. Grilla " & NroGrilla
        Exit Sub
    End If
    Flog.writeline "La escala de Vacaciones está configurada correctamente en la Politica 1502. Grilla " & NroGrilla
    
    Call Politica(1505)
    If Not PoliticaOK Then
        Flog.writeline "Error cargando configuracion de la Politica 1505. Tipo Base antiguedad estandar."
        VersionBaseAntig = 0
    Else
        VersionBaseAntig = st_BaseAntiguedad
    End If
    
    
    'El tipo Base de la antiguedad
    TipoBase = 4
    
    continuar = True
    ant = 1
    Do While (ant <= rs_cabgrilla!cgrdimension) And continuar
        Select Case ant
        Case 1:
            StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_1
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    continuar = False
                Else
                    ant = ant + 1
                End If
            End If
            
        Case 2:
            StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_2
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    continuar = False
                Else
                    ant = ant + 1
                End If
            End If
        Case 3:
            StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_3
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    continuar = False
                Else
                    ant = ant + 1
                End If
            End If
        Case 4:
            StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_4
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    continuar = False
                Else
                    ant = ant + 1
                End If
            End If
        Case 5:
            StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_5
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    continuar = False
                Else
                    ant = ant + 1
                End If
            End If
        End Select
    Loop
    
            
    'Setea la proporcion de dias
    Call Politica(1501)

    For j = 1 To rs_cabgrilla!cgrdimension
        Select Case j
        Case ant:
            'la busqueda es la de antiguedad
            Select Case VersionBaseAntig
            Case 0:
                Flog.writeline "Antiguedad Standard " ' Se computa al año actual
                Call bus_Antiguedad_R("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
                If (((antmes * 30) + antdia >= st_Dias) Or antanio <> 0) Then
                       antdia = 0
                       antmes = 0
                       antanio = 0
                       Call bus_Antiguedad("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
                 End If

            Case 1:
                Flog.writeline "Antiguedad Sin redondeo "
                Call bus_Antiguedad_R("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
                 If (((antmes * 30) + antdia) >= st_Dias) Or antanio <> 0 Then
                       antdia = 0
                       antmes = 0
                       antanio = 0
                      Call bus_Antiguedad_G("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
                 End If
            Case 2:
                Flog.writeline "Antiguedad Uruguay " ' Se computa al año anterior
                'Call bus_Antiguedad_G("VACACIONES", CDate("31/12/" & Year(fecha_desde)), antdia, antmes, antanio, q)
                Call bus_Antiguedad("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
            Case 3:
                 Flog.writeline "Antiguedad Standard " ' Se computa al año actual
                 Call bus_Antiguedad_R("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
                 If (((antmes * 30) + antdia) >= st_Dias) Or antanio <> 0 Then
                       antdia = 0
                       antmes = 0
                       antanio = 0
                       Call bus_Antiguedad("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
                 End If
            Case 4: ' Anguedad a una fecha dada por dia y mes del año
                Flog.writeline "Antiguedad a una fecha dada año siguiente"
                If Not (st_Dia = 0 Or st_Mes = 0) Then
                     Call bus_Antiguedad_R("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
                     If (((antmes * 30) + antdia >= st_Dias) Or antanio <> 0) Then
                       antdia = 0
                       antmes = 0
                       antanio = 0
                       Call bus_Antiguedad_G("VACACIONES", CDate(st_Dia & "/" & st_Mes & "/" & Periodo_Anio + 1), antdia, antmes, antanio, q)
                    End If
                 End If
            Case 5: ' Anguedad a una fecha dada por dia y mes del año
                Flog.writeline "Antiguedad a una fecha dada año actual"
                Call bus_Antiguedad_R("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
                If (((antmes * 30) + antdia >= st_Dias) Or antanio <> 0) Then
                       antdia = 0
                       antmes = 0
                       antanio = 0
                 Call bus_Antiguedad("VACACIONES", CDate(st_Dia & "/" & st_Mes & "/" & Periodo_Anio), antdia, antmes, antanio, q)
                 End If

            Case Else
                Flog.writeline "Antiguedad Mal configurada. Estandar "
                'Call bus_Antiguedad("VACACIONES", CDate("31/12/" & Year(fecha_desde)), antdia, antmes, antanio, q)
                Call bus_Antiguedad("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
            End Select

            parametros(j) = (antanio * 12) + antmes
            
            Flog.writeline "Años " & antanio
            Flog.writeline "Meses " & antmes
            Flog.writeline "Dias " & antdia

        Case Else:
            Select Case j
            Case 1:
                Call bus_Estructura(rs_cabgrilla!grparnro_1)
            Case 2:
                Call bus_Estructura(rs_cabgrilla!grparnro_2)
            Case 3:
                Call bus_Estructura(rs_cabgrilla!grparnro_3)
            Case 4:
                Call bus_Estructura(rs_cabgrilla!grparnro_4)
            Case 5:
                Call bus_Estructura(rs_cabgrilla!grparnro_5)
            End Select
            parametros(j) = valor
        End Select
    Next j

    'Busco la primera antiguedad de la escala menor a la del empleado
    ' de abajo hacia arriba
    StrSql = "SELECT * FROM valgrilla WHERE cgrnro = " & NroGrilla
    For j = 1 To rs_cabgrilla!cgrdimension
        If j <> ant Then
            StrSql = StrSql & " AND vgrcoor_" & j & "= " & parametros(j)
        End If
    Next j
        StrSql = StrSql & " ORDER BY vgrcoor_" & ant & " DESC "
    OpenRecordset StrSql, rs_valgrilla


    Encontro = False
    Do While Not rs_valgrilla.EOF And Not Encontro
        Select Case ant
        Case 1:
            If parametros(ant) >= rs_valgrilla!vgrcoor_1 Then
                 If rs_valgrilla!vgrvalor <> 0 Then
                    cantdias = rs_valgrilla!vgrvalor
                    Encontro = True
                    Columna = rs_valgrilla!vgrorden
                 End If
            End If
        Case 2:
            If parametros(ant) >= rs_valgrilla!vgrcoor_2 Then
                 If rs_valgrilla!vgrvalor <> 0 Then
                    cantdias = rs_valgrilla!vgrvalor
                    Encontro = True
                    Columna = rs_valgrilla!vgrorden
                 End If
            End If
        Case 3:
            If parametros(ant) >= rs_valgrilla!vgrcoor_3 Then
                 If rs_valgrilla!vgrvalor <> 0 Then
                    cantdias = rs_valgrilla!vgrvalor
                    Encontro = True
                    Columna = rs_valgrilla!vgrorden
                 End If
            End If
        Case 4:
            If parametros(ant) >= rs_valgrilla!vgrcoor_4 Then
                 If rs_valgrilla!vgrvalor <> 0 Then
                    cantdias = rs_valgrilla!vgrvalor
                    Encontro = True
                    Columna = rs_valgrilla!vgrorden
                 End If
            End If
        Case 5:
            If parametros(ant) >= rs_valgrilla!vgrcoor_5 Then
                 If rs_valgrilla!vgrvalor <> 0 Then
                    cantdias = rs_valgrilla!vgrvalor
                    Encontro = True
                    Columna = rs_valgrilla!vgrorden
                 End If
            End If
        End Select

        rs_valgrilla.MoveNext
    Loop

    
    
    '------------------------------
    'llamada politica 1513
    '------------------------------
    
    'EAM- Tiene en cuenta los dias trabajados en el ultimo año
    Call Politica(1513)
    
    
    If Dias_efect_trab_anio Then
        Flog.writeline "Tiene en cuenta el ultimo año. Politica 1513."
        antdia = 0
        antmes = 0
        antanio = 0
        'EAM- Calcula los dias correspondientes segun los meses trabajados en el ultimo año
        Call bus_Antiguedad_U("VACACIONES", CDate("01/01/" & Periodo_Anio), CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
        aux_antmes = (antanio * 12) + antmes
        'FGZ - 19/11/2010 -----------------------------------------
        'Aux_Dias_trab = ((20 / 12) * aux_antmes)
        Aux_Dias_trab = ((((antmes * 30) + antdia) / 30)) * 1.6667
        'FGZ - 19/11/2010 -----------------------------------------
        cantdias = RedondearNumero(Int(Aux_Dias_trab), (Aux_Dias_trab - Int(Aux_Dias_trab)))
        
        Encontro = True
        Flog.writeline "Empleado " & Ternro & " con meses trabajado en el último año: " & antmes
        Flog.writeline "Días Correspondientes:" & cantdias
        Flog.writeline "Tipo de redondeo:" & st_redondeo
        Flog.writeline
    End If
    
    
    If Not Encontro Then
        'EAM- Calcula el proporcional de vacaciones que le corresponde en funcion de los meses trabajados
        'Aux_Dias_trab = ((20 / 12) * Parametros(ant))
        
        Aux_Dias_trab = ((((antmes * 30) + antdia) / 30)) * 1.6667
        cantdias = RedondearNumero(Int(Aux_Dias_trab), (Aux_Dias_trab - Int(Aux_Dias_trab)))
        
        Flog.writeline "Dias Proporcion " & cantdias
        
    End If
                        
    'EAM(18891)- Descunto a los días de vacaciones (acordados o por escala) la proporción de días por licencias
    Call Politica(1516)
    
    If PoliticaOK And (cantdias > 0) Then
        Select Case st_Opcion
            Case 1:
                Aux_Dias_trab = LicenciaGozadas(Ternro, CDate("01/12/" & (Periodo_Anio - 1)), CDate("30/11/" & (Periodo_Anio)))
                Aux_Dias_trab = (Aux_Dias_trab / DiasProporcion)
                Flog.writeline "Cantidad de días de descuento por Licencias: " & Aux_Dias_trab
                
                Aux_Dias_trab = RedondearNumero(Int(Aux_Dias_trab), (Aux_Dias_trab - Int(Aux_Dias_trab)))
                Flog.writeline "Cantidad de días de descuento por Licencias con redondeo: " & Aux_Dias_trab
                
                cantdias = (cantdias - (st_Dias * Aux_Dias_trab))
                
                If (cantdias < 0) Then
                    cantidas = 0
                End If
                Flog.writeline "Cantidad de días correspondientes: " & cantdias
            Case Else:
                Flog.writeline "No se aplica el descuento de licencia. Versión incorrecta"
        End Select
    End If

Genera = True
    
' Cierro todo y libero
If rs_cabgrilla.State = adStateOpen Then rs_cabgrilla.Close
If rs_valgrilla.State = adStateOpen Then rs_valgrilla.Close

Set rs_cabgrilla = Nothing
Set rs_valgrilla = Nothing
End Sub


Public Sub bus_DiasVac_CR_OLD(ByVal Ternro As Long, ByRef NroVac As Long, ByVal fechaAlta As Date, ByRef FechaHasta As Date, ByRef cantdias As Double, ByRef Columna As Integer, ByRef Mensaje As String, ByRef Genera As Boolean, ByVal Anio As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Calcula los días de vacaciones
' Autor      : Margiotta, Emanuel
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim totalDias As Long
Dim fechadesde As Date
Dim ultFechaProcesada As Date
Dim fechaDesdePeriodo As Date
Dim fechaHastaPeriodo As Date
Dim fechaCorte As Date
Dim rsDiasCorresp As New ADODB.Recordset
Dim rsRegHorario As New ADODB.Recordset
Dim DiasHorario As Double
Dim fin_de_semana As Date
Dim fecDesdeSemana As Date
Dim fecHastaSemana As Date
Dim tieneFaseBaja As Boolean
Dim fechaBaja As Date
Dim cantDiasProp As Double

    'Activo el manejador de errores
    On Error GoTo CE
    
    fechaBaja = Empty
    Genera = False
    tieneFaseBaja = False
    
    'EAM- Obtiene la fecha de la ultima fecha de cálculo de dias corresp y la fase del empleado
    StrSql = "SELECT vdiascorcant,vdiascorcant,vdiasfechasta,bajfec,venc,estado FROM vacdiascor " & _
            "LEFT JOIN fases ON vacdiascor.ternro = fases.empleado" & _
            "  WHERE ternro= " & Ternro & " AND venc= 0 AND estado=-1 ORDER BY vdiasfechasta DESC"
    OpenRecordset StrSql, rsDiasCorresp
    
    If Not rsDiasCorresp.EOF Then
        ultFechaProcesada = rsDiasCorresp!vdiasfechasta
        cantdias = rsDiasCorresp!vdiascorcant
    Else
        StrSql = "SELECT * FROM vacdiascor WHERE ternro=" & Ternro & " ORDER BY vdiasfechasta DESC"
        OpenRecordset StrSql, rsDiasCorresp
        
        If Not rsDiasCorresp.EOF Then
            ultFechaProcesada = rsDiasCorresp!vdiasfechasta
            Flog.writeline "Ultima fecha de calculo de vacaciones " & rsDiasCorresp!vdiasfechasta
        Else
            ultFechaProcesada = fechaAlta
            Flog.writeline "No se encontro días correspondientes para el empleado y se toma la fecha de alta." & fechaAlta
        End If
    End If
    
    
    
    If (CDate(ultFechaProcesada) >= CDate(FechaHasta)) Then
        Flog.writeline "La fecha ingresada ya fue procesada. Ultimo fecha de procesamiento (" & ultFechaProcesada & ")"
        Exit Sub
    Else
       
       'Setea la proporcion de dias
       Call Politica(1501)
       
        Select Case FactorDivision
            'Empleados que tienen asignación horaria
            Case 0:
                'EAM- OBS- Siempre se procesa en la semana trabajada, si no se carga la semana trabajada se basa en la proyectada y no se
                'puede volver atras.
                
                'Busco el último movimiento en vigencia
                StrSql = " SELECT * FROM WC_MOV_HORARIOS WHERE ternro = " & Ternro & _
                         " AND fecdesde <= " & ConvFecha(FechaHasta) & " AND fechasta >= " & ConvFecha(FechaHasta) & _
                         " ORDER BY fecdesde desc, fechasta desc"
                 OpenRecordset StrSql, rsRegHorario
                 
                'EAM- Obtiene la fecha de la ultima fecha de cálculo de dias corresp y la fase del empleado
                StrSql = "SELECT vdiasfechasta,bajfec FROM vacdiascor " & _
                        "LEFT JOIN fases ON vacdiascor.ternro = fases.empleado" & _
                        "  WHERE ternro= " & Ternro & " AND venc= 0 AND estado=-1 ORDER BY vdiasfechasta DESC"
                OpenRecordset StrSql, rsDiasCorresp
                
                If Not rsDiasCorresp.EOF Then
                    If FechaHasta > rsDiasCorresp!bajfec Then
                        FechaHasta = CDate(rsDiasCorresp!bajfec)
                    End If
                End If
                rsDiasCorresp.Close
                
                'EAM- Busca la fecha hasta de vigencia del período
                StrSql = "SELECT * FROM vacacion WHERE vacfecdesde <= " & ConvFecha(ultFechaProcesada) & " AND vacfechasta >= " & ConvFecha(ultFechaProcesada) & _
                        " AND ternro= " & Ternro & " ORDER BY vacfecdesde DESC"
                OpenRecordset StrSql, rsDiasCorresp
                
                If Not rsDiasCorresp.EOF Then
                    If rsDiasCorresp!vacfechasta < FechaHasta Then
                        FechaHasta = rsDiasCorresp!vacfechasta
                        NroVac = rsDiasCorresp!vacnro
                    End If
                End If
                
               
                If Not rsRegHorario.EOF Then
                    fechaCorte = FechaHasta
                                
                    Do While ultFechaProcesada < fechaCorte
                        fechadesde = DateAdd("d", 1, ultFechaProcesada)
                        
                        'EAM- Obtiene la fecha de fin de semana
                        If IsDate(fechadesde) Then
                            fecHastaSemana = DateAdd("d", 7 - (Weekday(fechadesde) - 1), fechadesde)
                            fecDesdeSemana = DateAdd("d", -6, fecHastaSemana)
                        End If
    
        
                        If CDate(fecHastaSemana) < CDate(fechaCorte) Then
                            FechaHasta = fecHastaSemana
                        Else
                            FechaHasta = fechaCorte
                            fecHastaSemana = fechaCorte
                        End If
                        
                        
                        'Busco el ultimo movimiento en vigencia
                        StrSql = " SELECT DISTINCT fechor FROM WC_MOV_HORARIOS WHERE ternro = " & Ternro & _
                                " AND fechor  <= " & ConvFecha(fecHastaSemana) & " AND fechor >= " & ConvFecha(fechadesde)
                        OpenRecordset StrSql, rsRegHorario
        
                        If Not rsRegHorario.EOF Then
        
                            'EAM- Calcula la proporcion de los días trabajado en el rango de 2 semana
                            DiasHorario = DiasHorario + rsRegHorario.RecordCount
        
'                            'EAM- Obtiene la fecha de la ultima fecha de cálculo de dias corresp y la fase del empleado
'                            StrSql = "SELECT vdiasfechasta,bajfec FROM vacdiascor " & _
'                                    "LEFT JOIN fases ON vacdiascor.ternro = fases.empleado" & _
'                                    "  WHERE ternro= " & Ternro & " AND venc= 0 AND estado=-1 ORDER BY vdiasfechasta DESC"
'                            OpenRecordset StrSql, rsDiasCorresp
'                            rsDiasCorresp.Close
                        End If
                        
                
                        totalDias = totalDias + (DateDiff("d", fechadesde, FechaHasta) + 1)
                        ultFechaProcesada = FechaHasta
                        Genera = True
                    Loop
                Else
                    Flog.writeline "No se encontro asignación horaria para la fecha: " & FechaHasta
                    FechaHasta = ultFechaProcesada
                    rsDiasCorresp.Close
                    GoTo sinDatos
                End If
            
                cantDiasProp = Round(totalDias / 350 * 2 * DiasHorario, 4)
            Case 30:
                'EAM- Se proporciona 1 día vacación cada 30 días trabajados.
                'EAM- Busca la fecha hasta de vigencia del período
                StrSql = "SELECT * FROM vacacion WHERE vacfecdesde <= " & ConvFecha(ultFechaProcesada) & " AND vacfechasta >= " & ConvFecha(ultFechaProcesada) & _
                        " AND ternro= " & Ternro & " ORDER BY vacfecdesde DESC"
                OpenRecordset StrSql, rsDiasCorresp
                
                If Not rsDiasCorresp.EOF Then
                    If rsDiasCorresp!vacfechasta < FechaHasta Then
                        FechaHasta = rsDiasCorresp!vacfechasta
                        NroVac = rsDiasCorresp!vacnro
                    End If
                End If
                Flog.writeline "Empleado que proporcionan 1 día por 30 trabajados."
                'ultFechaProcesada
                                
                cantDiasProp = Int((DateDiff("d", ultFechaProcesada, FechaHasta) / 30))
                
                If cantDiasProp > 0 Then
                    FechaHasta = DateAdd("d", (cantDiasProp * 30), ultFechaProcesada)
                    Genera = True
                Else
                    FechaHasta = ultFechaProcesada
                    Genera = False
                End If
                
                
        
            Case Else:
                'EAM- Empleados Mensuales, su asignacion horaria se rige por el tipo de vacacion
                'EAM- Busca la fecha hasta de vigencia del período
                StrSql = "SELECT * FROM vacacion WHERE vacfecdesde <= " & ConvFecha(ultFechaProcesada) & " AND vacfechasta > " & ConvFecha(ultFechaProcesada) & _
                        " AND ternro= " & Ternro & " ORDER BY vacfecdesde DESC"
                OpenRecordset StrSql, rsDiasCorresp
                
                If Not rsDiasCorresp.EOF Then
                    If rsDiasCorresp!vacfechasta < FechaHasta Then
                        FechaHasta = rsDiasCorresp!vacfechasta
                        NroVac = rsDiasCorresp!vacnro
                    End If
                End If
                
                Flog.writeline "Empleado Mensuales. Busca los días laborables del tipo de vacacion: " & TipoVacacionProporcion
                
                StrSql = " SELECT * FROM tipovacac WHERE tipvacnro = " & TipoVacacionProporcion
                OpenRecordset StrSql, rsRegHorario
                
                'EAM- Obtiene la cantidad de dias laborables
                DiasHorario = 0
                Do While Not rsRegHorario.EOF
                    If rsRegHorario!tpvhabiles__1 Then DiasHorario = DiasHorario + 1
                    If rsRegHorario!tpvhabiles__2 Then DiasHorario = DiasHorario + 1
                    If rsRegHorario!tpvhabiles__3 Then DiasHorario = DiasHorario + 1
                    If rsRegHorario!tpvhabiles__4 Then DiasHorario = DiasHorario + 1
                    If rsRegHorario!tpvhabiles__5 Then DiasHorario = DiasHorario + 1
                    If rsRegHorario!tpvhabiles__6 Then DiasHorario = DiasHorario + 1
                    If rsRegHorario!tpvhabiles__7 Then DiasHorario = DiasHorario + 1
                    rsRegHorario.MoveNext
                Loop
                
                totalDias = DateDiff("d", ultFechaProcesada, FechaHasta)
                cantDiasProp = Round(totalDias / 350 * 2 * DiasHorario, 4)
                Genera = True
        End Select
        
        'Obtiene el saldo del período al cual calculó los dias correspondientes
         StrSql = " SELECT vdiascorcant FROM vacdiascor " & _
                "  WHERE ternro= " & Ternro & " AND venc= 0 AND vacnro=" & NroVac & " ORDER BY vdiasfechasta DESC"
        OpenRecordset StrSql, rsDiasCorresp
        
        If Not rsDiasCorresp.EOF Then
            cantdias = rsDiasCorresp!vdiascorcant
        Else
            cantdias = 0
        End If
        
        Flog.writeline "Cantidad de días acumulados: " & cantdias
        Flog.writeline "Total de días a procesar: " & totalDias
        Flog.writeline "Días Laborables: " & DiasHorario
        
        'cantDias = cantDias + Round(totalDias / 365 * 2 * DiasHorario, 4)
        cantdias = cantdias + cantDiasProp
        Flog.writeline "Cantidad de días generados: " & cantdias
        
        ultFechaProcesada = FechaHasta
        
    End If

    GoTo finalizado
CE:
    Flog.writeline " ------------------------------------------------------------"
    Flog.writeline "Error procesando Empleado:" & Ternro & " Año " & Anio
    Flog.writeline Err.Description
    Flog.writeline "SQL: " & StrSql
    Flog.writeline " ------------------------------------------------------------"

sinDatos:
    Exit Sub
finalizado:
    Set rsDiasCorresp = Nothing
    Set rsRegHorario = Nothing
End Sub

Sub arrmarArregloPeriodoDepurado(ByRef arrIntervalos, ByVal strAux As String)
 Dim i, j As Long
 Dim strArrAux
 Dim strAuxdatos
 Dim strDepurado
 
    strArrAux = Split(strAux, "@@")
    'depura el string
    For i = 1 To UBound(strArrAux)
        strAuxdatos = Split(strArrAux(i), "@")
        If CDate(strAuxdatos(0)) <= CDate(strAuxdatos(1)) Then
            strDepurado = strDepurado & "@@" & strAuxdatos(0) & "@" & strAuxdatos(1) & "@" & strAuxdatos(2) & "@" & strAuxdatos(3)
        End If
    Next
    
    strArrAux = Split(strDepurado, "@@")
   If UBound(strArrAux) > 0 Then 'mdf
    ReDim arrIntervalos(UBound(strArrAux), 4)
    
    For i = 1 To UBound(strArrAux)
        strAuxdatos = Split(strArrAux(i), "@")
        If CDate(strAuxdatos(0)) <= CDate(strAuxdatos(1)) Then
            For j = 0 To UBound(strAuxdatos)
                arrIntervalos(i, j + 1) = strAuxdatos(j)
            Next
        End If
    Next
   End If
End Sub


Sub arrmarArregloPeriodo(ByRef arrIntervalos, ByVal strAux As String)
 Dim i, j As Long
 Dim strArrAux
 Dim strAuxdatos

 
    strArrAux = Split(strAux, "@@")
    ReDim arrIntervalos(UBound(strArrAux), 4)
    
    For i = 1 To UBound(strArrAux)
        strAuxdatos = Split(strArrAux(i), "@")
'        If CDate(strAuxdatos(0)) <= CDate(strAuxdatos(1)) Then
            For j = 0 To UBound(strAuxdatos)
                arrIntervalos(i, j + 1) = strAuxdatos(j)
            Next
'        End If
    Next
End Sub
'EAM- Arma los nuevos intervalos que se generan de solaparse los movimientos
Sub generarIntervaloMovGenerado(ByVal Ternro As Long, ByVal ultFecProcesamiento As Date, ByVal fechaHastaProcesamiento As Date, ByRef arrIntervalos, ByRef ErrorDatos As Boolean)
 Dim rsMovimiento As New ADODB.Recordset
 Dim strAux As String
 Dim i, j As Long
 Dim strAuxdatos
 Dim Intervalos
 
    ErrorDatos = False
    
        
    StrSql = "SELECT fecdesde, fechasta FROM wc_mov_horarios " & _
            " WHERE fecdesde <= " & ConvFecha(fechaHastaProcesamiento) & " and fechasta >= " & ConvFecha(ultFecProcesamiento) & _
            " AND ternro = " & Ternro & _
            " GROUP BY fecdesde, fechasta" & _
            " ORDER BY fecdesde ASC, fechasta ASC"
    OpenRecordset StrSql, rsMovimiento
    
    ReDim Intervalos(rsMovimiento.RecordCount, 3)
    
    i = 1
    Do While Not rsMovimiento.EOF
        Intervalos(i, 1) = rsMovimiento!fecDesde
        Intervalos(i, 2) = rsMovimiento!fecHasta
        Intervalos(i, 3) = rsMovimiento!fecDesde & "@" & rsMovimiento!fecHasta
        rsMovimiento.MoveNext
        i = i + 1
    Loop
    
    If rsMovimiento.RecordCount <= 0 Then
        Flog.writeline "No se encontraron períodos para Analizar. "
        ErrorDatos = True
        Exit Sub
    End If
    
    
    If ultFecProcesamiento < CDate(Intervalos(1, 1)) Then
       ErrorDatos = True
    Else
        If UBound(Intervalos, 1) = 1 Then
            If (CDate(Intervalos(1, 2)) >= fechaHastaProcesamiento) And (CDate(Intervalos(1, 1)) < fechaHastaProcesamiento) Then
                strAux = strAux & "@@" & ultFecProcesamiento & "@" & fechaHastaProcesamiento & "@" & Intervalos(1, 3)
            Else
                ErrorDatos = True
            End If
        Else
            'EAM- Asigno el primer registro para empezar a comparar
            strAux = strAux & "@@" & ultFecProcesamiento & "@" & Intervalos(1, 2) & "@" & Intervalos(1, 3)
            Call arrmarArregloPeriodo(arrIntervalos, strAux)
            
            For i = 2 To UBound(Intervalos, 1)
                Call arrmarArregloPeriodo(arrIntervalos, strAux)
                For j = 1 To UBound(arrIntervalos, 1)
                    
                    If CDate(Intervalos(i, 1)) < CDate(arrIntervalos(j, 2)) Then
                        'cambia la fecha de finalizacion por la del periodo siguiente. Se solapa
                        Dim fechaAuxiliar As Date
                        Dim cantAux As Long
                        fechaAuxiliar = arrIntervalos(j, 2)
                                         
                                         
                        If ultFecProcesamiento > CDate(Intervalos(i, 1)) Then
                            arrIntervalos(j, 2) = DateAdd("d", -1, ultFecProcesamiento)
                           ' Intervalos(i, 1) = DateAdd("d", 1, ultFecProcesamiento)
                            Intervalos(i, 1) = ultFecProcesamiento
                        Else
                            arrIntervalos(j, 2) = DateAdd("d", -1, Intervalos(i, 1))
                        End If
                       
                        strAux = ""
                        For cantAux = 1 To j
                            strAux = strAux & "@@" & arrIntervalos(cantAux, 1) & "@" & arrIntervalos(cantAux, 2) & "@" & arrIntervalos(cantAux, 3) & "@" & arrIntervalos(cantAux, 4)
                        Next
                                                
                        'El periodo cae dentro y se forma 2 nuevos sub-intervalo
                        If Intervalos(i, 2) < fechaAuxiliar And (fechaHastaProcesamiento > Intervalos(i, 2)) Then
                            strAux = strAux & "@@" & Intervalos(i, 1) & "@" & Intervalos(i, 2) & "@" & Intervalos(i, 3)
                            
                            If fechaHastaProcesamiento > fechaAuxiliar Then
                                strAux = strAux & "@@" & DateAdd("d", 1, Intervalos(i, 2)) & "@" & fechaAuxiliar & "@" & arrIntervalos(j, 3) & "@" & arrIntervalos(j, 4)
                            Else
                                strAux = strAux & "@@" & DateAdd("d", 1, Intervalos(i, 2)) & "@" & fechaHastaProcesamiento & "@" & arrIntervalos(j, 3) & "@" & arrIntervalos(j, 4)
                            End If
                        Else
                            'strAux = strAux & "@@" & Intervalos(i, 1) & "@" & fechaHastaProcesamiento & "@" & Intervalos(i, 3)
                            strAux = strAux & "@@" & Intervalos(i, 1) & "@" & Intervalos(i, 2) & "@" & Intervalos(i, 3)
                        End If
                        
                        
                    Else
                        'Pregunto si el final de ese intervalo es mayor que el inicio del que estoy analizando porque si es asi se genera otro intervalo
                        
                        'EAM- si es igual se da este caso
                        '| -----------------------------------------|
                        '       |--------|          |-------------|
                        '   UP                      NP
                        If CDate(Intervalos(i, 1)) = CDate(arrIntervalos(j, 2)) Then
                            arrIntervalos(j, 2) = DateAdd("d", -1, Intervalos(i, 1))
                            
                            strAux = ""
                            For cantAux = 1 To j
                                strAux = strAux & "@@" & arrIntervalos(cantAux, 1) & "@" & arrIntervalos(cantAux, 2) & "@" & arrIntervalos(cantAux, 3) & "@" & arrIntervalos(cantAux, 4)
                            Next
                            
                            If CDate(Intervalos(i, 2)) < fechaHastaProcesamiento Then
                                strAux = strAux & "@@" & Intervalos(i, 1) & "@" & Intervalos(i, 2) & "@" & Intervalos(i, 3)
                            Else
                                strAux = strAux & "@@" & Intervalos(i, 1) & "@" & fechaHastaProcesamiento & "@" & Intervalos(i, 3)
                            End If
                        Else
                            'EAM- se puede dar por periodos consecutivos
                            If CDate(Intervalos(i, 1)) = CDate(DateAdd("d", 1, arrIntervalos(j, 2))) Then
                                'EAM- si es mayor cae adentro
                                If CDate(Intervalos(i, 2)) > fechaHastaProcesamiento Then
                                    strAux = strAux & "@@" & Intervalos(i, 1) & "@" & fechaHastaProcesamiento & "@" & Intervalos(i, 3)
                                Else
                                    strAux = strAux & "@@" & Intervalos(i, 1) & "@" & Intervalos(i, 2) & "@" & Intervalos(i, 3)
                                End If
                                
                            End If
                            
                        End If
                    End If
                Next
            Next
        End If
        Call arrmarArregloPeriodoDepurado(arrIntervalos, strAux)
        If strAux <> "" Then 'mdf
            If CDate(arrIntervalos(UBound(arrIntervalos, 1), 2)) < CDate(fechaHastaProcesamiento) Then
                ErrorDatos = True
            Else
                arrIntervalos(UBound(arrIntervalos, 1), 2) = fechaHastaProcesamiento
    '            For i = 1 To UBound(arrIntervalos, 1) - 1
    '                If (DateAdd("d", 1, arrIntervalos(i, 2)) = CDate(arrIntervalos(i + 1, 1))) Then
    '                    ErrorDatos = False
    '                End If
    '
    '            Next
            End If
        Else 'mdf
         ErrorDatos = True
        End If
    End If
     
    If Not ErrorDatos Then
        Flog.writeline "-------------------------- Intervalos Perídos ---------------------------------------"
        For i = 1 To UBound(arrIntervalos, 1)
            Flog.writeline "Fecha Desde: " & arrIntervalos(i, 1) & " Fecha Hasta " & arrIntervalos(i, 2) & " Fecha Desde Período: " & arrIntervalos(i, 3) & " Fecha Hasta Período: " & arrIntervalos(i, 4)
            strAux = strAux & "@@" & arrIntervalos(i, 1) & "@" & arrIntervalos(i, 2) & "@" & arrIntervalos(i, 3) & "@" & arrIntervalos(i, 4)
        Next
        Flog.writeline "-------------------------------------------------------------------------------------"
    Else
        ReDim arrIntervalos(0, 0)
        Flog.writeline "No se encontro movimientos para la fecha, o los movimientos no son consecutivos. "
    End If

End Sub


'Busca el periodod de vacaciones para Costa Rica
Sub BuscarPeriodoVacCR(ByVal fechaProcesar As Date, ByVal Ternro As Long, ByRef NroVac As Long, ByRef FechaHasta As Date)
 Dim rsVacCR As New ADODB.Recordset
 
   'EAM- Busca la fecha hasta de vigencia del período
    StrSql = "SELECT * FROM vacacion WHERE vacfecdesde <= " & ConvFecha(fechaProcesar) & " AND vacfechasta >= " & ConvFecha(fechaProcesar) & _
            " AND ternro= " & Ternro & " ORDER BY vacfecdesde DESC"
    OpenRecordset StrSql, rsVacCR
                
    If Not rsVacCR.EOF Then
        If (rsVacCR!vacfechasta >= fechaProcesar) Then
            FechaHasta = rsVacCR!vacfechasta
            NroVac = rsVacCR!vacnro
        End If
    End If
End Sub


'EAM- Procedimiento que actualiza los dias correspondientes para CR
Sub actualizarDiasVacCR(ByVal auxNroVac As Long, ByVal DiasCorraGen As Double, ByVal NroTPV As Long, ByVal FechaHasta As Date)
 Dim rs As New ADODB.Recordset

    StrSql = "SELECT * FROM tipovacac WHERE tpvnrocol = " & NroTPV
    OpenRecordset StrSql, rs
    
    Flog.writeline "Tipovac-NroTPV:" & StrSql
    If Not rs.EOF Then
        NroTPV = rs!tipvacnro
        rs.Close
    Else
        'EAM- Verifica si tiene el tipo de días de vacaciones configurado Pol(1501)
        'sino pone el Primero de la tabla por Default
        If (st_TipoDia1 > 0) Then
            NroTPV = st_TipoDia1
        Else
            NroTPV = 1 ' por default
        End If
    End If

    Flog.writeline "NroTPV:" & NroTPV
    StrSql = "SELECT * FROM vacdiascor WHERE vacnro = " & auxNroVac & " AND Ternro = " & Ternro & " AND (venc = 0 OR venc IS NULL)"
    OpenRecordset StrSql, rs
     Flog.writeline "seteo NroTPV. Busca:" & StrSql
    If Not rs.EOF Then
        If Reproceso Then
             Flog.writeline "Existe vacnro y reprocesa"
            If Not IsNull(NroTPV) And NroTPV > 0 Then
                StrSql = "UPDATE vacdiascor SET vdiascormanual = 0, vdiascorcant = " & DiasCorraGen & ", tipvacnro = " & NroTPV
                    
                'EAM- Se agrego para CR el campo de la ultima generacion de Dias Correspondiente
                If Not IsNull(FechaHasta) Then
                    StrSql = StrSql & ",vdiasfechasta = " & ConvFecha(FechaHasta)
                End If
                StrSql = StrSql & " WHERE vacnro = " & auxNroVac & " AND Ternro = " & Ternro & " AND (venc = 0 OR venc IS NULL)"
                    
            Else
               Flog.writeline "Error al actualizar los dias correspondientes. Tipo de vacación incorrecto: " & NroTPV
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
        Else
           Flog.writeline "Existe vacnro y no reprocesa"
            If Not IsNull(NroTPV) And Not NroTPV > 0 Then
                StrSql = "INSERT INTO vacdiascor(vacnro,vdiascorcant,vdiascormanual,ternro,tipvacnro,vdiasfechasta) VALUES (" & _
                         auxNroVac & "," & DiasCorraGen & ",0," & Ternro & "," & NroTPV & "," & ConvFecha(FechaHasta) & ")"
            Else
                Flog.writeline "Error al insertar los dias correspondientes. Tipo de vacación incorrecto: " & NroTPV
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
    
    Else
        Flog.writeline "no Existe vacnro"
        If Not IsNull(NroTPV) And Not NroTPV > 0 Then
             StrSql = "INSERT INTO vacdiascor(vacnro,vdiascorcant,vdiascormanual,ternro,tipvacnro,vdiasfechasta) VALUES (" & _
                      auxNroVac & "," & DiasCorraGen & ",0," & Ternro & "," & NroTPV & "," & ConvFecha(FechaHasta) & ")"
                Flog.writeline "Rompe1:" & StrSql
         Else
            StrSql = "INSERT INTO vacdiascor(vacnro,vdiascorcant,vdiascormanual,ternro,tipvacnro,vdiasfechasta) VALUES (" & _
                      auxNroVac & "," & DiasCorraGen & ",0," & Ternro & ",1," & ConvFecha(FechaHasta) & ")"
                       Flog.writeline "Rompe2:" & StrSql
         End If
         objConn.Execute StrSql, , adExecuteNoRecords
    End If
End Sub
Sub actualizarDiasVacPY(ByVal auxNroVac As Long, ByVal DiasCorraGen As Double, ByVal NroTPV As Long, ByVal FechaHasta As Date)
 Dim rs As New ADODB.Recordset

    StrSql = "SELECT * FROM tipovacac WHERE tpvnrocol = " & NroTPV
    OpenRecordset StrSql, rs
        
    If Not rs.EOF Then
        NroTPV = rs!tipvacnro
    Else
        'Verifica si tiene el tipo de días de vacaciones configurado Pol(1501)
        'sino pone el Primero de la tabla por Default
        If (st_TipoDia1 > 0) Then
            NroTPV = st_TipoDia1
        Else
            NroTPV = 1 ' por default
        End If
    End If

 
    StrSql = "SELECT * FROM vacdiascor WHERE vacnro = " & auxNroVac & " AND Ternro = " & Ternro & " AND (venc = 0 OR venc IS NULL)"
    OpenRecordset StrSql, rs
        
    If Not rs.EOF Then
        If Reproceso Then
            If Not IsNull(NroTPV) And NroTPV > 0 Then
                StrSql = "UPDATE vacdiascor SET vdiascormanual = 0, vdiascorcant = " & DiasCorraGen & ", tipvacnro = " & NroTPV
                    
                ' Se agrego para CR el campo de la ultima generacion de Dias Correspondiente
                If Not IsNull(FechaHasta) Then
                    StrSql = StrSql & ",vdiasfechasta = " & ConvFecha(FechaHasta)
                End If
                StrSql = StrSql & " WHERE vacnro = " & auxNroVac & " AND Ternro = " & Ternro & " AND (venc = 0 OR venc IS NULL)"
                    
            Else
               Flog.writeline "Error al actualizar los dias correspondientes. Tipo de vacación incorrecto: " & NroTPV
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
        Else
            If Not IsNull(NroTPV) And Not NroTPV > 0 Then
                StrSql = "INSERT INTO vacdiascor(vacnro,vdiascorcant,vdiascormanual,ternro,tipvacnro,vdiasfechasta) VALUES (" & _
                         auxNroVac & "," & DiasCorraGen & ",0," & Ternro & "," & NroTPV & "," & ConvFecha(FechaHasta) & ")"
            Else
                Flog.writeline "Error al insertar los dias correspondientes. Tipo de vacación incorrecto: " & NroTPV
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
    
    Else
        If Not IsNull(NroTPV) And Not NroTPV > 0 Then
             StrSql = "INSERT INTO vacdiascor(vacnro,vdiascorcant,vdiascormanual,ternro,tipvacnro,vdiasfechasta) VALUES (" & _
                      auxNroVac & "," & DiasCorraGen & ",0," & Ternro & "," & NroTPV & "," & ConvFecha(FechaHasta) & ")"
         Else
            StrSql = "INSERT INTO vacdiascor(vacnro,vdiascorcant,vdiascormanual,ternro,tipvacnro,vdiasfechasta) VALUES (" & _
                      auxNroVac & "," & DiasCorraGen & ",0," & Ternro & "," & NroTPV & "," & ConvFecha(FechaHasta) & ")"
         End If
         objConn.Execute StrSql, , adExecuteNoRecords
    End If
End Sub

Public Sub bus_DiasVac_CR(ByVal Ternro As Long, ByRef NroVac As Long, ByVal fechaAlta As Date, ByRef FechaHasta As Date, ByRef cantdias As Double, ByRef Columna As Integer, ByRef Mensaje As String, ByRef Genera As Boolean, ByVal Anio As Long, Optional ByVal generarRegVacDiasCor As Boolean)
' ---------------------------------------------------------------------------------------------
' Descripcion: Calcula los días de vacaciones
' Autor      : Margiotta, Emanuel
' Fecha      :
' Ultima Mod.: 16/05/2012 - Javier Irastorza
' Descripcion: Se cambio la manera de determinar los días hábiles de la jornada.
' ---------------------------------------------------------------------------------------------
Dim totalDias As Long

Dim ultFechaProcesada As String
Dim fechaCorte As Date
Dim rsDiasCorresp As New ADODB.Recordset
Dim rsRegHorario As New ADODB.Recordset
Dim DiasHorario As Double
Dim cantDiasProp As Double
Dim arrIntervalos()
Dim i As Integer
Dim actualizarDatos As Boolean
Dim IntervalosIncorrectos As Boolean
Dim cantDia As Integer
Dim cantMes As Integer
Dim cantAnio As Integer

    'Activo el manejador de errores
    On Error GoTo CE
    
    'fechaBaja = Empty
    Genera = False
    'tieneFaseBaja = False
    
    'EAM- Obtiene la fecha de la ultima fecha de cálculo de dias corresp y la fase del empleado
    
    '----------------------
    StrSql = "SELECT vdiascorcant,vdiascorcant,vdiasfechasta,bajfec,venc,estado FROM vacdiascor " & _
            "LEFT JOIN fases ON vacdiascor.ternro = fases.empleado" & _
            "  WHERE ternro= " & Ternro & " AND venc= 0 AND estado=-1 ORDER BY vdiasfechasta DESC"
    OpenRecordset StrSql, rsDiasCorresp
    
    If Not rsDiasCorresp.EOF Then
        ultFechaProcesada = rsDiasCorresp!vdiasfechasta
        cantdias = rsDiasCorresp!vdiascorcant
        Flog.writeline "Ultima fecha de calculo de vacaciones " & rsDiasCorresp!vdiasfechasta
    Else
        StrSql = "SELECT * FROM vacdiascor WHERE ternro=" & Ternro & " ORDER BY vdiasfechasta DESC"
        OpenRecordset StrSql, rsDiasCorresp
        
        If Not rsDiasCorresp.EOF Then
            ultFechaProcesada = rsDiasCorresp!vdiasfechasta
            Flog.writeline "Ultima fecha de calculo de vacaciones " & rsDiasCorresp!vdiasfechasta
        Else
            ultFechaProcesada = fechaAlta
            Flog.writeline "No se encontro días correspondientes para el empleado y se toma la fecha de alta." & fechaAlta
        End If
    End If
    
    
    
    If (CDate(ultFechaProcesada) >= CDate(FechaHasta)) Then
        Flog.writeline "La fecha ingresada ya fue procesada. Ultimo fecha de procesamiento (" & ultFechaProcesada & ")"
        Exit Sub
    Else
       
       'Setea la proporcion de dias. Asume que si no tiene alcance el empleado tiene parte, osea FactorDivision = 0
       FactorDivision = 0
       Call Politica(1501)
       actualizarDatos = False
       
        Select Case FactorDivision
            'Empleados que tienen asignación horaria
            Case 0:
                If (generarRegVacDiasCor) Then
                    If TipoVacacionProporcion = 0 Then
                        StrSql = "UPDATE vacdiascor SET tipvacnro = 1 WHERE vacnro = " & NroVac & " AND Ternro = " & Ternro & " AND (venc = 0 OR venc IS NULL)"
                    Else
                        StrSql = "UPDATE vacdiascor SET tipvacnro = " & TipoVacacionProporcion & " WHERE vacnro = " & NroVac & " AND Ternro = " & Ternro & " AND (venc = 0 OR venc IS NULL)"
                    End If
                    objConn.Execute StrSql, , adExecuteNoRecords
                    Flog.writeline "Se actualiza el tipo de vacacion para el vacnro: " & NroVac
                End If
                'EAM- OBS- Siempre se procesa en la semana trabajada, si no se carga la semana trabajada se basa en la proyectada y no se
                'puede volver atras.
                Call generarIntervaloMovGenerado(Ternro, DateAdd("d", 1, ultFechaProcesada), FechaHasta, arrIntervalos, IntervalosIncorrectos)
                
                'EAM- Si intervalo no es correcto aborta el procesamiento. Esto se chequea en el procedimiento generarIntervaloMovGenerado
                If IntervalosIncorrectos Then
                    Flog.writeline "Los intervalos de los movimientos no son correctos."
                    Exit Sub
                End If
                
                'EAM- Obtiene la fecha de la ultima fecha de cálculo de dias corresp y la fase del empleado
                StrSql = "SELECT vdiasfechasta,bajfec FROM vacdiascor " & _
                        "LEFT JOIN fases ON vacdiascor.ternro = fases.empleado" & _
                        "  WHERE ternro= " & Ternro & " AND venc= 0 AND estado=-1 ORDER BY vdiasfechasta DESC"
                OpenRecordset StrSql, rsDiasCorresp

                If Not rsDiasCorresp.EOF Then
                    If FechaHasta > rsDiasCorresp!bajfec Then
                        FechaHasta = CDate(rsDiasCorresp!bajfec)
                    End If
                End If
                rsDiasCorresp.Close
                
'                'EAM- Busca la fecha hasta de vigencia del período
'                StrSql = "SELECT * FROM vacacion WHERE vacfecdesde <= " & ConvFecha(ultFechaProcesada) & " AND vacfechasta >= " & ConvFecha(ultFechaProcesada) & _
'                        " AND ternro= " & Ternro & " ORDER BY vacfecdesde DESC"
'                OpenRecordset StrSql, rsDiasCorresp
'
'                If Not rsDiasCorresp.EOF Then
'                    'EAM- Se agrego la validacion (DateAdd("d", 1, rsDiasCorresp!vacfechasta) < FechaHasta) para verificar que no sea el ultimo dia del período del empleado sino corta siempre esa fecha. 10/07/2012
'                    If (rsDiasCorresp!vacfechasta < FechaHasta) And (rsDiasCorresp!vacfechasta > ultFechaProcesada) Then
'                        FechaHasta = rsDiasCorresp!vacfechasta
'                        NroVac = rsDiasCorresp!vacnro
'                    End If
'                End If
                
                'EAM- Se agrego el log para saber hasta donde se procesará. 10/07/2012
                Flog.writeline "Última fecha Procesamiento: " & ultFechaProcesada & "."
                Flog.writeline "Nueva fecha Procesamiento: " & FechaHasta & "."
               
                fechaCorte = CDate(Day(fechaAlta) & "/" & Month(fechaAlta) & "/" & Anio)
                ultFechaProcesada = arrIntervalos(UBound(arrIntervalos, 1), 2)
                'fechaCorte = FechaHasta
                If (UBound(arrIntervalos, 1) > 0) Then
                    For i = UBound(arrIntervalos, 1) To 1 Step -1


                        If CDate(arrIntervalos(i, 2)) < fechaCorte Then
                            If fechaCorte >= arrIntervalos(i, 1) Then
                            
                                'EAM- (v3.29) bloque if then se agrego y sus 3 lineas.
                                If i > 0 And fechaCorte > arrIntervalos(i, 1) Then
                                    actualizarDatos = True
                                    fechaCorte = DateAdd("yyyy", -1, fechaCorte)
                                    i = i + 1
                                Else
                                
                                StrSql = " SELECT DISTINCT fechor FROM WC_MOV_HORARIOS WHERE ternro = " & Ternro & _
                                        " AND fechor  <= " & ConvFecha(arrIntervalos(i, 2)) & " AND fechor >= " & ConvFecha(arrIntervalos(i, 1)) & " AND fecdesde= " & ConvFecha(arrIntervalos(i, 3)) & " AND fechasta= " & ConvFecha(arrIntervalos(i, 4))
                                OpenRecordset StrSql, rsRegHorario

                                Flog.writeline "----------------- Nuevo Intervalos Perídos por Cambio de período -----------------"
                                Flog.writeline "Fecha Desde: " & arrIntervalos(i, 1) & " Fecha Hasta " & arrIntervalos(i, 2) & " Fecha Desde Período: " & arrIntervalos(i, 3) & " Fecha Hasta Período: " & arrIntervalos(i, 4)
                                Flog.writeline "----------------------------------------------------------------------------------"
                                'arrIntervalos(i, 1) = DateAdd("d", 1, fechaCorte)
                                'i = i + 1

                                'Setea para que actualize los dias del período
                                'actualizarDatos = True
                                'EAM- Imputa a Período Nuevo
                                Call BuscarPeriodoVacCR(arrIntervalos(i, 2), Ternro, NroVac, fechaCorte)
                                
                                totalDias = totalDias + (DateDiff("d", arrIntervalos(i, 1), arrIntervalos(i, 2)) + 1)
                                ultFechaProcesada = arrIntervalos(i, 2)
                                End If
                            Else
                                'EAM- Busca los movimientos asignados para el intervalo de fechas y periodos q corresponden.
                                StrSql = " SELECT DISTINCT fechor FROM WC_MOV_HORARIOS WHERE ternro = " & Ternro & _
                                        " AND fechor  <= " & ConvFecha(arrIntervalos(i, 2)) & " AND fechor >= " & ConvFecha(DateAdd("d", 1, fechaCorte)) & " AND fecdesde= " & ConvFecha(arrIntervalos(i, 3)) & " AND fechasta= " & ConvFecha(arrIntervalos(i, 4))
                                OpenRecordset StrSql, rsRegHorario

                                'EAM- Imputa a Período Nuevo
                                Call BuscarPeriodoVacCR(arrIntervalos(i, 2), Ternro, NroVac, fechaCorte)
                                
                                totalDias = totalDias + (DateDiff("d", arrIntervalos(i, 1), arrIntervalos(i, 2)) + 1)
                                ultFechaProcesada = arrIntervalos(i, 2)
                            End If
                            
                            
                        Else

                            If CDate(fechaCorte) < CDate(arrIntervalos(i, 2)) Then
                                
                                If CDate(fechaCorte) <= CDate(arrIntervalos(i, 1)) Then
                                    'EAM- Busca los movimientos asignados para el intervalo de fechas y periodos q corresponden.
                                    StrSql = " SELECT DISTINCT fechor FROM WC_MOV_HORARIOS WHERE ternro = " & Ternro & _
                                            " AND fechor  <= " & ConvFecha(arrIntervalos(i, 2)) & " AND fechor >= " & ConvFecha(arrIntervalos(i, 1)) & " AND fecdesde= " & ConvFecha(arrIntervalos(i, 3)) & " AND fechasta= " & ConvFecha(arrIntervalos(i, 4))
                                    
                                    totalDias = totalDias + (DateDiff("d", arrIntervalos(i, 1), arrIntervalos(i, 2)) + 1)
                                Else
                                    StrSql = " SELECT DISTINCT fechor FROM WC_MOV_HORARIOS WHERE ternro = " & Ternro & _
                                            " AND fechor  <= " & ConvFecha(arrIntervalos(i, 2)) & " AND fechor >= " & ConvFecha(fechaCorte) & " AND fecdesde= " & ConvFecha(arrIntervalos(i, 3)) & " AND fechasta= " & ConvFecha(arrIntervalos(i, 4))
                                    totalDias = totalDias + (DateDiff("d", fechaCorte, arrIntervalos(i, 2)) + 1)
                                    arrIntervalos(i, 2) = DateAdd("d", -1, fechaCorte)
                                    i = i + 1
                                    actualizarDatos = True
                                End If
                                'ultFechaProcesada = arrIntervalos(i, 2)
                            End If
                            OpenRecordset StrSql, rsRegHorario
                            
                        End If


                        'EAM- (v3.29) Se agrego condicion acualizadatos.
                        If Not rsRegHorario.EOF And Not actualizarDatos Then
                            'EAM- Calcula la proporcion de los días trabajado en el rango de 2 semana
                            DiasHorario = DiasHorario + rsRegHorario.RecordCount
                        End If


                        'ultFechaProcesada = FechaHasta
                        If actualizarDatos Then
                            'Obtiene el saldo del período al cual calculó los dias correspondientes
                             StrSql = " SELECT vdiascorcant FROM vacdiascor " & _
                                    "  WHERE ternro= " & Ternro & " AND venc= 0 AND vacnro=" & NroVac & " ORDER BY vdiasfechasta DESC"
                            OpenRecordset StrSql, rsDiasCorresp

                            If Not rsDiasCorresp.EOF Then
                                cantdias = rsDiasCorresp!vdiascorcant
                            Else
                                cantdias = 0
                            End If

                            cantDiasProp = Round(DiasHorario * (14 / 364), 4)
                            DiasHorario = 0
                            cantdias = cantdias + cantDiasProp

                            'Primero acutalizo
                            actualizarDatos = False
                            Call actualizarDiasVacCR(NroVac, cantdias, TipoVacacionProporcion, ultFechaProcesada)
                            Flog.writeline "Se actualizo los días correspondientes del vacnro : " & NroVac & " Cant. Dias: " & cantdias & " Fecha: " & ultFechaProcesada
                            Flog.writeline ""
                            cantDiasProp = 0
                            Genera = True
                        End If
                       ' Loop
                        Next

                Else
                    Flog.writeline "No se encontro asignación horaria para la fecha: " & FechaHasta
                    FechaHasta = ultFechaProcesada
                    rsDiasCorresp.Close
                    GoTo sinDatos
                End If
                
                Genera = True

                'EAM- Ecuación original
                'cantDiasProp = Round(totalDias / 364 * 2 * DiasHorario, 4)
                cantDiasProp = Round(DiasHorario * (14 / 364), 4)
            Case 30:
                'EAM- Se proporciona 1 día vacación cada 30 días trabajados.
                'EAM- Busca la fecha hasta de vigencia del período
                StrSql = "SELECT * FROM vacacion WHERE vacfecdesde <= " & ConvFecha(DateAdd("d", 1, ultFechaProcesada)) & " AND vacfechasta >= " & ConvFecha(DateAdd("d", 1, ultFechaProcesada)) & _
                        " AND ternro= " & Ternro & " ORDER BY vacfecdesde DESC"
                OpenRecordset StrSql, rsDiasCorresp
                                         
                totalDias = DateDiff("d", ultFechaProcesada, FechaHasta)
                If Not rsDiasCorresp.EOF Then
                    NroVac = rsDiasCorresp!vacnro
                    fechaCorte = rsDiasCorresp!vacfechasta
                Else
                    Flog.writeline "No existe el Período de Vacacion. Fecha: " & ConvFecha(DateAdd("d", 1, ultFechaProcesada))
                    Exit Sub
                End If
                
                Do While ultFechaProcesada < FechaHasta
                                        
                    Flog.writeline "Empleado que proporcionan 1 día por 30 trabajados."
                    
                    'EAM- Se agrego la validacion (DateAdd("d", 1, rsDiasCorresp!vacfechasta) < FechaHasta) para verificar que no sea el ultimo dia del período del empleado sino corta siempre esa fecha. 10/07/2012
                    If (ultFechaProcesada <= FechaHasta) And (fechaCorte <= FechaHasta) Then
                        Call DIF_FECHAS2(CDate(ultFechaProcesada), CDate(fechaCorte), cantDia, cantMes, cantAnio)
                        ultFechaProcesada = rsDiasCorresp!vacfechasta
                    Else
                        Call DIF_FECHAS2(CDate(ultFechaProcesada), CDate(FechaHasta), cantDia, cantMes, cantAnio)
                        ultFechaProcesada = FechaHasta
                    End If
                                                        
                    cantDiasProp = cantMes + (cantAnio * 12)
                                    
                    'EAM- Si es mayor que 0 actualizo la cantidad de días
                    If (cantDiasProp > 0) Or (ultFechaProcesada = fechaCorte) Then
                        'Obtiene el saldo del período al cual calculó los dias correspondientes
                         StrSql = " SELECT vdiascorcant FROM vacdiascor " & _
                                "  WHERE ternro= " & Ternro & " AND venc= 0 AND vacnro=" & NroVac & " ORDER BY vdiasfechasta DESC"
                        OpenRecordset StrSql, rsDiasCorresp
                        
                        If Not rsDiasCorresp.EOF Then
                            cantdias = rsDiasCorresp!vdiascorcant
                        Else
                            cantdias = 0
                        End If
                        cantdias = cantdias + cantDiasProp
                        Call actualizarDiasVacCR(NroVac, cantdias, TipoVacacionProporcion, DateAdd("d", -(cantDia), ultFechaProcesada))
                        Flog.writeline Espacios(Tabulador * 1) & "Actualizo días vacaciones. VacNro: " & NroVac & " Cant Días: " & cantdias
                        cantDiasProp = 0
                        cantdias = 0
                        Genera = True
                    End If
                    
                    If (fechaCorte <= FechaHasta) Then
                        'EAM- Imputa a Período Nuevo
                        NroVac = 0
                        Call BuscarPeriodoVacCR(DateAdd("d", 1, ultFechaProcesada), Ternro, NroVac, fechaCorte)
                        If NroVac = 0 Then
                            Flog.writeline Espacios(Tabulador * 1) & "No se encontro el periodo de Vacación para la fecha ." & ultFechaProcesada
                            Exit Sub
                        End If
                    End If
                Loop
                
                
        
            Case Else:
                'EAM- Empleados Mensuales, su asignacion horaria se rige por el tipo de vacacion
                'EAM- Busca la fecha hasta de vigencia del período
                StrSql = "SELECT * FROM vacacion WHERE vacfecdesde <= " & ConvFecha(ultFechaProcesada) & " AND vacfechasta > " & ConvFecha(ultFechaProcesada) & _
                        " AND ternro= " & Ternro & " ORDER BY vacfecdesde DESC"
                OpenRecordset StrSql, rsDiasCorresp
                
                If Not rsDiasCorresp.EOF Then
                    'EAM- Se agrego la validacion (DateAdd("d", 1, rsDiasCorresp!vacfechasta) < FechaHasta) para verificar que no sea el ultimo dia del período del empleado sino corta siempre esa fecha. 10/07/2012
                    If (rsDiasCorresp!vacfechasta < FechaHasta) And (DateAdd("d", 1, rsDiasCorresp!vacfechasta) < FechaHasta) Then
                        FechaHasta = rsDiasCorresp!vacfechasta
                        NroVac = rsDiasCorresp!vacnro
                    End If
                End If
                
                Flog.writeline "Empleado Mensuales. Busca los días laborables del tipo de vacacion: " & TipoVacacionProporcion
                
                StrSql = " SELECT * FROM tipovacac WHERE tipvacnro = " & TipoVacacionProporcion
                OpenRecordset StrSql, rsRegHorario
                
                'EAM- Obtiene la cantidad de dias laborables
                DiasHorario = 0
                Do While Not rsRegHorario.EOF
                    If rsRegHorario!tpvhabiles__1 Then DiasHorario = DiasHorario + 1
                    If rsRegHorario!tpvhabiles__2 Then DiasHorario = DiasHorario + 1
                    If rsRegHorario!tpvhabiles__3 Then DiasHorario = DiasHorario + 1
                    If rsRegHorario!tpvhabiles__4 Then DiasHorario = DiasHorario + 1
                    If rsRegHorario!tpvhabiles__5 Then DiasHorario = DiasHorario + 1
                    If rsRegHorario!tpvhabiles__6 Then DiasHorario = DiasHorario + 1
                    If rsRegHorario!tpvhabiles__7 Then DiasHorario = DiasHorario + 1
                    rsRegHorario.MoveNext
                Loop
                
                totalDias = DateDiff("d", ultFechaProcesada, FechaHasta)
                cantDiasProp = Round(totalDias / 364 * 2 * DiasHorario, 4)
                Genera = True
                'ultFechaProcesada = FechaHasta
                Call actualizarDiasVacCR(NroVac, cantdias, TipoVacacionProporcion, ultFechaProcesada)
        End Select
        
        'Obtiene el saldo del período al cual calculó los dias correspondientes
         StrSql = " SELECT vdiascorcant FROM vacdiascor " & _
                "  WHERE ternro= " & Ternro & " AND venc= 0 AND vacnro=" & NroVac & " ORDER BY vdiasfechasta DESC"
        OpenRecordset StrSql, rsDiasCorresp
        
        If Not rsDiasCorresp.EOF Then
            cantdias = rsDiasCorresp!vdiascorcant
        Else
            cantdias = 0
        End If
            
    
        
        
        
        
        
'        'ultFechaProcesada = FechaHasta
        If FactorDivision <> 30 Then
            Flog.writeline "Cantidad de días acumulados: " & cantdias
            Flog.writeline "Total de días a procesar: " & totalDias
            Flog.writeline "Días Laborables: " & DiasHorario
            
            Flog.writeline "Cantidad de días generados: " & cantDiasProp
            cantdias = cantdias + cantDiasProp
            
            Call actualizarDiasVacCR(NroVac, cantdias, TipoVacacionProporcion, ultFechaProcesada)
            Flog.writeline "Se actualizo los días correspondientes del vacnro : " & NroVac & " Cant. Dias: " & cantdias & " Fecha: " & ultFechaProcesada
            Genera = True
        End If
        
    End If

    GoTo finalizado
CE:
    Flog.writeline " ------------------------------------------------------------"
    Flog.writeline "Error procesando Empleado:" & Ternro & " Año " & Anio
    Flog.writeline Err.Description
    Flog.writeline "SQL: " & StrSql
    Flog.writeline " ------------------------------------------------------------"

sinDatos:
    Exit Sub
finalizado:
    Set rsDiasCorresp = Nothing
    Set rsRegHorario = Nothing
End Sub

Public Function FechaAltaEmpleado(Ternro) As Date
    Dim StrSql As String
    Dim rsFases As New ADODB.Recordset
    Dim i_dia As Integer
    Dim i_mes As Integer
    Dim i_anio As Integer
    
    'StrSql = "SELECT * FROM fases where fases.empleado = " & Ternro & " AND fases.fasrecofec = -1 "
    
    'Agregado el 18/10/2013
    StrSql = " SELECT altfec "
    StrSql = StrSql & " FROM fases "
    StrSql = StrSql & " LEFT JOIN empant ON fases.empantnro=empant.empantnro "
    StrSql = StrSql & " LEFT JOIN causa ON fases.caunro=causa.caunro "
    StrSql = StrSql & " WHERE fases.empleado=" & Ternro & ""
    StrSql = StrSql & " ORDER BY altfec DESC, estado ASC"
    'fin
    
    OpenRecordset StrSql, rsFases
    If rsFases.EOF Then
        FechaAltaEmpleado = Empty
        'Flog.writeline "El empleado no tiene fase con alta reconocida."
        Flog.writeline "El empleado no tiene fase activa."
    Else
        If IsNull(rsFases("altfec")) Then
            FechaAltaEmpleado = ""
        Else
            FechaAltaEmpleado = CDate(rsFases("altfec"))
            
            'EAM (V3.28) - Verifico que el empleado no tenga fecha de alta un 29 de febrero (bisiesto).
            If (Day(FechaAltaEmpleado) = 29) And (Month(FechaAltaEmpleado) = 2) Then
                FechaAltaEmpleado = CDate("28/02/" & Year(FechaAltaEmpleado))
            End If
        End If
    End If
    
    If rsFases.State = adStateOpen Then rsFases.Close
    Set rsFases = Nothing
    
End Function

Sub generarPeriodoVacacion(Ternro As Long, Anio As Integer, Optional modeloPais As Integer, Optional cambiaFase As Integer)
'14/09/2012 - Gonzalez Nicolás - Se setea vacestado en -1 por default
    Dim rs_vacacion As New ADODB.Recordset
    Dim fechaAlta As Date
    Dim fechaInicio As Date
    Dim fechaFin As Date
    Dim vacdesc As String
    Dim vacestado As Integer

    fechaAlta = FechaAltaEmpleado(Ternro)
    
    If modeloPais = 4 Then
        Flog.writeline "voy a comprobar si cambia de fase:" & Anio
        If cambiaFase = -1 Then
            Anio = Year(fechaAlta)
              Flog.writeline "Cambia fase.Año:" & Anio
        End If
         Flog.writeline "Lo reimprimo :" & Anio
    Else
        If Anio < Year(fechaAlta) Then
            Flog.writeline "Error al querer generar un periodo anterior a la fecha de alta del empleado."
            Exit Sub
        End If
    End If
    
    fechaInicio = formatFecha(CStr(Day(fechaAlta)), CStr(Month(fechaAlta)), CStr(Anio))
    fechaFin = DateAdd("d", -1, DateAdd("yyyy", 1, fechaInicio))
    vacdesc = CStr(Anio) & " - " & CStr(Anio + 1) & " - (" & fechaInicio & ") "
    
    vacestado = -1 ' NG
    
    If modeloPais = 3 Then
        If Anio < Year(Date) - 4 Then
            vacestado = 0
        Else
            vacestado = -1
        End If
    End If
    
    StrSql = " INSERT INTO vacacion (vacdesc, vacfecdesde, vacfechasta, vacanio, empnro, ternro, vacestado) "
    StrSql = StrSql & " VALUES ('" & vacdesc & "'," & ConvFecha(fechaInicio) & "," & ConvFecha(fechaFin) & "," & Anio & ",0," & Ternro & "," & vacestado & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.writeline "Inserto en vacacion.Sql :" & StrSql
    
    
End Sub

Function formatFecha(ByVal Dia As String, ByVal mes As String, ByVal Anio As String) As Date
    If Len(Day(Dia)) = 1 Then
        Dia = "0" & Dia
    End If
    If Len(Month(mes)) = 1 Then
        mes = "0" & mes
    End If
    If Len(Anio) = 2 Then
        Anio = "20" & Anio
    End If
    formatFecha = CDate(Dia & "/" & mes & "/" & Anio)
End Function



Public Sub bus_DiasVac_old(ByVal Ternro As Long, ByVal NroVac As Long, ByRef cantdias As Integer, ByRef Columna As Integer, ByRef Mensaje As String, ByRef Genera As Boolean)
' ---------------------------------------------------------------------------------------------
' Descripcion: Obtencion del valor de la escala para vacaciones.
' Autor      : FGZ
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Valor_Grilla(10) As Boolean ' Elemento de una coordenada de una grilla
Dim tipoBus As Long
Dim concnro As Long
Dim prog As Long

Dim tdinteger3 As Integer

Dim ValAnt As Single
Dim Busq As Integer
Dim dias_maternidad As Integer

Dim j As Integer
Dim antig As Boolean
Dim pvariable As Boolean
Dim pvar As Integer
Dim ant As Integer
Dim continuar As Boolean
Dim parametros(5) As Integer
Dim grilla_val(10) As Boolean     ' para alojar los valores de:  valgrilla.val(i)

Dim vgrcoor_ant As Long
Dim vgrcoor_pvar As Long

Dim rs_valgrilla As New ADODB.Recordset
Dim rs_cabgrilla As New ADODB.Recordset
Dim rs_tbase As New ADODB.Recordset
Dim TipoBase As Long

Dim NroBusqueda As Long

Dim antdia As Long
Dim antmes As Long
Dim antanio As Long
Dim q As Integer

Dim Aux_Dias_trab As Double
Dim aux_redondeo As Double
Dim ValorCoord As Single
Dim Encontro As Boolean
Dim VersionBaseAntig As Integer
Dim habiles As Integer
Dim ExcluyeFeriados As Boolean
Dim rs As New ADODB.Recordset

    Genera = False
    
    Call Politica(1502)
    If Not PoliticaOK Then
        Flog.writeline "Error cargando configuracion de la Politica 1502"
        Exit Sub
    End If
    
'    StrSql = "SELECT * FROM tipdia WHERE tdnro = 2 " '2 es vacaciones
'    OpenRecordset StrSql, objRs
'    If Not objRs.EOF Then
'        NroGrilla = objRs!tdgrilla
'        tdinteger3 = objRs!tdinteger3
'
'        If tdinteger3 <> 20 And tdinteger3 <> 365 And tdinteger3 <> 360 Then
'            'El campo auxiliar3 del Tipo de Día para Vacaciones no está configurado para Proporcionar la cant. de días de Vacaciones.
'            Exit Sub
'        End If
'    End If

    StrSql = "SELECT * FROM cabgrilla " & _
             " WHERE cabgrilla.cgrnro = " & NroGrilla
    OpenRecordset StrSql, rs_cabgrilla

    If rs_cabgrilla.EOF Then
        'La escala de Vacaciones no esta configurada en el tipo de dia para vacaciones
        Flog.writeline "La escala de Vacaciones no esta configurada o el nro de grilla no esta bien configurado bien en la Politica 1502. Grilla " & NroGrilla
        Exit Sub
    End If
    
    Call Politica(1505)
    If Not PoliticaOK Then
        Flog.writeline "Error cargando configuracion de la Politica 1505. Tipo Base antiguedad estandar."
        VersionBaseAntig = 0
    Else
        VersionBaseAntig = st_BaseAntiguedad
    End If
    
    
    'El tipo Base de la antiguedad
    TipoBase = 4
    
    continuar = True
    ant = 1
    Do While (ant <= rs_cabgrilla!cgrdimension) And continuar
        Select Case ant
        Case 1:
            StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_1
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    continuar = False
                Else
                    ant = ant + 1
                End If
            End If
            
        Case 2:
            StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_2
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    continuar = False
                Else
                    ant = ant + 1
                End If
            End If
        Case 3:
            StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_3
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    continuar = False
                Else
                    ant = ant + 1
                End If
            End If
        Case 4:
            StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_4
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    continuar = False
                Else
                    ant = ant + 1
                End If
            End If
        Case 5:
            StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_5
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    continuar = False
                Else
                    ant = ant + 1
                End If
            End If
        End Select
    Loop

    'setea la proporcion de dias
    Call Politica(1501)
        
    For j = 1 To rs_cabgrilla!cgrdimension
        Select Case j
        Case ant:
            'la busqueda es la de antiguedad
            Select Case VersionBaseAntig
            Case 0:
                Flog.writeline "Antiguedad estandar "
                'Call bus_Antiguedad("VACACIONES", CDate("31/12/" & Year(fecha_desde)), antdia, antmes, antanio, q)
                'Call bus_Antiguedad("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
                Flog.writeline "Antiguedad Standard " ' Se computa al año actual
                Call bus_Antiguedad_R("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
                If (((antmes * 30) + antdia >= st_Dias) Or antanio <> 0) Then
                       antdia = 0
                       antmes = 0
                       antanio = 0
                       Call bus_Antiguedad("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
                 End If
            
            Case 1:
                Flog.writeline "Antiguedad Customizada "
                'Call bus_Antiguedad_G("VACACIONES", CDate("31/12/" & Year(fecha_desde)), antdia, antmes, antanio, q)
                Call bus_Antiguedad_G("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
            Case 2:
                Flog.writeline "Antiguedad Uruguay " ' Se computa al año anterior
                'Call bus_Antiguedad_G("VACACIONES", CDate("31/12/" & Year(fecha_desde)), antdia, antmes, antanio, q)
                Call bus_Antiguedad("VACACIONES", CDate("31/12/" & Periodo_Anio - 1), antdia, antmes, antanio, q)
            Case 3:
                 Flog.writeline "Antiguedad Standard " ' Se computa al año actual
                 Call bus_Antiguedad_R("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
                 If (((antmes * 30) + antdia) >= st_Dias) Or antanio <> 0 Then
                       antdia = 0
                       antmes = 0
                       antanio = 0
                       Call bus_Antiguedad("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
                 End If
            Case Else
                Flog.writeline "Antiguedad Mal configurada. Estandar "
                'Call bus_Antiguedad("VACACIONES", CDate("31/12/" & Year(fecha_desde)), antdia, antmes, antanio, q)
                Call bus_Antiguedad("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
            End Select
            parametros(j) = (antanio * 12) + antmes
            'Call bus_Antiguedad("VACACIONES", CDate("31/12/" & Year(fecha_desde)), antdia, antmes, antanio, q)
            Flog.writeline "Años " & antanio
            Flog.writeline "Meses " & antmes
            Flog.writeline "Dias " & antdia
            
        Case Else:
            Select Case j
            Case 1:
                Call bus_Estructura(rs_cabgrilla!grparnro_1)
            Case 2:
                Call bus_Estructura(rs_cabgrilla!grparnro_2)
            Case 3:
                Call bus_Estructura(rs_cabgrilla!grparnro_3)
            Case 4:
                Call bus_Estructura(rs_cabgrilla!grparnro_4)
            Case 5:
                Call bus_Estructura(rs_cabgrilla!grparnro_5)
            End Select
            parametros(j) = valor
        End Select
    Next j

    'Busco la primera antiguedad de la escala menor a la del empleado
    ' de abajo hacia arriba
    StrSql = "SELECT * FROM valgrilla WHERE cgrnro = " & NroGrilla
    For j = 1 To rs_cabgrilla!cgrdimension
        If j <> ant Then
            StrSql = StrSql & " AND vgrcoor_" & j & "= " & parametros(j)
        End If
    Next j
        StrSql = StrSql & " ORDER BY vgrcoor_" & ant & " DESC "
    OpenRecordset StrSql, rs_valgrilla


    Encontro = False
    Do While Not rs_valgrilla.EOF And Not Encontro
        Select Case ant
        Case 1:
            If parametros(ant) >= rs_valgrilla!vgrcoor_1 Then
                 If rs_valgrilla!vgrvalor <> 0 Then
                    cantdias = rs_valgrilla!vgrvalor
                    Encontro = True
                    Columna = rs_valgrilla!vgrorden
                 End If
            End If
        Case 2:
            If parametros(ant) >= rs_valgrilla!vgrcoor_2 Then
                 If rs_valgrilla!vgrvalor <> 0 Then
                    cantdias = rs_valgrilla!vgrvalor
                    Encontro = True
                    Columna = rs_valgrilla!vgrorden
                 End If
            End If
        Case 3:
            If parametros(ant) >= rs_valgrilla!vgrcoor_3 Then
                 If rs_valgrilla!vgrvalor <> 0 Then
                    cantdias = rs_valgrilla!vgrvalor
                    Encontro = True
                    Columna = rs_valgrilla!vgrorden
                 End If
            End If
        Case 4:
            If parametros(ant) >= rs_valgrilla!vgrcoor_4 Then
                 If rs_valgrilla!vgrvalor <> 0 Then
                    cantdias = rs_valgrilla!vgrvalor
                    Encontro = True
                    Columna = rs_valgrilla!vgrorden
                 End If
            End If
        Case 5:
            If parametros(ant) >= rs_valgrilla!vgrcoor_5 Then
                 If rs_valgrilla!vgrvalor <> 0 Then
                    cantdias = rs_valgrilla!vgrvalor
                    Encontro = True
                    Columna = rs_valgrilla!vgrorden
                 End If
            End If
        End Select
                    
        rs_valgrilla.MoveNext
    Loop
    
    If Not Encontro Then
        'Busco si existe algun valor para la estructura y ...
        'si hay que carga la columna correspondiente
        StrSql = "SELECT * FROM valgrilla WHERE cgrnro = " & NroGrilla
        StrSql = StrSql & " AND vgrvalor is not null"
        For j = 1 To rs_cabgrilla!cgrdimension
            If j <> ant Then
                StrSql = StrSql & " AND vgrcoor_" & j & "= " & parametros(j)
            End If
        Next j
        'StrSql = StrSql & " ORDER BY vgrcoor_" & ant & " DESC "
        OpenRecordset StrSql, rs_valgrilla
        If Not rs_valgrilla.EOF Then
            Columna = rs_valgrilla!vgrorden
        Else
            Columna = 1
        End If
        
'        Aux_Dias_trab = antanio * 365
'        Aux_Dias_trab = Aux_Dias_trab + antmes * 30
'        Aux_Dias_trab = Aux_Dias_trab + antdia
'        dias_trabajados = CLng(Aux_Dias_trab)
        
        dias_trabajados = ((antanio * 365) + (antmes * 30) + antdia)
        Flog.writeline "Dias trabajados " & dias_trabajados
        
        Flog.writeline "ANT " & ant
        
        If parametros(ant) <= BaseAntiguedad Then
            
            'FGZ - 16/02/2006
            habiles = 0
            StrSql = "SELECT * FROM tipovacac WHERE tipvacnro = " & TipoVacacionProporcion
            OpenRecordset StrSql, rs
            If Not rs.EOF Then
                If rs!tpvhabiles__1 Then habiles = habiles + 1
                If rs!tpvhabiles__2 Then habiles = habiles + 1
                If rs!tpvhabiles__3 Then habiles = habiles + 1
                If rs!tpvhabiles__4 Then habiles = habiles + 1
                If rs!tpvhabiles__5 Then habiles = habiles + 1
                If rs!tpvhabiles__6 Then habiles = habiles + 1
                If rs!tpvhabiles__7 Then habiles = habiles + 1
                
                ExcluyeFeriados = CBool(rs!tpvferiado)
            Else
                'por default tomo 7
                habiles = 7
            End If
            'Para que la proporcion sea lo mas exacto posible tengo que
            'restar a los dias trabajados (que caen dentro de una fase) los dias feiados que son habiles
            'antes de proporcionar
            If ExcluyeFeriados Then
                'deberia revisar dia por dia de los dias contemplados para la antiguedad revisando si son feriados y dia habil
                
            End If
            
            
            Flog.writeline "Dias Proporcion " & DiasProporcion
            Flog.writeline "Factor de Division " & FactorDivision
            Flog.writeline "Tipo Base Antiguedad " & BaseAntiguedad
            Flog.writeline "Dias habiles " & habiles
            
            
'            If DiasProporcion = 20 Then
'                If (dias_trabajados / DiasProporcion) / 7 * 5 > Fix((dias_trabajados / DiasProporcion) / 7 * 5) Then
'                    cantdias = Fix((dias_trabajados / DiasProporcion) / 7 * 5) + 1<d
'                Else
'                    cantdias = Fix((dias_trabajados / DiasProporcion) / 7 * 5)
'                End If
'            Else
'                cantdias = Round((dias_trabajados / DiasProporcion) / FactorDivision, 0)
'            End If
            
'            Agregue el control del parámetro redondeo. Gustavo
             
             If dias_trabajados < 20 Then
                cantdias = 0
             Else
                If DiasProporcion = 20 Then
                        cantdias = Fix((dias_trabajados / DiasProporcion) / 7 * habiles)
                    Else
                        cantdias = Fix(20 * (dias_trabajados / DiasProporcion) / FactorDivision)
                End If
                
                aux_redondeo = ((dias_trabajados / DiasProporcion) / 7 * habiles) - Fix((dias_trabajados / DiasProporcion) / 7 * habiles)
                
                Select Case st_redondeo
                
                Case 0 ' Redondea hacia abajo - Ya se realizo el cálculo
                    
                Case 1 ' Redondea hacia arriba
                    If aux_redondeo <> 0 Then
                        cantdias = cantdias + 1
                    End If
                    
                Case Else ' redondea hacia abajo si la parte decimal <.5 sino hacia arriba
                    If aux_redondeo >= 0.5 Then
                        cantdias = cantdias + 1
                    End If
                
                End Select
            End If
            Flog.writeline "Días Correspondientes:" & cantdias
            Flog.writeline "Tipo de redondeo:" & st_redondeo
            Flog.writeline "Parte decimal de los días correspondientes:" & aux_redondeo
            Flog.writeline
            
            'FGZ - 25/06/2009 ------------- Vacaciones Acordadas ------------------------------
            PoliticaOK = False
            DiasAcordados = False
            Call Politica(1511)
            If PoliticaOK And DiasAcordados Then
                 StrSql = "SELECT tipvacnro, diasacord FROM vacdiasacord "
                 StrSql = StrSql & " WHERE ternro = " & Ternro
                 OpenRecordset StrSql, rs
                 If Not rs.EOF Then
                     If rs!diasacord > cantdias Then
                         Flog.writeline "La cantidad de dias correspondientes es menor a la cantidad de dias acordados. " & rs!diasacord
                         Flog.writeline "Se utilizará la cantidad de dias acordados"
                         cantdias = rs!diasacord
                     End If
                 End If
            End If
            'FGZ - 25/06/2009 ------------- Vacaciones Acordadas ------------------------------
            Flog.writeline
            
            ' NF - 03/07/06
            PoliticaOK = False
            Call Politica(1508)
            If PoliticaOK Then
                Flog.writeline "Politica 1508 activa. Analizando Licencias por Maternidad (" & Tipo_Dia_Maternidad & ")."
                dias_maternidad = 0
                'StrSql = "SELECT * FROM emplic "
                StrSql = "SELECT SUM(elcantdias) total FROM emp_lic "
                StrSql = StrSql & " WHERE tdnro = " & Tipo_Dia_Maternidad
                StrSql = StrSql & " AND empleado = " & Ternro
                StrSql = StrSql & " AND elfechadesde >= " & ConvFecha("01/01/" & (Periodo_Anio - 1))
                StrSql = StrSql & " AND elfechahasta <= " & ConvFecha("31/12/" & (Periodo_Anio - 1))
                OpenRecordset StrSql, rs
                If Not rs.EOF And (Not IsNull(rs!total)) Then
                    dias_maternidad = rs!total
                    If dias_maternidad <> 0 Then
                        Flog.writeline "  Dias por maternidad: " & dias_maternidad
                        Flog.writeline "  Dias = " & cantdias & " - (" & dias_maternidad & " x " & Factor & ")"
                        cantdias = cantdias - CInt(dias_maternidad * Factor)
                    End If
                Else
                    Flog.writeline "  No se encontraron dias por maternidad."
                End If
                rs.Close
            End If
        Else
            Flog.writeline "No se encontro la escala para el convenio"
            Genera = False
        End If
    Else
        'FGZ - 25/06/2009 ------------- Vacaciones Acordadas ------------------------------
        PoliticaOK = False
        DiasAcordados = False
        Call Politica(1511)
        If PoliticaOK And DiasAcordados Then
             StrSql = "SELECT tipvacnro, diasacord FROM vacdiasacord "
             StrSql = StrSql & " WHERE ternro = " & Ternro
             OpenRecordset StrSql, rs
             If Not rs.EOF Then
                 If rs!diasacord > cantdias Then
                     Flog.writeline "La cantidad de dias correspondientes es menor a la cantidad de dias acordados. " & rs!diasacord
                     Flog.writeline "Se utilizará la cantidad de dias acordados"
                     cantdias = rs!diasacord
                 End If
             End If
        End If
        'FGZ - 25/06/2009 ------------- Vacaciones Acordadas ------------------------------
        Flog.writeline
    End If
   
Genera = True
    
' Cierro todo y libero
If rs_cabgrilla.State = adStateOpen Then rs_cabgrilla.Close
If rs_valgrilla.State = adStateOpen Then rs_valgrilla.Close

Set rs_cabgrilla = Nothing
Set rs_valgrilla = Nothing
End Sub


Public Sub Actulizar_Cartera_Vac(ByVal Ternro As Long, ByVal NroVac As Long, ByVal dias As Integer, ByVal Reproceso As Boolean)
' ---------------------------------------------------------------------------------------------
' Descripcion: Actualizacionde la cartera de vacaciones. Customizacion TTI.
' Autor      : FGZ
' Fecha      : 29/04/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_Cartera As New ADODB.Recordset
Dim rs_His_Cartera As New ADODB.Recordset


    StrSql = " SELECT * FROM ee_cartera_vac "
    StrSql = StrSql & " WHERE ternro =" & Ternro
    If rs_Cartera.State = adStateOpen Then rs_Cartera.Close
    OpenRecordset StrSql, rs_Cartera
    If rs_Cartera.EOF Then
        StrSql = "INSERT INTO ee_cartera_vac (ternro,saldoant,saldoact,saldofut) "
        StrSql = StrSql & " VALUES ("
        StrSql = StrSql & Ternro
        StrSql = StrSql & ",0"
        StrSql = StrSql & "," & dias
        StrSql = StrSql & ",0"
        StrSql = StrSql & " )"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        'Inserto en el historico
        StrSql = "INSERT INTO ee_cartera_vac_his (ternro,vacnro,saldoant,saldoact,saldofut) "
        StrSql = StrSql & " VALUES ("
        StrSql = StrSql & Ternro
        StrSql = StrSql & "," & NroVac
        StrSql = StrSql & ",0"
        StrSql = StrSql & "," & dias
        StrSql = StrSql & ",0"
        StrSql = StrSql & " )"
        objConn.Execute StrSql, , adExecuteNoRecords
    Else
        If Reproceso Then
            StrSql = " SELECT * FROM ee_cartera_vac_his "
            StrSql = StrSql & " WHERE ternro =" & Ternro
            StrSql = StrSql & " AND vacnro =" & NroVac
            If rs_His_Cartera.State = adStateOpen Then rs_His_Cartera.Close
            OpenRecordset StrSql, rs_His_Cartera
            If Not rs_His_Cartera.EOF Then
                'Actualizo
                StrSql = "UPDATE ee_cartera_vac SET "
                StrSql = StrSql & " saldoant = " & rs_His_Cartera!saldoant
                StrSql = StrSql & " ,saldoact = " & rs_His_Cartera!saldoact
                StrSql = StrSql & " ,saldofut = " & rs_His_Cartera!saldofut
                StrSql = StrSql & " WHERE ternro =" & Ternro
                objConn.Execute StrSql, , adExecuteNoRecords
                
                StrSql = " SELECT * FROM ee_cartera_vac "
                StrSql = StrSql & " WHERE ternro =" & Ternro
                If rs_Cartera.State = adStateOpen Then rs_Cartera.Close
                OpenRecordset StrSql, rs_Cartera
                
            End If
        End If
        
        'Actualizo
        StrSql = "UPDATE ee_cartera_vac SET "
        StrSql = StrSql & " saldoant = saldoant + " & rs_Cartera!saldoact
        StrSql = StrSql & " ,saldoact = " & dias - rs_Cartera!saldofut
        StrSql = StrSql & " ,saldofut = 0 "
        StrSql = StrSql & " WHERE ternro =" & Ternro
        objConn.Execute StrSql, , adExecuteNoRecords
    End If

' Cierro todo y libero
If rs_Cartera.State = adStateOpen Then rs_Cartera.Close
If rs_His_Cartera.State = adStateOpen Then rs_His_Cartera.Close

Set rs_Cartera = Nothing
Set rs_His_Cartera = Nothing
End Sub


'EAM- Obtiene los dias habiles trabajado en el ultimo año
Public Function DiasHabilTrabajado(ByVal Ternro As Long, ByVal fDesde As Date, ByVal fHasta As Date) As Double
  Dim rsRegHorario As New ADODB.Recordset
  Dim cantdias As Long
  Dim totalDias As Double
  Dim fecDesde As Date
  Dim fecHasta As Date
  
    'Seta las variables
    cantdias = 0
    totalDias = 0
  
    'EAM- Obtiene los Régimen Horario del periodo
    StrSql = "SELECT estrcodext, htetdesde, htethasta FROM tipoestructura " & _
             "INNER JOIN estructura ON tipoestructura.tenro= estructura.tenro " & _
             "INNER JOIN his_estructura on estructura.estrnro= his_estructura.estrnro " & _
             "WHERE tipoestructura.tenro = 21 AND his_estructura.htetdesde<= " & ConvFecha(fHasta) & _
             " AND (his_estructura.htethasta>= " & ConvFecha(fDesde) & " OR his_estructura.htethasta IS NULL) " & _
             " AND his_estructura.ternro= " & Ternro & " ORDER BY his_estructura.htetdesde ASC "
    OpenRecordset StrSql, rsRegHorario
    
    'EAM- Recorre todos los regimen horarios del empleado
    Do While Not rsRegHorario.EOF
        If Not (IsNull(rsRegHorario!estrcodext)) And (rsRegHorario!estrcodext <> "") Then
            'EAM- Obtiene la fecha de inicio del Regimen Horario
            If CDate(rsRegHorario!htetdesde) <= CDate(fDesde) Then
                fecDesde = fDesde
            Else
                fecDesde = rsRegHorario!htetdesde
            End If
            
            'EAM- Obtiene la fecha de fin del Regimen Horario
            If IsNull(rsRegHorario!htethasta) Then
                fecHasta = fHasta
            Else
                If CDate(rsRegHorario!htethasta) >= CDate(fHasta) Then
                    fecHasta = fHasta
                Else
                    fecHasta = rsRegHorario!htethasta
                End If
            End If
            
            'cantDias = CInt(DateDiff("d", rsRegHorario!htetdesde, fhasta))
            cantdias = CLng(DateDiff("d", fecDesde, fecHasta) + 1)
            cantdias = cantdias - LicenciaGozadas(Ternro, fDesde, fHasta)
            
            'EAM- Le descuenta la cantidad de dias feriados que se encuentran en el rango de fecha
            If CLng(st_Modifica) = -1 Then
                cantdias = cantdias - cantDiasFeriados(fecDesde, fecHasta)
            End If
            totalDias = totalDias + ((cantdias / 7) * CInt(rsRegHorario!estrcodext))
            totalDias = Round(totalDias, 2)
        Else
            Flog.writeline "No se encuentra configurada la cantidad de Hs diarias trabajadas"
        End If
        rsRegHorario.MoveNext
    Loop
    totalDias = RedondearNumero(Int(totalDias), (totalDias - Int(totalDias)))
     DiasHabilTrabajado = totalDias
End Function

Public Function LicenciaGozadas(ByVal Ternro As Long, ByVal fDesde As Date, ByVal fHasta As Date) As Integer
  Dim rsLicencias As New ADODB.Recordset
  Dim fecDesde As Date
  Dim fecHasta As Date
  Dim cantLicencias As Double
  If Lic_Descuento = "" Or Lic_Descuento = " " Then
    Lic_Descuento = "0"
  End If
    cantLicencias = 0
  
    StrSql = "SELECT * FROM emp_lic " & _
             "WHERE empleado= " & Ternro & " AND emp_lic.elfechadesde<= " & ConvFecha(fHasta) & _
             " AND (emp_lic.elfechahasta>= " & ConvFecha(fDesde) & " OR emp_lic.elfechahasta IS NULL) " & _
             " AND emp_lic.tdnro in (" & Lic_Descuento & ")"
    OpenRecordset StrSql, rsLicencias
    
    
    'EAM- Recorre todos las Licencias
    Do While Not rsLicencias.EOF
        
        'EAM- Obtiene la fecha de inicio de la Licencia
        If CDate(rsLicencias!elfechadesde) <= CDate(fDesde) Then
            fecDesde = fDesde
        Else
            fecDesde = rsLicencias!elfechadesde
        End If
        
        'EAM- Obtiene la fecha de fin del Regimen Horario
        If IsNull(rsLicencias!elfechahasta) Then
            fecHasta = fHasta
        Else
            If CDate(rsLicencias!elfechahasta) >= CDate(fHasta) Then
                fecHasta = fHasta
            Else
                fecHasta = rsLicencias!elfechahasta
            End If
        End If
        
        cantLicencias = cantLicencias + CInt(DateDiff("d", fecDesde, fecHasta) + 1)
        rsLicencias.MoveNext
    Loop
    
     LicenciaGozadas = cantLicencias
End Function

'EAM- Obtiene el regimen horario actual del empleado
Public Function BuscarRegHorarioActual(ByVal Ternro As Long) As Integer
  Dim rsRegHorario As New ADODB.Recordset
  
    'EAM- Obtiene los Régimen Horario del Actual
    StrSql = "SELECT estrcodext FROM estructura " & _
             "INNER JOIN his_estructura on estructura.estrnro= his_estructura.estrnro " & _
             "WHERE his_estructura.tenro = 21 AND his_estructura.htetdesde<= " & ConvFecha(Date) & _
             " AND (his_estructura.htethasta>= " & ConvFecha(Date) & " OR his_estructura.htethasta IS NULL) " & _
             " AND his_estructura.ternro= " & Ternro & " ORDER BY his_estructura.htetdesde ASC "
    OpenRecordset StrSql, rsRegHorario
    
    If Not rsRegHorario.EOF Then
        If rsRegHorario!estrcodext <> "" Then
            BuscarRegHorarioActual = rsRegHorario!estrcodext
        End If
    Else
        BuscarRegHorarioActual = 7
        Flog.writeline "No se encontro regimen Horario para el empleado. Los 7 dias de la semana son Hábiles para el Empleado" & Ternro
    End If
    
End Function

'EAM- Calcula la proporcion de dias de vacaciones dado la cantidad de dias trabajado.Se calcula según la proporcion de dias.
Public Function CalcularProporcionDiasVac(ByVal dias_trabajados As Long)
 Dim cantdias As Long
 Dim aux_redondeo As Double
 
    If dias_trabajados < 20 Then
        cantdias = 0
    Else
        If DiasProporcion = 20 Then
            cantdias = Fix((dias_trabajados / DiasProporcion))
        Else
            cantdias = Fix(20 * (dias_trabajados / DiasProporcion) / FactorDivision)
        End If

        aux_redondeo = ((dias_trabajados / DiasProporcion)) - Fix((dias_trabajados / DiasProporcion))
        cantdias = RedondearNumero(cantdias, aux_redondeo)
                
    End If
    
    CalcularProporcionDiasVac = cantdias
End Function

'EAM- Redondea un numero decimal a un integer segun la configuracion de redondeo
Public Function RedondearNumero(ByVal NumEntero As Long, ByVal NumDecimal As Double) As Long
 Dim Numero As Long
 
    Select Case st_redondeo
        Case 0 ' Redondea hacia abajo - Ya se realizo el cálculo

        Case 1 ' Redondea hacia arriba
            If NumDecimal <> 0 Then
                Numero = NumEntero + 1
            End If
        Case 2 ' Redondea hacia abajo - Dimatz Rafael 24-10-13
            Numero = NumEntero
            
        Case Else ' redondea hacia abajo si la parte decimal <.5 sino hacia arriba
            If NumDecimal >= 0.5 Then
                Numero = NumEntero + 1
            Else
                Numero = NumEntero
            End If
    End Select
    
    RedondearNumero = Numero
End Function
Public Sub bus_DiasVac_PT(ByVal Ternro As Long, ByVal NroVac As Long, ByRef cantdias As Integer, ByRef Columna As Integer, ByRef Mensaje As String, ByRef Genera As Boolean _
    , ByRef cantdiasCorr As Integer, ByRef columna2 As Integer)
' ---------------------------------------------------------------------------------------------
' Descripcion: Obtencion del valor de la escala para vacaciones. PORTUGAL (Se utilizo de base el modelo de Argentina)
' Autor      : Gonzalez Nicolás
' Fecha      : 07/05/2012
' Ultima Mod.:
' ---------------------------------------------------------------------------------------------
Dim Valor_Grilla(10) As Boolean ' Elemento de una coordenada de una grilla
Dim tipoBus As Long
Dim concnro As Long
Dim prog As Long

Dim tdinteger3 As Integer

Dim ValAnt As Single
Dim Busq As Integer
Dim dias_maternidad As Integer

Dim j As Integer
Dim antig As Boolean
Dim pvariable As Boolean
Dim pvar As Integer
Dim ant As Integer
Dim continuar As Boolean
Dim parametros(5) As Integer
Dim grilla_val(10) As Boolean     ' para alojar los valores de:  valgrilla.val(i)

Dim vgrcoor_ant As Long
Dim vgrcoor_pvar As Long

Dim rs_valgrilla As New ADODB.Recordset
Dim rs_cabgrilla As New ADODB.Recordset
Dim rs_tbase As New ADODB.Recordset
Dim TipoBase As Long

Dim NroBusqueda As Long

Dim antdia As Long
Dim antmes As Long
Dim antanio As Long
Dim q As Integer

Dim Aux_Dias_trab As Double
Dim aux_redondeo As Double
Dim ValorCoord As Single
Dim Encontro As Boolean
Dim VersionBaseAntig As Integer
Dim habiles, habilesCorr As Integer
Dim ExcluyeFeriados As Boolean
Dim ExcluyeFeriadosCorr  As Boolean
Dim rs As New ADODB.Recordset
'EAM- 08-07-2010
Dim dias_efect_trabajado As Long
Dim regHorarioActual As Integer
'Dim arrEscala()
'ReDim Preserve arrEscala(5, 0)  'la escala la carga al (total de registros y )

    Genera = False
    Encontro = False
    
    Call Politica(1502)
    If Not PoliticaOK Then
        Flog.writeline "Error cargando configuracion de la Politica 1502"
        Exit Sub
    End If
    

    StrSql = "SELECT * FROM cabgrilla " & _
             " WHERE cabgrilla.cgrnro = " & NroGrilla
    OpenRecordset StrSql, rs_cabgrilla

    If rs_cabgrilla.EOF Then
        'La escala de Vacaciones no esta configurada en el tipo de dia para vacaciones
        Flog.writeline "La escala de Vacaciones no esta configurada o el nro de grilla no esta bien configurado bien en la Politica 1502. Grilla " & NroGrilla
        Exit Sub
    End If
    
    Call Politica(1505)
    If Not PoliticaOK Then
        Flog.writeline "Error cargando configuracion de la Politica 1505. Tipo Base antiguedad estandar."
        VersionBaseAntig = 0
    Else
        VersionBaseAntig = st_BaseAntiguedad
    End If
    
    
    'El tipo Base de la antiguedad
    TipoBase = 4
    
    continuar = True
    ant = 1
    Do While (ant <= rs_cabgrilla!cgrdimension) And continuar
        Select Case ant
        Case 1:
            StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_1
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    continuar = False
                Else
                    ant = ant + 1
                End If
            End If
            
        Case 2:
            StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_2
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    continuar = False
                Else
                    ant = ant + 1
                End If
            End If
        Case 3:
            StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_3
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    continuar = False
                Else
                    ant = ant + 1
                End If
            End If
        Case 4:
            StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_4
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    continuar = False
                Else
                    ant = ant + 1
                End If
            End If
        Case 5:
            StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_5
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    continuar = False
                Else
                    ant = ant + 1
                End If
            End If
        End Select
    Loop

    'setea la proporcion de dias
    Call Politica(1501)
        
    For j = 1 To rs_cabgrilla!cgrdimension
        Select Case j
        Case ant:
            'la busqueda es la de antiguedad
            Select Case VersionBaseAntig
            Case 0:
                Flog.writeline "Antiguedad estandar "
                Flog.writeline "Antiguedad En el ultimo año " ' Se computa al año actual
                Call bus_Antiguedad_R("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
                If (((antmes * 30) + antdia >= st_Dias) Or antanio <> 0) Then
                       antdia = 0
                       antmes = 0
                       antanio = 0
                       Call bus_Antiguedad("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
                 End If
            
            Case 1:
                Flog.writeline "Antiguedad Sin redondeo "
                Call bus_Antiguedad_R("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
                 If (((antmes * 30) + antdia) >= st_Dias) Or antanio <> 0 Then
                       antdia = 0
                       antmes = 0
                       antanio = 0
                      Call bus_Antiguedad_G("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
                 End If
            Case 2:
                Flog.writeline "Antiguedad Uruguay " ' Se computa al año anterior
                'Call bus_Antiguedad_G("VACACIONES", CDate("31/12/" & Year(fecha_desde)), antdia, antmes, antanio, q)
                Call bus_Antiguedad("VACACIONES", CDate("31/12/" & Periodo_Anio - 1), antdia, antmes, antanio, q)
            Case 3:
                 Flog.writeline "Antiguedad Standard " ' Se computa al año actual
                 Call bus_Antiguedad_R("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
                 If (((antmes * 30) + antdia) >= st_Dias) Or antanio <> 0 Then
                       antdia = 0
                       antmes = 0
                       antanio = 0
                       Call bus_Antiguedad("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
                 End If
            Case 4: ' Anguedad a una fecha dada por dia y mes del año
                Flog.writeline "Antiguedad a una fecha dada año siguiente"
                If Not (st_Dia = 0 Or st_Mes = 0) Then
                     Call bus_Antiguedad_R("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
                     If (((antmes * 30) + antdia >= st_Dias) Or antanio <> 0) Then
                       antdia = 0
                       antmes = 0
                       antanio = 0
                       Call bus_Antiguedad_G("VACACIONES", CDate(st_Dia & "/" & st_Mes & "/" & Periodo_Anio + 1), antdia, antmes, antanio, q)
                    End If
                 End If
            Case 5: ' Anguedad a una fecha dada por dia y mes del año
                Flog.writeline "Antiguedad a una fecha dada año actual"
                Call bus_Antiguedad_R("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
                If (((antmes * 30) + antdia >= st_Dias) Or antanio <> 0) Then
                       antdia = 0
                       antmes = 0
                       antanio = 0
                 Call bus_Antiguedad("VACACIONES", CDate(st_Dia & "/" & st_Mes & "/" & Periodo_Anio), antdia, antmes, antanio, q)
                 End If
            
            Case Else
                Flog.writeline "Antiguedad Mal configurada. Estandar "
                'Call bus_Antiguedad("VACACIONES", CDate("31/12/" & Year(fecha_desde)), antdia, antmes, antanio, q)
                Call bus_Antiguedad("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
            End Select
            parametros(j) = (antanio * 12) + antmes
            'Call bus_Antiguedad("VACACIONES", CDate("31/12/" & Year(fecha_desde)), antdia, antmes, antanio, q)
            Flog.writeline "Años " & antanio
            Flog.writeline "Meses " & antmes
            Flog.writeline "Dias " & antdia
            
        Case Else:
            Select Case j
            Case 1:
                Call bus_Estructura(rs_cabgrilla!grparnro_1)
            Case 2:
                Call bus_Estructura(rs_cabgrilla!grparnro_2)
            Case 3:
                Call bus_Estructura(rs_cabgrilla!grparnro_3)
            Case 4:
                Call bus_Estructura(rs_cabgrilla!grparnro_4)
            Case 5:
                Call bus_Estructura(rs_cabgrilla!grparnro_5)
            End Select
            parametros(j) = valor
        End Select
    Next j

'--------------------------------------------------------------------------------------------------
    cantdias = buscarDiasVacEscala(ant, rs_cabgrilla!cgrdimension, parametros, TipoVacacionProporcion, Encontro)
    Columna = TipoVacacionProporcion
    
    'CAPAZ NO USO
    cantdiasCorr = buscarDiasVacEscala(ant, rs_cabgrilla!cgrdimension, parametros, TipoVacacionProporcionCorr)
    columna2 = TipoVacacionProporcionCorr




    '------------------------------
    'llamada politica 1513
    '------------------------------
    'Tiene en cuenta los dias trabajados en el ultimo año
    Call Politica(1513)
    
    If Dias_efect_trab_anio Then
        Flog.writeline "Tiene en cuenta el ultimo año. Politica 1513."
            
        'Obtiene la proporcion de dias_trabajados  -->  (dias trabajados / 7) * regimen horio
        dias_efect_trabajado = DiasHabilTrabajado(Ternro, CDate("01/01/" & Periodo_Anio), CDate("31/12/" & Periodo_Anio))
        regHorarioActual = BuscarRegHorarioActual(Ternro)
        Aux_Dias_trab = ((180 / 7) * regHorarioActual)
        Aux_Dias_trab = RedondearNumero(Int(Aux_Dias_trab), (Aux_Dias_trab - Int(Aux_Dias_trab)))
        
        If dias_efect_trabajado <= Aux_Dias_trab Then
            Encontro = True
            cantdias = CalcularProporcionDiasVac(dias_efect_trabajado)
            
            Flog.writeline "Empleado " & Ternro & " con dias trabajado menor a mitad de año: " & dias_efect_trabajado
            Flog.writeline "Días Correspondientes: " & cantdias
            Flog.writeline "Tipo de redondeo: " & st_redondeo
            Flog.writeline "Parte decimal de los días correspondientes: " & aux_redondeo
            Flog.writeline
        End If
        
    End If
            
            
                                
    If Not Encontro Then
                
        'EAM- Si la columna1 es = a vacion no tiene el tipo de vacacion sino ya se configuro por la politica 1501 (columna,columna2) 10/02/2011
        If (Columna = 0) Then
            'Busco si existe algun valor para la estructura y ...
            'si hay que carga la columna correspondiente
            StrSql = "SELECT * FROM valgrilla WHERE cgrnro = " & NroGrilla
            StrSql = StrSql & " AND vgrvalor is not null"
            For j = 1 To rs_cabgrilla!cgrdimension
                If j <> ant Then
                    StrSql = StrSql & " AND vgrcoor_" & j & "= " & parametros(j)
                End If
            Next j
            OpenRecordset StrSql, rs_valgrilla
            If Not rs_valgrilla.EOF Then
                Columna = rs_valgrilla!vgrorden
            Else
                Columna = 1
            End If
        End If
        
                
        dias_trabajados = ((antanio * 365) + (antmes * 30) + antdia)
        Flog.writeline "Dias trabajados " & dias_trabajados
        
        Flog.writeline "ANT " & ant
        
        Flog.writeline "------------"
        Flog.writeline "parametros(ant) = " & parametros(ant)
        Flog.writeline "BaseAntiguedad = " & BaseAntiguedad
        Flog.writeline "------------"
        
        'If parametros(ant) <= BaseAntiguedad Then
        If antanio = 0 Then
            
            habiles = cantDiasLaborable(TipoVacacionProporcion, ExcluyeFeriados)
            habilesCorr = cantDiasLaborable(TipoVacacionProporcionCorr, ExcluyeFeriadosCorr)
            
            If ExcluyeFeriados Then
                'deberia revisar dia por dia de los dias contemplados para la antiguedad revisando si son feriados y dia habil
                
            End If
            
            Flog.writeline "Empleado " & Ternro & " con menos de 1 año de trabajo."
            Flog.writeline "Dias Proporcion " & DiasProporcion
            Flog.writeline "Factor de Division " & FactorDivision
            Flog.writeline "Tipo Base Antiguedad " & BaseAntiguedad
            Flog.writeline "Dias habiles " & habiles
            Flog.writeline "Dias habiles Corrido" & habilesCorr
            
            Flog.writeline "-------------"
            Flog.writeline "Años " & antanio
            Flog.writeline "Meses " & antmes
            Flog.writeline "Dias " & antdia
            Flog.writeline "-------------"
            
            If antmes <= 10 Then
                cantdias = antmes * 2
            ElseIf antmes > 10 Then
                cantdias = 20
            Else
                cantdias = 0
            End If
          

            
'             If dias_trabajados < 30 Then
'                cantdias = 0
'             Else
'                If DiasProporcion = 30 Then
'                        cantdias = Fix((dias_trabajados / DiasProporcion) / 7 * habiles)
'                        Flog.Writeline dias_trabajados & "/" & DiasProporcion & "/ 7 * " & habiles
'                    Else
'                        cantdias = Fix(20 * (dias_trabajados / DiasProporcion) / FactorDivision)
'                End If
'
'                aux_redondeo = ((dias_trabajados / DiasProporcion) / 7 * habiles) - Fix((dias_trabajados / DiasProporcion) / 7 * habiles)
'                cantdias = RedondearNumero(cantdias, aux_redondeo)
'
'                'EAM(13972)- Obtiene los dias corridos de vacaciones a partir de los dias correspondientes
'                cantdiasCorr = (cantdias * habilesCorr) / habiles
'                aux_redondeo = ((cantdias * habilesCorr) / habiles) - Fix(((cantdias * habilesCorr) / habiles))
'                cantdiasCorr = RedondearNumero(cantdiasCorr, aux_redondeo)
'
'            End If
'            Flog.writeline "Días Correspondientes:" & cantdias
'            Flog.writeline "Días Correspondientes Corridos:" & cantdiasCorr
'            Flog.writeline "Tipo de redondeo:" & st_redondeo
'            Flog.writeline "Parte decimal de los días correspondientes:" & aux_redondeo
'            Flog.writeline
            
            'FGZ - 25/06/2009 ------------- Vacaciones Acordadas ------------------------------
            PoliticaOK = False
            DiasAcordados = False
            Call Politica(1511)
            If PoliticaOK And DiasAcordados Then
                 StrSql = "SELECT tipvacnro, diasacord FROM vacdiasacord "
                 StrSql = StrSql & " WHERE ternro = " & Ternro
                 OpenRecordset StrSql, rs
                 If Not rs.EOF Then
                     If rs!diasacord > cantdias Then
                         Flog.writeline "La cantidad de dias correspondientes es menor a la cantidad de dias acordados. " & rs!diasacord
                         Flog.writeline "Se utilizará la cantidad de dias acordados"
                         cantdias = rs!diasacord
                     End If
                 End If
            End If
            'FGZ - 25/06/2009 ------------- Vacaciones Acordadas ------------------------------
            Flog.writeline
            
            ' NF - 03/07/06
            PoliticaOK = False
            Call Politica(1508)
            If PoliticaOK Then
                Flog.writeline "Politica 1508 activa. Analizando Licencias por Maternidad (" & Tipo_Dia_Maternidad & ")."
                dias_maternidad = 0
                'StrSql = "SELECT * FROM emplic "
                StrSql = "SELECT SUM(elcantdias) total FROM emp_lic "
                StrSql = StrSql & " WHERE tdnro = " & Tipo_Dia_Maternidad
                StrSql = StrSql & " AND empleado = " & Ternro
                StrSql = StrSql & " AND elfechadesde >= " & ConvFecha("01/01/" & (Periodo_Anio - 1))
                StrSql = StrSql & " AND elfechahasta <= " & ConvFecha("31/12/" & (Periodo_Anio - 1))
                OpenRecordset StrSql, rs
                If Not rs.EOF And (Not IsNull(rs!total)) Then
                    dias_maternidad = rs!total
                    If dias_maternidad <> 0 Then
                        Flog.writeline "  Dias por maternidad: " & dias_maternidad
                        Flog.writeline "  Dias = " & cantdias & " - (" & dias_maternidad & " x " & Factor & ")"
                        cantdias = cantdias - CInt(dias_maternidad * Factor)
                    End If
                Else
                    Flog.writeline "  No se encontraron dias por maternidad."
                End If
                rs.Close
            End If
        Else
            Flog.writeline "No se encontro la escala para el convenio"
            Genera = False
        End If
    Else
        'FGZ - 25/06/2009 ------------- Vacaciones Acordadas ------------------------------
        PoliticaOK = False
        DiasAcordados = False
        Call Politica(1511)
        If PoliticaOK And DiasAcordados Then
             StrSql = "SELECT tipvacnro, diasacord FROM vacdiasacord "
             StrSql = StrSql & " WHERE ternro = " & Ternro
             OpenRecordset StrSql, rs
             If Not rs.EOF Then
                 If rs!diasacord > cantdias Then
                     Flog.writeline "La cantidad de dias correspondientes es menor a la cantidad de dias acordados. " & rs!diasacord
                     Flog.writeline "Se utilizará la cantidad de dias acordados"
                     cantdias = rs!diasacord
                 End If
             End If
        End If
        'FGZ - 25/06/2009 ------------- Vacaciones Acordadas ------------------------------
        Flog.writeline
    End If
   
    
    'EAM(18891)- Descunto a los días de vacaciones (acordados o por escala) la proporción de días por licencias
    Call Politica(1516)
    
    If PoliticaOK And (cantdias > 0) Then
        Select Case st_Opcion
            Case 1:
                Aux_Dias_trab = LicenciaGozadas(Ternro, CDate("01/12/" & (Periodo_Anio - 1)), CDate("30/11/" & (Periodo_Anio)))
                Aux_Dias_trab = (Aux_Dias_trab / DiasProporcion)
                Flog.writeline "Cantidad de días de descuento por Licencias: " & Aux_Dias_trab
                
                Aux_Dias_trab = RedondearNumero(Int(Aux_Dias_trab), (Aux_Dias_trab - Int(Aux_Dias_trab)))
                Flog.writeline "Cantidad de días de descuento por Licencias con redondeo: " & Aux_Dias_trab
                
                cantdias = (cantdias - (st_Dias * Aux_Dias_trab))
                
                If (cantdias < 0) Then
                    cantidas = 0
                End If
                Flog.writeline "Cantidad de días correspondientes: " & cantdias
            Case Else:
                Flog.writeline "No se aplica el descuento de licencia. Versión incorrecta"
        End Select
    End If
Genera = True

Flog.writeline ""
Flog.writeline "CANTIDAD DE DIAS: " & cantdias
Flog.writeline ""
' Cierro todo y libero
If rs_cabgrilla.State = adStateOpen Then rs_cabgrilla.Close
If rs_valgrilla.State = adStateOpen Then rs_valgrilla.Close

Set rs_cabgrilla = Nothing
Set rs_valgrilla = Nothing
End Sub

'MR - Calcula cantidad de dias trabajados en los casos que el empleado tiene menos de 6 meses de antiguedad
Function cantDiasProporcion(ByVal fHasta As Date, ByVal habiles As Integer) As Long
Dim fDesde As Date
Dim dias As Long
Dim semanas As Long
Dim rsFases As New ADODB.Recordset

StrSql = " SELECT altfec "
StrSql = StrSql & " FROM fases "
StrSql = StrSql & " WHERE fases.Empleado = " & Ternro
StrSql = StrSql & " AND estado =-1 "
StrSql = StrSql & " ORDER by altfec DESC "
OpenRecordset StrSql, rsFases

If rsFases.EOF Then
    dias = 0
Else
    fDesde = rsFases!altfec
    dias = DateDiff("d", fDesde, fHasta) + 1
    semanas = DateDiff("w", fDesde, fHasta)
    dias = dias - semanas * (7 - habiles)
End If
    cantDiasProporcion = dias
End Function

'EAM- Obtiene la cantidad de días feriados cargados en el sistema para un rango de fecha
Function cantDiasFeriados(ByVal fDesde As Date, ByVal fHasta As Date) As Long
 Dim rsFeriados As New ADODB.Recordset
 Dim objFeriado As New Feriado
 Dim cantFeriado
 
 
    cantFeriado = 0
 
    'Busco todos los Feriados
    StrSql = "SELECT * FROM feriado WHERE ferifecha >= " & ConvFecha(fDesde) & " AND ferifecha < " & ConvFecha(fHasta)
    OpenRecordset StrSql, rsFeriados
    
    Do While Not rsFeriados.EOF
        If objFeriado.Feriado(rsFeriados!ferifecha, Ternro, False) Then
            cantFeriado = cantFeriado + 1
        End If

        rsFeriados.MoveNext
    Loop
    cantDiasFeriados = cantFeriado
    Set objFeriado = Nothing
End Function


Public Sub bus_DiasVac_Col(ByVal Ternro As Long, ByVal NroVac As Long, ByRef cantdias As Integer, ByRef Columna As Integer, ByRef Mensaje As String, ByRef Genera As Boolean)
' ---------------------------------------------------------------------------------------------
' Descripcion: Obtencion del valor de la escala para vacaciones Colombia.
' Autor      : Lisandro Moro
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Valor_Grilla(10) As Boolean ' Elemento de una coordenada de una grilla
Dim tipoBus As Long
Dim concnro As Long
Dim prog As Long

Dim tdinteger3 As Integer

Dim ValAnt As Single
Dim Busq As Integer
Dim dias_maternidad As Integer

Dim j As Integer
Dim antig As Boolean
Dim pvariable As Boolean
Dim pvar As Integer
Dim ant As Integer
Dim continuar As Boolean
Dim parametros(5) As Integer
Dim grilla_val(10) As Boolean     ' para alojar los valores de:  valgrilla.val(i)

Dim vgrcoor_ant As Long
Dim vgrcoor_pvar As Long

Dim rs_valgrilla As New ADODB.Recordset
Dim rs_cabgrilla As New ADODB.Recordset
Dim rs_vacacion As New ADODB.Recordset
Dim rs_tbase As New ADODB.Recordset
Dim TipoBase As Long

Dim NroBusqueda As Long

Dim antdia As Long
Dim antmes As Long
Dim antanio As Long
Dim q As Integer

Dim Aux_Dias_trab As Double
Dim aux_redondeo As Double
Dim ValorCoord As Single
Dim Encontro As Boolean
Dim VersionBaseAntig As Integer
Dim habiles As Integer
Dim ExcluyeFeriados As Boolean
Dim rs As New ADODB.Recordset
'EAM- 08-07-2010
Dim dias_efect_trabajado As Long
Dim regHorarioActual As Integer
Dim aux_antmes As Long


    Genera = False
    
    Call Politica(1502)
    If Not PoliticaOK Then
        Flog.writeline "Error cargando configuracion de la Politica 1502"
        Exit Sub
    End If
    

    StrSql = "SELECT * FROM cabgrilla " & _
             " WHERE cabgrilla.cgrnro = " & NroGrilla
    OpenRecordset StrSql, rs_cabgrilla

    If rs_cabgrilla.EOF Then
        'La escala de Vacaciones no esta configurada en el tipo de dia para vacaciones
        Flog.writeline "La escala de Vacaciones no esta configurada o el nro de grilla no esta bien configurado bien en la Politica 1502. Grilla " & NroGrilla
        Exit Sub
    End If
    Flog.writeline "La escala de Vacaciones está configurada correctamente en la Politica 1502. Grilla " & NroGrilla
    
    Call Politica(1505)
    If Not PoliticaOK Then
        Flog.writeline "Error cargando configuracion de la Politica 1505. Tipo Base antiguedad estandar."
        VersionBaseAntig = 0
    Else
        VersionBaseAntig = st_BaseAntiguedad
    End If
    
    
    'El tipo Base de la antiguedad
    TipoBase = 4
    
    continuar = True
    ant = 1
    Do While (ant <= rs_cabgrilla!cgrdimension) And continuar
        Select Case ant
        Case 1:
            StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_1
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    continuar = False
                Else
                    ant = ant + 1
                End If
            End If
            
        Case 2:
            StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_2
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    continuar = False
                Else
                    ant = ant + 1
                End If
            End If
        Case 3:
            StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_3
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    continuar = False
                Else
                    ant = ant + 1
                End If
            End If
        Case 4:
            StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_4
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    continuar = False
                Else
                    ant = ant + 1
                End If
            End If
        Case 5:
            StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_5
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    continuar = False
                Else
                    ant = ant + 1
                End If
            End If
        End Select
    Loop
    
            
    'Setea la proporcion de dias
    Call Politica(1501)

    For j = 1 To rs_cabgrilla!cgrdimension
        Select Case j
        Case ant:
            'la busqueda es la de antiguedad
            Select Case VersionBaseAntig
            Case 0:
                Flog.writeline "Antiguedad Standard " ' Se computa al año actual
                Call bus_Antiguedad_R("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
                If (((antmes * 30) + antdia >= st_Dias) Or antanio <> 0) Then
                       antdia = 0
                       antmes = 0
                       antanio = 0
                       Call bus_Antiguedad("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
                 End If

            Case 1:
                Flog.writeline "Antiguedad Sin redondeo "
                Call bus_Antiguedad_R("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
                 If (((antmes * 30) + antdia) >= st_Dias) Or antanio <> 0 Then
                       antdia = 0
                       antmes = 0
                       antanio = 0
                      Call bus_Antiguedad_G("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
                 End If
            Case 2:
                Flog.writeline "Antiguedad Uruguay " ' Se computa al año anterior
                'Call bus_Antiguedad_G("VACACIONES", CDate("31/12/" & Year(fecha_desde)), antdia, antmes, antanio, q)
                Call bus_Antiguedad("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
            Case 3:
                 Flog.writeline "Antiguedad Standard " ' Se computa al año actual
                 Call bus_Antiguedad_R("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
                 If (((antmes * 30) + antdia) >= st_Dias) Or antanio <> 0 Then
                       antdia = 0
                       antmes = 0
                       antanio = 0
                       Call bus_Antiguedad("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
                 End If
            Case 4: ' Anguedad a una fecha dada por dia y mes del año
                Flog.writeline "Antiguedad a una fecha dada año siguiente"
                If Not (st_Dia = 0 Or st_Mes = 0) Then
                     Call bus_Antiguedad_R("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
                     If (((antmes * 30) + antdia >= st_Dias) Or antanio <> 0) Then
                       antdia = 0
                       antmes = 0
                       antanio = 0
                       Call bus_Antiguedad_G("VACACIONES", CDate(st_Dia & "/" & st_Mes & "/" & Periodo_Anio + 1), antdia, antmes, antanio, q)
                    End If
                 End If
            Case 5: ' Anguedad a una fecha dada por dia y mes del año
                Flog.writeline "Antiguedad a una fecha dada año actual"
                Call bus_Antiguedad_R("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
                If (((antmes * 30) + antdia >= st_Dias) Or antanio <> 0) Then
                       antdia = 0
                       antmes = 0
                       antanio = 0
                 Call bus_Antiguedad("VACACIONES", CDate(st_Dia & "/" & st_Mes & "/" & Periodo_Anio), antdia, antmes, antanio, q)
                 End If
             
            Case 7:
                Flog.writeline "Antiguedad Standard con fecha de alta desde empfaltagr de tabla empleado" ' Se computa al año actual
                Call bus_Antiguedad_RV7("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
                If (((antmes * 30) + antdia >= st_Dias) Or antanio <> 0) Then
                       antdia = 0
                       antmes = 0
                       antanio = 0
                       Call bus_Antiguedad("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
                 End If
                 
            Case Else
                Flog.writeline "Antiguedad Mal configurada. Estandar "
                'Call bus_Antiguedad("VACACIONES", CDate("31/12/" & Year(fecha_desde)), antdia, antmes, antanio, q)
                Call bus_Antiguedad("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
            End Select

            parametros(j) = (antanio * 12) + antmes
            
            Flog.writeline "Años " & antanio
            Flog.writeline "Meses " & antmes
            Flog.writeline "Dias " & antdia

        Case Else:
            Select Case j
            Case 1:
                Call bus_Estructura(rs_cabgrilla!grparnro_1)
            Case 2:
                Call bus_Estructura(rs_cabgrilla!grparnro_2)
            Case 3:
                Call bus_Estructura(rs_cabgrilla!grparnro_3)
            Case 4:
                Call bus_Estructura(rs_cabgrilla!grparnro_4)
            Case 5:
                Call bus_Estructura(rs_cabgrilla!grparnro_5)
            End Select
            parametros(j) = valor
        End Select
    Next j

    'Busco la primera antiguedad de la escala menor a la del empleado
    ' de abajo hacia arriba
    StrSql = "SELECT * FROM valgrilla WHERE cgrnro = " & NroGrilla
    For j = 1 To rs_cabgrilla!cgrdimension
        If j <> ant Then
            StrSql = StrSql & " AND vgrcoor_" & j & "= " & parametros(j)
        End If
    Next j
        StrSql = StrSql & " ORDER BY vgrcoor_" & ant & " DESC "
    OpenRecordset StrSql, rs_valgrilla


    Encontro = False
    Do While Not rs_valgrilla.EOF And Not Encontro
        Select Case ant
        Case 1:
            If parametros(ant) >= rs_valgrilla!vgrcoor_1 Then
                 If rs_valgrilla!vgrvalor <> 0 Then
                    cantdias = rs_valgrilla!vgrvalor
                    Encontro = True
                    Columna = rs_valgrilla!vgrorden
                 End If
            End If
        Case 2:
            If parametros(ant) >= rs_valgrilla!vgrcoor_2 Then
                 If rs_valgrilla!vgrvalor <> 0 Then
                    cantdias = rs_valgrilla!vgrvalor
                    Encontro = True
                    Columna = rs_valgrilla!vgrorden
                 End If
            End If
        Case 3:
            If parametros(ant) >= rs_valgrilla!vgrcoor_3 Then
                 If rs_valgrilla!vgrvalor <> 0 Then
                    cantdias = rs_valgrilla!vgrvalor
                    Encontro = True
                    Columna = rs_valgrilla!vgrorden
                 End If
            End If
        Case 4:
            If parametros(ant) >= rs_valgrilla!vgrcoor_4 Then
                 If rs_valgrilla!vgrvalor <> 0 Then
                    cantdias = rs_valgrilla!vgrvalor
                    Encontro = True
                    Columna = rs_valgrilla!vgrorden
                 End If
            End If
        Case 5:
            If parametros(ant) >= rs_valgrilla!vgrcoor_5 Then
                 If rs_valgrilla!vgrvalor <> 0 Then
                    cantdias = rs_valgrilla!vgrvalor
                    Encontro = True
                    Columna = rs_valgrilla!vgrorden
                 End If
            End If
        End Select

        rs_valgrilla.MoveNext
    Loop

    
    
    '------------------------------
    'llamada politica 1513
    '------------------------------
    
    'EAM- Tiene en cuenta los dias trabajados en el ultimo año
    Call Politica(1513)
    
    If Dias_efect_trab_anio Then
        Flog.writeline "Tiene en cuenta el ultimo año. Politica 1513."
        antdia = 0
        antmes = 0
        antanio = 0
        
        StrSql = "SELECT * FROM vacacion WHERE vacnro = " & NroVac
        OpenRecordset StrSql, rs_vacacion
        If Not rs_vacacion.EOF Then
            cantdias = bus_Antiguedad_Col(CDate(rs_vacacion("vacfecdesde")), CDate(rs_vacacion("vacfechasta")))
        Else
            Flog.writeline "Empleado " & Ternro & " no se encontro vacacion: " & NroVac
        End If
        
        If cantdias >= 360 Then
            cantdias = 15
        Else
            cantdias = cantdias / 24
        End If
        
        Aux_Dias_trab = cantdias
        cantdias = RedondearNumero(Int(Aux_Dias_trab), (Aux_Dias_trab - Int(Aux_Dias_trab)))
        
        Encontro = True
        Flog.writeline "Empleado " & Ternro '& " con meses trabajado en el último año: " & antmes
        Flog.writeline "Días Correspondientes:" & cantdias
        Flog.writeline "Tipo de redondeo:" & st_redondeo
        Flog.writeline
    End If
    
    If Not Encontro Then
        'Aux_Dias_trab = ((((antmes * 30) + antdia) / 30)) * 1.6667
        'cantdias = RedondearNumero(Int(Aux_Dias_trab), (Aux_Dias_trab - Int(Aux_Dias_trab)))
        cantdias = 0
        Flog.writeline "Dias Proporcion " & 0
    End If
                        
   
Genera = True
    
' Cierro todo y libero
If rs_cabgrilla.State = adStateOpen Then rs_cabgrilla.Close
If rs_valgrilla.State = adStateOpen Then rs_valgrilla.Close
If rs_vacacion.State = adStateOpen Then rs_vacacion.Close

Set rs_cabgrilla = Nothing
Set rs_valgrilla = Nothing
Set rs_vacacion = Nothing
End Sub


'EAM- Calcula los dias de Beneficio de vacaciones para los Empleados con entrada anteriores a la fecha (01/01/2007)
'SYKES - COSTA RICA
Sub CalcularBeneficioVac(ByVal Ternro As Long, ByVal NroVac As Long, ByVal tipoVac As Integer, ByVal Reproceso As Boolean, ByVal vdiasfechasta As Date)
'Modificado: 24/05/2012 - Gonzalez Nicolás - se toma vdiasfechasta (Fecha desde del periodo de vacaciones) para calcular los dias de antiguedad.

Dim fechaAlta As Date
Dim cantMeses As Long
Dim diasBeneficio As Integer
Dim l_rsAux As New ADODB.Recordset

    'EAM- Obtiene la fecha de Ingreso
    fechaAlta = FechaAltaEmpleado(Ternro)
    
    If fechaAlta < CDate("01/01/2007") Then
    
        'Le resto 1 día a la fecha desde del periodo
        Flog.writeline "Antes de restar un dia: " & vdiasfechasta
        vdiasfechasta = DateAdd("d", -1, vdiasfechasta)
        Flog.writeline "Despues de restar un dia: " & vdiasfechasta
        'Calcula la antiguedad en Años
        'cantMeses = DateDiff("m", fechaAlta, Date)
        cantMeses = DateDiff("m", fechaAlta, vdiasfechasta)
                
        'Obtiene los dias de Beneficio
     '   Select Case cantMeses
      '      Case Is < 24:
       '         diasBeneficio = 0
       '     Case Is <= 24:
       '         diasBeneficio = 1
       '     Case Is <= 36:
       '         diasBeneficio = 2
       '     Case Is <= 48:
       '         diasBeneficio = 3
       '     Case Is <= 60:
       '         diasBeneficio = 4
       '     Case Is <= 72:
       '         diasBeneficio = 5
       '     Case Is <= 84:
       '         diasBeneficio = 6
       '     Case Is <= 96:
       '         diasBeneficio = 7
       '     Case Is <= 108:
       '         diasBeneficio = 8
       '     Case Is <= 120:
       '         diasBeneficio = 9
       '     Case Else
       '         diasBeneficio = 10
       ' End Select
       
       '------------------------- mdf inicio
       Select Case cantMeses
             
            Case 0 To 23
                  diasBeneficio = 0
            
            Case 24 To 35
                  diasBeneficio = 1
            
            Case 36 To 47
                  diasBeneficio = 2
            
            Case 48 To 59
                  diasBeneficio = 3
            
            Case 60 To 71
                  diasBeneficio = 4
            
            Case 72 To 83
                  diasBeneficio = 5
            
            Case 84 To 95
                  diasBeneficio = 6
            
            Case 96 To 107
                  diasBeneficio = 7
            
            Case 108 To 119
                  diasBeneficio = 8
            
            Case 120 To 131
                  diasBeneficio = 9
            
            Case Else
                  diasBeneficio = 10
                  
        End Select
        Flog.writeline "Corresponden " & diasBeneficio & " dias de beneficio por antiguedad"
       
       '-------------------------------------------mdf fin
        
        StrSql = "SELECT * FROM vacdiascor WHERE vacnro= " & NroVac & " AND ternro= " & Ternro & " AND venc=3"
        OpenRecordset StrSql, l_rsAux
        
        If l_rsAux.EOF Then
            StrSql = "INSERT INTO vacdiascor (vacnro,ternro,vdiascorcant,vdiapednro,vdiascormanual,tipvacnro,venc,vdiasfechasta) " & _
                    "VALUES(" & NroVac & "," & Ternro & "," & diasBeneficio & ",0,0," & tipoVac & ",3," & ConvFecha(vdiasfechasta) & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
            'Flog.writeline "Dias de Beneficio: " & diasBeneficio & " a la fecha " & Date
            Flog.writeline "Dias de Beneficio: " & diasBeneficio & " a la fecha " & vdiasfechasta
            
            Flog.writeline ""
            Flog.writeline StrSql
            Flog.writeline ""
        Else
            Flog.writeline "Los días de Beneficio ya estan calculados."
            If Reproceso Then
                StrSql = "UPDATE vacdiascor SET vdiascorcant= " & diasBeneficio & ",vdiasfechasta=" & ConvFecha(vdiasfechasta) & _
                        " WHERE ternro= " & Ternro & " AND vacnro = " & NroVac & " AND venc=3"
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
        End If
    Else
             Flog.writeline "La fecha de alta es mayor al 01/01/2007"
       
    End If
End Sub

Sub CalcularBeneficioVac_PT(ByVal Ternro As Long, ByVal NroVac As Long, ByVal tipoVac As Integer, ByVal fecha_desde As Date, fecha_hasta As Date, ByVal Lic_Descuento As String, ByVal Reproceso As Boolean, ByVal vdiasfechasta As Date)
' ---------------------------------------------------------------------------------------------
' Descripcion: Calcula los dias de Beneficio de vacaciones para los Empleados que tienen 1 año calendario completo o más. PORTUGAL
' Autor      : Gonzalez Nicolás
' Fecha      : 10/05/2012
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
'Dim fechaAlta As Date
Dim cantMeses As Long
Dim diasBeneficio As Integer
Dim l_rsAux As New ADODB.Recordset

    'EAM- Obtiene la fecha de Ingreso
    'fechaAlta = FechaAltaEmpleado(Ternro)
        '---------------------------------------
        'BUSCA LICENCIAS QUE GENEREN DESCUENTO EN EL PLUS
        '---------------------------------------
        StrSql = " SELECT COUNT(*) total from emp_lic"
        StrSql = StrSql & " WHERE"
        StrSql = StrSql & " tdnro NOT IN (" & Lic_Descuento & ")"
        StrSql = StrSql & " AND empleado = " & Ternro
        StrSql = StrSql & " AND ( "
        StrSql = StrSql & " (elfechadesde >= " & ConvFecha(fecha_desde) & " AND elfechahasta <= " & ConvFecha(fecha_hasta) & " ) "
        StrSql = StrSql & " OR (elfechadesde >= " & ConvFecha(fecha_desde) & " AND elfechadesde <= " & ConvFecha(fecha_hasta) & " and elfechahasta >= " & ConvFecha(fecha_hasta) & ")"
        StrSql = StrSql & " OR (elfechadesde <= " & ConvFecha(fecha_desde) & " AND elfechahasta <= " & ConvFecha(fecha_hasta) & ")"
        StrSql = StrSql & ")"
        OpenRecordset StrSql, l_rsAux
        '---------------------------------------
        'Obtiene los dias de Beneficio
        If l_rsAux.EOF Then
            diasBeneficio = 3
        Else
            diasBeneficio = 3 - l_rsAux!total
        End If
        
        If diasBeneficio <= 0 Then
            Flog.writeline "El empleado " & Ternro & " No tiene días de beneficio"
            Exit Sub
        
        End If

        
    'BUSCA SI EXISTEN DIAS DE PLUS GENERADOS.
    StrSql = "SELECT * FROM vacdiascor WHERE vacnro= " & NroVac & " AND ternro= " & Ternro & " AND venc=3"
    OpenRecordset StrSql, l_rsAux
        
    If l_rsAux.EOF Then
        StrSql = "INSERT INTO vacdiascor (vacnro,ternro,vdiascorcant,vdiapednro,vdiascormanual,tipvacnro,venc,vdiasfechasta) " & _
         "VALUES(" & NroVac & "," & Ternro & "," & diasBeneficio & ",0,0," & tipoVac & ",3," & ConvFecha(vdiasfechasta) & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        Flog.writeline "Dias de Beneficio: " & diasBeneficio & " a la fecha " & Date
     Else
        If Reproceso = True Then
            Flog.writeline "Reproceso días de Beneficio."
            StrSql = "UPDATE vacdiascor SET vdiascorcant= " & diasBeneficio & ",vdiasfechasta=" & ConvFecha(vdiasfechasta) & _
                        " WHERE ternro= " & Ternro & " AND vacnro = " & NroVac & " AND venc=3"
                objConn.Execute StrSql, , adExecuteNoRecords
            
        Else
            Flog.writeline "Los días de Beneficio ya estan calculados."
        End If
        
     End If

End Sub
Sub CalcularBeneficioVac_UY(ByVal Ternro As Long, ByVal NroVac As Long, ByVal tipoVac As Integer, ByVal fecha_desde As Date, fecha_hasta As Date, ByVal Lic_Descuento As String, ByVal Reproceso As Boolean, ByVal vdiasfechasta As Date)
' ---------------------------------------------------------------------------------------------
' Descripcion: Calcula los dias de Beneficio de vacaciones para los Empleados que tienen 1 año calendario completo o más. PORTUGAL
' Autor      : Gonzalez Nicolás
' Fecha      : 10/05/2012
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
'Dim fechaAlta As Date
Dim cantMeses As Long
Dim diasBeneficio As Integer
Dim l_rsAux As New ADODB.Recordset


Dim Dia As Long
Dim mes As Long
Dim Anio As Long

Dim Cant As Integer
Dim diasBenef As Long
Dim cantAnios As Long
Dim cantdias As Integer
Dim Columna As Integer
Dim Mensaje As String
Dim Genera As Boolean

Dim antdia As Long
Dim antmes As Long
Dim antanio As Long
Dim q As Integer

Select Case st_BaseAntiguedad
    Case 0:
        Flog.writeline "Antiguedad Standard " ' Se computa al año actual
        Call bus_Antiguedad_R("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
        If (((antmes * 30) + antdia >= st_Dias) Or antanio <> 0) Then
               antdia = 0
               antmes = 0
               antanio = 0
               Call bus_Antiguedad("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
         End If
    
    Case 1:
        Flog.writeline "Antiguedad Sin redondeo "
        Call bus_Antiguedad_R("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
         If (((antmes * 30) + antdia) >= st_Dias) Or antanio <> 0 Then
               antdia = 0
               antmes = 0
               antanio = 0
              Call bus_Antiguedad_G("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
         End If
    Case 2:
        Flog.writeline "Antiguedad Uruguay " ' Se computa al año anterior
        'Call bus_Antiguedad_G("VACACIONES", CDate("31/12/" & Year(fecha_desde)), antdia, antmes, antanio, q)
        Call bus_Antiguedad("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
    Case 3:
         Flog.writeline "Antiguedad Standard " ' Se computa al año actual
         Call bus_Antiguedad_R("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
         If (((antmes * 30) + antdia) >= st_Dias) Or antanio <> 0 Then
               antdia = 0
               antmes = 0
               antanio = 0
               Call bus_Antiguedad("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
         End If
    Case 4: ' Anguedad a una fecha dada por dia y mes del año
        Flog.writeline "Antiguedad a una fecha dada año siguiente"
        If Not (st_Dia = 0 Or st_Mes = 0) Then
             Call bus_Antiguedad_R("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
             If (((antmes * 30) + antdia >= st_Dias) Or antanio <> 0) Then
               antdia = 0
               antmes = 0
               antanio = 0
               Call bus_Antiguedad_G("VACACIONES", CDate(st_Dia & "/" & st_Mes & "/" & Periodo_Anio + 1), antdia, antmes, antanio, q)
            End If
         End If
    Case 5: ' Anguedad a una fecha dada por dia y mes del año
        Flog.writeline "Antiguedad a una fecha dada año actual"
        Call bus_Antiguedad_R("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
        If (((antmes * 30) + antdia >= st_Dias) Or antanio <> 0) Then
               antdia = 0
               antmes = 0
               antanio = 0
         Call bus_Antiguedad("VACACIONES", CDate(st_Dia & "/" & st_Mes & "/" & Periodo_Anio), antdia, antmes, antanio, q)
         End If
    
    Case Else
        Flog.writeline "Antiguedad Mal configurada. Estandar "
        'Call bus_Antiguedad("VACACIONES", CDate("31/12/" & Year(fecha_desde)), antdia, antmes, antanio, q)
        Call bus_Antiguedad("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
End Select

'parametros(j) = (antanio * 12) + antmes

Flog.writeline "Años " & antanio
Flog.writeline "Meses " & antmes
Flog.writeline "Dias " & antdia



'If antanio = 4 Then
'    diasBeneficio = 66
'End If
'Flog.writeline "Corresponden " & diasBeneficio & " dias de beneficio por antiguedad"
   
If antanio = 25 Then
    diasBeneficio = 66
    Flog.writeline "Corresponden " & diasBeneficio & " dias de beneficio por antiguedad"
    StrSql = "SELECT * FROM vacdiascor WHERE vacnro= " & NroVac & " AND ternro= " & Ternro & " AND venc=0"
    OpenRecordset StrSql, l_rsAux
    If Not l_rsAux.EOF Then
        diasBeneficio = 66 - CDbl(l_rsAux!vdiascorcant)
    Else
        diasBeneficio = 66
    End If
    StrSql = "SELECT * FROM vacdiascor WHERE vacnro= " & NroVac & " AND ternro= " & Ternro & " AND venc=3"
    OpenRecordset StrSql, l_rsAux

    If l_rsAux.EOF Then
        StrSql = "INSERT INTO vacdiascor (vacnro,ternro,vdiascorcant,vdiapednro,vdiascormanual,tipvacnro,venc,vdiasfechasta) " & _
                "VALUES(" & NroVac & "," & Ternro & "," & diasBeneficio & ",0,0," & tipoVac & ",3," & ConvFecha(vdiasfechasta) & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        Flog.writeline "Dias de Beneficio: " & diasBeneficio & " a la fecha " & vdiasfechasta

        Flog.writeline ""
        Flog.writeline StrSql
        Flog.writeline ""
    Else
        Flog.writeline "Los días de Beneficio ya estan calculados."
        If Reproceso Then
            StrSql = "UPDATE vacdiascor SET vdiascorcant= " & diasBeneficio & ",vdiasfechasta=" & ConvFecha(vdiasfechasta) & _
                    " WHERE ternro= " & Ternro & " AND vacnro = " & NroVac & " AND venc=3"
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
    End If

End If
   
   
   
   '-------------------------------------------mdf fin
'
'    StrSql = "SELECT * FROM vacdiascor WHERE vacnro= " & NroVac & " AND ternro= " & Ternro & " AND venc=3"
'    OpenRecordset StrSql, l_rsAux
'
'    If l_rsAux.EOF Then
'        StrSql = "INSERT INTO vacdiascor (vacnro,ternro,vdiascorcant,vdiapednro,vdiascormanual,tipvacnro,venc,vdiasfechasta) " & _
'                "VALUES(" & NroVac & "," & Ternro & "," & diasBeneficio & ",0,0," & tipoVac & ",3," & ConvFecha(vdiasfechasta) & ")"
'        objConn.Execute StrSql, , adExecuteNoRecords
'        'Flog.writeline "Dias de Beneficio: " & diasBeneficio & " a la fecha " & Date
'        Flog.writeline "Dias de Beneficio: " & diasBeneficio & " a la fecha " & vdiasfechasta
'
'        Flog.writeline ""
'        Flog.writeline StrSql
'        Flog.writeline ""
'    Else
'        Flog.writeline "Los días de Beneficio ya estan calculados."
'        If Reproceso Then
'            StrSql = "UPDATE vacdiascor SET vdiascorcant= " & diasBeneficio & ",vdiasfechasta=" & ConvFecha(vdiasfechasta) & _
'                    " WHERE ternro= " & Ternro & " AND vacnro = " & NroVac & " AND venc=3"
'            objConn.Execute StrSql, , adExecuteNoRecords
'        End If
'    End If
'Else
'         Flog.writeline "La fecha de alta es mayor al 01/01/2007"
'
'End If



'Call bus_Antiguedad_R("VACACIONES", CDate("31/12/" & Periodo_Anio), dia, mes, anio, cant)

'If anio = 5 Then
'    StrSql = "SELECT * FROM vacdiascor WHERE vacnro= " & NroVac & " AND ternro= " & Ternro & " AND venc=0"
'    OpenRecordset StrSql, l_rsAux
'    If Not l_rsAux.EOF Then
'        diasBeneficio = 66 - CDbl(l_rsAux!vdiascorcat)
'    Else
'        diasBeneficio = 66
'    End If
'    'inserto los dias de beneficios
'    StrSql = "INSERT INTO vacdiascor (vacnro,ternro,vdiascorcant,vdiapednro,vdiascormanual,tipvacnro,venc,vdiasfechasta) " & _
'     "VALUES(" & NroVac & "," & Ternro & "," & diasBeneficio & ",0,0," & tipoVac & ",3," & ConvFecha(vdiasfechasta) & ")"
'    objConn.Execute StrSql, , adExecuteNoRecords
'    Flog.writeline "Dias de Beneficio: " & diasBeneficio & " a la fecha " & Date
'End If



    'EAM- Obtiene la fecha de Ingreso
    'fechaAlta = FechaAltaEmpleado(Ternro)
        '---------------------------------------
        'BUSCA LICENCIAS QUE GENEREN DESCUENTO EN EL PLUS
        '---------------------------------------
'        StrSql = " SELECT COUNT(*) total from emp_lic"
'        StrSql = StrSql & " WHERE"
'        StrSql = StrSql & " tdnro NOT IN (" & Lic_Descuento & ")"
'        StrSql = StrSql & " AND empleado = " & Ternro
'        StrSql = StrSql & " AND ( "
'        StrSql = StrSql & " (elfechadesde >= " & ConvFecha(fecha_desde) & " AND elfechahasta <= " & ConvFecha(fecha_hasta) & " ) "
'        StrSql = StrSql & " OR (elfechadesde >= " & ConvFecha(fecha_desde) & " AND elfechadesde <= " & ConvFecha(fecha_hasta) & " and elfechahasta >= " & ConvFecha(fecha_hasta) & ")"
'        StrSql = StrSql & " OR (elfechadesde <= " & ConvFecha(fecha_desde) & " AND elfechahasta <= " & ConvFecha(fecha_hasta) & ")"
'        StrSql = StrSql & ")"
'        OpenRecordset StrSql, l_rsAux
'        '---------------------------------------
'        'Obtiene los dias de Beneficio
'        If l_rsAux.EOF Then
'            diasBeneficio = 3
'        Else
'            diasBeneficio = 3 - l_rsAux!total
'        End If
'
'        If diasBeneficio <= 0 Then
'            Flog.writeline "El empleado " & Ternro & " No tiene días de beneficio"
'            Exit Sub
'
'        End If
'
'
'    'BUSCA SI EXISTEN DIAS DE PLUS GENERADOS.
'    StrSql = "SELECT * FROM vacdiascor WHERE vacnro= " & NroVac & " AND ternro= " & Ternro & " AND venc=3"
'    OpenRecordset StrSql, l_rsAux
'
'    If l_rsAux.EOF Then
'        StrSql = "INSERT INTO vacdiascor (vacnro,ternro,vdiascorcant,vdiapednro,vdiascormanual,tipvacnro,venc,vdiasfechasta) " & _
'         "VALUES(" & NroVac & "," & Ternro & "," & diasBeneficio & ",0,0," & tipoVac & ",3," & ConvFecha(vdiasfechasta) & ")"
'        objConn.Execute StrSql, , adExecuteNoRecords
'        Flog.writeline "Dias de Beneficio: " & diasBeneficio & " a la fecha " & Date
'     Else
'        If Reproceso = True Then
'            Flog.writeline "Reproceso días de Beneficio."
'            StrSql = "UPDATE vacdiascor SET vdiascorcant= " & diasBeneficio & ",vdiasfechasta=" & ConvFecha(vdiasfechasta) & _
'                        " WHERE ternro= " & Ternro & " AND vacnro = " & NroVac & " AND venc=3"
'                objConn.Execute StrSql, , adExecuteNoRecords
'
'        Else
'            Flog.writeline "Los días de Beneficio ya estan calculados."
'        End If
'
'     End If

End Sub
'EAM- busca en la escala de vacaciones los dias correspondientes segun el tipo de vacaciones pasado por parametro
Public Function buscarDiasVacEscala(ByVal ant As Integer, ByVal dimensionEscala, ByVal parametros, ByVal tipovacacion As Long, Optional ByRef Encontro As Boolean) As Integer
 '07/11/2013 - Gonzalez Nicolás - Se cambio de Private a Public
 
 Dim rs_valgrilla As New ADODB.Recordset
 Dim j, cantdias As Integer
 
    
    'Busco la primera antiguedad de la escala menor a la del empleado de abajo hacia arriba
    StrSql = "SELECT * FROM valgrilla WHERE cgrnro = " & NroGrilla
    For j = 1 To dimensionEscala
        If j <> ant Then
            StrSql = StrSql & " AND vgrcoor_" & j & "= " & parametros(j)
        End If
    Next j
    StrSql = StrSql & " AND vgrorden= " & tipovacacion & "  ORDER BY vgrcoor_" & ant & " DESC "
    OpenRecordset StrSql, rs_valgrilla


    Encontro = False
    cantdias = 0
    Do While Not rs_valgrilla.EOF And Not Encontro
        Select Case ant
        Case 1:
            If parametros(ant) >= rs_valgrilla!vgrcoor_1 Then
                 If rs_valgrilla!vgrvalor <> 0 Then
                    cantdias = rs_valgrilla!vgrvalor
                    Encontro = True
                 End If
            End If
        Case 2:
            If modeloPais <> 7 Then 'mdf
                If parametros(ant) >= rs_valgrilla!vgrcoor_2 Then
                     If rs_valgrilla!vgrvalor <> 0 Then
                        cantdias = rs_valgrilla!vgrvalor
                        Encontro = True
                     End If
                End If
            Else   'mdf
                  If (rs_valgrilla!vgrcoor_2 - parametros(ant)) <= 6 Or (rs_valgrilla!vgrcoor_2 - parametros(ant)) < 0 Then
                     If rs_valgrilla!vgrvalor <> 0 Then
                        cantdias = rs_valgrilla!vgrvalor
                        Encontro = True
                     End If
                End If
            End If
        Case 3:
            If parametros(ant) >= rs_valgrilla!vgrcoor_3 Then
                 If rs_valgrilla!vgrvalor <> 0 Then
                    cantdias = rs_valgrilla!vgrvalor
                    Encontro = True
                 End If
            End If
        Case 4:
            If parametros(ant) >= rs_valgrilla!vgrcoor_4 Then
                 If rs_valgrilla!vgrvalor <> 0 Then
                    cantdias = rs_valgrilla!vgrvalor
                    Encontro = True
                 End If
            End If
        Case 5:
            If parametros(ant) >= rs_valgrilla!vgrcoor_5 Then
                 If rs_valgrilla!vgrvalor <> 0 Then
                    cantdias = rs_valgrilla!vgrvalor
                    Encontro = True
                 End If
            End If
        End Select
                    
        rs_valgrilla.MoveNext
    Loop
        
    buscarDiasVacEscala = cantdias
    
    rs_valgrilla.Close
    Set rs_valgrilla = Nothing
    
End Function


Function cantDiasLaborable(ByVal tipvacnro As Long, ByRef ExcluyeFeriados As Boolean) As Integer
 Dim rsHabiles As New ADODB.Recordset
 Dim diasLab As Integer
 
    
    StrSql = "SELECT * FROM tipovacac WHERE tipvacnro = " & tipvacnro
    OpenRecordset StrSql, rsHabiles
    
    'EAM- Analiza los dias de la semana que son laborable para el tipo de Vac.
    If Not rsHabiles.EOF Then
        If rsHabiles!tpvhabiles__1 Then diasLab = diasLab + 1
        If rsHabiles!tpvhabiles__2 Then diasLab = diasLab + 1
        If rsHabiles!tpvhabiles__3 Then diasLab = diasLab + 1
        If rsHabiles!tpvhabiles__4 Then diasLab = diasLab + 1
        If rsHabiles!tpvhabiles__5 Then diasLab = diasLab + 1
        If rsHabiles!tpvhabiles__6 Then diasLab = diasLab + 1
        If rsHabiles!tpvhabiles__7 Then diasLab = diasLab + 1
                
        ExcluyeFeriados = CBool(rsHabiles!tpvferiado)
    Else
        diasLab = 7
    End If
    cantDiasLaborable = diasLab
End Function
