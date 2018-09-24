Attribute VB_Name = "mdlDiasCorresp"
Option Explicit
'Version: 1.01  'Inicial

'Const Version = 1.01    'Dias Correspondientes Parciales
'Const FechaVersion = "15/10/2005"

'Const Version = 1.02    'Version con otra conexion para el progreso
'Const FechaVersion = "23/11/2005"


'Const Version = 2.01     'Revision general
'Const FechaVersion = "01/12/2005"

'Const Version = 2.02     'se calcula la fecha_hasta como 31/12/anio del periodo de vac
'Const FechaVersion = "25/01/2006"

'Const Version = 2.03     'los dias habiles considerados para hacer la proporcion de dias cuando
                         'la antiguedad es menor a 6 meses
                         'se saca de la configuracion del tipo de vacaciones (Corridos, Habiles de L-V, de L-S, etc) de acuerdo a un tercer parametro en la politica 1501
'Const FechaVersion = "16/02/2006"

'Const Version = 2.04     'se agrega nueva version de la politica 1505
                         'se agrega un nuevo parametro a la politica 1501
                         'se crea la politica 1508, se agrego la logica de la misma
                         'se agrego dos nuevos tipos de parametros:
                         '       18-BaseAntiguedad (Int)
                         '       19-Factor (Double)
                         'se agrego un case a la Version Base Antiguedad que usa Uruguay
'Const FechaVersion = "29/06/2006"

'Const Version = 2.05    'Lisandro Moro
'                        'Se corrigio la forma de calculode las vacaciones
'                        ' en la sub Bus_diasVac, cuando DiasProporcion <> 20
'Const FechaVersion = "13/11/2006"

'---------------------------------------------------------------
'Const Version = "2.06"
'Const FechaVersion = "13/11/2007" 'FGZ
'Se cambió la fecha para la cual se resuelve el alcance por estructura de las politicas (sub politica)
'               Se cambió el uso de fecha_desde en los querys por aux_fecha
'                If fecha_desde > Date Then
'                    Aux_fecha = fecha_desde
'                Else
'                    If fecha_hasta > Date Then
'                        Aux_fecha = Date
'                    Else
'                        Aux_fecha = fecha_hasta
'                    End If
'                End If
'---------------------------------------------------------------

'----------------------------------------------------------------------------------------
'Const Version = "2.07"
'Const FechaVersion = "29/01/2008"
' Gustavo Ring - Se agrego redondeo para calcular los dias correspondientes
'                cuando el empleado tiene menos de 6 meses trabajados
'----------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------
'Const Version = "2.08"
'Const FechaVersion = "01/02/2008"
' Gustavo Ring - Se cambio el nro de parámetro 22 redondeo por el 16 que ya existia
'----------------------------------------------------------------------------------------

'Const Version = "2.09"
'Const FechaVersion = "24/02/2009"
' Gustavo Ring - Se creo custom para calcular antigüedad para Radiotronica
'----------------------------------------------------------------------------------------

'Const Version = "2.10"
'Const FechaVersion = "11/03/2009"
' Gustavo Ring - Se modifico antigüedad para Radiotronica
'----------------------------------------------------------------------------------------

'Const Version = "2.11"
'Const FechaVersion = "12/03/2009"
' Gustavo Ring - Soporte encriptación de cadena de conexión
'----------------------------------------------------------------------------------------

'Const Version = "2.12"
'Const FechaVersion = "14/04/2009"
' Gustavo Ring - Se modifico para tomar o no los dias feriados como trabajados.
'----------------------------------------------------------------------------------------

'Const Version = "2.13"
'Const FechaVersion = "07/05/2009"
' Gustavo Ring - Se modifico la query de las fases.
'----------------------------------------------------------------------------------------

'Const Version = "2.14"
'Const FechaVersion = "05/06/2009"
''Gustavo Ring - Se utiliza el parametro 11 para la ver cuando se proporciona
''----------------------------------------------------------------------------------------


'Const Version = "2.15"
'Const FechaVersion = "25/06/2009" 'FGZ
''           Nueva Politica 1511 - Vacaciones Acordadas.
''               Esta politica revisa las vacaciones acordadas del empleado
''                   y se queda con lo mas conveniente para el empleado.


'Const Version = "2.16"
'Const FechaVersion = "22/10/2009" 'FGZ
''           Nueva Politica 1512 - Vencimiento de vacaciones.
''               Esta politica calcula la cantidad de dias que vencen del periodo anterior y traspasa lo que se pueda al periodo actual.

'Const Version = "2.17"
'Const FechaVersion = "16/11/2009" 'FGZ
''           Problema con la funcion de validacion de version.

'Const Version = "2.18"
'Const FechaVersion = "09/02/2010" 'MB
''           Politica 1508 se agregó la opcion de base de antiguedad 4 y 5 a una fecha dada con los paramtros 30-dia y 31-mes.
''           la base 4 calcula la fecha como dia/mes/año del periodo + 1 y la base 5 calcula la fecha como dia/mes/año del periodo


'Const Version = "2.19"
'Const FechaVersion = "04/03/2010" 'FGZ
''           Integracion de la version anterior.
''               sincronizacion de numeros de parametros

'------------------------------------------------------------------------------------
'Const Version = "3.00"
'Const FechaVersion = "14/04/2010" 'FGZ
'           Ahora los periodos de vacaciones ahora pueden tener alcance por estructura
'           hubo que agregar 2 parametros mas


'Const Version = "3.01"
'Const FechaVersion = "15/07/2010" 'Margiotta, Emanuel
'''           Nuevo manera para determinar los días correspondientes para el caso de empleados que hayan trabajado menos de la mitad del año.
'''           Se controla los dias efectivamente trabajados en el ultimo año

'Const Version = "3.02"
'Const FechaVersion = "27/08/2010" 'Margiotta, Emanuel
'''           Se agregó el cálculo de dias correspondientes para distintos paises.
'''           Se agregó el cálculo de días corresp. para Uruguay.

'Const Version = "3.03"
'Const FechaVersion = "23/09/2010" 'FGZ
''           Se agregó detalle de log cuando levanta los parametros

'Const Version = "3.04"
'Const FechaVersion = "08/10/2010" 'Margiotta, Emanuel
''           Se comentó una linea en la función de bus_DiasVac_uy la busqueda de antiguedad 2 "Uruguay" porque hacia 2 busquedas seguidas

'Const Version = "3.05"
'Const FechaVersion = "08/10/2010" 'Margiotta, Emanuel
''           Se agregó la validación en la política 1513 cuando no tiene configurada la cantidad de hs diarias

'Const Version = "3.06"
'Const FechaVersion = "04/11/2010" 'Margiotta, Emanuel
''           Se Corrigió la descripción del tipo de dia de vacaciones cuando no lo encuentra en la escala y lo tiene que levantar de la Politica 1501
''           Se agrego a la Pol. 1501 el parametro de redondeo.

'Const Version = "3.07"
'Const FechaVersion = "04/11/2010" 'Margiotta, Emanuel
''           Se cambio en la funcion busqueda de días de vacaciones de uruguay El tipo Base de la antiguedad para que calcule los dias
''           correspondientes al periodo que se esta generando y no al anterior.


'Const Version = "3.08"
'Const FechaVersion = "16/11/2010" 'Margiotta, Emanuel
''           Se saco el calculo de Vencimiento de vacaciones para La Caja


'Const Version = "3.09"
'Const FechaVersion = "19/11/2010" 'FGZ
''           Se cambió el calculo proporcional de dias para Uruguay


'Const Version = "3.10"
'Const FechaVersion = "03/12/2010" 'FGZ
''           Politica 1505. Habia quedado mal los paramtros que utiliza ademas de que no estaba configurable.
''               Los parametros configurables que debe utilizar son
''                   35 - Dia de una fecha
''                   36 - Mes de una fecha

'Const Version = "3.11"
'Const FechaVersion = "12/04/2011" 'Lisandro Moro
''               cas 11821 - Se creo la version colombia

'Const Version = "3.12"
'Const FechaVersion = "12/04/2011" 'Margiotta, Emanuel
''               Se creo la version para Costa Rica - SYKE

'Const Version = "3.13"
'Const FechaVersion = "27/07/2011" 'FGZ
''               Hay un parametro que si viene en NULL aborta. Se controla
''               Se modificó el calculo de antiguedad en el ultimo año

'Const Version = "3.14"
'Const FechaVersion = "11/10/2011" 'EAM
''               Se modificó el calculo de vacaciones de Costa Rica

'Const Version = "3.15"
'Const FechaVersion = "18/10/2011" 'EAM
''               Se agregó la versión 2 de la politica 1513 para que se descuenten los días feriados - Andreani

'Const Version = "3.16"
'Const FechaVersion = "31/10/2011" 'EAM
''               Se corrigio para el procesos planificado de CR- ya que procesa todos los empleado y originariamente los levantaba de batch_empleado y es este caso no existen.

'Const Version = "3.17"
'Const FechaVersion = "06/12/2011" 'EAM
''               Se corrigio la politica 1501, estaba tomando el redondeo siempre con el valor 2
''               Se modifico el comentario del log para CR cuando no tiene Parte de Asignacion Horaria

'Const Version = "3.18"
'Const FechaVersion = "10/02/2012" 'EAM
''               CAS(13972) Se agrego a la política 1501 tipodia2 que sirve para buscar en la escala de vacaiones la cantidad de dias que le corresponde
''               segun el tipo de dia y en tipodia1 se configura el tipo de vacaciones por default.
''               Se agrego el calculo para los empleado con menos de 6 meses de trabajo (sin escala) y relacionado la politica 1513.


'Const Version = "3.19"
'Const FechaVersion = "17/04/2012" 'EAM
'               CAS(15527) Se modifico la función de bus_DiasVac_CR para los empleados con Asignación Horaria que busque los movimientos a partir
'               de la última fecha se proceso, ya que si es por semana completa Ej. del 22/03/12 al 29/03/12 y la primer semana tiene 5 movimientos
'               y la segunda 4 da como resultado que trabaja 9 dias a la semana y es incorrecto.
'               Se seteo la variable NroTPVCorr de la version 3.18 con la variable Columna2 ya que sino quedaba sin valor y daba error.
'               Se modifico el factor de división en el calculo de dias correspondientes por 350

Const Version = "3.20"
Const FechaVersion = "14/05/2012" 'Gonzalez Nicolás -  DEMO PORTUGAL
'                                  - Se creó módulo mdlValidarBD, el cual realiza el control de versiones.
'                                  - Se agregó versionado para la política 1514
'                                  - Se agregó función CalcularBeneficioVac_PT() - Calcula los dias PLUS
'                                  - Se agrego para el modelo standard | AnioaProc = Periodo_Anio
'                                  - Se agrego para el modelo standard | auxNroVac = NroVac
'------------------------------------------------------------------------------------
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



Public Sub Main()
Dim Fecha As Date
Dim cantdias As Integer
Dim cantdiasCorr As Integer
Dim CantdiasCR As Double
Dim DiasCorraGen As Double
Dim dias_maternidad As Integer
Dim Columna As Integer
Dim Columna2 As Integer
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
    
    'Activo el manejador de errores
    On Error GoTo CE
    
    'EAM - 18/01/2012 --------- Control de versiones ------
    'Version_Valida = ValidarV(Version, 10, TipoBD) 'NG - Ahora se valida desde la funcion ValidarVBD
    
    '___________________________________________________________________________________________________
    'Se agrego un parametro para determinar el país con el cual se quiere calcular el modelo de vacacion
    'El parametro 7 hace referencia al modelo de vacaciones configurado en la tabla confper
    modeloPais = Pais_Modelo(7)
    'NG - 04/05/2012 --------- Control de versiones ------
    Version_Valida = ValidarVBD(Version, 10, TipoBD, modeloPais)
        
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
    Flog.writeline "Parametros " & parametros
    
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
        StrSql = " SELECT * FROM batch_empleado WHERE batch_empleado.bpronro = " & NroProceso
        OpenRecordset StrSql, objReg
        
        If objReg.EOF Then
            objReg.Close
            StrSql = "SELECT distinct empleado.ternro FROM empleado WHERE empest= -1"
            OpenRecordset StrSql, objReg
        End If
    Else
        StrSql = " SELECT * FROM batch_empleado WHERE batch_empleado.bpronro = " & NroProceso
        OpenRecordset StrSql, objReg
    End If

    
    CEmpleadosAProc = objReg.RecordCount
    IncPorc = (100 / CEmpleadosAProc)
    
    SinError = True
    HuboErrores = False
    Do While Not objReg.EOF
   
        Ternro = objReg!Ternro
        Empleado.Ternro = objReg!Ternro
        
        Flog.writeline "Inicio Empleado:" & Ternro
   
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
                bprcfechasta = bprcfechastaAux
                
                Flog.writeline "Busco el periodo para Costa Rica."
                fechaAlta = FechaAltaEmpleado(Ternro)
                
                'EAM- Si el empleado no tiene fase, pasa al siguiente empleado
                If (fechaAlta = Empty) Then
                    GoTo siguiente
                End If
                                
                 Flog.writeline "Fecha Procesamiento: " & bprcfechasta & "."
                'EAM- Obtiene el año que se va a procesar a partir de la fase del empleado.
                If bprcfechasta < CDate(Day(fechaAlta) & "/" & Month(fechaAlta) & "/" & Year(bprcfechasta)) Then
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
                    Flog.writeline "Genero el periodo para Costa Rica."
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
                Call bus_DiasVac(Ternro, auxNroVac, cantdias, Columna, Mensaje, Genera, cantdiasCorr, Columna2)
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
                Call bus_DiasVac_CR(Ternro, auxNroVac, fechaAlta, bprcfechasta, CantdiasCR, Columna, Mensaje, Genera, Periodo_Anio)
                DiasCorraGen = CantdiasCR
            Case 5: 'Portugal
                Flog.writeline "Modelo de vacaciones de Portugal Nro." & modeloPais
                Call bus_DiasVac_PT(Ternro, auxNroVac, cantdias, Columna, Mensaje, Genera, cantdiasCorr, Columna2)
                DiasCorraGen = cantdias
        End Select
       
       'FGZ - 25/03/2010 --------------------------
        
        ''Flog.writeline ""
        'Flog.writeline "genera: " & Genera
        'Flog.writeline ""
        If Not Genera Then GoTo siguiente
        
        If st_TipoDia2 = 0 Then
            NroTPVCorr = Columna2
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
            NroTPVCorr = Columna2
        End If
        
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
        Flog.writeline "VERSION DE POLITICA 1514: " & st_Opcion
        
        If PoliticaOK Then
            Select Case st_Opcion
                Case 0: 'Costa Rica
                    Flog.writeline "Politica 1514 Activa. Cálculo de Beneficio Adicional de vacaciones. COSTA RICA"
                    Call CalcularBeneficioVac(Ternro, auxNroVac, NroTPV, Reproceso, bprcfechasta)
                Case 1: 'Portugal
                    If cantdias = 22 Then
                        Flog.writeline "Politica 1514 Activa. Cálculo de PLUS Adicional de vacaciones. PORTUGAL"
                        Call CalcularBeneficioVac_PT(Ternro, auxNroVac, NroTPV, fecha_desde, fecha_hasta, Lic_Descuento, Reproceso, bprcfechasta)
                    End If
            End Select
        End If
        
        MyCommitTrans
' ----------------------------------------------------------
siguiente:
        MyCommitTrans
          
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

Public Sub bus_DiasVac(ByVal Ternro As Long, ByVal NroVac As Long, ByRef cantdias As Integer, ByRef Columna As Integer, ByRef Mensaje As String, ByRef Genera As Boolean _
    , ByRef cantdiasCorr As Integer, ByRef Columna2 As Integer)
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
Dim Continuar As Boolean
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
    
    Continuar = True
    ant = 1
    Do While (ant <= rs_cabgrilla!cgrdimension) And Continuar
        Select Case ant
        Case 1:
            StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_1
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    Continuar = False
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
                    Continuar = False
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
                    Continuar = False
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
                    Continuar = False
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
                    Continuar = False
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
    Columna2 = TipoVacacionProporcionCorr

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
        
        dias_trabajados = ((antanio * 365) + (antmes * 30) + antdia)
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
                        cantdias = Fix((dias_trabajados / DiasProporcion) / 7 * habiles)
                    Else
                        cantdias = Fix(20 * (dias_trabajados / DiasProporcion) / FactorDivision)
                End If
                
                aux_redondeo = ((dias_trabajados / DiasProporcion) / 7 * habiles) - Fix((dias_trabajados / DiasProporcion) / 7 * habiles)
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
Dim Continuar As Boolean
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
    
    Continuar = True
    ant = 1
    Do While (ant <= rs_cabgrilla!cgrdimension) And Continuar
        Select Case ant
        Case 1:
            StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_1
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    Continuar = False
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
                    Continuar = False
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
                    Continuar = False
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
                    Continuar = False
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
                    Continuar = False
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
                        
   
Genera = True
    
' Cierro todo y libero
If rs_cabgrilla.State = adStateOpen Then rs_cabgrilla.Close
If rs_valgrilla.State = adStateOpen Then rs_valgrilla.Close

Set rs_cabgrilla = Nothing
Set rs_valgrilla = Nothing
End Sub


Public Sub bus_DiasVac_CR(ByVal Ternro As Long, ByRef NroVac As Long, ByVal fechaAlta As Date, ByRef FechaHasta As Date, ByRef cantdias As Double, ByRef Columna As Integer, ByRef Mensaje As String, ByRef Genera As Boolean, ByVal Anio As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Calcula los días de vacaciones
' Autor      : Margiotta, Emanuel
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim totalDias As Long
Dim fechaDesde As Date
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
Dim SemEvaluadas AS Double
Dim DiasHorarioUltSem AS Double

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
                        fechaDesde = DateAdd("d", 1, ultFechaProcesada)
                        
                        'EAM- Obtiene la fecha de fin de semana
                        If IsDate(fechaDesde) Then
                            fecHastaSemana = DateAdd("d", 7 - (Weekday(fechaDesde) - 1), fechaDesde)
                            fecDesdeSemana = DateAdd("d", -6, fecHastaSemana)
                        End If
    

                        'Busco el ultimo movimiento en vigencia - Dias habiles ultima semana. No se puede hacer abajo porque se puede afectar con la fecha de corte
                        StrSql = " SELECT DISTINCT fechor FROM WC_MOV_HORARIOS WHERE ternro = " & Ternro & _
                                " AND fechor  <= " & ConvFecha(fecHastaSemana) & " AND fechor >= " & ConvFecha(fechaDesde)
                        OpenRecordset StrSql, rsRegHorario
			DiasHorarioUltSem = 5
			If Not rsRegHorario.EOF Then
				DiasHorarioUltSem = rsRegHorario.RecordCount
			End If
			rsRegHorario.Close
        
                        If CDate(fecHastaSemana) < CDate(fechaCorte) Then
                            FechaHasta = fecHastaSemana
                        Else
                            FechaHasta = fechaCorte
                            fecHastaSemana = fechaCorte
                        End If
                        
                        
                        'Busco el ultimo movimiento en vigencia
                        StrSql = " SELECT DISTINCT fechor FROM WC_MOV_HORARIOS WHERE ternro = " & Ternro & _
                                " AND fechor  <= " & ConvFecha(fecHastaSemana) & " AND fechor >= " & ConvFecha(fechaDesde)
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
                        
                
                        totalDias = totalDias + (DateDiff("d", fechaDesde, FechaHasta) + 1)
                        ultFechaProcesada = FechaHasta
                        Genera = True
                    Loop
                Else
                    Flog.writeline "No se encontro asignación horaria para la fecha: " & FechaHasta
                    FechaHasta = ultFechaProcesada
                    rsDiasCorresp.Close
                    GoTo sinDatos
                End If
            
		if DiasHorario < DiasHorarioUltSem Then
			DiasHorario = DiasHorarioUltSem
		Else
			semEvaluadas = totalDias / 7
			DiasHorario  = DiasHorario / semEvaluadas
			if DiasHorario > DiasHorarioUltSem Then
				DiasHorario = DiasHorarioUltSem
			End If
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

Public Function FechaAltaEmpleado(Ternro) As Date
    Dim StrSql As String
    Dim rsFases As New ADODB.Recordset
    Dim i_dia As Integer
    Dim i_mes As Integer
    Dim i_anio As Integer
    
    StrSql = "SELECT * FROM fases where fases.empleado = " & Ternro & " AND fases.fasrecofec = -1 "
    OpenRecordset StrSql, rsFases
    If rsFases.EOF Then
        FechaAltaEmpleado = Empty
        Flog.writeline "El empleado no tiene fase con alta reconocida."
    Else
        If IsNull(rsFases("altfec")) Then
            FechaAltaEmpleado = ""
        Else
            'i_dia = Day(CDate(rsFases("altfec")))
            'i_mes = Month(CDate(rsFases("altfec")))
            'i_anio = Year(CDate(rsFases("altfec")))
            FechaAltaEmpleado = CDate(rsFases("altfec"))
        End If
    End If
    
    If rsFases.State = adStateOpen Then rsFases.Close
    Set rsFases = Nothing
    
End Function

Sub generarPeriodoVacacion(Ternro As Long, Anio As Integer, Optional modeloPais As Integer)
    Dim rs_vacacion As New ADODB.Recordset
    Dim fechaAlta As Date
    Dim FechaInicio As Date
    Dim fechaFin As Date
    Dim vacdesc As String
    Dim vacestado As Integer

    fechaAlta = FechaAltaEmpleado(Ternro)
    If Anio < Year(fechaAlta) Then
        Flog.writeline "Error al querer generar un periodo anterior a la fecha de alta del empleado."
        Exit Sub
    End If
    FechaInicio = formatFecha(CStr(Day(fechaAlta)), CStr(Month(fechaAlta)), CStr(Anio))
    fechaFin = DateAdd("d", -1, DateAdd("yyyy", 1, FechaInicio))
    vacdesc = CStr(Anio) & " - " & CStr(Anio + 1)
    
    If modeloPais = 3 Then
        If Anio < Year(Date) - 4 Then
            vacestado = 0
        Else
            vacestado = -1
        End If
    End If
    StrSql = " INSERT INTO vacacion (vacdesc, vacfecdesde, vacfechasta, vacanio, empnro, ternro, vacestado) "
    StrSql = StrSql & " VALUES ('" & vacdesc & "'," & ConvFecha(FechaInicio) & "," & ConvFecha(fechaFin) & "," & Anio & ",0," & Ternro & "," & vacestado & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
    
    
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
Dim Continuar As Boolean
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
    
    Continuar = True
    ant = 1
    Do While (ant <= rs_cabgrilla!cgrdimension) And Continuar
        Select Case ant
        Case 1:
            StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_1
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    Continuar = False
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
                    Continuar = False
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
                    Continuar = False
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
                    Continuar = False
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
                    Continuar = False
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
Public Function DiasHabilTrabajado(ByVal Ternro As Long, ByVal fDesde As Date, ByVal fhasta As Date) As Double
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
             "WHERE tipoestructura.tenro = 21 AND his_estructura.htetdesde<= " & ConvFecha(fhasta) & _
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
                fecHasta = fhasta
            Else
                If CDate(rsRegHorario!htethasta) >= CDate(fhasta) Then
                    fecHasta = fhasta
                Else
                    fecHasta = rsRegHorario!htethasta
                End If
            End If
            
            'cantDias = CInt(DateDiff("d", rsRegHorario!htetdesde, fhasta))
            cantdias = CLng(DateDiff("d", fecDesde, fecHasta) + 1)
            cantdias = cantdias - LicenciaGozadas(Ternro, fDesde, fhasta)
            
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

Public Function LicenciaGozadas(ByVal Ternro As Long, ByVal fDesde As Date, ByVal fhasta As Date) As Integer
  Dim rsLicencias As New ADODB.Recordset
  Dim fecDesde As Date
  Dim fecHasta As Date
  Dim cantLicencias As Double
  
    cantLicencias = 0
  
    StrSql = "SELECT * FROM emp_lic " & _
             "WHERE empleado= " & Ternro & " AND emp_lic.elfechadesde<= " & ConvFecha(fhasta) & _
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
            fecHasta = fhasta
        Else
            If CDate(rsLicencias!elfechahasta) >= CDate(fhasta) Then
                fecHasta = fhasta
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
    , ByRef cantdiasCorr As Integer, ByRef Columna2 As Integer)
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
Dim Continuar As Boolean
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
    
    Continuar = True
    ant = 1
    Do While (ant <= rs_cabgrilla!cgrdimension) And Continuar
        Select Case ant
        Case 1:
            StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_1
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    Continuar = False
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
                    Continuar = False
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
                    Continuar = False
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
                    Continuar = False
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
                    Continuar = False
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
    Columna2 = TipoVacacionProporcionCorr




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
'EAM- Obtiene la cantidad de días feriados cargados en el sistema para un rango de fecha
Function cantDiasFeriados(ByVal fDesde As Date, ByVal fhasta As Date) As Long
 Dim rsFeriados As New ADODB.Recordset
 Dim objFeriado As New Feriado
 Dim cantFeriado
 
 
    cantFeriado = 0
 
    'Busco todos los Feriados
    StrSql = "SELECT * FROM feriado WHERE ferifecha >= " & ConvFecha(fDesde) & " AND ferifecha < " & ConvFecha(fhasta)
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
Dim Continuar As Boolean
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
    
    Continuar = True
    ant = 1
    Do While (ant <= rs_cabgrilla!cgrdimension) And Continuar
        Select Case ant
        Case 1:
            StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_1
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    Continuar = False
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
                    Continuar = False
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
                    Continuar = False
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
                    Continuar = False
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
                    Continuar = False
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
Sub CalcularBeneficioVac(ByVal Ternro As Long, ByVal NroVac As Long, ByVal TipoVac As Integer, ByVal Reproceso As Boolean, ByVal vdiasfechasta As Date)
Dim fechaAlta As Date
Dim cantMeses As Long
Dim diasBeneficio As Integer
Dim l_rsAux As New ADODB.Recordset

    'EAM- Obtiene la fecha de Ingreso
    fechaAlta = FechaAltaEmpleado(Ternro)
    
    
    If fechaAlta < CDate("01/01/2007") Then
        
        'Calcula la antiguedad en Años
        cantMeses = DateDiff("m", fechaAlta, Date)
                
        'Obtiene los dias de Beneficio
        Select Case cantMeses
            Case Is < 24:
                diasBeneficio = 0
            Case Is <= 24:
                diasBeneficio = 1
            Case Is <= 36:
                diasBeneficio = 2
            Case Is <= 48:
                diasBeneficio = 3
            Case Is <= 60:
                diasBeneficio = 4
            Case Is <= 72:
                diasBeneficio = 5
            Case Is <= 84:
                diasBeneficio = 6
            Case Is <= 96:
                diasBeneficio = 7
            Case Is <= 108:
                diasBeneficio = 8
            Case Is <= 120:
                diasBeneficio = 9
            Case Else
                diasBeneficio = 10
        End Select
        
        StrSql = "SELECT * FROM vacdiascor WHERE vacnro= " & NroVac & " AND ternro= " & Ternro & " AND venc=3"
        OpenRecordset StrSql, l_rsAux
        
        If l_rsAux.EOF Then
            StrSql = "INSERT INTO vacdiascor (vacnro,ternro,vdiascorcant,vdiapednro,vdiascormanual,tipvacnro,venc,vdiasfechasta) " & _
                    "VALUES(" & NroVac & "," & Ternro & "," & diasBeneficio & ",0,0," & TipoVac & ",3," & ConvFecha(vdiasfechasta) & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
            Flog.writeline "Dias de Beneficio: " & diasBeneficio & " a la fecha " & Date
        Else
            Flog.writeline "Los días de Beneficio ya estan calculados."
            'If Reproceso Then
'                StrSql = "UPDATE vacdiascor SET vdiascorcant= " & diasBeneficio & ",vdiasfechasta=" & ConvFecha(vdiasfechasta) & _
'                        " WHERE ternro= " & Ternro & " AND vacnro = " & NroVac & " AND venc=3"
'                objConn.Execute StrSql, , adExecuteNoRecords
'            End If
        End If
    Else
        
    End If
End Sub

Sub CalcularBeneficioVac_PT(ByVal Ternro As Long, ByVal NroVac As Long, ByVal TipoVac As Integer, ByVal fecha_desde As Date, fecha_hasta As Date, ByVal Lic_Descuento As String, ByVal Reproceso As Boolean, ByVal vdiasfechasta As Date)
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
         "VALUES(" & NroVac & "," & Ternro & "," & diasBeneficio & ",0,0," & TipoVac & ",3," & ConvFecha(vdiasfechasta) & ")"
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

'EAM- busca en la escala de vacaciones los dias correspondientes segun el tipo de vacaciones pasado por parametro
Private Function buscarDiasVacEscala(ByVal ant As Integer, ByVal dimensionEscala, ByVal parametros, ByVal tipovacacion As Long, Optional ByRef Encontro As Boolean) As Integer
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
            If parametros(ant) >= rs_valgrilla!vgrcoor_2 Then
                 If rs_valgrilla!vgrvalor <> 0 Then
                    cantdias = rs_valgrilla!vgrvalor
                    Encontro = True
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
