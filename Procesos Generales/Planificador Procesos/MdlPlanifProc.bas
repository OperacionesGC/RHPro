Attribute VB_Name = "MdlPlanifProc"
Option Explicit

Global IdUser As String
Global Fecha As Date
Global Hora As String

'Global Const Version = "1.00"
'Global Const FechaModificacion = "28/08/2007"
'Global Const UltimaModificacion = " " 'Martin Ferraro - Version Inicial

'Global Const Version = "1.01"
'Global Const FechaModificacion = "26/11/2008"
'Global Const UltimaModificacion = " " 'Martin Ferraro - Se agrego que inserte el parametro al insertar en batch_proceso


'Global Const Version = "1.02"
'Global Const FechaModificacion = "13/02/2009"
'Global Const UltimaModificacion = " FGZ " 'FGZ - Encriptacion de string de conexion y SCHEMA de ORACLE
'                                   Modificaciones generales de manejo de error

'Global Const Version = "1.03"9
'Global Const FechaModificacion = "16/09/2009"
'Global Const UltimaModificacion = " MB " ' Modificaciones al LOG

'Global Const Version = "1.04"
'Global Const FechaModificacion = "01/12/2009"
'Global Const UltimaModificacion = " EAM " ' Se le agregó logica para poder planificar procesos de HC y AD

'Global Const Version = "1.05"
'Global Const FechaModificacion = "08/09/2010"
'Global Const UltimaModificacion = " FGZ " 'Manejo particular para procesos planificados de interface general

'Global Const Version = "1.06"
'Global Const FechaModificacion = "21/09/2010"
'Global Const UltimaModificacion = " FGZ " 'Se cambió la convencion de nombre que deben tener los archivos de interfaces a procesar
''                                            detalles de los archivos
''                                            Los mismos deben respetar la forma : encabezado_nnnnaaaammddhhmm.csv donde:
''                                               encabezado son 4 caracteres alfanumericos seguidos de un _
''                                               nnnn: 'Numero de interface - interface number (ie. 605)
''                                               aaaa: Año de creacion de archivo - file creation year
''                                               mm: Mes de creacion de archivo - file creation month
''                                               dd: Dia de creacion de archivo - file creation date
''                                               hh: Hora de creacion de archivo - file creation hour
''                                               mm: Minutos de creacion de archivo - file creation minutes

'Global Const Version = "1.07"
'Global Const FechaModificacion = "23/09/2010"
'Global Const UltimaModificacion = " FGZ " 'Se cambió la convencion de nombre que deben tener los archivos de interfaces a procesar
''                                            detalles de los archivos
''                                            Los mismos deben respetar la forma : nnnnaaaammddhhmm.csv donde:
''                                               nnnn: 'Numero de interface - interface number (ie. 605)
''                                               aaaa: Año de creacion de archivo - file creation year
''                                               mm: Mes de creacion de archivo - file creation month
''                                               dd: Dia de creacion de archivo - file creation date
''                                               hh: Hora de creacion de archivo - file creation hour
''                                               mm: Minutos de creacion de archivo - file creation minutes
''Ejemplo ttec_0605201009211450.csv

'Global Const Version = "1.08"
'Global Const FechaModificacion = "03/11/2010"
'Global Const UltimaModificacion = " FGZ " 'Se cambió la convencion de nombre que deben tener los archivos de interfaces a procesar
'                                            detalles de los archivos
'                                            Los mismos deben respetar la forma : aaaammddhhmmss_nnnn.csv donde:
'                                               aaaa: Año de creacion de archivo - file creation year
'                                               mm: Mes de creacion de archivo - file creation month
'                                               dd: Dia de creacion de archivo - file creation date
'                                               hh: Hora de creacion de archivo - file creation hour
'                                               mm: Minutos de creacion de archivo - file creation minutes
'                                               Ss: Secuencia. Esta secuencia se determina en base al orden de ejecución
'                                                       de interfaces que nos solicito RH pro.
'                                                       Ejemplo, las interfaz 605 =00, la interfaz 630 =05, etc.
'                                               nnnn: Numero de interface - interface number (ie. 605)
'20101028123100_0605.txt
'20101028123105_0630.txt


'Global Const Version = "1.09"
'Global Const FechaModificacion = "17/02/2011"
'Global Const UltimaModificacion = " Leticia A." ' Guardar información del las Interfaces que generó el Planificador - (Price)

'Global Const Version = "1.10"
'Global Const FechaModificacion = "13/04/2011"
'Global Const UltimaModificacion = " FGZ" ' Cuando se planifica interfaces y no se sigue la linea automatica, es decir, se guarda como la planificacion estandar
''                                           cuando inserta en batch_proceso le pasa como parametros los parametros cargados en la configuracion del proceso.

'Global Const Version = "1.11"
'Global Const FechaModificacion = "09/06/2011"
'Global Const UltimaModificacion = " FGZ" ' Cambio de definicion de variable de batch_proceso

'Global Const Version = "1.12"
'Global Const FechaModificacion = "08/05/2012"
'Global Const UltimaModificacion = " FGZ" ' la 1er linea del mail (On Error GoTo ME_Gral) activa un manejador de error
'           que de activarse antes de crear el archivo de log va a dar un error. Se comentó la linea.

'Global Const Version = "1.13"
'Global Const FechaModificacion = "22/01/2015"
'Global Const UltimaModificacion = " Dimatz Rafael" ' Se creo el Modelo 422 para la Interfaz F572 Nro 292
'           Mueve los archivos del Modelo 1 del F572 a la carpeta correspondiente \Form572 para que luego el planificador ejecute la Interfaz 292

'Global Const Version = "1.14"
'Global Const FechaModificacion = "27/01/2015"
'Global Const UltimaModificacion = " Dimatz Rafael" ' CAS - 29124. Se Modifico el modelo 422 para que planifique una sola vez. Verifica que si LeerArchivo tiene XML muestre que no es el Formato Correcto

'Global Const Version = "1.15"
'Global Const FechaModificacion = "02/06/2015"
'Global Const UltimaModificacion = " Dimatz Rafael" ' CAS - 31108. Descomprime el Zip en Form572, copie el Zip de la carpeta del Modelo 1 en Form572 y elimine el Zip

'Global Const Version = "1.16"
'Global Const FechaModificacion = "04/06/2015"
'Global Const UltimaModificacion = " Dimatz Rafael" ' CAS - 31108. Se descomprime desde la carpeta de cgi-bin\servicios\zip

'Global Const Version = "1.17"
'Global Const FechaModificacion = "09/06/2015"
'Global Const UltimaModificacion = " Dimatz Rafael" ' CAS - 31108. Se creo un shell para que descomprima los XML y elimine los Zip

'Global Const Version = "1.18"
'Global Const FechaModificacion = "25/06/2015"
'Global Const UltimaModificacion = " Dimatz Rafael" ' CAS - 31714. Se corrigio para que inserte correctamente en la tabla Batch_Proceso

'Global Const Version = "1.19"
'Global Const FechaModificacion = "28/07/2015"
'Global Const UltimaModificacion = " Dimatz Rafael" ' CAS - 32248. Se corrigio para que borre el zip

'Global Const Version = "1.20"
'Global Const FechaModificacion = "02/09/2015"
'Global Const UltimaModificacion = " Fernandez, Matias" ' CAS-32717 - Prudential - Bug Descompresión ZIP SIRADIG
                                                       ' correccion en el armado de rutas para descomprimir

Global Const Version = "1.21"
Global Const FechaModificacion = "07/09/2015"
Global Const UltimaModificacion = " Fernandez, Matias " '- CAS-32773 - MERCADO CENTRAL - Bug en planificador del 572 "
                                                        ' replanificacion, y elimina todos los zip primero y despues descomprime


'--------MDF
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

'--------MDF
   
   
Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial del Proceso.
' Autor      : Martin Ferraro
' Fecha      : 31/07/2007
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim strCmdLine
Dim Nombre_Arch As String
Dim HuboError As Boolean
Dim rs_batch_proceso As New ADODB.Recordset
Dim PID As String
Dim bprcparam As String
Dim ArrParametros

    'On Error GoTo ME_Gral
    
    strCmdLine = Command()
    ArrParametros = Split(strCmdLine, " ", -1)
    
    If UBound(ArrParametros) > 0 Then
        Etiqueta = ArrParametros(0)
        EncriptStrconexion = CBool(ArrParametros(1))
        c_seed = ArrParametros(1)
    Else
        Etiqueta = ArrParametros(0)
    End If
    
    
    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas

    'Creacion del archivo de log
    Nombre_Arch = PathFLog & "planificador " & Format(Date, "dd-mm-yyyy") & ".log"
    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.fileexists(Nombre_Arch) Then
        ' lo abro para agregar
        Set Flog = fs.OpenTextFile(Nombre_Arch, 8, 0)
    Else
        ' no existe, entonces lo creo
        Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    End If
        
        
    'Abro la conexion
    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
        
        
    Flog.writeline
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Version = " & Version
    Flog.writeline "Modificacion = " & UltimaModificacion
    Flog.writeline "Fecha = " & FechaModificacion
    Flog.writeline "Ult. Ejec. = " & Now
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline
    
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Planificacion de Procesos"
    Flog.writeline "-----------------------------------------------------------------"
    
    On Error GoTo ME_Main
    
    TiempoInicialProceso = GetTickCount
    
    Call BuscarPlanificados
    
Final:
    Flog.writeline Espacios(Tabulador * 0) & "Fin :" & Now
    TiempoFinalProceso = GetTickCount
    Flog.Close
    objConn.Close
Exit Sub

ME_Main:
    HuboError = True
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 0) & " Error " & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline
    GoTo Final
    
ME_Gral:
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
        Flog.writeline Espacios(Tabulador * 0) & " Error General " & Err.Description
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "SQL Ejecutado: " & StrSql
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
        Flog.writeline
        GoTo Final
End Sub

Public Sub KillProcess(ByVal processName As String)
On Error GoTo ErrHandler
Dim oWMI
Dim ret
Dim sService
Dim oWMIServices
Dim oWMIService
Dim oServices
Dim oService
Dim servicename
Set oWMI = GetObject("winmgmts:")
Set oServices = oWMI.InstancesOf("win32_process")
For Each oService In oServices

servicename = LCase(Trim(CStr(oService.Name) & ""))

If InStr(1, servicename, LCase(processName), vbTextCompare) > 0 Then
ret = oService.Terminate
End If

Next

Set oServices = Nothing
Set oWMI = Nothing

ErrHandler:
Err.Clear
End Sub

Public Sub BuscarPlanificados()
' -----------------------------------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que se encarga de buscar procesos planificados, insertarlos en batch
'              y calcular su proxima ejecucion de acuerdo a lo configurado
' Autor      : Martin Ferraro
' Fecha      : 21/06/2007
' Ult. Mod   :
' -----------------------------------------------------------------------------------------------------------------------
Dim rs_Consult As New ADODB.Recordset
Dim rs_Consult2 As New ADODB.Recordset

Dim fechaActual As Date
Dim horaActual As String
Dim fecProxEjec As Date
Dim horProxEjec As String
Dim planifCorrecta As Boolean

Dim Path
Dim NArchivo
Dim Directorio As String
Dim DirectorioDestino As String
Dim OK As Boolean
Dim CArchivos
Dim Archivo
Dim Nombre_Archivo
Dim folder
Dim modInterface As Long
Dim modParametros As String
Dim fs, f
Dim Carpeta

Dim Parametros() As String
Dim FechaDesde As Date
Dim FechaHasta As Date
Dim Id_batproc As Long

Dim Encontro_Arch As Boolean
Dim Archivo_XML As String

Dim Modelo
Dim l_path

fechaActual = Date
horaActual = Time

On Error GoTo E_BuscarPlanificados

Set fs = CreateObject("Scripting.FileSystemObject")
Set f = CreateObject("Scripting.FileSystemObject")

'-------------------------------------------------------------------------------------------------
'Barro la tabla de proceso planificados en busca de procesos listos a planificar
'-------------------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 0) & "Buscando Procesos a ejecutar el " & fechaActual & " " & horaActual
StrSql = "SELECT  batch_procplan.bproplanro, batch_procplan.btprcnro, batch_procplan.schednro, batch_procplan.param, batch_procplan.iduser, batch_procplan.fecproxejec, batch_procplan.fecultejec,"
StrSql = StrSql & " batch_procplan.horaproxejec, batch_procplan.horaultejec, batch_procplan.bpropladesabr, ale_sched.frectipnro"
StrSql = StrSql & " FROM batch_procplan "
StrSql = StrSql & " INNER JOIN ale_sched ON batch_procplan.schednro = ale_sched.schednro "
StrSql = StrSql & " WHERE ((fecproxejec < " & ConvFecha(fechaActual) & ")"
StrSql = StrSql & " OR (fecproxejec = " & ConvFecha(fechaActual) & " AND horaproxejec <= '" & Format(horaActual, "HH:mm") & "'))"
OpenRecordset StrSql, rs_Consult

If rs_Consult.EOF Then
    Flog.writeline Espacios(Tabulador * 0) & "No hay procesos planificados pendientes.."
End If

Do While Not rs_Consult.EOF

    Flog.writeline Espacios(Tabulador * 1) & "Proceso Planificado " & rs_Consult!bproplanro & " - " & rs_Consult!bpropladesabr & " ProxEjec: " & rs_Consult!fecProxEjec & " " & rs_Consult!horaproxejec & " " & " UltEjec: " & rs_Consult!fecultejec & " " & rs_Consult!horaultejec
    
    'Calculo la proxima planificacion
    Call CalcularProxEjec(rs_Consult!schednro, fechaActual, horaActual, fecProxEjec, horProxEjec, planifCorrecta)
    
    If planifCorrecta Then
    
        Select Case rs_Consult("btprcnro")
            Case 1, 2:
                'Dim Parametros() As String
                'Dim FechaDesde As Date
                'Dim FechaHasta As Date
                'Dim Id_batproc As Long
                
                'Param -> Rango@legajo desde@legajo hasta@ estado@TE1@Est1@TE2@Est2@TE3@Est3@Análisis detallado@Cant Proc en paralelo
                Parametros = Split(rs_Consult("Param"), "@")
                
                'Calcula las fechas del Periodo de procesamiento
                Call CalcFechaProcesamiento(Parametros(0), rs_Consult("frectipnro"), FechaDesde, FechaHasta)
                                
                'Inserto en batch_proceso para la ejecucion del mismo
                StrSql = "INSERT INTO Batch_Proceso (btprcnro, bprcfecha, iduser, bprchora, bprcfecdesde, bprcfechasta,"
                StrSql = StrSql & " bprcestado, empnro, bprcConfirmado, bprcparam)"
                StrSql = StrSql & " values (" & rs_Consult!btprcnro & "," & ConvFecha(fechaActual) & ", '" & rs_Consult!IdUser & "'" & ",'" & Format(horaActual, "hh:mm:ss ") & "'"
                StrSql = StrSql & " , " & ConvFecha(FechaDesde) & ", " & ConvFecha(FechaHasta)
                StrSql = StrSql & " , 'Pendiente', 0, -1,'" & Parametros(10) & "')"
                
                Flog.writeline Espacios(Tabulador * 1) & "Inserto Proceso: " & StrSql
                objConn.Execute StrSql, , adExecuteNoRecords
                
                Id_batproc = getLastIdentity(objConn, "batch_Proceso")
                    
                If Id_batproc = -1 Then
                    Flog.writeline Espacios(Tabulador * 1) & "Error al Obtener getLastIdentity del Proceso: "
                Else
                    If Parametros(8) <> "" And Parametros(8) <> "0" Then
                        StrSql = "SELECT DISTINCT empleado.ternro " & _
                        "FROM empleado INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro AND estact1.htethasta IS NULL AND estact1.tenro  = " & Parametros(4)
                        If Parametros(4) <> "" And Parametros(4) <> "0" Then
                            StrSql = StrSql & " AND estact1.estrnro =" & Parametros(5)
                        End If
                        StrSql = StrSql & " INNER JOIN his_estructura estact2 ON empleado.ternro = estact2.ternro AND estact2.htethasta IS NULL AND estact2.tenro  = " & Parametros(6)
                        If Parametros(6) <> "" And Parametros(6) <> "0" Then
                            StrSql = StrSql & " AND estact2.estrnro =" & Parametros(7)
                        End If
                        StrSql = StrSql & " INNER JOIN his_estructura estact3 ON empleado.ternro = estact3.ternro AND estact3.htethasta IS NULL AND estact3.tenro  = " & Parametros(8)
                        If Parametros(10) <> "" And Parametros(10) <> "0" Then
                            StrSql = StrSql & " AND estact3.estrnro =" & Parametros(10)
                        End If
                        StrSql = StrSql & " WHERE (empest = " & Parametros(3) & ") And (empleado.empleg >= " & Parametros(1) & ") And (empleado.empleg <= " & Parametros(2) & ")"
                        
                        ElseIf Parametros(6) <> "" And Parametros(6) <> "0" Then
                            StrSql = "SELECT DISTINCT empleado.ternro " & _
                            "FROM empleado INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro AND estact1.htethasta IS NULL AND estact1.tenro  = " & Parametros(4)
                            If Parametros(5) <> "" And Parametros(5) <> "0" Then
                                StrSql = StrSql & " AND estact1.estrnro =" & Parametros(5)
                            End If
                            StrSql = StrSql & " INNER JOIN his_estructura estact2 ON empleado.ternro = estact2.ternro AND estact2.htethasta IS NULL AND estact2.tenro  = " & Parametros(6)
                            If Parametros(7) <> "" And Parametros(7) <> "0" Then
                                StrSql = StrSql & " AND estact2.estrnro =" & Parametros(7)
                            End If
                            StrSql = StrSql & " WHERE (empest = " & Parametros(3) & ") And (empleado.empleg >= " & Parametros(1) & ") And (empleado.empleg <= " & Parametros(2) & ")"
                            
                            ElseIf Parametros(4) <> "" And Parametros(4) <> "0" Then
                                StrSql = "SELECT DISTINCT empleado.ternro " & _
                                " FROM empleado INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro AND estact1.htethasta IS NULL AND estact1.tenro  = " & Parametros(4)
                                If Parametros(5) <> "" And Parametros(5) <> "0" Then
                                    StrSql = StrSql & " AND estact1.estrnro =" & Parametros(5)
                                End If
                                StrSql = StrSql & " WHERE (empest = " & Parametros(3) & ") And (empleado.empleg >= " & Parametros(1) & ") And (empleado.empleg <= " & Parametros(2) & ")"
    
                            Else
                                StrSql = "SELECT DISTINCT empleado.ternro " & _
                                "FROM empleado "
                                StrSql = StrSql & " WHERE (empest = " & Parametros(3) & ") And (empleado.empleg >= " & Parametros(1) & ") And (empleado.empleg <= " & Parametros(2) & ")"
                    End If
                    'Ejecuta la cunsulta Filtro
                    OpenRecordset StrSql, rs_Consult2
                    
                    While Not rs_Consult2.EOF
                        StrSql = "INSERT INTO Batch_empleado (bpronro, ternro)"
                        StrSql = StrSql & " values (" & Id_batproc & "," & rs_Consult2("ternro") & ")"
                                        
                        Flog.writeline Espacios(Tabulador * 1) & "Inserto el empleado del Proceso: " & StrSql
                        objConn.Execute StrSql, , adExecuteNoRecords
                        
                        rs_Consult2.MoveNext
                    Wend
                
                    'Actualizo la fecha de ultima corrida y proxima ejecucion
                    StrSql = "UPDATE batch_procplan "
                    StrSql = StrSql & "SET fecultejec = " & ConvFecha(fechaActual)
                    StrSql = StrSql & ",horaultejec = '" & Format(horaActual, "HH:mm") & "'"
                    StrSql = StrSql & ",fecproxejec = " & ConvFecha(fecProxEjec)
                    StrSql = StrSql & ",horaproxejec = '" & Format(horProxEjec, "HH:mm") & "'"
                    StrSql = StrSql & " WHERE bproplanro = " & rs_Consult!bproplanro
                    Flog.writeline Espacios(Tabulador * 1) & "Actualizo Planif : " & StrSql
                    objConn.Execute StrSql, , adExecuteNoRecords
                
                End If
                
            Case 23:    'FGZ - 07/09/2010
                'Param -> Rango@legajo desde@legajo hasta@ estado@TE1@Est1@TE2@Est2@TE3@Est3@Análisis detallado@Cant Proc en paralelo
                Parametros = Split(rs_Consult("Param"), "@")
                
                'Modelo
                If Parametros(0) = 1 Then
                    'Buscar la carpeta del modelo 1
                    Call BuscarModelo(Parametros(0), Directorio, OK)
                    If Not OK Then
                        Flog.writeline Espacios(Tabulador * 1) & "No se puede Automatizar el proceso de interfaces"
                        Exit Sub
                    End If
                    
                    'Busco los archivos en la carpeta
                    'Set fs = CreateObject("Scripting.FileSystemObject")
                    
                    Path = Directorio
                    Set folder = fs.GetFolder(Directorio)
                    Set CArchivos = folder.Files
                    For Each Archivo In CArchivos
                        NArchivo = Archivo.Name
                        Flog.writeline Espacios(Tabulador * 1) & "Procesando archivo " & Archivo.Name
                        
                        ' por cada archivo
                        '   determino de que modelo se trata y luego
                        '   inserto el proceso en batch proceso con los parametros correspondientes (estado preparando)
                        '   lo subo a la carpeta que corresponda y con el nombre que corresponda
                        '   Actualizo el estado del proceso
                        
                        
                        Call LeeArchivo(Archivo.Name, modInterface, modParametros)
                        
                        Call BuscarModelo(modInterface, DirectorioDestino, OK)
                        'If Not OK Then
                        '    Flog.writeline Espacios(Tabulador * 1) & "No se puede encontrar el la carpeta correspondiente al modelo destino " & modInterface
                        '    Exit Sub
                        'End If
                        If OK Then
                            'Inserto en batch_proceso para la ejecucion del mismo
                            StrSql = "INSERT INTO Batch_Proceso (btprcnro, bprcfecha, iduser, bprchora, bprcfecdesde, bprcfechasta,"
                            StrSql = StrSql & " bprcestado, empnro, bprcConfirmado, bprcparam)"
                            StrSql = StrSql & " values (" & rs_Consult!btprcnro & "," & ConvFecha(fechaActual) & ", '" & rs_Consult!IdUser & "'" & ",'" & Format(horaActual, "hh:mm:ss ") & "'"
                            StrSql = StrSql & " , " & ConvFecha(FechaDesde) & ", " & ConvFecha(FechaHasta)
                            StrSql = StrSql & " , 'Insertando', 0, -1,' ')"
                            objConn.Execute StrSql, , adExecuteNoRecords
                            
                            Id_batproc = getLastIdentity(objConn, "batch_Proceso")
                            
                                                    
                            ' Insertar información de las Interfaces que se ejecutan en dicha planificacion
                            StrSql = "INSERT INTO batch_interplan(bproplanro,bpronro,bpromodnro)"
                            StrSql = StrSql & " VALUES (" & rs_Consult!bproplanro & "," & Id_batproc & "," & modInterface & " )"
                            objConn.Execute StrSql, , adExecuteNoRecords
                            
                            
                            'Inserto las lineas a procesar del archivo
                            Call InsertatLineas(Path & "\" & Archivo.Name, Id_batproc)
                            
    
                            
                            'upload del archivo
                            Nombre_Archivo = "interface_" & modInterface & "_" & Id_batproc & ".csv"
                            
                            modParametros = modInterface & "@" & Nombre_Archivo & "@" & modParametros
                            
                            Set f = fs.getfile(Directorio & "\" & Archivo.Name)
                            
                            On Error Resume Next
                            f.Copy DirectorioDestino & "\" & Nombre_Archivo
                            
                            On Error GoTo E_BuscarPlanificados
                            StrSql = "UPDATE Batch_Proceso SET "
                            StrSql = StrSql & " bprcestado = 'Pendiente'"
                            StrSql = StrSql & " ,bprcparam ='" & modParametros & "'"
                            StrSql = StrSql & " WHERE bpronro = " & Id_batproc
                            Flog.writeline Espacios(Tabulador * 1) & "Inserto Proceso: " & StrSql
                            objConn.Execute StrSql, , adExecuteNoRecords
                                                                        
                        
                                                                        
                            On Error Resume Next
                            Err.Number = 0
                            f.Move Directorio & "\bk\" & Archivo.Name
                            If Err.Number <> 0 Then
                                If Err.Number = 58 Then
                                    'el archivo ya existe. ==> le agrego la hora.
                                    'f.Move Directorio & "\bk\" & Archivo.Name & "_" & Time()
                                    Err.Number = 0
                                    f.Move Directorio & "\bk\" & Mid(Archivo.Name, 1, Len(Archivo.Name) - 4) & "_" & Left(Format(Time(), "HH-mm-ss"), 8) & "." & Right(Archivo.Name, 3)
                                    If Err.Number <> 0 Then
                                        Flog.writeline Espacios(Tabulador * 0) & "No se puede hacer bk del archivo original."
                                    End If
                                    Err.Number = 0
                                Else
                                    Flog.writeline Espacios(Tabulador * 0) & "La carpeta Destino no existe. Se creará."
                                    Set Carpeta = fs.CreateFolder(Path & "\bk")
                                    f.Move Directorio & "\bk\" & Archivo.Name
                                End If
                            End If
                            Flog.writeline Espacios(Tabulador * 1) & "siguiente archivo "
                            Flog.writeline
                            
                            On Error GoTo E_BuscarPlanificados
                        Else
                            If modInterface = 0 Then
                                Flog.writeline Espacios(Tabulador * 1) & "Archivo invalido. Se descarta. " & Archivo.Name
                            Else
                                Flog.writeline Espacios(Tabulador * 1) & "No se puede encontrar el la carpeta correspondiente al modelo destino " & modInterface
                            End If
                        End If
                    Next
                Else
                    'Modelo estandar
                
                    'Inserto en batch_proceso para la ejecucion del mismo
                    StrSql = "INSERT INTO Batch_Proceso (btprcnro, bprcfecha, iduser, bprchora, bprcfecdesde, bprcfechasta,"
                    StrSql = StrSql & " bprcestado, empnro, bprcConfirmado, bprcparam)"
                    StrSql = StrSql & " values (" & rs_Consult!btprcnro & "," & ConvFecha(fechaActual) & ", '" & rs_Consult!IdUser & "'" & ",'" & Format(horaActual, "hh:mm:ss ") & "'"
                    StrSql = StrSql & " , " & ConvFecha(FechaDesde) & ", " & ConvFecha(FechaHasta)
                    If UBound(Parametros()) >= 10 Then
                        StrSql = StrSql & " , 'Pendiente', 0, -1,'" & Parametros(10) & "')"
                    Else
                        'StrSql = StrSql & " , 'Pendiente', 0, -1,'" & Parametros(0) & "')"
                        StrSql = StrSql & " , 'Pendiente', 0, -1,'" & rs_Consult("Param") & "')"
                    End If
                    
                    Flog.writeline Espacios(Tabulador * 1) & "Inserto Proceso: " & StrSql
                    objConn.Execute StrSql, , adExecuteNoRecords
                    
                    Id_batproc = getLastIdentity(objConn, "batch_Proceso")
                End If
                
                'Actualizo la fecha de ultima corrida y proxima ejecucion
                StrSql = "UPDATE batch_procplan "
                StrSql = StrSql & "SET fecultejec = " & ConvFecha(fechaActual)
                StrSql = StrSql & ",horaultejec = '" & Format(horaActual, "HH:mm") & "'"
                StrSql = StrSql & ",fecproxejec = " & ConvFecha(fecProxEjec)
                StrSql = StrSql & ",horaproxejec = '" & Format(horProxEjec, "HH:mm") & "'"
                StrSql = StrSql & " WHERE bproplanro = " & rs_Consult!bproplanro
                Flog.writeline Espacios(Tabulador * 1) & "Actualizo Planif : " & StrSql
                objConn.Execute StrSql, , adExecuteNoRecords
            
            Case 422:
                 FechaDesde = Date
                 FechaHasta = Date
                     Parametros = Split(rs_Consult("Param"), "@")
                    'Buscar la carpeta del modelo 1
                    Call BuscarModelo(Parametros(0), Directorio, OK)
                    If Not OK Then
                        Flog.writeline Espacios(Tabulador * 1) & "No se puede Automatizar el proceso de interfaces"
                        Exit Sub
                    End If
                    Path = Directorio
                    Set folder = fs.GetFolder(Directorio)
                    Set CArchivos = folder.Files
                    Encontro_Arch = False
                    Archivo_XML = ""
                    For Each Archivo In CArchivos
                        NArchivo = Archivo.Name
                        Flog.writeline Espacios(Tabulador * 1) & "Procesando archivo " & Archivo.Name
                        Call LeeArchivo_F572(Archivo.Name, modInterface, modParametros, Encontro_Arch, Archivo_XML)
                        If modParametros <> "" Then
                            Call BuscarModelo(modInterface, DirectorioDestino, OK)
                            If Not OK Then
                                Flog.writeline Espacios(Tabulador * 1) & "No se puede encontrar el la carpeta correspondiente al modelo destino " & modInterface
                                Exit Sub
                            End If
                            
                            modParametros = modInterface & "@" & Nombre_Archivo & "@" & modParametros
                            Flog.writeline ("Modelo Parametros") & modParametros
                            
                            'Set f = fs.getfile(Directorio & "\" & Archivo.Name)
                            
'                            On Error Resume Next
'                            f.Copy DirectorioDestino & "\" & Archivo.Name
'                            Flog.writeline (" Directorio Destino Modelo") & DirectorioDestino
'                            On Error GoTo E_BuscarPlanificados
'
'                            On Error Resume Next
'                            Err.Number = 0
'                            f.Move Directorio & "\bk\" & Archivo.Name
'                            Flog.writeline (" Directorio Destino a mover bk") & Directorio
'                            If Err.Number <> 0 Then
'                                If Err.Number = 58 Then
'                                    Err.Number = 0
'                                    f.Move Directorio & "\bk\" & Mid(Archivo.Name, 1, Len(Archivo.Name) - 4) & "_" & Left(Format(Time(), "HH-mm-ss"), 8) & "." & Right(Archivo.Name, 3)
'                                    If Err.Number <> 0 Then
'                                        Flog.writeline Espacios(Tabulador * 0) & "No se puede hacer bk del archivo original."
'                                    End If
'                                    Err.Number = 0
'                                Else
'                                    Flog.writeline Espacios(Tabulador * 0) & "La carpeta Destino no existe. Se creará."
'                                    Set Carpeta = fs.CreateFolder(Path & "\bk")
'                                    f.Move Directorio & "\bk\" & Archivo.Name
'                                End If
'                            End If
                            Flog.writeline Espacios(Tabulador * 1) & "siguiente archivo "
                            Flog.writeline
                            On Error GoTo E_BuscarPlanificados
                       End If
                    Next
                    
                    'Flog.writeline "Busco Archivos zip para eliminar"
                    
                    StrSql = "SELECT modarchdefault FROM modelo WHERE modnro = 292"
                    OpenRecordset StrSql, rs_Consult2
                        If Not rs_Consult2.EOF Then
                            Modelo = rs_Consult2!modarchdefault
                        End If
                    
                    StrSql = "SELECT sis_dirsalidas FROM sistema"
                    OpenRecordset StrSql, rs_Consult2
                    If Not rs_Consult2.EOF Then
                        l_path = rs_Consult2!sis_dirsalidas
                    End If
                    Directorio = l_path & Modelo
                    
                    Set folder = fs.GetFolder(Directorio)
                    Set CArchivos = folder.Files
                    Set fs = CreateObject("Scripting.FileSystemObject")
                    Set f = CreateObject("Scripting.FileSystemObject")

'                    For Each Archivo In CArchivos
'                        NArchivo = Archivo.Name
'                        If Right(NArchivo, 3) = "zip" Then
'                            Set f = fs.getfile(Directorio & "\" & Archivo.Name)
'
'                            On Error Resume Next
'                            Directorio = Directorio & "\" & NArchivo
'                            fs.DeleteFile Directorio, True
'                        End If
'                    Next
                    
                    Flog.writeline "Encontro Archivo: " & Encontro_Arch
                    If Encontro_Arch Then
                                                    'Inserto en batch_proceso para la ejecucion del mismo
                            StrSql = "INSERT INTO Batch_Proceso (btprcnro, bprcfecha, iduser, bprchora, bprcfecdesde, bprcfechasta,"
                            StrSql = StrSql & " bprcestado, empnro, bprcConfirmado, bprcparam)"
                            StrSql = StrSql & " values (" & rs_Consult!btprcnro & "," & ConvFecha(fechaActual) & ", '" & rs_Consult!IdUser & "'" & ",'" & Format(horaActual, "hh:mm:ss ") & "'"
                            StrSql = StrSql & " , " & ConvFecha(FechaDesde) & ", " & ConvFecha(FechaHasta)
                            StrSql = StrSql & " , 'Insertando', 0, -1,' ')"
                            objConn.Execute StrSql, , adExecuteNoRecords

                            Id_batproc = getLastIdentity(objConn, "batch_Proceso")

                            StrSql = "UPDATE Batch_Proceso SET "
                            StrSql = StrSql & " bprcestado = 'Pendiente'"
                            StrSql = StrSql & " ,bprcparam ='" & modParametros & "'"
                            StrSql = StrSql & " WHERE bpronro = " & Id_batproc
                            Flog.writeline Espacios(Tabulador * 1) & "Inserto Proceso: " & StrSql
                            objConn.Execute StrSql, , adExecuteNoRecords

                            ' Insertar información de las Interfaces que se ejecutan en dicha planificacion
                            StrSql = "INSERT INTO batch_interplan(bproplanro,bpronro,bpromodnro)"
                            StrSql = StrSql & " VALUES (" & rs_Consult!bproplanro & "," & Id_batproc & "," & modInterface & " )"
                            objConn.Execute StrSql, , adExecuteNoRecords
                            
                            StrSql = "UPDATE batch_procplan "
                            StrSql = StrSql & "SET fecultejec = " & ConvFecha(fechaActual)
                            StrSql = StrSql & ",horaultejec = '" & Format(horaActual, "HH:mm") & "'"
                            StrSql = StrSql & ",fecproxejec = " & ConvFecha(fecProxEjec)
                            StrSql = StrSql & ",horaproxejec = '" & Format(horProxEjec, "HH:mm") & "'"
                            StrSql = StrSql & " WHERE bproplanro = " & rs_Consult!bproplanro
                            Flog.writeline Espacios(Tabulador * 1) & "Actualizo Planif : " & StrSql
                            objConn.Execute StrSql, , adExecuteNoRecords
                            
                            Flog.writeline "replanica f572: " & StrSql

                            'Inserto las lineas a procesar del archivo
                            'Call InsertatLineas(Path & "\bk\" & Archivo_XML, Id_batproc)
                    End If
                Case Else
    
                'Inserto en batch_proceso para la ejecucion del mismo
                StrSql = "INSERT INTO Batch_Proceso (btprcnro, bprcfecha, iduser, bprchora, bprcfecdesde, bprcfechasta,"
                StrSql = StrSql & " bprcestado, empnro, bprcConfirmado, bprcparam)"
                StrSql = StrSql & " values (" & rs_Consult!btprcnro & "," & ConvFecha(fechaActual) & ", '" & rs_Consult!IdUser & "'" & ",'" & Format(horaActual, "hh:mm:ss ") & "'"
                StrSql = StrSql & " , " & ConvFecha(fechaActual) & ", " & ConvFecha(fechaActual)
                StrSql = StrSql & " , 'Pendiente', 0, -1"
                If EsNulo(rs_Consult!param) Then
                StrSql = StrSql & " ,NULL)"
                Else
                StrSql = StrSql & " ,'" & rs_Consult!param & "')"
                End If
                Flog.writeline Espacios(Tabulador * 1) & "Inserto Proceso: " & StrSql
                objConn.Execute StrSql, , adExecuteNoRecords
                
                'Actualizo la fecha de ultima corrida y proxima ejecucion
                StrSql = "UPDATE batch_procplan "
                StrSql = StrSql & "SET fecultejec = " & ConvFecha(fechaActual)
                StrSql = StrSql & ",horaultejec = '" & Format(horaActual, "HH:mm") & "'"
                StrSql = StrSql & ",fecproxejec = " & ConvFecha(fecProxEjec)
                StrSql = StrSql & ",horaproxejec = '" & Format(horProxEjec, "HH:mm") & "'"
                StrSql = StrSql & " WHERE bproplanro = " & rs_Consult!bproplanro
                Flog.writeline Espacios(Tabulador * 1) & "Actualizo Planif : " & StrSql
                objConn.Execute StrSql, , adExecuteNoRecords
    
        End Select
    
    End If
    
    Flog.writeline
    rs_Consult.MoveNext
Loop

If rs_Consult.State = adStateOpen Then rs_Consult.Close
Set rs_Consult = Nothing

Exit Sub

E_BuscarPlanificados:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: BuscarPlanificados"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="
End Sub


Public Sub CalcularProxEjec(ByVal planificador As Long, ByVal fechaActual As Date, ByVal horaActual, ByRef fecProxEjec As Date, ByRef horProxEjec As String, ByRef planifCorrecta As Boolean)

Dim frectipnro
Dim schedhora
Dim alesch_fecini
Dim alesch_fecfin
Dim alesch_frecrep
   
Dim rs As New ADODB.Recordset

On Error GoTo E_CalcularProxEjec

    planifCorrecta = False
    
    'Busco los datos del planificador
    StrSql = "SELECT frectipnro, alesch_fecini, schedhora, alesch_frecrep, alesch_fecfin "
    StrSql = StrSql & "FROM ale_sched "
    StrSql = StrSql & "WHERE schednro = " & planificador
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        
        frectipnro = rs!frectipnro
        schedhora = rs!schedhora
        alesch_fecini = rs!alesch_fecini
        alesch_fecfin = rs!alesch_fecfin
        alesch_frecrep = rs!alesch_frecrep
        
        'Verifico que el planificador este activado
        If (DateValue(fechaActual) >= DateValue(alesch_fecini)) And (DateValue(fechaActual) <= DateValue(alesch_fecfin)) Then
            
            
            Select Case frectipnro
                
                '---------------------------------------------------------------------------------
                'Diariamiente
                '---------------------------------------------------------------------------------
                Case 1
                    
                    Flog.writeline Espacios(Tabulador * 2) & "Planificacion Diaria a las " & schedhora
                    
                    'Sigt dia, misma hora
                    fecProxEjec = DateAdd("d", 1, fechaActual)
                    horProxEjec = Trim(schedhora) & ":00"
                    Flog.writeline Espacios(Tabulador * 2) & "Proxima Ejecucion: " & fecProxEjec & " a las " & horProxEjec
                    planifCorrecta = True
                    
                '---------------------------------------------------------------------------------
                ' Semanalmente
                '---------------------------------------------------------------------------------
                Case 2
                    
                    Flog.writeline Espacios(Tabulador * 2) & "Planificacion Semanal. Todos los " & WeekdayName(alesch_frecrep) & "  a las " & schedhora

                    If ((alesch_frecrep - Weekday(fechaActual)) > 0) Then
                        'El dia es dentro de la misma semana
                        fecProxEjec = DateAdd("d", alesch_frecrep - Weekday(fechaActual), fechaActual)
                    Else
                        'Mismo dia de la semana, pero la siguiente semana
                        fecProxEjec = DateAdd("d", 7 + alesch_frecrep - Weekday(fechaActual), fechaActual)
                    End If
                    
                    horProxEjec = Trim(schedhora) & ":00"
                    Flog.writeline Espacios(Tabulador * 2) & "Proxima Ejecucion: " & fecProxEjec & " a las " & horProxEjec
                    planifCorrecta = True
                    
                '---------------------------------------------------------------------------------
                ' Mensualmente
                '---------------------------------------------------------------------------------
                Case 3
                    
                    Flog.writeline Espacios(Tabulador * 2) & "Planificacion Mensual. Todos los dias " & alesch_frecrep & " del mes, a las " & schedhora
                    
                    'El control verifica casos para cuando se elige programacion mensual con dias 31 que no todos los meses lo contienen
                    If IsDate(alesch_frecrep & "/" & Month(DateAdd("m", 1, fechaActual)) & "/" & Year(DateAdd("m", 1, fechaActual))) Then
                        fecProxEjec = DateValue(alesch_frecrep & "/" & Month(DateAdd("m", 1, fechaActual)) & "/" & Year(DateAdd("m", 1, fechaActual)))
                        horProxEjec = Trim(schedhora) & ":00"
                        Flog.writeline Espacios(Tabulador * 2) & "Proxima Ejecucion: " & fecProxEjec & " a las " & horProxEjec
                        planifCorrecta = True
                    Else
                        Flog.writeline Espacios(Tabulador * 2) & "No existe la fecha para el mes. Fecha: " & alesch_frecrep & "/" & Month(DateAdd("m", 1, fechaActual)) & "/" & Year(DateAdd("m", 1, fechaActual))
                        planifCorrecta = False
                    End If
                    
                        
                '---------------------------------------------------------------------------------
                'Temporal
                '---------------------------------------------------------------------------------
                Case 4
                    
                    Flog.writeline Espacios(Tabulador * 2) & "Planificacion Temporal. Cada " & alesch_frecrep & " dias."
                    fecProxEjec = DateAdd("d", alesch_frecrep, fechaActual)
                    horProxEjec = horaActual
                    Flog.writeline Espacios(Tabulador * 2) & "Proxima Ejecucion: " & fecProxEjec & " a las " & horProxEjec
                    planifCorrecta = True
                    
                End Select
            
            
            ' Fecha siguiente fuera del tope maximo
            If DateValue(fecProxEjec) > DateValue(alesch_fecfin) Then
                planifCorrecta = False
                Flog.writeline Espacios(Tabulador * 2) & "La planificacion calculada excede el rango final del planificador."
            End If
            
        Else
            'Fecha fuera de rango
            planifCorrecta = False
            Flog.writeline Espacios(Tabulador * 2) & "Planificacion fuera de rango. Revise las fechas desde/hasta del Planificador"
        End If
    
    Else
        planifCorrecta = False
        Flog.writeline Espacios(Tabulador * 2) & "Error Planificacion. No se encontro el planifcador " & planificador
    End If
    
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    
Exit Sub

E_CalcularProxEjec:
    HuboError = True
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: CalcularProxEjec"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="
End Sub


'Calucula la fechas en que tiene que procesar los procesos planificados
Private Sub CalcFechaProcesamiento(ByVal rango As Integer, ByVal frecuencia As Integer, ByRef FDesde As Date, ByRef FHasta As Date)
 Dim Fecha As Date
 Dim dias As Integer
 
 FDesde = Now
 FHasta = Now
 Fecha = Now
 
    Select Case frecuencia
    '-1 -> Anterior. 0 Actual
    
        'Diariamiente
        Case 1
            If rango = -1 Then
              FDesde = DateAdd("d", -1, Fecha)
            End If
             
        'Semanal
        Case 2
            If rango = 0 Then
                FDesde = Inicio_de_Semana(Fecha)
            Else
                Fecha = DateAdd("d", -7, Fecha)
                FDesde = Inicio_de_Semana(Fecha)
                FHasta = fin_de_Semana(Fecha)
            End If
             
        'Mes actual
        Case 3
            If rango = 0 Then
                FDesde = DateAdd("d", -(Day(Now) - 1), Fecha)
            Else
                FHasta = DateAdd("m", -1, Fecha)
                FDesde = DateAdd("d", -(Day(Now) - 1), Fecha)
                FHasta = fin_del_Mes(Fecha)
            End If
        
        'Temporal
        Case Else
            FDesde = DateAdd("d", CDbl(rango), Fecha)
        
    End Select

End Sub


'Devuelve el último días del Mes
Public Function fin_del_Mes(Fecha As Variant) As Date
    If IsDate(Fecha) Then
        fin_del_Mes = DateAdd("m", 1, Fecha)
        fin_del_Mes = DateSerial(Year(fin_del_Mes), Month(fin_del_Mes), 1)
        fin_del_Mes = DateAdd("d", -1, fin_del_Mes)
    End If
End Function

'Devuelve la Fecha del último día de la semana
Function fin_de_Semana(ByVal Fecha As Date) As Date
    If IsDate(Fecha) And (Weekday(Fecha) <> 1) Then
        fin_de_Semana = FormatDateTime(Fecha - Weekday(Fecha) + 8, vbGeneralDate)
    Else
        fin_de_Semana = Fecha
    End If
End Function

'Devuelve la Fecha del primer día de la semana
Function Inicio_de_Semana(ByVal Fecha As Date) As Date
    If IsDate(Fecha) And Weekday(Fecha) <> 1 Then
        Inicio_de_Semana = FormatDateTime(Fecha - (Weekday(Fecha) - 2), vbGeneralDate)
    Else
        Inicio_de_Semana = FormatDateTime(Fecha - 6, vbGeneralDate)
    End If
End Function




Public Sub BuscarModelo(ByVal NroModelo As Long, ByRef Directorio As String, ByRef OK As Boolean)
' -----------------------------------------------------------------------------------------------------------------------
' Descripcion: Busca el modelo y retorna la ruta completa dende se alojan los archivos a levantar
' Autor      : FGZ
' Fecha      : 07/09/2010
' Ult. Mod   :
' -----------------------------------------------------------------------------------------------------------------------
Dim rs As New ADODB.Recordset

On Error GoTo E_BuscarModelo



    StrSql = "SELECT sis_direntradas FROM sistema WHERE sisnro = 1 "
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        Directorio = Trim(rs!sis_direntradas)
    Else
        Flog.writeline Espacios(Tabulador * 1) & "No se encontró el registro de la tabla sistema nro 1"
        OK = False
        Exit Sub
    End If
    
    StrSql = "SELECT * FROM modelo WHERE modnro = " & NroModelo
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        Directorio = Directorio & Trim(objRs!modarchdefault)
     Else
        Flog.writeline Espacios(Tabulador * 1) & "No se encontró el modelo " & NroModelo
        OK = False
        Exit Sub
    End If
    OK = True
Exit Sub

E_BuscarModelo:
    HuboError = True
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: BuscarModelo"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="
    OK = False
End Sub



Private Sub InsertatLineas(ByVal NombreArchivo As String, ByVal Bpronro As Long)
' -----------------------------------------------------------------------------------------------------------------------
' Descripcion: Abre el archivo y revisa de que modelo se trata. La 1er linea tiene la informacion
' Autor      : FGZ
' Fecha      : 07/09/2010
' Ult. Mod   :
' -----------------------------------------------------------------------------------------------------------------------
Const ForReading = 1
Const TristateFalse = 0
Dim strLinea As String
Dim datos
Dim f
Dim NroLinea As Long
Dim Ciclos As Long

    'Espero hasta que se crea el archivo
    On Error Resume Next
    Err.Number = 1
    Ciclos = 0
    Do Until Err.Number = 0
        Err.Number = 0
        Set f = fs.getfile(NombreArchivo)
        If f.Size = 0 Then
            If Ciclos > 100 Then
                Flog.writeline Espacios(Tabulador * 0) & "No anda el getfile "
            Else
                Err.Number = 1
                Ciclos = Ciclos + 1
            End If
        End If
    Loop
    On Error GoTo 0
   
   'Abro el archivo
    On Error GoTo CE
    
    Set f = fs.OpenTextFile(NombreArchivo, ForReading, TristateFalse)
    
    Do While Not f.AtEndOfStream
        strLinea = f.ReadLine
        NroLinea = NroLinea + 1
    
        If Trim(strLinea) <> "" Then
            StrSql = "INSERT INTO modelo_filas (bpronro, fila) "
            StrSql = StrSql & " VALUES (" & Bpronro & "," & NroLinea & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
    Loop
    f.Close
    
   
Fin_InsertatLineas:
    Exit Sub
    
CE:
    HuboError = True
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 0) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline Espacios(Tabulador * 0) & "Error: " & Err.Number
    Flog.writeline Espacios(Tabulador * 0) & "Decripcion: " & Err.Description
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline
    GoTo Fin_InsertatLineas
End Sub


Private Sub LeeArchivo(ByVal NombreArchivo As String, ByRef Modelo As Long, ByRef Parametros As String)
' -----------------------------------------------------------------------------------------------------------------------
' Descripcion: El nomre del archivo determina el modelo.
' Autor      : FGZ
' Fecha      : 08/09/2010
' Ult. Mod   :
' -----------------------------------------------------------------------------------------------------------------------
Dim Aux_modelo
Dim Aux_parametros

'detalles de los archivos
'Los mismos deben respetar la forma : nnnnaaaammddhhmm.csv donde:
'   encabezado son 4 caracteres alfanumericos seguidos de un _
'   nnnn: 'Numero de interface - interface number (ie. 605)
'   aaaa: Año de creacion de archivo - file creation year
'   mm: Mes de creacion de archivo - file creation month
'   dd: Dia de creacion de archivo - file creation date
'   hh: Hora de creacion de archivo - file creation hour
'   mm: Minutos de creacion de archivo - file creation minutes

'Ejemplo 0605201009211450.csv

'FGZ - 21/09/2010
'Aux_modelo = Left(NombreArchivo, 3)
'Aux_modelo = Mid(NombreArchivo, 6, 4)
'FGZ - 23/09/2010
'Aux_modelo = Left(NombreArchivo, 4)

'FGZ - 03/11/2010 -redefinicion del nombre de los archivos
'El formato es el siguiente: aaaammddhhmmss_nnnn
'    aaaa = Año
'    mm = mes
'    dd = dia
'    Hh = Hora
'    mm = minuto
'    Ss= secuencia. Esta secuencia se determina en base al orden de ejecución de interfaces que nos solicito RH pro. Ejemplo, las interfaz 605 =00, la interfaz 630 =05, etc.
'    Nnnn= numero de interfaz en 4 digitos.
'Ejemplo 20101028123100_0605.txt
Aux_modelo = Right(NombreArchivo, 8)
Aux_modelo = Left(Aux_modelo, 4)

'If Not IsNumeric(Aux_modelo)  Then
If Not IsNumeric(Aux_modelo) Or UCase(Right(NombreArchivo, 3)) = "XML" Then
    Flog.writeline Espacios(Tabulador * 1) & "Error. El nombre del archivo " & NombreArchivo & " no respeta formato (nnnnaaaammddhhmm.csv) "
    Modelo = 0
    Parametros = ""
Else
    'De acuerdo al nro de modelo los parametros que lleva
    Select Case CLng(Aux_modelo)
        Case 211: 'Interface de Novedades
            'AccionNovedad
                '-1 Reemplazar Novedades
                '0 Mantiene Novedades
                '1 Sumar Novedades
            
            'Default
            Aux_parametros = "0"
        Case 214: 'Tickets
            'TikPedNro
            'Default
            Aux_parametros = "0"
        Case 226: 'Interface de Postulantes Deloitte
            'Default
            Aux_parametros = "0"
        Case 239: 'Interfase Deloitte
            'Default
            Aux_parametros = "0"
        Case 267: 'Interface Wella Cab de Facturacion
            'Pliqnro
            
            'Default
            Aux_parametros = "0"
        Case 268: 'Migracion de Empleados
            'Pisa
            
            'Default
            Aux_parametros = CStr(CBool(0)) 'False
        Case 272:   'Interface Novedades SPEC
            'Pisa Novedades
            'Tiene vigencia
            
            'Default
            Aux_parametros = CStr(CBool(0)) 'false
            Aux_parametros = Aux_parametros & "@" & CStr(CBool(0)) 'false
        Case 275: 'Interface de Postulantes Estandar
            'Default
            Aux_parametros = "0"
        Case 281: 'Alta masiva de Remuneraciones (LA CAJA)
            'Pisa
            
            'Default
            Aux_parametros = CStr(CBool(0)) 'False
        Case 300: 'Migracion de Empleados para TELEPERFORMANCE ( + 3 columnas)
            'Pisa
            
            'Default
            Aux_parametros = CStr(CBool(0)) 'False
        Case 313: 'Interfase de Novedades con seguridad
            'PisaNovedad
            
            'Default
            Aux_parametros = CStr(CBool(0)) 'False
        Case 318: 'Interface Valor Plan Obra Social - PRICE
            'PisaPlan
            
            'Default
            Aux_parametros = CStr(CBool(0)) 'False
        Case 605: 'Migracion de Empleados
            'Pisa
            
            'Default
            Aux_parametros = CStr(CBool(0)) 'False
        Case 606: 'Migracion de Empleados URU
            'Pisa
            
            'Default
            Aux_parametros = CStr(CBool(0)) 'False
        Case 607: 'Migracion de Empleados CHILE
            'Pisa
            
            'Default
            Aux_parametros = CStr(CBool(0)) 'False
        Case 608: 'Migracion de Empleados COLOMBIA
            'Pisa
            
            'Default
            Aux_parametros = CStr(CBool(0)) 'False
        Case 609: 'Migracion de Empleados 'Radiotronica
            'Pisa
            
            'Default
            Aux_parametros = CStr(CBool(0)) 'False
        Case 630: 'Migracion de Historico de estructuras
            'Pisa
            
            'Default
            Aux_parametros = CStr(CBool(0)) 'False
        Case 661 ' Migración de Imágenes de Empleados
            'Pisa
            
            'Default
            Aux_parametros = CStr(CBool(0)) 'False
    End Select
    
    Modelo = Aux_modelo
    Parametros = Aux_parametros
    
End If
'Modelo = Aux_modelo
'Parametros = Aux_parametros

End Sub

'--------------------------------------------------------------------
' Se encarga de descomprimir archivos Zip
'--------------------------------------------------------------------

Public Sub Descomprimir(ByVal NombreArchivo As String)

Dim StrSql As String
Dim rsConsult As New ADODB.Recordset
Dim rsConsult2 As New ADODB.Recordset
Dim ArchivoZip As String
Dim Modelo
Dim l_path
Dim destino
Dim destinoZip
Dim PID
Dim fs, f
Dim folder
Dim l_path_zip
Dim l_zip_del
Dim tiempo


Set fs = CreateObject("Scripting.FileSystemObject")
Set f = CreateObject("Scripting.FileSystemObject")

Flog.writeline "Busco Directorios a donde descomprimir"

StrSql = "SELECT modarchdefault FROM modelo WHERE modnro = 1"
OpenRecordset StrSql, rsConsult2
    If Not rsConsult2.EOF Then
        Modelo = rsConsult2!modarchdefault
    End If

StrSql = "SELECT sis_dirsalidas FROM sistema"
OpenRecordset StrSql, rsConsult2
If Not rsConsult2.EOF Then
    l_path = rsConsult2!sis_dirsalidas
End If
l_path = l_path & Modelo

Flog.writeline "Obtengo Archivo Origen en la carpeta del Modelo 1"

Set f = fs.getfile(l_path & "\" & NombreArchivo)

StrSql = "SELECT modarchdefault FROM modelo WHERE modnro = 292"
OpenRecordset StrSql, rsConsult2
    If Not rsConsult2.EOF Then
        Modelo = rsConsult2!modarchdefault
    End If

StrSql = "SELECT sis_dirsalidas FROM sistema"
OpenRecordset StrSql, rsConsult2
If Not rsConsult2.EOF Then
    l_path = rsConsult2!sis_dirsalidas
End If
l_zip_del = l_path & "\Form572\"
l_path = l_path & Modelo

Flog.writeline "Muevo Archivo a destino"

'f.Move l_path & "\" & NombreArchivo 'MDF

destino = l_path
Flog.writeline "Directorio a descomprimir " & destino

destinoZip = l_path & "\" & NombreArchivo
Flog.writeline "Archivo a Descomprimir " & destinoZip

StrSql = "SELECT sisdir FROM sistema"
OpenRecordset StrSql, rsConsult2
If Not rsConsult2.EOF Then
    l_path_zip = rsConsult2!sisdir
End If

l_path_zip = """" & l_path_zip & "\cgi-bin\servicios\zip\7za.exe" & """ "

l_path_zip = l_path_zip & " e "

Flog.writeline "Servicio 7Zip " & l_path_zip

Flog.writeline "Empieza la descompresion"
Dim I
Dim ruta_final As String

ruta_final = ""
Dim Elimina_Archivo
Dim termino As Boolean
Dim hProc As Long
'Dim fs
Dim fs5
    I = 0
    'Elimina_Archivo = l_zip_del & NombreArchivo
     Elimina_Archivo = l_zip_del
     Set fs5 = CreateObject("Scripting.FileSystemObject")
    'ruta_final = "cmd.exe /c " & l_path_zip & " """ & destinoZip & """ -o""" & destino & """ & " & " DEL /F /S /Q /A " & """" & Elimina_Archivo & """"
    
    '-----------------
      'fs5.DeleteFile (Elimina_Archivo)
      Dim arch
      Dim CArchivos
      Dim arreglo_carpetas
      Dim arreglo_archivo
      Set folder = fs5.GetFolder(Elimina_Archivo)
      Set CArchivos = folder.Files
      For Each arch In CArchivos
        arreglo_carpetas = Split(arch, "\")
        arreglo_archivo = Split(arreglo_carpetas(UBound(arreglo_carpetas)), ".")
        If UBound(arreglo_archivo) > 0 Then
          If arreglo_archivo(1) = "zip" Then
            Flog.writeline "Eliminando:" & arch
            fs5.DeleteFile (arch)
          End If
        End If
      Next
      
      f.Move l_path & "\" & NombreArchivo 'MDF
    
    '------------------
    
    
    ruta_final = l_path_zip & " """ & destinoZip & """ -o""" & destino & """"
    
    Flog.writeline "comando de descompresion:" & ruta_final
    
    PID = Shell(ruta_final, vbHide)
    
    
    
    Flog.writeline "Descompresion exitosa..."

   
    If PID <> 0 Then
        Flog.writeline "    Ejecutando Planificador ... PID = " & PID
    End If

Exit Sub

MError:
    Flog.writeline "Error: " & Err.Description
    'HuboErrores = True
    'EmpErrores = True
    Exit Sub
End Sub

Private Sub LeeArchivo_F572(ByVal NombreArchivo As String, ByRef Modelo As Long, ByRef Parametros As String, ByRef Encontro_Arch As Boolean, ByRef Archivo_XML As String)
' -----------------------------------------------------------------------------------------------------------------------
' Descripcion: El nomre del archivo determina el modelo.
' Autor      : FGZ
' Fecha      : 08/09/2010
' Ult. Mod   :
' -----------------------------------------------------------------------------------------------------------------------
Dim Aux_modelo
Dim Aux_modelo2
Dim Aux_modelo3
Dim Aux_parametros
Dim mostrar_error
Aux_modelo = Left(NombreArchivo, 11)
Aux_modelo2 = Mid(NombreArchivo, 13, 4)
Aux_modelo3 = Right(NombreArchivo, 3)
mostrar_error = True
If Aux_modelo3 = "zip" Then
    Call Descomprimir(NombreArchivo)
    Encontro_Arch = True
    mostrar_error = False
End If

If (Not IsNumeric(Aux_modelo) Or Aux_modelo3 <> "xml") And mostrar_error Then
    Flog.writeline Espacios(Tabulador * 1) & "Error. El nombre del archivo " & NombreArchivo & " no respeta formato(xml) "
    Modelo = 0
    Parametros = ""
Else
    Parametros = "0"
    Encontro_Arch = True
    Archivo_XML = NombreArchivo
End If
Modelo = 292
End Sub

