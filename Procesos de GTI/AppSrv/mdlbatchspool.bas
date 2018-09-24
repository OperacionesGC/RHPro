Attribute VB_Name = "mdlbatchspool"
'Spooler de procesos
'Creado el 7/3/2003
'Alvaro Bayon - 'Fernando Zwenger
'----------------------------------------
Option Explicit

'************************************************************************************
'Versiones

'Const Version = 2.01    'Inicial.
'Const FechaVersion = "27/04/2006"

'Const Version = 2.02
'Const FechaVersion = "14/08/2006"   'FGZ - 'No estaba registrando bien la hora de cada movimiento, funcion now.

'Const Version = 2.03
'Const FechaVersion = "21/11/2006"   'FGZ - 'Estaba controlando mal los limites de los array de pendientes y ejecutando
'                                    '       Ademas si la cantidad de procesos pendientes es mayor = al limite ==> la funcion calcularPesos daba error

'Const Version = "2.04"
'Const FechaVersion = "20/04/2007"   'FGZ - 'se le agregaron estas variables pa efectos estadisticos
''                                           Global Cantidad_de_OpenRecordsetAS As Long
''                                           Global Cantidad_Call_Politicas As Long
''                                           Global Usurio As String

'Const Version = "2.05"
'Const FechaVersion = "10/09/2007"   'FGZ - Se agreg� el proceso de Planificador
''                                       Esta modificacion requiere agregar dos parametros en RHProappSrvDefaults.ini
''                                       ----------------------------------------------------------------------------
''                                        Tiempo de Espera No Responde (Minutos) = [5]
''                                        Tiempo de Espera Sin Progreso (Minutos) = [5]
''                                        Tiempo de lectura de Registraciones (Minutos) = [1]
''                                        Tiempo de Dormida (segundos) = [1]
''                                        Usa Lectura de Registraciones = [-1]
''                                        Maximo Nro de Procesos Concurrentes (Tipicamente 5) = [1]
''                                        Genera multiples Archivos de LOG (uno por dia) = [-1]
''                                        Cantidad de reintentos de Mensajeria = [3]
''                                        Tiempo entre reintentos de Mensajeria (Minutos) = [1]
''                                        Usa Planificador = [-1]
''                                        Tiempo entre ejecuciones del planificador (Minutos) = [5]
''                                       ----------------------------------------------------------------------------

'Const Version = "2.06"
'Const FechaVersion = "02/10/2007"
''                                       ----------------------------------------------------------------------------
''                                       FGZ - Se quitaron los mensajes de log innecesarios
''                                       ----------------------------------------------------------------------------

'Const Version = "2.07"
'Const FechaVersion = "06/02/2008"
''                                       ----------------------------------------------------------------------------
''                                       FGZ - Se hicieron unos cambios en el sub EliminarProcesosMarcados
''                                       ----------------------------------------------------------------------------

'Const Version = "2.08"
'Const FechaVersion = "26/06/2008"
''                                       ----------------------------------------------------------------------------
''                                       FGZ - Cambio de schema para Oracle
''                                       ----------------------------------------------------------------------------
'''                                       Esta modificacion requiere agregar un parametros en RHProappSrv.ini
'''                                       ----------------------------------------------------------------------------
'''                                        SAP = [path donde genera el archivo de exportacion para sap]
'''                                        Procesos = [path de los ejecutables de los procesos]
'''                                        Flog = [path de los archivos de log]
'''                                        Fecha = [formato de fecha]
'''                                        conexion = [string de conexion]
'''                                        TipoDB = [3 o 4]
'''                                        Etiqueta = [etiqueta en procesos.ini]
'''                                        SCHEMA = [nombre del schema]
'''                                       ----------------------------------------------------------------------------


'Const Version = "2.08"
'Const FechaVersion = "26/06/2008"
''                                       ----------------------------------------------------------------------------
''                                       FGZ - se modific� el control de incompatibilidades
''                                       ----------------------------------------------------------------------------

'Const Version = "2.09"
'Const FechaVersion = "21/10/2008"
''                                       ----------------------------------------------------------------------------
''                                       FGZ - se modific� el control de incompatibilidades para hacerlo mas especifico de acuerdo al tipo de proceso
''                                       ----------------------------------------------------------------------------



'Const Version = "2.10"
'Const FechaVersion = "13/02/2009"
''                                       ----------------------------------------------------------------------------
''                                       FGZ - se agreg� un parametro en defaults.ini para saber si tengo que desencriptar el string de conexion
''                                       ----------------------------------------------------------------------------
''''                                       Esta modificacion requiere agregar un parametros en RHProappSrvDefaults.ini
''''                                       ----------------------------------------------------------------------------
''''                                        Tiempo de Espera No Responde (Minutos) = [5]
''''                                        Tiempo de Espera Sin Progreso (Minutos) = [5]
''''                                        Tiempo de lectura de Registraciones (Minutos) = [1]
''''                                        Tiempo de Dormida (segundos) = [1]
''''                                        Usa Lectura de Registraciones = [0]
''''                                        Maximo Nro de Procesos Concurrentes (Tipicamente 5) = [3]
''''                                        Genera multiples Archivos de LOG (uno por dia) = [-1]
''''                                        Cantidad de reintentos de Mensajeria = [3]
''''                                        Tiempo entre reintentos de Mensajeria (Minutos) = [1]
''''                                        Usa Planificador = [0]
''''                                        Tiempo entre ejecuciones del planificador (Minutos) = [5]
''''                                        conexion ecriptada = [0]
''''                                       ----------------------------------------------------------------------------


'Const Version = "2.11"
'Const FechaVersion = "16/02/2009"
''                                       ----------------------------------------------------------------------------
''                                       FGZ - se agregaron 2 parametros en defaults.ini
''                                       ----------------------------------------------------------------------------
'''                                       Esta modificacion requiere agregar un parametros en RHProappSrvDefaults.ini
'''                                       ----------------------------------------------------------------------------
'''                                        Tiempo de Espera No Responde (Minutos) = [5]
'''                                        Tiempo de Espera Sin Progreso (Minutos) = [5]
'''                                        Tiempo de lectura de Registraciones (Minutos) = [1]
'''                                        Tiempo de Dormida (segundos) = [1]
'''                                        Usa Lectura de Registraciones = [0]
'''                                        Maximo Nro de Procesos Concurrentes (Tipicamente 5) = [3]
'''                                        Genera multiples Archivos de LOG (uno por dia) = [-1]
'''                                        Cantidad de reintentos de Mensajeria = [3]
'''                                        Tiempo entre reintentos de Mensajeria (Minutos) = [1]
'''                                        Usa Planificador = [0]
'''                                        Tiempo entre ejecuciones del planificador (Minutos) = [5]
'''                                        conexion ecriptada = [0]
'''                                        Procesos habilitados = [0]
'''                                        Procesos no habilitados = [0]
'''                                       ----------------------------------------------------------------------------


'Const Version = "2.12"
'Const FechaVersion = "26/06/2009"
''                                       ----------------------------------------------------------------------------
''                                       FGZ - habia un problema cuando la semilla de encriptacion es 0
''                                       ----------------------------------------------------------------------------


'Const Version = "2.13"
'Const FechaVersion = "08/10/2009"
''                                       ----------------------------------------------------------------------------
''                                       FGZ - habia un problema cuando cambiaba de dia y la conexion esta encriptada
''                                       ----------------------------------------------------------------------------


'Const Version = "2.14"
'Const FechaVersion = "16/10/2009"
''                                       ----------------------------------------------------------------------------
''                                       FGZ - Ahora hay un nuevo estado en el STD de los procesos
''                                                   Cuando se hace el shell ahora se actualiza el estado del proceso a Iniciando.
''
''                                               Se agreg� el control de procesos que nunca inician.
''                                               Si en el Tiempo de Espera No Responde configurado en el ini el proceso se dispar� pero nunca inici�
''                                               Entonces el AppSrv mata los procesos y los pone en estado Abortado.
''                                       ----------------------------------------------------------------------------

'Const Version = "2.15"
'Const FechaVersion = "13/01/2010"
''                                       ----------------------------------------------------------------------------
''                                       FGZ - Se cambi� mensaje de log cuando aborta un proceso que no inicia (ponia la hora en 00:00:00)
''                                             Se modific� el sub EliminarProcesosNoInician porque estaba haciendo mal los controles
''                                             Ademas se cambiaron algunos mensajes de log


'Const Version = "2.16"
'Const FechaVersion = "30/06/2010"
''                                       ----------------------------------------------------------------------------
''                                       FGZ - Se cambi� el sub CargarConfiguracionesBasicasAppSrv
''                                             controla que la linea de schequema este vacia


'Const Version = "2.17"
'Const FechaVersion = "29/12/2011"   'FGZ
''                                       ----------------------------------------------------------------------------
''                                       Se cambi� el sub PuedeEjecutar.
''                                             controla por procesos en estado iniciando
''                                       Se cambi� el sub EliminarProcesosNoInician
''                                           Aborta procesos que no inician

'Const Version = "2.18"
'Const FechaVersion = "03/01/2012"   'FGZ
''                                       ----------------------------------------------------------------------------
''                                       Se le sacaron los mensajes de log cuando todavia no es la hora de ejecucion (as� no crece tanto el archivo de log).

'-------------------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------------
'Const Version = "3.00"
'Const FechaVersion = "27/06/2012"   'FGZ
''                                       ----------------------------------------------------------------------------
''                                       Se recompil� el proceso para que utilice MDAC 2.8.

'Const Version = "3.01"
'Const FechaVersion = "08/08/2012"   'FGZ
''                                       ----------------------------------------------------------------------------
''                                       Se modific� el control de incompatibilidades para los procesos de Simulacion
''                                       Los procesos de Copiado, simulacion y borrado de simulacion son incompatibles y si hay alguno corriendo no se puede ejecutar otro
''                                       Se cambiaron algunos mensajes de log que estaban fijos y/o mal
''
''#Const Version = "3.0.1"

'Const Version = "3.02"
'Const FechaVersion = "18/10/2012"   'FGZ
''                                       ----------------------------------------------------------------------------
''                                       Se modific� el formato de fecha interno (sin Hora) para el control de procesos incompatibles
'
''                                       Se redefini� el OpenRecordset por los problemas de corte de conexion. Ahora en este proceso se usa OpenrecordsetAS
'
''                                       Se agreg� funcionalidad de Depuracion de carpeta de logs. Los parametros salen
''                                       FGZ - se agregaron 2 parametros en defaults.ini
''                                       ----------------------------------------------------------------------------
''                                       Esta modificacion requiere agregar un parametros en RHProappSrvDefaults.ini
''                                       ----------------------------------------------------------------------------
''                                        Tiempo de Espera No Responde (Minutos) = [5]
''                                        Tiempo de Espera Sin Progreso (Minutos) = [5]
''                                        Tiempo de lectura de Registraciones (Minutos) = [1]
''                                        Tiempo de Dormida (segundos) = [1]
''                                        Usa Lectura de Registraciones = [0]
''                                        Maximo Nro de Procesos Concurrentes (Tipicamente 5) = [3]
''                                        Genera multiples Archivos de LOG (uno por dia) = [-1]
''                                        Cantidad de reintentos de Mensajeria = [3]
''                                        Tiempo entre reintentos de Mensajeria (Minutos) = [1]
''                                        Usa Planificador = [0]
''                                        Tiempo entre ejecuciones del planificador (Minutos) = [5]
''                                        conexion ecriptada = [0]
''                                        Procesos habilitados = [0]
''                                        Procesos no habilitados = [0]
''                                        Accion Depuracion = [0]
''                                        Mantiene Dias De Depuracion (Dias) = [30]
''                                       ----------------------------------------------------------------------------
''#Const Version = "3.0.2"


'Const Version = "3.03"
'Const FechaVersion = "05/12/2012"   'FGZ
''                                       ----------------------------------------------------------------------------
''                                       Se agreg� funcionalidad de Depuracion de carpeta de logs. Los parametros salen
'
''                                       FGZ - se agregaron 2 parametros en defaults.ini
''                                       ----------------------------------------------------------------------------
''                                       Esta modificacion requiere agregar un parametros en Notificacion.ini
''                                       Ejemplo de la estructura
''                                       ----------------------------------------------------------------------------
''                                        FromName = [RHPro Saas]
''                                        Subject = [RHDESA.DES_R3 - Servicio Detenido por error]
''                                        Body = [Appsrv detenido por Error]
''                                        TO=[fzwenger@rhpro.com]
''                                        CC = [lmoro@rhpro.com; emargiotta@rhpro.com]
''                                        CCO = []
''                                        FromAddress = [rhpro@lisandromoro.com.ar]
''                                        Host = [mail.lisandromoro.com.ar]
''                                        Port = [25]
''                                        USER = [rhpro@lisandromoro.com.ar]
''                                        Password = [suipacha]
''                                        HTMLBody = []
''                                        HTMLMailHeader = []
''                                        Attachment = []
''                                       ----------------------------------------------------------------------------
''#Const Version = "3.0.3"


'Const Version = "3.04"
'Const FechaVersion = "21/08/2013"   'FGZ
''                                       ----------------------------------------------------------------------------
''                                       Se agreg� funcionalidad de verificacion y notificacion de Alertas activas y no planificadas
''#Const Version = "3.0.4"

'Const Version = "3.05"
'Const FechaVersion = "18/09/2013"   'FGZ
''                                       ----------------------------------------------------------------------------
''                                       Se agreg� funcionalidad de verificacion y notificacion de Alertas activas y no planificadas
''#Const Version = "3.0.5"


'Const Version = "3.06"
'Const FechaVersion = "20/09/2013"   'FGZ
''                                       ----------------------------------------------------------------------------
''                                       Cuando cambiaba de dia habia un problema en el chequeo de alertas y trataba de escribir en el log (que ya se habia cerrado por cambio de dia).
''                                       Esto hacia que diera un error y cuando trataba de escribir en el log el error daba otro error (porque el log esta cerrado) y como consecuencia el proceso quedaba colgado en memoria.
''#Const Version = "3.0.6"


'Const Version = "3.07"
'Const FechaVersion = "01/11/2013"   'FGZ
''                                       ----------------------------------------------------------------------------
''                                       Las notificaciones por error se hacer por reintentos de conexion se hacen cada 60 minutos.
''                                       Correcciones en el control de procesos en estado "NO Responde"
''#Const Version = "3.0.7"



'Const Version = "3.08"
'Const FechaVersion = "26/01/2015"   'FGZ
'                                       ----------------------------------------------------------------------------
'                                       Las notificaciones por error se hacer por reintentos de conexion se hacen cada 60 minutos.
'                                           Control sobre conexiones reestablecidas
'                                   EAM- Se comento las lineas que  inicializan del arreglo Ejecutando porque da error de abortado por el usuario.
'                                   Miriam Ruiz- CAS-29182 - BDO Base Kraft CHILE - BUG EN PROCESAR- se libera la versi�n

'#Const Version = "3.0.8"

'Const Version = "3.09"
'Const FechaVersion = "21/04/2015" ' Carmen Quintero - CAS-29445 - HyA - App Server detenido por error - Se agrego validacion para cuando los procesos quedan iniciando.

'#Const Version = "3.0.9"

Const Version = "3.10"
Const FechaVersion = "14/07/2015" ' Miriam Ruiz - CAS-31836 - Rhpro - Inconvenientes con opcion borrado de appsrv - Se arregl� el borrado de logs

'************************************************************************************
'************************************************************************************

Const MaxPendientes = 1000
Const ForReading = 1
Const ForAppending = 8
Const ForWriting = 2
Const FormatoInternoFecha = "dd/mm/yyyy HH:mm:ss"
Const FormatoInternoHora = "HH:mm:ss"
Const FormatoInternoFechaCorto = "dd/mm/yyyy"
Type TCelda
    Proceso As Long
    NombreProceso As String
    Peso As Single
    TipoProceso As Integer
    IdUser As String
    bprchora As String
    Fecha As Date
End Type

Type TCeldaEj
    Proceso As Long
    pid As Long
    Progreso As Single
    HoraInicioEj As Date
    HoraFinEj As Date
    Fecha As Date
    Intentos As Integer
End Type

'FGZ - 21/11/2006 - le cambi� la definicion de las variables
'Global Pendientes(MaxPendientes) As TCelda
'Global Ejecutando(MaxPendientes) As TCeldaEj
Global Pendientes(MaxPendientes + 1) As TCelda
Global Ejecutando(MaxPendientes + 1) As TCeldaEj

Private Const PROCESS_TERMINATE As Long = &H1
Private Const SYNCHRONIZE = &H100000

Private Declare Function OpenProcess Lib _
   "kernel32" (ByVal dwDesiredAccess As Long, _
   ByVal bInheritHandle As Long, _
   ByVal dwProcessID As Long) As Long

Private Declare Function TerminateProcess Lib _
   "kernel32" (ByVal hProcess As Long, _
   ByVal uExitCode As Long) As Long
   
Public Declare Function Sleep Lib _
   "kernel32" (ByVal dwMilliseconds As Long) As Long

Global objrsProcesosPendientes As New ADODB.Recordset
Global objrsRegistraciones As New ADODB.Recordset

Global Flog

' -------- Variables de control de tiempo ------------
' minutos que espera el spool antes de abortar el proceso
Global TiempoDeEsperaNoResponde As Integer
Global TiempoDeEsperaNoInician As Integer

' minutos que espera el spool antes de poner al proceso que se esta ejecutando en estado de No Responde
Global TiempoDeEsperaSinProgreso As Integer

' Tiempo entre lectura y lectura
Global TiempoDeLecturadeRegistraciones As Integer

' Tiempo de Dormida del Spool
Global TiempodeDormida As Integer

' Variable booleana que maneja si se usa Lectura de Registraciones o no
Global UsaLecturaRegistraciones As Boolean

'Maximo nro de Procesos Concurrentes
Global MaxConcurrentes As Integer

'Genera multiples Archivos de LOG (uno por dia)
Global MultiplesLOGs As Boolean
Global DiaAnterior As Date

'FGZ - 19/05/2004
Global UltimaRegInsertadaWFTurno As String  '(N) - Ninguna, (E) - Entrada y (S) - Salida

'FGZ - 19/1/2005
Global Etiqueta

'FGZ - 20/1/2005
Global Cantidad_Reintentos As Long
Global Tiempo_Reintentos As Long

'FGZ - 21/07/2005
Global FinDia As Boolean
'FGZ - 20/04/2007
Global Cantidad_de_OpenRecordset As Long
Global Cantidad_Call_Politicas As Long
Global Usuario As String

'FGZ - 20/07/2007 - Se agregaron estas 2 variables para manejar el proceso planificador -------
'Variable booleana que maneja si se usa proceso de Planificacion o no
Global UsaPlanificador As Boolean
'Tiempo entre ejecuciones del planificador
Global TiempoDePlanificador As Long
'FGZ - 20/07/2007 - Se agregaron estas 2 variables para manejar el proceso planificador -------
Global Proc_Hab As String
Global Proc_NoHab As String
'FGZ - 18/10/2012 --------------------------------------------
Global AccionConLogs As Long   '0 no depura, 1 Hace Bk en carpeta y '2 Elimina
Global DiasDepuracionLogs As Long
'                                x: Mantiene los logs de los ultimos X dias.

Global NotificaError As Boolean
Global ErrorNotificado As Boolean

Global NotifAnterior
Global NotifActual
Global TiempoEntreNotif

'
'------------------------------------------------------------
'------------------------------------------------------------


Public Sub Main()
Dim archivo As String
Dim fs, f
Dim strline As String
Dim tiposIncomp As String
Dim pos1 As Integer
Dim pos2 As Integer
'Dim path As String 'En esta variable va el path en que se encuentran los procesos
Dim pid
Dim cerrado As Boolean

Dim Actual As Integer
Dim Ultimo As Integer
Dim seguir As Boolean

Dim HoraEntre1 As Date
Dim HoraEntre2 As Date
Dim Nombre_Arch As String
Dim LecturaAnterior

Dim UltimaLectura
Dim LecturaActual
Dim TiempoEntreLecturas As Long

Dim PlanificacionAnterior
Dim UltimaPlanificacion
Dim PlanificacionActual
Dim TiempoEntrePlanificaciones As Long

Dim rs_MuestraPendientes As New ADODB.Recordset
Dim Contador As Long



Do While True
    'carga las configuraciones basicas, formato de fecha, string de conexion,
    'tipo de BD y ubicacion del archivo de log
    'Call CargarConfiguracionesBasicas
    Call CargarConfiguracionesBasicasAppSrv
    Call SetarDefaults
    'FGZ - 04/12/2012
    Call CargarMensaje("")
    DiaAnterior = Date
    'FGZ - 08/10/2009 - cuando cambia de dia cierra todo y vuelve a levantar la conf del ini ==>
    '                   vuelve a levantar el string de conexion encriptado
    Ya_Encripto = False
    PrimerConexion = True
    '--------------------------------------------------------------------------------
    ' FGZ 25/07/2003
    ' Abre para append el archivo de log, si no lo encuentra ==> lo crea
    If MultiplesLOGs Then
        Nombre_Arch = PathFLog & "RHProAppSrv " & Format(Date, "dd-mm-yyyy") & ".log"
    Else
        Nombre_Arch = PathFLog & "RHProAppSrv " & ".log"
    End If
    
    ' Primero tendr�a que chequear si existe, si es asi lo abro para appending y sino lo creo
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    ' crea o abre el archivo de log, segun corresponda
    Call AbrirArchivoLog(fs, Nombre_Arch)
    
    'Obtiene los datos de como esta configurado el servidor actualmente
    Call ObtenerConfiguracionRegional
   
    
Continuar:
    On Error GoTo 0
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaVersion
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Numero, separador decimal    : " & NumeroSeparadorDecimal
    Flog.writeline "Numero, separador de miles   : " & NumeroSeparadorMiles
    Flog.writeline "Moneda, separador decimal    : " & MonedaSeparadorDecimal
    Flog.writeline "Moneda, separador de miles   : " & MonedaSeparadorMiles
    Flog.writeline "Formato de Fecha del Servidor: " & FormatoDeFechaCorto
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline
    Flog.writeline "Inicio RHProAppSrv " & Format(Now, FormatoInternoFecha)
    On Error Resume Next
    Err.Number = 0
    OpenConnection strconexion, objConn
    
    'If Error_Encrypt Then
    If Error_Encrypt And PrimerConexion Then
        Flog.writeline Format(Now, FormatoInternoFecha) & ". No se pudo establecer la conexion con la Base de Datos. FIN"
        End
    End If
    Do While Err.Number <> 0
        Flog.writeline Format(Now, FormatoInternoFecha) & ". No se pudo establecer la conexion con la Base de Datos. Intenta nuevamente en 10 segundos"
        Flog.writeline Err.Description
        
        'pongo un delay y vuelvo a intentar
        'TiempodeDormida = 10
        Sleep (TiempodeDormida * 1000)
        Err.Number = 0
        OpenConnection strconexion, objConn
    Loop
    PrimerConexion = False
    'FGZ - 18/10/2012 --------------------------------------------------------------
    'AccionConLogs = 0
    'DiasDepuracionLogs = 10 'Dias
    Call DepurarLogs(AccionConLogs, DiasDepuracionLogs)
    'FGZ - 18/10/2012 --------------------------------------------------------------
    
    'Habilito el control de errores
    On Error GoTo ce
    
    'FGZ - 20/04/2007 - determino el usuario con el cual esta levantando el proceso / servicio
    'Usuario = GetCurrentUserId
    'FGZ - 20/04/2007 - determino el usuario con el cual esta levantando el proceso / servicio
    
    
    'FGZ - 16/08/2013 -----------------
    'Control de Alertas Activas y no pendientes.
    Call ControlarAlertas
    'FGZ - 16/08/2013 -----------------
    
    'TiempoDeLecturadeRegistraciones = 1 ' minutos
    LecturaAnterior = Format(C_Date(Date - 1), FormatoInternoFecha)
    PlanificacionAnterior = Format(C_Date(Date - 1), FormatoInternoFecha)
    FinDia = False
    Contador = 1
    Do While Not FinDia
        If Contador = 30 Then
            Flog.writeline "Analizando pendientes ...." & Format(Now, FormatoInternoFecha)
            Contador = 1
        Else
            Contador = Contador + 1
        End If
        ' Ac� tendria que lanzar el leer registraciones bajo dos condiciones
        ' que supere el tiempo preestablecido entre ejecuciones para este tipo de proceso
        ' que no haya otro leer registraciones ni prc30 ejecutandose
        If UsaLecturaRegistraciones Then
            UltimaLectura = Format(LecturaAnterior, "dd/mm/yyyy hh:mm:ss")
            LecturaActual = Format(Now, "dd/mm/yyyy hh:mm:ss")
            TiempoEntreLecturas = DateDiff("n", UltimaLectura, LecturaActual)
            'Flog.writeline "*********************************************************"
            'Flog.writeline "Ultima Lectura: " & UltimaLectura
            'Flog.writeline "Lectura Actual: " & LecturaActual
            'Flog.writeline "Tiempo Entre Lecturas: " & TiempoEntreLecturas
            'Flog.writeline "*********************************************************"
        
            If TiempoEntreLecturas > TiempoDeLecturadeRegistraciones Then
                'Flog.writeline "Chequea Registraciones " & Format(Now, FormatoInternoFecha)
                'si hay alguno pendiente ==> no tiene sentido que inserte otro
                StrSql = "SELECT * FROM batch_proceso INNER JOIN Batch_tipproc ON batch_proceso.btprcnro = batch_tipproc.btprcnro WHERE bprcestado = 'Pendiente' and batch_proceso.btprcnro = 22"
                OpenRecordsetAS StrSql, objrsRegistraciones
                If objrsRegistraciones.EOF Then
                    ' no hay ==> veo si puedo lanzarlo
                    If PuedeEjecutar(0, 22) Then
                        ' insertar en batch_proceso un leer registraciones
                        StrSql = "INSERT INTO Batch_Proceso (btprcnro, bprcfecha, iduser, bprchora, bprcfecdesde, bprcfechasta, " & _
                                 "bprcestado, empnro) " & _
                                 "values (" & 22 & "," & ConvFecha(Date) & ", 'super'" & ",'" & Format(Now, "hh:mm:ss ") & "' " & _
                                 ", " & ConvFecha(Date) & ", " & ConvFecha(Date) & _
                                 ", 'Pendiente', 0)"
                        objConn.Execute StrSql, , adExecuteNoRecords
                    End If
                End If
            End If
        End If
        
        'FGZ -20/07/2007 - Planificador ----------------------------
        If UsaPlanificador Then
            UltimaPlanificacion = Format(PlanificacionAnterior, "dd/mm/yyyy hh:mm:ss")
            PlanificacionActual = Format(Now, "dd/mm/yyyy hh:mm:ss")
            TiempoEntrePlanificaciones = DateDiff("n", UltimaPlanificacion, PlanificacionActual)
            'Flog.writeline "*********************************************************"
            'Flog.writeline "Ultima Planificaci�n: " & UltimaPlanificacion
            'Flog.writeline "Planificaci�n Actual: " & PlanificacionActual
            'Flog.writeline "Tiempo Entre Planificaciones: " & TiempoEntrePlanificaciones
            'Flog.writeline "*********************************************************"
        
            If TiempoEntrePlanificaciones > TiempoDePlanificador Then
                If fs.fileexists(PathProcesos & "RHProPlan.exe") Then
                    'FGZ - Se agreg� un nuevo parametro al proceso (si tiene el string de conexion encriptado o no)
                    'pid = Shell(PathProcesos & "RHProPlan.exe" & " " & Etiqueta, vbHide)
                    Flog.writeline "Se obtenido el PID Parametros: " & PathProcesos & "RHProPlan.exe" & " " & Etiqueta & " " & c_seed
                    
                    pid = Shell(PathProcesos & "RHProPlan.exe" & " " & Etiqueta & " " & c_seed, vbHide)
                    
                    Flog.writeline "SHELL " & PathProcesos & "RHProPlan.exe" & " " & Etiqueta & " " & c_seed & " " & Format(Now, "dd/mm/yyyy hh:mm:ss")
                    If pid <> 0 Then
                        Flog.writeline "    Ejecutando Planificador ... PID = " & pid
                    End If
                Else
                    Flog.writeline "No se encuentra el Programa Asociado al Planificador: RHProPlan.exe  " & Format(Now, "dd/mm/yyyy hh:mm:ss")
                End If
                PlanificacionAnterior = Format(Now, FormatoInternoFecha)
            End If
        End If
        'FGZ -20/07/2007 - Planificador ----------------------------
        
        'Chequeo que no exista ninguno en estado procesando que que realmente no se este ejecutando
        'Flog.writeline "Monitorea " & Format(Now, FormatoInternoFecha)
        Call Monitor
      
        'Inicializo el valor del arreglo Pendientes
        Call InicializoPendientes
      
        'Flog.writeline "Busca Pendientes " & Format(Now, FormatoInternoFecha)
                 
        'Para evitar el problema de la hora (hh:mm y h:mm)
        StrSql = "SELECT * FROM batch_proceso INNER JOIN Batch_tipproc ON batch_proceso.btprcnro = batch_tipproc.btprcnro WHERE bprcestado = 'Pendiente' "
        StrSql = StrSql & " AND (bprcfecha <=" & ConvFecha(Date) & ")"
        'FGZ - 13/02/2009 --------
        If Trim(Proc_Hab) <> "0" Then   'Solo se debe levantar ciertos tipos de procesos
            StrSql = StrSql & " AND batch_proceso.btprcnro IN (" & Proc_Hab & ")"
        End If
        If Trim(Proc_NoHab) <> "0" Then   'Solo se debe levantar ciertos tipos de procesos
            StrSql = StrSql & " AND batch_proceso.btprcnro NOT IN (" & Proc_NoHab & ")"
        End If
        'FGZ - 13/02/2009 --------
        StrSql = StrSql & " ORDER BY  bprcurgente, bpronro"
        OpenRecordsetAS StrSql, objrsProcesosPendientes
        If objrsProcesosPendientes.EOF Then
            'Flog.writeline " STRSQL : " & StrSql
            'Flog.writeline "No Hay Pendientes " & Format(Now, FormatoInternoFecha)
        End If
        'Flog.writeline " STRSQL : " & StrSql
        
        'Si hay procesos pendientes y puedo correrlos entonces
        If Not objrsProcesosPendientes.EOF And PuedeEjecutarConcurrente() Then
            'Flog.writeline "Encontr� Pendientes " & Format(Now, FormatoInternoFecha)
            ' Ordeno los pendientes por alg�n criterio
            Ultimo = CalcularPesos
            Actual = 1
            seguir = True
            
            If objrsProcesosPendientes.State = adStateOpen Then objrsProcesosPendientes.Close
            HoraEntre1 = Now
            ' Trato de levantar todos lo procesos que puedo
            Do While (Actual <= Ultimo) And seguir
                If PuedeEjecutar(Pendientes(Actual).Proceso, Pendientes(Actual).TipoProceso) Then
                    If Format(Pendientes(Actual).Fecha, FormatoInternoFecha) = Format(Date, FormatoInternoFecha) Then
                        If Format(Pendientes(Actual).bprchora, FormatoInternoHora) <= Format(Time, FormatoInternoHora) Then
                            pid = EjecutarProceso(PathProcesos, Pendientes(Actual).NombreProceso & " ", Pendientes(Actual).Proceso, Actual)
                            If Pendientes(Actual).TipoProceso = 22 Then
                                LecturaAnterior = Format(Now, FormatoInternoFecha)
                            End If
                        Else
                            'FGZ - 03/01/2012
                            ' le saqu� los mensajes de log para que no crezca tanto el archivo de log
                            'Flog.writeline "No puede ejecutar el proceso " & Pendientes(Actual).Proceso & " de tipo " & Pendientes(Actual).TipoProceso
                            'Flog.writeline "Hora del proceso (" & Format(Pendientes(Actual).bprchora, FormatoInternoHora) & ") posterior a la hora actual del servidor (" & Format(Time, FormatoInternoHora) & ")"
                        End If
                    Else
                        pid = EjecutarProceso(PathProcesos, Pendientes(Actual).NombreProceso & " ", Pendientes(Actual).Proceso, Actual)
                        If Pendientes(Actual).TipoProceso = 22 Then
                            LecturaAnterior = Format(Now, FormatoInternoFecha)
                        End If
                    End If
                Else
                    Flog.writeline "No puede ejecutar el proceso " & Pendientes(Actual).Proceso & " de tipo " & Pendientes(Actual).TipoProceso
                    Flog.writeline "Ya hay un proceso incompatible corriendo "
                End If
                'Flog.writeline "Actual = Actual + 1 "
                Actual = Actual + 1
                'Flog.writeline "LOOP"
            Loop
        End If
        If objrsProcesosPendientes.State = adStateOpen Then objrsProcesosPendientes.Close
           
        'Flog.writeline "A Dormir " & Format(Now, FormatoInternoFecha)
        'A dormir por x segundos
        'TiempodeDormida = 10
        Sleep (TiempodeDormida * 1000)
           
        'Flog.writeline "Despierta " & Format(Now, FormatoInternoFecha)
        
        'Actualizo los procesos que terminaron de ejecutar
        Call ActualizarTerminaronSuEjecucion
        'Flog.writeline "Pas� por ActualizarTerminaronSuEjecucion " & Format(Now, FormatoInternoFecha)
        
        'Busco los procesos que pudieren estar colgados y si es as�, los termino y �los relanzo?
        HoraEntre2 = Format(Now, FormatoInternoFecha)
        Call BuscoProcesosColgados(HoraEntre1, HoraEntre2)
        'Flog.writeline "Pas� por BuscarProcesosColgados " & Format(Now, FormatoInternoFecha)
        
        'Actualizar los procesos que no responden
        Call EliminarProcesosNoResponden
        'Flog.writeline "Pas� por EliminarProcesosNoResponden " & Format(Now, FormatoInternoFecha)
        
        'FGZ - 14/10/2009 - Le agregu� el control sobre los procesos que nunca arrancan
        'Actualizar los procesos que no responden Arrancaron
        Call EliminarProcesosNoInician
        
        
        'Elimino los procesos marcados por el usuario para eliminar
        Call EliminarProcesosMarcados
        'Flog.writeline "Pas� por EliminarProcesosMarcados " & Format(Now, FormatoInternoFecha)
    
        'Flog.writeline "Otro ciclo " & Format(Now, FormatoInternoFecha)
        
        'Chequea si el nombre del archivo de log es el que corresponde
        Call ChequeaLog(fs, Nombre_Arch)
    Loop

    ''FGZ - 16/08/2013 -----------------
    ''Control de Alertas Activas y no pendientes.
    'Call ControlarAlertas

    'Cierro Todo
    If objrsProcesosPendientes.State = adStateOpen Then objrsProcesosPendientes.Close
    If objRs.State = adStateOpen Then objRs.Close
    If objConn.State = adStateOpen Then objConn.Close
Loop

If objrsProcesosPendientes.State = adStateOpen Then objrsProcesosPendientes.Close
If objRs.State = adStateOpen Then objRs.Close
If objConn.State = adStateOpen Then objConn.Close
Set objRs = Nothing
Set objrsProcesosPendientes = Nothing
Set objConn = Nothing

Exit Sub

ce:
' -------------------------------------------------------------------------------------
' FGZ 25/07/2003
' Este manejador de errores esta habilitado unicamente para controlar el archivo de log
' se ejecuta siempre y cuando el archivo de log no exista aun.
    Flog.writeline "RHProAppSrv detenido por Error ( " & Err.Description & " )"
    Flog.writeline "============================================================="
    GoTo Continuar
End Sub


Private Function ProcesosEjecutando(Usuario As String) As Boolean
Dim rs_proc As New ADODB.Recordset
    StrSql = "SELECT * FROM batch_proceso WHERE (iduser = '" & Usuario & "') AND (bprcestado = 'Procesando')"
    OpenRecordsetAS StrSql, rs_proc
    ProcesosEjecutando = Not rs_proc.EOF
End Function


Private Sub Monitor()
' Chequea que si un proceso que est� en tabla en estado de ejecuci�n
' realmente se est� ejecutando.
' Si no es as� lo pone en estado de error
    
    Dim rs As New ADODB.Recordset
    Dim pid
    Dim hProc As Long
    Dim nRet As Long
    Const fdwAccess = SYNCHRONIZE

    'Obtiene los procesos que figuran en estado de ejecuci�n
    ' 25/07/2003 FGZ
    ' se agreg� " ... OR bprcestado = 'Procesando'" para que
    'tambien mate los procesos que no responden que no estan en memoria
    
    StrSql = "SELECT * FROM batch_proceso WHERE (bprcestado = 'Procesando' OR bprcestado = 'No Responde' )"
    'strsql = strsql & " AND btprcnro <> 8 AND btprcnro <> 25 "
    'FGZ - 16/02/2009 --------
    If Trim(Proc_Hab) <> "0" Then   'Solo se debe levantar ciertos tipos de procesos
        StrSql = StrSql & " AND batch_proceso.btprcnro IN (" & Proc_Hab & ")"
    End If
    If Trim(Proc_NoHab) <> "0" Then   'Solo se debe levantar ciertos tipos de procesos
        StrSql = StrSql & " AND batch_proceso.btprcnro NOT IN (" & Proc_NoHab & ")"
    End If
    'FGZ - 16/02/2009 --------
    OpenRecordsetAS StrSql, rs
    
    Do While Not rs.EOF
        ' Obtengo el identificador de proceso del SO
        pid = 0 & rs!bprcpid
        
        'Verifico si existe un proceso con ese PID
        hProc = OpenProcess(fdwAccess, False, pid)
        
        ' Si no existe, actualizo el estado de la tabla batch_proceso
        If hProc = 0 Then
        
            StrSql = "UPDATE batch_proceso SET bprcestado = 'Error' WHERE bpronro = " & rs!bpronro
            objConn.Execute StrSql, , adExecuteNoRecords
            
            Flog.writeline "Proceso " & rs!bpronro & " Abortado Manualmente por Usuario " & Now

        End If
        rs.MoveNext
    Loop
    If rs.State = adStateOpen Then rs.Close
End Sub


Private Function ProcesosenEjecucion() As Boolean
Dim rs As New ADODB.Recordset
    StrSql = "SELECT * FROM batch_proceso WHERE (bprcestado = 'Procesando')"
    OpenRecordsetAS StrSql, rs
    ProcesosenEjecucion = Not rs.EOF
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
End Function

Private Function EjecutarProceso(path As String, Nombre As String, NroProc As Long, Actual As Integer) As Long
' Lanza un proceso y actualiza la tabla de procesos
Dim MiPid As Long
Dim MiIndice As Integer
Dim fs

Set fs = CreateObject("Scripting.FileSystemObject")

'Flog.writeline "Inicio Proceso:" & path & Nombre & " " & NroProc & " " & Format(Now, "dd/mm/yyyy hh:mm:ss")

If fs.fileexists(path & Nombre) Then
    ' Ejecuto y obtengo el pid
    
            'FGZ - 14/10/2009 -
            'Le cambio el estado al proceso ------------------------------------
            StrSql = "UPDATE batch_proceso SET bprcpid = " & MiPid & ",  bprchorainicioej = '" & Format(Time, FormatoInternoHora) & "', bprcfecinicioej = " & ConvFecha(Now) & ", bprcestado = 'Iniciando'" & _
            " WHERE bpronro = " & NroProc
            objConn.Execute StrSql, , adExecuteNoRecords
            Flog.writeline
            Flog.writeline "Cambio estado a Iniciando, Proceso " & NroProc
            'Le cambio el estado al proceso ------------------------------------
    
    
    Flog.writeline "SHELL " & path & Nombre & NroProc & " " & Etiqueta & " " & c_seed & " " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    MiPid = Shell(path & Nombre & NroProc & " " & Etiqueta & " " & c_seed, vbHide)
    
    'Flog.writeline "MiPid: " & MiPid
    
    If MiPid <> 0 Then
        If Actual <> -1 Then
            
            'Inserto en conjunto de procesos en ejecuci�n
            Call InsertoEjecutando(Actual, MiPid)
            Flog.writeline "PID = " & MiPid
            'Actualizo el estado de la tabla
'            StrSql = "UPDATE batch_proceso SET bprcpid = " & MiPid & _
'                " WHERE bpronro = " & NroProc
'            objConn.Execute StrSql, , adExecuteNoRecords
        End If
    
    End If
    EjecutarProceso = MiPid
Else
    Flog.writeline "No se encuentra el Programa Asociado al Proceso:" & Nombre & Format(Now, "dd/mm/yyyy hh:mm:ss")
    
    ' actualizo el estado del proceso
    StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Time, FormatoInternoHora) & "', bprcfecfinej = " & ConvFecha(Now) & ", bprcestado = 'Inexistente'" & _
    " WHERE bpronro = " & NroProc
    objConn.Execute StrSql, , adExecuteNoRecords
    
End If
End Function

Private Sub InsertoEjecutando(NroActual As Integer, P_pid As Long)
Dim i As Integer

    i = BuscarIndiceEjecutando
    Ejecutando(i).pid = P_pid
    Ejecutando(i).Proceso = Pendientes(NroActual).Proceso
    Ejecutando(i).Progreso = 0
    Ejecutando(i).HoraInicioEj = Format(Now, FormatoInternoHora)
    Ejecutando(i).HoraFinEj = Ejecutando(i).HoraInicioEj
    Ejecutando(i).Intentos = Ejecutando(i).Intentos + 1
    Ejecutando(i).Fecha = Format(Pendientes(NroActual).Fecha, "dd/mm/yyyy")
End Sub

Private Function BuscarIndiceEjecutando() As Integer
' Busca un �ndice de un elemento que est� vac�o
Dim i As Integer
Dim Continuo As Boolean

    Continuo = True
    i = 1
    Do While i <= UBound(Ejecutando) And Continuo
        If Ejecutando(i).Proceso = 0 Then
            Continuo = False
        Else
            i = i + 1
        End If
    Loop
    
    BuscarIndiceEjecutando = i
    
End Function
Private Function PuedeEjecutarConcurrente() As Boolean
Dim rsProcesos As New ADODB.Recordset
Dim CantProc As Integer

StrSql = "SELECT count(*) as cantidad FROM batch_proceso WHERE (bprcestado = 'Procesando')"
OpenRecordsetAS StrSql, rsProcesos
CantProc = rsProcesos("cantidad")

If CantProc >= MaxConcurrentes Then
    Flog.writeline "Estan corriendo el maximo Posible de Procesos" & Format(Now, "dd/mm/yyyy hh:mm:ss")
End If
PuedeEjecutarConcurrente = (CantProc < MaxConcurrentes)

End Function


Private Function PuedeEjecutar(nroproceso As Long, Tipo As Integer) As Boolean
Dim rsProcesos As New ADODB.Recordset
Dim CantProc As Integer
Dim Puede As Boolean

StrSql = "SELECT count(*) as cantidad FROM batch_proceso WHERE (bprcestado = 'Procesando' OR bprcestado = 'Iniciando')"
OpenRecordsetAS StrSql, rsProcesos
CantProc = rsProcesos("cantidad")

If CantProc < MaxConcurrentes Then
    Puede = Not HayIncompatibleCorriendo(Tipo, nroproceso)
Else
    Puede = False
End If

PuedeEjecutar = Puede
    
If rsProcesos.State = adStateOpen Then rsProcesos.Close
Set rsProcesos = Nothing
End Function


Private Function HayIncompatibleCorriendo(ByVal Tipo As Long, ByVal nroproceso As Long) As Boolean
Dim rsProcesos As New ADODB.Recordset
Dim Puede As Boolean
Dim Hay As Boolean
Dim cadena As String
Dim AuxIncompatible As String
Dim pos1 As Integer
Dim pos2 As Integer

' levantar los procesos incompatibles
StrSql = "SELECT btprcincompat FROM batch_tipproc WHERE btprcnro = " & Tipo
OpenRecordsetAS StrSql, rsProcesos

Hay = False
AuxIncompatible = ""
cadena = ""

If Not rsProcesos.EOF Then
    If Not IsNull(rsProcesos!btprcincompat) Then
        cadena = rsProcesos!btprcincompat
        
        If Len(cadena) >= 1 Then
            pos1 = 1
            pos2 = InStr(pos1, cadena, ",") - 1
            If pos2 > 0 Then
                AuxIncompatible = Mid(cadena, pos1, pos2)
                cadena = Mid(cadena, pos2 + 2, Len(cadena))
            Else
                AuxIncompatible = cadena
                cadena = ""
            End If
        End If
    End If
End If

Do While Trim(AuxIncompatible) <> "" And Not Hay
    Hay = HayOtro(CInt(AuxIncompatible), nroproceso)
    
    AuxIncompatible = ""
    ' siguiente tipo incompatible
    If Len(cadena) >= 1 Then
        pos1 = 1
        pos2 = InStr(pos1, cadena, ",") - 1
        If pos2 > 0 Then
            AuxIncompatible = Mid(cadena, pos1, pos2)
            cadena = Mid(cadena, pos2 + 2, Len(cadena))
        Else
            AuxIncompatible = cadena
            cadena = ""
        End If
    End If
Loop

HayIncompatibleCorriendo = Hay
If rsProcesos.State = adStateOpen Then rsProcesos.Close
Set rsProcesos = Nothing
End Function



Private Function HayOtro(Tipo As Integer, nroproceso As Long) As Boolean
Dim rsHay As New ADODB.Recordset
Dim Esta As Boolean
Dim rsEnEj As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim rsp As New ADODB.Recordset

Esta = False

' busco todos los proceso que estan corriendo
'StrSql = "SELECT * FROM Batch_Proceso WHERE btprcnro = " & Tipo & " AND (bprcestado = 'Procesando' OR bprcestado = 'No Responde')"
StrSql = "SELECT * FROM Batch_Proceso WHERE btprcnro = " & Tipo & " AND (bprcestado = 'Procesando' OR bprcestado = 'No Responde' OR bprcestado = 'Iniciando')"
OpenRecordsetAS StrSql, rsEnEj

'levanto los datos del proceso que quiero ejecutar
StrSql = "SELECT * FROM Batch_Proceso WHERE bpronro = " & nroproceso
OpenRecordsetAS StrSql, rs

If Not rs.EOF Then
    ' hay proceso ejecutando de tipo incompatibles
    ' entonces chequeo interseccion de rango de fechas y empleados
        Select Case rs!btprcnro
        Case 1, 2, 4, 22:
            Do While Not rsEnEj.EOF And Not Esta
                'Hay incompatibilidad si estan procesando los mismos empleados en las mismas fechas
                'reviso interseccion de fechas de Procesamiento
                If Not IsNull(rs!bprcfecdesde) And Not IsNull(rs!bprcfechasta) And Not IsNull(rsEnEj!bprcfecdesde) And Not IsNull(rsEnEj!bprcfechasta) Then
                    If EstaEnRangoDeFechas(rs!bprcfecdesde, rs!bprcfechasta, rsEnEj!bprcfecdesde, rsEnEj!bprcfechasta) Then
                        ' la interseccion es <> de vacio, ==> chequeo la interseccion de Empleados
                        StrSql = "SELECT ternro FROM Batch_empleado WHERE bpronro = " & nroproceso & " AND (ternro IN ( SELECT ternro FROM Batch_empleado WHERE bpronro = " & rsEnEj!bpronro & "))"
                        OpenRecordsetAS StrSql, rsHay
                                            
                        If Not rsHay.EOF Then
                            ' la interseccion no es vacia
                            Esta = True
                        End If
                    End If
                Else
                    StrSql = "SELECT ternro FROM Batch_empleado WHERE bpronro = " & nroproceso & " AND (ternro IN ( SELECT ternro FROM Batch_empleado WHERE bpronro = " & rsEnEj!bpronro & "))"
                    OpenRecordsetAS StrSql, rsHay
                                        
                    If Not rsHay.EOF Then
                        ' la interseccion no es vacia
                        Esta = True
                    End If
                    If rsHay.State = adStateOpen Then rsHay.Close
                End If
                rsEnEj.MoveNext
            Loop
        Case 5, 9:
            'hay incompatibilidad si estan corriendo el mismo procesode AP
            Do While Not rsEnEj.EOF And Not Esta
                Select Case rsEnEj!btprcnro
                Case 1, 2:
                    'Debo revisar
                    '   Si las fechas del proceso se intersectan con las fechas del proceso que quiero ejecutar
                    '   y en caso de que se intersecten si alguno de los empleados de ese proceso esta en el en el proceso que quiero ejecutar
                    StrSql = " SELECT gti_Procacum.gpanro, gpadesde,gpahasta From batch_proceso "
                    StrSql = StrSql & " INNER JOIN Batch_Procacum ON Batch_Procacum.bpronro = batch_proceso.bpronro "
                    StrSql = StrSql & " INNER JOIN gti_Procacum ON Batch_Procacum.gpanro = gti_Procacum.gpanro "
                    StrSql = StrSql & " WHERE batch_proceso.bpronro = " & nroproceso
                    OpenRecordsetAS StrSql, rsp
                    Do While Not rsp.EOF And Not Esta
                        If Not IsNull(rsEnEj!bprcfecdesde) And Not IsNull(rsEnEj!bprcfechasta) And Not IsNull(rsp!gpadesde) And Not IsNull(rsp!gpahasta) Then
                            If EstaEnRangoDeFechas(rsEnEj!bprcfecdesde, rsEnEj!bprcfechasta, rsp!gpadesde, rsp!gpahasta) Then
                                ' la interseccion es <> de vacio, ==> chequeo la interseccion de Empleados
                                StrSql = "SELECT ternro FROM gti_cab WHERE gpanro = " & rsp!gpanro & " AND (ternro IN ( SELECT ternro FROM Batch_empleado WHERE bpronro = " & rsEnEj!bpronro & "))"
                                OpenRecordsetAS StrSql, rsHay
                                If Not rsHay.EOF Then
                                    ' la interseccion no es vacia
                                    Esta = True
                                End If
                            End If
                        Else
                            StrSql = "SELECT ternro FROM gti_cab WHERE gpanro = " & rsp!gpanro & " AND (ternro IN ( SELECT ternro FROM Batch_empleado WHERE bpronro = " & rsEnEj!bpronro & "))"
                            OpenRecordsetAS StrSql, rsHay
                            If Not rsHay.EOF Then
                                Esta = True
                            End If
                            If rsHay.State = adStateOpen Then rsHay.Close
                        End If
                        rsp.MoveNext
                    Loop
                Case 4, 5, 9:
                    StrSql = " SELECT gti_Procacum.gpanro From batch_proceso "
                    StrSql = StrSql & " INNER JOIN Batch_Procacum ON Batch_Procacum.bpronro = batch_proceso.bpronro "
                    StrSql = StrSql & " INNER JOIN gti_Procacum ON Batch_Procacum.gpanro = gti_Procacum.gpanro "
                    StrSql = StrSql & " WHERE batch_proceso.bpronro = " & rsEnEj!bpronro
                    StrSql = StrSql & " AND gti_Procacum.gpanro IN ("
                    StrSql = StrSql & " SELECT gti_Procacum.gpanro From batch_proceso "
                    StrSql = StrSql & " INNER JOIN Batch_Procacum ON Batch_Procacum.bpronro = batch_proceso.bpronro "
                    StrSql = StrSql & " INNER JOIN gti_Procacum ON Batch_Procacum.gpanro = gti_Procacum.gpanro "
                    StrSql = StrSql & " WHERE batch_proceso.bpronro = " & nroproceso
                    StrSql = StrSql & " )"
                    OpenRecordsetAS StrSql, rsHay
                    If Not rsHay.EOF Then
                        Esta = True
                    End If
                Case Else
                    'Debo revisar
                    '   Si las fechas del proceso se intersectan con las fechas del proceso que quiero ejecutar
                    '   y en caso de que se intersecten si alguno de los empleados de ese proceso esta en el en el proceso que quiero ejecutar
                    StrSql = " SELECT gti_Procacum.gpanro, gpadesde,gpahasta From batch_proceso "
                    StrSql = StrSql & " INNER JOIN Batch_Procacum ON Batch_Procacum.bpronro = batch_proceso.bpronro "
                    StrSql = StrSql & " INNER JOIN gti_Procacum ON Batch_Procacum.gpanro = gti_Procacum.gpanro "
                    StrSql = StrSql & " WHERE batch_proceso.bpronro = " & nroproceso
                    OpenRecordsetAS StrSql, rsp
                    Do While Not rsp.EOF And Not Esta
                        If Not IsNull(rsEnEj!bprcfecdesde) And Not IsNull(rsEnEj!bprcfechasta) And Not IsNull(rsp!gpadesde) And Not IsNull(rsp!gpahasta) Then
                            If EstaEnRangoDeFechas(rsEnEj!bprcfecdesde, rsEnEj!bprcfechasta, rsp!gpadesde, rsp!gpahasta) Then
                                ' la interseccion es <> de vacio, ==> chequeo la interseccion de Empleados
                                StrSql = "SELECT ternro FROM gti_cab WHERE gpanro = " & rsp!gpanro & " AND (ternro IN ( SELECT ternro FROM Batch_empleado WHERE bpronro = " & rsEnEj!bpronro & "))"
                                OpenRecordsetAS StrSql, rsHay
                                If Not rsHay.EOF Then
                                    ' la interseccion no es vacia
                                    Esta = True
                                End If
                            End If
                        Else
                            StrSql = "SELECT ternro FROM gti_cab WHERE gpanro = " & rsp!gpanro & " AND (ternro IN ( SELECT ternro FROM Batch_empleado WHERE bpronro = " & rsEnEj!bpronro & "))"
                            OpenRecordsetAS StrSql, rsHay
                            If Not rsHay.EOF Then
                                Esta = True
                            End If
                            If rsHay.State = adStateOpen Then rsHay.Close
                        End If
                        rsp.MoveNext
                    Loop
                End Select
                rsEnEj.MoveNext
            Loop
        Case 221, 371: 'Procesos de Simulacion (Copiado y borrado de simulacion)
            'Si hay uno corriendo DEBE esperar
            Do While Not rsEnEj.EOF And Not Esta
                If (rsEnEj!btprcnro = 221 Or rsEnEj!btprcnro = 223 Or rsEnEj!btprcnro = 371) Then
                    Esta = True
                End If
                rsEnEj.MoveNext
            Loop
        Case 223:   'Simulacion
            Do While Not rsEnEj.EOF And Not Esta
                If (rsEnEj!btprcnro = 221 Or rsEnEj!btprcnro = 371) Then
                    Esta = True
                Else
                    'Hay incompatibilidad si estan procesando los mismos empleados en las mismas fechas
                    'reviso interseccion de fechas de Procesamiento
                    If Not IsNull(rs!bprcfecdesde) And Not IsNull(rs!bprcfechasta) And Not IsNull(rsEnEj!bprcfecdesde) And Not IsNull(rsEnEj!bprcfechasta) Then
                        If EstaEnRangoDeFechas(rs!bprcfecdesde, rs!bprcfechasta, rsEnEj!bprcfecdesde, rsEnEj!bprcfechasta) Then
                            ' la interseccion es <> de vacio, ==> chequeo la interseccion de Empleados
                            StrSql = "SELECT ternro FROM Batch_empleado WHERE bpronro = " & nroproceso & " AND (ternro IN ( SELECT ternro FROM Batch_empleado WHERE bpronro = " & rsEnEj!bpronro & "))"
                            OpenRecordsetAS StrSql, rsHay
                            If Not rsHay.EOF Then
                                ' la interseccion no es vacia
                                Esta = True
                            End If
                        End If
                    Else
                        StrSql = "SELECT ternro FROM Batch_empleado WHERE bpronro = " & nroproceso & " AND (ternro IN ( SELECT ternro FROM Batch_empleado WHERE bpronro = " & rsEnEj!bpronro & "))"
                        OpenRecordsetAS StrSql, rsHay
                        If Not rsHay.EOF Then
                            ' la interseccion no es vacia
                            Esta = True
                        End If
                        If rsHay.State = adStateOpen Then rsHay.Close
                    End If
                End If
                rsEnEj.MoveNext
            Loop
        Case Else:
            Do While Not rsEnEj.EOF And Not Esta
                'Hay incompatibilidad si estan procesando los mismos empleados en las mismas fechas
                'reviso interseccion de fechas de Procesamiento
                If Not IsNull(rs!bprcfecdesde) And Not IsNull(rs!bprcfechasta) And Not IsNull(rsEnEj!bprcfecdesde) And Not IsNull(rsEnEj!bprcfechasta) Then
                    If EstaEnRangoDeFechas(rs!bprcfecdesde, rs!bprcfechasta, rsEnEj!bprcfecdesde, rsEnEj!bprcfechasta) Then
                        ' la interseccion es <> de vacio, ==> chequeo la interseccion de Empleados
                        StrSql = "SELECT ternro FROM Batch_empleado WHERE bpronro = " & nroproceso & " AND (ternro IN ( SELECT ternro FROM Batch_empleado WHERE bpronro = " & rsEnEj!bpronro & "))"
                        OpenRecordsetAS StrSql, rsHay
                                            
                        If Not rsHay.EOF Then
                            ' la interseccion no es vacia
                            Esta = True
                        End If
                    End If
                Else
                    StrSql = "SELECT ternro FROM Batch_empleado WHERE bpronro = " & nroproceso & " AND (ternro IN ( SELECT ternro FROM Batch_empleado WHERE bpronro = " & rsEnEj!bpronro & "))"
                    OpenRecordsetAS StrSql, rsHay
                                        
                    If Not rsHay.EOF Then
                        ' la interseccion no es vacia
                        Esta = True
                    End If
                    If rsHay.State = adStateOpen Then rsHay.Close
                End If
                rsEnEj.MoveNext
            Loop
        End Select
End If

If rs.State = adStateOpen Then rs.Close
If rsHay.State = adStateOpen Then rsHay.Close
If rsEnEj.State = adStateOpen Then rsEnEj.Close
Set rs = Nothing
Set rsHay = Nothing
Set rsEnEj = Nothing

HayOtro = Esta
End Function




Private Function HayOtro_old(Tipo As Integer, nroproceso As Long) As Boolean
Dim rsHay As New ADODB.Recordset
Dim Esta As Boolean
Dim rsEnEj As New ADODB.Recordset
Dim rs As New ADODB.Recordset

Esta = False

' busco todos los proceso que estan corriendo
StrSql = "SELECT * FROM Batch_Proceso WHERE btprcnro = " & Tipo & " AND (bprcestado = 'Procesando' OR bprcestado = 'No Responde')"
OpenRecordsetAS StrSql, rsEnEj

'levanto los datos del proceso que quiero ejecutar
StrSql = "SELECT * FROM Batch_Proceso WHERE bpronro = " & nroproceso
OpenRecordsetAS StrSql, rs

If Not rs.EOF Then
    ' hay proceso ejecutando de tipo incompatibles
    ' entonces chequeo interseccion de rango de fechas y empleados
    Do While Not rsEnEj.EOF And Not Esta
        ' si hay algun carga registraciones ejecutando ==> no debo lanzar otro ni tampoco un prc30
'        If rsEnEj!btprcnro = 1 And rs!btprcnro = 22 Or rsEnEj!btprcnro = 22 And rs!btprcnro = 1 Or rsEnEj!btprcnro = 22 And rs!btprcnro = 22 Then
'            Esta = True
'        End If
        
''FGZ - 29/01/2004
'        If (rsEnEj!btprcnro = 1 And rs!btprcnro = 22) Or (rsEnEj!btprcnro = 22 And rs!btprcnro = 1) Then
'            Esta = True
'        End If
'
'        If (rsEnEj!btprcnro = 2 And rs!btprcnro = 22) Or (rsEnEj!btprcnro = 22 And rs!btprcnro = 2) Then
'            Esta = True
'        End If
'
'        If rsEnEj!btprcnro = 22 And rs!btprcnro = 22 Then
'            Esta = True
'        End If
'
'        If (rsEnEj!btprcnro = 1 And rs!btprcnro = 2) Or (rsEnEj!btprcnro = 2 And rs!btprcnro = 1) Then
'            Esta = True
'        End If
'
'        If rsEnEj!btprcnro = 2 And rs!btprcnro = 2 Then
'            Esta = True
'        End If
'
'        If rsEnEj!btprcnro = 1 And rs!btprcnro = 1 Then
'            Esta = True
'        End If
''FGZ - 29/01/2004

            'reviso interseccion de fechas de Procesamiento
            If Not IsNull(rs!bprcfecdesde) And Not IsNull(rs!bprcfechasta) And Not IsNull(rsEnEj!bprcfecdesde) And Not IsNull(rsEnEj!bprcfechasta) Then
                If EstaEnRangoDeFechas(rs!bprcfecdesde, rs!bprcfechasta, rsEnEj!bprcfecdesde, rsEnEj!bprcfechasta) Then
                    ' la interseccion es <> de vacio, ==> chequeo la interseccion de Empleados
                    StrSql = "SELECT ternro FROM Batch_empleado WHERE bpronro = " & nroproceso & " AND (ternro IN ( SELECT ternro FROM Batch_empleado WHERE bpronro = " & rsEnEj!bpronro & "))"
                    OpenRecordsetAS StrSql, rsHay
                                        
                    If Not rsHay.EOF Then
                        ' la interseccion no es vacia
                        Esta = True
                    End If
                End If
            Else
                StrSql = "SELECT ternro FROM Batch_empleado WHERE bpronro = " & nroproceso & " AND (ternro IN ( SELECT ternro FROM Batch_empleado WHERE bpronro = " & rsEnEj!bpronro & "))"
                OpenRecordsetAS StrSql, rsHay
                                    
                If Not rsHay.EOF Then
                    ' la interseccion no es vacia
                    Esta = True
                End If
                If rsHay.State = adStateOpen Then rsHay.Close
            End If
    
        rsEnEj.MoveNext
    Loop
End If

If rs.State = adStateOpen Then rs.Close
If rsHay.State = adStateOpen Then rsHay.Close
If rsEnEj.State = adStateOpen Then rsEnEj.Close
Set rs = Nothing
Set rsHay = Nothing
Set rsEnEj = Nothing

HayOtro_old = Esta
End Function



Private Function EstaEnRangoDeFechas(FD1 As Date, FH1 As Date, FD As Date, FH As Date)
' devuelve true si el rango (fechaDesde1--FechaDesde2) esta en el rango (fechahasta2--Fechahsta2)
Dim Esta As Boolean

Esta = False
'FGZ - 16/10/2012 ---------------------
FD = Format(FD, FormatoInternoFechaCorto)
FH = Format(FH, FormatoInternoFechaCorto)
FD1 = Format(FD1, FormatoInternoFechaCorto)
FH1 = Format(FH1, FormatoInternoFechaCorto)
'FGZ - 16/10/2012 ---------------------
If (FD <= FD1 And FD1 <= FH) Or (FD <= FH1 And FH1 <= FH) Or (FD1 <= FD And FD <= FH1) Then
    Esta = True
End If

EstaEnRangoDeFechas = Esta

End Function

Private Function CalcularPesos() As Integer
' ----------------------------------------------------------------
' Descripcion:  carga todos los procesos pendientes al arreglo
'               Devuelve la cantidad de procesos pendientes de ejecuci�n
' Fecha:
' Autor:        FGZ
' Ultima Mod:   FGZ - 10/08/2004
'               Se agreg� que tenga en cuenta el tipo de modelo
'               para los proceso de Liquidacion.
' ----------------------------------------------------------------
Dim P As Integer
Dim i As Integer

Dim rs_TipoLiquidador As New ADODB.Recordset

P = objrsProcesosPendientes.RecordCount
'FGZ - 21/11/2006
'Si la cantidad de registros que levanta es > maximo que maneja ==> da error
If P >= MaxPendientes Then
    P = MaxPendientes - 1
End If
For i = 1 To P
    Pendientes(i).Proceso = objrsProcesosPendientes!bpronro
    Pendientes(i).Peso = 1
    Pendientes(i).TipoProceso = objrsProcesosPendientes!btprcnro
    Pendientes(i).IdUser = objrsProcesosPendientes!IdUser
    Pendientes(i).bprchora = Format(objrsProcesosPendientes!bprchora, FormatoInternoHora)
    Pendientes(i).Fecha = Format(objrsProcesosPendientes!bprcfecha, FormatoInternoFecha)

    If objrsProcesosPendientes!btprcnro = 3 Then
        'Flog.writeline "Proceso de Liquidacion. "
        'Busco el tipo de modelo
        StrSql = " SELECT * FROM tipoliquidador "
        If Not IsNull(objrsProcesosPendientes!bprcTipoModelo) Then
            StrSql = StrSql & " WHERE tliqnro = " & objrsProcesosPendientes!bprcTipoModelo
        Else
            StrSql = StrSql & " WHERE tliqdefault = -1 "
        End If
        If rs_TipoLiquidador.State = adStateOpen Then rs_TipoLiquidador.Close
        OpenRecordsetAS StrSql, rs_TipoLiquidador
        
        If Not rs_TipoLiquidador.EOF Then
            'Ejecutable del modelo
            'Flog.writeline "Ejecutable del modelo " & rs_TipoLiquidador!tliqprog
            If Not IsNull(rs_TipoLiquidador!tliqprog) Then
                Pendientes(i).NombreProceso = rs_TipoLiquidador!tliqprog
            Else
                'Flog.writeline "Nombre del Ejecutable del Modelo en Null "
                'Flog.writeline "Ejecutable default " & objrsProcesosPendientes!btprcprog
                Pendientes(i).NombreProceso = objrsProcesosPendientes!btprcprog
            End If
        Else
            'Ejecutable default del modelo
            'Flog.writeline "No se encontr� el modelo de liquidacion "
            'Flog.writeline "Ejecutable default " & objrsProcesosPendientes!btprcprog
            Pendientes(i).NombreProceso = objrsProcesosPendientes!btprcprog
        End If
    Else
        Pendientes(i).NombreProceso = objrsProcesosPendientes!btprcprog
    End If
    objrsProcesosPendientes.MoveNext
Next i

CalcularPesos = P
'Flog.writeline "Peso = " & P

If rs_TipoLiquidador.State = adStateOpen Then rs_TipoLiquidador.Close
Set rs_TipoLiquidador = Nothing

End Function


Public Sub DepurarLogs(ByVal Accion As Long, ByVal Dias As Long)
' --------------------------------------------------------------
' Descripcion: Depura los archivos de logs.
' Autor: FGZ - 18/10/2012
' Parametros:
'               Accion:
'                               0: No depura
'                               1: mueve a carpeta Bk
'                               2: Elimina archivos
'               Dias
'                                x: Mantiene los logs de los ultimos X dias.
' Ultima modificacion:
' --------------------------------------------------------------
Dim fs, f
Dim CarpetaBK   'Carpeta de BK
Dim CArchivos   'Carpeta de archivos
Dim NArchivo    'Nombre del archivo
Dim FArchivo    'Fecha del Archivo
Dim archivo     'Obeto
Dim Folder      'Carpeta
Dim Carpeta     'Carpeta
Dim Aplica As Boolean

    

    Flog.writeline " Depuraci�n de logs ... "
    If Accion = 0 Then
        Flog.writeline "    Depuracion NO activa"
    Else
        'CarpetaBK = PathFLog & "bk\"
    
        Set fs = CreateObject("Scripting.FileSystemObject")
        Set Folder = fs.GetFolder(PathFLog)
        Set CArchivos = Folder.Files
        
        
        If Not CArchivos.Count = 0 Then
            Flog.writeline CArchivos.Count & " Archivos de logs encontrados "
        End If
        
        On Error Resume Next
        For Each archivo In CArchivos
            
            NArchivo = archivo.Name
            Set f = fs.getfile(PathFLog & "\" & NArchivo)
            'Flog.writeline "Archivo:  " & PathFLog & NArchivo
            Flog.writeline "Archivo1:  " & PathFLog & NArchivo
            FArchivo = Format(f.DateLastModified, FormatoInternoFechaCorto)
            Aplica = False
            If CDate(FArchivo) < CDate((Date - Dias)) Then
                Aplica = True
                CarpetaBK = PathFLog & Format(Year(FArchivo), "0000") & Format(Month(FArchivo), "00") & "\"
            End If
            ' Flog.writeline "Accion:" & Accion
            If Aplica Then
                Select Case Accion
                Case 1: 'Mueve
                    f.Move CarpetaBK & NArchivo
                    If Err.Number <> 0 Then
                        Err.Clear
                        Flog.writeline "La carpeta Destino no existe. Se crear�."
                        Set Carpeta = fs.CreateFolder(CarpetaBK)
                        f.Move CarpetaBK & NArchivo
                    End If
                Case 2: ' Elimina
                    'No implementada porque el metodo no funciona
                  '  f.DeleteFile archivo
                  Flog.writeline "borrando archivo:" & PathFLog & NArchivo
                    fs.DeleteFile (PathFLog & NArchivo)
                Case Else:
                    'No hace nada
                End Select
            End If
        Next
    End If
    Flog.writeline " Fin Depuracion de logs"

End Sub


Public Sub ChequeaLog(ByVal fs, Nombre_Arch As String)
Dim Nombre_Arch_Corresponde As String

If MultiplesLOGs Then
    Nombre_Arch_Corresponde = PathFLog & "RHProAppSrv " & Format(Date, "dd-mm-yyyy") & ".log"

    If Nombre_Arch_Corresponde <> Nombre_Arch Then
        Call AbrirArchivoLog(fs, Nombre_Arch)
    End If
Else
    If Format(C_Date(DiaAnterior), "ddmmyyyy") <> Format(C_Date(Date), "ddmmyyyy") Then
        'cambio de dia
        Call AbrirArchivoLog(fs, Nombre_Arch)
    End If
End If
End Sub


Public Sub AbrirArchivoLog(ByVal fs, Nombre_Arch As String)
Dim Nombre_Arch_Corresponde As String

If MultiplesLOGs Then 'Genera un archivo por dia
    Nombre_Arch_Corresponde = PathFLog & "RHProAppSrv " & Format(Date, "dd-mm-yyyy") & ".log"

    If Nombre_Arch_Corresponde = Nombre_Arch Then
        If fs.fileexists(Nombre_Arch) Then
            ' lo abro para agregar
            Set Flog = fs.OpenTextFile(Nombre_Arch, ForAppending, 0)
        Else
            ' no existe, entonces lo creo
            Set Flog = fs.CreateTextFile(Nombre_Arch, True)
        End If
    Else
        Flog.writeline "Fin. Cambia d�a RHProAppSrv " & Format(Now, "dd/mm/yyyy hh:mm:ss")
        Flog.Close
        FinDia = True

'        Nombre_Arch = Nombre_Arch_Corresponde
'        Set Flog = fs.CreateTextFile(Nombre_Arch, True)
'        Flog.writeline "Inicio RHProAppSrv " & Format(Now, "dd/mm/yyyy hh:mm:ss")

    End If
Else 'trabaja siempre con el mismo archivo
    'una vez por dia lo inicializa
    If fs.fileexists(Nombre_Arch) Then
        On Error Resume Next
        Flog.Close
        On Error GoTo 0
        Set Flog = fs.OpenTextFile(Nombre_Arch, ForWriting, 0)
    Else
        ' no existe, entonces lo creo
        Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    End If
    Flog.writeline "Inicio RHProAppSrv " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    DiaAnterior = Date
End If
End Sub

Private Sub ActualizarTerminaronSuEjecucion()
Dim i As Integer
Dim hProc As Long
' Actualizo el conjunto de procesos en ejecuci�n
' borrando los procesos ya procesados

Const fdwAccess = SYNCHRONIZE

For i = 1 To UBound(Ejecutando)
    If Ejecutando(i).Proceso <> 0 Then
        'Verifico si existe un proceso con ese PID
        hProc = OpenProcess(fdwAccess, False, Ejecutando(i).pid)
           
        ' Si no existe, actualizo el estado de la tabla batch_proceso
        If hProc = 0 Then ' ya no esta en memoria
            'Flog.writeline "El proceso " & Ejecutando(i).Proceso & " ya no est� en memoria." & Format(Now, FormatoInternoHora)
            Ejecutando(i).pid = 0
            Ejecutando(i).Proceso = 0
            Ejecutando(i).Progreso = 0
            Ejecutando(i).HoraInicioEj = Format(Time, FormatoInternoFecha)
            Ejecutando(i).HoraFinEj = Ejecutando(i).HoraInicioEj
        End If
    End If
Next i

End Sub

Private Sub EliminarProcesosMarcados()
Dim rsEj As New ADODB.Recordset
Dim Ok As Long

    StrSql = "SELECT * "
    StrSql = StrSql & " FROM  batch_proceso "
    'StrSql = StrSql & " WHERE bprcfecha   >= " & ConvFecha(C_Date(Date - 10))
    StrSql = StrSql & " WHERE bprcfecha   >= " & ConvFecha(C_Date(DateAdd("d", -10, Date)))
    StrSql = StrSql & " AND   bprcterminar = -1 "
    StrSql = StrSql & " AND   bprcestado  <> 'Abortado por Usuario'"
    'FGZ - 16/02/2009 --------
    If Trim(Proc_Hab) <> "0" Then   'Solo se debe levantar ciertos tipos de procesos
        StrSql = StrSql & " AND batch_proceso.btprcnro IN (" & Proc_Hab & ")"
    End If
    If Trim(Proc_NoHab) <> "0" Then   'Solo se debe levantar ciertos tipos de procesos
        StrSql = StrSql & " AND batch_proceso.btprcnro NOT IN (" & Proc_NoHab & ")"
    End If
    'FGZ - 16/02/2009 --------
    
    'Flog.writeline "        SQL: " & StrSql
    OpenRecordsetAS StrSql, rsEj
    If rsEj.EOF Then
        'Flog.writeline "       No hay procesos marcados para terminar"
    Else
        Flog.writeline "       Se encontraron proceso marcados para terminar.... procesando"
    End If
    Do While Not rsEj.EOF
        Flog.writeline "    Proceso " & rsEj!bpronro & " Abortado por Usuario " & Format(C_Date(Date), FormatoInternoFecha)
                    
        If Not IsNull(rsEj!bprcpid) Then
            Ok = ANULAR_PROCESO(rsEj!bprcpid)
        End If
        
        ' actualizo los datos del proceso
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Time, FormatoInternoHora) & "'"
        StrSql = StrSql & ", bprcfecfinej = " & ConvFecha(Date) & ""
        StrSql = StrSql & ", bprcestado = 'Abortado por Usuario'"
        StrSql = StrSql & ", bprcterminar = 0"
        StrSql = StrSql & " WHERE bpronro = " & rsEj!bpronro
        objConn.Execute StrSql, , adExecuteNoRecords
            
        rsEj.MoveNext
    Loop
    
    If rsEj.State = adStateOpen Then rsEj.Close
    Set rsEj = Nothing
End Sub


Private Sub EliminarProcesosNoResponden()
Dim rsEj As New ADODB.Recordset
Dim Ok As Long
Dim TerminarProceso As Boolean

    'TiempoDeEsperaNoResponde = 3
        
    'StrSql = "SELECT * FROM batch_proceso WHERE empnro = 0 AND bprcterminar = 0 and bprcestado = 'No Responde'"
    StrSql = "SELECT * FROM batch_proceso WHERE (bprcterminar = 0 OR bprcterminar is null) and bprcestado = 'No Responde'"
    'FGZ - 16/02/2009 --------
    If Trim(Proc_Hab) <> "0" Then   'Solo se debe levantar ciertos tipos de procesos
        StrSql = StrSql & " AND batch_proceso.btprcnro IN (" & Proc_Hab & ")"
    End If
    If Trim(Proc_NoHab) <> "0" Then   'Solo se debe levantar ciertos tipos de procesos
        StrSql = StrSql & " AND batch_proceso.btprcnro NOT IN (" & Proc_NoHab & ")"
    End If
    'FGZ - 16/02/2009 --------
    '" AND btprcnro <> 8 AND btprcnro <> 25 "
    OpenRecordsetAS StrSql, rsEj
    
    Do While Not rsEj.EOF
        'Flog.writeline "Abortando Proceso " & rsEj!bpronro & " porque No Responde ...."
        If IsNull(rsEj!bprcHoraFinEj) Then
            TerminarProceso = True
        Else
            If Abs(DateDiff("n", Format(Time, FormatoInternoHora), Format(rsEj!bprcHoraFinEj, FormatoInternoHora))) > TiempoDeEsperaNoResponde Then
                TerminarProceso = True
                'Flog.writeline "datediff " & Abs(DateDiff("n", Format(Time, FormatoInternoHora), Format(rsEj!bprcHoraFinEj, FormatoInternoHora)))
                'Flog.writeline "Terminar"
            Else
                TerminarProceso = False
                'Flog.writeline "datediff " & Abs(DateDiff("n", Format(Time, FormatoInternoHora), Format(rsEj!bprcHoraFinEj, FormatoInternoHora)))
                'Flog.writeline "NO Terminar"
            End If
        End If
        If TerminarProceso Then
            Flog.writeline "Proceso " & rsEj!bpronro & " Abortado porque No Responde" & Format(Now, FormatoInternoHora)
                        
            If Not IsNull(rsEj!bprcpid) Then
                Ok = ANULAR_PROCESO(rsEj!bprcpid)
            End If
            
            ' actualizo los datos del proceso
            StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Time, "hh:mm:ss") & "', bprcfecfinej = " & ConvFecha(Now) & ", bprcestado = 'Abortado'" & _
            " WHERE bpronro = " & rsEj!bpronro
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        rsEj.MoveNext
    Loop
    
    If rsEj.State = adStateOpen Then rsEj.Close
    Set rsEj = Nothing
End Sub


Private Sub EliminarProcesosNoInician()
Dim rsEj As New ADODB.Recordset
Dim Ok As Long
Dim TerminarProceso As Boolean

    TiempoDeEsperaNoInician = TiempoDeEsperaNoResponde
        
    StrSql = "SELECT * FROM batch_proceso WHERE (bprcterminar = 0 OR bprcterminar is null) and bprcestado = 'Iniciando'"
    StrSql = StrSql & " ORDER BY bpronro desc "
    OpenRecordsetAS StrSql, rsEj
    Do While Not rsEj.EOF
        
        'FGZ - 12/01/2010 - le modifiqu� este control porque debe buscar la hora en que el appsrv lo dispar� y no la del proceso que para este
        '                   punto puede que todavia no haya iniciado.
      If Ejecutando(BuscarIndice(rsEj!bpronro)).pid <> 0 Then
        
        If DateDiff("n", Format(Ejecutando(BuscarIndice(rsEj!bpronro)).HoraInicioEj, FormatoInternoHora), Format(Time, FormatoInternoHora)) > TiempoDeEsperaNoInician Then
                        
            'Si ya no esta en memoria
            If Ejecutando(BuscarIndice(rsEj!bpronro)).pid = 0 Then
                'FGZ - 29/12/2011 -------
                'TerminarProceso = false
                TerminarProceso = True
                Flog.writeline "Proceso " & rsEj!bpronro & " qued� en estado Iniciando y no esta en memoria. (Revisar ejecutable del proceso) " & Format(Now, FormatoInternoHora)
            Else
                TerminarProceso = True
                 Flog.writeline "entr� por ac�"
            End If

        Else
            TerminarProceso = False
        End If
      Else
        TerminarProceso = False
      End If
'        If IsNull(rsEj!bprchorainicioej) Then
'            TerminarProceso = True
'        Else
'            If Date > rsEj!bprcfecinicioej Then
'                TerminarProceso = True
'            Else
'                If Date = rsEj!bprcfecinicioej Then
'                    If DateDiff("n", Format(rsEj!bprchorainicioej, FormatoInternoHora), Format(Time, FormatoInternoHora)) > TiempoDeEsperaNoInician Then
'                        TerminarProceso = True
'                    Else
'                        TerminarProceso = False
'                    End If
'                End If
'            End If
'        End If
        
        
        If TerminarProceso Then
            'Flog.writeline "Proceso " & rsEj!bpronro & " Abortado porque No Inicia. (Posible problema de version) " & Format(Now, FormatoInternoHora)
            Flog.writeline "    Proceso abortado porque No Inicia. (Posible problema de version) " & Format(Now, FormatoInternoHora)
            If Not IsNull(rsEj!bprcpid) Then
                If Ejecutando(BuscarIndice(rsEj!bpronro)).pid <> 0 Then
                    Ok = ANULAR_PROCESO(rsEj!bprcpid)
                End If
            End If
            
            'actualizo los datos del proceso
            StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Time, "hh:mm:ss") & "', bprcfecfinej = " & ConvFecha(Now) & ", bprcestado = 'Abortado'" & _
            " WHERE bpronro = " & rsEj!bpronro
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        rsEj.MoveNext
    Loop
    
    If rsEj.State = adStateOpen Then rsEj.Close
    Set rsEj = Nothing
End Sub



Private Sub BuscoProcesosColgados(H1 As Date, H2 As Date)
' Busco los procesos que est�n colgados en memoria y actualizo su estado a "No Responde"

Dim rsEj As New ADODB.Recordset
Dim strBusco As String
Dim Ok As Long
Dim MiIndice As Integer
Dim pid
Dim hProc As Long
Const fdwAccess = SYNCHRONIZE

    'TiempoDeEsperaSinProgreso = 5
    
    StrSql = " SELECT bpronro,bprcprogreso,iduser,bprcpid "
    StrSql = StrSql & " FROM   batch_proceso "
    StrSql = StrSql & " WHERE  batch_proceso.bprcestado = 'Procesando' "
    'FGZ - 16/02/2009 --------
    If Trim(Proc_Hab) <> "0" Then   'Solo se debe levantar ciertos tipos de procesos
        StrSql = StrSql & " AND batch_proceso.btprcnro IN (" & Proc_Hab & ")"
    End If
    If Trim(Proc_NoHab) <> "0" Then   'Solo se debe levantar ciertos tipos de procesos
        StrSql = StrSql & " AND batch_proceso.btprcnro NOT IN (" & Proc_NoHab & ")"
    End If
    'FGZ - 16/02/2009 --------
    StrSql = StrSql & " ORDER BY bpronro desc "
    OpenRecordsetAS StrSql, rsEj
    
'    strsql = "SELECT bpronro,bprcprogreso,iduser,bprcpid " & _
'             "FROM   batch_proceso " & _
'             "WHERE  batch_proceso.bprcestado = 'Procesando' " & _
'             " AND btprcnro <> 8 AND btprcnro <> 25 " & _
'             "ORDER BY bpronro desc "
    
'    StrSql = "SELECT bpronro,bprcprogreso,iduser,bprcpid " & _
'             "FROM   batch_proceso " & _
'             "WHERE  batch_proceso.empnro     = 0 " & _
'             "AND    batch_proceso.bprcestado = 'Procesando' " & _
'             "ORDER BY bpronro desc "
    'Flog.writeline "Busco procesos en estado Procesando  - " & Format(C_Date(Now), FormatoInternoFecha)
    
    Do While Not rsEj.EOF
        'Flog.writeline "Encontr� Procesando  - " & rsEj!bpronro & Format(C_Date(Now), FormatoInternoFecha)
        Flog.writeline "Encontr� Procesando  - " & rsEj!bpronro & Format(Now, FormatoInternoFecha)
        MiIndice = BuscarIndice(rsEj!bpronro)
        
        If Ejecutando(MiIndice).Progreso = rsEj!bprcprogreso Then
           'Flog.writeline "No avanz� el progreso. espero"
           If DateDiff("n", Format(Ejecutando(MiIndice).HoraFinEj, FormatoInternoHora), Format(Time, FormatoInternoHora)) > TiempoDeEsperaSinProgreso Then
                Flog.writeline "No avanz� el progreso en " & TiempoDeEsperaSinProgreso & " minutos. Pone Proceso " & rsEj!bpronro & " en estado NO RESPONDE - " & Format(Now, FormatoInternoHora)
                ' si hace mas de 5 minutos que no avanza entonces ponemos su estado en No Responde
                StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Time, FormatoInternoHora) & "', bprcfecfinej = " & ConvFecha(Now) & ", bprcestado = 'No Responde'" & _
                " WHERE bpronro = " & rsEj!bpronro 'Ejecutando(MiIndice).Proceso
                objConn.Execute StrSql, , adExecuteNoRecords
                
                'FGZ - 01/11/2013 --------------------------------
                'Flog.writeline "SQL Actualizado: " & StrSql
                'FGZ - 01/11/2013 --------------------------------

            End If
        Else
            'Flog.writeline "Actualizo el progreso "
            ' hora y fecha del ultimo progreso detectado
            If IsNull(rsEj!bprcprogreso) Then
                'Flog.writeline "Proceso " & rsEj!bpronro & " con progreso en NULO "
                
                ' Obtengo el identificador de proceso del SO
                pid = 0 & rsEj!bprcpid
                
                'Verifico si existe un proceso con ese PID
                hProc = OpenProcess(fdwAccess, False, pid)
                
                ' Si no existe, actualizo el estado de la tabla batch_proceso
                If hProc = 0 Then
                    StrSql = "UPDATE batch_proceso SET bprcestado = 'Error' WHERE bpronro = " & rsEj!bpronro
                    objConn.Execute StrSql, , adExecuteNoRecords
                    
                    'Flog.writeline "Proceso abortado (no estaba en memoria) "
                    Call LimpioProceso(MiIndice)
                    
                    'Flog.writeline "Proceso " & rsEj!bpronro & " Abortado Manualmente por Usuario " & Format(Now, FormatoInternoHora)
                    Flog.writeline "Proceso " & rsEj!bpronro & " se aborta porque No est� en memoria. Estado --> ERROR " & Format(Now, FormatoInternoHora)
                Else
                    ' el progreso est� en nulo y no deberia ocurrir
                    ' lo pongo en estado "No Responde" con progreso en 0
                    StrSql = "UPDATE batch_proceso SET bprcprogreso = 0, bprcestado = 'No Responde'" & _
                    " WHERE bpronro = " & Ejecutando(MiIndice).Proceso
                    objConn.Execute StrSql, , adExecuteNoRecords
                    
                    Flog.writeline "Proceso a No Responde "
                    'FGZ - 01/11/2013 --------------------------------
                    Flog.writeline "SQL Actualizado: " & StrSql
                    'FGZ - 01/11/2013 --------------------------------
                    
                    Call LimpioProceso(MiIndice)
                End If
            Else
                'Flog.writeline "Proceso " & rsEj!bpronro & " indice : " & MiIndice & ", HoraFinEj: " & Format(Time, FormatoInternoFecha) & " - " & Format(C_Date(Now), FormatoInternoFecha)
                Ejecutando(MiIndice).Progreso = rsEj!bprcprogreso
                Ejecutando(MiIndice).HoraFinEj = Format(Time, FormatoInternoHora)
            End If
        End If
        
        rsEj.MoveNext
    Loop
    
    If rsEj.State = adStateOpen Then rsEj.Close
    Set rsEj = Nothing
End Sub

Public Sub LimpioProceso(ByVal Indice As Long)

    Ejecutando(Indice).pid = 0
    Ejecutando(Indice).Proceso = 0
    Ejecutando(Indice).Progreso = 0
    Ejecutando(Indice).HoraInicioEj = Format(Time, FormatoInternoHora)
    Ejecutando(Indice).HoraFinEj = Format(Time, FormatoInternoHora)
    Ejecutando(Indice).Intentos = 0
    Ejecutando(Indice).Fecha = Format(Date, FormatoInternoFecha)
End Sub

Public Function ANULAR_PROCESO(ByVal id As Long) As Long
' Variables para control del proceso
Dim hProcessId, hThreadId, hProcess As Long
Const fdwAccess = PROCESS_TERMINATE

hProcess = OpenProcess(fdwAccess, False, id)

TerminateProcess hProcess, 0

End Function



Private Function BuscarIndice(nroproceso As Long) As Integer
Dim i As Integer
Dim Continuo As Boolean
Dim Indice As Integer

Indice = 0
Continuo = True
i = 1
Do While i <= UBound(Ejecutando) And Continuo
    If Ejecutando(i).Proceso = nroproceso Then
        Indice = i
        Continuo = False
    End If
    i = i + 1
Loop

BuscarIndice = Indice
End Function

Private Function CantidadEjecutando() As Integer
Dim i As Integer
Dim Corte As Boolean

Corte = False
i = 0
Do While i <= MaxPendientes And Not Corte
    If Ejecutando(i).pid <> 0 Then
        i = i + 1
    Else
        Corte = True
    End If
Loop

CantidadEjecutando = i

End Function

Private Sub InicializoPendientes()
Dim i As Integer

i = 1

Do While i <= UBound(Pendientes)
    Pendientes(i).Proceso = 0
    Pendientes(i).Peso = 0
    Pendientes(i).NombreProceso = ""
    Pendientes(i).TipoProceso = 0
    Pendientes(i).IdUser = ""
    Pendientes(i).bprchora = ""
    Pendientes(i).Fecha = Format(C_Date(Date), FormatoInternoFecha)
    
    'EAM(v3.08) - Se comento la inicializaci�n del arreglo ejecuntando porque no se debe setear ya que da error de abortado por el usuario
'    Ejecutando(i).pid = 0
'    Ejecutando(i).Proceso = 0
'    Ejecutando(i).Progreso = 0
'    'FGZ - 12/01/2010 - Estaba mal inicializado las horas
'    'Ejecutando(i).HoraFinEj = Format(C_Date(Now), FormatoInternoFecha)
'    'Ejecutando(i).HoraInicioEj = Format(C_Date(Now), FormatoInternoFecha)
'    Ejecutando(i).HoraFinEj = Format(C_Date(Now), FormatoInternoHora)
'    Ejecutando(i).HoraInicioEj = Format(C_Date(Now), FormatoInternoHora)
'    Ejecutando(i).Intentos = 0
'    Ejecutando(i).Fecha = Format(C_Date(Date), FormatoInternoFecha)

    Ejecutando(i).pid = 0
    Ejecutando(i).Proceso = 0
    Ejecutando(i).Progreso = 0
    'FGZ - 12/01/2010 - Estaba mal inicializado las horas
    'Ejecutando(i).HoraFinEj = Format(C_Date(Now), FormatoInternoFecha)
    'Ejecutando(i).HoraInicioEj = Format(C_Date(Now), FormatoInternoFecha)
    Ejecutando(i).HoraFinEj = Format(C_Date(Now), FormatoInternoHora)
    Ejecutando(i).HoraInicioEj = Format(C_Date(Now), FormatoInternoHora)
    Ejecutando(i).Intentos = 0
    Ejecutando(i).Fecha = Format(C_Date(Date), FormatoInternoFecha)
    
    i = i + 1
Loop
End Sub

Private Sub InicializoPendientes_old()
Dim i As Integer
Dim Fin As Boolean

Fin = False
i = 1

While Not Fin And i <= UBound(Pendientes)
    If Pendientes(i).Proceso <> 0 Then
        Pendientes(i).Proceso = 0
        Pendientes(i).Peso = 0
        Pendientes(i).NombreProceso = ""
        i = i + 1
    Else
        Fin = True
    End If
    
Wend

End Sub


Public Sub SetarDefaults()
' carga las configuraciones basicas para los procesos
Const ForReading = 1
Const ForAppending = 8

Dim f, fs
Dim strline As String
Dim pos1 As Integer
Dim pos2 As Integer

    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.OpenTextFile(App.path & "\rhproappsrvDefaults.ini", ForReading, 0)
    
    ' minutos que espera el spool antes de abortar el proceso
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        TiempoDeEsperaNoResponde = Mid(strline, pos1, pos2 - pos1)
    End If
    
    ' minutos que espera el spool antes de poner al proceso que se esta ejecutando en estado de No Responde
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        TiempoDeEsperaSinProgreso = Mid(strline, pos1, pos2 - pos1)
    End If

    'Tiempo entre lectura y lectura
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        TiempoDeLecturadeRegistraciones = Mid(strline, pos1, pos2 - pos1)
    End If

    'Tiempo de Dormida del Spool
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        TiempodeDormida = Mid(strline, pos1, pos2 - pos1)
    End If

    'Variable booleana que maneja si se usa Lectura de Registraciones o no
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        UsaLecturaRegistraciones = CBool(Mid(strline, pos1, pos2 - pos1))
    End If

    'Maximo Nro de Procesos concurrentes
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        MaxConcurrentes = Mid(strline, pos1, pos2 - pos1)
    End If

    'FGZ - 22/11/2004
    'Genera multiples Archivos de LOG (uno por dia)
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        MultiplesLOGs = CBool(Mid(strline, pos1, pos2 - pos1))
    End If

    'Cantidad de Reintentos de Mensajeria
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        Cantidad_Reintentos = CLng(Mid(strline, pos1, pos2 - pos1))
    End If

    'Tiempo entre reintentos
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        Tiempo_Reintentos = CLng(Mid(strline, pos1, pos2 - pos1))
    End If

    'FGZ - 20/07/2007 - Se agregaron estos 2 parametros
    'Usa Planificador
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        UsaPlanificador = CBool(Mid(strline, pos1, pos2 - pos1))
    End If
    'Tiempo entre ejecuciones del planificador
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        TiempoDePlanificador = CLng(Mid(strline, pos1, pos2 - pos1))
    End If
    'FGZ - 20/07/2007 - Se agregaron estos 2 parametros
    
    
    'FGZ - 14/01/2009 - Se agreg� este parametro
    'Conexion encriptada
    EncriptStrconexion = False
    c_seed = "0"
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        EncriptStrconexion = CBool(Mid(strline, pos1, pos2 - pos1))
        If EncriptStrconexion Then
            c_seed = Mid(strline, pos1, pos2 - pos1)
            If EsNulo(c_seed) Then
                c_seed = "0"
            End If
        End If
    End If
    'FGZ - 14/01/2009 - Se agreg� este parametro
    
    
    
    'FGZ - 13/02/2009 - Se agregaron 2 parametros
    'Procesos Habilitados
    'Procesos No Habilitados
    
    'Procesos habilitados
    Proc_Hab = "0"
    '= [0] significa TODOS
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        Proc_Hab = Mid(strline, pos1, pos2 - pos1)
    End If
    
    'Procesos No habilitados
    Proc_NoHab = "0"
    '= [0] significa NINGUNO
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        Proc_NoHab = Mid(strline, pos1, pos2 - pos1)
    End If
    'FGZ - 13/02/2009 - Se agregaron 2 parametros
    
    'FGZ - 18/10/2012 -----------------------------------------
    '   Se agregaron 2 parametros mas para depuracion de logs
    
    'Accion de Depuracion
    AccionConLogs = 0 'Nada
    '= [1] significa Hacer BK (muevo los archivos a la carpeta BK)
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        AccionConLogs = Mid(strline, pos1, pos2 - pos1)
    End If
    
    'Dias de Depuracion
    DiasDepuracionLogs = 30 'Dias
    '= [0] significa que se mantienen solo los archivos del dia analizado
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        DiasDepuracionLogs = Mid(strline, pos1, pos2 - pos1)
        If DiasDepuracionLogs < 0 Then
            DiasDepuracionLogs = 30 'Dias
        End If
    End If
    'FGZ - 18/10/2012 -----------------------------------------
    
    f.Close
End Sub


Public Function EsNulo(ByVal Objeto) As Boolean
    If IsNull(Objeto) Then
        EsNulo = True
    Else
        If UCase(Objeto) = "NULL" Or UCase(Objeto) = "" Then
            EsNulo = True
        Else
            EsNulo = False
        End If
    End If
End Function



Public Sub OpenRecordsetAS(strSQLQuery As String, ByRef objRs As ADODB.Recordset, Optional lockType As LockTypeEnum = adLockReadOnly)
Dim Asunto As String
Dim Mensaje As String
Dim AuxErr


    NotifAnterior = Format(Format(C_Date(Date - 1), FormatoInternoFecha), "dd/mm/yyyy hh:mm:ss")

    On Error Resume Next

Continuar:
    On Error Resume Next
    If objRs.State <> adStateClosed Then
        If objRs.lockType <> adLockReadOnly Then objRs.UpdateBatch
        objRs.Close
    End If
    Err.Clear
    objRs.Open strSQLQuery, objConn, adOpenDynamic, lockType, adCmdText
    If CLng(Err.Number) <> 0 Then
        AuxErr = Err.Number
        Flog.writeline "Error. Causa:" & Err.Number & " - " & Err.Description
        If ((CLng(Err.Number) = 2147467259) Or (CLng(Err.Number) = -2147467259)) Or (CLng(Err.Number) = 3709) Then
            'Error de conxion
            If NotificaError And Not ErrorNotificado Then
                Asunto = "Servicio detenido por error"
                Mensaje = "Error: " & Err.Number & " - " & Err.Description & Chr(13) + Chr(10)
                            
                NotifActual = Format(Now, "dd/mm/yyyy hh:mm:ss")
                TiempoEntreNotif = DateDiff("n", NotifAnterior, NotifActual)
                If TiempoEntreNotif > 60 Then
                    Call NotificarError(Asunto, Mensaje)
                    NotifAnterior = NotifActual
                    ErrorNotificado = True
                End If
            End If
            'Bloque que queda intentando conectar cada X segundos
            Err.Number = AuxErr
            Do While Err.Number <> 0
                Err.Clear
                Flog.writeline "       Reintentando conexi�n en " & TiempodeDormida & " segundos."
                Sleep (TiempodeDormida * 1000)
                Flog.writeline
                Flog.writeline "Intentando establecer conexi�n nuevamente ..."
                If objConn.State <> adStateClosed Then objConn.Close
                OpenConnectionAS strconexion, objConn
            Loop
            If NotificaError Then
                'FGZ - 17/10/2014 ----------------------------
                If objConn.State = adStateOpen Then
                    Asunto = "Servicio reestablecido"
                    Mensaje = " Se reestableci� la conexi�n del servicio"
                    Call NotificarError(Asunto, Mensaje)
                    ErrorNotificado = False
                Else
                    'Nada
                End If
                'FGZ - 17/10/2014 ----------------------------
            End If
            GoTo Continuar
        Else
            Flog.writeline "Error. Causa:" & Err.Number & " - " & Err.Description
            If NotificaError And Not ErrorNotificado Then
                Asunto = "Servicio detenido por error"
                Mensaje = "Error: " & Err.Number & " - " & Err.Description & Chr(13) + Chr(10)
                
                NotifActual = Format(Now, "dd/mm/yyyy hh:mm:ss")
                TiempoEntreNotif = DateDiff("n", NotifAnterior, NotifActual)
                If TiempoEntreNotif > 60 Then
                    Call NotificarError(Asunto, Mensaje)
                    NotifAnterior = NotifActual
                    ErrorNotificado = True
                End If
            End If
            GoTo Continuar
        End If
    End If
    
    Cantidad_de_OpenRecordset = Cantidad_de_OpenRecordset + 1
    
    
    
    
    
'    If CLng(Err.Number) <> 0 Then
'        Flog.writeline "Error. Causa:" & Err.Number & " - " & Err.Description
'        Flog.writeline "       Sleep " & TiempodeDormida & " segundos."
'        Sleep (TiempodeDormida * 1000)
'        If ((CLng(Err.Number) = 2147467259) Or (CLng(Err.Number) = -2147467259)) Then
'            Flog.writeline
'            Flog.writeline "Intentando establecer conexi�n nuevamente ..."
'            If objConn.State <> adStateClosed Then objConn.Close
'            OpenConnectionAS strconexion, objConn
'        Else
'            Flog.writeline "Error. Causa:" & Err.Number & " - " & Err.Description
'        End If
'        If NotificaError Then
'            Call NotificarError("Error: " & Err.Number & " - " & Err.Description)
'        End If
'        Err.Clear
'        GoTo Continuar
'    End If
'    Cantidad_de_OpenRecordset = Cantidad_de_OpenRecordset + 1
    
End Sub


Public Sub OpenConnectionAS(strConnectionString As String, ByRef objConn As ADODB.Connection)
    Error_Encrypt = False
    On Error GoTo Manejador
    If Not Ya_Encripto Then
        If EncriptStrconexion Then
            'strconexion = Encrypt(c_seed, strconexion)
            strConnectionString = Decrypt(c_seed, strConnectionString)
            Ya_Encripto = True
        End If
    End If
    
    If objConn.State <> adStateClosed Then objConn.Close
    'objConn.CursorLocation = adUseServer
    objConn.CursorLocation = adUseClient
    'objconn.IsolationLevel =
    objConn.IsolationLevel = adXactCursorStability
    'objConn.IsolationLevel = adXactBrowse
    objConn.CommandTimeout = 120
    objConn.ConnectionTimeout = 60
    objConn.Open strConnectionString
'   If TipoBD = 2 Then
'       objConn.Execute "SET LOCK MODE TO WAIT 60"
'   End If
    Select Case TipoBD
    Case 2:
        objConn.Execute "SET LOCK MODE TO WAIT 60"
    Case 4:
    If Not EsNulo(SCHEMA) Then
        'objConn.Execute "ALTER SESSION SET NLS_SORT = BINARY"
        objConn.Execute "ALTER SESSION SET CURRENT_SCHEMA = " & SCHEMA
    End If
   Case Else
   End Select
Exit Sub
Manejador:
    If EncriptStrconexion Then
        Flog.writeline "No se puede establecer conexion. Revise el string de conexion configurado en el .ini. Posible problema de encriptaci�n."
    Else
        Flog.writeline "No se puede establecer conexion. Revise el string de conexion configurado en el .ini."
    End If
    Error_Encrypt = True
End Sub


Public Sub ControlarAlertas()
Dim rs As New ADODB.Recordset
Dim Ok As Long
Dim ListaAlertas
Dim arr
Dim Asunto As String
Dim Mensaje As String
Dim MensajeAle As String
Dim HayAlertas As Boolean

    ListaAlertas = "0"
    HayAlertas = False
    'MensajeAle = " " & Chr(13) + Chr(10)
    MensajeAle = "<tr><td>  </td></tr>"
    
    StrSql = "SELECT bprcparam FROM batch_proceso WHERE btprcnro = 8 AND bprcestado = 'Pendiente'"
    OpenRecordset StrSql, rs
    Do While Not rs.EOF
        If Not EsNulo(rs!bprcparam) Then
            arr = Split(rs!bprcparam, " ")
            ListaAlertas = ListaAlertas & "," & arr(0)
        End If
        rs.MoveNext
    Loop
        
    'Reviso si hay activas que no esten entre las pendientes
    StrSql = "SELECT alenro,aledes, aledesext FROM alertas WHERE aleactiva = -1"
    StrSql = StrSql & " AND alenro NOT IN (" & ListaAlertas & ")"
    OpenRecordset StrSql, rs
    Do While Not rs.EOF
        'Hay alertas
        HayAlertas = True
        'MensajeAle = MensajeAle & "Alerta " & rs!alenro & " " & rs!aledes & Chr(13) + Chr(10)
        
        MensajeAle = MensajeAle & "<tr><td> " & "Alerta " & rs!alenro & " " & rs!aledes & " </td></tr>"
        
        rs.MoveNext
    Loop
    
    If NotificaError And HayAlertas Then
        Flog.writeline "Notificacion de Planificaci�n de alertas" & Format(Now, FormatoInternoFecha)
    
        Asunto = "RHPRO - Warning de Planificaci�n"
        
        'Mensaje = "Hay alertas activas que no tienen planificacion activa. Por favor revise la configuracion" & Chr(13) + Chr(10)
        'Mensaje = Mensaje & " " & Chr(13) + Chr(10)
        'Mensaje = Mensaje & MensajeAle
        'Mensaje = Mensaje & "Appsrv (RHPro) " & Chr(13) + Chr(10)
        
        'Escribo el Mensaje a enviar
        Mensaje = "<TABLE cellpadding=" & Chr(34) & "0" & Chr(34) & "cellspacing=" & Chr(34) & "0" & Chr(34) & ">"
        Mensaje = Mensaje & "<TR><TH> Aviso Importante</TH></TR>"
        Mensaje = Mensaje & "<tr><td>&nbsp;</td></tr>"
        Mensaje = Mensaje & "<tr><td> Hay alertas activas que no tienen planificacion activa. Por favor revise la configuracion </td></tr>"
        Mensaje = Mensaje & "<tr><td>&nbsp;</td></tr>"
        Mensaje = Mensaje & MensajeAle
        'Mensaje = Mensaje & "<tr><td>  </td></tr>"
        Mensaje = Mensaje & "<tr><td>&nbsp;</td></tr>"
        Mensaje = Mensaje & "<tr><td>&nbsp;</td></tr>"
        Mensaje = Mensaje & "<tr><td>&nbsp;</td></tr>"
        Mensaje = Mensaje & "<tr><td align=" & Chr(34) & "right" & Chr(34) & ">RHPRO AppSrv</td></tr>"
        Mensaje = Mensaje & "</TABLE>"
        
        
        Call NotificarError(Asunto, Mensaje)
    End If
    
    
If rs.State = adStateOpen Then rs.Close
Set rs = Nothing

End Sub


