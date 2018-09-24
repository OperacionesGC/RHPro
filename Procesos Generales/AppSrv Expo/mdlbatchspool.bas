Attribute VB_Name = "mdlbatchspool"
'----------------------------------------
'Spooler de procesos
'Creado el 7/3/2003
'Alvaro Bayon
'Fernando Zwenger
'----------------------------------------

Option Explicit

Type TCelda
    Proceso As Long
    NombreProceso As String
    Peso As Single
    TipoProceso As Integer
    IdUser As String
    bprchora As String
End Type

Type TCeldaEj
    Proceso As Long
    Pid As Long
    Progreso As Single
    HoraInicioEj As Date
    HoraFinEj As Date
End Type

Const MaxPendientes = 1000
Const MaxConcurrentes = 3

Const ForReading = 1
Const ForAppending = 8

Global Pendientes(MaxPendientes) As TCelda
Global Ejecutando(MaxPendientes) As TCeldaEj

Private Const PROCESS_TERMINATE As Long = &H1
Private Const SYNCHRONIZE = &H100000

Private Declare Function OpenProcess Lib _
   "kernel32" (ByVal dwDesiredAccess As Long, _
   ByVal bInheritHandle As Long, _
   ByVal dwProcessID As Long) As Long

Private Declare Function TerminateProcess Lib _
   "kernel32" (ByVal hProcess As Long, _
   ByVal uExitCode As Long) As Long
   
Private Declare Function Sleep Lib _
   "kernel32" (ByVal dwMilliseconds As Long) As Long
   
Global objrsProcesosPendientes As New ADODB.Recordset
Global objrsRegistraciones As New ADODB.Recordset

Global Flog

' -------- Variables de control de tiempo ------------
' minutos que espera el spool antes de abortar el proceso
Global TiempoDeEsperaNoResponde As Integer

' minutos que espera el spool antes de poner al proceso que se esta ejecutando en estado de No Responde
Global TiempoDeEsperaSinProgreso As Integer

' Tiempo entre lectura y lectura
Global TiempoDeLecturadeRegistraciones As Integer

' Tiempo de Dormida del Spool
Global TiempodeDormida As Integer

' Variable booleana que maneja si se usa Lectura de Registraciones o no
Global UsaLecturaRegistraciones As Boolean
Global Etiqueta


Private Function ProcesosEjecutando(usuario As String) As Boolean
Dim rs_proc As New ADODB.Recordset
    StrSql = "SELECT * FROM batch_proceso WHERE (iduser = '" & usuario & "') AND (bprcestado = 'Procesando')"
    OpenRecordset StrSql, rs_proc
    ProcesosEjecutando = Not rs_proc.EOF
End Function


Private Sub Monitor()
' Chequea que si un proceso que está en tabla en estado de ejecución
' realmente se está ejecutando.
' Si no es así lo pone en estado de error
    
    Dim rs As New ADODB.Recordset
    Dim Pid
    Dim hproc As Long
    Dim nRet As Long
    Const fdwAccess = SYNCHRONIZE

    'Obtiene los procesos que figuran en estado de ejecución
    ' 25/07/2003 FGZ
    ' se agregó " ... OR bprcestado = 'Procesando'" para que
    'tambien mate los procesos que no responden que no estan en memoria
    
    StrSql = "SELECT * " & _
             "FROM batch_proceso " & _
             "WHERE (batch_proceso.empnro     = 0 AND " & _
             "       batch_proceso.bprcestado = 'Procesando') " & _
             "OR    (batch_proceso.empnro     = 0 AND " & _
             "       batch_proceso.bprcestado = 'No Responde')"
    OpenRecordset StrSql, rs
    
    Do While Not rs.EOF
        ' Obtengo el identificador de proceso del SO
        Pid = 0 & rs!bprcpid
        
        'Verifico si existe un proceso con ese PID
        hproc = OpenProcess(fdwAccess, False, Pid)
        
        ' Si no existe, actualizo el estado de la tabla batch_proceso
        If hproc = 0 Then
        
            StrSql = "UPDATE batch_proceso SET bprcestado = 'Error' WHERE bpronro = " & rs!bpronro
            objConn.Execute StrSql, , adExecuteNoRecords
            
            Flog.writeline "Proceso " & rs!bpronro & " Abortado Manualmente por Usuario " & Now

        End If
        rs.MoveNext
    Loop
    
End Sub


Private Function ProcesosenEjecucion() As Boolean
Dim rs As New ADODB.Recordset
    StrSql = "SELECT * " & _
             "FROM   batch_proceso " & _
             "WHERE  batch_proceso.empnro     = 0 " & _
             "AND    batch_proceso.bprcestado = 'Procesando'"
    OpenRecordset StrSql, rs
    ProcesosenEjecucion = Not rs.EOF
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
End Function

Private Function EjecutarProceso(path As String, Nombre As String, NroProc As Long, Actual As Integer) As Long
' Lanza un proceso y actualiza la tabla de procesos
Dim MiPid As Long
Dim MiIndice As Integer
Dim file

Set file = CreateObject("Scripting.FileSystemObject")

MiPid = 0
If file.fileexists(path & Nombre) Then
    Flog.writeline "Inicio Proceso:" & Nombre & " " & NroProc & " " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    
    ' Ejecuto y obtengo el pid
    MiPid = Shell(path & Nombre & NroProc, vbHide)

Else
    Flog.writeline "Proceso: " & Nombre & " inexistente " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    StrSql = "UPDATE batch_proceso SET bprcestado = 'Error'" & _
        " WHERE bpronro = " & NroProc
    objConn.Execute StrSql, , adExecuteNoRecords
End If

If MiPid <> 0 Then
    If Actual <> -1 Then
        'Inserto en conjunto de procesos en ejecución
        Call InsertoEjecutando(Actual, MiPid)
        Flog.writeline "PID = " & MiPid
        'Actualizo el estado de la tabla
        StrSql = "UPDATE batch_proceso SET bprcpid = " & MiPid & _
            " WHERE bpronro = " & NroProc
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
End If
EjecutarProceso = MiPid
    
End Function

Private Sub InsertoEjecutando(NroActual As Integer, P_pid As Long)
Dim i As Integer

    i = BuscarIndiceEjecutando
    Ejecutando(i).Pid = P_pid
    Ejecutando(i).Proceso = Pendientes(NroActual).Proceso
    Ejecutando(i).Progreso = 0
    Ejecutando(i).HoraInicioEj = Format(Now, "dd/mm/yyyy hh:mm:ss")
    Ejecutando(i).HoraFinEj = Ejecutando(i).HoraInicioEj

End Sub

Private Function BuscarIndiceEjecutando() As Integer
' Busca un índice de un elemento que esté vacío
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

StrSql = "SELECT count(*) AS cantidad " & _
         "FROM   batch_proceso " & _
         "WHERE  batch_proceso.empnro     = 0 " & _
         "AND    batch_proceso.bprcestado = 'Procesando'"
OpenRecordset StrSql, rsProcesos
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

StrSql = "SELECT count(*) as cantidad FROM batch_proceso WHERE batch_proceso.empnro     = 0 AND (batch_proceso.bprcestado = 'Procesando')"
OpenRecordset StrSql, rsProcesos
CantProc = rsProcesos("cantidad")

If CantProc < MaxConcurrentes Then
    Puede = Not HayIncompatibleCorriendo(Tipo, nroproceso)
Else
    Flog.writeline "Estan corriendo el maximo Posible de Procesos" & Format(Now, "dd/mm/yyyy hh:mm:ss")
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
Dim Cadena As String
Dim AuxIncompatible As String
Dim pos1 As Integer
Dim pos2 As Integer

' levantar los procesos incompatibles
StrSql = "SELECT btprcincompat FROM batch_tipproc WHERE btprcnro = " & Tipo
OpenRecordset StrSql, rsProcesos

Hay = False
AuxIncompatible = ""
Cadena = ""

If Not rsProcesos.EOF Then
    If Not IsNull(rsProcesos!btprcincompat) Then
        Cadena = rsProcesos!btprcincompat
        
        If Len(Cadena) >= 1 Then
            pos1 = 1
            pos2 = InStr(pos1, Cadena, ",") - 1
            If pos2 > 0 Then
                AuxIncompatible = Mid(Cadena, pos1, pos2)
                Cadena = Mid(Cadena, pos2 + 2, Len(Cadena))
            Else
                AuxIncompatible = Cadena
                Cadena = ""
            End If
        End If
    End If
End If

Do While AuxIncompatible <> "" And Not Hay
    Hay = HayOtro(CInt(AuxIncompatible), nroproceso)
    
    AuxIncompatible = ""
    ' siguiente tipo incompatible
    If Len(Cadena) >= 1 Then
        pos1 = 1
        pos2 = InStr(pos1, Cadena, ",") - 1
        If pos2 > 0 Then
            AuxIncompatible = Mid(Cadena, pos1, pos2)
            Cadena = Mid(Cadena, pos2 + 2, Len(Cadena))
        Else
            AuxIncompatible = Cadena
            Cadena = ""
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

Esta = False

' busco todos los proceso que estan corriendo
StrSql = "SELECT * " & _
         "FROM   batch_proceso " & _
         "WHERE (batch_proceso.empnro      = 0             AND " & _
         "       batch_proceso.bprcestado  = 'Procesando'  AND " & _
         "       batch_proceso.bprcfecha   >=" & ConvFecha(Date - 10) & " AND " & _
         "       batch_proceso.btprcnro    = " & Tipo & ")" & _
         "OR    (batch_proceso.empnro      = 0             AND " & _
         "       batch_proceso.bprcestado  = 'No Responde' AND " & _
         "       batch_proceso.bprcfecha   >=" & ConvFecha(Date - 10) & " AND " & _
         "       batch_proceso.btprcnro    = " & Tipo & ")"
'StrSql = "SELECT * " & _
'         "FROM  batch_Proceso " & _
'         "WHERE batch_proceso.empnro     = 0 " & _
'         "AND   batch_proceso.btprcnro   = " & tipo & _
'         "AND  (batch_proceso.bprcestado = 'Procesando' OR " & _
'         "      batch_proceso.bprcestado = 'No Responde')"
' Se modifica esta consulta para optimizar el acceso a la tabla. O.D.A. 17/10/2003
OpenRecordset StrSql, rsEnEj

' levanto los datos del proceso que quiero ejecutar
StrSql = "SELECT * FROM Batch_Proceso WHERE bpronro = " & nroproceso
OpenRecordset StrSql, rs

If Not rs.EOF Then
    ' hay proceso ejecutando de tipo incompatibles
    ' entonces chequeo interseccion de rango de fechas y empleados
    Do While Not rsEnEj.EOF And Not Esta
        ' si hay algun carga registraciones ejecutando ==> no debo lanzar otro ni tampoco un prc30
'        If rsEnEj!btprcnro = 1 And rs!btprcnro = 22 Or rsEnEj!btprcnro = 22 And rs!btprcnro = 1 Or rsEnEj!btprcnro = 22 And rs!btprcnro = 22 Then
'            Esta = True
'        End If
        
'FGZ - 29/01/2004
        If (rsEnEj!btprcnro = 1 And rs!btprcnro = 22) Or (rsEnEj!btprcnro = 22 And rs!btprcnro = 1) Then
            Esta = True
        End If
        
        If (rsEnEj!btprcnro = 2 And rs!btprcnro = 22) Or (rsEnEj!btprcnro = 22 And rs!btprcnro = 2) Then
            Esta = True
        End If
        
        If rsEnEj!btprcnro = 22 And rs!btprcnro = 22 Then
            Esta = True
        End If
        
        If (rsEnEj!btprcnro = 1 And rs!btprcnro = 2) Or (rsEnEj!btprcnro = 2 And rs!btprcnro = 1) Then
            Esta = True
        End If
        
        If rsEnEj!btprcnro = 2 And rs!btprcnro = 2 Then
            Esta = True
        End If
        
        If rsEnEj!btprcnro = 1 And rs!btprcnro = 1 Then
            Esta = True
        End If
        
'FGZ - 29/01/2004

        ' reviso interseccion de fechas de Procesamiento
        If Not IsNull(rs!bprcfecdesde) And Not IsNull(rs!bprcfechasta) And Not IsNull(rsEnEj!bprcfecdesde) And Not IsNull(rsEnEj!bprcfechasta) Then
            If EstaEnRangoDeFechas(rs!bprcfecdesde, rs!bprcfechasta, rsEnEj!bprcfecdesde, rsEnEj!bprcfechasta) Then
                ' la interseccion es <> de vacio, ==> chequeo la interseccion de Empleados
                StrSql = "SELECT ternro FROM Batch_empleado WHERE bpronro = " & nroproceso & " AND (ternro IN ( SELECT ternro FROM Batch_empleado WHERE bpronro = " & rsEnEj!bpronro & "))"
                OpenRecordset StrSql, rsHay
                                    
                If Not rsHay.EOF Then
                    ' la interseccion no es vacia
                    Esta = True
                End If
            End If
        Else
            StrSql = "SELECT ternro FROM Batch_empleado WHERE bpronro = " & nroproceso & " AND (ternro IN ( SELECT ternro FROM Batch_empleado WHERE bpronro = " & rsEnEj!bpronro & "))"
            OpenRecordset StrSql, rsHay
                                
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

HayOtro = Esta

End Function



Private Function EstaEnRangoDeFechas(FD1 As Date, FH1 As Date, FD As Date, FH As Date)
' devuelve true si el rango (fechaDesde1--FechaDesde2) esta en el rango (fechahasta2--Fechahsta2)
Dim Esta As Boolean

Esta = False

If (FD <= FD1 And FD1 <= FH) Or (FD <= FH1 And FH1 <= FH) Or (FD1 <= FD And FD <= FH1) Then
    Esta = True
End If

EstaEnRangoDeFechas = Esta

End Function

Private Function CalcularPesos() As Integer
' carga todos los procesos pendientes al arreglo
' Devuelve la cantidad de procesos pendientes de ejecución

Dim P As Integer
Dim i As Integer

P = objrsProcesosPendientes.RecordCount
For i = 1 To P
    Pendientes(i).Proceso = objrsProcesosPendientes!bpronro
    Pendientes(i).NombreProceso = objrsProcesosPendientes!btprcprog
    Pendientes(i).Peso = 1
    Pendientes(i).TipoProceso = objrsProcesosPendientes!btprcnro
    Pendientes(i).IdUser = objrsProcesosPendientes!IdUser
    Pendientes(i).bprchora = objrsProcesosPendientes!bprchora
    objrsProcesosPendientes.MoveNext
    
Next i

CalcularPesos = P
Flog.writeline "Peso = " & P
End Function

Public Sub ChequeaLog(ByVal fs, Nombre_Arch As String)
Dim Nombre_Arch_Corresponde As String

    Nombre_Arch_Corresponde = PathFLog & "RHProAppSrv " & Format(Date, "dd-mm-yyyy") & ".log"

    If Nombre_Arch_Corresponde <> Nombre_Arch Then
        Call AbrirArchivoLog(fs, Nombre_Arch)
    End If
End Sub


Public Sub AbrirArchivoLog(ByVal fs, Nombre_Arch As String)
Dim Nombre_Arch_Corresponde As String

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
        Flog.writeline "Fin. Cambia día RHProAppSrv " & Format(Now, "dd/mm/yyyy hh:mm:ss")
        Flog.Close

        Nombre_Arch = Nombre_Arch_Corresponde
        Set Flog = fs.CreateTextFile(Nombre_Arch, True)
        Flog.writeline "Inicio RHProAppSrv " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    End If

End Sub


Public Sub Main()
Dim Archivo As String
Dim fs, f
Dim strline As String
Dim tiposIncomp As String
Dim pos1 As Integer
Dim pos2 As Integer
'Dim path As String 'En esta variable va el path en que se encuentran los procesos
Dim Pid
Dim cerrado As Boolean

Dim Actual As Integer
Dim Ultimo As Integer
Dim seguir As Boolean

Dim HoraEntre1 As Date
Dim HoraEntre2 As Date
Dim Nombre_Arch As String
Dim LecturaAnterior

Dim rs_MuestraPendientes As New ADODB.Recordset


' carga las configuraciones basicas, formato de fecha, string de conexion,
' tipo de BD y ubicacion del archivo de log
Call CargarConfiguracionesBasicas
    
' Usa Lectura de Registraciones
UsaLecturaRegistraciones = False

'--------------------------------------------------------------------------------
' FGZ 25/07/2003
' Abre para append el archivo de log, si no lo encuentra ==> lo crea

    Nombre_Arch = PathFLog & "RHProAppSrv " & Format(Date, "dd-mm-yyyy") & ".log"
    
    ' Primero tendría que chequear si existe, si es asi lo abro para appending y sino lo creo
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    ' crea o abre el archivo de log, segun corresponda
    Call AbrirArchivoLog(fs, Nombre_Arch)
    
    ' Habilito el control de errores
    On Error GoTo CE
    
Continuar:
Flog.writeline "Inicio RHProAppSrv " & Format(Now, "dd/mm/yyyy hh:mm:ss")

OpenConnection strconexion, objConn

TiempoDeLecturadeRegistraciones = 1 ' minutos
LecturaAnterior = Format(Date - 1, "dd/mm/yyyy hh:mm:ss")

Do While True
    ' Chequea si el nombre del archivo de log es el que corresponde
    Call ChequeaLog(fs, Nombre_Arch)
    
    ' Acá tendria que lanzar el leer registraciones bajo dos condiciones
    ' que supere el tiempo preestablecido entre ejecuciones para este tipo de proceso
    ' que no haya otro leer registraciones ni prc30 ejecutandose
    
    If UsaLecturaRegistraciones Then
        If DateDiff("n", LecturaAnterior, Format(Now, "dd/mm/yyyy hh:mm:ss")) > TiempoDeLecturadeRegistraciones Then
            Flog.writeline "Chequea Registraciones " & Format(Now, "dd/mm/yyyy hh:mm:ss")
            ' FGZ 24/07/2003
            ' si hay alguno pendiente ==> no tiene sentido que inserte otro
            StrSql = "SELECT * " & _
                     "FROM batch_proceso " & _
                     "INNER JOIN Batch_tipproc " & _
                     "ON    batch_proceso.btprcnro = batch_tipproc.btprcnro " & _
                     "WHERE batch_proceso.empnro     = 0 " & _
                     "AND   batch_proceso.bprcestado = 'Pendiente' " & _
                     "AND   batch_proceso.btprcnro   = 22"
            OpenRecordset StrSql, objrsRegistraciones
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
    If objrsRegistraciones.State = adStateOpen Then objrsRegistraciones.Close
    
    'Chequeo que no exista ninguno en estado procesando que
    'que realmente no se este ejecutando
    Flog.writeline "Monitorea " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Call Monitor
  
    'Inicializo el valor del arreglo Pendientes
    InicializoPendientes
  
    Flog.writeline "Busca Pendientes " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    'Busco los procesos pendientes en la tabla de procesos
    
    ' FGZ - 11/03/2004 Para evitar el problema de la hora (hh:mm y h:mm)
    StrSql = "SELECT * FROM batch_proceso " & _
             " INNER JOIN Batch_tipproc ON batch_proceso.btprcnro = batch_tipproc.btprcnro " & _
             " WHERE batch_proceso.empnro = 0" & _
             " AND bprcestado = 'Pendiente' " & _
             " AND (bprcfecha <=" & ConvFecha(Date) & ")" & _
             " ORDER BY  bprcurgente, bpronro"
    OpenRecordset StrSql, objrsProcesosPendientes
   
    If objrsProcesosPendientes.EOF Then
        Flog.writeline "No Hay Pendientes " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    End If
    
    'Si hay procesos pendientes y puedo correrlos entonces
    If Not objrsProcesosPendientes.EOF And PuedeEjecutarConcurrente() Then
        Flog.writeline "Encontró Pendientes " & Format(Now, "dd/mm/yyyy hh:mm:ss")
        ' Ordeno los pendientes por algún criterio
        Ultimo = CalcularPesos
        Actual = 1
        seguir = True
        
        If objrsProcesosPendientes.State = adStateOpen Then objrsProcesosPendientes.Close
        HoraEntre1 = Now
        ' Trato de levantar todos lo procesos que puedo
        Do While (Actual <= Ultimo) And seguir
            If PuedeEjecutar(Pendientes(Actual).Proceso, Pendientes(Actual).TipoProceso) Then
                If Format(Pendientes(Actual).bprchora, "hh:mm:ss") <= Format(Now, "hh:mm:ss") Then
                    Pid = EjecutarProceso(PathProcesos, Pendientes(Actual).NombreProceso & " ", Pendientes(Actual).Proceso, Actual)
                    If Pendientes(Actual).TipoProceso = 22 Then
                        LecturaAnterior = Now
                    End If
                Else
                    Flog.writeline "No puede ejecutar el proceso por la hora " & Pendientes(Actual).Proceso & " de tipo " & Pendientes(Actual).TipoProceso
                    Flog.writeline "Hora del proceso:" & Format(Pendientes(Actual).bprchora, "hh:mm:ss") & ". Hora del servidor:" & Format(Now, "hh:mm:ss")
                End If
            Else
                Flog.writeline "No puede ejecutar el proceso " & Pendientes(Actual).Proceso & " de tipo " & Pendientes(Actual).TipoProceso
            End If
            Actual = Actual + 1
        Loop
    End If
    If objrsProcesosPendientes.State = adStateOpen Then objrsProcesosPendientes.Close
       
    Flog.writeline "A Dormir " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    ' A dormir por x segundos
    TiempodeDormida = 10
    Sleep (TiempodeDormida * 1000)
       
    Flog.writeline "Despierta " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    
    ' actualizo los procesos que terminaron de ejecutar
    Call ActualizarTerminaronSuEjecucion
    Flog.writeline "Pasó por ActualizarTerminaronSuEjecucion " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    
    ' Busco los procesos que pudieren estar colgados y si es así, los termino y ¿los relanzo?
    HoraEntre2 = Now
    Call BuscoProcesosColgados(HoraEntre1, HoraEntre2)
    Flog.writeline "Pasó por BuscarProcesosColgados " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    
    ' Actualizar los procesos que no responden
    Call EliminarProcesosNoResponden
    Flog.writeline "Pasó por EliminarProcesosNoResponden " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    
    ' Elimino los procesos marcados por el usuario para eliminar
    Call EliminarProcesosMarcados
    Flog.writeline "Pasó por EliminarProcesosMarcados " & Format(Now, "dd/mm/yyyy hh:mm:ss")

    Flog.writeline "Otro ciclo " & Format(Now, "dd/mm/yyyy hh:mm:ss")
Loop

Flog.writeline "RHProAppSrv detenido " & Format(Now, "dd/mm/yyyy hh:mm:ss")

Flog.Close

If objrsProcesosPendientes.State = adStateOpen Then objrsProcesosPendientes.Close
If objRs.State = adStateOpen Then objRs.Close
If objConn.State = adStateOpen Then objConn.Close
Set objRs = Nothing
Set objrsProcesosPendientes = Nothing
Set objConn = Nothing

Exit Sub

' -------------------------------------------------------------------------------------
' FGZ 25/07/2003
' Este manejador de errores esta habilitado unicamente para controlar el archivo de log
' se ejecuta siempre y cuando el archivo de log no exista aun.
CE:
       
    Flog.writeline "RHProAppSrv detenido por Error ( " & Err.Description & " )"
    Flog.writeline "============================================================="
    GoTo Continuar
End Sub



Private Sub ActualizarTerminaronSuEjecucion()
Dim i As Integer
Dim hproc As Long
' Actualizo el conjunto de procesos en ejecución
' borrando los procesos ya procesados

Const fdwAccess = SYNCHRONIZE

For i = 1 To UBound(Ejecutando)
    If Ejecutando(i).Proceso <> 0 Then
        'Verifico si existe un proceso con ese PID
        hproc = OpenProcess(fdwAccess, False, Ejecutando(i).Pid)
           
        ' Si no existe, actualizo el estado de la tabla batch_proceso
        If hproc = 0 Then ' ya no esta en memoria
            Ejecutando(i).Pid = 0
            Ejecutando(i).Proceso = 0
            Ejecutando(i).Progreso = 0
            Ejecutando(i).HoraInicioEj = Format(Now, "dd/mm/yyyy hh:mm:ss")
            Ejecutando(i).HoraFinEj = Ejecutando(i).HoraInicioEj
        End If
    End If
Next i

End Sub

Private Sub EliminarProcesosMarcados()
Dim rsEj As New ADODB.Recordset
Dim Ok As Long

    StrSql = "SELECT * " & _
             " FROM  batch_proceso" & _
             " WHERE bprcfecha   >=" & ConvFecha(Date - 10) & _
             " AND   empnro       = 0" & _
             " AND   bprcterminar = -1" & _
             " AND   bprcestado  <> 'Abortado por Usuario'"
    OpenRecordset StrSql, rsEj

    Do While Not rsEj.EOF
        Flog.writeline "Proceso " & rsEj!bpronro & " Abortado por Usuario " & Now
                    
        If Not IsNull(rsEj!bprcpid) Then
            Ok = ANULAR_PROCESO(rsEj!bprcpid)
        End If
        
        ' actualizo los datos del proceso
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss") & "', bprcfecfinej = " & ConvFecha(Now) & ", bprcestado = 'Abortado por Usuario'" & _
        " WHERE bpronro = " & rsEj!bpronro
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

    TiempoDeEsperaNoResponde = 3
        
    'StrSql = "SELECT * FROM batch_proceso WHERE empnro = 0 AND bprcterminar = 0 and bprcestado = 'No Responde'"
    StrSql = "SELECT * FROM batch_proceso WHERE empnro = 0 AND (bprcterminar = 0 OR bprcterminar is null) and bprcestado = 'No Responde'"
    OpenRecordset StrSql, rsEj
    
    Do While Not rsEj.EOF
        If IsNull(rsEj!bprcHoraFinEj) Then
            TerminarProceso = True
        Else
            If DateDiff("n", Format(Now, "hh:mm:ss"), Format(rsEj!bprcHoraFinEj, "hh:mm:ss")) > TiempoDeEsperaNoResponde Then
                TerminarProceso = True
            Else
                TerminarProceso = False
            End If
        End If
        If TerminarProceso Then
            Flog.writeline "Proceso " & rsEj!bpronro & " Abortado porque No Responde" & Now
                        
            If Not IsNull(rsEj!bprcpid) Then
                Ok = ANULAR_PROCESO(rsEj!bprcpid)
            End If
            
            ' actualizo los datos del proceso
            StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss") & "', bprcfecfinej = " & ConvFecha(Now) & ", bprcestado = 'Abortado'" & _
            " WHERE bpronro = " & rsEj!bpronro
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        rsEj.MoveNext
    Loop
    
    If rsEj.State = adStateOpen Then rsEj.Close
    Set rsEj = Nothing
End Sub


Private Sub BuscoProcesosColgados(H1 As Date, H2 As Date)
' Busco los procesos que están colgados en memoria y actualizo su estado a "No Responde"

Dim rsEj As New ADODB.Recordset
Dim strBusco As String
Dim Ok As Long
Dim MiIndice As Integer
Dim Pid
Dim hproc As Long
Const fdwAccess = SYNCHRONIZE

    TiempoDeEsperaSinProgreso = 5
    
    StrSql = "SELECT bpronro,bprcprogreso,iduser,bprcpid " & _
             "FROM   batch_proceso " & _
             "WHERE  batch_proceso.empnro     = 0 " & _
             "AND    batch_proceso.bprcestado = 'Procesando' " & _
             "ORDER BY bpronro desc "
    OpenRecordset StrSql, rsEj
    ' Agregado de la columna BPRCPID al SELECT. O.D.A. 23/02/2004
    
    Flog.writeline "Busco procesos en estado Procesando  - " & Now
    
    Do While Not rsEj.EOF
        Flog.writeline "Encontró Procesando  - " & rsEj!bpronro & Now
        MiIndice = BuscarIndice(rsEj!bpronro)
        
        If Ejecutando(MiIndice).Progreso = rsEj!bprcprogreso Then
           Flog.writeline "No avansó el progreso. espero"
           If DateDiff("n", Format(Ejecutando(MiIndice).HoraFinEj, "hh:mm:ss"), Format(Now, "hh:mm:ss")) > TiempoDeEsperaSinProgreso Then
                Flog.writeline "No avansó el progreso en 5 minutos. Pone Proceso " & rsEj!bpronro & " en estado NO RESPONDE - " & Now
                ' si hace mas de 5 minutos que no avanza entonces ponemos su estado en No Responde
                StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss") & "', bprcfecfinej = " & ConvFecha(Now) & ", bprcestado = 'No Responde'" & _
                " WHERE bpronro = " & Ejecutando(MiIndice).Proceso
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
        Else
            Flog.writeline "Actualizo el progreso "
            ' hora y fecha del ultimo progreso detectado
            If IsNull(rsEj!bprcprogreso) Then
                Flog.writeline "Proceso " & rsEj!bpronro & " con progreso en NULO "
                
                ' Obtengo el identificador de proceso del SO
                Pid = 0 & rsEj!bprcpid
                
                'Verifico si existe un proceso con ese PID
                hproc = OpenProcess(fdwAccess, False, Pid)
                
                ' Si no existe, actualizo el estado de la tabla batch_proceso
                If hproc = 0 Then
                    StrSql = "UPDATE batch_proceso SET bprcestado = 'Error' WHERE bpronro = " & rsEj!bpronro
                    objConn.Execute StrSql, , adExecuteNoRecords
                    
                    Flog.writeline "Proceso abortado (no estaba en memoria) "
                    Call LimpioProceso(MiIndice)
                    
                    Flog.writeline "Proceso " & rsEj!bpronro & " Abortado Manualmente por Usuario " & Now
                Else
                    ' el progreso está en nulo y no deberia ocurrir
                    ' lo pongo en estado "No Responde" con progreso en 0
                    StrSql = "UPDATE batch_proceso SET bprcprogreso = 0, bprcestado = 'No Responde'" & _
                    " WHERE bpronro = " & Ejecutando(MiIndice).Proceso
                    objConn.Execute StrSql, , adExecuteNoRecords
                    
                    Flog.writeline "Proceso a No Responde "
                    Call LimpioProceso(MiIndice)
                End If
            Else
                Flog.writeline "Proceso " & rsEj!bpronro & " indice : " & MiIndice & ", HoraFinEj: " & Format(Now, "dd/mm/yyyy hh:mm:ss") & " - " & Now
                Ejecutando(MiIndice).Progreso = rsEj!bprcprogreso
                Ejecutando(MiIndice).HoraFinEj = Format(Now, "hh:mm:ss")
            End If
        End If
        
        rsEj.MoveNext
    Loop
    
    If rsEj.State = adStateOpen Then rsEj.Close
    Set rsEj = Nothing
End Sub

Public Sub LimpioProceso(ByVal Indice As Long)

    Ejecutando(Indice).Pid = 0
    Ejecutando(Indice).Proceso = 0
    Ejecutando(Indice).Progreso = 0
    Ejecutando(Indice).HoraInicioEj = CStr(Now)
    Ejecutando(Indice).HoraFinEj = CStr(Now)

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
    If Ejecutando(i).Pid <> 0 Then
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
    
    Ejecutando(i).Pid = 0
    Ejecutando(i).Proceso = 0
    Ejecutando(i).Progreso = 0
    Ejecutando(i).HoraFinEj = Now
    Ejecutando(i).HoraInicioEj = Now
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

