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
End Type

Type TCeldaEj
    Proceso As Integer
    pid As Long
    Progreso As Single
    HoraInicioEj As Date
    HoraFinEj As Date
End Type

Const MaxPendientes = 100
Const MaxConcurrentes = 3

Const ForReading = 1
Const ForAppending = 8
Const FormatoInternoFecha = "dd/mm/yyyy hh:mm:ss"
Const FormatoInternoHora = "hh:mm:ss"

Global Pendientes(100) As TCelda
Global Ejecutando(100) As TCeldaEj

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
    Dim pid
    Dim hProc As Long
    Dim nRet As Long
    Const fdwAccess = SYNCHRONIZE

    'Obtiene los procesos que figuran en estado de ejecución
    ' 25/07/2003 FGZ
    ' se agregó " ... OR bprcestado = 'Procesando'" para que
    'tambien mate los procesos que no responden que no estan en memoria
    
    StrSql = "SELECT * FROM batch_proceso WHERE (bprcestado = 'Procesando' OR bprcestado = 'No Responde' )"
    OpenRecordset StrSql, rs
    
    Do While Not rs.EOF
        ' Obtengo el identificador de proceso del SO
        pid = 0 & rs!bprcpid
        
        'Verifico si existe un proceso con ese PID
        hProc = OpenProcess(fdwAccess, False, pid)
        
        ' Si no existe, actualizo el estado de la tabla batch_proceso
        If hProc = 0 Then
        
            StrSql = "UPDATE batch_proceso SET bprcestado = 'Error' WHERE bpronro = " & rs!bpronro
            objConn.Execute StrSql, , adExecuteNoRecords
            
            Flog.writeline "Proceso " & rs!bpronro & " Abortado Manualmente por Usuario " & Format(CDate(Now), FormatoInternoFecha)

        End If
        rs.MoveNext
    Loop
    
End Sub


Private Function ProcesosenEjecucion() As Boolean
Dim rs As New ADODB.Recordset
    StrSql = "SELECT * FROM batch_proceso WHERE (bprcestado = 'Procesando')"
    OpenRecordset StrSql, rs
    ProcesosenEjecucion = Not rs.EOF
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
End Function

Private Function EjecutarProceso(path As String, Nombre As String, NroProc As Long, Actual As Integer) As Long
' Lanza un proceso y actualiza la tabla de procesos
Dim MiPid As Long
Dim MiIndice As Integer

Flog.writeline "Inicio Proceso:" & Nombre & " " & NroProc & " " & Format(Now, "dd/mm/yyyy hh:mm:ss")

    ' Ejecuto y obtengo el pid
    MiPid = Shell(path & Nombre & NroProc, vbHide)
    If MiPid <> 0 Then
    
        If Actual <> -1 Then
            'Inserto en conjunto de procesos en ejecución
            Call InsertoEjecutando(Actual, MiPid)
            
            'Actualizo el estado de la tabla
            StrSql = "UPDATE batch_proceso SET bprcpid = '" & MiPid & _
                "' WHERE bpronro = " & NroProc
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
    
    End If
    EjecutarProceso = MiPid
    
End Function

Private Sub InsertoEjecutando(NroActual As Integer, P_pid As Long)
Dim i As Integer

    i = BuscarIndiceEjecutando
    Ejecutando(i).pid = P_pid
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

StrSql = "SELECT count(*) as cantidad FROM batch_proceso WHERE (bprcestado = 'Procesando')"
OpenRecordset StrSql, rsProcesos
CantProc = rsProcesos("cantidad")

PuedeEjecutarConcurrente = (CantProc < MaxConcurrentes)

End Function



Private Function PuedeEjecutar(nroproceso As Long, Tipo As Integer) As Boolean
Dim rsProcesos As New ADODB.Recordset
Dim CantProc As Integer
Dim Puede As Boolean

StrSql = "SELECT count(*) as cantidad FROM batch_proceso WHERE (bprcestado = 'Procesando')"
OpenRecordset StrSql, rsProcesos
CantProc = rsProcesos("cantidad")

If CantProc < MaxConcurrentes Then
    ' puedo ejecutar mas procesos
    Select Case Tipo
    Case 1: ' PRC30
        Puede = PuedeEjecutarPRC30(nroproceso)
    Case 2: ' PRC01
        Puede = PuedeEjecutarPRC01(nroproceso)
    Case 4: 'PRC06
        Puede = PuedeEjecutarPRC06(nroproceso)
    Case 5: 'ACUNOV
        Puede = PuedeEjecutarACUNOV(nroproceso)
    Case 8: 'Alertas
        Puede = PuedeEjecutarAlertas(nroproceso)
    Case 9: 'NovLiq
        Puede = PuedeEjecutarNOVLIQ(nroproceso)
    Case 10, 11, 12, 13, 14: 'Vacaciones
        Puede = PuedeEjecutarVACACIONES(nroproceso)
    Case 15: ' Exp SAP
        Puede = PuedeEjecutarEXPSAP(nroproceso)
    Case 16: ' DesgloceAD
        Puede = PuedeEjecutarDesgloceAD(nroproceso)
    Case 17: ' Reporte ASP 01
        Puede = PuedeEjecutarREPORTES(nroproceso)
    Case 18: ' Proceso Gerencial
        Puede = PuedeEjecutarGerencial(nroproceso)
    Case 19: ' Proceso Gerencial
        Puede = PuedeEjecutarTrimestral(nroproceso)
    Case 20: ' Feriados Nacionales
        Puede = PuedeEjecutarFeriados(nroproceso)
    Case 21: ' Proceso Semestral
        Puede = PuedeEjecutarSemestral(nroproceso)
    Case 22: ' Leer Registaciones
        Puede = PuedeEjecutarLeerRegistraciones(nroproceso)
    Case 23: ' Mensajeria
        Puede = True
    Case Else
        Puede = False
    End Select
Else
    Puede = False
End If

PuedeEjecutar = Puede
    
If rsProcesos.State = adStateOpen Then rsProcesos.Close
Set rsProcesos = Nothing
End Function


Private Function PuedeEjecutarPRC30(nroproceso As Long) As Boolean
Dim rsProcesos As New ADODB.Recordset
Dim Puede As Boolean
Dim Cadena As String
Dim AuxIncompatible As String
Dim pos1 As Integer
Dim pos2 As Integer


' levantar los procesos incompatibles
StrSql = "SELECT btprcincompat FROM Batch_tipproc WHERE btprcnro = 1"
OpenRecordset StrSql, rsProcesos

Puede = True
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

Do While AuxIncompatible <> "" And Puede
    If HayOtro(CInt(AuxIncompatible), nroproceso) Then
            Puede = False
    End If
    
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

PuedeEjecutarPRC30 = Puede
If rsProcesos.State = adStateOpen Then rsProcesos.Close
Set rsProcesos = Nothing
End Function

Private Function PuedeEjecutarPRC01(nroproceso As Long) As Boolean
Dim rsProcesos As New ADODB.Recordset
Dim Puede As Boolean
Dim Cadena As String
Dim AuxIncompatible As String
Dim pos1 As Integer
Dim pos2 As Integer


' levantar los procesos incompatibles
StrSql = "SELECT btprcincompat FROM batch_tipproc WHERE btprcnro = 2"
OpenRecordset StrSql, rsProcesos

Puede = True
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

Do While AuxIncompatible <> "" And Puede
    If HayOtro(CInt(AuxIncompatible), nroproceso) Then
            Puede = False
    End If
    
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

PuedeEjecutarPRC01 = Puede
If rsProcesos.State = adStateOpen Then rsProcesos.Close
Set rsProcesos = Nothing
End Function

Private Function PuedeEjecutarPRC06(nroproceso As Long) As Boolean
Dim rsProcesos As New ADODB.Recordset
Dim Puede As Boolean
Dim Cadena As String
Dim AuxIncompatible As String
Dim pos1 As Integer
Dim pos2 As Integer


' levantar los procesos incompatibles
StrSql = "SELECT btprcincompat FROM batch_tipproc WHERE btprcnro = 4"
OpenRecordset StrSql, rsProcesos

Puede = True
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

Do While AuxIncompatible <> "" And Puede
    If HayOtro(CInt(AuxIncompatible), nroproceso) Then
            Puede = False
    End If
    
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

PuedeEjecutarPRC06 = Puede
If rsProcesos.State = adStateOpen Then rsProcesos.Close
Set rsProcesos = Nothing
End Function

Private Function PuedeEjecutarACUNOV(nroproceso As Long) As Boolean
Dim rsProcesos As New ADODB.Recordset
Dim Puede As Boolean
Dim Cadena As String
Dim AuxIncompatible As String
Dim pos1 As Integer
Dim pos2 As Integer


' levantar los procesos incompatibles
StrSql = "SELECT btprcincompat FROM batch_tipproc WHERE btprcnro = 5"
OpenRecordset StrSql, rsProcesos

Puede = True
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

Do While AuxIncompatible <> "" And Puede
    If HayOtro(CInt(AuxIncompatible), nroproceso) Then
            Puede = False
    End If
    
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

PuedeEjecutarACUNOV = Puede
If rsProcesos.State = adStateOpen Then rsProcesos.Close
Set rsProcesos = Nothing
End Function

Private Function PuedeEjecutarEXPSAP(nroproceso As Long) As Boolean
Dim Puede As Boolean

' NADA LO FRENA
Puede = True
PuedeEjecutarEXPSAP = Puede

End Function

Private Function PuedeEjecutarGerencial(nroproceso As Long) As Boolean
Dim Puede As Boolean

' NADA LO FRENA
Puede = True
PuedeEjecutarGerencial = Puede

End Function


Private Function PuedeEjecutarNOVLIQ(nroproceso As Long) As Boolean
Dim rsProcesos As New ADODB.Recordset
Dim Puede As Boolean
Dim Cadena As String
Dim AuxIncompatible As String
Dim pos1 As Integer
Dim pos2 As Integer


' levantar los procesos incompatibles
StrSql = "SELECT btprcincompat FROM batch_tipproc WHERE btprcnro = 9"
OpenRecordset StrSql, rsProcesos

Puede = True
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

Do While AuxIncompatible <> "" And Puede
    If HayOtro(CInt(AuxIncompatible), nroproceso) Then
            Puede = False
    End If
    
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

PuedeEjecutarNOVLIQ = Puede
If rsProcesos.State = adStateOpen Then rsProcesos.Close
Set rsProcesos = Nothing
End Function

Private Function PuedeEjecutarVACACIONES(nroproceso As Long) As Boolean
Dim Puede As Boolean

Puede = True
PuedeEjecutarVACACIONES = Puede

End Function
Private Function PuedeEjecutarDesgloceAD(nroproceso As Long)
Dim Puede As Boolean

    Puede = True
    PuedeEjecutarDesgloceAD = Puede
End Function


Private Function PuedeEjecutarREPORTES(nroproceso As Long) As Boolean
Dim rsProcesos As New ADODB.Recordset
Dim Puede As Boolean
Dim Cadena As String
Dim AuxIncompatible As String
Dim pos1 As Integer
Dim pos2 As Integer


' levantar los procesos incompatibles
StrSql = "SELECT btprcincompat FROM batch_tipproc WHERE btprcnro = 17"
OpenRecordset StrSql, rsProcesos

Puede = True
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

Do While AuxIncompatible <> "" And Puede
    If HayOtro(CInt(AuxIncompatible), nroproceso) Then
            Puede = False
    End If
    
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

PuedeEjecutarREPORTES = Puede
If rsProcesos.State = adStateOpen Then rsProcesos.Close
Set rsProcesos = Nothing

End Function

Private Function PuedeEjecutarFeriados(nroproceso As Long) As Boolean
Dim rsProcesos As New ADODB.Recordset
Dim Puede As Boolean
Dim Cadena As String
Dim AuxIncompatible As String
Dim pos1 As Integer
Dim pos2 As Integer


' levantar los procesos incompatibles
StrSql = "SELECT btprcincompat FROM batch_tipproc WHERE btprcnro = 20"
OpenRecordset StrSql, rsProcesos

Puede = True
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

Do While AuxIncompatible <> "" And Puede
    If HayOtro(CInt(AuxIncompatible), nroproceso) Then
            Puede = False
    End If
    
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

PuedeEjecutarFeriados = Puede
If rsProcesos.State = adStateOpen Then rsProcesos.Close
Set rsProcesos = Nothing

End Function
Private Function PuedeEjecutarLeerRegistraciones(nroproceso As Long) As Boolean
Dim rsProcesos As New ADODB.Recordset
Dim Puede As Boolean
Dim Cadena As String
Dim AuxIncompatible As String
Dim pos1 As Integer
Dim pos2 As Integer


' levantar los procesos incompatibles
StrSql = "SELECT btprcincompat FROM batch_tipproc WHERE btprcnro = 22"
OpenRecordset StrSql, rsProcesos

Puede = True
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

Do While AuxIncompatible <> "" And Puede
    If HayOtro(CInt(AuxIncompatible), nroproceso) Then
            Puede = False
    End If
    
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

PuedeEjecutarLeerRegistraciones = Puede
If rsProcesos.State = adStateOpen Then rsProcesos.Close
Set rsProcesos = Nothing

End Function


Private Function PuedeEjecutarTrimestral(nroproceso As Long) As Boolean
Dim rsProcesos As New ADODB.Recordset
Dim Puede As Boolean
Dim Cadena As String
Dim AuxIncompatible As String
Dim pos1 As Integer
Dim pos2 As Integer


' levantar los procesos incompatibles
StrSql = "SELECT btprcincompat FROM batch_tipproc WHERE btprcnro = 19"
OpenRecordset StrSql, rsProcesos

Puede = True
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

Do While AuxIncompatible <> "" And Puede
    If HayOtro(CInt(AuxIncompatible), nroproceso) Then
            Puede = False
    End If
    
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

PuedeEjecutarTrimestral = Puede
If rsProcesos.State = adStateOpen Then rsProcesos.Close
Set rsProcesos = Nothing
End Function

Private Function PuedeEjecutarSemestral(nroproceso As Long) As Boolean
Dim rsProcesos As New ADODB.Recordset
Dim Puede As Boolean
Dim Cadena As String
Dim AuxIncompatible As String
Dim pos1 As Integer
Dim pos2 As Integer


' levantar los procesos incompatibles
StrSql = "SELECT btprcincompat FROM batch_tipproc WHERE btprcnro = 21"
OpenRecordset StrSql, rsProcesos

Puede = True
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

Do While AuxIncompatible <> "" And Puede
    If HayOtro(CInt(AuxIncompatible), nroproceso) Then
            Puede = False
    End If
    
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

PuedeEjecutarSemestral = Puede
If rsProcesos.State = adStateOpen Then rsProcesos.Close
Set rsProcesos = Nothing
End Function


Private Function HayOtro_Old(Tipo As Integer, ProcesoNro As Long) As Boolean
Dim rsHay As New ADODB.Recordset
Dim rsEstaEJ As New ADODB.Recordset
Dim Esta As Boolean

' busco todos los empleados del proceso que se quiere ejecutar
StrSql = "SELECT * FROM Batch_Proceso INNER JOIN Batch_empleado ON Batch_proceso.bpronro = Batch_Empleado.bpronro " _
& " WHERE Batch_Proceso.bpronro = " & ProcesoNro
OpenRecordset StrSql, rsEstaEJ

Esta = False
Do While Not rsEstaEJ.EOF And Not Esta
    ' todos los empleados de procesos del mismo tipo que estan corriendo
    StrSql = "SELECT * FROM Batch_Proceso INNER JOIN Batch_empleado ON Batch_proceso.bpronro = Batch_Empleado.bpronro " _
    & " WHERE Batch_Proceso.bprcestado = 'Procesando' AND Batch_Proceso.btprcnro = " & Tipo & " ORDER BY Batch_Proceso.bpronro"
    OpenRecordset StrSql, rsHay
    
    Do While Not rsHay.EOF And Not Esta
        ' interseccion de conjuntos de empleados en fechas
        If rsHay!Ternro = rsEstaEJ!Ternro Then
            If EstaEnRangoDeFechas(rsHay!fecdesde, rsHay!fechasta, rsEstaEJ!fecdesde, rsEstaEJ!fechasta) Then
                Esta = True
            End If
        End If
            
        ' siguiente legajo de los procesos que estan corriendo de ese tipo de procesos
        rsHay.MoveNext
    Loop
    
    ' siguiente legajo del proceso que se quiere correr
    rsEstaEJ.MoveNext
Loop

If rsHay.State = adStateOpen Then rsHay.Close
If rsEstaEJ.State = adStateOpen Then rsEstaEJ.Close
Set rsHay = Nothing
Set rsEstaEJ = Nothing

HayOtro_Old = Esta

End Function


Private Function HayOtro(Tipo As Integer, nroproceso As Long) As Boolean
Dim rsHay As New ADODB.Recordset
Dim Esta As Boolean
Dim rsEnEj As New ADODB.Recordset
Dim rs As New ADODB.Recordset

Esta = False

' busco todos los proceso que estan corriendo
StrSql = "SELECT * FROM Batch_Proceso WHERE btprcnro = " & Tipo & " AND (bprcestado = 'Procesando' OR bprcestado = 'No Responde')"
OpenRecordset StrSql, rsEnEj

' levanto los datos del proceso que quiero ejecutar
StrSql = "SELECT * FROM Batch_Proceso WHERE bpronro = " & nroproceso
OpenRecordset StrSql, rs

If Not rs.EOF Then
    ' hay proceso ejecutando de tipo incompatibles
    ' entonces chequeo interseccion de rango de fechas y empleados
    Do While Not rsEnEj.EOF And Not Esta
        ' si hay algun carga registraciones ejecutando ==> no debo lanzar otro ni tampoco un prc30
        If rsEnEj!btprcnro = 1 And rs!btprcnro = 22 Or rsEnEj!btprcnro = 22 And rs!btprcnro = 1 Or rsEnEj!btprcnro = 22 And rs!btprcnro = 22 Then
            Esta = True
        End If
        
        
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
    objrsProcesosPendientes.MoveNext
    
Next i

CalcularPesos = P
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
        Flog.writeline "Fin. Cambia día RHProAppSrv " & Format(CDate(Now), FormatoInternoFecha)
        Flog.Close

        Nombre_Arch = Nombre_Arch_Corresponde
        Set Flog = fs.CreateTextFile(Nombre_Arch, True)
        Flog.writeline "Inicio RHProAppSrv " & Format(CDate(Now), FormatoInternoFecha)
    End If

End Sub


Public Sub Main()
Dim Archivo As String
Dim fs, f
Dim strline As String
Dim tiposIncomp As String
Dim pos1 As Integer
Dim pos2 As Integer
Dim path As String 'En esta variable va el path en que se encuentran los procesos
Dim pid
Dim cerrado As Boolean

Dim Actual As Integer
Dim Ultimo As Integer
Dim seguir As Boolean

Dim HoraEntre1 As Date
Dim HoraEntre2 As Date
Dim Nombre_Arch As String
Dim LecturaAnterior

' carga las configuraciones basicas, formato de fecha, string de conexion,
' tipo de BD y ubicacion del archivo de log
Call CargarConfiguracionesBasicas
    
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
Flog.writeline "Inicio RHProAppSrv " & Format(CDate(Now), FormatoInternoFecha)

' FGZ 25/07/2003
'--------------------------------------------------------------------------------

    'Crea el archivo de log
' ----------
    'Nombre_Arch = PathFLog & "batchspool " & Format(Date, "dd-mm-yyyy") & ".log"
    'Set fs = CreateObject("Scripting.FileSystemObject")
    'Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
'--------------------------------------------------------------------------------
    'Abre el archivo INI
    path = ""
    'Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.OpenTextFile(App.path & "\batchspool.INI", ForReading, 0)
    
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        path = Mid(strline, pos1, pos2 - pos1)
        If Right(path, 1) <> "\" Then path = path & "\"
    End If
    f.Close
'--------------------------------------------------------------------------------
  OpenConnection strconexion, objConn

TiempoDeLecturadeRegistraciones = 10 ' minutos
'LecturaAnterior = Date - 1
LecturaAnterior = Format(CDate(Date - 1), FormatoInternoFecha)

Do While True
    ' Chequea si el nombre del archivo de log es el que corresponde
    Call ChequeaLog(fs, Nombre_Arch)
    
    ' Acá tendria que lanzar el leer registraciones bajo dos condiciones
    ' que supere el tiempo preestablecido entre ejecuciones para este tipo de proceso
    ' que no haya otro leer registraciones ni prc30 ejecutandose
    
    If DateDiff("n", LecturaAnterior, Format(CDate(Now), FormatoInternoFecha)) > TiempoDeLecturadeRegistraciones Then
        Flog.writeline "Chequea Registraciones " & Format(CDate(Now), FormatoInternoFecha)
        ' FGZ 24/07/2003
        ' si hay alguno pendiente ==> no tiene sentido que inserte otro
        StrSql = "SELECT * FROM batch_proceso INNER JOIN Batch_tipproc ON batch_proceso.btprcnro = batch_tipproc.btprcnro WHERE bprcestado = 'Pendiente' and batch_proceso.btprcnro = 22"
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
    
    'Chequeo que no exista ninguno en estado procesando que
    'que realmente no se este ejecutando
    Flog.writeline "Monitorea " & Format(CDate(Now), FormatoInternoFecha)
    Call Monitor
  
    'Inicializo el valor del arreglo Pendientes
    InicializoPendientes
  
    Flog.writeline "Busca Pendientes " & Format(CDate(Now), FormatoInternoFecha)
    'Busco los procesos pendientes en la tabla de procesos
    StrSql = "SELECT * FROM batch_proceso INNER JOIN Batch_tipproc ON batch_proceso.btprcnro = batch_tipproc.btprcnro WHERE bprcestado = 'Pendiente' " & _
         " AND (bprcfecha < " & ConvFecha(Date) & " OR ( bprcfecha = " & ConvFecha(Date) & " AND bprchora < '" & Format(Now, "hh:mm:ss ") & "'))" & _
         " ORDER BY  bprcurgente, bpronro"

    OpenRecordset StrSql, objrsProcesosPendientes
   
    'Si hay procesos pendientes y puedo correrlos entonces
    If Not objrsProcesosPendientes.EOF And PuedeEjecutarConcurrente() Then
        Flog.writeline "Encontró Pendientes " & Format(CDate(Now), FormatoInternoFecha)
        ' Ordeno los pendientes por algún criterio
        Ultimo = CalcularPesos
        Actual = 1
        seguir = True
        
        HoraEntre1 = Format(CDate(Now), FormatoInternoFecha)
        ' Trato de levantar todos lo procesos que puedo
        Do While (Actual <= Ultimo) And seguir
            If PuedeEjecutar(Pendientes(Actual).Proceso, Pendientes(Actual).TipoProceso) Then
                pid = EjecutarProceso(path, Pendientes(Actual).NombreProceso & " ", Pendientes(Actual).Proceso, Actual)
                If Pendientes(Actual).TipoProceso = 22 Then
                    LecturaAnterior = Format(CDate(Now), FormatoInternoFecha)
                End If
            End If
            Actual = Actual + 1
        Loop
    End If
    
    Flog.writeline "A Dormir " & Format(CDate(Now), FormatoInternoFecha)
    ' A dormir por x segundos
    TiempodeDormida = 10
    Sleep (TiempodeDormida * 1000)
       
    Flog.writeline "Despierta " & Format(CDate(Now), FormatoInternoFecha)
    
    ' actualizo los procesos que terminaron de ejecutar
    Call ActualizarTerminaronSuEjecucion
    Flog.writeline "Pasó por ActualizarTerminaronSuEjecucion " & Format(CDate(Now), FormatoInternoFecha)
    
    ' Busco los procesos que pudieren estar colgados y si es así, los termino y ¿los relanzo?
    'HoraEntre2 = DateAdd("s", Segundos, HoraEntre1)
    HoraEntre2 = Format(CDate(Now), FormatoInternoFecha)
    Call BuscoProcesosColgados(HoraEntre1, HoraEntre2)
    Flog.writeline "Pasó por BuscarProcesosColgados " & Format(CDate(Now), FormatoInternoFecha)
    
    ' Actualizar los procesos que no responden
    Call EliminarProcesosNoResponden
    Flog.writeline "Pasó por EliminarProcesosNoResponden " & Format(CDate(Now), FormatoInternoFecha)
    
    ' Elimino los procesos marcados por el usuario para eliminar
    Call EliminarProcesosMarcados
    Flog.writeline "Pasó por EliminarProcesosMarcados " & Format(CDate(Now), FormatoInternoFecha)

    Flog.writeline "Otro ciclo " & Format(CDate(Now), FormatoInternoFecha)
Loop

Flog.writeline "RHProAppSrv detenido " & Format(CDate(Now), FormatoInternoFecha)

Flog.Close

If objrsProcesosPendientes.State = adStateOpen Then objrsProcesosPendientes.Close
If objRs.State = adStateOpen Then objRs.Close
Set objRs = Nothing
If objConn.State = adStateOpen Then objConn.Close
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
Dim hProc As Long
' Actualizo el conjunto de procesos en ejecución
' borrando los procesos ya procesados

Const fdwAccess = SYNCHRONIZE

For i = 1 To UBound(Ejecutando)
    
    If Ejecutando(i).Proceso <> 0 Then
        'Verifico si existe un proceso con ese PID
        hProc = OpenProcess(fdwAccess, False, Ejecutando(i).pid)
           
        ' Si no existe, actualizo el estado de la tabla batch_proceso
        If hProc = 0 Then ' ya no esta en memoria
            
            Ejecutando(i).pid = 0
            Ejecutando(i).Proceso = 0
            Ejecutando(i).Progreso = 0
            Ejecutando(i).HoraInicioEj = Format(CDate(Now), FormatoInternoHora)
            Ejecutando(i).HoraFinEj = Format(CDate("00:00:00"), FormatoInternoHora)
            
        End If
    End If

Next i

End Sub

Private Sub EliminarProcesosMarcados()
Dim rsEj As New ADODB.Recordset
Dim Ok As Long

    StrSql = "SELECT * FROM batch_proceso WHERE bprcterminar = -1 AND bprcestado <> 'Abortado por Usuario'"
    OpenRecordset StrSql, rsEj
    
    Do While Not rsEj.EOF
        Flog.writeline "Proceso " & rsEj!bpronro & " Abortado por Usuario " & Format(CDate(Now), FormatoInternoFecha)
                    
        If Not IsNull(rsEj!bprcpid) Then
            Ok = ANULAR_PROCESO(rsEj!bprcpid)
        End If
        
        ' actualizo los datos del proceso
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Time, FormatoInternoHora) & "', bprcfecfinej = " & ConvFecha(Now) & ", bprcestado = 'Abortado por Usuario'" & _
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

    TiempoDeEsperaNoResponde = 3
    
    StrSql = "SELECT * FROM batch_proceso WHERE bprcterminar = 0 and bprcestado = 'No Responde'"
    OpenRecordset StrSql, rsEj
    
    Do While Not rsEj.EOF
        If DateDiff("n", rsEj!bprchorafinej, Now) > TiempoDeEsperaNoResponde Then
            Flog.writeline "Proceso " & rsEj!bpronro & " Abortado porque No Responde" & Format(CDate(Now), FormatoInternoFecha)
                        
            If Not IsNull(rsEj!bprcpid) Then
                Ok = ANULAR_PROCESO(rsEj!bprcpid)
            End If
            
            ' actualizo los datos del proceso
            StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Time, FormatoInternoHora) & "', bprcfecfinej = " & ConvFecha(Now) & ", bprcestado = 'Abortado'" & _
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

    TiempoDeEsperaSinProgreso = 5
    
    StrSql = "SELECT * FROM batch_proceso WHERE (bprcestado = 'Procesando')"
    OpenRecordset StrSql, rsEj
    Flog.writeline "Busco procesos en estado Procesando  - " & Format(CDate(Now), FormatoInternoFecha)
        
    Do While Not rsEj.EOF
        Flog.writeline "Encontró Procesando  - " & rsEj!bpronro & Format(CDate(Now), FormatoInternoFecha)
        MiIndice = BuscarIndice(rsEj!bpronro)
        
        If Ejecutando(MiIndice).Progreso = rsEj!bprcprogreso Then
           Flog.writeline "No avansó el progreso. espero"
           If DateDiff("n", Format(Ejecutando(MiIndice).HoraFinEj, FormatoInternoFecha), Format(CDate(Now), FormatoInternoFecha)) > TiempoDeEsperaSinProgreso Then
                Flog.writeline "No avansó el progreso en 5 minutos. Pone Proceso " & rsEj!bpronro & " en estado NO RESPONDE - " & Format(CDate(Now), FormatoInternoFecha)
                ' si hace mas de 5 minutos que no avanza entonces ponemos su estado en No Responde
                StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Time, FormatoInternoHora) & "', bprcfecfinej = " & ConvFecha(Now) & ", bprcestado = 'No Responde'" & _
                " WHERE bpronro = " & Ejecutando(MiIndice).Proceso
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
        Else
            Flog.writeline "Actualizo el progreso "
            ' hora y fecha del ultimo progreso detectado
            Flog.writeline "Proceso " & rsEj!bpronro & " indice : " & MiIndice & ", Progreso: " & rsEj!bprcprogreso & " - " & Format(CDate(Now), FormatoInternoFecha)
            If IsNull(rsEj!bprcprogreso) Then
                Ejecutando(MiIndice).Progreso = 0
            Else
                Ejecutando(MiIndice).Progreso = rsEj!bprcprogreso
            End If
            Flog.writeline "Proceso " & rsEj!bpronro & " indice : " & MiIndice & ", HoraFinEj: " & Format(CDate(Time), FormatoInternoFecha) & " - " & Format(CDate(Now), FormatoInternoFecha)
            Ejecutando(MiIndice).HoraFinEj = Format(CDate(Time), FormatoInternoFecha)
        End If
        
        rsEj.MoveNext
    Loop
    
    If rsEj.State = adStateOpen Then rsEj.Close
    Set rsEj = Nothing
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
i = 0
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


