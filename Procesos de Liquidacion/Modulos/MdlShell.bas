Attribute VB_Name = "MdlShell"
Option Explicit
' ---------------------------------------------------------------------------------------------
' Descripcion:Modulo Funcionalidad del Shell
' Autor      :JPB
' Fecha      :14/12/2016
' Ultima Mod.:
' Descripcion:

 
'JPB: Se definen las referencias a las funciones para el manejo de procesos
Private Declare Function OpenProcess Lib "kernel32" _
  (ByVal dwDesiredAccess As Long, _
   ByVal bInheritHandle As Long, _
   ByVal dwProcessId As Long) As Long
  
Private Declare Function GetExitCodeProcess Lib "kernel32" _
  (ByVal hProcess As Long, lpExitCode As Long) As Long
  
Private Declare Function CloseHandle Lib "kernel32" _
  (ByVal hObject As Long) As Long

Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
  
 
 
Public Function ValidarArchivo(ByVal FileName As String) As VbTriState
    Dim intFile As Integer

    On Error Resume Next
    GetAttr FileName
    If Err.Number Then
        ValidarArchivo = vbUseDefault 'Archivo no existe o el servidor no esta disponible
    Else
        Err.Clear
        intFile = FreeFile(0)
        Open FileName For Binary Lock Read Write As #intFile
        If Err.Number Then
            ValidarArchivo = vbFalse 'El archivo ya esta abierto
        Else
            Close #intFile
            ValidarArchivo = vbTrue 'El archivo esta disponible para ser usado
        End If
    End If
End Function

 

'JPB: Ejecuta un shell sincronicamente esperando a que finalice el proceso lanzado
'Utiliza las funciones de kernel32 para verificar estados de los procesos
Public Function Shell_Sync(ByVal programa As String, ByVal minutos_espera As Integer, ByRef CodigoSalida As Integer)
  
    Const PROCESS_QUERY_INFORMATION = &H400
    Const STATUS_PENDING = &H103&
    
    Dim handle_Process As Long
    Dim id_process As Long
    Dim lp_ExitCode As Long
    Dim horaInicial As Date
    Dim horaMaxima As Date
    Dim res As String
    Dim objShell
    Dim objExec
            
    horaInicial = Format$(Now, "hh:mm:ss")
    'minutos_espera = 20
    'Calcula el tiempo maximo de espera al proceso sumandole 20 minutos a la hora de ejecucion
    horaMaxima = DateAdd("n", minutos_espera, horaInicial)
    ' Abre el proceso con el shell
    Set objShell = CreateObject("WScript.Shell")
    Set objExec = objShell.Exec(programa)
    'Si objExec.ExitCode=0 ejecuto correctamente
     
    Shell_Sync = id_process
    ' Creo un handle hacia el proceso ejecutado por el shell
    handle_Process = OpenProcess(PROCESS_QUERY_INFORMATION, False, id_process)
    
    'Espero a que finalice el proceso disparado por el shell o sale si cumple un derminado tiempo maximo de espera
    Do
         horaInicial = Format$(Now, "hh:mm:ss")
         'Consulta sobre el codigo de salida del proceso
         res = GetExitCodeProcess(handle_Process, lp_ExitCode)
         DoEvents
     'Repetir mientras el proceso este pendiente y no haya sobrepasado el tiempo maximo de espera
    Loop While (lp_ExitCode = STATUS_PENDING) And (horaInicial < horaMaxima)
    
    'Espera 1 Segundo luego de la ejecucion del shell.
    Sleep 1000
    
    'Recupero el codigo de salida del resultado de ejecutar el Shells
    CodigoSalida = objExec.ExitCode
       
    'Cierro el handle
    Call CloseHandle(handle_Process)
    
    Flog.writeline "Finaliza ejecución del proceso por shell: " & programa & " con codigo de salida: " & CodigoSalida
 
End Function



'JPB: Ejecuta un shell asincronicamente sin esperar a que finalice el proceso lanzado
Public Function Shell_Async(programa As String)
     Shell_Async = Shell(programa, 1)
End Function
