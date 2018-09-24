Attribute VB_Name = "BajarRegistraciones"
Option Explicit

Const INFINITE = &HFFFF
Const STARTF_USESHOWWINDOW = &H1
Private Enum enSW
    SW_HIDE = 0
    SW_NORMAL = 1
    SW_MAXIMIZE = 3
    SW_MINIMIZE = 6
End Enum
Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type
Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Byte
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
Private Enum enPriority_Class
    NORMAL_PRIORITY_CLASS = &H20
    IDLE_PRIORITY_CLASS = &H40
    HIGH_PRIORITY_CLASS = &H80
End Enum

Global fs, f, f2, flog

Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Function SuperShell(ByVal App As String, ByVal WorkDir As String, dwMilliseconds As Long, ByVal start_size As enSW, ByVal Priority_Class As enPriority_Class) As Boolean
    Dim pclass As Long
    Dim sinfo As STARTUPINFO
    Dim pinfo As PROCESS_INFORMATION
    Dim Rta
    'Not used, but needed
    Dim sec1 As SECURITY_ATTRIBUTES
    Dim sec2 As SECURITY_ATTRIBUTES
    'Set the structure size
    sec1.nLength = Len(sec1)
    sec2.nLength = Len(sec2)
    sinfo.cb = Len(sinfo)
    'Set the flags
    sinfo.dwFlags = STARTF_USESHOWWINDOW
    'Set the window's startup position
    sinfo.wShowWindow = start_size
    'Set the priority class
    pclass = Priority_Class
    
    'Start the program
    If CreateProcess(vbNullString, App, sec1, sec2, False, pclass, 0&, WorkDir, sinfo, pinfo) Then
      'Wait
      WaitForSingleObject pinfo.hProcess, INFINITE
      SuperShell = True
    Else
      Rta = MsgBox("No se pudo ejecutar " & App & " en " & WorkDir, vbOKOnly, "Ejecución de SuperShell.")
      SuperShell = False
    End If
End Function

Public Sub TransferirViejo(path, pathlocal, NombreArchivo)
Dim archivo As String

Dim strline As String
Dim pos1 As Integer
Dim pos2 As Integer

Dim Archivo_Reg As String
Dim Archivo_Aux As String

Dim cantidad As Integer
  
    Archivo_Reg = pathlocal & NombreArchivo
    
    On Error Resume Next
    Set f = fs.getfile(Archivo_Reg)
    Archivo_Aux = Archivo_Reg
    If Err.Number = 0 Then
        
    ' El archivo existe. Debo transferirlo
        cantidad = 1
        flog.writeline Format(Now, "yyyy-mm-dd hh:mm:ss") & " Quedo sin trasnferir " & Archivo_Aux
                
        Do
            f.Copy path & NombreArchivo
            cantidad = cantidad + 1
        Loop While (Err.Number <> 0) And (cantidad < 50)
        
        If cantidad = 50 Then
            'MsgBox "Problemas con la conexión. Imposible copiar registraciones.", vbCritical, "Bajar Registraciones"
            flog.writeline Format(Now, "yyyy-mm-dd hh:mm:ss") & " Fallo transferencia " & Archivo_Aux

            Exit Sub
        Else
        ' Transmiti exitosamente.
        ' Renombro en el cliente
            f.Move Mid(Archivo_Aux, 1, Len(Archivo_Aux) - 3) & "tx"
        End If
        
    End If
    On Error GoTo 0
    
End Sub

Public Sub Transferir(path, pathlocal, NombreArchivo)
Dim archivo As String

Dim strline As String
Dim pos1 As Integer
Dim pos2 As Integer
Dim Archivo_Reg As String
Dim Archivo_Aux As String
Dim cantidad As Integer
Dim Rta
Dim TamOrigen
Dim TamDestino

    Rta = 0
    Archivo_Reg = pathlocal & NombreArchivo
    
    On Error Resume Next
    Set f = fs.getfile(Archivo_Reg)
    TamOrigen = f.Size
    
    Archivo_Aux = Archivo_Reg
    If Err.Number = 0 Then
        
        Do
            Rta = 0
            cantidad = 1
            ' El archivo existe. Debo transferirlo
            ' intento copiarlo x cantidad de veces, si no tengo exito pregunto
            Do
                Err.Number = 0
                f.Copy path & NombreArchivo
                cantidad = cantidad + 1
            Loop While (Err.Number <> 0) And (cantidad < 10)
            
            If cantidad = 10 Then
                
                flog.writeline Format(Now, "yyyy-mm-dd hh:mm:ss") & " " & NombreArchivo & " " & "0" & " Fallo transferencia. "
                ' FGZ 10/09/2003
                flog.writeline "Error: " & Err.Number
                flog.writeline "Descripción: " & Err.Description
                
                Rta = MsgBox("Problemas con la conexión. Imposible transferir registraciones.", vbRetryCancel, "Transferencia de Registraciones")
                
            Else
                Set f2 = fs.getfile(path & NombreArchivo)
                TamDestino = f.Size
                
                If TamOrigen <> TamDestino Then
                    flog.writeline Format(Now, "yyyy-mm-dd hh:mm:ss") & " " & NombreArchivo & " " & TamDestino & " Fallo transferencia. "
                    ' FGZ 10/09/2003
                    flog.writeline "Error: " & Err.Number
                    flog.writeline "Descripción: " & Err.Description
                            
                    Rta = MsgBox("Problemas de Transferencia. Archivo transferido Dañado.", vbRetryCancel, "Transferencia de Registraciones")
                
                Else
                    ' Transmiti exitosamente.
                    flog.writeline Format(Now, "yyyy-mm-dd hh:mm:ss") & " " & NombreArchivo & " " & TamDestino & " Transferencia OK."
                    ' Renombro en el cliente
                    
                    Err.Number = 0
                    f.Move Mid(Archivo_Aux, 1, Len(Archivo_Aux) - 3) & "tx"
                    If (Err.Number <> 0) Then
                      flog.writeline Format(Now, "yyyy-mm-dd hh:mm:ss") & " " & NombreArchivo & " No se puede renombrar a " & Mid(Archivo_Aux, 1, Len(Archivo_Aux) - 3) & "tx"
                      flog.writeline "Error: " & Err.Number
                      flog.writeline "Descripción: " & Err.Description
                    End If
' Si no se puede renombrar, se registra cuál es el error. O.D.A. 17/10/2003

                End If
            End If
        Loop While Rta = vbRetry
    
    End If
    
    On Error GoTo 0
    
End Sub


Public Sub Main()

Const ForReading = 1
Const ForAppending = 8

Dim strline As String
Dim pos1 As Integer
Dim pos2 As Integer
Dim strNombreArchivo As String
Dim path As String
Dim pathlocal As String
Dim nombEjec As String
Dim dirEjec As String

Dim Archivo_Reg As String
Dim Archivo_Aux As String

Dim cantidad As Integer
Dim CArchivos
Dim MiArchivo
Dim Folder
Dim HabiaPendientes As Boolean
Dim Rta
    
'----FGZ  02/04/2003----------------------------------------
    strNombreArchivo = ""

    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.OpenTextFile(App.path & "\BajarReg.INI", ForReading, 0)
    
    ' seteo del archivo de log
    If fs.fileexists("c:\Registraciones.log") Then
        ' lo abro
        Set flog = fs.OpenTextFile("c:\Registraciones.log", ForAppending, 0)
    Else
        ' no existe, entonces lo creo
        Set flog = fs.CreateTextFile("c:\Registraciones.log", True)
    End If
    
    If Not f.AtEndOfStream Then
        'Leo el nombre del archivo de registraciones
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        strNombreArchivo = Mid(strline, pos1, pos2 - pos1)
    End If
    If Not f.AtEndOfStream Then
        'Leo el path del server en donde se graban las reg.
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        path = Mid(strline, pos1, pos2 - pos1)
        If Right(path, 1) <> "\" Then path = path & "\"
    End If
    If Not f.AtEndOfStream Then
        'Leo el path local del archivo de registraciones
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        pathlocal = Mid(strline, pos1, pos2 - pos1)
    End If
    If Not f.AtEndOfStream Then
        'Leo el nombre del ejecutable
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        nombEjec = Mid(strline, pos1, pos2 - pos1)
    End If
    If Not f.AtEndOfStream Then
        'Leo el directorio del ejecutable
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        dirEjec = Mid(strline, pos1, pos2 - pos1)
    End If
' Se recuperan del archivo de parametros el nombre y argumentos del ejecutable
' asi como el directorio sobre el que se tiene que ejecutar. Se recomienda
' utilizar especificaciones de archivo completas.
' Por ejemplo: nombEjec = c:\rhpro\mnetwin\mnetwin.exe /com /mini
'              dirEjec  = c:\rhpro\mnetwin
' O.D.A. 17/10/2003

    f.Close

    HabiaPendientes = False
    ' Chequeo si hay algo sin transferir en la maquina local
    Set Folder = fs.GetFolder(pathlocal)
    Set CArchivos = Folder.Files
    'If CArchivos.Count <> 0 Then
    '    Rta = MsgBox("Se encontraron registraciones sin transmitir. Se intentará transmitir ahora.", vbOKOnly, "Transferencia de Registraciones")
    'End If
    For Each MiArchivo In CArchivos
        If (UCase(Right(MiArchivo.Name, 4)) = ".REG") Then
            Rta = MsgBox("Se encontó una registración sin transmitir. Se intentará transmitir ahora.", vbOKOnly, "Transferencia de Registraciones")
            Call Transferir(path, pathlocal, MiArchivo.Name)
            HabiaPendientes = True
        End If
    Next


    If HabiaPendientes Then Exit Sub
    
    ' Chequea que la no existencia de una sesion
    ' abierta para el proceso, si existe alguna, no se ejecuta
    If App.PrevInstance Then Exit Sub

    SuperShell nombEjec, dirEjec, 0, SW_NORMAL, NORMAL_PRIORITY_CLASS

'-- Renombro el arhivo de registraciones -----------------------------
    
    On Error Resume Next
    Archivo_Reg = pathlocal & strNombreArchivo
    
    ' Renombro el archivo .reg
    On Error Resume Next
    Set f = fs.getfile(Archivo_Reg)
    If Err.Number = 0 Then
        'Existe regis.reg. Por ende lo renombro
        Archivo_Aux = pathlocal & Replace(Format(Now, "yyyy-mm-dd hh:mm:ss"), ":", "-") & " " & strNombreArchivo
        f.Move Archivo_Aux
    End If
    On Error GoTo 0

        
    ' Chequeo si hay algo sin transferir en la maquina local
    Set Folder = fs.GetFolder(pathlocal)
    Set CArchivos = Folder.Files
    For Each MiArchivo In CArchivos
        If (UCase(Right(MiArchivo.Name, 4)) = ".REG") Then
            Call Transferir(path, pathlocal, MiArchivo.Name)
        End If
    Next
    
'---FGZ  02/04/2003-------------------------------

End Sub
