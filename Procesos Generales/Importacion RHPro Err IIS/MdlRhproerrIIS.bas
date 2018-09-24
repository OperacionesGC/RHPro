Attribute VB_Name = "MdlRHProImpErrIIS"
Option Explicit

Global Const Version = "1.00" 'Importacion de log de errores del IIS
Global Const FechaModificacion = "25/04/2014"
Global Const UltimaModificacion = "" 'Miriam Ruiz - Version Inicial - CAS-24137 - H&A - Calidad - Confiabilidad - Tolerancia a fallas

Global Seed As String 'Usado como clave de encriptacion/desencriptacion
Global encryptAct As Boolean
Global Sep As String
Global CantRegErr As Long
Global ArchOpen As Boolean
Global encabezados(1 To 22, 1 To 2) As String
Global f
Global RegLeidos As Long
Global NroLinea As Long
Global RegError As Long
Global ArrErrores() As Double




Global ProcPendiente As Boolean




Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial
' Autor      : Miriam Ruiz
' Fecha      : 14/04/2014
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim objconnMain As New ADODB.Connection
Dim strCmdLine
Dim rs_batch_proceso As New ADODB.Recordset
Dim PID As String
Dim bprcparam As String
Dim bprcfecha As Date
Dim Nombre_Arch As String

Dim ArrParametros

    strCmdLine = Command()
    ArrParametros = Split(strCmdLine, " ", -1)
    If UBound(ArrParametros) > 1 Then
        If IsNumeric(ArrParametros(0)) Then
            NroProcesoBatch = ArrParametros(0)
            Etiqueta = ArrParametros(1)
            EncriptStrconexion = CBool(ArrParametros(2))
            c_seed = ArrParametros(2)
        Else
            Exit Sub
        End If
    Else
        If UBound(ArrParametros) > 0 Then
            If IsNumeric(ArrParametros(0)) Then
                NroProcesoBatch = ArrParametros(0)
                Etiqueta = ArrParametros(1)
            Else
                Exit Sub
            End If
        Else
            If IsNumeric(strCmdLine) Then
                NroProcesoBatch = strCmdLine
            Else
                Exit Sub
            End If
        End If
    End If
    
    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas

    
    Nombre_Arch = PathFLog & "Imp.Log Errores IIS" & " - " & NroProcesoBatch & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaModificacion
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-------------------------------------------------"
    Flog.writeline
    
    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    
    On Error Resume Next
    OpenConnection strconexion, objconnProgreso
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    
    On Error GoTo ME_Main
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0 ,bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Flog.writeline "Pone el estado en procesando"
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 416 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        If IsNull(rs_batch_proceso!bprcparam) Then
            bprcparam = ""
        Else
            bprcparam = rs_batch_proceso!bprcparam
        End If
        bprcfecha = rs_batch_proceso!bprcfecha
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        Call ImportErr(NroProcesoBatch, bprcparam, bprcfecha)
    Else
        Flog.writeline "no encontró el proceso"
    End If
    
    TiempoFinalProceso = GetTickCount
    Flog.writeline
    Flog.writeline "--------------------------------------------------------------"
    Flog.writeline "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    
    If HuboError Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error', bprcprogreso = 100  WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado', bprcprogreso = 100 WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
    End If
    
    
Fin:
    Flog.Close
    If objconnProgreso.State = adStateOpen Then objconnProgreso.Close
Exit Sub

ME_Main:
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
        Flog.writeline Espacios(Tabulador * 0) & " Error General " & Err.Description
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "SQL Ejecutado: " & StrSql
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
        Flog.writeline
        
    'Actualizo el progreso
  '  MyBeginTrans
 '       StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Error General', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
 '       objconnProgreso.Execute StrSql, , adExecuteNoRecords
 '   MyCommitTrans
End Sub

Public Sub InicializarEncab()

'date
encabezados(1, 2) = "#"
'time
encabezados(2, 2) = "#"
's-sitename
encabezados(3, 2) = "#"
's-computername
encabezados(4, 2) = "#"
's-ip
encabezados(5, 2) = "#"
'cs-method
encabezados(6, 2) = "#"
'cs-uri-stem
encabezados(7, 2) = "#"
'cs-uri-query
encabezados(8, 2) = "#"
's-port
encabezados(9, 2) = "#"
'cs-username
encabezados(10, 2) = "#"
'c-ip
encabezados(11, 2) = "#"
'cs-version
encabezados(12, 2) = "#"
'cs(User-Agent)
encabezados(13, 2) = "#"
'cs(Cookie)
encabezados(14, 2) = "#"
'cs(Referer)
encabezados(15, 2) = "#"
'cs-host
encabezados(16, 2) = "#"
'sc-status
encabezados(17, 2) = "#"
'sc-substatus
encabezados(18, 2) = "#"
'sc-win32-status
encabezados(19, 2) = "#"
'sc-bytes
encabezados(20, 2) = "#"
'cs-bytes
encabezados(21, 2) = "#"
'time-taken
encabezados(22, 2) = "#"


End Sub

Public Sub CrearEncab()

'date
encabezados(1, 1) = "date"
'time
encabezados(2, 1) = "time"
's-sitename
encabezados(3, 1) = "s-sitename"
's-computername
encabezados(4, 1) = "s-computername"
's-ip
encabezados(5, 1) = "s-ip"
'cs-method
encabezados(6, 1) = "cs-method"
'cs-uri-stem
encabezados(7, 1) = "cs-uri-stem"
'cs-uri-query
encabezados(8, 1) = "cs-uri-query"
's-port
encabezados(9, 1) = "s-port"
'cs-username
encabezados(10, 1) = "cs-username"
'c-ip
encabezados(11, 1) = "c-ip"
'cs-version
encabezados(12, 1) = "cs-version"
'cs(User-Agent)
encabezados(13, 1) = "cs(User-Agent)"
'cs(Cookie)
encabezados(14, 1) = "cs(Cookie)"
'cs(Referer)
encabezados(15, 1) = "cs(Referer)"
'cs-host
encabezados(16, 1) = "cs-host"
'sc-status
encabezados(17, 1) = "sc-status"
'sc-substatus
encabezados(18, 1) = "sc-substatus"
'sc-win32-status
encabezados(19, 1) = "sc-win32-status"
'sc-bytes
encabezados(20, 1) = "sc-bytes"
'cs-bytes
encabezados(21, 1) = "cs-bytes"
'time-taken
encabezados(22, 1) = "time-taken"


End Sub

Function YaFueProcesado(archivo)
Dim rs_Consult As New ADODB.Recordset
Dim ArchivoAux

ArchivoAux = Left(archivo, Len(archivo) - 4) & ".prc"
StrSql = "select arch_procesado from rhpro_err where arch_procesado = '" & Trim(ArchivoAux) & "'"
OpenRecordset StrSql, rs_Consult
 If Not rs_Consult.EOF Then
     YaFueProcesado = True
 Else
    YaFueProcesado = False
 End If
End Function


Function Buscarencab(ByVal nombre As String)

Dim I As Integer
Dim flag As Boolean

flag = False
I = 1
Do While (I < 23) And (Not flag)
    If encabezados(I, 1) = nombre Then
           flag = True
    Else
        I = I + 1
    End If
Loop
Buscarencab = I
    
End Function

Private Sub CrearArrErrores()
Dim rs_Consult As New ADODB.Recordset
Dim I As Integer
Dim MaxArr As Integer

I = 0
StrSql = "select count(sc_status_nro) cant from rhpro_err_status where err_activo = 0"
OpenRecordset StrSql, rs_Consult
 If Not rs_Consult.EOF Then
     MaxArr = rs_Consult!Cant
 End If
ReDim ArrErrores(MaxArr) As Double

StrSql = "select sc_status_nro from rhpro_err_status where err_activo = 0"
OpenRecordset StrSql, rs_Consult
Do While Not rs_Consult.EOF
       
    ArrErrores(I) = rs_Consult!sc_status_nro
    I = I + 1
    rs_Consult.MoveNext
Loop

End Sub

Function EstaEnArreglo(error As String) As Boolean
Dim I As Integer
Dim flag As Boolean

flag = False
I = 0
Do While I <= UBound(ArrErrores) And Not flag
    If ArrErrores(I) = CInt(error) Then
        flag = True
    Else
        I = I + 1
    End If
Loop
EstaEnArreglo = flag

End Function


Private Sub GuardarDatos(nombreArch)
Dim Campo
Dim ruta As String
Dim archivo As String
Dim parametros As String
Dim nrolineas As Integer
Dim CodError As String
Dim DescError As String
Dim I As Integer

CodError = encabezados(17, 2)

 If EstaEnArreglo(CodError) Then

    Campo = Split(encabezados(7, 2), "/")
    ruta = ""
    If UBound(Campo) > 0 Then
        For I = 0 To UBound(Campo) - 1
             ruta = ruta & Campo(I) & "/"
        Next I
        archivo = Campo(UBound(Campo))
        Campo = Split(encabezados(8, 2), "|")
        parametros = "#"
        nrolineas = 0
        CodError = "#"
        DescError = "#"
        If UBound(Campo) > 0 Then
            If Campo(1) = "-" Then
                Campo(1) = "0"
            End If
        End If
        Select Case UBound(Campo)
        Case 0: parametros = Campo(0)
        Case 1: parametros = Campo(0)
                nrolineas = Campo(1)
        Case 2: parametros = Campo(0)
                nrolineas = Campo(1)
                CodError = Campo(2)
        Case Else
                parametros = Campo(0)
                nrolineas = CInt(Campo(1))
                CodError = Campo(2)
                DescError = Campo(3)
       End Select
  
        DescError = Replace(DescError, "'", "")
        StrSql = "INSERT INTO rhpro_err (s_date , s_time, s_sitename, s_computername, s_ip" & _
                 ",cs_method, ruta_uri_stem, archivo_uri_stem, parametros_uri_query,nrolines_uri_query" & _
                 ",cod_error_uri_query,desc_error_uri_query,s_port,cs_username,c_ip,cs_version,cs_User_Agent" & _
                 ",cs_Cookie,cs_Referer,cs_host,sc_status_nro,sc_substatus_nro,sc_win32_status_nro,sc_bytes,cs_bytes" & _
                 ",time_taken, tecnro,arch_procesado) VALUES ('" & _
                 encabezados(1, 2) & "','" & encabezados(2, 2) & "','" & encabezados(3, 2) & "','" & encabezados(4, 2) & "','" & _
                 encabezados(5, 2) & "','" & encabezados(6, 2) & "','" & ruta & "','" & archivo & "','" & _
                 parametros & "'," & nrolineas & ",'" & CodError & "','" & DescError & "'," & _
                 encabezados(9, 2) & ",'" & encabezados(10, 2) & "','" & encabezados(11, 2) & "','" & encabezados(12, 2) & "','" & _
                 encabezados(13, 2) & "','" & encabezados(14, 2) & "','" & encabezados(15, 2) & "','" & encabezados(16, 2) & "'," & _
                 encabezados(17, 2) & "," & encabezados(18, 2) & "," & encabezados(19, 2) & "," & encabezados(20, 2) & "," & _
                 encabezados(21, 2) & "," & encabezados(22, 2) & "," & _
                 "4,'" & nombreArch & "' )"
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
End If

End Sub

Public Sub ProcesaArchivo(archivo, directorio)
Dim fImport
Dim carpeta
Dim folder
Dim file
Dim fileItem
Dim strLineaArch
Dim CErroresAProc


Dim CantReg As Long


Dim Destino As String

    
    Dim datos() As String
    Dim I As Long
    Dim numeral As String
    Dim Campo As String
    Dim strLinea As String
    Dim Separador As String
    Dim encabpos() As Integer
    
'-------------------------------------------------------------------------------------------------
'Renombro el archivo para que no lo tome otro proceso
'-------------------------------------------------------------------------------------------------
fs.MoveFile directorio & "\" & archivo, directorio & "\" & Left(archivo, Len(archivo) - 4) & ".prc"
archivo = Left(archivo, Len(archivo) - 4) & ".prc"


'-------------------------------------------------------------------------------------------------
'Apertura del archivo
'-------------------------------------------------------------------------------------------------
On Error Resume Next
Flog.writeline Espacios(Tabulador * 0) & "Buscando el archivo " & archivo
'Archivo = Directorio & "\" & Archivo
Set fImport = fs.OpenTextFile(directorio & "\" & archivo, 1, 0)
If Err.Number <> 0 Then
    Flog.writeline Espacios(Tabulador * 1) & "Error. No se encontro el archivo " & archivo
    HuboError = True
    Exit Sub
End If
'On Error GoTo E_ImportLog
'Flog.writeline

'Marco que abri el archivo
ArchOpen = True

'-------------------------------------------------------------------------------------------------
'Calculo la cantidad de lineas
'-------------------------------------------------------------------------------------------------
CantReg = 0
Do While Not fImport.AtEndOfStream
    strLineaArch = fImport.ReadLine
    CantReg = CantReg + 1
Loop
fImport.Close


'-------------------------------------------------------------------------------------------------
'Seteo de las variables de progreso
'-------------------------------------------------------------------------------------------------
Progreso = 0
CErroresAProc = CantReg
Flog.writeline
If CErroresAProc = 0 Then
    Flog.writeline Espacios(Tabulador * 0) & "No hay Datos a Importar"
    CErroresAProc = 1
Else
    Flog.writeline Espacios(Tabulador * 0) & "Cantidad de Registros a Importar: " & CErroresAProc
End If
IncPorc = (100 / CErroresAProc)
Flog.writeline

        
Set fImport = fs.OpenTextFile(directorio & "\" & archivo, 1, 0)

Flog.writeline Espacios(Tabulador * 0) & "------------------------------------------------------------------------------------"
Flog.writeline Espacios(Tabulador * 0) & "INICIO DE LECTURA DE LINEAS DEL ARCHIVO"
Flog.writeline Espacios(Tabulador * 0) & "------------------------------------------------------------------------------------"
  Separador = Chr(32) ' caracter vacío
Do While Not fImport.AtEndOfStream
    
    'Leo la linea del archivo
    strLineaArch = fImport.ReadLine
    
    'Aplico desencriptacion
    ' strLinea = DesEncriptar(strLineaArch)
   strLinea = strLineaArch
    
       datos = Split(strLinea, Separador)
            numeral = Mid(datos(0), 1, 1)
            If numeral = "#" Then
                Campo = Mid(datos(0), 2, 6)
                If Campo = "Fields" Then
                  Call InicializarEncab
                  ReDim encabpos(0 To UBound(datos))
                    For I = 1 To UBound(datos)
                      
                        encabpos(I) = Buscarencab(datos(I))
                    Next I
                End If
            Else
                For I = 0 To UBound(datos)
                    encabezados(encabpos(I + 1), 2) = Trim(datos(I))
                Next I
                
                Call GuardarDatos(archivo)
            End If
    
    CantReg = CantReg + 1
    
    'comentar desde aca
    'Actualizo el Progreso
    Progreso = Progreso + IncPorc
  '  Progreso = 0
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
    StrSql = StrSql & " , bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso))
    StrSql = StrSql & "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
' hasta aca
Loop

'Para el ultimo registro
If CantReg <> 0 Then
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "Cant. de Reg. Procesados = " & CantReg
    Flog.writeline Espacios(Tabulador * 0) & "Cant. de Reg. con Error  = " & CantRegErr
    Flog.writeline
End If


fImport.Close
ArchOpen = False

'Muevo el archivo a la carpeta de backup o error dependiendo si lo proceso exitosamente o no
Set fImport = fs.GetFile(directorio & "\" & archivo)

'Seteo el destino despendiendo si hubo algun error o no
If HuboError Then
    Destino = directorio & "\Err\"
Else
    Destino = directorio & "\bk\"
End If
Flog.writeline
Flog.writeline Espacios(Tabulador * 0) & "Moviendo el archivo a la carpeta " & Destino
'Muevo el archivo creando la carpeta respectiva si no existe
On Error Resume Next
    fImport.Move Destino & archivo
    If Err.Number <> 0 Then
        Flog.writeline Espacios(Tabulador * 0) & "La carpeta Destino no existe. Se creará."
        Set carpeta = fs.CreateFolder(Destino)
        fImport.Move Destino & archivo
    End If
'On Error GoTo E_ImportLog:
'Flog.writeline

End Sub


Public Sub ImportErr(ByVal bpronro As Long, ByVal parametros As String, ByVal bprcfecha As Date)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de Importacion del log de errores del ISS
' Autor      : Miriam Ruiz
' Fecha      : 14/04/2014
' --------------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------------------
'Parametros
'-------------------------------------------------------------------------------------------------
Dim Auto As Boolean
Dim Periodo As Long
Dim Fecha As Date

'-------------------------------------------------------------------------------------------------
'Variables
'-------------------------------------------------------------------------------------------------

Dim directorio As String


Dim archivo As String
Dim fImport
Dim carpeta
Dim folder
Dim file
Dim fileItem
Dim strLineaArch
Dim CErroresAProc

Dim CantReg As Long


Dim Destino As String

    
    Dim datos() As String
    Dim I As Long
    Dim numeral As String
    Dim Campo As String
    Dim strLinea As String
    Dim Separador As String
    Dim encabpos() As Integer
    Dim subnombre As String
    Dim FechaActual
    Dim procesado
            

'-------------------------------------------------------------------------------------------------
'RecordSets
'-------------------------------------------------------------------------------------------------
Dim rs_Consult As New ADODB.Recordset

'Inicio codigo ejecutable
On Error GoTo E_ImportLog

'Valores default de encriptacion: Activa y semilla = 56238
Seed = "56238"
encryptAct = True
HuboError = False
ArchOpen = False
archivo = parametros

If parametros = "" Then
    Auto = True
Else
    Auto = False
End If
Call CrearEncab
Call CrearArrErrores
'-------------------------------------------------------------------------------------------------
'Configuracion del Directorio de entrada
'-------------------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 0) & "Buscando directorio de entrada."
StrSql = "SELECT sis_direrr FROM sistema WHERE sisnro = 1 "
OpenRecordset StrSql, rs_Consult
If Not rs_Consult.EOF Then
    directorio = Trim(rs_Consult!sis_direrr)
Else
    Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontró el registro de la tabla sistema "
    HuboError = True
    Exit Sub
End If
Flog.writeline
    
    

'-------------------------------------------------------------------------------------------------
'Busqueda del archivo en caso de disparo automatico
'-------------------------------------------------------------------------------------------------
Set fs = CreateObject("Scripting.FileSystemObject")

If Auto Then
    Flog.writeline Espacios(Tabulador * 0) & "Buscando archivo a procesar."

    'Seteo el nombre del archivo generado
    On Error Resume Next
    
    'Busco Directorio
    Set folder = fs.GetFolder(directorio)
    If Err.Number <> 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontró el directorio " & directorio
        HuboError = True
        Exit Sub
    End If
    
    'Busco el primer archivo con extension log
    Set file = folder.Files
    archivo = ""
    For Each fileItem In file
        If fs.GetExtensionName(directorio & "\" & fileItem.Name) = "log" Then
            archivo = fileItem.Name
            Flog.writeline Espacios(Tabulador * 1) & "Archivo pendiente de procesamiento encontrado " & archivo
            subnombre = Mid(archivo, 5, 6)
            FechaActual = CDate(Date)
            FechaActual = Right(FechaActual, 2) & Mid(FechaActual, 4, 2) & Left(FechaActual, 2)
            procesado = YaFueProcesado(archivo)
            If subnombre <> FechaActual And Not (procesado) Then
                Call ProcesaArchivo(archivo, directorio)
            End If
            If procesado Then
               Flog.writeline Espacios(Tabulador * 1) & "El archivo " & archivo & " ya fue procesado anteriormente"
            
            End If
            
        End If
    Next
    
    If Len(archivo) = 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "No se encontró ningun archivo a procesar (.log)"
        HuboError = True
        Exit Sub
    End If
    
    Set file = Nothing
    On Error GoTo E_ImportLog
Else
    procesado = YaFueProcesado(archivo)
     If Not (procesado) Then
        Call ProcesaArchivo(archivo, directorio)
     End If
     If procesado Then
           Flog.writeline Espacios(Tabulador * 1) & "El archivo " & archivo & " ya fue procesado anteriormente"
     End If
End If
Flog.writeline


If rs_Consult.State = adStateOpen Then rs_Consult.Close

Set rs_Consult = Nothing
Set fImport = Nothing
Set fs = Nothing

Exit Sub

E_ImportLog:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: ImportErrIIS"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="
    MyBeginTrans
    Progreso = Round(Progreso + IncPorc, 4) 'comentar esta
    ' Progreso = 0
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
             "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    MyCommitTrans
    
'------------------------------------------------------------------
'Movimiento del archivo
'------------------------------------------------------------------
    If ArchOpen Then
        'Cierro el archivo
        fImport.Close
        
        'Muevo el archivo a la carpeta de backup o error dependiendo si lo proceso exitosamente o no
        Set fImport = fs.GetFile(directorio & "\" & archivo)
        
        'Seteo el destino despendiendo si hubo algun error o no
        Destino = directorio & "\Err\"
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "Moviendo el archivo a la carpeta " & Destino
        'Muevo el archivo creando la carpeta respectiva si no existe
        On Error Resume Next
            fImport.Move Destino & archivo
            If Err.Number <> 0 Then
                Flog.writeline Espacios(Tabulador * 0) & "La carpeta Destino no existe. Se creará."
                Set carpeta = fs.CreateFolder(Destino)
                fImport.Move Destino & archivo
            End If
        On Error GoTo E_ImportLog:
        Flog.writeline
    
    End If
    
    HuboError = True
    Flog.writeline " Error: " & Err.Description

End Sub


Public Function CtrlNuloTXT(ByVal Valor) As String

    If IsNull(Valor) Then
        CtrlNuloTXT = "NULL"
    Else
        If UCase(Valor) = "NULL" Then
            CtrlNuloTXT = "NULL"
        Else
            CtrlNuloTXT = "'" & Valor & "'"
        End If
    End If
    
End Function




Public Function cambiaFecha(ByVal Fecha As String) As String

    If EsNulo(Fecha) Then
        cambiaFecha = "NULL"
    Else
        cambiaFecha = ConvFecha(Fecha)
    End If

End Function









