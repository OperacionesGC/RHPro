Attribute VB_Name = "MdlinterfazBDO"
Option Explicit

Const Version = "1.00"
Const FechaVersion = "21/11/2013" ' cas XXXX - Deluchi Ezequiel - Version Inicial

'---------------------------------------------------------------------------------------------------------------------------------------------
Dim dirsalidas As String
Dim usuario As String
Dim Incompleto As Boolean
'-------------------------------------------------------------------------------------------------
'Conexion Externa
'-------------------------------------------------------------------------------------------------
Global ExtConn As New ADODB.Connection
Global ExtConnOra As New ADODB.Connection
Global ExtConnAccess As New ADODB.Connection
Global ExtConnAccess2 As New ADODB.Connection
Global ConnLE As New ADODB.Connection
Global Usa_LE As Boolean
Global Misma_BD As Boolean
Private Type ConexionEmpresa
    ConexionOracle As String    'Guarda la conexion de oracle para la empresa
    ConexionAcces As String     'Guarda la conexion de Acces para la empresa
    estrnroEmpresa As String    'Codigo de la estrucutra empresa configurada
End Type




Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial de la interface Waldbott.
' Autor      : Deluchi Ezequiel
' Fecha      : 04/09/2012
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim objconnMain As New ADODB.Connection
Dim strCmdLine
Dim Nombre_Arch As String
Dim rs_batch_proceso As New ADODB.Recordset
Dim PID As String
Dim bprcparam As String
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

    Nombre_Arch = PathFLog & "InterfaceBDO_" & NroProcesoBatch & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaVersion
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-------------------------------------------------"
    Flog.writeline
    
    
    On Error Resume Next
    'Abro la conexion
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If

    
    OpenConnection strconexion, objconnProgreso
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If
    On Error GoTo 0
    
    On Error GoTo ME_Main
    
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0 ,bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Flog.writeline "Acutaliza el estado en procesando"
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 407 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        bprcparam = IIf(EsNulo(rs_batch_proceso!bprcparam), "", rs_batch_proceso!bprcparam)
        usuario = rs_batch_proceso!iduser
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        Call interfazBDO(NroProcesoBatch, bprcparam)
    Else
        Flog.writeline "no se encontró el proceso"
    End If
    
    TiempoFinalProceso = GetTickCount
    Flog.writeline
    Flog.writeline "--------------------------------------------------------------"
    Flog.writeline "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    
    If HuboError Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Else
        If Incompleto Then
            StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProcesoBatch
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
        Else
            StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
        End If
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
    MyBeginTrans
        StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Error General', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
    MyCommitTrans
End Sub

Public Sub interfazBDO(ByVal bpronro As Long, ByVal Parametros As String)
' Parametros =  1º Si es Exportacion (0) o Importacion (1)
'
Dim arrayParametros
Dim empresa
Dim Origen
Dim destino
Dim modelos
Dim ProcManual As Integer
Dim legdesde As Long
Dim leghasta As Long

    
    arrayParametros = Split(Parametros, "@")
        
    'Exportacion
    If CInt(arrayParametros(2)) = 0 Then
        Call exportacion(bpronro)
    End If
    
    'Importacion
    If CInt(arrayParametros(2)) = 1 Then
        Call importacion(bpronro)
    End If
    
End Sub

'Public Sub exportacion(ByVal ProcManual As Long, ByVal bpronro As Long, ByVal empresa As Long, ByVal Origen As Long, ByVal destino As Long, ByVal modelos As String, ByVal legDesde As Long, ByVal legHasta As Long)
Public Sub exportacion(ByVal bpronro As Long)
 
 Dim directorio As String
 Dim Nombre_Arch As String
 Dim rsEmpleados As New ADODB.Recordset
 Dim separador As String
 Dim strLineaExp As String
 Dim archSalida
 Dim porc As Double
 Dim cantEmpleados As Integer
 'Dim Progreso As Double
 Dim usaencabezado As String
 'Dim lineaUTF8 As String

'--------------
Dim oStream As Object
 
  Set oStream = CreateObject("ADODB.Stream") 'Create the stream
  oStream.Type = adTypeText
  oStream.Charset = "UTF-8" 'Indicate the charactor encoding
  oStream.Open 'Initialize the stream
  oStream.Position = 0 'Reset the position
  


'--------------

    'hay q levantar los empleados de batch_empleado ya los filtro el asp los desincronizados
    'preguntar si es todas o no y hacer la cosnutla
    On Error GoTo CE
    
    Flog.writeline Espacios(Tabulador * 0) & "---------------------------------------------------------"
    Flog.writeline Espacios(Tabulador * 0) & "Comienza la exportacion "
    Flog.writeline Espacios(Tabulador * 0) & "---------------------------------------------------------"
    'Progreso = 0
    
    StrSql = " SELECT * FROM empleado "
    'StrSql = " SELECT * FROM empleado WHERE empleg in (212357260,212357261,212357266) "
    OpenRecordset StrSql, rsEmpleados
            
    cantEmpleados = rsEmpleados.RecordCount
    
    'porc = CLng(50 / cantEmpleados)
    If cantEmpleados = 0 Then
        Progreso = 100
        cantEmpleados = 1
        Flog.writeline Espacios(Tabulador * 0) & "No se encontraron empleados."
    End If
    StrSql = "SELECT sis_direntradas FROM sistema WHERE sisnro = 1 "
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        directorio = Trim(objRs!sis_direntradas)
    Else
        Flog.writeline "No esta configurado el directorio de sistema."
        Exit Sub
    End If
    
    If objRs.State = adStateOpen Then objRs.Close
    
    StrSql = "SELECT * FROM modelo WHERE modnro = 389 " '& NroModelo
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        directorio = directorio & Trim(objRs!modarchdefault)
        
        If Right(directorio, 1) <> "\" Then
            directorio = directorio & "\export"
        Else
            directorio = directorio & "export"
        End If
        
       separador = IIf(Not IsNull(objRs!modseparador), objRs!modseparador, ",")
       usaencabezado = IIf(Not IsNull(objRs!modencab), CBool(objRs!modencab), False)
       Flog.writeline "Directorio de exportacion: " & directorio
    Else
        Flog.writeline "No se encontró el modelo para obtener el directorio de lectura."
        Exit Sub
    End If
            
    'EAM- Crea el archivo de exportación de la empresa
    Nombre_Arch = directorio & "\MAG01_Synch.txt"
    
    'Set fs = CreateObject("Scripting.FileSystemObject")
    'Set archSalida = fs.CreateTextFile(Nombre_Arch, False, True)

    porc = CLng(Progreso) / CLng(cantEmpleados)
    Dim pp As String
    Dim bRet() As Byte
    Do While Not rsEmpleados.EOF
        Progreso = Progreso + porc
        Flog.writeline Espacios(Tabulador * 1) & " Ternro: " & rsEmpleados!ternro
        strLineaExp = expEmpleado(rsEmpleados!ternro, separador)
                        
        If Not EsNulo(strLineaExp) Then
            bRet = AsciiToUTF8(strLineaExp)
            
            'archSalida.writeline strLineaExp
            'archSalida.writeline bRet
            'archSalida.writeline UTF8ToAscii(bRet)
            oStream.WriteText strLineaExp & Chr(13) & Chr(10) 'Write to the steam
             
        End If
        
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & " WHERE bpronro = " & bpronro
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
        
        rsEmpleados.MoveNext
    Loop
    oStream.Position = 0
    If fs.FileExists(Nombre_Arch) Then
        fs.deletefile Nombre_Arch, True
    End If
    oStream.SaveToFile Nombre_Arch 'Save the stream to a file
    oStream.Close
    Set oStream = Nothing
    
    Progreso = 100
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & " WHERE bpronro = " & bpronro
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
  
  

GoTo Procesado
CE:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 0) & "Error: " & Err.Description
    Flog.writeline Espacios(Tabulador * 0) & "SQL: " & StrSql
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline
Procesado:
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 0) & "Los datos fueron Exportados Exitosamente."
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline
End Sub

'Public Sub importacion(ByVal bpronro As Long, ByVal empresa As Long, ByVal Origen As Long, ByVal destino As Long)
Public Sub importacion(ByVal bpronro As Long)
Dim directorio As String
Dim directorioBackup As String
Dim CArchivos
Dim archivo
Dim Folder
Dim separador
Dim EncontroAlguno
Dim Path
Dim cantArchivos As Long
Dim porc As Long
Dim usaencabezado As Boolean

    StrSql = "SELECT sis_direntradas FROM sistema WHERE sisnro = 1 "
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        directorio = Trim(objRs!sis_direntradas)
    Else
        Flog.writeline "No esta configurado el directorio de sistema."
        Exit Sub
    End If

    If objRs.State = adStateOpen Then objRs.Close
    
    StrSql = "SELECT * FROM modelo WHERE modnro = 389 " '& NroModelo
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        directorio = directorio & Trim(objRs!modarchdefault)
        
        If Right(directorio, 1) <> "\" Then
            directorio = directorio & "\import"
        Else
            directorio = directorio & "import"
        End If

        If Right(directorio, 1) <> "\" Then
            directorioBackup = directorio & "\backup\"
        Else
            directorioBackup = directorio & "backup\"
        End If
        separador = IIf(Not IsNull(objRs!modseparador), objRs!modseparador, ",")
        usaencabezado = IIf(Not IsNull(objRs!modencab), CBool(objRs!modencab), False)
        Flog.writeline "Directorio de importación: " & directorio
     Else
        Flog.writeline "No se encontró el modelo para obtener el directorio de lectura."
        Exit Sub
    End If
    If objRs.State = adStateOpen Then objRs.Close
    
    Set fs = CreateObject("Scripting.FileSystemObject")
        
        Path = directorio
        
        Dim fc, F1, s2
        Set Folder = fs.GetFolder(directorio)
        Set CArchivos = Folder.Files
        
        EncontroAlguno = False
        
        cantArchivos = CArchivos.Count
        If cantArchivos = 0 Then
            cantArchivos = 1
            Progreso = 50
        Else
            Progreso = 0
        End If
        porc = 50 / cantArchivos
        
        For Each archivo In CArchivos
            Progreso = Progreso + porc
            EncontroAlguno = True
            Flog.writeline "Procesando archivo " & archivo.Name
            Call LeeArchivo(directorio & "\" & archivo.Name, usaencabezado, separador, directorioBackup & archivo.Name)
                
                
            StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & " WHERE bpronro = " & bpronro
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
        Next
        
        Call exportacion(bpronro)
        
        If Not EncontroAlguno Then
            Flog.writeline "No se encontró ningun archivo."
            Progreso = 100
            StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & " WHERE bpronro = " & bpronro
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
            
        End If
End Sub

Private Sub LeeArchivo(ByVal nombreArchivo As String, ByVal usaencabezado As Boolean, ByVal separador As String, ByVal directorioBackup As String)
' Descripcion: Lee todos los archivos del directorio y linea por linea
' Autor      : Deluchi Ezequiel
' Fecha      :
' Modificado :

Const ForReading = 1
Const TristateFalse = 0
Dim strLinea As String
Dim Archivo_Aux As String
Dim rs_Lineas As New ADODB.Recordset
Dim rs_Modelo As New ADODB.Recordset
Dim Ciclos As Long
Dim str_error As String
Dim nroLinea As Long

    If App.PrevInstance Then
        Flog.writeline Espacios(Tabulador * 0) & "Hay una instancia previa del proceso corriendo - El proceso actual queda en estado Pendiente."
        ProcPendiente = True ' para dejar el proceso pendiente
        Exit Sub
    End If
    
    'Espero hasta que se crea el archivo
    
    On Error Resume Next
    Err.Number = 1
    Ciclos = 0
    Do Until Err.Number = 0
        Err.Number = 0
        Set f = fs.GetFile(nombreArchivo)
        If f.Size = 0 Then
            If Ciclos > 100 Then
                Flog.writeline "No anda el getfile."
            Else
                Err.Number = 1
                Ciclos = Ciclos + 1
            End If
        End If
    Loop
    On Error GoTo 0
    Flog.writeline "Archivo encontrado: " & nombreArchivo
   
   'Abro el archivo
    On Error GoTo CE
    Set f = fs.OpenTextFile(nombreArchivo, ForReading, TristateFalse)
    
    str_error = ""
    nroLinea = 0
    Do While Not f.AtEndOfStream

        strLinea = f.ReadLine
        nroLinea = nroLinea + 1
        If nroLinea = 1 And usaencabezado Then
            strLinea = f.ReadLine
        End If
        
        If Trim(strLinea) <> "" Then
            Call import(strLinea, nroLinea, separador)
        End If
        
    Loop

    f.Close
                    
    FileCopy nombreArchivo, directorioBackup
    
    'Borrar el archivo
    fs.deletefile nombreArchivo, True

    Flog.writeline
    Flog.writeline "Archivo procesado: " & nombreArchivo & " " & Format(Now, "dd/mm/yyyy hh:mm:ss")
        
Fin:
    If rs_Lineas.State = adStateOpen Then rs_Lineas.Close
    Set rs_Lineas = Nothing
    Exit Sub
    
CE:
    HuboError = True
    
    MyRollbackTrans
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 0) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline Espacios(Tabulador * 0) & "Error: " & Err.Number
    Flog.writeline Espacios(Tabulador * 0) & "Descripcion: " & Err.Description
    Flog.writeline
    
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "SQL Ejecutado: " & StrSql
        Flog.writeline
    End If
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline
    GoTo Fin
End Sub


Public Function cambiaFecha(ByVal fecha As String) As String
' ---------------------------------------------------------------------------------------------
' Descripcion: Setea la fecha al formato de insercion de la base de datos.
' Autor      : Martin Ferraro
' Fecha      : 18/04/2009
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------


    If EsNulo(fecha) Then
        cambiaFecha = "NULL"
    Else
        cambiaFecha = ConvFecha(fecha)
    End If

End Function


'Obtiene el directorio configurado para el modelo
Public Function PathModelo(nroModelo)
 Dim directorio As String
 Dim rsAux As New ADODB.Recordset
 
    StrSql = "SELECT sis_direntradas FROM sistema WHERE sisnro = 1 "
    OpenRecordset StrSql, rsAux
    
    If Not rsAux.EOF Then
        directorio = Trim(rsAux!sis_direntradas)
    Else
        Flog.writeline "No esta configurado el directorio de sistema."
        Exit Function
    End If
    
    StrSql = "SELECT * FROM modelo WHERE modnro= " & nroModelo
    OpenRecordset StrSql, rsAux
    
    If Not rsAux.EOF Then
        directorio = directorio & Trim(rsAux!modarchdefault)
        Flog.writeline "Directorio del modelo: " & directorio
     Else
        Flog.writeline "No se encontró el modelo para obtener el directorio de lectura."
        Exit Function
    End If
        
    PathModelo = directorio
End Function


Public Function SeparadorModelo(nroModelo)
 Dim separador As String
 Dim rsAux As New ADODB.Recordset

    StrSql = "SELECT modseparador FROM modelo WHERE modnro= " & nroModelo
    OpenRecordset StrSql, rsAux
    
    If Not rsAux.EOF Then
        separador = Trim(rsAux!modseparador)
        Flog.writeline "Separador del modelo: " & separador
     Else
        Flog.writeline "No se encontró el modelo para obtener el directorio de lectura."
        Exit Function
    End If
        
    SeparadorModelo = separador
End Function

Public Function AsciiToUTF8(InputStr As String) As Byte()
Dim bytSrc() As Byte
Dim bytDest() As Byte
Dim I As Long

   bytSrc = InputStr
   ReDim bytDest(UBound(bytSrc) \ 2)
   For I = 0 To UBound(bytDest)
      bytDest(I) = bytSrc(I * 2)
   Next
   
   AsciiToUTF8 = bytDest
End Function

Public Function UTF8ToAscii(ByRef InputByt() As Byte) As String
Dim bytDest() As Byte
Dim I As Long

   ReDim bytDest(UBound(InputByt) * 2 + 1)
   For I = 0 To UBound(InputByt)
      bytDest(I * 2) = InputByt(I)
   Next
   
   UTF8ToAscii = CStr(bytDest)
End Function


