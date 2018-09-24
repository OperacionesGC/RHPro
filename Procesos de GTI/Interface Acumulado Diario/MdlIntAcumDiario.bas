Attribute VB_Name = "MdlIntAcumDiario"
Option Explicit

Global NroProcesoBatch As Long
Global exito As Boolean

Global fs
Global Flog
Global rs As New ADODB.Recordset

'Variables de Progreso
Global CEmpleadosAProc As Integer
Global CConceptosAProc As Integer
Global IncPorc As Single
Global IncPorcEmpleado As Single
Global Progreso As Single

Global NombreArchivo As String
Global NroLinea As Long
Dim crpNro As Long
Dim RegLeidos As Long
Dim RegError As Long
Dim RegFecha As Date
Dim NroProceso As Long

Global f
Global HuboError As Boolean
Global PisaNovedad As Boolean
Global Path
Global NArchivo
Global Separador As String
Global UsaEncabezado As Boolean
Global ErroresNov As Boolean

Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long


Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial de Interface
' Autor      : FGZ
' Fecha      : 16/01/2003
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim objconnMain As New ADODB.Connection
Dim strCmdLine
Dim Nombre_Arch As String
Dim rs_batch_proceso As New ADODB.Recordset
Dim bprcparam As String
Dim PID As String

' carga las configuraciones basicas, formato de fecha, string de conexion,
' tipo de BD y ubicacion del archivo de log
Call CargarConfiguracionesBasicas
    
'Abro la conexion
    OpenConnection strconexion, objConn
    
    strCmdLine = Command()
    If IsNumeric(strCmdLine) Then
        NroProcesoBatch = strCmdLine
    Else
        Exit Sub
    End If
    
    Nombre_Arch = PathFLog & "InterfaceAcumDiario" & "-" & NroProcesoBatch & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "PID = " & PID
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objConn.Execute StrSql, , adExecuteNoRecords
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 52 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    ErroresNov = False
    
    If Not rs_batch_proceso.EOF Then
        bprcparam = rs_batch_proceso!bprcparam
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        Call LevantarParamteros(bprcparam)
        Call ComenzarTransferencia
    End If
    
    If Not HuboError Then
        If ErroresNov Then
            StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcprogreso = 100, bprcestado = 'Incompleto' WHERE bpronro = " & NroProcesoBatch
        Else
            StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcprogreso = 100, bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
        End If
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcprogreso = 100, bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
    End If
    objConn.Execute StrSql, , adExecuteNoRecords
    
    objConn.Close
    Flog.Close

End Sub


Private Sub LeeArchivo(ByVal NombreArchivo As String)
Const ForReading = 1
Const TristateFalse = 0
Dim strLinea As String
Dim Archivo_Aux As String

    If App.PrevInstance Then Exit Sub

    'Espero hasta que se crea el archivo
    On Error Resume Next
    Err.Number = 1
    Do Until Err.Number = 0
        Err.Number = 0
        Set f = fs.getfile(NombreArchivo)
        If f.Size = 0 Then Err.Number = 1
    Loop
    On Error GoTo 0
   
   'Abro el archivo
    On Error GoTo CE
    Set f = fs.OpenTextFile(NombreArchivo, ForReading, TristateFalse)
    
    NroLinea = 0
    RegLeidos = 0
    RegError = 0
    If Not f.AtEndOfStream Then
        StrSql = "INSERT INTO inter_pin(bpronro,modnro,crpnarchivo,crpnregleidos,crpnregerr,crpnfecha,crpndesc,crpnestado) VALUES ( " & _
                                      NroProcesoBatch & ",213,'" & NombreArchivo & "',0,0," & ConvFecha(Date) & ",'Interface GTI_AcumDiario: " & Date & "','I')"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        crpNro = getLastIdentity(objConn, "inter_pin")
    End If
                
    Do While Not f.AtEndOfStream
        strLinea = f.ReadLine
        NroLinea = NroLinea + 1
        If NroLinea = 1 And UsaEncabezado Then
            strLinea = f.ReadLine
            NroLinea = NroLinea + 1
        End If
        If Trim(strLinea) <> "" Then
            Call InsertarLinea(strLinea)
        End If
        
        'Como actualizo el progreso aca si no se cuantas lineas tiene el archivo
        'Incremento el progreso para que el servidor de aplicaciones no vea a este proceso
        'como colgado
        Progreso = Progreso + IncPorc
        If Progreso > 100 Then Progreso = 99
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & " WHERE bpronro = " & NroProcesoBatch
        objConn.Execute StrSql, , adExecuteNoRecords
        
    Loop
    
    StrSql = "UPDATE inter_pin SET crpnregleidos = " & RegLeidos & _
             ",crpnregerr = " & RegError & _
             " WHERE crpnnro = " & crpNro
    objConn.Execute StrSql, , adExecuteNoRecords
    
    f.Close
    
    Flog.writeline "Archivo procesado: " & NombreArchivo & " " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    
    'Borrar el archivo
    fs.Deletefile NombreArchivo, True
    
fin:
    Exit Sub
    
CE:
    HuboError = True
    
    Flog.writeline "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline "Error: " & Err.Number
    Flog.writeline "Decripcion: " & Err.Description
    
    GoTo fin
        
End Sub


Public Sub LevantarParamteros(ByVal parametros As String)
Dim pos1 As Integer
Dim pos2 As Integer


If Not IsNull(parametros) Then
    If Len(parametros) >= 1 Then

        'Pisa o no
        pos1 = 1
        pos2 = Len(parametros)
        PisaNovedad = CBool(Mid(parametros, pos1, pos2))

'        Pos1 = 1
'        Pos2 = InStr(Pos1, parametros, ".") - 1
'        PisaNovedad = Mid(parametros, Pos1, Pos2)
        
'        pos1 = pos2 + 2
'        pos2 = InStr(pos1, Parametros, ".") - 1
'        Mantener_Liq = CBool(Mid(Parametros, pos1, pos2 - pos1 + 1))
        
'        Pos1 = Pos2 + 2
'        Pos2 = Len(parametros)
'        HACE_TRAZA = CBool(Mid(parametros, Pos1, Pos2 - Pos1 + 1))
        
    End If
End If

End Sub



Public Sub InsertarLinea(ByVal strLinea As String)
Dim pos1 As Integer
Dim pos2 As Integer
    
Dim tercero As Long
Dim NroLegajo As Long
Dim thnro As Long
Dim Cantidad As Single
Dim FechaAD As Date

Dim rs_Empleado As New ADODB.Recordset
Dim rs_TipHora As New ADODB.Recordset
Dim rs_GTI_AcumDiario As New ADODB.Recordset


' El formato es:
' Legajo; Fecha; Thnro; cantidad

    'Nro de Legajo
    pos1 = 1
    pos2 = InStr(pos1, strLinea, Separador)
    If IsNumeric(Mid$(strLinea, pos1, pos2 - pos1)) Then
        NroLegajo = Mid$(strLinea, pos1, pos2 - pos1)
    Else
        InsertaError 1, 8
        Exit Sub
    End If
    
    'Fecha
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strLinea, Separador)
    FechaAD = Mid(strLinea, pos1, pos2 - pos1)

    'Tipo de Hora
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strLinea, Separador)
    thnro = Mid(strLinea, pos1, pos2 - pos1)

    'Cantidada
    pos1 = pos2 + 1
    pos2 = Len(strLinea)
    Cantidad = Mid(strLinea, pos1, pos2)

' ====================================================================
'   Validar los parametros Levantados

'Que exista el legajo
StrSql = "SELECT * FROM empleado where empleg = " & NroLegajo
OpenRecordset StrSql, rs_Empleado
If rs_Empleado.EOF Then
    Flog.writeline "No se encontro el legajo " & NroLegajo
    InsertaError 1, 8
    Exit Sub
Else
    tercero = rs_Empleado!Ternro
End If

'Que la fecha sea valida
If Not IsDate(FechaAD) Then
    Flog.writeline "Facha no Valida " & FechaAD
    InsertaError 2, 4
    Exit Sub
End If
'Que exista el tipo de hora
StrSql = "SELECT * FROM tiphora WHERE thnro = " & thnro
OpenRecordset StrSql, rs_TipHora
If rs_TipHora.EOF Then
    Flog.writeline "No se encontro el tipo de hora " & thnro
    InsertaError 3, 37
    Exit Sub
End If

'Que sea numerico la cantidad
If Not IsNumeric(Cantidad) Then
    Flog.writeline "la cantidad no es numerica " & Cantidad
    InsertaError 4, 38
    Exit Sub
End If

'=============================================================
'Busco si existe
StrSql = "SELECT * FROM gti_acumdiario WHERE " & _
         " thnro = " & thnro & _
         " AND ternro = " & tercero & _
         " AND adfecha = " & ConvFecha(FechaAD)
OpenRecordset StrSql, rs_GTI_AcumDiario

If rs_GTI_AcumDiario.EOF Then
        StrSql = "INSERT INTO gti_acumdiario (" & _
                 "ternro,thnro,adfecha,adcanthoras,admanual,advalido,adestado" & _
                 ") VALUES (" & tercero & _
                 "," & thnro & _
                 "," & ConvFecha(FechaAD) & _
                 "," & Cantidad & _
                 ",0,0,'L'" & _
                 " )"
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.writeline "Acum Diario insertado "
Else
    StrSql = "SELECT * FROM gti_acumdiario WHERE " & _
             " thnro = " & thnro & _
             " AND ternro = " & tercero & _
             " AND admanual = 0 AND adValido = 0 AND adestado = 'L'"
    If rs_GTI_AcumDiario.State = adStateOpen Then rs_GTI_AcumDiario.Close
    OpenRecordset StrSql, rs_GTI_AcumDiario
    
    If Not rs_GTI_AcumDiario.EOF Then
            StrSql = "UPDATE gti_acumdiario SET adcanthoras = " & Cantidad & _
                     " WHERE thnro = " & thnro & _
                     " AND ternro = " & tercero & _
                     " AND admanual = 0 AND adValido = 0 AND adestado = 'L'"
        objConn.Execute StrSql, , adExecuteNoRecords
        Flog.writeline "Acum Diario Actualizado "
    Else
        Flog.writeline "El Acum Diario ya existe "
        InsertaError 1, 54
        Exit Sub
    End If
End If

End Sub




Public Sub ComenzarTransferencia()
Dim Directorio As String
Dim CArchivos
Dim archivo
Dim Folder

    StrSql = "SELECT sis_direntradas FROM sistema WHERE sisnro = 1 "
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        Directorio = Trim(objRs!sis_direntradas)
    Else
        Exit Sub
    End If
    
    Flog.writeline "Directorio de Acum Diario:  " & Directorio
    
    StrSql = "SELECT * FROM modelo WHERE modnro = 213 "
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        Directorio = Directorio & Trim(objRs!modarchdefault)
        Separador = objRs!modseparador
        UsaEncabezado = CBool(objRs!modencab)
     Else
        Exit Sub
    End If
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    Path = Directorio
    
    Dim fc, F1, s2
    Set Folder = fs.GetFolder(Directorio)
    Set CArchivos = Folder.Files
    
    'Determino la proporcion de progreso
    Progreso = 0
    If Not CArchivos.Count = 0 Then
        Flog.writeline CArchivos.Count & " Archivos de Acum Diario " & Format(Now, "dd/mm/yyyy hh:mm:ss")
        CEmpleadosAProc = CArchivos.Count
        If CEmpleadosAProc = 0 Then
            CEmpleadosAProc = 1
        End If
    End If
    IncPorc = ((100 / CEmpleadosAProc) * (100 / 200)) / 100
    
    
    HuboError = False
    For Each archivo In CArchivos
        If UCase(Right(archivo.Name, 4)) = ".CSV" Or UCase(Right(archivo.Name, 4)) = ".TXT" Then
            NArchivo = archivo.Name
            Call LeeArchivo(Directorio & "\" & archivo.Name)
        End If
    Next
    
End Sub

Private Sub InsertaError(NroCampo As Byte, nroError As Long)
    StrSql = "INSERT INTO inter_err(crpnnro,inerrnro,nrolinea,campnro) VALUES (" & _
             crpNro & "," & nroError & "," & NroLinea & "," & NroCampo & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
    
    RegError = RegError + 1
    ErroresNov = True
End Sub


