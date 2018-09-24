Attribute VB_Name = "MdlInterface"
Option Explicit

'-----------------------------------------------------------------------------
'Global Const Version = "1.00"  'LM
'Global Const FechaModificacion = "16/04/2007"
'Global Const UltimaModificacion = "Nueva interfas para Universobit"

'Global Const Version = "1.01"  'LM
'Global Const FechaModificacion = "16/04/2007"
'Global Const UltimaModificacion = "Se corrigio el numero de tipo de proceso"

'Global Const Version = "1.02"  'LM
'Global Const FechaModificacion = "13/05/2008"
'Global Const UltimaModificacion = "Agregados de mensajes de log"

'Global Const Version = "1.03"  'FGZ
'Global Const FechaModificacion = "07/10/2008"
'Global Const UltimaModificacion = "Modificacion en sub Universo y leerArchivo para que actualice el id de postulantes ya levantados"
''                                   Se agregó una opcion mas en case "usuario id" para que tome el nro de id que tienen en
''                                   la BD de Universo y actualice en RHPro en pos_postulante.arepronro


'Global Const Version = "1.04"  'FGZ
'Global Const FechaModificacion = "02/09/2009"
'Global Const UltimaModificacion = " "
''                                   Se modificó el insert de empleos anteriores porque cuando no cargaban la empresa rompia


Global Const Version = "1.05"  'FGZ
Global Const FechaModificacion = "06/10/2009"
Global Const UltimaModificacion = " "
'                                   Se agregaron mensajes de log y se cambiaron algunos detalles



'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------

Global crpNro As Long
Global RegLeidos As Long
Global RegError As Long
Global RegFecha As Date
Global NroProceso As Long

Global f
Global HuboError As Boolean
Global Path
Global NArchivo
Global NroLinea As Long
Global usuario As String

Global separador As String
Global SeparadorDecimal As String
Global UsaEncabezado As Boolean

Global ErroresNov As Boolean
Global NroModelo As Long
Global DescripcionModelo As String
Global NombreArchivo As String
'--XML----------------------
Global adoRS As ADODB.Recordset       'ADODB.Recordset
 
'04/10/2004
'Dim objFeriado As New Feriado


Public Sub Main()
    ' ---------------------------------------------------------------------------------------------
    ' Descripcion: Procedimiento inicial de Interface.
    ' Autor      : Lisandro Moro
    ' Fecha      : 09/05/2007
    ' Ultima Mod.:
    ' Descripcion:
    ' ---------------------------------------------------------------------------------------------
    Dim objconnMain As New ADODB.Connection
    Dim strCmdLine
    Dim Nombre_Arch As String
    Dim rs_batch_proceso As New ADODB.Recordset
    Dim bprcparam As String
    Dim PID As String
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
    
    
    Nombre_Arch = PathFLog & "Migracion_Interface" & "-" & NroProcesoBatch & ".log"
    'Archivo de log
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Version = " & Version
    Flog.writeline "Modificacion = " & UltimaModificacion
    Flog.writeline "Fecha = " & FechaModificacion
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline
    Flog.writeline "PID = " & PID
    
    
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcprogreso = 0, bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 175 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    ErroresNov = False
    'Primera_Vez = True
    
    If Not rs_batch_proceso.EOF Then
        bprcparam = rs_batch_proceso!bprcparam
        usuario = rs_batch_proceso!iduser
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        Call LevantarParamteros(bprcparam)
        Call ComenzarTransferencia
    End If
    
    
Final:
    If Not HuboError Then
        If ErroresNov Then
            StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcprogreso = 100, bprcestado = 'Incompleto' WHERE bpronro = " & NroProcesoBatch
        Else
            StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcprogreso = 100, bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
        End If
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcprogreso = 100, bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
    End If
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    objConn.Close
    objconnProgreso.Close
    Flog.Close
    End


ME_Main:
    HuboError = True
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


Private Sub LeeArchivo(ByVal NombreArchivo As String)
    'Const ForReading = 1
    'Const TristateFalse = 0
    
    Const adStateOpen = &H1
    Const adChapter = 136
    
    
    Dim strLinea As String
    Dim Archivo_Aux As String
    Dim rs_Lineas As New ADODB.Recordset
    
    Dim Ciclos As Long
    
    
    If App.PrevInstance Then
        Flog.writeline "Hay una instancia previa del proceso corriendo. Este procesamiento se aborta."
        Exit Sub
    End If

    'Espero hasta que se crea el archivo
    Ciclos = 1000
    On Error Resume Next
    Err.Number = 1
    Do Until Err.Number = 0 Or Ciclos = 0
        Err.Number = 0
        Set f = fs.getfile(NombreArchivo)
        If f.Size = 0 Then
            Flog.writeline Espacios(Tabulador * 0) & "No anda el getfile "
            Err.Number = 1
            Ciclos = Ciclos - 1
        Else
            Ciclos = 0
        End If
    Loop
    If Ciclos = 0 And Err.Number <> 0 Then
        Flog.writeline Espacios(Tabulador * 0) & "No pudo leer el archivo. Procesamiento Abortado."
        Exit Sub
    End If
   ' On Error GoTo 0

   'Abro el archivo
    
    '''Set f = fs.OpenTextFile(NombreArchivo, ForReading, TristateFalse)
    'On Error Resume Next
    'Set adoRS = CreateObject("ADODB.Recordset")
    
    Set adoRS = New ADODB.Recordset
    adoRS.ActiveConnection = "Provider=MSDAOSP; Data Source=MSXML2.DSOControl.2.6;"
    
    'On Error Resume Next
    'Abrimos el archivo
    adoRS.Open NombreArchivo
    If Err Then
        Flog.writeline "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
        Flog.writeline "Error: " & Err.Number
        Flog.writeline "Decripcion: " & Err.Description
        Flog.writeline Error
        GoTo Fin
    End If
    
    On Error GoTo CE
    
    ' inicializo los valores de Universo
    'Call CargarDatosUniverso
    
    If Not adoRS.EOF Then
        'Contamos los Postulantes
        IncPorc = Round(CDbl(100 / adoRS.RecordCount), 0)
    
        NroLinea = 0
        RegLeidos = 0
        RegError = 0
    
        StrSql = "INSERT INTO inter_pin(bpronro,modnro,crpnarchivo,crpnregleidos,crpnregerr,crpnfecha,crpndesc,crpnestado) VALUES ( " & _
                                      NroProcesoBatch & "," & NroModelo & ",'" & Left(NombreArchivo, 60) & "',0,0," & ConvFecha(Date) & ",'" & Left(DescripcionModelo, 18) & ": " & Date & "','I')"
        objConn.Execute StrSql, , adExecuteNoRecords

        crpNro = getLastIdentity(objConn, "inter_pin")
    End If

    StrSql = "SELECT * FROM modelo_filas WHERE bpronro =" & NroProcesoBatch
    StrSql = StrSql & " ORDER BY fila "
    OpenRecordset StrSql, rs_Lineas
    If Not rs_Lineas.EOF Then
        rs_Lineas.MoveFirst
    End If
    
    'Obtengo el ultimo id ingresado
    Flog.writeline "Busco ultimo id postulante ingresado... "
    StrSql = "SELECT max(arepronro) ultimo FROM pos_postulante "
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        If IsNull(objRs!ultimo) Then
            ultimo_id = 0
        Else
            ultimo_id = (objRs!ultimo)
        End If
    Else
        ultimo_id = 0
    End If
    Flog.writeline "Ultimo id postulante ingresado: " & ultimo_id
    Flog.writeline
    
    'Do While Not rs_Lineas.EOF And NroLinea <= adoRS.RecordCount
        
    Do While Not rs_Lineas.EOF And NroLinea <= adoRS.RecordCount And Not adoRS.EOF
        RegLeidos = RegLeidos + 1
        
        If rs_Lineas("fila") = adoRS.AbsolutePosition Then
            If CLng(adoRS!ID) > ultimo_id Then
                Flog.writeline "Postulante nuevo ID " & adoRS!ID
                Call Insertar_Postulante_Segun_Modelo_Estandar(adoRS)
            'FGZ - 07/10/2008 - le agregré esto para arreglar los problemas de numeracion
            Else
                Flog.writeline "Postulante con ID " & adoRS!ID & " ya existe."
                ID = adoRS!ID
                Call Actualizar_Postulante_Segun_Modelo_Estandar(adoRS)
            End If
            rs_Lineas.MoveNext
        End If
        If Not adoRS.EOF Then
            adoRS.MoveNext
        End If
        
        'Como actualizo el progreso aca si no se cuantas lineas tiene el archivo
        'Incremento el progreso para que el servidor de aplicaciones no vea a este proceso
        'como colgado
        Progreso = Progreso + IncPorc
        If Progreso > 100 Then Progreso = 100
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & " WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
       'adoRS.MoveNext
    Loop

    'Cierro el xml
    If adoRS.State = adStateOpen Then adoRS.Close
    
    StrSql = "UPDATE inter_pin SET crpnregleidos = " & RegLeidos & _
             ",crpnregerr = " & RegError & _
             " WHERE crpnnro = " & crpNro
    objConn.Execute StrSql, , adExecuteNoRecords

    'f.Close
    Flog.writeline "Archivo procesado: " & NombreArchivo & " " & Format(Now, "dd/mm/yyyy hh:mm:ss")

    'Borrar el archivo
    fs.Deletefile NombreArchivo, True
    'RmDir (NombreArchivo)

Fin:
    If rs_Lineas.State = adStateOpen Then rs_Lineas.Close
    Set rs_Lineas = Nothing
    Exit Sub
'
CE:
    HuboError = True
    MyRollbackTrans
    Flog.writeline "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline "Error: " & Err.Number
    Flog.writeline "Decripcion: " & Err.Description
    Flog.writeline Error
    Flog.writeline "Linea " & RegLeidos & " del archivo procesado"
    'GoTo Fin

End Sub


Public Sub LevantarParamteros(ByVal parametros As String)
Dim pos1 As Integer
Dim pos2 As Integer


separador = "@"
If Not IsNull(parametros) Then
    If Len(parametros) >= 1 Then

        'Nro de Modelo
        pos1 = 1
        pos2 = InStr(pos1, parametros, separador) - 1
        NroModelo = Mid(parametros, pos1, pos2 - pos1 + 1)
        
        'Nombre del archivo a levantar
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, separador) - 1
        If pos2 > 0 Then
            NombreArchivo = Mid(parametros, pos1, pos2 - pos1 + 1)
        Else
            pos2 = Len(parametros)
            NombreArchivo = Mid(parametros, pos1, pos2 - pos1 + 1)
        End If
    End If
End If

End Sub


Public Sub ComenzarTransferencia()
    Dim Directorio As String
    Dim CArchivos
    Dim Archivo
    Dim Folder

    'Leo los datos del Sistema
    StrSql = "SELECT sis_direntradas FROM sistema WHERE sisnro = 1 "
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        Directorio = Trim(objRs!sis_direntradas)
    Else
        Flog.writeline "No se encontró el registro de la tabla sistema nro 1"
        Exit Sub
    End If
    
    'Leo los datos del modelo
    StrSql = "SELECT * FROM modelo WHERE modnro = " & NroModelo
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        Directorio = Directorio & Trim(objRs!modarchdefault)
        separador = IIf(Not IsNull(objRs!modseparador), objRs!modseparador, ",")
        SeparadorDecimal = IIf(Not IsNull(objRs!modsepdec), objRs!modsepdec, ".")
        UsaEncabezado = IIf(Not IsNull(objRs!modencab), CBool(objRs!modencab), False)
        DescripcionModelo = objRs!moddesc
        
        Flog.writeline "Directorio a buscar :  " & Directorio
     Else
        Flog.writeline "No se encontró el modelo " & NroModelo
        Exit Sub
    End If
    
    'Algunos modelos no se comportan de la misma manera ==>
    'Select Case NroModelo
'    Case 222:
'        Call LineaModelo_222
    'Case Else
        'Set fs = CreateObject("Scripting.FileSystemObject")

        'Path = Directorio

        Dim fc, F1, s2
        
        'Set Folder = fs.GetFolder(Directorio)
        'Set CArchivos = Folder.Files

        'Determino la proporcion de progreso
        Progreso = 0
        
        'If Not CArchivos.Count = 0 Then
        '    'Flog.writeline CArchivos.Count & " archivos a procesar " & Format(Now, "dd/mm/yyyy hh:mm:ss")
        '    CEmpleadosAProc = CArchivos.Count
        '    If CEmpleadosAProc = 0 Then
        '        CEmpleadosAProc = 1
        '    End If
        'End If
        'IncPorc = ((100 / CEmpleadosAProc) * (100 / 200)) / 100

        HuboError = False
        'For Each archivo In CArchivos
            'If UCase(Right(archivo.Name, 4)) = ".CSV" Or UCase(Right(archivo.Name, 4)) = ".TXT" Then
        '    If UCase(archivo.Name) = UCase(NombreArchivo) Then
                NArchivo = Directorio & "\" & NombreArchivo
                'NArchivo = archivo.Name
                'MyBeginTrans
                    Flog.writeline "Archivo Procesado: " & NombreArchivo
                    Call LeeArchivo(NArchivo)
                'MyCommitTrans
        '    End If
        'Next
    'End Select
End Sub

Public Sub InsertaError(NroCampo As Byte, nroError As Long)
    StrSql = "INSERT INTO inter_err(crpnnro,inerrnro,nrolinea,campnro) VALUES (" & _
             crpNro & "," & nroError & "," & NroLinea & "," & NroCampo & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
    
    RegError = RegError + 1
    ErroresNov = True
End Sub


