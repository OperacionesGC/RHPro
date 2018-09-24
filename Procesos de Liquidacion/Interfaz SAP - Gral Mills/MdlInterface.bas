Attribute VB_Name = "MdlInterface"
Option Explicit


'Global Const Version = "1.00"
'Global Const FechaModificacion = "02/08/2011"
'Global Const UltimaModificacion = "Version inicial"
'http://www.microsoft.com/downloads/es-es/details.aspx?FamilyID=993c0bcf-3bcf-4009-be21-27e85e1857b1

'Global Const Version = "1.01"
'Global Const FechaModificacion = "18/10/2011"
'Global Const UltimaModificacion = "Correccion por sis_direntradas from sistema"

'Global Const Version = "1.02"
'Global Const FechaModificacion = "09/11/2011"
'Global Const UltimaModificacion = "Se agregaron algunos manejadores de errores. y Renombro el archivo procesado. Le agrego el nro de proceso."

'Global Const Version = "1.03"
'Global Const FechaModificacion = "15/11/2011"
'Global Const UltimaModificacion = "Se busca la estructura cc por cod ext. Se agrego el dni y se dio formato al cuil."

'Global Const Version = "1.04"
'Global Const FechaModificacion = "22/11/2011"
'                                 Gonzalez Nicolás
'Global Const UltimaModificacion = "Se cambio referencia xml de 6.0 a 3.0"

'Global Const Version = "1.05"
'Global Const FechaModificacion = "12/12/2011"
'                                 Gonzalez Nicolás
'Global Const UltimaModificacion = "Se agregó validación de null cuando valida nacionalnro en función insertar_tercero().Se corrigió mensaje de error en ErrorEmpleado:."

'Global Const Version = "1.06"
'Global Const FechaModificacion = "15/12/2011"
''                                 Gonzalez Nicolás
'Global Const UltimaModificacion = "Se valida que exista la estructura configurada por mapeo, en función insertar_Estructura()"

'Global Const Version = "1.07"
'Global Const FechaModificacion = "12/01/2012"
'Global Const UltimaModificacion = "Lisandro Moro - Correcciones: LOCALIDAD, NACIONALID, CONFIDENC, CODSINDIC, IMPUTACION"

'Global Const Version = "1.08"
'Global Const FechaModificacion = "31/01/2012"
'Global Const UltimaModificacion = "Lisandro Moro - Correcciones: Correcciones varias y mensajes de log."

'Global Const Version = "1.09"
'Global Const FechaModificacion = "24/02/2012"
'Global Const UltimaModificacion = "Lisandro Moro - Correcciones: Se corrigio. Se mejoro la captura de errores. Mejora al resolver l_IMPUTACION "

'Global Const Version = "1.10"
'Global Const FechaModificacion = "17/12/2013"
'Global Const UltimaModificacion = "Carmen Quintero - Correcciones: Varias (CAS- 22825 - General Mills - Bug - Interface SAP/Rhpro)"
                                 ' 1.- Se agregó validacion para los casos cuando ciertas etiquetas no se configuran en el archivo .xml
                                 ' 2.- Se agregó validacion para el campo Nacionalidad
                                 
'Global Const Version = "1.11"
'Global Const FechaModificacion = "28/01/2014"
'Global Const UltimaModificacion = "Carmen Quintero - Correcciones: Varias (CAS- 22825 - General Mills - Bug - Interface SAP/Rhpro)"
                                 ' 1.- Cuando hay bajas dentro del archivo, se toma la fecha indicada en la etiqueta f_ADATE del archivo y no la actual.

Global Const Version = "1.12"
Global Const FechaModificacion = "04/02/2014"
Global Const UltimaModificacion = "Carmen Quintero - Correcciones: Varias (CAS- 22825 - General Mills - Bug - Interface SAP/Rhpro)"
                                 ' 1.- Se agregó funcion ConvFecha.


'------------------------------------------------------------------------------------

Global crpNro As Long
Global RegLeidos As Long
Global RegError As Long
Global RegFecha As Date
Global NroProceso As Long

Global f
Global HuboErrorLocal As Boolean
Global Path
Global NArchivo
Global NroLinea As Long
'Global usuario As String

Global separador As String
Global SeparadorDecimal As String
Global UsaEncabezado As Boolean

Global ErroresNov As Boolean
Global NroModelo As Integer
Global DescripcionModelo As String
Global NombreArchivo As String

Dim directorio As String
Dim CArchivos
Dim Archivo
Dim Folder
Global cantArch As Integer
'--XML----------------------
'Global adoRS As ADODB.Recordset       'ADODB.Recordset

Public Sub Main()
    ' ---------------------------------------------------------------------------------------------
    ' Descripcion: Procedimiento inicial de Interface.
    ' Autor      : Lisandro Moro
    ' Fecha      : 12/07/2011
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
    
    'Abro la conexion
    OpenConnection strconexion, objConn
    OpenConnection strconexion, objconnProgreso
    
    Nombre_Arch = PathFLog & "Interface_SAP_GralMills " & "-" & NroProcesoBatch & ".log"
    'Archivo de log
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "PID = " & PID
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Version = " & Version
    Flog.writeline "Modificacion = " & UltimaModificacion
    Flog.writeline "Fecha = " & FechaModificacion
    Flog.writeline "-----------------------------------------------------------------"

    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcprogreso = 0, bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 304 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    ErroresNov = False
    'Primera_Vez = True
    
    If Not rs_batch_proceso.EOF Then
        'bprcparam = rs_batch_proceso!bprcparam
        usuario = rs_batch_proceso!iduser
        'rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        'Call LevantarParamteros(bprcparam)
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
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    objConn.Close
    objconnProgreso.Close
    Flog.Close
    End
End Sub


Private Sub LeeArchivo(ByVal NombreArchivo As String)
    Flog.writeline "Leer Archivo"
    'Const ForReading = 1
    'Const TristateFalse = 0
    'Dim strLinea As String
    'Dim Archivo_Aux As String
    'Dim rs_Lineas As New ADODB.Recordset
    
    'Const adStateOpen = &H1
    'Const adChapter = 136
    
    If App.PrevInstance Then
        Flog.writeline "Error: Instancia previa en ejecucion."
        Flog.writeline "Error: El Programa se cierra."
        End
    End If

    'Espero hasta que se crea el archivo
    On Error Resume Next
    Err.Number = 1
    Do Until Err.Number = 0
        Flog.writeline "Marca:" & Err.Number
        Err.Number = 0
        Set f = fs.GetFile(NombreArchivo)
        If f.Size = 0 Then Err.Number = 1
    Loop
   ' On Error GoTo 0

   'Abro el archivo
    Flog.writeline "Leyendo Archivo " & NombreArchivo
    '''Set f = fs.OpenTextFile(NombreArchivo, ForReading, TristateFalse)
    On Error Resume Next
    On Error GoTo 0
    On Error GoTo CE
    
    Dim doc As DOMDocument30
    Dim nodes As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim node2 As IXMLDOMNode
    

    Dim s As String

    Set doc = New DOMDocument30
    doc.Load NombreArchivo
    Set nodes = doc.selectNodes("//t_EMPLEADO")
    
    'Debug.Print nodes.length
    Flog.writeline "Leyendo Nodos " & nodes.length
    If nodes.length > 0 Then
        IncPorc = Round(CInt(CInt(100 / cantArch) / nodes.length), 0)
        'If IncPorc = 0 Then IncPorc = 1
        
        NroLinea = 0
        RegLeidos = 0
        RegError = 0
        
       'Call CargarDatosSapGralMills 'Borrar? si
        Dim a2 As Integer
        'Set nodes = doc.selectNodes("//t_EMPLEADO")
        'Set nodes = doc.selectNodes("//TABLE")
        
        For Each node In nodes 'doc.documentElement.childNodes
        'For a2 = 0 To nodes.length
            'Set node = nodes.Item(a2)
            NroLinea = NroLinea + 1
            RegLeidos = RegLeidos + 1
            'node2 = node.selectSingleNode("//t_EMPLEADO")
            
            Flog.writeline "Generar Empleados."
            Call Generar_Empleado(node)
            
            Progreso = Progreso + IncPorc
            If Progreso > 100 Then Progreso = 100
            StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & " WHERE bpronro = " & NroProcesoBatch
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
        Next
    Else
        Progreso = 100
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & " WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
    End If
    
    
'    'Set adoRS = CreateObject("ADODB.Recordset")
'    Set adoRS = New ADODB.Recordset
'    adoRS.ActiveConnection = "Provider=MSDAOSP; Data Source=MSXML6.DSOControl.2.6;"
'
'    'On Error Resume Next
'    'Abrimos el archivo
'    adoRS.Open NombreArchivo
    If Err Then
        Flog.writeline "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
        Flog.writeline "Error: " & Err.Number
        Flog.writeline "Decripcion: " & Err.Description
        Flog.writeline error
        GoTo Fin
    End If
    
    On Error GoTo CE
    
    ' inicializo los valores de bumeran
    
    
'    If Not adoRS.EOF Then
'        'Contamos los Postulantes
'        IncPorc = Round(CInt(100 / adoRS.RecordCount), 0)
'
'        NroLinea = 0
'        RegLeidos = 0
'        RegError = 0
'
'
'        StrSql = "INSERT INTO inter_pin(bpronro,modnro,crpnarchivo,crpnregleidos,crpnregerr,crpnfecha,crpndesc,crpnestado) VALUES ( " & _
'                                      NroProcesoBatch & "," & NroModelo & ",'" & Left(NombreArchivo, 60) & "',0,0," & ConvFecha(Date) & ",'" & Left(DescripcionModelo, 18) & ": " & Date & "','I')"
'        objConn.Execute StrSql, , adExecuteNoRecords
'
'        crpNro = getLastIdentity(objConn, "inter_pin")
'    End If

    'StrSql = "SELECT * FROM modelo_filas WHERE bpronro =" & NroProcesoBatch
    'StrSql = StrSql & " ORDER BY fila "
    'OpenRecordset StrSql, rs_Lineas
    'If Not rs_Lineas.EOF Then
    '    rs_Lineas.MoveFirst
    'End If
    
'    Do While Not rs_Lineas.EOF And NroLinea <= adoRS.RecordCount
        'strLinea = f.ReadLine
        'NroLinea = NroLinea + 1
        'If NroLinea > adoRS.RecordCount Then
            'strLinea = f.ReadLine
            'NroLinea = NroLinea + 1
            'rs_Lineas.MoveNext
        'End If
        'If Trim(strLinea) <> "" And NroLinea = rs_Lineas!fila Then
        'RegLeidos = RegLeidos + 1
        
'        If rs_Lineas("fila") = adoRS.AbsolutePosition Then
'            Call Insertar_Postulante_Segun_Modelo_Estandar(adoRS)
'            rs_Lineas.MoveNext
'        End If
'        adoRS.MoveNext
        'Como actualizo el progreso aca si no se cuantas lineas tiene el archivo
        'Incremento el progreso para que el servidor de aplicaciones no vea a este proceso
        'como colgado
'        Progreso = Progreso + IncPorc
 '       If Progreso > 100 Then Progreso = 100
 '       StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & " WHERE bpronro = " & NroProcesoBatch
 '       objconnProgreso.Execute StrSql, , adExecuteNoRecords
       'adoRS.MoveNext
'    Loop

    'Cierro el xml
 '   adoRS.Close
    
    
    StrSql = "UPDATE inter_pin SET crpnregleidos = " & RegLeidos & _
             ",crpnregerr = " & RegError & _
             " WHERE crpnnro = " & crpNro
    objConn.Execute StrSql, , adExecuteNoRecords

    'f.Close
    
    'fs.Deletefile NombreArchivo, True
    'RmDir (NombreArchivo)

Fin:
    'If rs_Lineas.State = adStateOpen Then rs_Lineas.Close
    'Set rs_Lineas = Nothing
    Set doc = Nothing
    Set nodes = Nothing
    Set node = Nothing
    Exit Sub
'
CE:
    HuboError = True
    MyRollbackTrans
    Flog.writeline "Error CE - Leer Archivo"
    Flog.writeline "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline "Error: " & Err.Number
    Flog.writeline "Decripcion: " & Err.Description
    Flog.writeline error
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
        
        'Dependiendo del modelo puede que vengan mas parametros
'        Select Case NroModelo
'        Case 211: 'Novedades
'            pos1 = pos2 + 2
'            pos2 = Len(parametros)
'            PisaNovedad = CBool(Mid(parametros, pos1, pos2 - pos1 + 1))
''        Case 212: 'GTI - Mega Alarmas
''        Case 213: 'GTI - Acumulado Diario
''            pos1 = pos2 + 2
''            pos2 = pos2 = Len(parametros)
''            PisaNovedad = CBool(Mid(parametros, pos1, pos2)) 'No se esta usando
'        Case 214: 'Tickets
'            pos1 = pos2 + 2
'            pos2 = Len(parametros)
'            TikPedNro = Mid(parametros, pos1, pos2 - pos1 + 1)
'        Case 215: 'Acumuladores de Agencia
''            pos1 = pos2 + 2
''            pos2 = Len(parametros)
''            TikPedNro = Mid(parametros, pos1, pos2)
'        Case 216: 'Acumuladores de Agencia para Citrusvil
''            pos1 = pos2 + 2
''            pos2 = Len(parametros)
''            TikPedNro = Mid(parametros, pos1, pos2)
'        Case 217: 'Vales
''            pos1 = pos2 + 2
''            pos2 = Len(parametros)
''            TikPedNro = Mid(parametros, pos1, pos2)
'        Case 218: 'Migracion de Novedades
'        End Select
    End If
End If

End Sub


Public Sub ComenzarTransferencia()
    
    'Leo los datos del Sistema
    StrSql = "SELECT sis_direntradas FROM sistema WHERE sisnro = 1 "
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        directorio = Trim(objRs!sis_direntradas)
    Else
        Flog.writeline "No se encontró el registro de la tabla sistema nro 1"
        Exit Sub
    End If
    
    If Right(directorio, 1) = "\" Then
        directorio = directorio & "SapGralMills"
    Else
        directorio = directorio & "\SapGralMills"
    End If
    'Leo los datos del modelo
    'StrSql = "SELECT * FROM modelo WHERE modnro = " & NroModelo
    'OpenRecordset StrSql, objRs
    'If Not objRs.EOF Then
    '    Directorio = Directorio & Trim(objRs!modarchdefault)
    '    separador = IIf(Not IsNull(objRs!modseparador), objRs!modseparador, ",")
    '    SeparadorDecimal = IIf(Not IsNull(objRs!modsepdec), objRs!modsepdec, ".")
    '    UsaEncabezado = IIf(Not IsNull(objRs!modencab), CBool(objRs!modencab), False)
    '    DescripcionModelo = objRs!moddesc
    '
        Flog.writeline "Directorio a buscar :  " & directorio
    'Else
    '    Flog.writeline "No se encontró el modelo " & NroModelo
    '    Exit Sub
    'End If
    
    'Algunos modelos no se comportan de la misma manera ==>
    'Select Case NroModelo
'    Case 222:
'        Call LineaModelo_222
    'Case Else
        'Set fs = CreateObject("Scripting.FileSystemObject")

        'Path = Directorio

        Dim fc, F1, s2
        
        Set Folder = fs.GetFolder(directorio)
        Set CArchivos = Folder.Files

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
        HuboErrorLocal = False
        cantArch = CArchivos.Count
        
        If CLng(cantArch) < 1 Then
            Flog.writeline "No se encontraron archivos a procesar."
        End If
        
        'On Error GoTo ErroroTransferencia
        
        For Each Archivo In CArchivos
            'If UCase(Right(archivo.Name, 4)) = ".CSV" Or UCase(Right(archivo.Name, 4)) = ".TXT" Then
        '    If UCase(archivo.Name) = UCase(NombreArchivo) Then
            If UBound(Split(CStr(Archivo.Name), ".")) > 0 Then
                If Split(CStr(Archivo.Name), ".")(1) = "xml" Then
                    NombreArchivo = CStr(Archivo.Name)
                    NArchivo = directorio & "\" & NombreArchivo
                    'NArchivo = archivo.Name
                    'MyBeginTrans
                    Flog.writeline "----------------------------------------------------------"
                    Flog.writeline "Archivo Procesado: " & NombreArchivo
                    Flog.writeline "----------------------------------------------------------"
                    HuboErrorLocal = False
                    
                    Call LeeArchivo(NArchivo)
                    'MyCommitTrans
                    
                    Flog.writeline "Archivo procesado: " & NombreArchivo & " " & Format(Now, "dd/mm/yyyy hh:mm:ss")
                    
                    'Borrar el archivo
                    If HuboErrorLocal Then
                        'Mantengo el archivo
                        Flog.writeline "Error: Se encontro un error al procesar el archivo." & NArchivo
                        Flog.writeline "Error: El archivo no se movera a la carpeta BK."
                    Else
                        If fs.FileExists(NArchivo) Then
                            
                            On Error Resume Next
                            If InStr(Folder, "/") > 0 Then
                                'fs.MoveFile NArchivo, Folder & "/bk/"
                                fs.MoveFile directorio & "/" & NombreArchivo, directorio & "/bk/" & Split(CStr(NombreArchivo), ".")(0) & "_" & NroProceso & "." & Split(CStr(NombreArchivo), ".")(1)
                            Else
                                'fs.MoveFile NArchivo, Folder & "\bk\" 'error si ya existie
                                'fs.CopyFile NArchivo, Folder & "\bk\", True
                                'fs.DeleteFile NArchivo
                                'fs.MoveFile Directorio & "\" & NombreArchivo, Directorio & "\bk\" & Split(CStr(NombreArchivo), ".")(0) & "_" & Replace(Format(Date, "yyyy-MM-DD"), "/", "-") & "-" & Replace(Format(Time, "hh:mm:ss"), ":", "-") & "." & Split(CStr(NombreArchivo), ".")(1)
                                fs.MoveFile directorio & "\" & NombreArchivo, directorio & "\bk\" & Split(CStr(NombreArchivo), ".")(0) & "_" & NroProceso & "." & Split(CStr(NombreArchivo), ".")(1)
                            End If
                            If Err.Number <> 0 Then
                                Flog.writeline "Error: Se produjo un error al querer mover el archivo." & NArchivo
                                Flog.writeline "Error: El archivo no se movera a la carpeta BK."
                                Flog.writeline "Error: " & Err.Description
                            Else
                                Flog.writeline "Moviendo archivo:" & NArchivo
                            End If
                            Err.Clear
                        End If
                    End If
                    HuboErrorLocal = False
                    
                End If
            End If
            On Error GoTo ErroroTransferencia
        Next
    'End Select
    Exit Sub
ErroroTransferencia:
    HuboError = True
    Flog.writeline "Error Transferencia."
    Flog.writeline "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline "Error: " & Err.Number
    Flog.writeline "Decripcion: " & Err.Description
    Flog.writeline error
    Flog.writeline "Linea " & RegLeidos & " del archivo procesado"
    Err.Clear
End Sub

Public Sub InsertaError(NroCampo As Byte, nroError As Long)
    StrSql = "INSERT INTO inter_err(crpnnro,inerrnro,nrolinea,campnro) VALUES (" & _
             crpNro & "," & nroError & "," & NroLinea & "," & NroCampo & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
    
    RegError = RegError + 1
    ErroresNov = True
End Sub


Public Sub Escribir_Log(ByVal TipoLog As String, ByVal Lin As Long, ByVal Col As Long, ByVal msg As String, ByVal CantTab As Long, ByVal strLinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Escribe un mensage determinado en uno de 3 archivos de log
' Autor      : FGZ
' Fecha      : 18/04/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Select Case UCase(TipoLog)
    Case "FLOG" 'Archivo de Informacion de resumen
            Flog.writeline Espacios(Tabulador * CantTab) & msg
    Case "FLOGE" 'Archivo de Errores
            FlogE.writeline Espacios(Tabulador * CantTab) & "Linea " & Lin & " Columna " & Col & ": " & msg
            FlogE.writeline Espacios(Tabulador * CantTab) & strLinea
    Case "FLOGP" 'Archivo de lineas procesadas
            FlogP.writeline Espacios(Tabulador * CantTab) & "Linea " & Lin & " Columna " & Col & ": " & msg
    Case Else
        Flog.writeline Espacios(Tabulador * CantTab) & "Nombre de archivo de log incorrecto " & TipoLog
End Select

End Sub

