Attribute VB_Name = "migradorEntidades"
Option Explicit

Global Const Version = "2.10"
Global Const FechaModificacion = "20/02/2006"
Global Const UltimaModificacion = "Inicial"

Dim fs, f
Global Flog

Dim NroLinea As Long
Dim crpNro As Long
Dim RegLeidos As Long
Dim RegError As Long
Dim RegFecha As Date
Dim NroProceso As Long

Global Path As String
Global NArchivo As String
Global Rta
Global HuboErrores As Boolean
Global EmpErrores As Boolean

Global NombreArchivo As String
Global tipo As Integer
Global entnro As Integer

Global NroColumna As Long
Global Tabulador As Long
Global Tabs As Long

Global EncontroAlguno As Boolean

Global IdUser As String
Global Fecha As Date
Global Hora As String



Private Sub Main()

Dim Directorio As String
Dim CArchivos
Dim archivo
Dim Folder
Dim strCmdLine As String
Dim Nombre_Arch As String

Dim StrSql As String
Dim objRs As New ADODB.Recordset
Dim rsPeriodos As New ADODB.Recordset
Dim tipoDepuracion
Dim historico As Boolean
Dim param
Dim rsConceptos As New ADODB.Recordset
Dim i
Dim totalAcum
Dim cantRegistros
Dim PID As String
Dim parametros As String
Dim ArrParametros
Dim concnro As Integer

    strCmdLine = Command()
    ArrParametros = Split(strCmdLine, " ", -1)
    If UBound(ArrParametros) > 0 Then
        If IsNumeric(ArrParametros(0)) Then
            NroProceso = ArrParametros(0)
            Etiqueta = ArrParametros(1)
        Else
            Exit Sub
        End If
    Else
        If IsNumeric(strCmdLine) Then
            NroProceso = strCmdLine
        Else
            Exit Sub
        End If
    End If
    
    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas

    TiempoInicialProceso = GetTickCount
    OpenConnection strconexion, objConn
    OpenConnection strconexion, objconnProgreso
    
    HuboErrores = False
    Tabulador = 5
    
    Nombre_Arch = PathFLog & "MigracionEntidades" & "-" & NroProceso & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.Writeline "-----------------------------------------------------------------"
    Flog.Writeline "Version = " & Version
    Flog.Writeline "Modificacion = " & UltimaModificacion
    Flog.Writeline "Fecha = " & FechaModificacion
    Flog.Writeline "-----------------------------------------------------------------"
    
    Flog.Writeline
    Flog.Writeline "PID = " & PID
    
    Flog.Writeline "Inicio Proceso de Migracion de Entidades : " & Now
    Flog.Writeline "Cambio el estado del proceso a Procesando"
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & ", bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Flog.Writeline "Obtengo los parametros del proceso"
    
    TiempoAcumulado = GetTickCount
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, objRs
    
    If Not objRs.EOF Then
       IdUser = objRs!IdUser
       Fecha = objRs!bprcfecha
       Hora = objRs!bprchora
       
       'Obtengo los parametros del proceso
       parametros = objRs!bprcparam
       ArrParametros = Split(parametros, "@")
       
       'Obtengo el tipo de migracion (importacion)
       ' 3 --> entity03.p
       ' 4 --> entity04.p
       ' 5 --> entity05.p
       ' 6 --> entity06.p
       tipo = ArrParametros(0)
       
       entnro = ArrParametros(1)
       
       NombreArchivo = ArrParametros(2)
       
       ' Proceso que migra (importa) los datos
       Call ComenzarTransferencia
       
    Else
        Exit Sub
    End If
   
    'Actualizo el estado del proceso
    If Not HuboErrores Then
       StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
       Flog.Writeline "Proceso Finalizado Correctamente"
    Else
       StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProceso
       Flog.Writeline "Proceso Incompleto"
    End If
    
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Flog.Writeline "Fin :" & Now
    Flog.Close

    Exit Sub
    
CE:
    HuboErrores = True
    Flog.Writeline " Error: " & Err.Description & Now

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
        Flog.Writeline Espacios(Tabulador * 1) & "No se encontró el registro de la tabla sistema nro 1"
        Exit Sub
    End If
    
    Directorio = Directorio & "\Migraciones"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
        
    Path = Directorio
        
    Dim fc, F1, s2
    Set Folder = fs.GetFolder(Directorio)
    Set CArchivos = Folder.Files
        
    HuboError = False
    EncontroAlguno = False
    For Each archivo In CArchivos
        EncontroAlguno = True
        If UCase(archivo.Name) = UCase(NombreArchivo) Then
            NArchivo = archivo.Name
            Flog.Writeline Espacios(Tabulador * 1) & "Procesando archivo: " & archivo.Name
            Call LeeArchivo(Directorio & "\" & archivo.Name)
        End If
    Next
    
    If Not EncontroAlguno Then
        Flog.Writeline Espacios(Tabulador * 1) & "No se encontró el archivo: " & NombreArchivo & " en el directorio " & Directorio
    End If
End Sub

Private Sub LeeArchivo(ByVal NombreArchivo As String)
Const ForReading = 1
Const TristateFalse = 0
Dim strlinea As String
Dim Archivo_Aux As String
Dim rs_Lineas As New ADODB.Recordset
Dim rs_Modelo As New ADODB.Recordset

    MyBeginTrans

    If App.PrevInstance Then
        Flog.Writeline Espacios(Tabulador * 0) & "Hay una instancia previa del proceso corriendo "
        Exit Sub
    End If
    'Espero hasta que se crea el archivo
    On Error Resume Next
    Err.Number = 1
    Do Until Err.Number = 0
        Err.Number = 0
        Set f = fs.getfile(NombreArchivo)
        If f.Size = 0 Then
            Flog.Writeline Espacios(Tabulador * 0) & "No anda el getfile "
            Err.Number = 1
        End If
    Loop
    On Error GoTo 0
    Flog.Writeline Espacios(Tabulador * 0) & "Archivo creado " & NombreArchivo
   
   'Abro el archivo
    On Error GoTo CE
    Set f = fs.OpenTextFile(NombreArchivo, ForReading, TristateFalse)
    
    NroLinea = 0
    RegLeidos = 0
    RegError = 0
    If f.AtEndOfStream Then
        Flog.Writeline Espacios(Tabulador * 0) & "No se pudo abrir el archivo " & NombreArchivo
    End If
                
    StrSql = "SELECT * FROM modelo_filas WHERE bpronro =" & NroProceso
    StrSql = StrSql & " ORDER BY fila "
    OpenRecordset StrSql, rs_Lineas
    If Not rs_Lineas.EOF Then
        rs_Lineas.MoveFirst
    Else
        Flog.Writeline Espacios(Tabulador * 0) & "No hay filas seleccionadas"
    End If
    
    'Determino la proporcion de progreso
    Progreso = 0
    CEmpleadosAProc = rs_Lineas.RecordCount
    If CEmpleadosAProc = 0 Then
        CEmpleadosAProc = 1
    End If
    IncPorc = (99 / CEmpleadosAProc)
    
    Do While Not f.AtEndOfStream And Not rs_Lineas.EOF
        strlinea = f.ReadLine
        NroLinea = NroLinea + 1
        If Trim(strlinea) <> "" And NroLinea = rs_Lineas!fila Then
            Select Case tipo
                Case 3:
                    Call Insertar_Linea_Segun_entity03(strlinea)
                    RegLeidos = RegLeidos + 1
                Case 4:
                    Call Insertar_Linea_Segun_entity04(strlinea)
                    RegLeidos = RegLeidos + 1
                Case 6:
                    Call Insertar_Linea_Segun_entity06(strlinea)
                    RegLeidos = RegLeidos + 1
                Case Else
                    Flog.Writeline Espacios(Tabulador * 0) & "El Prog. de Migración no esta configurado. Estan configurados los 3,4 y 6"
            End Select
            
            'Actualizo el progreso
            Progreso = Progreso + IncPorc
            StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & " WHERE bpronro = " & NroProceso
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
        End If
        
        rs_Lineas.MoveNext
            
    Loop
    
    f.Close
    Flog.Writeline
    Flog.Writeline Espacios(Tabulador * 0) & "Archivo procesado: " & NombreArchivo & " " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    
    'Borrar el archivo
    fs.Deletefile NombreArchivo, True
    
    
    MyCommitTrans
    
Fin:
    If rs_Lineas.State = adStateOpen Then rs_Lineas.Close
    Set rs_Lineas = Nothing
    Exit Sub
    
CE:
    HuboError = True
    
    MyRollbackTrans
    Flog.Writeline
    Flog.Writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.Writeline Espacios(Tabulador * 0) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.Writeline Espacios(Tabulador * 0) & "Error: " & Err.Number
    Flog.Writeline Espacios(Tabulador * 0) & "Decripcion: " & Err.Description
    Flog.Writeline
    Flog.Writeline Espacios(Tabulador * 0) & "Linea: " & strlinea
    Flog.Writeline Espacios(Tabulador * 0) & "Linea " & RegLeidos & " del archivo procesado"
    If InStr(1, Err.Description, "ODBC") > 0 Then
        Flog.Writeline
        Flog.Writeline Espacios(Tabulador * 0) & "SQL Ejecutado: " & StrSql
        Flog.Writeline
    End If
    Flog.Writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.Writeline
    GoTo Fin
End Sub

'--------------------------------------------------------------------
' Migracion de datos segun entity03.p. Solo carga los datos de las columnas 3,5, y 7
' <codetosend> <tab><tab> <aditifield1> <tab><tab> <adtifield2> [ <tab><tab> <aditifield3> [ <tab><tab> <aditifield4> ]]
'--------------------------------------------------------------------
Private Sub Insertar_Linea_Segun_entity03(ByVal strLin As String)
    
Dim codetosend As String
Dim aditfield1 As String
Dim aditfield2 As String
Dim aditfield3 As String
Dim aditfield4 As String
Dim Separador As String
Dim pos1 As Integer
Dim pos2 As Integer
Dim seguir As Boolean
Dim rs_entity_value As New ADODB.Recordset

    On Error GoTo Manejador_De_Error
    
    Tabs = 1
    seguir = True
    Separador = Chr("9")
    strLin = strLin & Chr("9")
    
    Texto = "Procesando linea: " & strLin
    Flog.Writeline Espacios(Tabulador * Tabs) & Texto
    
    aditfield1 = ""
    aditfield2 = ""
    aditfield3 = ""
    
    Tabs = 2
    
    NroColumna = NroColumna + 1
    ' Primer columna: codetosend
    pos1 = 1
    pos2 = InStr(pos1, strLin, Separador)
    If pos2 = 0 Then
        Flog.Writeline Espacios(Tabulador * Tabs) & "La primera columna no esta definida."
        GoTo Fin
    Else
        codetosend = Mid(strLin, pos1, pos2 - pos1)
    End If

    NroColumna = NroColumna + 1
    ' Segunda Columna: Vacia. Se saltea
    pos1 = pos2 + 1
    pos2 = InStr(pos1, strLin, Separador)
    If pos2 = 0 Then
        Flog.Writeline Espacios(Tabulador * Tabs) & "La tercera columna no esta definida."
        GoTo Fin
    End If
    
    NroColumna = NroColumna + 1
    ' Tercera columna: aditfield1
    pos1 = pos2 + 1
    pos2 = InStr(pos1, strLin, Separador)
    If pos2 = 0 Then
        Flog.Writeline Espacios(Tabulador * Tabs) & "La tercera columna no esta definida."
        GoTo Fin
    Else
        aditfield1 = Mid(strLin, pos1, pos2 - pos1)
        If aditfield1 = "" Then
            Flog.Writeline Espacios(Tabulador * Tabs) & "La tercera columna esta en blanco."
            GoTo Fin
        End If
    End If
    
    NroColumna = NroColumna + 1
    ' Cuarta Columna: Vacia. Se saltea
    pos1 = pos2 + 1
    pos2 = InStr(pos1, strLin, Separador)
    If pos2 = 0 Then
        seguir = False
    End If
    
    If seguir Then
        NroColumna = NroColumna + 1
        ' Quinta Columna: aditfield2
        pos1 = pos2 + 1
        pos2 = InStr(pos1, strLin, Separador)
        If pos2 = 0 Then
            seguir = False
        Else
            aditfield2 = Mid(strLin, pos1, pos2 - pos1)
        End If
    End If

    If seguir Then
        NroColumna = NroColumna + 1
        ' Sexta Columna: Vacia. Se saltea
        pos1 = pos2 + 1
        pos2 = InStr(pos1, strLin, Separador)
        If pos2 = 0 Then
            seguir = False
        End If
    End If
    
    If seguir Then
        NroColumna = NroColumna + 1
        ' Septima Columna: aditfield3
        pos1 = pos2 + 1
        pos2 = InStr(pos1, strLin, Separador)
        If pos2 = 0 Then
            seguir = False
        Else
            aditfield3 = Mid(strLin, pos1, pos2 - pos1)
        End If
    End If

    ' ====================================================================
    ' Validar los parametros Levantados
    ' ====================================================================
    ' Verifico si existe el codetosend
    StrSql = "SELECT * FROM entity_value WHERE codetosend = '" & codetosend & "' AND entnro = " & entnro
    OpenRecordset StrSql, rs_entity_value
    If rs_entity_value.EOF Then
        StrSql = "INSERT INTO entity_value (codetosend,entnro,aditfield1,aditfield2,aditfield3,aditfield4) "
        StrSql = StrSql & " VALUES ('" & codetosend & "'," & entnro & ",'" & aditfield1 & "','" & aditfield2 & "','" & aditfield3 & "','')"
        objConn.Execute StrSql, , adExecuteNoRecords
        Flog.Writeline Espacios(Tabulador * Tabs) & "Se insertaron los datos."
    Else
        StrSql = "UPDATE entity_value SET aditfield1 = '" & aditfield1 & "'"
        StrSql = StrSql & ",aditfield2 = '" & aditfield2 & "'"
        StrSql = StrSql & ",aditfield3 = '" & aditfield3 & "'"
        StrSql = StrSql & ",aditfield4 = ''"
        StrSql = StrSql & " WHERE codetosend = '" & codetosend & "' AND entnro = " & entnro
        objConn.Execute StrSql, , adExecuteNoRecords
        Flog.Writeline Espacios(Tabulador * Tabs) & "Se actualizaron los datos."
    End If
    
Fin:
'Cierro todo y libero
If rs_entity_value.State = adStateOpen Then rs_entity_value.Close

Set rs_entity_value = Nothing
Exit Sub

Manejador_De_Error:
    HuboError = True

    Flog.Writeline
    Flog.Writeline Espacios(Tabulador * 1) & "Error insalvable en la linea " & strLin
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.Writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.Writeline Espacios(Tabulador * 1) & "Descripción: " & Err.Description
    Flog.Writeline
    Flog.Writeline
    Flog.Writeline Espacios(Tabulador * 1) & "SQL Ejecutada: " & StrSql
    Flog.Writeline
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline
    
    GoTo Fin
End Sub

'--------------------------------------------------------------------
' Migracion de datos segun entity04.p. Solo carga los datos de las columnas 5,7 y 9
'--------------------------------------------------------------------
Private Sub Insertar_Linea_Segun_entity04(ByVal strLin As String)
    
Dim codetosend As String
Dim aditfield1 As String
Dim aditfield2 As String
Dim aditfield3 As String
Dim aditfield4 As String
Dim Separador As String
Dim pos1 As Integer
Dim pos2 As Integer
Dim seguir As Boolean
Dim rs_entity_value As New ADODB.Recordset

    On Error GoTo Manejador_De_Error
    
    Tabs = 1
    seguir = True
    Separador = Chr("9")
    strLin = strLin & Chr("9")
    
    Texto = "Procesando linea: " & strLin
    Flog.Writeline Espacios(Tabulador * Tabs) & Texto
    
    aditfield1 = ""
    aditfield2 = ""
    aditfield3 = ""
    aditfield4 = ""
    
    Tabs = 2
    
    NroColumna = NroColumna + 1
    ' Primer columna: aditfield1
    pos1 = 1
    pos2 = InStr(pos1, strLin, Separador)
    If pos2 = 0 Then
        Flog.Writeline Espacios(Tabulador * Tabs) & "La primera columna no esta definida."
        GoTo Fin
    Else
        aditfield1 = Mid(strLin, pos1, pos2 - pos1)
        If aditfield1 = "" Then
            Flog.Writeline Espacios(Tabulador * Tabs) & "La Primer columna esta en blanco."
            GoTo Fin
        End If
    End If
    
    NroColumna = NroColumna + 1
    ' Segunda Columna: Vacia. Se saltea
    pos1 = pos2 + 1
    pos2 = InStr(pos1, strLin, Separador)
    If pos2 = 0 Then
        Flog.Writeline Espacios(Tabulador * Tabs) & "La tercera columna no esta definida."
        GoTo Fin
    End If
    
    NroColumna = NroColumna + 1
    ' Tercera columna: codetosend
    pos1 = pos2 + 1
    pos2 = InStr(pos1, strLin, Separador)
    If pos2 = 0 Then
        Flog.Writeline Espacios(Tabulador * Tabs) & "La tercera columna no esta definida."
        GoTo Fin
    Else
        codetosend = Mid(strLin, pos1, pos2 - pos1)
    End If

    NroColumna = NroColumna + 1
    ' Cuarta Columna: Vacia. Se saltea
    pos1 = pos2 + 1
    pos2 = InStr(pos1, strLin, Separador)
    If pos2 = 0 Then
        seguir = False
    End If
    
    If seguir Then
        NroColumna = NroColumna + 1
        ' Quinta Columna: aditfield2
        pos1 = pos2 + 1
        pos2 = InStr(pos1, strLin, Separador)
        If pos2 = 0 Then
            seguir = False
        Else
            aditfield2 = Mid(strLin, pos1, pos2 - pos1)
        End If
    End If

    If seguir Then
        NroColumna = NroColumna + 1
        ' Sexta Columna: Vacia. Se saltea
        pos1 = pos2 + 1
        pos2 = InStr(pos1, strLin, Separador)
        If pos2 = 0 Then
            seguir = False
        End If
    End If
    
    If seguir Then
        NroColumna = NroColumna + 1
        ' Septima Columna: aditfield3
        pos1 = pos2 + 1
        pos2 = InStr(pos1, strLin, Separador)
        If pos2 = 0 Then
            seguir = False
        Else
            aditfield3 = Mid(strLin, pos1, pos2 - pos1)
        End If
    End If

    If seguir Then
        NroColumna = NroColumna + 1
        ' Octava Columna: Vacia. Se saltea
        pos1 = pos2 + 1
        pos2 = InStr(pos1, strLin, Separador)
        If pos2 = 0 Then
            seguir = False
        End If
    End If
    
    If seguir Then
        NroColumna = NroColumna + 1
        ' Novena Columna: aditfield4
        pos1 = pos2 + 1
        pos2 = InStr(pos1, strLin, Separador)
        If pos2 = 0 Then
            seguir = False
        Else
            aditfield4 = Mid(strLin, pos1, pos2 - pos1)
        End If
    End If

    ' ====================================================================
    ' Validar los parametros Levantados
    ' ====================================================================
    ' Verifico si existe el codetosend
    StrSql = "SELECT * FROM entity_value "
    StrSql = StrSql & "WHERE codetosend = '" & codetosend & "' AND aditfield1 = '" & aditfield1 & "' AND entnro = " & entnro
    OpenRecordset StrSql, rs_entity_value
    
    If rs_entity_value.EOF Then
        StrSql = "INSERT INTO entity_value (codetosend,entnro,aditfield1,aditfield2,aditfield3,aditfield4) "
        StrSql = StrSql & " VALUES ('" & codetosend & "'," & entnro & ",'" & aditfield1 & "','" & aditfield2 & "','" & aditfield3 & "','" & aditfield4 & "')"
        objConn.Execute StrSql, , adExecuteNoRecords
        Flog.Writeline Espacios(Tabulador * Tabs) & "Se insertaron los datos."
    Else
        StrSql = "UPDATE entity_value aditfield2 = '" & aditfield2 & "'"
        StrSql = StrSql & ",aditfield3 = '" & aditfield3 & "'"
        StrSql = StrSql & ",aditfield4 = '" & aditfield4 & "'"
        StrSql = StrSql & " WHERE codetosend = '" & codetosend & "'"
        StrSql = StrSql & " AND aditfield1 = '" & aditfield1 & "' AND entnro = " & entnro
        objConn.Execute StrSql, , adExecuteNoRecords
        Flog.Writeline Espacios(Tabulador * Tabs) & "Se actualizaron los datos."
    End If
    
Fin:
'Cierro todo y libero
If rs_entity_value.State = adStateOpen Then rs_entity_value.Close

Set rs_entity_value = Nothing
Exit Sub

Manejador_De_Error:
    HuboError = True

    Flog.Writeline
    Flog.Writeline Espacios(Tabulador * 1) & "Error insalvable en la linea " & strLin
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.Writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.Writeline Espacios(Tabulador * 1) & "Descripción: " & Err.Description
    Flog.Writeline
    Flog.Writeline
    Flog.Writeline Espacios(Tabulador * 1) & "SQL Ejecutada: " & StrSql
    Flog.Writeline
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline
    
    GoTo Fin
End Sub

'--------------------------------------------------------------------
' Migracion de datos segun entity06.p. Solo carga los datos de las columnas 2,3,4, y 5
'--------------------------------------------------------------------
Private Sub Insertar_Linea_Segun_entity06(ByVal strLin As String)
    
Dim codetosend As String
Dim aditfield1 As String
Dim aditfield2 As String
Dim aditfield3 As String
Dim aditfield4 As String
Dim Separador As String
Dim pos1 As Integer
Dim pos2 As Integer
Dim seguir As Boolean
Dim rs_entity_value As New ADODB.Recordset

    On Error GoTo Manejador_De_Error
    
    Tabs = 1
    seguir = True
    Separador = Chr("9")
    strLin = strLin & Chr("9")
    
    Texto = "Procesando linea: " & strLin
    Flog.Writeline Espacios(Tabulador * Tabs) & Texto
    
    aditfield1 = ""
    aditfield2 = ""
    aditfield3 = ""
    
    Tabs = 2
    
    NroColumna = NroColumna + 1
    ' Primer columna: codetosend
    pos1 = 1
    pos2 = InStr(pos1, strLin, Separador)
    If pos2 = 0 Then
        Flog.Writeline Espacios(Tabulador * Tabs) & "La primer columna no esta definida."
        GoTo Fin
    Else
        codetosend = Mid(strLin, pos1, pos2 - pos1)
    End If

    NroColumna = NroColumna + 1
    ' Segunda columna: aditfield1
    pos1 = pos2 + 1
    pos2 = InStr(pos1, strLin, Separador)
    If pos2 = 0 Then
        Flog.Writeline Espacios(Tabulador * Tabs) & "La segunda columna no esta definida."
        GoTo Fin
    Else
        aditfield1 = Mid(strLin, pos1, pos2 - pos1)
        If aditfield1 = "" Then
            Flog.Writeline Espacios(Tabulador * Tabs) & "La segunda columna esta en blanco."
            GoTo Fin
        End If
    End If
    
    If seguir Then
        NroColumna = NroColumna + 1
        ' Tercera Columna: aditfield2
        pos1 = pos2 + 1
        pos2 = InStr(pos1, strLin, Separador)
        If pos2 = 0 Then
            seguir = False
        Else
            aditfield2 = Mid(strLin, pos1, pos2 - pos1)
        End If
    End If

    If seguir Then
        NroColumna = NroColumna + 1
        ' Cuarta Columna: aditfield3
        pos1 = pos2 + 1
        pos2 = InStr(pos1, strLin, Separador)
        If pos2 = 0 Then
            seguir = False
        Else
            aditfield3 = Mid(strLin, pos1, pos2 - pos1)
        End If
    End If

    If seguir Then
        NroColumna = NroColumna + 1
        ' Quinta Columna: aditfield4
        pos1 = pos2 + 1
        pos2 = InStr(pos1, strLin, Separador)
        If pos2 = 0 Then
            seguir = False
        Else
            aditfield4 = Mid(strLin, pos1, pos2 - pos1)
        End If
    End If

    ' ====================================================================
    ' Validar los parametros Levantados
    ' ====================================================================
    ' Verifico si existe el codetosend
    StrSql = "SELECT * FROM entity_value WHERE codetosend = '" & codetosend & "' AND entnro = " & entnro
    OpenRecordset StrSql, rs_entity_value
    If rs_entity_value.EOF Then
        StrSql = "INSERT INTO entity_value (codetosend,entnro,aditfield1,aditfield2,aditfield3,aditfield4) "
        StrSql = StrSql & " VALUES ('" & codetosend & "'," & entnro & ",'" & aditfield1 & "','" & aditfield2 & "','" & aditfield3 & "','" & aditfield4 & "')"
        objConn.Execute StrSql, , adExecuteNoRecords
        Flog.Writeline Espacios(Tabulador * Tabs) & "Se insertaron los datos."
    Else
        StrSql = "UPDATE entity_value SET aditfield1 = '" & aditfield1 & "'"
        StrSql = StrSql & ",aditfield2 = '" & aditfield2 & "'"
        StrSql = StrSql & ",aditfield3 = '" & aditfield3 & "'"
        StrSql = StrSql & ",aditfield4 = '" & aditfield4 & "'"
        StrSql = StrSql & " WHERE codetosend = '" & codetosend & "' AND entnro = " & entnro
        objConn.Execute StrSql, , adExecuteNoRecords
        Flog.Writeline Espacios(Tabulador * Tabs) & "Se actualizaron los datos."
    End If
    
Fin:
'Cierro todo y libero
If rs_entity_value.State = adStateOpen Then rs_entity_value.Close

Set rs_entity_value = Nothing
Exit Sub

Manejador_De_Error:
    HuboError = True

    Flog.Writeline
    Flog.Writeline Espacios(Tabulador * 1) & "Error insalvable en la linea " & strLin
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline Espacios(Tabulador * 1) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.Writeline Espacios(Tabulador * 1) & "Error: " & Err.Number
    Flog.Writeline Espacios(Tabulador * 1) & "Descripción: " & Err.Description
    Flog.Writeline
    Flog.Writeline
    Flog.Writeline Espacios(Tabulador * 1) & "SQL Ejecutada: " & StrSql
    Flog.Writeline
    Flog.Writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.Writeline
    
    GoTo Fin
End Sub


