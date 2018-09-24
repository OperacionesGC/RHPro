Attribute VB_Name = "MdlExportacion"
Option Explicit

'Global Const Version = "1.01"
'Global Const FechaModificacion = "26/01/2006"
'Global Const UltimaModificacion = " " 'Etapa 1 - Requerimiento 7
Global Const Version = "1.02"
Global Const FechaModificacion = "23/05/2006"
Global Const UltimaModificacion = " " 'Mariano C. - Si el Garante es <> 1 then Garante =0
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------

Global IdUser As String
Global Fecha As Date
Global Hora As String


Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial del Proceso.
' Autor      : FGZ
' Fecha      : 26/01/2006
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim strCmdLine
Dim Nombre_Arch As String
Dim HuboError As Boolean
Dim rs_batch_proceso As New ADODB.Recordset
Dim PID As String
Dim bprcparam As String
Dim ArrParametros

    strCmdLine = Command()
    ArrParametros = Split(strCmdLine, " ", -1)
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
    
    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas

    'Abro la conexion
    OpenConnection strconexion, objConn
    OpenConnection strconexion, objconnProgreso
    
    Nombre_Arch = PathFLog & "Exp_Empleados_Codelco" & "-" & NroProcesoBatch & ".log"
    
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
    
    'Cambio el estado del proceso a Procesando y el PID
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 123 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        IdUser = rs_batch_proceso!IdUser
        Fecha = rs_batch_proceso!bprcfecha
        Hora = rs_batch_proceso!bprchora
        bprcparam = rs_batch_proceso!bprcparam
        
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        
        Call LevantarParamteros(NroProcesoBatch, bprcparam)
    End If
    
    TiempoFinalProceso = GetTickCount
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "=================================================="
    Flog.writeline Espacios(Tabulador * 0) & "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    Flog.writeline Espacios(Tabulador * 0) & "=================================================="
    
    If Not HuboError Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
    End If
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Flog.Close
    objconnProgreso.Close
    objConn.Close
    
End Sub

Public Sub Generacion_Datos(ByVal Bpronro As Long, ByVal Fecha_Hasta As Date, ByVal Modelo As Long, ByVal Todos As Boolean)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del archivo de Datos de los legajos
' Autor      : FGZ
' Fecha      : 26/01/2006
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Const ForReading = 1
Const TristateFalse = 0

Dim fExport
Dim fAuxiliar
Dim Directorio As String
Dim Archivo As String
Dim Intentos As Integer
Dim carpeta

Dim NroLiq As Long
Dim strLinea As String
Dim Aux_Linea As String
Dim Texto As String
Dim Separador As String
Dim RUT As String
Dim Estructura As String
Dim vbULTIMO_VALOR As String

'Registros
Dim rs_Modelo As New ADODB.Recordset
Dim rs_Detalles As New ADODB.Recordset
Dim rs_rep As New ADODB.Recordset
Dim rs_REP_A As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim rs_Fases As New ADODB.Recordset

'Archivo de exportacion
StrSql = "SELECT sis_dirsalidas FROM sistema WHERE sisnro = 1 "
OpenRecordset StrSql, rs
If Not rs.EOF Then
    Directorio = Trim(rs!sis_dirsalidas)
End If

StrSql = "SELECT * FROM modelo WHERE modnro = 650"
OpenRecordset StrSql, rs_Modelo
If Not rs_Modelo.EOF Then
    If Not IsNull(rs_Modelo!modarchdefault) Then
        Directorio = Directorio & Trim(rs_Modelo!modarchdefault)
        Separador = rs_Modelo!modseparador
    Else
        Flog.writeline Espacios(Tabulador * 1) & "El modelo no tiene configurada la carpeta desteino. El archivo será generado en el directorio default"
    End If
Else
    Flog.writeline Espacios(Tabulador * 1) & "No se encontró el modelo. El archivo será generado en el directorio default"
End If


'Seteo el nombre del archivo generado
Archivo = Directorio & "\Exp_Datos_" & Format(Now, "dd-mm-yyyy hh-mm-ss") & ".csv"
Set fs = CreateObject("Scripting.FileSystemObject")
'Activo el manejador de errores
On Error Resume Next
Set fExport = fs.CreateTextFile(Archivo, True)
If Err.Number <> 0 Then
    Flog.writeline Espacios(Tabulador * 1) & "La carpeta Destino no existe. Se creará."
    Set carpeta = fs.CreateFolder(Directorio)
    Err.Number = 0
    Set fExport = fs.CreateTextFile(Archivo, True)
End If
'desactivo el manejador de errores
On Error GoTo 0

On Error GoTo CE

'Busco los legajos a evaluar
StrSql = "SELECT * FROM empleado "
If Not Todos Then
    StrSql = StrSql & " INNER JOIN batch_empleado ON empleado.ternro = batch_empleado.ternro AND batch_empleado.bpronro = " & Bpronro
Else
    StrSql = StrSql & " WHERE empleado.empest = -1"
End If
StrSql = StrSql & " ORDER BY empleg "
OpenRecordset StrSql, rs_Detalles

'seteo de las variables de progreso
Progreso = 0
CConceptosAProc = rs_Detalles.RecordCount
If CConceptosAProc = 0 Then
    CConceptosAProc = 1
    Flog.writeline Espacios(Tabulador * 1) & " No hay empleados para Exportar. SQL " & StrSql
End If
IncPorc = (100 / CConceptosAProc)

    
    'Genero encabezado
    Flog.writeline Espacios(Tabulador * 2) & "Genero encabezado"
    Aux_Linea = "Legajo(Nro SAP)" & Separador & "Apellido" & Separador & "Apellido2" & Separador & "Nombre" & Separador & "Nombre2" & Separador & "Legajo Supervisor" & Separador & "Fec. Ingreso a Codelco" & Separador & "Fecha de Baja" & Separador & "RUT" & Separador & "E-mail"

    'mas las estructuras configuradas
    Flog.writeline Espacios(Tabulador * 2) & "mas las estructuras configuradas"
    StrSql = " SELECT * FROM confrep WHERE repnro = 120"
    StrSql = StrSql & " AND conftipo = 'TE'"
    StrSql = StrSql & " ORDER BY confnrocol"
    If rs_rep.State = adStateOpen Then rs_rep.Close
    OpenRecordset StrSql, rs_rep
    Do While Not rs_rep.EOF
    
        Aux_Linea = Aux_Linea & Separador & rs_rep!confetiq
        
        rs_rep.MoveNext
    Loop
    fExport.writeline Aux_Linea

    'Detalles
    Flog.writeline Espacios(Tabulador * 2) & "Detalles"
    Do While Not rs_Detalles.EOF
        Aux_Linea = rs_Detalles!empleg
        
        StrSql = "SELECT * FROM tercero WHERE tercero.ternro = " & rs_Detalles!ternro
        If rs.State = adStateOpen Then rs.Close
        OpenRecordset StrSql, rs
        If Not rs.EOF Then
            'nombre y apellidos
            Flog.writeline Espacios(Tabulador * 2) & "nombre y apellidos"
            Aux_Linea = Aux_Linea & Separador & rs!terape & Separador & IIf(Not EsNulo(rs!terape2), rs!terape2, "") & Separador & rs!ternom & Separador & IIf(Not EsNulo(rs!ternom2), rs!ternom2, "")
            
            'Supervisor
            Flog.writeline Espacios(Tabulador * 2) & "Supervisor"
            StrSql = "SELECT * FROM empleado where ternro=" & rs_Detalles!empreporta
            OpenRecordset StrSql, rs_REP_A
            'Aux_Linea = Aux_Linea & Separador & IIf(Not EsNulo(rs_Detalles!empreporta), rs_Detalles!empreporta, "")
            If EsNulo(rs_REP_A!empleg) = True Or rs_REP_A!empleg = "" Then
                Aux_Linea = Aux_Linea & Separador & " "
            Else
                Aux_Linea = Aux_Linea & Separador & rs_REP_A!empleg
            End If
            'Aux_Linea = Aux_Linea & Separador & IIf(Not EsNulo(rs_REP_A!empleg), rs_REP_A!empleg, "")
            'Fecha de ingreso
            Flog.writeline Espacios(Tabulador * 2) & "Fecha de Ingreso"
            StrSql = "SELECT * FROM fases WHERE empleado = " & rs_Detalles!ternro
            StrSql = StrSql & " AND altfec <= " & ConvFecha(Fecha)
            StrSql = StrSql & " ORDER BY altfec"
            If rs_Fases.State = adStateOpen Then rs_Fases.Close
            OpenRecordset StrSql, rs_Fases
            If Not rs_Fases.EOF Then
                rs_Fases.MoveLast
                Aux_Linea = Aux_Linea & Separador & Format(rs_Fases!altfec, "dd/mm/yyyy")
                
                'fecha de baja
                Flog.writeline Espacios(Tabulador * 2) & "Fecha de baja"
                Aux_Linea = Aux_Linea & Separador & IIf(Not EsNulo(rs_Fases!bajfec), Format(rs_Fases!bajfec, "dd/mm/yyyy"), "")
            Else
                Aux_Linea = Aux_Linea & Separador & "" & Separador & ""
            End If
            
            
            'RUT
            Flog.writeline Espacios(Tabulador * 2) & "RUT"
            StrSql = "SELECT * FROM tipodocu WHERE UPPER(tidsigla) = 'RUT'"
            If rs.State = adStateOpen Then rs.Close
            OpenRecordset StrSql, rs
            If Not rs.EOF Then
                StrSql = "SELECT * FROM ter_doc WHERE ternro = " & rs_Detalles!ternro
                StrSql = StrSql & " AND tidnro = " & rs!tidnro
                If rs.State = adStateOpen Then rs.Close
                OpenRecordset StrSql, rs
                If rs.EOF Then
                    RUT = ""
                Else
                    RUT = rs!nrodoc
                End If
            Else
                'Esto esta porque vi que la interface lo buscaba fijo
                StrSql = "SELECT * FROM ter_doc WHERE ternro = " & rs_Detalles!ternro
                StrSql = StrSql & " AND tidnro = 21"
                If rs.State = adStateOpen Then rs.Close
                OpenRecordset StrSql, rs
                If rs.EOF Then
                    RUT = ""
                Else
                    RUT = rs!nrodoc
                End If
            End If
            Aux_Linea = Aux_Linea & Separador & RUT
            
            'MAIL
            Flog.writeline Espacios(Tabulador * 2) & "E-MAIL"
            Aux_Linea = Aux_Linea & Separador & IIf(Not EsNulo(rs_Detalles!empemail), rs_Detalles!empemail, "")
        Else
            GoTo siguiente
        End If
        
        
        'busco las estructuras
        Flog.writeline Espacios(Tabulador * 2) & "Estructuras"
        rs_rep.MoveFirst
        Do While Not rs_rep.EOF
            'Busco el tipo de estructura a la fecha
            StrSql = " SELECT estrdabr FROM his_estructura "
            StrSql = StrSql & " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro "
            StrSql = StrSql & " WHERE his_estructura.ternro = " & rs_Detalles!ternro & " AND "
            StrSql = StrSql & " his_estructura.tenro =" & rs_rep!confval & " AND "
            StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(Fecha) & ") AND "
            StrSql = StrSql & " ((" & ConvFecha(Fecha) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
            StrSql = StrSql & " ORDER BY his_estructura.htetdesde"
            If rs.State = adStateOpen Then rs.Close
            OpenRecordset StrSql, rs
            If Not rs.EOF Then
                rs.MoveLast
                Estructura = rs!estrdabr
            Else
                Estructura = ""
            End If
        
            vbULTIMO_VALOR = Estructura
            Aux_Linea = Aux_Linea & Separador & Estructura
            rs_rep.MoveNext
        Loop
        
        If vbULTIMO_VALOR <> 1 Then Aux_Linea = Aux_Linea & 0
        Flog.writeline Espacios(Tabulador * 2) & "==========================================="
        ' ------------------------------------------------------------------------
        'Escribo en el archivo de texto
        fExport.writeline Aux_Linea
            
siguiente:
        'Actualizo el progreso del Proceso
        Progreso = Progreso + IncPorc
        TiempoAcumulado = GetTickCount
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
                 ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
                 "' WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
                
        'Siguiente proceso
        rs_Detalles.MoveNext
    Loop

Fin:
'Cierro el archivo creado
fExport.Close

If rs_Detalles.State = adStateOpen Then rs_Detalles.Close
Set rs_Detalles = Nothing

If rs_REP_A.State = adStateOpen Then rs_REP_A.Close
Set rs_REP_A = Nothing

If rs_Modelo.State = adStateOpen Then rs_Modelo.Close
Set rs_Modelo = Nothing

If rs_rep.State = adStateOpen Then rs_rep.Close
Set rs_rep = Nothing

If rs.State = adStateOpen Then rs.Close
Set rs = Nothing

If rs_Fases.State = adStateOpen Then rs_Fases.Close
Set rs_Fases = Nothing
Flog.writeline Espacios(Tabulador * 1) & "FIN"
Exit Sub
CE:
    Resume Next
    HuboError = True
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Description
    Flog.writeline Espacios(Tabulador * 1) & "Ultimo SQL: " & StrSql
    Flog.writeline
    Flog.writeline
    GoTo Fin
End Sub

Public Sub Generacion_Estructuras(ByVal Bpronro As Long, ByVal Fecha_Hasta As Date, ByVal Modelo As Long, ByVal Todos As Boolean)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del archivo de Historico de estructuras de los legajos
' Autor      : FGZ
' Fecha      : 26/01/2006
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Const ForReading = 1
Const TristateFalse = 0

Dim fExport
Dim fAuxiliar
Dim Directorio As String
Dim Archivo As String
Dim Intentos As Integer
Dim carpeta

Dim NroLiq As Long
Dim strLinea As String
Dim Aux_Linea As String
Dim Texto As String
Dim Separador As String
Dim RUT As String
Dim Estructura As String

'Registros
Dim rs_Modelo As New ADODB.Recordset
Dim rs_Detalles As New ADODB.Recordset
Dim rs_rep As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim rs_Fases As New ADODB.Recordset

'Archivo de exportacion
StrSql = "SELECT sis_dirsalidas FROM sistema WHERE sisnro = 1 "
OpenRecordset StrSql, rs
If Not rs.EOF Then
    Directorio = Trim(rs!sis_dirsalidas)
End If

StrSql = "SELECT * FROM modelo WHERE modnro = 652"
OpenRecordset StrSql, rs_Modelo
If Not rs_Modelo.EOF Then
    If Not IsNull(rs_Modelo!modarchdefault) Then
        Directorio = Directorio & Trim(rs_Modelo!modarchdefault)
        Separador = rs_Modelo!modseparador
    Else
        Flog.writeline Espacios(Tabulador * 1) & "El modelo no tiene configurada la carpeta desteino. El archivo será generado en el directorio default"
    End If
Else
    Flog.writeline Espacios(Tabulador * 1) & "No se encontró el modelo. El archivo será generado en el directorio default"
End If


'Seteo el nombre del archivo generado
Archivo = Directorio & "\Exp_Estructuras_" & Format(Now, "dd-mm-yyyy hh-mm-ss") & ".csv"
Set fs = CreateObject("Scripting.FileSystemObject")
'Activo el manejador de errores
On Error Resume Next
Set fExport = fs.CreateTextFile(Archivo, True)
If Err.Number <> 0 Then
    Flog.writeline Espacios(Tabulador * 1) & "La carpeta Destino no existe. Se creará."
    Set carpeta = fs.CreateFolder(Directorio)
    Err.Number = 0
    Set fExport = fs.CreateTextFile(Archivo, True)
End If
'desactivo el manejador de errores
On Error GoTo 0

On Error GoTo CE

'Busco los legajos a evaluar
StrSql = "SELECT * FROM empleado "
If Not Todos Then
    StrSql = StrSql & " INNER JOIN batch_empleado ON empleado.ternro = batch_empleado.ternro AND batch_empleado.bpronro = " & Bpronro
Else
    StrSql = StrSql & " WHERE empleado.empest = -1"
End If
StrSql = StrSql & " ORDER BY empleg "
OpenRecordset StrSql, rs_Detalles

'seteo de las variables de progreso
Progreso = 0
CConceptosAProc = rs_Detalles.RecordCount
If CConceptosAProc = 0 Then
    CConceptosAProc = 1
    Flog.writeline Espacios(Tabulador * 1) & " No hay empleados para Exportar. SQL " & StrSql
End If
IncPorc = (100 / CConceptosAProc)

    
    'Genero encabezado
    Flog.writeline Espacios(Tabulador * 2) & "Genero encabezado"
    Aux_Linea = "Tipo de Estructura" & Separador & "RUT" & Separador & "Estructura" & Separador & "Fecha Desde" & Separador & "Fecha Hasta"
    fExport.writeline Aux_Linea
    
    'mas las estructuras configuradas
    Flog.writeline Espacios(Tabulador * 2) & "mas las estructuras configuradas"
    StrSql = " SELECT * FROM confrep WHERE repnro = 156"
    StrSql = StrSql & " AND conftipo = 'TE'"
    StrSql = StrSql & " ORDER BY confnrocol"
    If rs_rep.State = adStateOpen Then rs_rep.Close
    OpenRecordset StrSql, rs_rep

    'Detalles
    Flog.writeline Espacios(Tabulador * 2) & "Detalles"
    Do While Not rs_Detalles.EOF
        'RUT
        Flog.writeline Espacios(Tabulador * 2) & "RUT"
        StrSql = "SELECT * FROM tipodocu WHERE UPPER(tidsigla) = 'RUT'"
        If rs.State = adStateOpen Then rs.Close
        OpenRecordset StrSql, rs
        If Not rs.EOF Then
            StrSql = "SELECT * FROM ter_doc WHERE ternro = " & rs_Detalles!ternro
            StrSql = StrSql & " AND tidnro = " & rs!tidnro
            If rs.State = adStateOpen Then rs.Close
            OpenRecordset StrSql, rs
            If rs.EOF Then
                RUT = ""
            Else
                RUT = rs!nrodoc
            End If
        Else
            'Esto esta porque vi que la interface lo buscaba fijo
            StrSql = "SELECT * FROM ter_doc WHERE ternro = " & rs_Detalles!ternro
            StrSql = StrSql & " AND tidnro = 21"
            If rs.State = adStateOpen Then rs.Close
            OpenRecordset StrSql, rs
            If rs.EOF Then
                RUT = ""
            Else
                RUT = rs!nrodoc
            End If
        End If
        
        'Estructuras
        Flog.writeline Espacios(Tabulador * 2) & "Estructuras"
        
        rs_rep.MoveFirst
        Do While Not rs_rep.EOF
            'Busco el tipo de estructura a la fecha
            StrSql = " SELECT his_estructura.tenro, tipoestructura.tedabr ,estructura.estrdabr, his_estructura.htetdesde, his_estructura.htethasta FROM his_estructura "
            StrSql = StrSql & " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro "
            StrSql = StrSql & " INNER JOIN tipoestructura ON tipoestructura.tenro = estructura.tenro "
            StrSql = StrSql & " WHERE his_estructura.ternro = " & rs_Detalles!ternro & " AND "
            StrSql = StrSql & " his_estructura.tenro =" & rs_rep!confval & " AND "
            StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(Fecha) & ") AND "
            StrSql = StrSql & " ((" & ConvFecha(Fecha) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
            StrSql = StrSql & " ORDER BY his_estructura.tenro, his_estructura.htetdesde"
            If rs.State = adStateOpen Then rs.Close
            OpenRecordset StrSql, rs
            If Not rs.EOF Then
                rs.MoveLast
                Estructura = rs!estrdabr
            
                'Tipo de Estructura
                Aux_Linea = rs!tedabr
                'RUT
                Aux_Linea = Aux_Linea & Separador & RUT
                'Estructura
                Aux_Linea = Aux_Linea & Separador & Estructura
                'Fecha desde
                Aux_Linea = Aux_Linea & Separador & IIf(Not EsNulo(rs!htetdesde), Format(rs!htetdesde, "dd/mm/yyyy"), "")
                'Fecha hasta
                Aux_Linea = Aux_Linea & Separador & IIf(Not EsNulo(rs!htethasta), Format(rs!htethasta, "dd/mm/yyyy"), "")
                'Escribo en el archivo de texto
                fExport.writeline Aux_Linea
            End If
        
            rs_rep.MoveNext
        Loop
        ' ------------------------------------------------------------------------
                
        'Actualizo el progreso del Proceso
        Progreso = Progreso + IncPorc
        TiempoAcumulado = GetTickCount
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
                 ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
                 "' WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
                
                
        'Siguiente proceso
        rs_Detalles.MoveNext
    Loop

Fin:
'Cierro el archivo creado
fExport.Close

If rs_Detalles.State = adStateOpen Then rs_Detalles.Close
Set rs_Detalles = Nothing

If rs_Modelo.State = adStateOpen Then rs_Modelo.Close
Set rs_Modelo = Nothing

If rs_rep.State = adStateOpen Then rs_rep.Close
Set rs_rep = Nothing

If rs.State = adStateOpen Then rs.Close
Set rs = Nothing

If rs_Fases.State = adStateOpen Then rs_Fases.Close
Set rs_Fases = Nothing
Flog.writeline Espacios(Tabulador * 1) & "FIN"
Exit Sub

CE:
    HuboError = True
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Description
    Flog.writeline Espacios(Tabulador * 1) & "Ultimo SQL: " & StrSql
    Flog.writeline
    Flog.writeline
    GoTo Fin
End Sub


Public Sub LevantarParamteros(ByVal Bpronro As Long, ByVal parametros As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento para levantar los parametros pasados en batch_proceso en bprcparam
' Autor      : FGZ
' Fecha      :
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim pos1 As Integer
Dim pos2 As Integer
Dim Separador As String

Dim Modelo As Long
Dim Fecha As Date
Dim aux As String
Dim Todos_Empleados As Boolean

Separador = "@"
' Levanto cada parametro por separado
If Not IsNull(parametros) Then
    If Len(parametros) >= 1 Then
    
        pos1 = 1
        pos2 = InStr(pos1, parametros, Separador) - 1
        Modelo = Mid(parametros, pos1, pos2 - pos1 + 1)
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        Todos_Empleados = CBool(Mid(parametros, pos1, pos2 - pos1 + 1))
        
        pos1 = pos2 + 2
        pos2 = Len(parametros)
        Fecha = CDate(Mid(parametros, pos1, pos2 - pos1 + 1))
    End If
End If
    Select Case Modelo
    Case 650:
        Call Generacion_Datos(Bpronro, Fecha, Modelo, Todos_Empleados)
    Case 652:
        Call Generacion_Estructuras(Bpronro, Fecha, Modelo, Todos_Empleados)
    Case Else
        Call Generacion_Datos(Bpronro, Fecha, Modelo, Todos_Empleados)
    End Select
End Sub

