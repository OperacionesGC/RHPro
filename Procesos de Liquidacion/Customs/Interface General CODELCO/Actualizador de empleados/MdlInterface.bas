Attribute VB_Name = "MdlInterface"
Option Explicit

Global Const Version = "2.01"
Global Const FechaModificacion = "19/01/2006"
Global Const UltimaModificacion = " " 'Etapa 1 - Requerimiento 1
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
Global crpNro As Long
Global RegLeidos As Long
Global RegError As Long
Global RegFecha As Date
Global NroProceso As Long

Global f
'Global HuboError As Boolean
Global Path
Global NArchivo
Global NroLinea As Long
Global LineaCarga As Long

Global Separador As String
Global SeparadorDecimal As String
Global UsaEncabezado As Boolean

Global ErroresNov As Boolean

Global ErrCarga
Global LineaError
Global LineaOK

Global PisaNovedad As Boolean
Global Pisa As Boolean
Global TikPedNro As Long
Global NombreArchivo As String
Global acunro As Long 'se usa en el modelo 216 de Citrusvil y se carga por confrep
Global nro_ModOrg  As Long

Global NroModelo As Long
Global DescripcionModelo As String
Global Primera_Vez As Boolean
Global Banco As Long
Global usuario As String
Global EncontroAlguno As Boolean


Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial de Interface.
' Autor      : FGZ
' Fecha      : 29/07/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim objconnMain As New ADODB.Connection
Dim strCmdLine
Dim Nombre_Arch As String
Dim Nombre_Arch_Errores As String
Dim Nombre_Arch_Errores2 As String
Dim Nombre_Arch_Correctos As String
Dim rs_batch_proceso As New ADODB.Recordset
Dim bprcparam As String
Dim PID As String
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

    'Obtiene los datos de como esta configurado el servidor actualmente
    Call ObtenerConfiguracionRegional
    
    'Abro la conexion
    OpenConnection strconexion, objConn
    OpenConnection strconexion, objconnProgreso
    
    Nombre_Arch = PathFLog + "EmpleadosCODELCO_Lineas_con_Errores_" + Format(Now, "dd-mm-yyyy hh-mm-ss") + ".exc"
    Nombre_Arch_Errores = PathFLog + "EmpleadosCODELCO_Errores_de_Carga_" + Format(Now, "dd-mm-yyyy hh-mm-ss") + ".log"
    Nombre_Arch_Errores2 = PathFLog + "EmpleadosCODELCO_Lineas_Importadas_" + Format(Now, "dd-mm-yyyy hh-mm-ss") + ".ok"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    Set FlogE = fs.CreateTextFile(Nombre_Arch_Errores, True)
    Set FlogP = fs.CreateTextFile(Nombre_Arch_Errores2, True)
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Version = " & Version
    Flog.writeline "Modificacion = " & UltimaModificacion
    Flog.writeline "Fecha = " & FechaModificacion
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Numero, separador decimal    : " & NumeroSeparadorDecimal
    Flog.writeline "Numero, separador de miles   : " & NumeroSeparadorMiles
    Flog.writeline "Moneda, separador decimal    : " & MonedaSeparadorDecimal
    Flog.writeline "Moneda, separador de miles   : " & MonedaSeparadorMiles
    Flog.writeline "Formato de Fecha del Servidor: " & FormatoDeFechaCorto
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline
    Flog.writeline "PID = " & PID
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcprogreso = 0, bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 23 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    ErroresNov = False
    Primera_Vez = False
    tplaorden = 0
    
    If Not rs_batch_proceso.EOF Then
        bprcparam = rs_batch_proceso!bprcparam
        usuario = rs_batch_proceso!iduser
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        Call LevantarParamteros(bprcparam)
        LineaCarga = 0
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
    FlogE.Close
End Sub


Private Sub LeeArchivo(ByVal NombreArchivo As String)
Const ForReading = 1
Const TristateFalse = 0
Dim strlinea As String
Dim Archivo_Aux As String
Dim rs_Lineas As New ADODB.Recordset
Dim rs_Modelo As New ADODB.Recordset


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
   
Comienzo:

   'Abro el archivo
    On Error GoTo CE
    Set f = fs.OpenTextFile(NombreArchivo, ForReading, TristateFalse)
    
    NroLinea = 0
    RegLeidos = 0
    RegError = 0
    If Not f.AtEndOfStream Then
        StrSql = "INSERT INTO inter_pin(bpronro,modnro,crpnarchivo,crpnregleidos,crpnregerr,crpnfecha,crpndesc,crpnestado) VALUES ( " & _
                                      NroProcesoBatch & "," & NroModelo & ",'" & Left(NombreArchivo, 60) & "',0,0," & ConvFecha(Date) & ",'" & Left(DescripcionModelo, 18) & ": " & Date & "','I')"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        crpNro = getLastIdentity(objConn, "inter_pin")
    End If
                
    StrSql = "SELECT * FROM modelo WHERE modnro = " & NroModelo
    OpenRecordset StrSql, rs_Modelo
    If rs_Modelo.EOF Then
        Exit Sub
    End If
                
    StrSql = "SELECT * FROM modelo_filas WHERE bpronro =" & NroProcesoBatch
    StrSql = StrSql & " ORDER BY fila "
    OpenRecordset StrSql, rs_Lineas
    If Not rs_Lineas.EOF Then
        rs_Lineas.MoveFirst
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
        If NroLinea = 1 And UsaEncabezado Then
            strlinea = f.ReadLine
        
        End If
        If Trim(strlinea) <> "" And NroLinea = rs_Lineas!fila Then
            RegLeidos = RegLeidos + 1
            Flog.writeline Espacios(Tabulador * 0) & "Linea " & NroLinea
            Select Case rs_Modelo!modinterface
                
                Case 1: Call Insertar_Linea_Segun_Modelo_Estandar(strlinea)
                
                Case 2: Call Insertar_Linea_Segun_Modelo_Custom(strlinea)
                
                Case 3: Call Insertar_Linea_Segun_Modelo_MigraInicial(strlinea)
            
            End Select
            
            rs_Lineas.MoveNext
            
            'Como actualizo el progreso aca si no se cuantas lineas tiene el archivo
            'Incremento el progreso para que el servidor de aplicaciones no vea a este proceso
            'como colgado
            Progreso = Progreso + IncPorc
            StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & " WHERE bpronro = " & NroProcesoBatch
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
        End If
        
    Loop
    
    StrSql = "UPDATE inter_pin SET crpnregleidos = " & RegLeidos & _
             ",crpnregerr = " & RegError & _
             " WHERE crpnnro = " & crpNro
    objConn.Execute StrSql, , adExecuteNoRecords
    
    f.Close
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "Archivo procesado: " & NombreArchivo & " " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    
    'Borrar el archivo
    If NroModelo <> 650 Then
        fs.Deletefile NombreArchivo, True
    Else
        NroModelo = 653
        GoTo Comienzo
    End If
    
Fin:
    If rs_Lineas.State = adStateOpen Then rs_Lineas.Close
    Set rs_Lineas = Nothing
    Exit Sub
    
CE:
    Debug.Print StrSql
    Debug.Print Err.Description

    HuboError = True
    
    MyRollbackTrans
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 0) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline Espacios(Tabulador * 0) & "Error: " & Err.Number
    Flog.writeline Espacios(Tabulador * 0) & "Decripcion: " & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "Linea " & RegLeidos & " del archivo procesado"
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline
    GoTo Fin
End Sub


Public Sub LevantarParamteros(ByVal parametros As String)
Dim pos1 As Long
Dim pos2 As Long

Dim NombreArchivo1 As String
Dim NombreArchivo2 As String
Dim NombreArchivo3 As String


Separador = "@"
If Not IsNull(parametros) Then
    If Len(parametros) >= 1 Then

        'Nro de Modelo
        pos1 = 1
        pos2 = InStr(pos1, parametros, Separador) - 1
        NroModelo = Mid(parametros, pos1, pos2 - pos1 + 1)
        
        'Nombre del archivo a levantar
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        If pos2 > 0 Then
            NombreArchivo = Mid(parametros, pos1, pos2 - pos1 + 1)
        Else
            pos2 = Len(parametros)
            NombreArchivo = Mid(parametros, pos1, pos2 - pos1 + 1)
        End If
        
        'Dependiendo del modelo puede que vengan mas parametros
        Select Case NroModelo
        Case 211: 'Interface de Novedades
            pos1 = pos2 + 2
            pos2 = Len(parametros)
            PisaNovedad = CBool(Mid(parametros, pos1, pos2 - pos1 + 1))
        Case 212: 'GTI - Mega Alarmas
        Case 213: 'GTI - Acumulado Diario
        Case 214: 'Tickets
            pos1 = pos2 + 2
            pos2 = Len(parametros)
            TikPedNro = Mid(parametros, pos1, pos2 - pos1 + 1)
        Case 215: 'Interface de Acumuladores de Agencia
        Case 216: 'Interface de Acumuladores de Agencia para Citrusvil
        Case 217: 'Interface de Vales
        Case 218: 'Libre
        Case 219: 'Libre
        Case 220: 'Libre
        Case 221: 'Libre
        Case 222: 'Libre
        Case 223: 'Libre
        Case 224: 'Libre
        Case 225: 'Libre
        Case 226: 'Interface de Postulantes
        Case 227: 'Libre
        Case 228: 'Declaracion Jurada (LA ESTRELLA)
        Case 229: 'Interface de Prestamos
        Case 230: 'Interface de Pedidos de Vacaciones
        Case 231: 'Exportacion / Interface Banco Nacion
        Case 232: 'Interface Bumerang
        Case 233: 'Interface de Licencias
        Case 234: 'Exportacin JDE
        Case 235: 'Interface de Estadisticas de Accidentes
        Case 236: 'Interface de Bultos
        Case 239: 'Interfse Deloitte
            'Cargar_datos_deloitte
        Case 241: 'Interface Dabra
        Case 242: 'Interface SAP
        Case 243: 'Interface Cuentas Bancarias
        Case 244: '
        Case 245: '
        Case 246: '
        Case 247: 'Interface de Acumulado de Horas TELEPERFORMANCE
        Case 300:  'Migracion de Empleados para TELEPERFORMANCE ( + 3 columnas)
            Pisa = False
            
            'Logs de Error - Genera linea por linea, cual fue la que genera error
                NombreArchivo1 = PathFLog + "Empleados_Errores_de_Carga_" + Format(Now, "dd-mm-yyyy hh-mm-ss") + ".log"
                Set fs = CreateObject("Scripting.FileSystemObject")
                Set ErrCarga = fs.CreateTextFile(NombreArchivo1, True)

            'Logs de Lineas que no se cargaron
                NombreArchivo2 = PathFLog + "Empleados_Lineas_con_Errores_" + Format(Now, "dd-mm-yyyy hh-mm-ss") + ".exc"
                Set fs = CreateObject("Scripting.FileSystemObject")
                Set LineaError = fs.CreateTextFile(NombreArchivo2, True)

            'Logs de Lineas que se cargaron al sistema
                NombreArchivo3 = PathFLog + "Empleados_Lineas_Importadas_" + Format(Now, "dd-mm-yyyy hh-mm-ss") + ".ok"
                Set fs = CreateObject("Scripting.FileSystemObject")
                Set LineaOK = fs.CreateTextFile(NombreArchivo3, True)
        Case 630:
            'Logs de Error - Genera linea por linea, cual fue la que genera error
                NombreArchivo1 = PathFLog + "Estructuras_Errores_de_Carga_" + Format(Now, "dd-mm-yyyy hh-mm-ss") + ".log"
                Set fs = CreateObject("Scripting.FileSystemObject")
                Set ErrCarga = fs.CreateTextFile(NombreArchivo1, True)

            'Logs de Lineas que no se cargaron
                NombreArchivo2 = PathFLog + "Estructuras_Lineas_con_Errores_" + Format(Now, "dd-mm-yyyy hh-mm-ss") + ".exc"
                Set fs = CreateObject("Scripting.FileSystemObject")
                Set LineaError = fs.CreateTextFile(NombreArchivo2, True)

            'Logs de Lineas que se cargaron al sistema
                NombreArchivo3 = PathFLog + "Estructuras_Lineas_Importadas_" + Format(Now, "dd-mm-yyyy hh-mm-ss") + ".ok"
                Set fs = CreateObject("Scripting.FileSystemObject")
                Set LineaOK = fs.CreateTextFile(NombreArchivo3, True)

        Case 605:
'            pos1 = pos2 + 2
'            pos2 = Len(parametros)
'            Pisa = CBool(Mid(parametros, pos1, pos2 - pos1 + 1))
            Pisa = False
            
            'Logs de Error - Genera linea por linea, cual fue la que genera error
                NombreArchivo1 = PathFLog + "Empleados_Errores_de_Carga_" + Format(Now, "dd-mm-yyyy hh-mm-ss") + ".log"
                Set fs = CreateObject("Scripting.FileSystemObject")
                Set ErrCarga = fs.CreateTextFile(NombreArchivo1, True)

            'Logs de Lineas que no se cargaron
                NombreArchivo2 = PathFLog + "Empleados_Lineas_con_Errores_" + Format(Now, "dd-mm-yyyy hh-mm-ss") + ".exc"
                Set fs = CreateObject("Scripting.FileSystemObject")
                Set LineaError = fs.CreateTextFile(NombreArchivo2, True)

            'Logs de Lineas que se cargaron al sistema
                NombreArchivo3 = PathFLog + "Empleados_Lineas_Importadas_" + Format(Now, "dd-mm-yyyy hh-mm-ss") + ".ok"
                Set fs = CreateObject("Scripting.FileSystemObject")
                Set LineaOK = fs.CreateTextFile(NombreArchivo3, True)

        Case 600:
            'Logs de Error - Genera linea por linea, cual fue la que genera error
                NombreArchivo1 = PathFLog + "Familiares_Errores_de_Carga_" + Format(Now, "dd-mm-yyyy hh-mm-ss") + ".log"
                Set fs = CreateObject("Scripting.FileSystemObject")
                Set ErrCarga = fs.CreateTextFile(NombreArchivo1, True)

            'Logs de Lineas que no se cargaron
                NombreArchivo2 = PathFLog + "Familiares_Lineas_con_Errores_" + Format(Now, "dd-mm-yyyy hh-mm-ss") + ".exc"
                Set fs = CreateObject("Scripting.FileSystemObject")
                Set LineaError = fs.CreateTextFile(NombreArchivo2, True)

            'Logs de Lineas que se cargaron al sistema
                NombreArchivo3 = PathFLog + "Familiares_Lineas_Importadas_" + Format(Now, "dd-mm-yyyy hh-mm-ss") + ".ok"
                Set fs = CreateObject("Scripting.FileSystemObject")
                Set LineaOK = fs.CreateTextFile(NombreArchivo3, True)

        Case 645
            'Logs de Error - Genera linea por linea, cual fue la que genera error
                NombreArchivo1 = PathFLog + "ACMensuales_Errores_de_Carga_" + Format(Now, "dd-mm-yyyy hh-mm-ss") + ".log"
                Set fs = CreateObject("Scripting.FileSystemObject")
                Set ErrCarga = fs.CreateTextFile(NombreArchivo1, True)

            'Logs de Lineas que no se cargaron
                NombreArchivo2 = PathFLog + "ACMensuales_Lineas_con_Errores_" + Format(Now, "dd-mm-yyyy hh-mm-ss") + ".exc"
                Set fs = CreateObject("Scripting.FileSystemObject")
                Set LineaError = fs.CreateTextFile(NombreArchivo2, True)

            'Logs de Lineas que se cargaron al sistema
                NombreArchivo3 = PathFLog + "ACMensuales_Lineas_Importadas_" + Format(Now, "dd-mm-yyyy hh-mm-ss") + ".ok"
                Set fs = CreateObject("Scripting.FileSystemObject")
                Set LineaOK = fs.CreateTextFile(NombreArchivo3, True)
        
        Case 650
'               'Logs de Error - Genera linea por linea, cual fue la que genera error
'                NombreArchivo1 = PathFLog + "EmpleadosCODELCO_Errores_de_Carga_" + Format(Now, "dd-mm-yyyy hh-mm-ss") + ".log"
'                Set fs = CreateObject("Scripting.FileSystemObject")
'                Set ErrCarga = fs.CreateTextFile(NombreArchivo1, True)
'
'                'Logs de Lineas que no se cargaron
'                NombreArchivo2 = PathFLog + "EmpleadosCODELCO_Lineas_con_Errores_" + Format(Now, "dd-mm-yyyy hh-mm-ss") + ".exc"
'                Set fs = CreateObject("Scripting.FileSystemObject")
'                Set LineaError = fs.CreateTextFile(NombreArchivo2, True)
'
'                'Logs de Lineas que se cargaron al sistema
'                NombreArchivo3 = PathFLog + "EmpleadosCODELCO_Lineas_Importadas_" + Format(Now, "dd-mm-yyyy hh-mm-ss") + ".ok"
'                Set fs = CreateObject("Scripting.FileSystemObject")
'                Set LineaOK = fs.CreateTextFile(NombreArchivo3, True)

        Case 652
'               'Logs de Error - Genera linea por linea, cual fue la que genera error
'                NombreArchivo1 = PathFLog + "His_Estr_CODELCO_Errores_de_Carga_" + Format(Now, "dd-mm-yyyy hh-mm-ss") + ".log"
'                Set fs = CreateObject("Scripting.FileSystemObject")
'                Set ErrCarga = fs.CreateTextFile(NombreArchivo1, True)
'
'                'Logs de Lineas que no se cargaron
'                NombreArchivo2 = PathFLog + "His_Estr_CODELCO_Lineas_con_Errores_" + Format(Now, "dd-mm-yyyy hh-mm-ss") + ".exc"
'                Set fs = CreateObject("Scripting.FileSystemObject")
'                Set LineaError = fs.CreateTextFile(NombreArchivo2, True)
'
'                'Logs de Lineas que se cargaron al sistema
'                NombreArchivo3 = PathFLog + "His_Estr_CODELCO_Lineas_Importadas_" + Format(Now, "dd-mm-yyyy hh-mm-ss") + ".ok"
'                Set fs = CreateObject("Scripting.FileSystemObject")
'                Set LineaOK = fs.CreateTextFile(NombreArchivo3, True)
        End Select
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
        Flog.writeline Espacios(Tabulador * 1) & "No se encontró el registro de la tabla sistema nro 1"
        Exit Sub
    End If
    
    StrSql = "SELECT * FROM modelo WHERE modnro = " & NroModelo
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        Directorio = Directorio & Trim(objRs!modarchdefault)
        Separador = IIf(Not IsNull(objRs!modseparador), objRs!modseparador, ",")
        SeparadorDecimal = IIf(Not IsNull(objRs!modsepdec), objRs!modsepdec, ".")
        UsaEncabezado = IIf(Not IsNull(objRs!modencab), CBool(objRs!modencab), False)
        DescripcionModelo = objRs!moddesc
        
        Flog.writeline Espacios(Tabulador * 1) & "Directorio a buscar :  " & Directorio
     Else
        Flog.writeline Espacios(Tabulador * 1) & "No se encontró el modelo " & NroModelo
        Exit Sub
    End If
    
    'Algunos modelos no se comportan de la misma manera ==>
    Select Case NroModelo
'    Case 222:
'        Call LineaModelo_222
    Case Else
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
                Flog.writeline Espacios(Tabulador * 1) & "Procesando archivo " & archivo.Name
                Call LeeArchivo(Directorio & "\" & archivo.Name)
            End If
        Next
        If Not EncontroAlguno Then
            Flog.writeline Espacios(Tabulador * 1) & "No se encontró el archivo " & NombreArchivo
        End If
    End Select
End Sub

Public Sub InsertaError(NroCampo As Byte, nroError As Long)
    StrSql = "INSERT INTO inter_err(crpnnro,inerrnro,nrolinea,campnro) VALUES (" & _
             crpNro & "," & nroError & "," & NroLinea & "," & NroCampo & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
    
    RegError = RegError + 1
    ErroresNov = True
End Sub



Public Sub Escribir_Log(ByVal TipoLog As String, ByVal Lin As Long, ByVal Col As Long, ByVal msg As String, ByVal CantTab As Long, ByVal strlinea As String)
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
            FlogE.writeline Espacios(Tabulador * CantTab) & strlinea
    Case "FLOGP" 'Archivo de lineas procesadas
            FlogP.writeline Espacios(Tabulador * CantTab) & "Linea " & Lin & " Columna " & Col & ": " & msg
    Case Else
        Flog.writeline Espacios(Tabulador * CantTab) & "Nombre de archivo de log incorrecto " & TipoLog
End Select

End Sub

