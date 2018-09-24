Attribute VB_Name = "MdlLiquidador"
Option Explicit

Global USA_DEBUG As Boolean
Global Tiempo As Date
Global TiempoTotal As Double
Global Usa_grossing As Boolean
Global Termino_gross As Boolean
Global v_netofijo As Double
Global v_nroitera As Long
Global v_nroConce As Long
Global val As Double
Global fec As Date
Global Actualizo_Bae As Boolean
Global ReusaTraza As Boolean
Global Const MaxIteraGross = 20
Global borrarProc As Boolean
'Global Usa_Estadisticas As Boolean
'Global Nueva_Tpa As Boolean
Global Usa_Nov_Dist As Boolean

Private Declare Function Sleep Lib _
   "kernel32" (ByVal dwMilliseconds As Long) As Long



Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: procedimiento inicial del liquidador.
' Autor      : FGZ
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim objconnMain As New ADODB.Connection
Dim strCmdLine
Dim Grupo As Long
Dim FechaInicio As Date
Dim FechaFin As Date
Dim Nombre_Arch As String
Dim rs_Batch_Proceso As New ADODB.Recordset
Dim rs_His_Batch_Proceso As New ADODB.Recordset
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
    
    'FGZ - 13/08/2004
    Cantidad_de_OpenRecordset = 0
    Borrar_Estadisticas = True
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Nombre_Arch = PathFLog & "Tiempos_Liq" & "-" & NroProcesoBatch & ".log"
        Set Flog = fs.CreateTextFile(Nombre_Arch, True)
                      
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If
    
    OpenConnection strconexion, objconnProgreso
    If Err.Number <> 0 Or Error_Encrypt Then
        Nombre_Arch = PathFLog & "Tiempos_Liq" & "-" & NroProcesoBatch & ".log"
        Set Flog = fs.CreateTextFile(Nombre_Arch, True)
        
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If
    On Error GoTo 0
    
    StrSql = "SELECT iduser, bprctipomodelo, bpronro, bprcparam FROM batch_proceso WHERE btprcnro = 3 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_Batch_Proceso
       
    
    If Not fs.FolderExists(PathFLog & rs_Batch_Proceso!iduser) Then fs.CreateFolder (PathFLog & rs_Batch_Proceso!iduser)
    Nombre_Arch = PathFLog & rs_Batch_Proceso!iduser & "\" & "Tiempos_Liq" & "-" & NroProcesoBatch & ".log"
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Version      = " & Version
    Flog.writeline "Modificacion = " & UltimaModificacion
    Flog.writeline "Modificacion = " & UltimaModificacion1
    Flog.writeline "Modificacion = " & UltimaModificacion2
    Flog.writeline "Fecha        = " & FechaModificacion
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline
    Flog.writeline "PID = " & PID
    
    On Error Resume Next
    '----------------------------------MDF ver 6.63
    'Abro la conexion
    'OpenConnection strconexion, objConn ---MDF
    'If Err.Number <> 0 Or Error_Encrypt Then
    '    Flog.writeline "Problemas en la conexion"
    '    Exit Sub
    'End If
    'OpenConnection strconexion, objconnProgreso
    'If Err.Number <> 0 Or Error_Encrypt Then
    '    Flog.writeline "Problemas en la conexion"
    '    Exit Sub
    'End If
    'On Error GoTo 0
    '----------------------------------MDF ver 6.63
    'FGZ - 27/04/2011 --------- Control de versiones ------
    Version_Valida = ValidarV(Version, 3, TipoBD)
    If Not Version_Valida Then
        'Actualizo el progreso
        MyBeginTransLiq
            StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Error de Version', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
        MyCommitTransLiq
        Flog.writeline
        GoTo Fin
    End If
    'FGZ - 05/08/2009 --------- Control de versiones ------
    
    On Error GoTo ME_Main
    
    'FGZ - 22/11/2013 --------------------------
    Usa_Estadisticas = CalculoEstadisticoActivo(3)
    'FGZ - 22/11/2013 --------------------------
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    'Obtengo los datos del proceso
    'FGZ - 05/06/2012 ---------------------------------------------------------------------
    'StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 3 AND bpronro =" & NroProcesoBatch
    '------MDF
    'StrSql = "SELECT iduser, bprctipomodelo, bpronro, bprcparam FROM batch_proceso WHERE btprcnro = 3 AND bpronro =" & NroProcesoBatch
    'OpenRecordset StrSql, rs_Batch_Proceso
    '------MDF
    HuboError = False
    
    If Not rs_Batch_Proceso.EOF Then
        'Call Establecer_Empresa(rs_batch_proceso!empnro)
        ContadorProgreso = 0
        usuario = rs_Batch_Proceso!iduser
        'Setea variable global que indica si el liquidador es Arg, Chile, Uru, etc.
        TipoProceso = IIf(EsNulo(rs_Batch_Proceso!bprctipomodelo), 0, rs_Batch_Proceso!bprctipomodelo)
        Flog.writeline Espacios(Tabulador * 0) & "Tipo de Proceso Liq (Arg - Chile -etc): " & TipoProceso
        
        'Flog.writeline Espacios(Tabulador * 0) & "Procedimiento EstablecerFirmas"
        Call EstablecerFirmas
        
        Flog.writeline Espacios(Tabulador * 0) & "Procedimiento Batliq06( " & rs_Batch_Proceso!Bpronro & "," & rs_Batch_Proceso!bprcparam & ")"
        Call batliq06(rs_Batch_Proceso!Bpronro, rs_Batch_Proceso!bprcparam)
    End If
    If rs_Batch_Proceso.State = adStateOpen Then rs_Batch_Proceso.Close
    Set rs_Batch_Proceso = Nothing
        
    If Not HuboError Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcprogreso = 100, bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProcesoBatch
    End If
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    On Error GoTo 0
    ' -----------------------------------------------------------------------------------
    'FGZ - 30/09/2004
    'Copio el proceso en el historico de batch_proceso y lo borro de batch_proceso
    If Not HuboError Then
        Call PasaraHistorico
    End If
    
Fin:
    'FGZ - 30/07/2013 -----------------------------
    'las estadisticas finales se muestran siempre
    'If CBool(USA_DEBUG) Then
        Call Estadisticas
    'Else
    '    Flog.writeline Espacios(Tabulador * 1) & "Cantidad de Lecturas en BD: " & Cantidad_de_OpenRecordset
    'End If
    Flog.Close

    If objConn.State = adStateOpen Then objConn.Close
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
    MyBeginTransLiq
        StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Error General', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
    MyCommitTransLiq
    GoTo Fin:
    
End Sub


Public Sub batliq06(ByVal Bpronro As Long, ByVal Parametros As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Programa que se ejecuta con el run value en batpro02.p con el cron para Liquidar
'              Configurado en el tipo de proceso batch 3
' Autor      : Maximiliano Breglia
' Fecha      : 01/12/01
' Traduccion : FGZ
' Fecha      : 17/11/2003
' --------------------------------------------------------------------------------------------
Dim p_grunro As Long

Dim Mantener_Liq As Boolean
Dim EmpleadoLiquidado As Boolean
'Dim Guardar_Nov As Boolean
Dim Analisis_Detallado As Boolean

Dim Todos As Boolean

Dim pos1 As Integer
Dim pos2 As Integer
Dim arrParam


Dim rs_batch_empleado As New ADODB.Recordset
Dim rs_Empleados As New ADODB.Recordset
Dim rs_cabliq As New ADODB.Recordset
Dim rs_Detliq As New ADODB.Recordset
Dim rs_HayAlgo As New ADODB.Recordset

    TiempoInicialProceso = GetTickCount

'    On Error GoTo CE
    
' El formato del mismo es (pronro.mantener Liq Ant.Guardar Nov.Analisis Det.Todos)
' Levanto cada parametro por separado, el separador de parametros es "."
Flog.writeline Espacios(Tabulador * 1) & "Levanto cada parametro por separado"
If Not EsNulo(Parametros) Then
    If Len(Parametros) >= 1 Then
'        pos1 = 1
'        pos2 = InStr(pos1, Parametros, ".") - 1
'        NroProc = CLng(Mid(Parametros, pos1, pos2))
'
'        pos1 = pos2 + 2
'        pos2 = InStr(pos1, Parametros, ".") - 1
'        Mantener_Liq = CBool(Mid(Parametros, pos1, pos2 - pos1 + 1))
'
'        pos1 = pos2 + 2
'        pos2 = InStr(pos1, Parametros, ".") - 1
'        guarda_nov = CBool(Mid(Parametros, pos1, pos2 - pos1 + 1))
'
'        pos1 = pos2 + 2
'        pos2 = InStr(pos1, Parametros, ".") - 1
'        HACE_TRAZA = CBool(Mid(Parametros, pos1, pos2 - pos1 + 1))
'
'        pos1 = pos2 + 2
'        pos2 = InStr(pos1, Parametros, ".") - 1
'        Todos = CInt(Mid(Parametros, pos1, pos2 - pos1 + 1))
'
'        pos1 = pos2 + 2
'        pos2 = InStr(pos1, Parametros, ".") - 1
'        NovedadesHist = CBool(Mid(Parametros, pos1, pos2 - pos1 + 1))
'
'        pos1 = pos2 + 2
'        pos2 = InStr(pos1, Parametros, ".") - 1
'        SoloLimpieza = CBool(Mid(Parametros, pos1, pos2 - pos1 + 1))
'
'        pos1 = pos2 + 2
'        pos2 = Len(Parametros)
'        USA_DEBUG = CBool(Mid(Parametros, pos1, pos2 - pos1 + 1))
        
        '24/09/2010 - Martin Ferraro - Se facilito la busqueda de la forma que se obtienen los parametros
        '                              y se agrego el parametro opcional borrarProc que indica si se borra
        '                              el proceso de liquidacion
        arrParam = Split(Parametros, ".")
        
        borrarProc = False
        NroProc = CLng(arrParam(0))
        Mantener_Liq = CBool(arrParam(1))
        guarda_nov = CBool(arrParam(2))
        HACE_TRAZA = CBool(arrParam(3))
        Todos = CInt(arrParam(4))
        NovedadesHist = CBool(arrParam(5))
        SoloLimpieza = CBool(arrParam(6))
        USA_DEBUG = CBool(arrParam(7))
        If UBound(arrParam) = 8 Then borrarProc = CBool(arrParam(8))
    
    End If
End If

'Martin Ferraro - 03/12/2008 - Seteo de la reutilizacion de traza
ReusaTraza = False
StrSql = "SELECT confnro, confactivo"
StrSql = StrSql & " FROM confper"
StrSql = StrSql & " WHERE confnro = 4"
OpenRecordset StrSql, rs_Empleados
If Not rs_Empleados.EOF Then
    ReusaTraza = CBool(rs_Empleados!confactivo)
End If
rs_Empleados.Close

Flog.writeline Espacios(Tabulador * 1) & "-----------------------------------------------------------------"
Flog.writeline Espacios(Tabulador * 1) & "Proceso                 : " & NroProc
Flog.writeline Espacios(Tabulador * 1) & "Mantiene la Liq anterior: " & Mantener_Liq
Flog.writeline Espacios(Tabulador * 1) & "Guarda las Novedades    : " & guarda_nov
Flog.writeline Espacios(Tabulador * 1) & "Genera Traza            : " & HACE_TRAZA
Flog.writeline Espacios(Tabulador * 1) & "Todos                   : " & Todos
Flog.writeline Espacios(Tabulador * 1) & "Liquida con Nov Hist    : " & NovedadesHist
Flog.writeline Espacios(Tabulador * 1) & "Solo Limpieza           : " & SoloLimpieza
Flog.writeline Espacios(Tabulador * 1) & "Usa Debug               : " & USA_DEBUG
Flog.writeline Espacios(Tabulador * 1) & "Reutiliza de Traza      : " & ReusaTraza
Flog.writeline Espacios(Tabulador * 1) & "Borrar Proceso          : " & borrarProc
Flog.writeline Espacios(Tabulador * 1) & "-----------------------------------------------------------------"
' Cargo el buliq_proceso

'FGZ - 20/02/2015 --------------------
If Todos Then
    USA_DEBUG = False
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "Se desactiva el modo DEBUG porque se están procesando demasiados empleados."
    Flog.writeline Espacios(Tabulador * 1) & "Usa Debug               : " & USA_DEBUG
    Flog.writeline Espacios(Tabulador * 1) & "-----------------------------------------------------------------"
End If
'FGZ - 20/02/2015 --------------------

Call Establecer_Proceso(NroProc)

'' FGZ - 04/02/2004
'' Cargo las Novedades Globales
'Call CargarNovedadesGlobales

'TOPES AMPO GLOBALES PARA EL PROCESO DE LIQUIDACIÓN
Call SetearAmpo
    
Fecha_Inicio = C_Date("01/" & CStr(buliq_periodo!pliqmes) & "/" & CStr(buliq_periodo!pliqanio))
If buliq_periodo!pliqmes = 12 Then
    Fecha_Fin = C_Date("01/01/" & CStr(buliq_periodo!pliqanio + 1)) - 1
Else
    Fecha_Fin = C_Date("01/" & CStr(buliq_periodo!pliqmes + 1) & "/" & CStr(buliq_periodo!pliqanio)) - 1
End If


If Not Todos Then
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 1) & "No son todos los empleados"
    End If
    StrSql = "SELECT * FROM batch_empleado " & _
             " WHERE bpronro =" & Bpronro
Else
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 1) & "Son todos los empleados"
        Flog.writeline Espacios(Tabulador * 1) & "busco los cabliq"
    End If
    
    'FGZ - 05/06/2012 ---------------------
    'StrSql = "SELECT * FROM cabliq "
    StrSql = "SELECT empleado FROM cabliq " & _
             " WHERE pronro =" & NroProc
    OpenRecordset StrSql, rs_cabliq
    If rs_cabliq.EOF Then
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 1) & "No hay cabeceras de liq asociadas al proceso"
        End If
    End If
    Do While Not rs_cabliq.EOF
        'Reviso que no esté
        StrSql = "SELECT ternro FROM batch_empleado " & _
                 " WHERE bpronro =" & Bpronro & _
                 " AND ternro =" & rs_cabliq!Empleado
        OpenRecordset StrSql, rs_batch_empleado
        
        If Not rs_batch_empleado.EOF Then
            StrSql = "INSERT INTO batch_empleado (bpronro,ternro) VALUES (" & _
            Bpronro & _
            "," & rs_cabliq!Empleado & _
            ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        
        rs_cabliq.MoveNext
    Loop
    StrSql = "SELECT ternro FROM batch_empleado " & _
             " WHERE bpronro =" & Bpronro
End If
OpenRecordset StrSql, rs_Empleados
If rs_Empleados.EOF Then
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 1) & "No hay empleados en Batch_empleados asociados al proceso"
    End If
End If
' si el empleado no estaba en cabliq entonces lo inserto
Do While Not rs_Empleados.EOF And Not Todos
    'FGZ - 05/06/2012 ------------------------------
    'StrSql = "SELECT * FROM cabliq "
    StrSql = "SELECT MAX(cliqnro) maximo FROM cabliq " & _
             " WHERE pronro =" & NroProc & _
             " AND empleado =" & rs_Empleados!Ternro
    OpenRecordset StrSql, rs_cabliq
             
    If rs_cabliq.EOF Then
        StrSql = "INSERT INTO cabliq (pronro, empleado) VALUES (" & _
        NroProc & _
        "," & rs_Empleados!Ternro & _
        ")"
        objConn.Execute StrSql, , adExecuteNoRecords
    Else
        'FGZ - 26/06/2012 --------
        If EsNulo(rs_cabliq!Maximo) Then
            StrSql = "INSERT INTO cabliq (pronro, empleado) VALUES (" & _
            NroProc & _
            "," & rs_Empleados!Ternro & _
            ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
    End If
             
    rs_Empleados.MoveNext
Loop

If rs_Empleados.State = adStateOpen Then rs_Empleados.Close
If rs_cabliq.State = adStateOpen Then rs_cabliq.Close


' FGZ - 11/08/2004 - Cargo el arreglo de conceptos
Call CargarConceptos(buliq_proceso!tprocnro, Fecha_Inicio, Fecha_Fin)

'Cargo las cabezeras de liquidacion en el arreglo global
Call CargarCabecerasLiq(Todos, NroProc, Bpronro)
Empleado_Actual = 1

'FGZ - 20/02/2015 ------------------------
If Cantidad_de_Empleados > 5 Then
    USA_DEBUG = False
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "Se desactiva el modo DEBUG porque se están procesando demasiados empleados."
    Flog.writeline Espacios(Tabulador * 1) & "Usa Debug               : " & USA_DEBUG
    Flog.writeline Espacios(Tabulador * 1) & "-----------------------------------------------------------------"
End If
'FGZ - 20/02/2015 ------------------------

'FGZ - 06/08/2015 -------------------------------------
    StrSql = "UPDATE batch_proceso SET bprcempleados = " & Cantidad_de_Empleados & _
             " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
'FGZ - 06/08/2015 -------------------------------------

'Cargo todas las tablas comunes durante toda la liquidacion del proceso
Call Cargar_Acumuladores

Call CargarBusquedas

Call Cargar_FunFormulas

Call Cargar_For_Tpa

Call Cargar_Con_For_Tpa

Call Cargar_Cge_Segun

Call Cargar_Cft_Segun

Call Cargar_Con_Acum(buliq_proceso!tprocnro, Fecha_Inicio, Fecha_Fin)

'cargo nombre de wf
Call CargarNombresTablasTemporales

'Call CreateTempTable(TTempWF_tpa)
'FGZ - 30/05/2011 --------------------------
Call CargarTiposEstructuras
Call CreateTempTable(TTempWF_impproarg)
Call CreateTempTable(TTempWF_Retroactivo)

'EAM (6.67) - Cargo novedades Globales una sola vez antes de comenzara a Ciclar
Set rs_NovGral = cargarNovedades(3) 'Globales

'La tabla tipoliquidador debe tener una entrada     3   RHProLiq.exe    Chile   Liquidador para Chile   0
        'Chile (TipoProceso = 3)- Si es el primer cto. de tipo 3, entra a calcular la gratificacion
        
If (TipoProceso = 3) Then
  Call CreateTempTable(TTempWF_EscalaUTM)
End If

' No se de donde salió ni para que sirve
p_grunro = 1

' Maximiliano Breglia 24/11/2006
v_nroitera = 1
Usa_grossing = False

'FGZ - 09/04/2014 -----------------------
'revisar si usa distribucion por novedad
Usa_Nov_Dist = HayNovDist()
'FGZ - 09/04/2014 -----------------------

'seteo las variables de progreso
CEmpleadosAProc = Cantidad_de_Empleados

Progreso = 0
HayAcuNoNeg = False

    
Do While Empleado_Actual <= Cantidad_de_Empleados 'Not rs_Empleados.EOF
    
    'Comienzo la transaccion
    MyBeginTransLiq

    EmpleadoSinError = True
    Termino_gross = True
    
    Fecha_Inicio = buliq_proceso!profecini
    Fecha_Fin = buliq_proceso!profecfin
    Call Establecer_Empleado(Arr_EmpCab(Empleado_Actual).Empleado, p_grunro, Arr_EmpCab(Empleado_Actual).cliqnro, Fecha_Inicio, Fecha_Fin)

    'FGZ - 21/09/2004
    EmpleadoLiquidado = False
    
    Call Establecer_Empresa(Fecha_Inicio, Fecha_Fin)
    
    ' Chequear si esta liquidado y no se deba borrar
    If Mantener_Liq Then 'no se puede borrar
        'StrSql = "SELECT * FROM detliq WHERE cliqnro = " & buliq_cabliq!cliqnro
        StrSql = "SELECT cliqnro FROM detliq WHERE cliqnro = " & buliq_cabliq!cliqnro
        If rs_HayAlgo.State = adStateOpen Then rs_HayAlgo.Close
        OpenRecordset StrSql, rs_HayAlgo
        If Not rs_HayAlgo.EOF Then
            'Empleado liquidado, no se hace nada
            EmpleadoLiquidado = True
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 0) & "------------------------------------------------------------------"
                Flog.writeline Espacios(Tabulador * 0) & "Se mantiene la liquidacion del Empleado: " & buliq_empleado!Empleg
            End If
        Else
            'StrSql = "SELECT * FROM acu_liq WHERE cliqnro = " & buliq_cabliq!cliqnro
            StrSql = "SELECT cliqnro FROM acu_liq WHERE cliqnro = " & buliq_cabliq!cliqnro
            If rs_HayAlgo.State = adStateOpen Then rs_HayAlgo.Close
            OpenRecordset StrSql, rs_HayAlgo
            If Not rs_HayAlgo.EOF Then
                'Empleado liquidado, no se hace nada
                EmpleadoLiquidado = True
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 0) & "------------------------------------------------------------------"
                    Flog.writeline Espacios(Tabulador * 0) & "Se mantiene la liquidacion del Empleado: " & buliq_empleado!Empleg
                End If
            End If
        End If
    Else
        'Eliminar los datos de la liquidación
        
        If CBool(USA_DEBUG) Then
            Flog.writeline
            Flog.writeline Espacios(Tabulador * 0) & "================================================================="
            Flog.writeline Espacios(Tabulador * 0) & "EMPLEADO: " & buliq_empleado!Empleg
            Flog.writeline
            Flog.writeline Espacios(Tabulador * 0) & "Usa Grossing: " & Usa_grossing
            Flog.writeline Espacios(Tabulador * 0) & "Iteracion (grossing): " & v_nroitera
            Flog.writeline Espacios(Tabulador * 0) & "================================================================="
            Flog.writeline
            Flog.writeline Espacios(Tabulador * 0) & "Eliminar los datos de la liquidación. Liqpro04"
        End If
        
        Call Liqpro04(buliq_cabliq!cliqnro, buliq_cabliq!Empleado, buliq_cabliq!pronro, buliq_proceso!PliqNro, False)
        
        'Actualizo el estado del Proceso de Liquidacion -  Se sacó porque causaba deadlock
        'If HayAcuNoNeg Then
        '    StrSql = "UPDATE proceso SET proaprob = 0, proestdesc = 'Acumuladores Negativos No permitidos' WHERE pronro = " & CStr(buliq_proceso!pronro)
        'Else
        '    StrSql = "UPDATE proceso SET proaprob = 0, proestdesc = null WHERE pronro = " & CStr(buliq_proceso!pronro)
        'End If
        'objConn.Execute StrSql, , adExecuteNoRecords
        
        'Verifico si termino bien el liqpro04
        If Not EmpleadoSinError Then
            GoTo SgtEmp
        End If
        
    End If
    
    If Not SoloLimpieza And ((Mantener_Liq And Not EmpleadoLiquidado) Or Not Mantener_Liq) Then
        'LIQUIDADOR
        TiempoInicialEmpleado = GetTickCount
        
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 0) & "------------------------------------------------------------------"
            Flog.writeline Espacios(Tabulador * 0) & "Comienza Liquidacion del empleado . Liqpro06"
        End If
        
        Call Liqpro06(buliq_proceso!tprocnro, buliq_proceso!pronro, buliq_proceso!PliqNro, buliq_cabliq!cliqnro, False)
        
        If CBool(USA_DEBUG) Then
            TiempoFinalEmpleado = GetTickCount
            Flog.writeline Espacios(Tabulador * 0) & "Tiempo para el empleado: " & (TiempoFinalEmpleado - TiempoInicialEmpleado)
        End If
    End If
    
    'Borro los cabliq
    If SoloLimpieza Then
        StrSql = "DELETE FROM cabliq " & _
                 " WHERE pronro =" & NroProc & _
                 " AND empleado =" & Arr_EmpCab(Empleado_Actual).Empleado
                 '" AND empleado =" & Arr_EmpCab(Empleado_Actual).Ternro
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
    
    'FGZ - 20/05/2011 ----------------------------------------------------------------
    ' por razones de eficiencia se actualiza el progreso una vez por cada empleado
    'FGZ - 28/07/2011 -------------
    'FGZ - si hace solo limpieza el tiempoAcumulado es 0 ==> la diferencia dá negativo
    TiempoAcumulado = GetTickCount
    'FGZ - 28/07/2011 -------------
    
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
             "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    'FGZ - 20/05/2011 ----------------------------------------------------------------
    
SgtEmp:

'Cheque si hubo algun error con el empleado
    If EmpleadoSinError Then
        'Borro de batch_empleado
        If Not Todos Then
            'StrSql = "DELETE FROM batch_empleado WHERE bpronro = " & Bpronro & " And Ternro = " & Arr_EmpCab(Empleado_Actual).Ternro
            StrSql = "DELETE FROM batch_empleado WHERE bpronro = " & Bpronro & " And Ternro = " & Arr_EmpCab(Empleado_Actual).Empleado
        Else
            StrSql = "DELETE FROM batch_empleado WHERE bpronro = " & Bpronro & " And Ternro = " & Arr_EmpCab(Empleado_Actual).Empleado
        End If
        objConn.Execute StrSql, , adExecuteNoRecords
        
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 0) & "Fin de la transaccion"
        End If
        'Fin de la transaccion
        MyCommitTransLiq
    End If
    

    ' Maximiliano Breglia 24/11/2006
    ' Si usa grossing vuelve a procesar el mismo empleado hasta que de el error
    If Not CBool(Usa_grossing) Then
         Empleado_Actual = Empleado_Actual + 1
         v_nroitera = 1
         v_netofijo = 0
         v_nroConce = 0
         Termino_gross = True
       Else
         v_nroitera = v_nroitera + 1
    End If
Loop

'Actualizo el estado del Proceso de Liquidacion
If HayAcuNoNeg Then
    StrSql = "UPDATE proceso SET proaprob = 1, proestdesc = 'Acumuladores Negativos No permitidos' WHERE pronro = " & CStr(buliq_proceso!pronro)
Else
    StrSql = "UPDATE proceso SET proaprob = 1, proestdesc = null WHERE pronro = " & CStr(buliq_proceso!pronro)
End If
objConn.Execute StrSql, , adExecuteNoRecords

If borrarProc Then
    StrSql = "DELETE proceso WHERE pronro = " & CStr(buliq_proceso!pronro)
    objConn.Execute StrSql, , adExecuteNoRecords
    If CBool(USA_DEBUG) Then Flog.writeline "Se borro el proceso " & CStr(buliq_proceso!pronro)
End If

TiempoFinalProceso = GetTickCount

If CBool(USA_DEBUG) Then
    Flog.writeline
    Flog.writeline "================================================================="
    Flog.writeline "Cantidad de empleados procesados: " & CEmpleadosAProc
    Flog.writeline "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
End If

'Borro los WF
'If Nueva_Tpa Then
'Else
'    Call BorrarTempTable(TTempWF_tpa)
'End If
Call BorrarTempTable(TTempWF_impproarg)
If (TipoProceso = 3) Then
 Call BorrarTempTable(TTempWF_EscalaUTM)
End If
Call BorrarTempTable(TTempWF_Retroactivo)

If rs_Empleados.State = adStateOpen Then rs_Empleados.Close
If rs_cabliq.State = adStateOpen Then rs_cabliq.Close

Set rs_cabliq = Nothing
Set rs_Empleados = Nothing
End Sub



Public Sub Liqpro16(ByVal Nro_empleado As Long, ByVal Nro_Concepto As Long, ByVal Cabecera_liq As Long)
' --------------------------------------------------------------------------
' Descripcion:  Realiza el proceso de ajuste de retroactivo para la liquidacion que se deshace.
' Autor:        13/08/2003 JMH
' Traducción:   FGZ
' Fecha:        10/09/2003
' Ultima Mod:
' --------------------------------------------------------------------------
Dim Mes_Aux As Integer

' Registros
Dim rs_Hisretroactivo As New ADODB.Recordset
Dim rs_Acu_Mes As New ADODB.Recordset

'RECORRE LOS REGISTROS HISTORICOS DE LOS ACUMULADORES DEL CONCEPTO QUE REQUIERE SER AJUSTADO


'StrSql = "SELECT * FROM hisretroactivo "
StrSql = "SELECT amanio, ammes, acunro, dlimonto  FROM hisretroactivo " & _
         " WHERE cliqnro = " & Cabecera_liq & _
         " AND concnro = " & Nro_Concepto
OpenRecordset StrSql, rs_Hisretroactivo

'FGZ
'lo cambié porque no tengo cosas como acu_mes.ammonto[mes_aux]
Do While Not rs_Hisretroactivo.EOF
'    ' Mes por Mes
'    For Mes_Aux = 1 To 12
        StrSql = "SELECT acunro FROM acu_mes " & _
                 " WHERE ternro = " & Nro_empleado & _
                 " AND acunro = " & rs_Hisretroactivo!acuNro & _
                 " AND amanio = " & rs_Hisretroactivo!amanio & _
                 " AND ammes = " & rs_Hisretroactivo!ammes ' Mes_Aux
        OpenRecordset StrSql, rs_Acu_Mes
        
        If Not rs_Acu_Mes.EOF Then
            StrSql = "UPDATE acu_mes SET ammonto = ammonto - " & rs_Hisretroactivo!dlimonto & _
                     " WHERE ternro = " & Nro_empleado & _
                     " AND acunro = " & rs_Hisretroactivo!acuNro & _
                     " AND amanio = " & rs_Hisretroactivo!amanio & _
                     " AND ammes = " & rs_Hisretroactivo!ammes
            objConn.Execute StrSql, , adExecuteNoRecords
            
            'Que pasa con dlicant???????
            ' el Update debería ser:
            'strsql = "UPDATE acu_mes SET ammonto = ammonto - " & rs_Hisretroactivo!dlimonto & _
            '         " ,amcant = amcant - " & rs_Hisretroactivo!dlicant & _
            '         " WHERE ternro = " & Nro_empleado & _
            '         " AND acunro = " & rs_Hisretroactivo!acunro & _
            '         " AND amanio = " & rs_Hisretroactivo!amanio & _
            '         " AND ammes = " & Mes_Aux

        Else
            'Nunca deberia fallar, salvo que hayan borrado el acumulador
        End If
            
'    Next Mes_Aux

    rs_Hisretroactivo.MoveNext
Loop

'borro los hisretroactivos
StrSql = "DELETE FROM hisretroactivo "
StrSql = StrSql & " WHERE cliqnro = " & Cabecera_liq
StrSql = StrSql & " AND concnro = " & Nro_Concepto
objConn.Execute StrSql, , adExecuteNoRecords

End Sub


Public Sub Liqpro04(ByVal Cabecera As Long, ByVal Nro_Ter As Long, ByVal Nro_Proceso As Long, ByVal Nro_Periodo As Long, ByVal Excluye_Liq As Boolean)
' --------------------------------------------------------------------------
' Descripcion:  Desmarca los registros generados por el liquidador.
'               Usado para reliquidaciones y para excluir empleados
'               de un proceso de Liquidación.
' Autor:        14/06/1998  JMH
' Traducción:   FGZ
' Fecha:        08/09/2003
' Ultima Mod:   FGZ - 18/04/2006 Desmarcado de pagos/dtos y licencias
'               Martin Ferraro - Adecuaciones para impuesto unico
'               25/09/2008 - Martin Ferraro - Desmarcar movimientos
'               FGZ - 01/10/2012 - Desmarcar_Comisiones
'               FGZ - 23/05/2014 - Borrado de ficharet aun cuando no existan detalles de liquidacion
' --------------------------------------------------------------------------

Dim Contador As Integer
Dim Aux_Valor As Double
Dim ultimoProceso As Boolean

' Registros
Dim rs_Periodo As New ADODB.Recordset
Dim rs_Detliq As New ADODB.Recordset
Dim rs_cabliq As New ADODB.Recordset
Dim rs_Acu_Liq As New ADODB.Recordset
Dim rs_Acu_Mes As New ADODB.Recordset
Dim rs_ImpPro As New ADODB.Recordset
Dim rs_ImpMes As New ADODB.Recordset
Dim rs_Prestamos As New ADODB.Recordset
Dim rs_Emp_Lic As New ADODB.Recordset
Dim rs_Procesos As New ADODB.Recordset
Dim rs_impUnico As New ADODB.Recordset
Dim rs_perHist As New ADODB.Recordset
Dim rs_AcuImpUnico As New ADODB.Recordset
Dim rs_ImpMesArg As New ADODB.Recordset


Dim modeloRecalculo As Boolean
Dim periHist As Long
Dim anioHist As Long
Dim mesHist As Long

On Error GoTo CELP04:

Principal:

    '30/06/2009 - Martin Ferraro - Si no existe det no entro a liqpro04
    StrSql = "SELECT cliqnro FROM detliq WHERE cliqnro = " & Cabecera
    OpenRecordset StrSql, rs_Procesos
    If rs_Procesos.EOF And Usa_grossing = False Then
        If CBool(USA_DEBUG) Then Flog.writeline Espacios(Tabulador * 1) & "No hay datos de liquidacion a borrar del legajo " & Legajo
        
        'Por alguna razon que no hemos podido detectar en ocaciones quedan registros dfe ficharet sin cabeceras ==> hago el delete de todas formas
        StrSql = "DELETE FROM ficharet " & _
                 " WHERE pronro =" & Nro_Proceso & _
                 " AND empleado =" & Nro_Ter
        objConn.Execute StrSql, , adExecuteNoRecords
        
        Exit Sub
    End If
    
    
    'Controlo si existe algun otro proceso para el empleado en el periodo
    StrSql = "SELECT cabliq.cliqnro FROM cabliq "
    StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro "
    StrSql = StrSql & " INNER JOIN periodo ON proceso.pliqnro = periodo.pliqnro "
    StrSql = StrSql & " WHERE periodo.pliqnro = " & Nro_Periodo
    StrSql = StrSql & "   AND cabliq.empleado = " & Nro_Ter
    StrSql = StrSql & "   AND cabliq.pronro  <> " & Nro_Proceso
    OpenRecordset StrSql, rs_Procesos
    
    ultimoProceso = rs_Procesos.EOF
    
    rs_Procesos.Close
    
    'Comienzo a desliquidar
    'FGZ - 20/05/2011 -----------------------------------------------
    'select *
    'StrSql = "SELECT * FROM periodo WHERE pliqnro =" & Nro_Periodo
    StrSql = "SELECT pliqanio, pliqmes FROM periodo WHERE pliqnro =" & Nro_Periodo
    OpenRecordset StrSql, rs_Periodo

    If Not rs_Periodo.EOF Then
        'Marcar datos generados de la liquidacion para REUSAR
'        If CBool(USA_DEBUG) Then
'            Flog.writeline Espacios(Tabulador * 1) & "Busco las liquidaciones empleado " & Legajo
'        End If
        
        'FGZ - 20/05/2011 -----------------------------------------------
        'select *
        'StrSql = "SELECT * FROM detliq WHERE cliqnro =" & Cabecera
        StrSql = "SELECT dliretro,ajustado,concnro FROM detliq WHERE cliqnro =" & Cabecera
        OpenRecordset StrSql, rs_Detliq
        
        Do While Not rs_Detliq.EOF
            'Ajuste retroactivo, debe volver los valores de los meses anteriores
            If CBool(rs_Detliq!dliretro) And CBool(rs_Detliq!ajustado) Then
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 2) & "Ajuste retroactivo, debe volver los valores de los meses anteriores. Liqpro16"
                End If
                Call Liqpro16(Nro_Ter, rs_Detliq!ConcNro, Cabecera)
            End If
            
            rs_Detliq.MoveNext
        Loop
        
        'Elimino los detliq
        StrSql = "DELETE FROM detliq WHERE cliqnro = " & Cabecera
        objConn.Execute StrSql, , adExecuteNoRecords
        
        'Verifico que se hayan borrado
        StrSql = "SELECT cliqnro FROM detliq WHERE cliqnro = " & Cabecera
        OpenRecordset StrSql, rs_cabliq
        If Not rs_cabliq.EOF Then
            Flog.writeline Espacios(Tabulador * 1) & "No se borraron los detliq del empleado. Empleado " & Legajo & ". Se aborta la liquidacion."
            'Exit Sub
            MyRollbackTransliq
            HuboError = True
            EmpleadoSinError = False
            Exit Sub
        Else
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 2) & "Elimino todo los detliq del empleado. Empleado " & Legajo
            End If
        End If
        
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 1) & "Acumuladores "
        End If
                
        'FGZ - 20/05/2011 ---------------------------
        'select *
        'StrSql = "SELECT * FROM acu_liq "
        StrSql = "SELECT acumulador.acumes, acumulador.acunro, acumulador.acuimponible,acumulador.acuimpcont " & _
                 " ,acu_liq.almonto,acu_liq.alcant,acu_liq.almontoreal " & _
                 " FROM acu_liq " & _
                 " INNER JOIN acumulador ON acu_liq.acunro = acumulador.acunro " & _
                 " WHERE cliqnro =" & Cabecera
        OpenRecordset StrSql, rs_Acu_Liq
        Do While Not rs_Acu_Liq.EOF
            If CBool(rs_Acu_Liq!acumes) Then
                'RESTAR AL ACUMULADOR MENSUAL
                'FGZ - 20/05/2011 ---------------------------
                'select *
                'StrSql = "SELECT * FROM acu_mes "
                StrSql = "SELECT acunro FROM acu_mes " & _
                         " WHERE amanio =" & rs_Periodo!pliqanio & _
                         " AND ternro =" & Nro_Ter & _
                         " AND ammes =" & rs_Periodo!pliqmes & _
                         " AND acunro =" & rs_Acu_Liq!acuNro
                OpenRecordset StrSql, rs_Acu_Mes
                
                If Not rs_Acu_Mes.EOF Then
                    If ultimoProceso Then
'                        StrSql = "UPDATE acu_mes SET ammontoreal = 0, ammonto = 0 " & _
'                                 " ,amcant = amcant - " & IIf(Not EsNulo(rs_Acu_Liq!alcant), rs_Acu_Liq!alcant, 0) & _
'                                 " WHERE amanio =" & rs_Periodo!pliqanio & _
'                                 " AND ternro =" & Nro_Ter & _
'                                 " AND ammes =" & rs_Periodo!pliqmes & _
'                                 " AND acunro =" & rs_Acu_Liq!acunro
                        StrSql = "UPDATE acu_mes SET ammontoreal = 0, ammonto = 0 " & _
                                 " ,amcant = 0 " & _
                                 " WHERE amanio =" & rs_Periodo!pliqanio & _
                                 " AND ternro =" & Nro_Ter & _
                                 " AND ammes =" & rs_Periodo!pliqmes & _
                                 " AND acunro =" & rs_Acu_Liq!acuNro
                    Else
                        StrSql = "UPDATE acu_mes SET ammontoreal = ammontoreal - " & rs_Acu_Liq!almontoreal & ", ammonto = ammonto - " & rs_Acu_Liq!almonto & _
                                 " ,amcant = amcant - " & IIf(Not EsNulo(rs_Acu_Liq!alcant), rs_Acu_Liq!alcant, 0) & _
                                 " WHERE amanio =" & rs_Periodo!pliqanio & _
                                 " AND ternro =" & Nro_Ter & _
                                 " AND ammes =" & rs_Periodo!pliqmes & _
                                 " AND acunro =" & rs_Acu_Liq!acuNro
                    End If
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
                
            End If 'Es Acumulador Mensual
            
            'ES IMPONIBLE
            If CBool(rs_Acu_Liq!acuimponible) Or CBool(rs_Acu_Liq!acuimpcont) Then
                'RESTAR DE LAS ESTRUCTURAS DE IMPONIBLE
                'FGZ - 20/05/2011 ------------------------------
                'StrSql = "SELECT * FROM impmesarg "
                'StrSql = "SELECT * FROM impproarg "
                StrSql = "SELECT tconnro, ipamonto,ipacant FROM impproarg " & _
                         " WHERE cliqnro =" & Cabecera & _
                         " AND acunro =" & rs_Acu_Liq!acuNro
                OpenRecordset StrSql, rs_ImpPro
            
                Do While Not rs_ImpPro.EOF  'por cada imponible
                    'FGZ - 20/05/2011 ------------------------------
                    'select *
                    'StrSql = "SELECT * FROM impmesarg "
                    StrSql = "SELECT acunro FROM impmesarg " & _
                             " WHERE imaanio =" & rs_Periodo!pliqanio & _
                             " AND ternro =" & Nro_Ter & _
                             " AND imames =" & rs_Periodo!pliqmes & _
                             " AND tconnro =" & rs_ImpPro!tconnro & _
                             " AND acunro =" & rs_Acu_Liq!acuNro
                    OpenRecordset StrSql, rs_ImpMes
                    
                    If Not rs_ImpMes.EOF Then
                        'Actualizo
                        If ultimoProceso Then
                            StrSql = "UPDATE impmesarg SET imamonto = 0 " & _
                                     " ,imacant = 0 " & _
                                 " WHERE imaanio =" & rs_Periodo!pliqanio & _
                                 " AND ternro =" & Nro_Ter & _
                                 " AND imames =" & rs_Periodo!pliqmes & _
                                 " AND tconnro =" & rs_ImpPro!tconnro & _
                                 " AND acunro =" & rs_Acu_Liq!acuNro
                            objConn.Execute StrSql, , adExecuteNoRecords
                            
                        Else
                            StrSql = "UPDATE impmesarg SET imamonto = imamonto - " & rs_ImpPro!ipamonto & _
                                     " ,imacant = imacant - " & rs_ImpPro!ipacant & _
                                 " WHERE imaanio =" & rs_Periodo!pliqanio & _
                                 " AND ternro =" & Nro_Ter & _
                                 " AND imames =" & rs_Periodo!pliqmes & _
                                 " AND tconnro =" & rs_ImpPro!tconnro & _
                                 " AND acunro =" & rs_Acu_Liq!acuNro
                            objConn.Execute StrSql, , adExecuteNoRecords
                            
                        End If
                        
                        StrSql = "UPDATE impproarg SET ipamonto = 0 , ipacant = 0 " & _
                             " WHERE cliqnro =" & Cabecera & _
                             " AND tconnro =" & rs_ImpPro!tconnro & _
                             " AND acunro =" & rs_Acu_Liq!acuNro
                        objConn.Execute StrSql, , adExecuteNoRecords
                        
                        If Excluye_Liq Then
                            'No es una reliquidacion, sino que se excluye de la misma.
                            'En la Reliq. 99% de probabilidad que se vuelva a generar, sino queda en cero.
                            StrSql = "DELETE FROM impmesarg " & _
                                 " WHERE imaanio =" & rs_Periodo!pliqanio & _
                                 " AND ternro =" & Nro_Ter & _
                                 " AND imames =" & rs_Periodo!pliqmes & _
                                 " AND tconnro =" & rs_ImpPro!tconnro & _
                                 " AND acunro =" & rs_Acu_Liq!acuNro
                            objConn.Execute StrSql, , adExecuteNoRecords
                        End If
                    End If
                    rs_ImpPro.MoveNext
                Loop 'por cada imponible
            End If 'ES IMPONIBLE
            
            rs_Acu_Liq.MoveNext
        Loop
        
        'Borro el acu_liq
        StrSql = "DELETE FROM acu_liq " & _
                 " WHERE cliqnro =" & Cabecera
        objConn.Execute StrSql, , adExecuteNoRecords
        
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 2) & "Borro el acu_liq"
        End If
        
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 1) & "BORRA LOS DATOS GENERADOS DE GANANCIAS"
        End If
        'BORRA LOS DATOS GENERADOS DE GANANCIAS
        StrSql = "DELETE FROM fichaded " & _
                 " WHERE pronro =" & Nro_Proceso & _
                 " AND empleado =" & Nro_Ter
        objConn.Execute StrSql, , adExecuteNoRecords
        
                
        StrSql = "DELETE FROM desliq " & _
                 " WHERE pronro =" & Nro_Proceso & _
                 " AND empleado =" & Nro_Ter
        objConn.Execute StrSql, , adExecuteNoRecords
                
                
        'StrSql = "DELETE FROM ficharem " & _
        '         " WHERE pronro =" & Nro_Proceso & _
        '         " AND empleado =" & Nro_Ter
        'objConn.Execute StrSql, , adExecuteNoRecords
        
        
        StrSql = "DELETE FROM ficharet " & _
                 " WHERE pronro =" & Nro_Proceso & _
                 " AND empleado =" & Nro_Ter
        objConn.Execute StrSql, , adExecuteNoRecords
        
         'StrSql = "SELECT * FROM ficharet " & _
         '        " WHERE pronro =" & Nro_Proceso & _
         '        " AND empleado =" & Nro_Ter
         ' OpenRecordset StrSql, rs_Prestamos
        
        'If CBool(USA_DEBUG) Then
        '    Flog.writeline Espacios(Tabulador * 1) & " NO Existe ficharet: " & rs_Prestamos.EOF
        'End If
        
        'FGZ - 26/05/2005
        'limpio traza_gan_item_top de ganancias
        StrSql = "DELETE FROM traza_gan_item_top "
        StrSql = StrSql & " WHERE ternro =" & Nro_Ter
        StrSql = StrSql & " AND pronro =" & Nro_Proceso
        objConn.Execute StrSql, , adExecuteNoRecords
        
        'FGZ - 30/12/2004
        'Limpio la traza de ganancias
        StrSql = "DELETE FROM traza_gan WHERE "
        StrSql = StrSql & " pliqnro =" & Nro_Periodo
        StrSql = StrSql & " AND pronro =" & Nro_Proceso
        StrSql = StrSql & " AND ternro =" & Nro_Ter
        objConn.Execute StrSql, , adExecuteNoRecords
        
        'Martin - 04/08/2010
        Call Desmarcar_Prestamos
        
'        'BORRA LAS CUOTAS DE PRESTAMOS GENERADOS POR EL PROCESO
'        If CBool(USA_DEBUG) Then
'            Flog.writeline Espacios(Tabulador * 1) & "BORRA LAS CUOTAS DE PRESTAMOS GENERADOS POR EL PROCESO"
'        End If
'
'        'Busco las cuotas que fueron generadas por el proceso
'        StrSql = "SELECT cuonro, cuototal, cuosaldo, cuogenera , pronrogenera, cuoordenliq "
'        StrSql = StrSql & " FROM pre_cuota"
'        StrSql = StrSql & " INNER JOIN prestamo ON prestamo.prenro = pre_cuota.prenro"
'        StrSql = StrSql & " WHERE pronrogenera = " & Nro_Proceso
'        StrSql = StrSql & " AND prestamo.ternro = " & Nro_Ter
'        StrSql = StrSql & " ORDER BY cuonrocuo, cuoordenliq"
'        OpenRecordset StrSql, rs_Prestamos
'        Do While Not rs_Prestamos.EOF
'
'            'Actualizo la cuota que particiono
'            StrSql = "UPDATE pre_cuota"
'            'If TodoNada Then
'            '    StrSql = StrSql & " SET cuototal = " & rs_Prestamos!cuototal
'            'Else
'                StrSql = StrSql & " SET cuototal = cuototal + " & rs_Prestamos!cuototal
'            'End If
'            StrSql = StrSql & " , cuosaldo = " & rs_Prestamos!cuosaldo
'            StrSql = StrSql & " WHERE cuonro = " & rs_Prestamos!cuogenera
'            objConn.Execute StrSql, , adExecuteNoRecords
'
'            'Borro la cuota generada
'            StrSql = "DELETE pre_cuota WHERE cuonro = " & rs_Prestamos!cuonro
'            objConn.Execute StrSql, , adExecuteNoRecords
'
'            rs_Prestamos.MoveNext
'        Loop
'        rs_Prestamos.Close
'
'
'        StrSql = "SELECT * FROM prestamo " & _
'                 "INNER JOIN pre_cuota ON prestamo.prenro = pre_cuota.prenro" & _
'                 " WHERE prestamo.ternro = " & Nro_Ter & _
'                 " AND pre_cuota.pronro =" & Nro_Proceso & _
'                 " AND pre_cuota.cuocancela = -1"
'        OpenRecordset StrSql, rs_Prestamos
'
'        Do While Not rs_Prestamos.EOF
'            'Es una cuota cancelada, la marco como no cancelada
'            StrSql = "UPDATE pre_cuota SET pronro = null, cuocancela = 0 "
'            StrSql = StrSql & " WHERE pronro = " & Nro_Proceso
'            StrSql = StrSql & " AND prenro = " & rs_Prestamos!prenro
'            StrSql = StrSql & " AND cuonro = " & rs_Prestamos!cuonro
'            objConn.Execute StrSql, , adExecuteNoRecords
'
'
'            If rs_Prestamos!estnro = 6 Then
'                ' Lo apruebo para re-liq la cuota
'                StrSql = "UPDATE prestamo SET estnro = 3 " & _
'                         " WHERE ternro = " & rs_Prestamos!ternro & _
'                         " AND prenro =" & rs_Prestamos!prenro
'                 objConn.Execute StrSql, , adExecuteNoRecords
'            End If
'
'            rs_Prestamos.MoveNext
'        Loop

        'BORRA LAS CUOTAS DE EMBARGO GENERADAS POR EL PROCESO
        Call Desmarcar_Embargos(Nro_Ter, Nro_Proceso)
        
        'FGZ - 26/05/2005
        'Desmarco los baes generados
        Call Desmarcar_BAE
        
        'Desmarco LOS VALES PAGADOS POR EL PROCESO
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 1) & "Desmarco LOS VALES PAGADOS POR EL PROCESO"
        End If
        
        'Martin Ferraro - 07/07/2006 - Se agrego " AND vales.empleado = " & Nro_Ter
        'Porque solo marcaba como liq el ultimo vale
        
        'FGZ - 29/05/2012 ---------------------------------------
        'StrSql = "UPDATE vales SET pronro = null " & _
        '         " WHERE pronro = " & NroProc & _
        '         " AND vales.empleado = " & Nro_Ter
        'objConn.Execute StrSql, , adExecuteNoRecords

        Call Desmarcar_Vales
        'FGZ - 29/05/2012 ---------------------------------------
        
'FGZ - 18/04/2006 - COMENTE LAS LINEAS DE ABAJO
'        'Desmarco LOS DIAS DE LAS LICENCIAS PAGADAS POR EL PROCESO
'        StrSql = "SELECT * FROM emp_lic " & _
'                 " WHERE empleado = " & Nro_Ter & _
'                 " AND tdnro = 2"
'        OpenRecordset StrSql, rs_Emp_Lic
'
'        Do While Not rs_Emp_Lic.EOF
'
'            StrSql = "UPDATE vacpagdesc SET pronro = null" & _
'                     " WHERE pronro =" & Nro_Proceso & _
'                     " AND ternro = " & Nro_Ter
'            objConn.Execute StrSql, , adExecuteNoRecords
'
'            rs_Emp_Lic.MoveNext
'        Loop
'
'        'FGZ - 14/01/2004
'        'Desmarco las licencias Marcadas por el Proceso
'        StrSql = "UPDATE emp_lic SET pronro = null " & _
'                 " WHERE pronro = " & Nro_Proceso
'        objConn.Execute StrSql, , adExecuteNoRecords
'        'Fin FGZ - 14/01/2004
        
        
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 1) & "Desmarco LOS PAGOS/DTOS MARCADOS POR EL PROCESO"
        End If
'FGZ - 18/04/2006 - REEMPLACE LAS LINEAS DE ARRIBA POR LO QUE SIGUE ABAJO
        'FGZ - 18/04/2006
        'Ojo que si no hubo licencias marcados por el proceso ==>
        'quedan afuera los pagos/dtos generados a partir de los dias correspondientes de vacaciones
        'y ... todos los pagos/dtos hechos por otro tipo de licencias que no sean de vacaciones tampoco se desmarcaban
        'en realidad el desmarcado no deberia estar dentro del loop
        
        'FGZ - 29/05/2012 ---------------------------------------
        'StrSql = "UPDATE vacpagdesc SET pronro = null "
        'StrSql = StrSql & " WHERE pronro = " & Nro_Proceso
        'StrSql = StrSql & " AND ternro = " & Nro_Ter
        'objConn.Execute StrSql, , adExecuteNoRecords
        Call Desmarcar_PagosDtos
        'FGZ - 29/05/2012 ---------------------------------------
        
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 1) & "Desmarco LAS LICENCIAS MARCADAS POR EL PROCESO"
        End If
        
        'FGZ - 29/05/2012 ---------------------------------------
        'Desmarco las licencias Marcadas por el Proceso
        'StrSql = "UPDATE emp_lic SET pronro = null "
        'StrSql = StrSql & " WHERE pronro = " & Nro_Proceso
        ''30/06/2009 - Martin Ferraro - Desmarcaba todas y no las del empleado
        'StrSql = StrSql & " AND empleado = " & Nro_Ter
        'objConn.Execute StrSql, , adExecuteNoRecords
'FGZ - 18/04/2006
        Call Desmarcar_Licencias
        'FGZ - 29/05/2012 ---------------------------------------
        
        'FGZ - 17/09/2012 --------------
        Call Desmarcar_Insalubridad
        
        'FGZ - 01/10/2012 --------------
        Call Desmarcar_Comisiones
        
        'FGZ 14/01/2013 ----------------
        Call Desmarcar_Gastos
        'FGZ 14/01/2013 ----------------
        
        'FGZ 28/02/2013 ----------------
        Call Desmarcar_DiasVacVendidos
        'FGZ 28/02/2013 ----------------
        
        'FGZ - 22/03/2005
        ' si se liquida guardando novedades historicas => entonces borrar sino no tocarlas
        If guarda_nov Then
            StrSql = "DELETE FROM hisnovemp " & _
                     " WHERE empleado = " & Nro_Ter & _
                     " AND pronro = " & Nro_Proceso
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        
        '04/12/2008 - Martin Ferraro - Borrar Novedades Generadas por Grossing
        If Not Usa_grossing Then
            StrSql = "DELETE FROM novemp " & _
                     " WHERE empleado = " & Nro_Ter & _
                     " AND pronro = " & Nro_Proceso
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        
        'FGZ - 10/02/2011 ----------------------------------------------
        'Desmarco las novedades retroactivas marcadas por el proceso
        StrSql = "UPDATE novretro SET pronropago = null "
        StrSql = StrSql & " WHERE pronropago = " & Nro_Proceso
        StrSql = StrSql & " AND ternro = " & Nro_Ter
        objConn.Execute StrSql, , adExecuteNoRecords
        'FGZ - 10/02/2011 ----------------------------------------------
        
        'Desmarcar Anticipo
        Call DesmarcarAnticipo

        'Si es proceso de recalculo
        modeloRecalculo = EsModeloRecalculo(Nro_Proceso)

        If modeloRecalculo Then
            'Borro impuni_cab
            StrSql = "DELETE impuni_cab"
            StrSql = StrSql & " WHERE cliqnro = " & Cabecera
            objConn.Execute StrSql, , adExecuteNoRecords
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 1) & "Es Modelo de Recalculo Chile - Borro impuni_cab"
            End If
        End If 'If modeloRecalculo Then
    
        Call Desmarcar_Mov

        'FGZ - 28/04/2011 -------------------
        'Desmarco los saldos de francos compensatorios generados por el proceso
         StrSql = "DELETE emp_fr_comp "
         StrSql = StrSql & " WHERE ternro = " & Nro_Ter
         StrSql = StrSql & " AND liq = -1"
         StrSql = StrSql & " AND pronro = " & Nro_Proceso
         objConn.Execute StrSql, , adExecuteNoRecords
        'FGZ - 28/04/2011 -------------------

        'EAM- Desmarca la venta de vacaciones
        Call Desmarcar_VentaVac(Nro_Proceso, Nro_Ter)

        Call Eliminar_Distribucion(Nro_Ter, Nro_Proceso)
            
        Call Desmarcar_Paros(Nro_Proceso, Nro_Ter)
            
        'Limpia la traza
        Call LimpiarTraza(Cabecera)
        
         If Excluye_Liq Then
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 1) & "Borro la cabecera de liquidación"
            End If
            StrSql = "DELETE FROM cabliq " & _
                     " WHERE cliqnro = " & Cabecera
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
            
        
            
    End If
    
Fin:
Exit Sub

'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
CELP04:
    'MyRollbackTransliq
    HuboError = True
    EmpleadoSinError = False
    'If CBool(USA_DEBUG) Then
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
        Flog.writeline Espacios(Tabulador * 0) & " Empleado abortado: " & buliq_empleado!Empleg
        Flog.writeline Espacios(Tabulador * 0) & " Error: " & Err.Description
        'Fue error de Consulta de SQL
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "SQL Ejecutado: " & StrSql
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
        Flog.writeline
    'End If
    MyRollbackTransliq
    GoTo Fin:
    
End Sub



Public Sub Liqpro06(Nrotipo As Long, NroProc As Long, NroPer As Long, NroCab As Long, Traza As Boolean)
' --------------------------------------------------------------------------
' Descripcion:  Liquidar un empleado
' Autor:        FGZ
' Fecha:
' Ultima Mod:
' --------------------------------------------------------------------------

' Definicion de Variable Locales
Dim OK As Boolean
Dim tipo As Long
Dim Resultado As Double
Dim fecha_retro As Date
Dim Dias_Trab As Integer
Dim tpa_parametro As Integer
Dim Calculo_Topes As Boolean
Dim Contador  As Integer
Dim Acumulado_Mes As Double
Dim Acumulado_Mes_Cant As Double
Dim tope_aplicar As Double
Dim tope_mes As Double

Dim Calculo_Gratif As Boolean
Dim modeloRecalculo As Boolean

Dim Tope_Monto_Proporcional As Double
Dim Tope_Monto_Semestral As Double

'FGZ - 06/01/2005
Dim Tope_Monto_Total As Double

Dim Cantidad_Proporcional As Double
Dim Alcance As Boolean

Dim Ampo_A_Aplicar As Double

Dim param1 As Integer
Dim param0 As Integer
Dim nro_tg As Integer

Dim YA_Ajustado As Boolean
Dim Ajustes As Double

' Fin de contrato
Dim Proporcion_Contrato As Double
Dim Resu_Proporcionado As Double
Dim A_Proporcionar As Boolean
Dim Fecha_Baja As Date
Dim Contrato_proporcion As Long
'Dim cant As Integer

Dim UltimoAcumulador As Long
Dim AcumuladoDeMonto As Double
Dim AcumuladoDeCant As Double


'Semestre del mes liquidando
Dim MesSemestreDesde As Integer
Dim MesSemestreHasta As Integer

' Auxiliares
Dim Aux_AcuNro As Long
Dim Aux_TconNro As Long
Dim HayAjuste As Boolean

Dim Actualiza As Boolean

Dim montoTope1 As Double
Dim montoTope2 As Double
Dim Desborde As Double

' Registros
Dim rs_Formulas As New ADODB.Recordset
Dim rs_Conceptos As New ADODB.Recordset
Dim rs_His_Estructura As New ADODB.Recordset
Dim rs_Alcance As New ADODB.Recordset
Dim rs_For_Tpa As New ADODB.Recordset
Dim rs_Cft_Segun As New ADODB.Recordset
Dim rs_Con_For_Tpa As New ADODB.Recordset
Dim rs_CftSegun As New ADODB.Recordset
Dim rs_CftSegun2 As New ADODB.Recordset
Dim rs_Programa As New ADODB.Recordset
Dim rs_Fases As New ADODB.Recordset
Dim rs_WF_impproarg As New ADODB.Recordset
Dim rs_ImpMesArg As New ADODB.Recordset
Dim rs_Ampo As New ADODB.Recordset
Dim rs_ImpPro As New ADODB.Recordset
Dim rs_AmpoConTpa As New ADODB.Recordset
Dim rs_ATC_acu As New ADODB.Recordset
Dim rs_acumulador As New ADODB.Recordset
Dim rs_Con_Acum As New ADODB.Recordset
Dim rs_AcuLiq As New ADODB.Recordset
Dim rs_tc_acu As New ADODB.Recordset
Dim rs_acumes As New ADODB.Recordset
Dim rs_ConceptoAux As New ADODB.Recordset
Dim rs_NovAjuste As New ADODB.Recordset
Dim rs_Contrato As New ADODB.Recordset
Dim rs_cystipo As New ADODB.Recordset
Dim rs_firmas As New ADODB.Recordset
Dim rs_tipoproc As New ADODB.Recordset

Dim Firmado As Boolean
Dim ChequeaFirmas As Boolean

Dim tipoNovedad As String   'Individual
                            'Estructura
                            'Global
                            
Dim Aux_Cantidad As Double

Dim Tope_Monto_Mensual As Double

'FGZ - 06/02/2004
Dim Acum
'Dim Aux_Acu_Monto as single
'Dim Aux_Acu_Cant As Single
'Dim Aux_Acu_MontoReal As Single
'Dim Aux_Nuevo_Monto As Single
'Dim Aux_Nuevo_Cant As Single

'FGZ - 19/07/2005
'pase de single a double porque
'cuando el valor numerico es de + de 5 digitos enteros me roba decimales (redoondea)
Dim Aux_Acu_Monto As Double
Dim Aux_Acu_Cant As Double
Dim Aux_Acu_MontoReal As Double
Dim Aux_Nuevo_Monto As Double
Dim Aux_Nuevo_Cant As Double

' FGZ - 21/05/2004
Dim Grupo As Long

'FGZ - 19/08/2004
Dim Selecc As String
Dim Aux_Acu_Actual As Long

'FGZ - 04/10/2004
Dim rs_Aux_WF_impproarg As New ADODB.Recordset
Dim Aux_impproarg1 As Double
Dim Aux_impproarg2 As Double

'FGZ -02/08/2005
'agregué un chequeo cuando inseta en los imponibles del proceso
Dim NoActualiza As Boolean

'FGZ -16/08/2005
'agregué que genere el acumulador de desborde cuando hay desborde
Dim Monto_Desborde As Double

Dim MesFase As Integer
Dim AnioFase As Integer
Dim MesSemestreDesdeAux As Integer
Dim AnioSemestreDesdeAux As Integer
Dim IndiceCon As Long
Dim I As Long

Dim rs_novEmp As New ADODB.Recordset    'EAM- (6.67) - Cargo todas las novedades del empleado
Dim rs_NovEstr As New ADODB.Recordset   'EAM- (6.67) - Cargo todas las novedades pro estructura

    ' Inicio codigo ejecutable
    On Error GoTo CE
    
    OK = False
    Alcance = False
    Actualizo_Bae = False
    
    Calculo_Gratif = False
    modeloRecalculo = False
    
    'Limpio el cache de acumuladores
    Call objCache_Acu_Liq_Monto.Limpiar
    Call objCache_Acu_Liq_MontoReal.Limpiar
    Call objCache_Acu_Liq_Cantidad.Limpiar
    
    ' Limpiar los wf de topes
    Call LimpiarTempTable(TTempWF_impproarg)
    
    'FGZ - 27/05/2011 -------------------------
    Call CargarHis_Estructura
    Call CargarHis_EstructuraEmp
    Call CargarHis_EstructuraPer
    
    'EAM- Cargo las novedades del empleado
    Set rs_novEmp = cargarNovedades(1)  'Individuales
    Set rs_NovEstr = cargarNovedades(2) 'Estructura
    
    ' FGZ - 28/01/2004
    'Limpio los ampos
    Cant_Ampo_Proporcionar_1 = 0
    Cant_Ampo_Proporcionar_2 = 0
    Cant_Ampo_Proporcionar_3 = 0
    Cant_Ampo_Proporcionar_4 = 0
    Cant_Ampo_Proporcionar_5 = 0
    
    Sumo_Cant_Ampo_Prop_1 = False
    Sumo_Cant_Ampo_Prop_2 = False
    Sumo_Cant_Ampo_Prop_3 = False
    Sumo_Cant_Ampo_Prop_4 = False
    Sumo_Cant_Ampo_Prop_5 = False
    
    ' FGZ - 28/01/2004
    
    ' -----------------------------------------------------------------------
    ' CALCULAR LA PROPORCION DE CONTRATO: ACTUAL O ANTERIOR
    Proporcion_Contrato = 0
    Calculo_Topes = False
    Dias_Trab = 30
    
'    StrSql = "SELECT * FROM fases WHERE real = -1 AND empleado = " & buliq_empleado!ternro & _
'            " AND ((altfec >= " & ConvFecha(fecha_inicio) & " AND altfec <= " & ConvFecha(fecha_fin) & ") " & _
'            " OR (bajfec <= " & ConvFecha(fecha_fin) & "))" & _
'            " ORDER BY altfec"
    'FGZ - 20/05/2011 ---------------------------------------------------
    'StrSql = "SELECT * FROM fases WHERE real = -1 AND empleado = " & buliq_empleado!Ternro
    StrSql = "SELECT altfec,bajfec,estado FROM fases WHERE real = -1 AND empleado = " & buliq_empleado!Ternro & _
            " AND (altfec <= " & ConvFecha(Fecha_Fin) & " AND ( bajfec >= " & ConvFecha(Fecha_Inicio) & _
            " OR bajfec is null ))" & _
            " ORDER BY altfec"
    OpenRecordset StrSql, rs_Fases
    If Not rs_Fases.EOF Then rs_Fases.MoveLast
    If Not rs_Fases.EOF Then
        Dias_Trab = CantidadDeDias(Fecha_Inicio, Fecha_Fin, rs_Fases!altfec, IIf(EsNulo(rs_Fases!bajfec), Fecha_Fin, rs_Fases!bajfec))
        If Not CBool(rs_Fases!Estado) Then
            Fecha_Baja = rs_Fases!bajfec
        End If
    End If
    
    If Fecha_Inicio <= buliq_empleado!empfbajaprev And buliq_empleado!empfbajaprev <= Fecha_Fin And Fecha_Baja > buliq_empleado!empfbajaprev Or EsNulo(Fecha_Baja) Then
        ' SOLO PROPORCIONO cuando se dio de baja posterior a la fecha de vencimiento de contrato o cuando aún no se dio de baja
        ' verificar si el contrato es anterior o actual
        If Fecha_Inicio <= buliq_empleado!empfbajaprev And buliq_empleado!empfbajaprev <= Fecha_Fin Then
'            Proporcion_Contrato = buliq_empleado!empactprop / Dias_Trab
'            Contrato_proporcion = buliq_empleado!tcnro
'        Else
'            If fecha_inicio <= buliq_empleado!empantvto And buliq_empleado!empantvto <= fecha_inicio Then
'                Proporcion_Contrato = buliq_empleado!empantprop / Dias_Trab
'                Contrato_proporcion = buliq_empleado!tcantnro
'            End If

            'FGZ - 06/12/2004
            'Se dió esta condicion y los campos ( empantprop tcnro y tcantnro ya no existen en x2)
            'Por ahora lo sacamos
            Proporcion_Contrato = 1
            Contrato_proporcion = 0
        End If
    End If

    If (Proporcion_Contrato >= 1) Or (Proporcion_Contrato = 0) Then
        Proporcion_Contrato = 0
    Else
        Proporcion_Contrato = 1 - Proporcion_Contrato
    End If

    ' default por si no encontró un contrato de proporcion y debe proporcionar, el actual
    'Buscar los contratos actual = 18, anterior = 26 y futuro = 27  tipos de estructura fijos  esto para sacar  el Contrato_proporcion */
    If Contrato_proporcion = 0 Then
'        StrSql = "SELECT replica_estr.origen, replica_estr.estrnro " & _
'        " FROM his_estructura, replica_estr " & _
'        " WHERE tenro = 18 " & _
'        " AND his_estructura.ternro = " & buliq_empleado!ternro & _
'        " AND his_estructura.htetdesde <=" & ConvFecha(Fecha_Fin) & _
'        " AND (his_estructura.htethasta >= " & ConvFecha(Fecha_Inicio) & _
'        " OR his_estructura.htethasta IS NULL)" & _
'        " AND replica_estr.estrnro = his_estructura.estrnro "
'            OpenRecordset StrSql, rs_Contrato
'            If Not rs_Contrato.EOF Then
'                Contrato_proporcion = rs_Contrato!Origen
'            End If
            
'            'FGZ - 19/05/2011
'            IndiceCon = IndiceContrato(buliq_empleado!Ternro)
'            If IndiceCon > 0 Then
'                Contrato_proporcion = Arr_Contrato(IndiceCon).proporcion
'            End If
    End If

'Contrato_proporcion = buliq_empleado!tcantnro
'    If CBool(USA_DEBUG) Then
'        Flog.writeline Espacios(Tabulador * 1) & "Contrato a Proporcionar " & Contrato_proporcion
'        If Not rs_contrato.EOF Then
'            Flog.writeline Espacios(Tabulador * 2) & "Estructura " & rs_contrato!estrnro
'        End If
'        Flog.writeline Espacios(Tabulador * 2) & "Dias Trabajados " & Dias_Trab
'        Flog.writeline Espacios(Tabulador * 1) & "Proporción de Sueldo " & Proporcion_Contrato
'    End If
    
       
    ' FIN DE PROPORCION DEL CONTRATO
    ' -----------------------------------------------------------------------

    ' calculo del mes del semestre del periodo que se está liquidando
    If buliq_periodo!pliqmes >= 1 And buliq_periodo!pliqmes <= 6 Then
        MesSemestreDesde = 1
        MesSemestreHasta = Minimo(6, buliq_periodo!pliqmes)
    Else
        MesSemestreDesde = 7
        MesSemestreHasta = Minimo(12, buliq_periodo!pliqmes)
    End If


   
    'CONCEPTOS:
    ' FGZ - 11/08/2004
    ' Antes ejecutaba el sql y recorria el recorset (esto se ejecutaba por cada empleado)
    ' Ahora se cargan Una sola vez todos los conceptos en un vector y se recorre ese vector por cada empleado
    ' se cambió toda referencia a rs_concepto por Arr_Conceptos(Concepto_actual)
    
    'seteo de las variables de progreso
    CConceptosAProc = Cantidad_de_Conceptos 'rs_Conceptos.RecordCount
    'FGZ - 22/10/2004
    If CEmpleadosAProc = 0 Then
        CEmpleadosAProc = 1
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 1) & "No se encontraron empleados para procesar "
    End If
    If CConceptosAProc = 0 Then
        CConceptosAProc = 1
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 1) & "No se encontraron conceptos a procesar "
    End If
    IncPorc = ((100 / CEmpleadosAProc) * (100 / CConceptosAProc)) / 100
    IncPorcEmpleado = (100 / CConceptosAProc)

    Concepto_Actual = 1
    Do While Concepto_Actual <= Cantidad_de_Conceptos 'Not rs_Conceptos.EOF
   
        ' PRIMER CONCEPTO, POSTERIOR AL ULTIMO REMUNERATIVO,  calcular: impproarg y topear acu_liq
        If (Arr_conceptos(Concepto_Actual).tconnro > 3) And (Not Calculo_Topes) Then
            Calculo_Topes = True
            
            If CBool(USA_DEBUG) Then Flog.writeline Espacios(Tabulador * 1) & "Inicio Calculo de Mopre (tipo Concepto > 3) "
            
            'FGZ - 01/02/2008 - Actualizo tomos los imponibles negativos a 0
            StrSql = "UPDATE " & TTempWF_impproarg & " set ipamonto = 0, ipacant = 0 "
            StrSql = StrSql & " WHERE ipamonto < 0 "
            objConn.Execute StrSql, Reg_Afected, adExecuteNoRecords
            If CBool(USA_DEBUG) Then
                Flog.writeline
                Flog.writeline Espacios(Tabulador * 1) & "Actualizo tomos los imponibles negativos a 0."
                Flog.writeline Espacios(Tabulador * 1) & "  registros afectados: " & Reg_Afected
                Flog.writeline
            End If
            
            'FGZ - 04/06/2012 ----------------------------
            'StrSql = "SELECT * FROM " & TTempWF_impproarg & " ORDER BY acunro, tconnro"
            StrSql = "SELECT acunro , ipacant, ipamonto, tconnro, desborde, tope_aporte FROM " & TTempWF_impproarg & " ORDER BY acunro, tconnro"
            OpenRecordset StrSql, rs_WF_impproarg
            
            AcumuladoDeMonto = 0
            AcumuladoDeCant = 0
            UltimoAcumulador = -1
            Do While Not rs_WF_impproarg.EOF
            
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 1) & " Acumulador a Topear: " & rs_WF_impproarg!acuNro
                    Flog.writeline Espacios(Tabulador * 1) & " Tipo de Tope: " & rs_WF_impproarg!tconnro
                    Flog.writeline Espacios(Tabulador * 1) & " Monto a Topear : " & rs_WF_impproarg!ipamonto
                    Flog.writeline Espacios(Tabulador * 1) & " Cant a Topear : " & rs_WF_impproarg!ipacant
                End If
               
                If rs_WF_impproarg!acuNro <> UltimoAcumulador Then
                    AcumuladoDeMonto = 0
                    AcumuladoDeCant = 0
                    UltimoAcumulador = rs_WF_impproarg!acuNro
                End If
                            
                ' Buscar el Acumulado Mensual Imponible que me interesa
                'FGZ - 08/06/2012 --------------
                StrSql = "SELECT imamonto,imacant FROM impmesarg " & _
                         " WHERE ternro = " & buliq_empleado!Ternro & _
                         " AND acunro = " & rs_WF_impproarg!acuNro & _
                         " AND tconnro = " & rs_WF_impproarg!tconnro & _
                         " AND imaanio = " & buliq_periodo!pliqanio & _
                         " AND imames = " & buliq_periodo!pliqmes
                OpenRecordset StrSql, rs_ImpMesArg
                            
                    'Buscar los imponibles ya utilizados para ser restados:
                    'Sueldo(1) y Vacaciones(2): solo lo de este mes
                    'SAC(3): lo del semestre que estoy liquidando
                                        
                    If rs_WF_impproarg!tconnro = 1 Or rs_WF_impproarg!tconnro = 2 Then
                        ' solo tomar lo del mes como ya acumulado
                        If Not rs_ImpMesArg.EOF Then
                            Acumulado_Mes = rs_ImpMesArg!imamonto
                            Acumulado_Mes_Cant = rs_ImpMesArg!imacant
                        Else
                            Acumulado_Mes = 0
                            Acumulado_Mes_Cant = 0
                        End If
                    End If
                    
                    If rs_WF_impproarg!tconnro = 3 Then
                        ' Tomar el semestre como ya acumulado
                        Acumulado_Mes = 0
                        Acumulado_Mes_Cant = 0
                                                

                        '03/07/2009 - Martin - Busco la fase para comparar con la fecha desde de semestre
                        MesFase = 0
                        AnioFase = 0
                        MesSemestreDesdeAux = MesSemestreDesde
                        AnioSemestreDesdeAux = buliq_periodo!pliqanio
                        StrSql = "SELECT altfec FROM fases WHERE estado = -1  AND real = -1 AND empleado = " & buliq_empleado!Ternro
                        OpenRecordset StrSql, rs_Fases
                        If Not rs_Fases.EOF Then
                            If Not EsNulo(rs_Fases!altfec) Then
                                MesFase = Month(rs_Fases!altfec)
                                AnioFase = Year(rs_Fases!altfec)
                                'Comparacion por año
                                If CInt(buliq_periodo!pliqanio) < AnioFase Then
                                    'Debo tomar la fecha de la fase como del semestre
                                    MesSemestreDesdeAux = MesFase
                                    AnioSemestreDesdeAux = AnioFase
                                Else
                                    'Comparacion por mes
                                    If CInt(buliq_periodo!pliqanio) = AnioFase Then
                                        'Debo tomar la fecha de la fase como del semestre
                                        If MesSemestreDesde < MesFase Then
                                            MesSemestreDesdeAux = MesFase
                                            AnioSemestreDesdeAux = AnioFase
                                        End If
                                    End If
                                End If
                            End If
                        End If
                        
                        
                        StrSql = "SELECT sum(imamonto) AS SUMA ,sum(imacant)AS CANTIDAD FROM impmesarg " & _
                                 " WHERE ternro = " & buliq_empleado!Ternro & _
                                 " AND acunro = " & rs_WF_impproarg!acuNro & _
                                  " AND tconnro = " & rs_WF_impproarg!tconnro & _
                                 " AND imaanio = " & AnioSemestreDesdeAux & _
                                 " AND imames >= " & MesSemestreDesdeAux & _
                                 " AND imames <= " & MesSemestreHasta
                        OpenRecordset StrSql, rs_ImpMesArg
                        
                        If Not rs_ImpMesArg.EOF Then
                            Acumulado_Mes = Acumulado_Mes + IIf(EsNulo(rs_ImpMesArg!Suma), 0, rs_ImpMesArg!Suma)
                            Acumulado_Mes_Cant = Acumulado_Mes_Cant + IIf(EsNulo(rs_ImpMesArg!Suma), 0, rs_ImpMesArg!Cantidad)
                        End If
                    End If
                    ' calculo de ya acumulado de SAC
                    
                    
                    ' Verificar los topes grales. contra mes mas el proceso actual
                    ' si lo supera se guarda al resto hasta el tope, nunca puede ser negativo.
                    
                    ' ------------ CONTROL PARA LOS TOPES DEL AMPO --------------------
                    ' Calcular el tope de AMPO proporcionado de acurdo a la cantidad de dias que estoy liquidando
                    
                    If CBool(rs_WF_impproarg!tope_aporte) Then
                        Ampo_A_Aplicar = Valor_Ampo
                    Else
                        Ampo_A_Aplicar = Valor_Ampo_Cont
                    End If
                    
                    If CBool(USA_DEBUG) Then
                        Flog.writeline Espacios(Tabulador * 1) & " Acumulado Mes: " & Acumulado_Mes
                        Flog.writeline Espacios(Tabulador * 1) & " Acumulado Mes Cant: " & Acumulado_Mes_Cant
                    End If
                   
                   
                
                    Select Case rs_WF_impproarg!tconnro
                    Case 1:
                        'EAM(6.75) - Se modificó la forma de calcular el tope del AMPO
                        Tope_Monto_Mensual = Ampo_A_Aplicar * (Ampo_Max_1 * Cant_Diaria_Ampos_1)
                        If (Cant_Ampo_Proporcionar_1 = 0 And Acumulado_Mes_Cant = 0) Or Not CBool(Ampo_Proporciona_1) Then
                            Tope_Monto_Proporcional = Valor_Ampo * Ampo_Max_1
                        Else
                            'Tope_Monto_Proporcional = (Valor_Ampo * Cant_Diaria_Ampos_1) * Acumulado_Mes_Cant
                            If Acumulado_Mes_Cant Then
                                Cant_Ampo_Proporcionar_1 = Acumulado_Mes_Cant
                            End If
                            Tope_Monto_Proporcional = (Valor_Ampo * Cant_Diaria_Ampos_1) * Cant_Ampo_Proporcionar_1
                        End If
                        

                        If CBool(USA_DEBUG) Then
                            Flog.writeline Espacios(Tabulador * 1) & " Valor Ampo Aplicar: " & Ampo_A_Aplicar
                            Flog.writeline Espacios(Tabulador * 1) & " Ampo Max Sueldo: " & Ampo_Max_1
                            Flog.writeline Espacios(Tabulador * 1) & " Cant Ampo a Proporcionar: " & Cant_Ampo_Proporcionar_1
                            Flog.writeline Espacios(Tabulador * 1) & " Cant Diaria Ampo Sueldo: " & Cant_Diaria_Ampos_1
                            Flog.writeline Espacios(Tabulador * 1) & " Tope Proporcional Aplicar Sueldo: " & Tope_Monto_Proporcional
                        End If
                    Case 2:
                        Tope_Monto_Mensual = Ampo_A_Aplicar * (Ampo_Max_2 * Cant_Diaria_Ampos_2)
                        If (Cant_Ampo_Proporcionar_2 = 0 And Not Sumo_Cant_Ampo_Prop_2) Or Not CBool(Ampo_Proporciona_2) Then
                        'If Not CBool(Ampo_Proporciona_2) Then
                            If rs_WF_impproarg!tope_aporte Then
                                Tope_Monto_Proporcional = Valor_Ampo * Ampo_Max_2
                            Else
                                Tope_Monto_Proporcional = Valor_Ampo_Cont * Ampo_Max_2
                            End If
                        Else
                            Tope_Monto_Proporcional = (Cant_Ampo_Proporcionar_2 * Cant_Diaria_Ampos_2) * Ampo_A_Aplicar
                        End If
                        If CBool(USA_DEBUG) Then
                            Flog.writeline Espacios(Tabulador * 1) & " Valor Ampo Aplicar: " & Ampo_A_Aplicar
                            Flog.writeline Espacios(Tabulador * 1) & " Ampo Max Vac: " & Ampo_Max_2
                            Flog.writeline Espacios(Tabulador * 1) & " Cant Ampo a Proporcionar 2: " & Cant_Ampo_Proporcionar_2
                            Flog.writeline Espacios(Tabulador * 1) & " Cant Diaria Ampo Vac: " & Cant_Diaria_Ampos_2
                            Flog.writeline Espacios(Tabulador * 1) & " Tope Proporcional Aplicar VAC: " & Tope_Monto_Proporcional
                        End If
                    Case 3:
                        'FGZ - 25/06/2007 - Agregué esta variable para tener en cuenta a la hora de topear
                        Tope_Monto_Semestral = Ampo_A_Aplicar * (Ampo_Max_3 * Cant_Diaria_Ampos_3)
                        'FGZ - 25/06/2007 - Agregué esta variable para tener en cuenta a la hora de topear
                        If (Cant_Ampo_Proporcionar_3 = 0 And Not Sumo_Cant_Ampo_Prop_3) Or Not CBool(Ampo_Proporciona_3) Then
                        'If Not CBool(Ampo_Proporciona_3) Then
                            If rs_WF_impproarg!tope_aporte Then
                                Tope_Monto_Proporcional = Valor_Ampo * Ampo_Max_3 * Cant_Diaria_Ampos_3
                            Else
                                Tope_Monto_Proporcional = Valor_Ampo_Cont * Ampo_Max_3 * Cant_Diaria_Ampos_3
                            End If
                        Else
                            Tope_Monto_Proporcional = (Cant_Ampo_Proporcionar_3 * Cant_Diaria_Ampos_3) * Ampo_A_Aplicar
                        End If
                        If CBool(USA_DEBUG) Then
                            Flog.writeline Espacios(Tabulador * 1) & " Valor Ampo Aplicar: " & Ampo_A_Aplicar
                            Flog.writeline Espacios(Tabulador * 1) & " Ampo Max SAC: " & Ampo_Max_3
                            Flog.writeline Espacios(Tabulador * 1) & " Cant Ampo a Proporcionar SAC: " & Cant_Ampo_Proporcionar_3
                            Flog.writeline Espacios(Tabulador * 1) & " Cant Diaria Ampo SAC: " & Cant_Diaria_Ampos_3
                            Flog.writeline Espacios(Tabulador * 1) & " Tope Proporcional Aplicar SAC: " & Tope_Monto_Proporcional
                        End If
                    End Select
                    
                    ' ------------------- topea contra proporciones, ej. Sueldo, Vacaciones, SAC ----------------------
                    ' Guardo al acunro para restaurar despues de actualizar el WF
                    Aux_AcuNro = rs_WF_impproarg!acuNro
                    Aux_TconNro = rs_WF_impproarg!tconnro
                    
                    'FGZ - 13/08/2007 - ---- Se hizo un cambio si proporciona porque estaba tomando el acumulado mensual y no deberia
                    'Aux_Cantidad = IIf((rs_WF_impproarg!ipacant - Acumulado_Mes_Cant < 0), 0, rs_WF_impproarg!ipacant - Acumulado_Mes_Cant)
                    Aux_Cantidad = IIf((rs_WF_impproarg!ipacant < 0), 0, rs_WF_impproarg!ipacant)
                    Select Case rs_WF_impproarg!tconnro
                    Case 1:
                    
                        If CBool(USA_DEBUG) Then Flog.writeline Espacios(Tabulador * 1) & " Tope de Sueldo "
             
                        'FGZ - 03/08/2007 - ---- Se hizo un cambio si proporciona porque estaba tomando el acumulado mensual y no deberia
                        If Not CBool(Ampo_Proporciona_1) Then
                            If (Acumulado_Mes + rs_WF_impproarg!ipamonto) > Tope_Monto_Proporcional Then
                                StrSql = "UPDATE " & TTempWF_impproarg & " SET desborde = (" & rs_WF_impproarg!ipamonto & ") - " & Tope_Monto_Proporcional - Acumulado_Mes & _
                                         ", ipamonto = " & Tope_Monto_Proporcional - Acumulado_Mes & _
                                         ",ipacant = " & Aux_Cantidad & _
                                         "WHERE acunro = " & rs_WF_impproarg!acuNro & " AND tconnro = " & rs_WF_impproarg!tconnro
                                objConn.Execute StrSql, , adExecuteNoRecords
                                
                                If CBool(USA_DEBUG) Then Flog.writeline Espacios(Tabulador * 1) & " Si (AcumMes + Monto a Topear >  Tope Proporcional Aplicar) guardo: " & Tope_Monto_Proporcional - Acumulado_Mes
                                
                                If Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde <> 0 Then
                                    Monto_Desborde = (rs_WF_impproarg!ipamonto) - (Tope_Monto_Proporcional - Acumulado_Mes)
                                    If objCache_Acu_Liq_Monto.EsSimboloDefinido(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde)) Then
                                        'levanto el monto
                                        Monto_Desborde = Monto_Desborde + objCache_Acu_Liq_Monto.Valor(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde))
                                        'borro y vuelvo a insertar
                                        Call objCache_Acu_Liq_Monto.Borrar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde))
                                        Call objCache_Acu_Liq_MontoReal.Borrar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde))
                                        Call objCache_Acu_Liq_Cantidad.Borrar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde))
                                        
                                        Call objCache_Acu_Liq_Monto.Insertar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde), Monto_Desborde)
                                        Call objCache_Acu_Liq_MontoReal.Insertar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde), Monto_Desborde)
                                        Call objCache_Acu_Liq_Cantidad.Insertar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde), 0)
                                    Else
                                        Call objCache_Acu_Liq_Monto.Insertar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde), Monto_Desborde)
                                        Call objCache_Acu_Liq_MontoReal.Insertar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde), Monto_Desborde)
                                        Call objCache_Acu_Liq_Cantidad.Insertar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde), 0)
                                    End If
                                End If
                            End If
                        Else
                        ' Se agregó el tope mes para el control de 2 liquidacion en el periodo MB 08/08/2008
                        If (rs_WF_impproarg!ipamonto) > Tope_Monto_Proporcional Then
                              tope_aplicar = Tope_Monto_Proporcional
                            Else
                              tope_aplicar = rs_WF_impproarg!ipamonto
                            End If
                            
                            'EAM(6.75) - Se cambio como se calcula el tope de mes
                            'tope_mes = Ampo_A_Aplicar * Ampo_Max_1
                            tope_mes = Tope_Monto_Proporcional
                            If (Acumulado_Mes + tope_aplicar) >= tope_mes Then
                                'EAM(6.74) - Se cambio el calculo. Topea sobre el tope a aplicar que es según la cant de días trabajados y no por mes
                                tope_aplicar = tope_mes - Acumulado_Mes
                                'tope_aplicar = tope_aplicar - Acumulado_Mes
                            End If
                            
                            If CBool(USA_DEBUG) Then
                                Flog.writeline Espacios(Tabulador * 1) & " Tope Mes: " & tope_mes
                                Flog.writeline Espacios(Tabulador * 1) & " guardo Tope verificando Tope Proporcional y el tope del mes: " & tope_aplicar
                            End If
                            
                            StrSql = "UPDATE " & TTempWF_impproarg & " SET desborde = (" & rs_WF_impproarg!ipamonto & ") - " & tope_aplicar & _
                                         ", ipamonto = " & tope_aplicar & _
                                         ",ipacant = " & Aux_Cantidad & _
                                         "WHERE acunro = " & rs_WF_impproarg!acuNro & " AND tconnro = " & rs_WF_impproarg!tconnro
                                objConn.Execute StrSql, , adExecuteNoRecords
                                
                                If Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde <> 0 Then
                                    Monto_Desborde = (rs_WF_impproarg!ipamonto) - tope_aplicar
                                    If objCache_Acu_Liq_Monto.EsSimboloDefinido(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde)) Then
                                        'levanto el monto
                                        Monto_Desborde = Monto_Desborde + objCache_Acu_Liq_Monto.Valor(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde))
                                        'borro y vuelvo a insertar
                                        Call objCache_Acu_Liq_Monto.Borrar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde))
                                        Call objCache_Acu_Liq_MontoReal.Borrar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde))
                                        Call objCache_Acu_Liq_Cantidad.Borrar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde))
                                        
                                        Call objCache_Acu_Liq_Monto.Insertar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde), Monto_Desborde)
                                        Call objCache_Acu_Liq_MontoReal.Insertar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde), Monto_Desborde)
                                        Call objCache_Acu_Liq_Cantidad.Insertar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde), 0)
                                    Else
                                        Call objCache_Acu_Liq_Monto.Insertar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde), Monto_Desborde)
                                        Call objCache_Acu_Liq_MontoReal.Insertar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde), Monto_Desborde)
                                        Call objCache_Acu_Liq_Cantidad.Insertar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde), 0)
                                    End If
                                End If
                            'End If
                        End If

           Case 2:
                    If CBool(USA_DEBUG) Then Flog.writeline Espacios(Tabulador * 1) & " Tope de LAR "
                    
                    'Primer topeo contra proporcional
                    If CBool(USA_DEBUG) Then Flog.writeline Espacios(Tabulador * 1) & " Tope contra el proporcional"
                    If rs_WF_impproarg!ipamonto <= Tope_Monto_Proporcional Then
                        montoTope1 = rs_WF_impproarg!ipamonto
                    Else
                        montoTope1 = Tope_Monto_Proporcional
                    End If
                    If CBool(USA_DEBUG) Then Flog.writeline Espacios(Tabulador * 2) & " Monto topeado " & montoTope1
                    
                    'Segundo topeo contra acumulador mensual
                    If CBool(USA_DEBUG) Then Flog.writeline Espacios(Tabulador * 1) & " Tope contra el Acumulado Mensual"
                    If (montoTope1 + Acumulado_Mes) <= Tope_Monto_Mensual Then
                        montoTope2 = montoTope1
                        Desborde = 0
                        Acumulado_Mes = Acumulado_Mes - montoTope2
                    Else
                        montoTope2 = Tope_Monto_Mensual - Acumulado_Mes
                        Desborde = (Acumulado_Mes + montoTope2) - Tope_Monto_Mensual
                        Acumulado_Mes = 0
                    End If
                    If CBool(USA_DEBUG) Then Flog.writeline Espacios(Tabulador * 2) & " Monto topeado " & montoTope2
                    
                    'Guardo valores
                    StrSql = "UPDATE " & TTempWF_impproarg & " SET desborde = " & Desborde & _
                             ", ipamonto = " & montoTope2 & _
                             ",ipacant = " & Aux_Cantidad & _
                             "WHERE acunro = " & rs_WF_impproarg!acuNro & " AND tconnro = " & rs_WF_impproarg!tconnro
                    objConn.Execute StrSql, , adExecuteNoRecords
                    If CBool(USA_DEBUG) Then Flog.writeline Espacios(Tabulador * 1) & " Guardo Valores"
                      
                    'VERRRRRRR
                    'If CBool(USA_DEBUG) Then Flog.writeline Espacios(Tabulador * 1) & " guardo Tope verificando Tope Proporcional y el Acumulado del mes: " & Tope_Monto_Proporcional - Acumulado_Mes
                    If Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde <> 0 Then
                        Monto_Desborde = Desborde
                        If objCache_Acu_Liq_Monto.EsSimboloDefinido(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde)) Then
                            'levanto el monto
                            Monto_Desborde = Monto_Desborde + objCache_Acu_Liq_Monto.Valor(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde))
                            'borro y vuelvo a insertar
                            Call objCache_Acu_Liq_Monto.Borrar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde))
                            Call objCache_Acu_Liq_MontoReal.Borrar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde))
                            Call objCache_Acu_Liq_Cantidad.Borrar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde))
                            
                            Call objCache_Acu_Liq_Monto.Insertar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde), Monto_Desborde)
                            Call objCache_Acu_Liq_MontoReal.Insertar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde), Monto_Desborde)
                            Call objCache_Acu_Liq_Cantidad.Insertar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde), 0)
                        Else
                            Call objCache_Acu_Liq_Monto.Insertar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde), Monto_Desborde)
                            Call objCache_Acu_Liq_MontoReal.Insertar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde), Monto_Desborde)
                            Call objCache_Acu_Liq_Cantidad.Insertar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde), 0)
                        End If
                    End If
                      
                        'Busco para tipo 1
                        'FGZ - 04/06/2012 ----------------------------
                        'StrSql = "SELECT * FROM " & TTempWF_impproarg
                        StrSql = "SELECT ipamonto FROM " & TTempWF_impproarg
                        StrSql = StrSql & " WHERE acunro = " & rs_WF_impproarg!acuNro & " AND tconnro = 1"
                        'If rs_Aux_WF_impproarg.State = adStateOpen Then rs_Aux_WF_impproarg.Close
                        OpenRecordset StrSql, rs_Aux_WF_impproarg
                        If Not rs_Aux_WF_impproarg.EOF Then
                            Aux_impproarg1 = rs_Aux_WF_impproarg!ipamonto
                        Else
                            Aux_impproarg1 = 0
                        End If

                        'Busco para tipo 2
                        'FGZ - 04/06/2012 ----------------------------
                        'StrSql = "SELECT * FROM " & TTempWF_impproarg
                        StrSql = "SELECT ipamonto FROM " & TTempWF_impproarg
                        StrSql = StrSql & " WHERE acunro = " & rs_WF_impproarg!acuNro & " AND tconnro = 2"
                        'If rs_Aux_WF_impproarg.State = adStateOpen Then rs_Aux_WF_impproarg.Close
                        OpenRecordset StrSql, rs_Aux_WF_impproarg
                        If Not rs_Aux_WF_impproarg.EOF Then
                            Aux_impproarg2 = rs_Aux_WF_impproarg!ipamonto
                        Else
                            Aux_impproarg2 = 0
                        End If '

                        If CBool(USA_DEBUG) Then Flog.writeline Espacios(Tabulador * 1) & " Nuevo Tope Sueldo + Lar: " & Aux_impproarg1 + Aux_impproarg2

                        If (Aux_impproarg1 + Aux_impproarg2) > Valor_Ampo_Cont * Ampo_Max_2 Then
                            ' Actualizo
                            StrSql = "UPDATE " & TTempWF_impproarg & " SET desborde = (" & Acumulado_Mes & " + ipamonto) - " & Tope_Monto_Proporcional & _
                                     ", ipamonto = " & Aux_impproarg2 & _
                                     ",ipacant = " & Aux_Cantidad & _
                                     "WHERE acunro = " & rs_WF_impproarg!acuNro & " AND tconnro = " & rs_WF_impproarg!tconnro
                            objConn.Execute StrSql, , adExecuteNoRecords
                        If CBool(USA_DEBUG) Then Flog.writeline Espacios(Tabulador * 1) & " Si Tope Sueldo + Lar > Tope Max Vac guardo: " & Aux_impproarg2

                        End If
                        
                    Case 3:
                      If CBool(USA_DEBUG) Then
                        Flog.writeline Espacios(Tabulador * 1) & " Tope de SAC "
                        Flog.writeline Espacios(Tabulador * 1) & " Tope Semestral: " & Tope_Monto_Semestral
                      End If
                      
                            If rs_WF_impproarg!ipamonto > Tope_Monto_Proporcional Then
                                If (Acumulado_Mes + rs_WF_impproarg!ipamonto > Tope_Monto_Semestral) Then
                                    If Tope_Monto_Proporcional < Tope_Monto_Semestral Then
                                        'Agregado por Maxi 25/06/2009
                                        If Acumulado_Mes = 0 Then
                                            StrSql = "UPDATE " & TTempWF_impproarg & " SET desborde = (" & rs_WF_impproarg!ipamonto & ") - " & (Tope_Monto_Proporcional) & _
                                                     ", ipamonto = " & (Tope_Monto_Proporcional) & _
                                                     ",ipacant = " & Aux_Cantidad & _
                                                     "WHERE acunro = " & rs_WF_impproarg!acuNro & " AND tconnro = " & rs_WF_impproarg!tconnro
                                            If CBool(USA_DEBUG) Then Flog.writeline Espacios(Tabulador * 1) & " Si Acum Mes + Monto > Tope Semestral < Tope Proporcional y  guardo: " & Tope_Monto_Proporcional
                                         Else
                                            StrSql = "UPDATE " & TTempWF_impproarg & " SET desborde = (" & rs_WF_impproarg!ipamonto & ") - " & (Tope_Monto_Proporcional - Acumulado_Mes) & _
                                                     ", ipamonto = " & (Tope_Monto_Proporcional - Acumulado_Mes) & _
                                                     ",ipacant = " & Aux_Cantidad & _
                                                     "WHERE acunro = " & rs_WF_impproarg!acuNro & " AND tconnro = " & rs_WF_impproarg!tconnro
                                            If CBool(USA_DEBUG) Then Flog.writeline Espacios(Tabulador * 1) & " Si Acum Mes + Monto > Tope Semestral < Tope Proporcional y  guardo: " & Tope_Monto_Proporcional
                                        End If
                                    Else
                                        StrSql = "UPDATE " & TTempWF_impproarg & " SET desborde = (" & rs_WF_impproarg!ipamonto & ") - " & (Tope_Monto_Semestral - Acumulado_Mes) & _
                                                 ", ipamonto = " & (Tope_Monto_Semestral - Acumulado_Mes) & _
                                                 ",ipacant = " & Aux_Cantidad & _
                                                 "WHERE acunro = " & rs_WF_impproarg!acuNro & " AND tconnro = " & rs_WF_impproarg!tconnro
                                        If CBool(USA_DEBUG) Then Flog.writeline Espacios(Tabulador * 1) & " Si Acum Mes + Monto > Tope Semestral > Tope Proporcional y  guardo: " & Tope_Monto_Semestral - Acumulado_Mes
                                    End If
                                    objConn.Execute StrSql, , adExecuteNoRecords
                                    If Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde <> 0 Then
                                        If Tope_Monto_Proporcional < Tope_Monto_Semestral Then
                                            Monto_Desborde = (rs_WF_impproarg!ipamonto) - (Tope_Monto_Proporcional)
                                        Else
                                            Monto_Desborde = (rs_WF_impproarg!ipamonto) - (Tope_Monto_Semestral - Acumulado_Mes)
                                        End If
                                        If objCache_Acu_Liq_Monto.EsSimboloDefinido(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde)) Then
                                            'levanto el monto
                                            Monto_Desborde = Monto_Desborde + objCache_Acu_Liq_Monto.Valor(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde))
                                            'borro y vuelvo a insertar
                                            Call objCache_Acu_Liq_Monto.Borrar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde))
                                            Call objCache_Acu_Liq_MontoReal.Borrar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde))
                                            Call objCache_Acu_Liq_Cantidad.Borrar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde))
                                            
                                            Call objCache_Acu_Liq_Monto.Insertar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde), Monto_Desborde)
                                            Call objCache_Acu_Liq_MontoReal.Insertar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde), Monto_Desborde)
                                            Call objCache_Acu_Liq_Cantidad.Insertar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde), 0)
                                        Else
                                            Call objCache_Acu_Liq_Monto.Insertar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde), Monto_Desborde)
                                            Call objCache_Acu_Liq_MontoReal.Insertar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde), Monto_Desborde)
                                            Call objCache_Acu_Liq_Cantidad.Insertar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde), 0)
                                        End If
                                    End If
                                Else
                                    If Acumulado_Mes = 0 Then
                                            StrSql = "UPDATE " & TTempWF_impproarg & " SET desborde = (" & rs_WF_impproarg!ipamonto & ") - " & (Tope_Monto_Proporcional) & _
                                                     ", ipamonto = " & (Tope_Monto_Proporcional) & _
                                                     ",ipacant = " & Aux_Cantidad & _
                                                     " WHERE acunro = " & rs_WF_impproarg!acuNro & " AND tconnro = " & rs_WF_impproarg!tconnro
                                            If CBool(USA_DEBUG) Then Flog.writeline Espacios(Tabulador * 1) & " Si Acum Mes + Monto > Tope Semestral < Tope Proporcional y  guardo: " & Tope_Monto_Proporcional
                                         Else
                                            StrSql = "UPDATE " & TTempWF_impproarg & " SET desborde = (" & rs_WF_impproarg!ipamonto & ") - " & (Tope_Monto_Proporcional - Acumulado_Mes) & _
                                                     ", ipamonto = " & (Tope_Monto_Proporcional - Acumulado_Mes) & _
                                                     ",ipacant = " & Aux_Cantidad & _
                                                     " WHERE acunro = " & rs_WF_impproarg!acuNro & " AND tconnro = " & rs_WF_impproarg!tconnro
                                            If CBool(USA_DEBUG) Then Flog.writeline Espacios(Tabulador * 1) & " Si Acum Mes + Monto > Tope Semestral < Tope Proporcional y  guardo: " & (Tope_Monto_Proporcional - Acumulado_Mes)
                                        End If
                                       objConn.Execute StrSql, , adExecuteNoRecords
                                    
                                    'If CBool(USA_DEBUG) Then Flog.writeline Espacios(Tabulador * 1) & "Si Tope_Monto_Proporcional < Tope Semestral guardo: " & Tope_Monto_Proporcional
                                    
                                    If Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde <> 0 Then
                                        If Acumulado_Mes = 0 Then
                                          Monto_Desborde = (rs_WF_impproarg!ipamonto) - Tope_Monto_Proporcional
                                        Else
                                          Monto_Desborde = (rs_WF_impproarg!ipamonto) - (Tope_Monto_Proporcional - Acumulado_Mes)
                                        End If
                                        If objCache_Acu_Liq_Monto.EsSimboloDefinido(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde)) Then
                                            'levanto el monto
                                            Monto_Desborde = Monto_Desborde + objCache_Acu_Liq_Monto.Valor(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde))
                                            'borro y vuelvo a insertar
                                            Call objCache_Acu_Liq_Monto.Borrar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde))
                                            Call objCache_Acu_Liq_MontoReal.Borrar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde))
                                            Call objCache_Acu_Liq_Cantidad.Borrar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde))
                                            
                                            Call objCache_Acu_Liq_Monto.Insertar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde), Monto_Desborde)
                                            Call objCache_Acu_Liq_MontoReal.Insertar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde), Monto_Desborde)
                                            Call objCache_Acu_Liq_Cantidad.Insertar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde), 0)
                                        Else
                                            Call objCache_Acu_Liq_Monto.Insertar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde), Monto_Desborde)
                                            Call objCache_Acu_Liq_MontoReal.Insertar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde), Monto_Desborde)
                                            Call objCache_Acu_Liq_Cantidad.Insertar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde), 0)
                                        End If
                                    End If
                                End If
                            Else
                                If (Acumulado_Mes + rs_WF_impproarg!ipamonto > Tope_Monto_Proporcional) Then
                                    'Maxi 04/05/2010 error en tope de SAC cuando tiene algo acum en el semestre
                                            
                                       StrSql = "UPDATE " & TTempWF_impproarg & " SET desborde = (" & rs_WF_impproarg!ipamonto & ") - " & (Tope_Monto_Proporcional - Acumulado_Mes) & _
                                             ", ipamonto = " & Tope_Monto_Proporcional - Acumulado_Mes & _
                                             ",ipacant = " & Aux_Cantidad & _
                                             "WHERE acunro = " & rs_WF_impproarg!acuNro & " AND tconnro = " & rs_WF_impproarg!tconnro
                                    
                                    If CBool(USA_DEBUG) Then Flog.writeline Espacios(Tabulador * 1) & "Si Monto + Acumulado > Tope Semestral guardo: " & Tope_Monto_Proporcional - Acumulado_Mes
                                    
                                    objConn.Execute StrSql, , adExecuteNoRecords
                                    If Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde <> 0 Then
                                        Monto_Desborde = (rs_WF_impproarg!ipamonto) - (Tope_Monto_Semestral - Acumulado_Mes)
                                        If objCache_Acu_Liq_Monto.EsSimboloDefinido(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde)) Then
                                            'levanto el monto
                                            Monto_Desborde = Monto_Desborde + objCache_Acu_Liq_Monto.Valor(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde))
                                            'borro y vuelvo a insertar
                                            Call objCache_Acu_Liq_Monto.Borrar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde))
                                            Call objCache_Acu_Liq_MontoReal.Borrar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde))
                                            Call objCache_Acu_Liq_Cantidad.Borrar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde))
                                            
                                            Call objCache_Acu_Liq_Monto.Insertar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde), Monto_Desborde)
                                            Call objCache_Acu_Liq_MontoReal.Insertar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde), Monto_Desborde)
                                            Call objCache_Acu_Liq_Cantidad.Insertar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde), 0)
                                        Else
                                            Call objCache_Acu_Liq_Monto.Insertar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde), Monto_Desborde)
                                            Call objCache_Acu_Liq_MontoReal.Insertar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde), Monto_Desborde)
                                            Call objCache_Acu_Liq_Cantidad.Insertar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde), 0)
                                        End If
                                    End If
                                    
                                Else
                               
                                  If (Acumulado_Mes + rs_WF_impproarg!ipamonto > Tope_Monto_Semestral) Then
                                    'Maxi 27/04/2009 error en tope de SAC cuando tiene algo acum en el semestre
                                            
                                       StrSql = "UPDATE " & TTempWF_impproarg & " SET desborde = (" & rs_WF_impproarg!ipamonto & ") - " & (Tope_Monto_Semestral - Acumulado_Mes) & _
                                             ", ipamonto = " & Tope_Monto_Semestral - Acumulado_Mes & _
                                             ",ipacant = " & Aux_Cantidad & _
                                             "WHERE acunro = " & rs_WF_impproarg!acuNro & " AND tconnro = " & rs_WF_impproarg!tconnro
                                    
                                    If CBool(USA_DEBUG) Then Flog.writeline Espacios(Tabulador * 1) & "Si Monto + Acum > Tope_Monto_Semestral guardo: " & Tope_Monto_Semestral - Acumulado_Mes
                                    
                                    objConn.Execute StrSql, , adExecuteNoRecords
                                    If Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde <> 0 Then
                                        Monto_Desborde = (rs_WF_impproarg!ipamonto) - (Tope_Monto_Semestral - Acumulado_Mes)
                                        If objCache_Acu_Liq_Monto.EsSimboloDefinido(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde)) Then
                                            'levanto el monto
                                            Monto_Desborde = Monto_Desborde + objCache_Acu_Liq_Monto.Valor(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde))
                                            'borro y vuelvo a insertar
                                            Call objCache_Acu_Liq_Monto.Borrar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde))
                                            Call objCache_Acu_Liq_MontoReal.Borrar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde))
                                            Call objCache_Acu_Liq_Cantidad.Borrar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde))
                                            
                                            Call objCache_Acu_Liq_Monto.Insertar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde), Monto_Desborde)
                                            Call objCache_Acu_Liq_MontoReal.Insertar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde), Monto_Desborde)
                                            Call objCache_Acu_Liq_Cantidad.Insertar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde), 0)
                                        Else
                                            Call objCache_Acu_Liq_Monto.Insertar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde), Monto_Desborde)
                                            Call objCache_Acu_Liq_MontoReal.Insertar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde), Monto_Desborde)
                                            Call objCache_Acu_Liq_Cantidad.Insertar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde), 0)
                                        End If
                                    End If
                                    
                                   Else
                                       StrSql = "UPDATE " & TTempWF_impproarg & " SET desborde = (" & rs_WF_impproarg!ipamonto & ") - " & rs_WF_impproarg!ipamonto & _
                                             ", ipamonto = " & rs_WF_impproarg!ipamonto & _
                                             ",ipacant = " & Aux_Cantidad & _
                                             "WHERE acunro = " & rs_WF_impproarg!acuNro & " AND tconnro = " & rs_WF_impproarg!tconnro
                                    
                                    If CBool(USA_DEBUG) Then Flog.writeline Espacios(Tabulador * 1) & "Si Monto + Acum < Tope_Monto_Semestral guardo: " & rs_WF_impproarg!ipamonto
                                    
                                    objConn.Execute StrSql, , adExecuteNoRecords
                                    If Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde <> 0 Then
                                        Monto_Desborde = 0
                                        If objCache_Acu_Liq_Monto.EsSimboloDefinido(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde)) Then
                                            'levanto el monto
                                            Monto_Desborde = Monto_Desborde + objCache_Acu_Liq_Monto.Valor(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde))
                                            'borro y vuelvo a insertar
                                            Call objCache_Acu_Liq_Monto.Borrar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde))
                                            Call objCache_Acu_Liq_MontoReal.Borrar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde))
                                            Call objCache_Acu_Liq_Cantidad.Borrar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde))
                                            
                                            Call objCache_Acu_Liq_Monto.Insertar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde), Monto_Desborde)
                                            Call objCache_Acu_Liq_MontoReal.Insertar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde), Monto_Desborde)
                                            Call objCache_Acu_Liq_Cantidad.Insertar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde), 0)
                                        Else
                                            Call objCache_Acu_Liq_Monto.Insertar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde), Monto_Desborde)
                                            Call objCache_Acu_Liq_MontoReal.Insertar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde), Monto_Desborde)
                                            Call objCache_Acu_Liq_Cantidad.Insertar_Simbolo(CStr(Arr_Acumulador(rs_WF_impproarg!acuNro).acudesborde), 0)
                                        End If
                                    End If
                                   End If
                                   
                                End If
                                   
                             End If
                   End Select
                    
                    'FGZ - 04/06/2012 --------------------------------------------------------
                    'StrSql = "SELECT * FROM " & TTempWF_impproarg & " ORDER BY acunro, tconnro"
                    StrSql = "SELECT acunro , ipacant, ipamonto, tconnro, desborde, tope_aporte FROM " & TTempWF_impproarg & " ORDER BY acunro, tconnro"
                    OpenRecordset StrSql, rs_WF_impproarg
                    Do While (Not rs_WF_impproarg.EOF) And ((rs_WF_impproarg!acuNro < Aux_AcuNro) Or (rs_WF_impproarg!acuNro = Aux_AcuNro) And (rs_WF_impproarg!tconnro < Aux_TconNro))
                        rs_WF_impproarg.MoveNext
                    Loop
                    
                ' Crear el imponible del proceso
                'FGZ - 08/06/2012 -----------
                StrSql = "SELECT acunro FROM impproarg " & _
                         " WHERE acunro = " & rs_WF_impproarg!acuNro & _
                         " AND tconnro = " & rs_WF_impproarg!tconnro & " AND cliqnro = " & buliq_cabliq!cliqnro
                OpenRecordset StrSql, rs_ImpPro
                If rs_ImpPro.EOF Then
                    StrSql = "INSERT INTO impproarg (cliqnro,acunro,ipacant,ipamonto,tconnro)" & _
                             " VALUES (" & buliq_cabliq!cliqnro & "," & rs_WF_impproarg!acuNro & _
                             "," & rs_WF_impproarg!ipacant & "," & rs_WF_impproarg!ipamonto & "," & rs_WF_impproarg!tconnro & ")"
                Else
                    ' Actualizo
                    StrSql = "UPDATE impproarg SET ipamonto = " & rs_WF_impproarg!ipamonto & _
                             ", ipacant = " & rs_WF_impproarg!ipacant & _
                             " WHERE acunro = " & rs_WF_impproarg!acuNro & " AND tconnro = " & rs_WF_impproarg!tconnro & _
                             " AND cliqnro = " & buliq_cabliq!cliqnro
                End If
                objConn.Execute StrSql, , adExecuteNoRecords
                
                AcumuladoDeMonto = AcumuladoDeMonto + rs_WF_impproarg!ipamonto
                AcumuladoDeCant = AcumuladoDeCant + rs_WF_impproarg!ipacant
                
                
                rs_WF_impproarg.MoveNext
                If (rs_WF_impproarg.EOF) Then
                    Actualiza = True
                    rs_WF_impproarg.MovePrevious
                Else
                    If ((rs_WF_impproarg!acuNro > Aux_AcuNro) Or (rs_WF_impproarg!acuNro = Aux_AcuNro) And (rs_WF_impproarg!tconnro > Aux_TconNro)) Then
                        Actualiza = True
                        rs_WF_impproarg.MovePrevious
                    End If
                End If
                
                If Actualiza Then
                    If objCache_Acu_Liq_Monto.EsSimboloDefinido(CStr(rs_WF_impproarg!acuNro)) Then
                        'levanto el monto
                        Aux_Acu_Monto = objCache_Acu_Liq_Monto.Valor(CStr(rs_WF_impproarg!acuNro))
                        
                        'borro y vuelvo a insertar
                        Call objCache_Acu_Liq_Monto.Borrar_Simbolo(CStr(rs_WF_impproarg!acuNro))
                        Call objCache_Acu_Liq_MontoReal.Borrar_Simbolo(CStr(rs_WF_impproarg!acuNro))
                        Call objCache_Acu_Liq_Cantidad.Borrar_Simbolo(CStr(rs_WF_impproarg!acuNro))
                        
                        Call objCache_Acu_Liq_Monto.Insertar_Simbolo(CStr(rs_WF_impproarg!acuNro), AcumuladoDeMonto)
                        Call objCache_Acu_Liq_MontoReal.Insertar_Simbolo(CStr(rs_WF_impproarg!acuNro), Aux_Acu_Monto)
                        Call objCache_Acu_Liq_Cantidad.Insertar_Simbolo(CStr(rs_WF_impproarg!acuNro), AcumuladoDeCant)
                    End If
                    Actualiza = False
                End If
                
                'FGZ - 20/05/2011 ---------------------------------------------------------
                ' Solo se llama a establecer_Impgralarg() cuando usa_debug
                '   por esa razon se saca
                'If CBool(USA_DEBUG) Then
                '    Call Establecer_Impgralarg(NroProc, rs_WF_impproarg!tconnro)
                '    Select Case rs_WF_impproarg!tconnro
                '    Case 1:
                '        Texto = "Imponible sueldo del proceso(Acu " + Format(rs_WF_impproarg!acuNro, "000") + ")"
                '        Flog.writeline Espacios(Tabulador * 1) & Texto & rs_WF_impproarg!ipamonto
                '
                '    Case 2:
                '        Texto = "Imponible LAR del proceso(Acu " + Format(rs_WF_impproarg!acuNro, "000") + ")"
                '        Flog.writeline Espacios(Tabulador * 1) & Texto & rs_WF_impproarg!ipamonto
                '
                '    Case 3:
                '        Texto = "Imponible SAC del proceso(Acu " + Format(rs_WF_impproarg!acuNro, "000") + ")"
                '        Flog.writeline Espacios(Tabulador * 1) & Texto & rs_WF_impproarg!ipamonto
                '
                '    End Select
                'End If
               '
               ' If CBool(USA_DEBUG) Then
               '     Call Establecer_Impgralarg(NroProc, rs_WF_impproarg!tconnro)
               '     If Not buliq_impgralarg.EOF Then
               '         Select Case rs_WF_impproarg!tconnro
               '         Case 1:
               '             Flog.writeline Espacios(Tabulador * 1) & "Imponible Sueldo de Argentina " & buliq_impgralarg!ipgtopemonto
               '         Case 2:
               '             Flog.writeline Espacios(Tabulador * 1) & "Imponible LAR de Argentina " & buliq_impgralarg!ipgtopemonto
               '         Case 3:
               '             Flog.writeline Espacios(Tabulador * 1) & "Imponible SAC de Argentina" & buliq_impgralarg!ipgtopemonto
               '         End Select
               '     End If
               ' End If
               'FGZ - 20/05/2011 ---------------------------------------------------------
                
                rs_WF_impproarg.MoveNext
            Loop
        End If 'If (Arr_Conceptos(Concepto_Actual).tconnro > 3) And (Not Calculo_Topes) Then
    
    TiempoInicialConcepto = GetTickCount
    If CBool(USA_DEBUG) Then
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 1) & "Concepto: " & Arr_conceptos(Concepto_Actual).ConcCod
    End If
            
        '-------------- Alcance del concepto -----------------------
        ' Nivel
        'si es 0 busca el origen con el empleado
        'si es 1 hay que buscar la estructura
        'si es 2 listo (global)
            
        Alcance = False
        ' Antes esta consulta se ejecutaba una vez por cada concepto y por cada empleado ==> Orden(NxM)
        ' Ahora esta consulta esta cargada en un vector al inicio del proceso y por tanto se ejecuta 1 sola vez en todo el proceso ==> Orden(1)
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 1) & "-------------- Alcance del concepto -----------------------"
        End If
        Indice_Actual_Cge_Segun = Indice_Arr_Cge_Segun()
        
        ' FGZ - 18/03/2004 fecha fin
        ' cambio buliq_proceso!profecfin x Empleado_Fecha_Fin
        If Not Indice_Actual_Cge_Segun = 0 Then
            ' Alcance por estructura
            If Arr_Cge_Segun(Indice_Actual_Cge_Segun).Nivel = 1 Then
                If CBool(USA_DEBUG) Then
                    'Flog.writeline Espacios(Tabulador * 1) & "If Arr_Cge_Segun(Indice_Actual_Cge_Segun).Nivel = 1 Then"
                    Flog.writeline Espacios(Tabulador * 1) & "Alcance por estructura"
                End If
                
                'FGZ - 27/05/2011 ---------------------------------------------------
                If Not EsNulo(Arr_His_EstructuraEmp(Arr_Cge_Segun(Indice_Actual_Cge_Segun).Entidad).Estrnro) Then
                    If CBool(USA_DEBUG) Then
                        Flog.writeline Espacios(Tabulador * 1) & "Estructura " & Arr_His_EstructuraEmp(Arr_Cge_Segun(Indice_Actual_Cge_Segun).Entidad).Estrnro
                    End If
                    If Existe_Origen(Indice_Actual_Cge_Segun, Arr_His_EstructuraEmp(Arr_Cge_Segun(Indice_Actual_Cge_Segun).Entidad).Estrnro) Then
                        Alcance = True
                    Else
                        Alcance = False
                    End If
                Else
                    If CBool(USA_DEBUG) Then
                        Flog.writeline Espacios(Tabulador * 1) & "No hay estructuras de tipo " & Arr_Cge_Segun(Indice_Actual_Cge_Segun).Entidad
                    End If
                    Alcance = False
                End If
                
                
                'If rs_His_Estructura.State = adStateOpen Then rs_His_Estructura.Close
                'StrSql = " SELECT tenro, estrnro FROM his_estructura " & _
                '         " WHERE ternro = " & buliq_empleado!Ternro & " AND " & _
                '         " tenro =" & Arr_Cge_Segun(Indice_Actual_Cge_Segun).Entidad & " AND " & _
                '         " (htetdesde <= " & ConvFecha(Empleado_Fecha_Fin) & ") AND " & _
                '         " ((" & ConvFecha(Empleado_Fecha_Fin) & " <= htethasta) or (htethasta is null))"
                'OpenRecordset StrSql, rs_His_Estructura
'               ' If CBool(USA_DEBUG) Then
'               '     'Flog.writeline Espacios(Tabulador * 1) & StrSql
'               ' End If
                'If Not rs_His_Estructura.EOF Then
                '    If CBool(USA_DEBUG) Then
                '        Flog.writeline Espacios(Tabulador * 1) & "Estructura " & rs_His_Estructura!Estrnro
                '    End If
                '    If Existe_Origen(Indice_Actual_Cge_Segun, rs_His_Estructura!Estrnro) Then
                '        Alcance = True
                '    Else
                '        Alcance = False
                '    End If
                'Else
                '    If CBool(USA_DEBUG) Then
                '        Flog.writeline Espacios(Tabulador * 1) & "No hay estructuras de tipo " & Arr_Cge_Segun(Indice_Actual_Cge_Segun).Entidad
                '    End If
                '    Alcance = False
                'End If
                'FGZ - 27/05/2011 ---------------------------------------------------
                
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 1) & "Alcance = " & CStr(CBool(Alcance))
                End If
                
                If Not Alcance Then
                    If CBool(USA_DEBUG) Then
                        Flog.writeline Espacios(Tabulador * 1) & "No se encontró alcance del concepto por estructura"
                    End If
                    If HACE_TRAZA Then
                        Texto = "No se encontró alcance del concepto por estructura"
                        Call InsertarTraza(NroCab, Arr_conceptos(Concepto_Actual).ConcNro, 0, Texto, 0)
                    End If
                    ' SIGUIENTE CONCEPTO
                    GoTo SiguienteConcepto
                End If
            End If
        Else
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 1) & "No se encontró el alcance del concepto"
            End If
            
            If HACE_TRAZA Then
                Texto = "No se encontró el alcance del concepto"
                Call InsertarTraza(NroCab, Arr_conceptos(Concepto_Actual).ConcNro, 0, Texto, 0)
            End If
            ' SIGUIENTE CONCEPTO
            GoTo SiguienteConcepto
        End If 'If Not rs_Alcance.EOF Then
        
        If CBool(USA_DEBUG) Then
            Select Case Arr_Cge_Segun(Indice_Actual_Cge_Segun).Nivel
            Case 0: 'Individual
                Texto = "Alcance Individual empleado " + CStr(buliq_empleado!Empleg)
            Case 1: 'Estructura
                Texto = "Alcance Estructura Tipo: " & Arr_Cge_Segun(Indice_Actual_Cge_Segun).Entidad & " Estructura: " + CStr(Arr_Cge_Segun(Indice_Actual_Cge_Segun).Origen)
            Case 2: 'Global
                Texto = "Alcance Global"
            Case Else
            End Select
            Flog.writeline Espacios(Tabulador * 1) & Texto
        End If

        If HACE_TRAZA Then
            Select Case Arr_Cge_Segun(Indice_Actual_Cge_Segun).Nivel
            Case 0: 'Individual
                Texto = "Alcance Individual empleado " + CStr(buliq_empleado!Empleg)
            Case 1: 'Estructura
                Texto = "Alcance Estructura Tipo: " & Arr_Cge_Segun(Indice_Actual_Cge_Segun).Entidad & " Estructura: " + CStr(Arr_Cge_Segun(Indice_Actual_Cge_Segun).Origen)
            Case 2: 'Global
                Texto = "Alcance Global"
            Case Else
            End Select
            
            Call InsertarTraza(NroCab, Arr_conceptos(Concepto_Actual).ConcNro, 0, Texto, 0)
        End If
        
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 1) & "-------------- Parametros -----------------------"
        End If
        
        
'================================================================================
'MACHETAZO
Machetazo:
    Resultado = 0
    Ajustes = 0
    HayAjuste = False
    I = 0
    
If CBool(Arr_conceptos(Concepto_Actual).Concajuste) Then
        'Multiples Novedades de ajuste
        'FGZ - 04/06/2012 --------------------------------
        'StrSql = "SELECT * FROM novaju "
        StrSql = "SELECT nanro, navalor, navigencia, nadesde, nahasta, napliqdesde, napliqhasta FROM novaju " & _
                 " WHERE concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro & _
                 " AND empleado = " & buliq_empleado!Ternro
        OpenRecordset StrSql, rs_NovAjuste
    
        'FGZ - 13/12/2013 --------------------------------------------------------
        If Not rs_NovAjuste.EOF Then
            ReDim Preserve ArrNov(rs_NovAjuste.RecordCount) As TDistrib
            I = 0
        End If
        'FGZ - 13/12/2013 --------------------------------------------------------
    
    
        Do While Not rs_NovAjuste.EOF
            If FirmaActiva20 Then
                '/* Verificar si esta en el NIVEL FINAL DE FIRMA */
                    StrSql = "select cysfirfin from cysfirmas where cysfiryaaut = -1 AND cysfirfin = -1 " & _
                             " AND cysfircodext = '" & rs_NovAjuste!nanro & "' and cystipnro = 20"
                    OpenRecordset StrSql, rs_firmas
                    If rs_firmas.EOF Then
                        Firmado = False
                    Else
                        Firmado = True
                    End If
                If rs_firmas.State = adStateOpen Then rs_firmas.Close
            Else
                Firmado = True
            End If
            
            If Firmado Then
                ' FGZ - 18/04/2005
                'Se decidió sacar este control y permitir liquidar ajustes con valor 0
                
                ' No considerar ajustes en cero
                'If Not Round(rs_NovAjuste!navalor, Arr_conceptos(Concepto_Actual).Conccantdec) = 0 Then
                    If CBool(USA_DEBUG) Then
                        Flog.writeline Espacios(Tabulador * 1) & "Novedad de Ajuste " & rs_NovAjuste!navalor
                    End If
                    If HACE_TRAZA Then
                        Call InsertarTraza(NroCab, Buliq_Concepto(Concepto_Actual).ConcNro, 99988, "Novedad de Ajuste", rs_NovAjuste!navalor)
                    End If
                    
                    ' Valido la Vigencia
                    If CBool(rs_NovAjuste!navigencia) Then
                        If ((rs_NovAjuste!nahasta < Fecha_Inicio) Or (Fecha_Fin < rs_NovAjuste!nadesde)) Then
                            GoTo SiguienteAjuste
                        End If
                    End If
                    HayAjuste = True
                    Resultado = Resultado + Round(rs_NovAjuste!navalor, Arr_conceptos(Concepto_Actual).Conccantdec)
                    
                    'FGZ - 13/12/2013 --------------------------------------------------------
                    I = I + 1
                    ArrNov(I).ID = rs_NovAjuste!nanro
                    ArrNov(I).Valor = rs_NovAjuste!navalor
                    'FGZ - 13/12/2013 --------------------------------------------------------
                    
                    
                    If Not (EsNulo(rs_NovAjuste!napliqdesde) Or EsNulo(rs_NovAjuste!napliqhasta)) Then
                        If rs_NovAjuste!napliqdesde <> 0 And rs_NovAjuste!napliqhasta <> 0 Then
                            Retroactivo = True
                            pliqdesde = rs_NovAjuste!napliqdesde
                            pliqhasta = rs_NovAjuste!napliqhasta
                            Call CrearRetro(pliqdesde, pliqhasta, rs_NovAjuste!navalor, Buliq_Concepto(Concepto_Actual).ConcNro)
                        Else
                            Retroactivo = False
                        End If
                    Else
                        Retroactivo = False
                    End If
                    
            End If
            
            
SiguienteAjuste:
                'Siguiente Novedad
                rs_NovAjuste.MoveNext
            Loop
            
            If HayAjuste Then
                'FGZ - 13/12/2013 --------------------------------------------------------
                If I <> 0 Then
                    'Hay novedades de ajuste ==> reviso si hay distribucion
                    Call RevisarDistribucionNov(2)
                End If
                'FGZ - 13/12/2013 --------------------------------------------------------
                
                StrSql = "INSERT INTO detliq (" & _
                          "cliqnro,concnro,dlimonto,dlicant,ajustado,dlitexto,dliretro" & _
                          ") VALUES (" & buliq_cabliq!cliqnro & _
                          "," & Arr_conceptos(Concepto_Actual).ConcNro & _
                          "," & Round(Resultado, Arr_conceptos(Concepto_Actual).Conccantdec) & _
                          ", 0,"
                If Retroactivo Then
                    StrSql = StrSql & " -1,' ',-1)"
                Else
                    StrSql = StrSql & " 0,' ',0)"
                End If
                          
                          
                          
                 objConn.Execute StrSql, , adExecuteNoRecords
                 
                 ' inserto en el cache de detliq
                 ' El monto
                 Call objCache_detliq_Monto.Insertar_Simbolo(CStr(Arr_conceptos(Concepto_Actual).ConcNro), Resultado)
                 ' La cantidad
                 Call objCache_detliq_Cantidad.Insertar_Simbolo(CStr(Arr_conceptos(Concepto_Actual).ConcNro), 0)
                 
                If CBool(Arr_conceptos(Concepto_Actual).concretro) And Retroactivo Then ' Concepto retroactivo y Novedad Retroactiva
                    Concepto_Retroactivo = Retroactivo
                    concepto_pliqdesde = pliqdesde
                    concepto_pliqhasta = pliqhasta
                End If
            End If
End If
    

'================================================================================
    If Not HayAjuste Then 'Por Formula
        ' ---------------
        'FGZ - 26/01/2004
        Parametro = 0
        tpa_parametro = 0
        Concepto_Retroactivo = False
        concepto_pliqdesde = 0
        concepto_pliqhasta = 0
        'FGZ - 26/01/2004
        ' ---------------
        
        ' Limpiar los wf de pasaje de parametros
        Call LimpiarTempTable(TTempWF_Retroactivo)
        
        ' Resolucion de los parámetros de la fórmula del Concepto
       
        ' Antes esta consulta se ejecutaba una vez por cada formula, por cada concepto y por cada empleado ==> Orden(NxMxP)
        ' Ahora esta consulta esta cargada en un vector al inicio del proceso y por tanto se ejecuta 1 sola vez en todo el proceso ==> Orden(1)
        Indice_Actual_For_Tpa = 0
        Indice_Actual_For_Tpa = BuscarSiguiente_For_Tpa(Arr_conceptos(Concepto_Actual).fornro)
        
        'FGZ - 23/05/2011 ------------------------
        'If Arr_conceptos(Concepto_Actual).Conccod = "13000" Then
        '    Nueva_Tpa = True
        'End If
        'If Nueva_Tpa Then
            Call InicializarWF_Tpa(Arr_conceptos(Concepto_Actual).fornro)
        'Else
        '    Call LimpiarTempTable(TTempWF_tpa)
        'End If
        'FGZ - 23/05/2011 ------------------------
        
        Do While Indice_Actual_For_Tpa <> 0
        'Do While Not rs_For_Tpa.EOF
            TiempoInicialParametro = GetTickCount
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 2) & "Parametro: " & Arr_For_Tpa(Indice_Actual_For_Tpa).tpanro
            End If
            
            ' Antes esta consulta se ejecutaba una vez por cada formula, por cada concepto , por cada parametro y por cada empleado ==> Orden(NxMxPxQ)
            ' Ahora esta consulta esta cargada en un vector al inicio del proceso y por tanto se ejecuta 1 sola vez en todo el proceso ==> Orden(1)
            ' FGZ - 19/08/2004
            Indice_Actual_Cft_Segun = 0
            Indice_Actual_Cft_Segun = Buscar_Sig_Indice_Arr_Cft_Segun(Arr_For_Tpa(Indice_Actual_For_Tpa).tpanro, Arr_For_Tpa(Indice_Actual_For_Tpa).fornro)
        
            ' FGZ - 14/04/2004 'Se cambió el if por un While
            Alcance = False
            
            Do While Not Indice_Actual_Cft_Segun = 0 And Not Alcance
            
            'Do While Not rs_CftSegun.EOF And Not Alcance
                If Arr_Cft_Segun(Indice_Actual_Cft_Segun).Nivel = 1 Then
                    'lo tengo que buscar
                    'If rs_His_Estructura.State = adStateOpen Then rs_His_Estructura.Close
                    'FGZ - 27/05/2011 -------------------------------------
                    If Arr_Cft_Segun(Indice_Actual_Cft_Segun).Nivel = 1 And Arr_Cft_Segun(Indice_Actual_Cft_Segun).Origen = Arr_His_EstructuraEmp(Arr_Cft_Segun(Indice_Actual_Cft_Segun).Entidad).Estrnro Then
                        Alcance = True
                        Grupo = Arr_Cft_Segun(Indice_Actual_Cft_Segun).Origen
                    Else
                        Alcance = False
                        Grupo = 0
                    End If
                    
                    'StrSql = " SELECT tenro, estrnro FROM his_estructura " & _
                    '         " WHERE ternro = " & buliq_empleado!Ternro & " AND " & _
                    '         " tenro =" & Arr_Cft_Segun(Indice_Actual_Cft_Segun).Entidad & " AND " & _
                    '         " (htetdesde <= " & ConvFecha(Empleado_Fecha_Fin) & ") AND " & _
                    '         " ((" & ConvFecha(Empleado_Fecha_Fin) & " <= htethasta) or (htethasta is null))"
                    'OpenRecordset StrSql, rs_His_Estructura
                    'If Not rs_His_Estructura.EOF Then
                    '    If Arr_Cft_Segun(Indice_Actual_Cft_Segun).Nivel = 1 And Arr_Cft_Segun(Indice_Actual_Cft_Segun).Origen = rs_His_Estructura!estrnro Then
                    '        Alcance = True
                    '        Grupo = Arr_Cft_Segun(Indice_Actual_Cft_Segun).Origen
                    '    Else
                    '        Alcance = False
                    '        Grupo = 0
                    '    End If
                    'Else
                    '    Alcance = False
                    'End If
                    'FGZ - 27/05/2011 -------------------------------------
                Else
                    'el alcance es global
                    Alcance = True
                    ' FGZ - 21/05/2004
                    Grupo = 0
                    ' FGZ - 21/05/2004
                End If 'If rs_CftSegun.nivel = 1 Then
                
                If Not Alcance Then
                    'rs_CftSegun.MoveNext
                    Indice_Actual_Cft_Segun = Buscar_Sig_Indice_Arr_Cft_Segun(Arr_For_Tpa(Indice_Actual_For_Tpa).tpanro, Arr_For_Tpa(Indice_Actual_For_Tpa).fornro)
                End If
            Loop
            
            If Not Alcance Then
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 2) & "Parametro sin configurar el alcance"
                End If
            
                If HACE_TRAZA Then
                    Call InsertarTraza(NroCab, Arr_conceptos(Concepto_Actual).ConcNro, Arr_For_Tpa(Indice_Actual_For_Tpa).tpanro, "Parametro sin configurar el alcance", 0)
                End If
                ' SIGUIENTE CONCEPTO
                GoTo SiguienteConcepto
            End If
            
            Selecc = IIf(EsNulo(Arr_Cft_Segun(Indice_Actual_Cft_Segun).Selecc), "", Trim(Arr_Cft_Segun(Indice_Actual_Cft_Segun).Selecc))
            Indice_Actual = Indice_Arr_con_for_tpa(Arr_conceptos(Concepto_Actual).ConcNro, Arr_For_Tpa(Indice_Actual_For_Tpa).fornro, Arr_For_Tpa(Indice_Actual_For_Tpa).tpanro, Arr_Cft_Segun(Indice_Actual_Cft_Segun).Nivel, Selecc)
            
            If Indice_Actual = 0 Then
            'If rs_con_for_tpa.EOF Then
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 2) & "Parametro sin configurar la busqueda"
                End If
                
                If HACE_TRAZA Then
                    Call InsertarTraza(NroCab, Arr_conceptos(Concepto_Actual).ConcNro, Arr_For_Tpa(Indice_Actual_For_Tpa).tpanro, "Parametro sin configurar la busqueda", 0)
                End If
                ' SIGUIENTE CONCEPTO
                GoTo SiguienteConcepto
            End If
            
            If EsNulo(Arr_Programa(Arr_con_for_tpa(Indice_Actual).Prognro).Prognro) Then
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 2) & "Nro de busqueda no identificado"
                End If
            
                If HACE_TRAZA Then
                    Call InsertarTraza(NroCab, Arr_conceptos(Concepto_Actual).ConcNro, Arr_For_Tpa(Indice_Actual_For_Tpa).tpanro, "Nro de busqueda no identificado", 0)
                End If
                ' SIGUIENTE CONCEPTO
                GoTo SiguienteConcepto
            End If
            
            val = 0
            Valor = 0
            OK = False
            Retroactivo = False
            pliqdesde = 0
            pliqhasta = 0
            
            TiempoInicialBusqueda = GetTickCount
            ' si es automatico y la busqueda esta marcada como que puede usar cache, verificar el cache del empleado
            If CBool(Arr_con_for_tpa(Indice_Actual).cftauto) Then
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 3) & "Busqueda Automatica: " & CStr(Arr_Programa(Arr_con_for_tpa(Indice_Actual).Prognro).Prognro)
                End If
                
                If CBool(Arr_Programa(Arr_con_for_tpa(Indice_Actual).Prognro).Progarchest) Then 'si esta generada
                    If CBool(Arr_Programa(Arr_con_for_tpa(Indice_Actual).Prognro).Progcache) Then
                        If CBool(USA_DEBUG) Then
                            Flog.writeline Espacios(Tabulador * 3) & "Busca en Cache de la Busqueda"
                        End If
                       If objCache.EsSimboloDefinido(CStr(Arr_Programa(Arr_con_for_tpa(Indice_Actual).Prognro).Prognro)) Then
                            If CBool(USA_DEBUG) Then
                                Flog.writeline Espacios(Tabulador * 3) & " está en Cache "
                            End If
                            val = objCache.Valor(CStr(Arr_Programa(Arr_con_for_tpa(Indice_Actual).Prognro).Prognro))
                            OK = True
                        Else
                            If CBool(USA_DEBUG) Then
                                Flog.writeline Espacios(Tabulador * 3) & " NO está en cache. Ejecuto la busqueda"
                            End If
                            Call EjecutarBusqueda(Arr_Programa(Arr_con_for_tpa(Indice_Actual).Prognro).Tprognro, Arr_conceptos(Concepto_Actual).ConcNro, Arr_con_for_tpa(Indice_Actual).Prognro, val, fec, OK)
                            If CBool(USA_DEBUG) Then
                                TiempoFinalBusqueda = GetTickCount
                                Flog.writeline Espacios(Tabulador * 3) & "Tiempo de la busqueda: " & (TiempoFinalBusqueda - TiempoInicialBusqueda)
                            End If
                            
                            If Not OK Then
                                If CBool(USA_DEBUG) Then
                                    Flog.writeline Espacios(Tabulador * 3) & "Error en búsqueda de parametro"
                                End If
                            
                                If HACE_TRAZA Then
                                    Call InsertarTraza(NroCab, Arr_conceptos(Concepto_Actual).ConcNro, Arr_For_Tpa(Indice_Actual_For_Tpa).tpanro, "Error en búsqueda de parametro", 0)
                                End If
                                ' SIGUIENTE CONCEPTO
                                GoTo SiguienteConcepto
                            Else
                                If CBool(USA_DEBUG) Then
                                    Flog.writeline Espacios(Tabulador * 3) & "Resultado de la busqueda: " & val
                                End If
                            End If
                            ' insertar en el cache del empleado
                            Call objCache.Insertar_Simbolo(CStr(Arr_Programa(Arr_con_for_tpa(Indice_Actual).Prognro).Prognro), val)
                        End If
                    Else
                        ' busqueda automatica, primera vez
                        Call EjecutarBusqueda(Arr_Programa(Arr_con_for_tpa(Indice_Actual).Prognro).Tprognro, Arr_conceptos(Concepto_Actual).ConcNro, Arr_con_for_tpa(Indice_Actual).Prognro, val, fec, OK)
                        If CBool(USA_DEBUG) Then
                            TiempoFinalBusqueda = GetTickCount
                            Flog.writeline Espacios(Tabulador * 3) & "Tiempo de la busqueda: " & (TiempoFinalBusqueda - TiempoInicialBusqueda)
                        End If
                        
                        If Not OK Then
                            If CBool(USA_DEBUG) Then
                                Flog.writeline Espacios(Tabulador * 3) & "Error en búsqueda de parametro"
                            End If
                        
                            If HACE_TRAZA Then
                                Call InsertarTraza(NroCab, Arr_conceptos(Concepto_Actual).ConcNro, Arr_For_Tpa(Indice_Actual_For_Tpa).tpanro, "Error en búsqueda de parametro", 0)
                            End If
                            ' SIGUIENTE CONCEPTO
                            GoTo SiguienteConcepto
                        Else
                            If CBool(USA_DEBUG) Then
                                Flog.writeline Espacios(Tabulador * 3) & "Resultado de la busqueda: " & val
                            End If
                        End If
                    End If
                Else 'La busqueda no está generada
                    If CBool(USA_DEBUG) Then
                        Flog.writeline Espacios(Tabulador * 3) & " busqueda no generada "
                    End If
                
                    If HACE_TRAZA Then
                        Call InsertarTraza(NroCab, Arr_conceptos(Concepto_Actual).ConcNro, Arr_For_Tpa(Indice_Actual_For_Tpa).tpanro, "Búsqueda no está generada", 0)
                    End If
                End If
            Else
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 3) & "Busqueda NO Automatica: " & CStr(Arr_Programa(Arr_con_for_tpa(Indice_Actual).Prognro).Prognro)
                End If
                
                ' Busqueda no Automatica, Novedades buscar
                    If Not NovedadesHist Then
                        'Call Bus_NovGegi(Arr_Programa(rs_Con_For_Tpa!Prognro).prognro, Arr_Conceptos(Concepto_Actual).concnro, rs_For_Tpa!tpanro, fecha_inicio, fecha_fin, IIf(EsNulo(rs_CftSegun!origen), 0, rs_CftSegun!origen), ok, val)
                        ' FGZ - 21/05/2004
                        'Call Bus_NovGegi(Arr_Programa(Arr_con_for_tpa(Indice_Actual).Prognro).Prognro, Arr_conceptos(Concepto_Actual).concnro, Arr_For_Tpa(Indice_Actual_For_Tpa).tpanro, Fecha_Inicio, Fecha_Fin, Grupo, Ok, val)
                        'EAM- 6.67
                        'Call Bus_NovGegi(Arr_Programa(Arr_con_for_tpa(Indice_Actual).Prognro).Prognro, Arr_conceptos(Concepto_Actual).ConcNro, Arr_For_Tpa(Indice_Actual_For_Tpa).tpanro, Empleado_Fecha_Inicio, Empleado_Fecha_Fin, Grupo, OK, val)
                        Call Bus_NovGegiNEW(Arr_Programa(Arr_con_for_tpa(Indice_Actual).Prognro).Prognro, Arr_For_Tpa(Indice_Actual_For_Tpa).tpanro, Empleado_Fecha_Inicio, Empleado_Fecha_Fin, Grupo, OK, val, rs_novEmp, rs_NovEstr, rs_NovGral)
                        
                    Else
                        If CBool(USA_DEBUG) Then
                            Flog.writeline Espacios(Tabulador * 3) & "en historico "
                        End If
                        'Call Bus_NovGegiHis(Arr_Programa(rs_Con_For_Tpa!Prognro).prognro, Arr_Conceptos(Concepto_Actual).concnro, rs_For_Tpa!tpanro, fecha_inicio, fecha_fin, IIf(EsNulo(rs_CftSegun!origen), 0, rs_CftSegun!origen), ok, val)
                        ' FGZ - 21/05/2004
                        'Call Bus_NovGegiHis(Arr_Programa(Arr_con_for_tpa(Indice_Actual).Prognro).Prognro, Arr_conceptos(Concepto_Actual).concnro, Arr_For_Tpa(Indice_Actual_For_Tpa).tpanro, Fecha_Inicio, Fecha_Fin, Grupo, Ok, val)
                        Call Bus_NovGegiHis(Arr_Programa(Arr_con_for_tpa(Indice_Actual).Prognro).Prognro, Arr_conceptos(Concepto_Actual).ConcNro, Arr_For_Tpa(Indice_Actual_For_Tpa).tpanro, Empleado_Fecha_Inicio, Empleado_Fecha_Fin, Grupo, OK, val)
                    End If
                    If CBool(USA_DEBUG) Then
                        TiempoFinalBusqueda = GetTickCount
                        Flog.writeline Espacios(Tabulador * 3) & "Tiempo de la busqueda: " & (TiempoFinalBusqueda - TiempoInicialBusqueda)
                    End If
                    
                    If Not OK Then
                        If CBool(USA_DEBUG) Then
                            Flog.writeline Espacios(Tabulador * 3) & "No se encontró la Novedad"
                        End If
                    
                        If HACE_TRAZA Then
                            Call InsertarTraza(NroCab, Arr_conceptos(Concepto_Actual).ConcNro, Arr_For_Tpa(Indice_Actual_For_Tpa).tpanro, "No se encontró la Novedad", 0)
                        End If
                        ' SIGUIENTE CONCEPTO
                        GoTo SiguienteConcepto
                    Else
                        If CBool(USA_DEBUG) Then
                            Flog.writeline Espacios(Tabulador * 3) & "Resultado de la busqueda: " & val
                        End If
                    End If
            End If 'Busqueda de Novedad

            If OK Then
                ' se obtuvo el parametro satisfactoriamente
                ' inserto en el wf_tpa
                'If Nueva_Tpa Then
                    Call Actualizar_WF_Tpa(Indice_Actual_For_Tpa, Arr_For_Tpa(Indice_Actual_For_Tpa).tpanro, Arr_For_Tpa(Indice_Actual_For_Tpa).ftorden, Arr_conceptos(Concepto_Actual).concabr, val, fec)
                'Else
                '    Call insertar_wf_tpa(Arr_For_Tpa(Indice_Actual_For_Tpa).tpanro, Arr_For_Tpa(Indice_Actual_For_Tpa).ftorden, Arr_conceptos(Concepto_Actual).concabr, val, fec)
                'End If
                'Parametro = val
                If Arr_For_Tpa(Indice_Actual_For_Tpa).ftimprime Then
                    ' Guarda el parametro imprimible
                    Parametro = val
                    tpa_parametro = Arr_For_Tpa(Indice_Actual_For_Tpa).tpanro
                End If
                
                If CBool(Arr_conceptos(Concepto_Actual).concretro) And Retroactivo Then ' Concepto retroactivo y Novedad Retroactiva
                    'fecha_retro = fec
                    Concepto_Retroactivo = Retroactivo
                    concepto_pliqdesde = pliqdesde
                    concepto_pliqhasta = pliqhasta
                End If
                
                If HACE_TRAZA Then
                    If Arr_Cft_Segun(Indice_Actual_Cft_Segun).Nivel = 0 Then
                        Texto = "Excep. Indiv. "
                    Else
                        If Arr_Cft_Segun(Indice_Actual_Cft_Segun).Nivel = 1 Then
                            Texto = "Excep. x Estr. "
                        Else
                            Texto = ""
                        End If
                    End If
                    'FGZ - 10/06/2013 ----------------------------------------
                    If CBool(Arr_con_for_tpa(Indice_Actual).cftauto) Then
                        Texto = Texto & "busqueda: " & Arr_Programa(Arr_con_for_tpa(Indice_Actual).Prognro).Prognro & " - " & Arr_Programa(Arr_con_for_tpa(Indice_Actual).Prognro).Prognom & "(" & Trim(tipoNovedad) & ")"
                    Else
                        'Texto = Texto & "Busqueda NO Automatica: " & CStr(Arr_Programa(Arr_con_for_tpa(Indice_Actual).Prognro).Prognro)
                        'Texto = Texto & "Busqueda por Novedad: " & CStr(Arr_Programa(Arr_con_for_tpa(Indice_Actual).Prognro).Prognro)
                        Texto = Texto & "Novedad "
                    End If
                    'FGZ - 10/06/2013 ----------------------------------------
                    Call InsertarTraza(NroCab, Arr_conceptos(Concepto_Actual).ConcNro, Arr_For_Tpa(Indice_Actual_For_Tpa).tpanro, Texto, val)
                End If
                If CBool(USA_DEBUG) Then
                    If Arr_Cft_Segun(Indice_Actual_Cft_Segun).Nivel = 0 Then
                        Texto = "Excep. Indiv. "
                    Else
                        If Arr_Cft_Segun(Indice_Actual_Cft_Segun).Nivel = 1 Then
                            Texto = "Excep. x Estr. "
                        Else
                            Texto = ""
                        End If
                    End If
                    If Texto <> "" Then
                        Texto = Texto & "busqueda: " & Arr_Programa(Arr_con_for_tpa(Indice_Actual).Prognro).Prognro & " - " & Arr_Programa(Arr_con_for_tpa(Indice_Actual).Prognro).Prognom & "(" & Trim(tipoNovedad) & ")"
                        Flog.writeline Espacios(Tabulador * 3) & Texto
                    End If
                End If
            Else
                ' No se pudo obtener el parametro.
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 2) & "No se obtuvo el parametro"
                End If
                If HACE_TRAZA Then
                    Call InsertarTraza(NroCab, Arr_conceptos(Concepto_Actual).ConcNro, Arr_For_Tpa(Indice_Actual_For_Tpa).tpanro, "No se obtuvo el parametro", 0)
                End If
                ' SIGUIENTE CONCEPTO
                GoTo SiguienteConcepto
            End If
        
            If CBool(USA_DEBUG) Then
                TiempoFinalParametro = GetTickCount
                Flog.writeline Espacios(Tabulador * 2) & "Tiempo del parametro: " & (TiempoFinalParametro - TiempoInicialParametro)
            End If
            
            ' Siguiente for_tpa
            Indice_Actual_For_Tpa = BuscarSiguiente_For_Tpa(Arr_conceptos(Concepto_Actual).fornro)
            'rs_For_Tpa.MoveNext
        Loop
        
        'EJECUCION DE LA FORMULA DEL CONCEPTO
        exito = False
        Resultado = 0
                
              
        ' setear todos los parametros en wf_tpa en la tabla de simbolos
        
        'If Nueva_Tpa Then
            Call CargarTablaParametros_Nueva
        'Else
        '    Call CargarTablaParametros
        'End If
        
        ' reviso si la formula es externa o interna
        If Arr_conceptos(Concepto_Actual).Fortipo = 3 Then 'Configurable
            ' Evalua la expresion de la formula
            TiempoInicialFormula = GetTickCount
            'Resultado = cdbl(eval.Evaluate(Trim("( 1 =1 AND 1 =2) OR ( 11 =11 AND 1 =12)"), exito, True))
            'Resultado = CDbl(eval.Evaluate(Trim("SI ((( 12 =5) OR ( 12 =9) OR ( 12 =12));(1527.19 / 120 );0)"), exito, True))
            'Resultado = CDbl(eval.Evaluate(Trim("SI ( 1 =1 OR ( 1 =2 AND  90 =50); SI ( 4500 <=5000; 4500 /1000*1.5; SI ( 4500 <=10000;(( 4500 -5000)/1000*2)+7.5; SI ( 4500 >=20000;(( 4500 -10000)/1000*2.5)+17.5; SI ( 4500 <=30000;(( 4500 -20000)/1000*3)+ 42.5; SI ( 4500 <=50000;(( 4500 -30000)/1000*3.5)+72.5; SI ( 4500 <=75000;(( 4500 -50000)/1000*3.75)+142.5; SI ( 4500 <=100000;(( 4500 -75000)/1000*4)+236.25; SI ( 4500 <=150000;(( 4500 -100000)/1000*5)+336.25;(( 4500 -150000)/1000*5.25)+586.25))"), exito, True))
            'Resultado = CDbl(eval.Evaluate(Trim("SI(3=1 OR (3=2 AND 50=50);SI( Monto <=5000; Monto/1000 * 1.5; SI(4500 <=10000;((4500 - 5000)/1000 * 2) + 7.5; SI(4500 >=20000;((4500 -10000)/1000 * 2.5) + 17.5 ; SI(4500 <=30000;((4500 -20000)/1000 * 3) + 42.5; SI(4500 <=50000;((4500 -30000)/1000 * 3.5) + 72.5; SI(4500 <=75000;((4500 -50000)/1000 * 3.75) + 142.5); SI(4500 <=100000;((4500 -75000)/1000 * 4) + 236.25; SI(4500 <=150000;((4500 -100000)/1000 * 5) + 336.25; (( 4500 -150000)/1000 * 5.25) + 586.25))))))));1500)"), exito, True))
            
            'Resultado = CDbl(eval.Evaluate(Trim("(( 4500 -150000)/1000 * 5.25) + 586.25"), exito, True))
            'Resultado = CDbl(eval.Evaluate(Trim(""), exito, True))
            'Resultado = CDbl(eval.Evaluate(Trim(""), exito, True))
            'Resultado = CDbl(eval.Evaluate(Trim(""), exito, True))
            'Resultado = CDbl(eval.Evaluate(Trim(""), exito, True))
            'Resultado = CDbl(eval.Evaluate(Trim(""), exito, True))
            'Resultado = CDbl(eval.Evaluate(Trim(""), exito, True))
            'Resultado = CDbl(eval.Evaluate(Trim(""), exito, True))
            'Resultado = CDbl(eval.Evaluate(Trim(""), exito, True))
            
            
            Resultado = CDbl(eval.Evaluate(Trim(Arr_conceptos(Concepto_Actual).Forexpresion), exito, True))
            
            If CBool(USA_DEBUG) Then
                TiempoFinalFormula = GetTickCount
                Flog.writeline Espacios(Tabulador * 1) & "Formula: " & Arr_conceptos(Concepto_Actual).fornro & " (" & Arr_conceptos(Concepto_Actual).Fordabr & ") - Tiempo: " & (TiempoFinalFormula - TiempoInicialFormula)
            End If
            
        Else 'es una formula codificada en vb (un procedimiento que la resuleve)
            'el tema es como resulevo a que formula llamar
            TiempoInicialFormula = GetTickCount
            If Arr_conceptos(Concepto_Actual).Fortipo = 2 Then 'No configurable
                Resultado = EjecutarFormulaNoConfigurable(Trim(Arr_conceptos(Concepto_Actual).Forprog))
            Else ' de sistema
                Resultado = EjecutarFormulaDeSistema(Trim(Arr_conceptos(Concepto_Actual).Forprog))
            End If
            If CBool(USA_DEBUG) Then
                TiempoFinalFormula = GetTickCount
                Flog.writeline Espacios(Tabulador * 1) & "Formula: " & Arr_conceptos(Concepto_Actual).Forprog & " - Tiempo: " & (TiempoFinalFormula - TiempoInicialFormula)
            End If
        End If
        
        If (Not exito) Then
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 1) & "La formula no devolvio un OK"
            End If
        
            If HACE_TRAZA Then
                Call InsertarTraza(NroCab, Arr_conceptos(Concepto_Actual).ConcNro, 9999, "La formula no devolvio un OK", 0)
            End If
            ' SIGUIENTE CONCEPTO
            GoTo SiguienteConcepto
        End If
            
        
        ' Redondeo por presicion para trabajar con dos decimales
        ' Hay que sacarlo del concepto
        Resultado = Round(Resultado, Arr_conceptos(Concepto_Actual).Conccantdec)
            
        If (Resultado = 0) Then
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 1) & "La formula devolvio CERO"
            End If
        
            If HACE_TRAZA Then
                Call InsertarTraza(NroCab, Arr_conceptos(Concepto_Actual).ConcNro, 9999, "La formula devolvio CERO", 0)
            End If
            ' Es por el tema de ganancias
            ' SIGUIENTE CONCEPTO
            GoTo SiguienteConcepto
        End If
     
       ' Si es grossing debe saltar
       If Not Termino_gross Then
          If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 1) & "Nueva Iteracion de Grossing"
            End If
          Exit Sub ' Sale sin grabar y pasa al mismo empleado con la proxima iteracion
       End If
       
'        ' Redondeo por presicion para trabajar con dos decimales
'        ' Hay que sacarlo del concepto
'        Resultado = Round(Resultado, Arr_Conceptos(Concepto_Actual).conccantdec)
        
        ' Guarda la cantidad de dias de AMPO para Proporcinar
        'FGZ - 04/06/2012 -----------------------------------------------------------------------------
        'StrSql = "SELECT * FROM ampocontpa WHERE concnro = " & Arr_conceptos(Concepto_Actual).ConcNro
        StrSql = "SELECT signo FROM ampocontpa WHERE concnro = " & Arr_conceptos(Concepto_Actual).ConcNro & _
                 " AND tpanro = " & tpa_parametro
        OpenRecordset StrSql, rs_AmpoConTpa
        
        If Not rs_AmpoConTpa.EOF Then
            If CBool(rs_AmpoConTpa!signo) Then
                Select Case Arr_conceptos(Concepto_Actual).tconnro
                Case 1
                    Cant_Ampo_Proporcionar_1 = Cant_Ampo_Proporcionar_1 + Parametro
                    If Parametro <> 0 Then
                        Sumo_Cant_Ampo_Prop_1 = True
                    End If
                Case 2
                    Cant_Ampo_Proporcionar_2 = Cant_Ampo_Proporcionar_2 + Parametro
                    If Parametro <> 0 Then
                        Sumo_Cant_Ampo_Prop_2 = True
                    End If
                Case 3
                    Cant_Ampo_Proporcionar_3 = Cant_Ampo_Proporcionar_3 + Parametro
                    If Parametro <> 0 Then
                        Sumo_Cant_Ampo_Prop_3 = True
                    End If
                Case 4
                    Cant_Ampo_Proporcionar_4 = Cant_Ampo_Proporcionar_4 + Parametro
                    If Parametro <> 0 Then
                        Sumo_Cant_Ampo_Prop_4 = True
                    End If
                Case 5
                    Cant_Ampo_Proporcionar_5 = Cant_Ampo_Proporcionar_5 + Parametro
                    If Parametro <> 0 Then
                        Sumo_Cant_Ampo_Prop_5 = True
                    End If
                End Select
            Else
                Select Case Arr_conceptos(Concepto_Actual).tconnro
                Case 1
                    Cant_Ampo_Proporcionar_1 = Cant_Ampo_Proporcionar_1 - Parametro
                    If Parametro <> 0 Then
                        Sumo_Cant_Ampo_Prop_1 = True
                    End If
                Case 2
                    Cant_Ampo_Proporcionar_2 = Cant_Ampo_Proporcionar_2 - Parametro
                    If Parametro <> 0 Then
                        Sumo_Cant_Ampo_Prop_2 = True
                    End If
                Case 3
                    Cant_Ampo_Proporcionar_3 = Cant_Ampo_Proporcionar_3 - Parametro
                    If Parametro <> 0 Then
                        Sumo_Cant_Ampo_Prop_3 = True
                    End If
                Case 4
                    Cant_Ampo_Proporcionar_4 = Cant_Ampo_Proporcionar_4 - Parametro
                    If Parametro <> 0 Then
                        Sumo_Cant_Ampo_Prop_4 = True
                    End If
                Case 5
                    Cant_Ampo_Proporcionar_5 = Cant_Ampo_Proporcionar_5 - Parametro
                    If Parametro <> 0 Then
                        Sumo_Cant_Ampo_Prop_5 = True
                    End If
                End Select
            End If
            
            
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 1) & "Resultado de la formula " & Resultado
            End If
            If HACE_TRAZA Then
                Call InsertarTraza(NroCab, Arr_conceptos(Concepto_Actual).ConcNro, 9999, "El Resultado de la formula ", Resultado)
            End If
            
        End If
        
        'saque la linea
        '"," & ConvFecha(fecha_retro) & _
        ' que se insertaba en dlifec

        ' inserto el concepto en detliq
        StrSql = "INSERT INTO detliq (" & _
                 "cliqnro,concnro,dlimonto,dlicant,dliretro,fornro,tconnro,ajustado,dliqdesde,dliqhasta,dlitexto" & _
                 ") VALUES (" & buliq_cabliq!cliqnro & _
                 "," & Arr_conceptos(Concepto_Actual).ConcNro & _
                 "," & Resultado & _
                 "," & Parametro & _
                 "," & Concepto_Retroactivo & _
                 "," & Arr_conceptos(Concepto_Actual).fornro & _
                 "," & Arr_conceptos(Concepto_Actual).tconnro & _
                 ", 0" & _
                 "," & concepto_pliqdesde & _
                 "," & concepto_pliqhasta & _
                 ",'" & IIf(EsNulo(Arr_conceptos(Concepto_Actual).Conctexto), "Concepto xxx ", Arr_conceptos(Concepto_Actual).Conctexto) & _
                 "' )"
        objConn.Execute StrSql, , adExecuteNoRecords
                
        'FGZ - 10/02/2004
        ' inserto en el cache de detliq
        ' El monto
        Call objCache_detliq_Monto.Insertar_Simbolo(CStr(Arr_conceptos(Concepto_Actual).ConcNro), Resultado)
        ' La cantidad
        Call objCache_detliq_Cantidad.Insertar_Simbolo(CStr(Arr_conceptos(Concepto_Actual).ConcNro), Parametro)
        'FGZ - 10/02/2004
                
        If CBool(USA_DEBUG) Then
            Texto = "Resultado de la formula (" & Arr_conceptos(Concepto_Actual).fornro & ") "
            Flog.writeline Espacios(Tabulador * 1) & Texto & Resultado
        End If
        If HACE_TRAZA Then
            Texto = "Resultado de la formula (" & Arr_conceptos(Concepto_Actual).fornro & ") "
            Call InsertarTraza(NroCab, Arr_conceptos(Concepto_Actual).ConcNro, 9999, Texto, Resultado)
        End If
                
    End If 'Machetazo
    
    
        ' Los Conceptos Retroactivos no deben sumar en Acumuladores Retroactivos
        '18/06/2010 - La condicion para que sea retro es que la tabla temp de retro sea <> vacia
        '             lo controla Liqpro15
        'If CBool(Concepto_Retroactivo) Then
            ' Hay que sumar el resultado a los Meses retroactivos
            'FGZ - 27/12/2004
            YA_Ajustado = False
             Call Liqpro15(buliq_cabliq!Empleado, concepto_pliqdesde, concepto_pliqhasta, Resultado, Buliq_Concepto(Concepto_Actual).ConcNro, NroCab, YA_Ajustado)
                
            If YA_Ajustado Then
                ' Actualizo Ajustado
                
                'saque del where " AND dlifec =" & ConvFecha(fecha_retro)
                
                StrSql = "UPDATE detliq set ajustado = -1, dliretro = -1 WHERE " & _
                         " cliqnro =" & buliq_cabliq!cliqnro & _
                         " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro & _
                         " AND dliqdesde =" & concepto_pliqdesde & _
                         " AND dliqhasta =" & concepto_pliqhasta
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
        'End If
        'Else ' Concepto no Retroactivo al resto de los acumuladores
        
            ' Actualizar los acumuladores del concepto: acu_liq
            ' FGZ - 23/08/2004
            ' ahora uso el vector de con_acum cargado una sola vez para todo el proceso
            Aux_Acu_Actual = 0
            Aux_Acu_Actual = Siguiente_Con_Acum(Aux_Acu_Actual)
            Do While Not Aux_Acu_Actual = 0
                'FGZ - 08/10/2004
                'Si el concepto es retroactivo y el acumulador no es retro entonces suma (sino quedan duplicados)
                If Not CBool(Concepto_Retroactivo) Or (CBool(Concepto_Retroactivo) And Not CBool(Arr_Acumulador(Aux_Acu_Actual).acuretro)) Then
                
                    A_Proporcionar = False
                    Resu_Proporcionado = 0
                                       
                    'FGZ - 19/05/2011 --------------------------------------------------------------
                    ' Se sacó esto porque ya no se usa
                    
                    '' Verificar si se exceptua para el contrato del empleado, si es así sigue al proximo
                    ''desde acá
                    'StrSql = "SELECT * FROM tc_acu_excep WHERE acunro = " & Arr_Acumulador(Aux_Acu_Actual).acuNro & _
                    '         " AND tcnro = " & Contrato_proporcion
                    'OpenRecordset StrSql, rs_tc_acu
   
                    'If Not rs_tc_acu.EOF Then
                    '    ' Verificar si el contrato vence en el mes de liquidacion
                    '    If Proporcion_Contrato = 1 Then
                    '        ' El contrato exceptua pero no termina este mes - No suma en el acumlador
                    '        GoTo siguienteCon_Acum
                    '    Else
                    '        ' El contrato Exceptua y Finaliza este mes
                    '        Resu_Proporcionado = Resultado * Proporcion_Contrato
                    '        A_Proporcionar = True
                    '    End If
                    'End If
                    ' hasta acá
                   'FGZ - 19/05/2011 --------------------------------------------------------------
                   
                    If objCache_Acu_Liq_Monto.EsSimboloDefinido(CStr(Arr_Acumulador(Aux_Acu_Actual).acuNro)) Then
                         'Acumulo
                         Aux_Acu_Monto = objCache_Acu_Liq_Monto.Valor(CStr(Arr_Acumulador(Aux_Acu_Actual).acuNro))
                         Aux_Acu_Cant = objCache_Acu_Liq_Cantidad.Valor(CStr(Arr_Acumulador(Aux_Acu_Actual).acuNro))
                         
                         If CBool(A_Proporcionar) Then
                             Aux_Nuevo_Monto = Resu_Proporcionado
                         Else
                             Aux_Nuevo_Monto = Resultado
                         End If
                         Aux_Nuevo_Cant = Parametro
                         
                         'borro y vuelvo a insertar
                         Call objCache_Acu_Liq_Monto.Borrar_Simbolo(CStr(Arr_Acumulador(Aux_Acu_Actual).acuNro))
                         Call objCache_Acu_Liq_MontoReal.Borrar_Simbolo(CStr(Arr_Acumulador(Aux_Acu_Actual).acuNro))
                         Call objCache_Acu_Liq_Cantidad.Borrar_Simbolo(CStr(Arr_Acumulador(Aux_Acu_Actual).acuNro))
                         
                         Call objCache_Acu_Liq_Monto.Insertar_Simbolo(CStr(Arr_Acumulador(Aux_Acu_Actual).acuNro), (Aux_Acu_Monto + Aux_Nuevo_Monto))
                         Call objCache_Acu_Liq_MontoReal.Insertar_Simbolo(CStr(Arr_Acumulador(Aux_Acu_Actual).acuNro), (Aux_Acu_Monto + Aux_Nuevo_Monto))
                         Call objCache_Acu_Liq_Cantidad.Insertar_Simbolo(CStr(Arr_Acumulador(Aux_Acu_Actual).acuNro), (Aux_Acu_Cant + Aux_Nuevo_Cant))

                     Else
                         If CBool(A_Proporcionar) Then
                             Aux_Nuevo_Monto = Resu_Proporcionado
                         Else
                             Aux_Nuevo_Monto = Resultado
                         End If
                         Aux_Nuevo_Cant = Parametro
                     
                         ' inserto en el cache
                         Call objCache_Acu_Liq_Monto.Insertar_Simbolo(CStr(Arr_Acumulador(Aux_Acu_Actual).acuNro), Aux_Nuevo_Monto)
                         Call objCache_Acu_Liq_Cantidad.Insertar_Simbolo(CStr(Arr_Acumulador(Aux_Acu_Actual).acuNro), Aux_Nuevo_Cant)
                     End If
    
            
                    If CBool(USA_DEBUG) Then
                        Texto = "Suma al acumulador: " + Format(Arr_Acumulador(Aux_Acu_Actual).acuNro, "000") + "-" + Arr_Acumulador(Aux_Acu_Actual).acudesabr & " "
                        Flog.writeline Espacios(Tabulador * 2) & Texto & Resultado
                    End If
               
              
                    If CBool(Arr_Acumulador(Aux_Acu_Actual).acutopea) Then
                        If CBool(USA_DEBUG) Then
                            Flog.writeline Espacios(Tabulador * 2) & "El acumulador topea"
                        End If
                        ' Sumarizar en el temporario de Imponibles del proceso
                        'FGZ - 04/06/2012 -----------------------------------
                        'StrSql = "SELECT * FROM " & TTempWF_impproarg & " WHERE acunro = " & Arr_Acumulador(Aux_Acu_Actual).acuNro
                        StrSql = "SELECT acunro FROM " & TTempWF_impproarg & " WHERE acunro = " & Arr_Acumulador(Aux_Acu_Actual).acuNro & _
                                 " AND tconnro =" & Arr_conceptos(Concepto_Actual).tconnro
                        OpenRecordset StrSql, rs_WF_impproarg
                                        
                        If rs_WF_impproarg.EOF Then
                            Call insertar_wf_impproarg(Arr_Acumulador(Aux_Acu_Actual).acuNro, CBool(Arr_Acumulador(Aux_Acu_Actual).acuimponible), CBool(Arr_Acumulador(Aux_Acu_Actual).acuimpcont), Arr_conceptos(Concepto_Actual).tconnro)
                        End If
                        
                        ' Actualiza
                            NoActualiza = False
                            Select Case Arr_conceptos(Concepto_Actual).tconnro
                            Case 1
                                StrSql = "UPDATE " & TTempWF_impproarg & " SET " & _
                                         " ipacant = " & Cant_Ampo_Proporcionar_1 & _
                                         ", ipamonto = ipamonto + "
                            Case 2
                                StrSql = "UPDATE " & TTempWF_impproarg & " SET " & _
                                         " ipacant = " & Cant_Ampo_Proporcionar_2 & _
                                         ", ipamonto = ipamonto + "
                            Case 3
                                StrSql = "UPDATE " & TTempWF_impproarg & " SET " & _
                                         " ipacant = " & Cant_Ampo_Proporcionar_3 & _
                                         ", ipamonto = ipamonto + "
                            Case 4
                                StrSql = "UPDATE " & TTempWF_impproarg & " SET " & _
                                         " ipacant = " & Cant_Ampo_Proporcionar_4 & _
                                         ", ipamonto = ipamonto + "
                            Case 5
                                StrSql = "UPDATE " & TTempWF_impproarg & " SET " & _
                                         " ipacant = " & Cant_Ampo_Proporcionar_5 & _
                                         ", ipamonto = ipamonto + "
                            Case Else:
                                NoActualiza = True
                                'If CBool(USA_DEBUG) Then
                                    Flog.writeline Espacios(Tabulador * 3) & "El Tipo de concepto no permite que sume a un Imponible. Error de Configuracion."
                                'End If
                            End Select
                            
                            If Not NoActualiza Then
                                If CBool(A_Proporcionar) Then
                                    StrSql = StrSql & Resu_Proporcionado
                                Else
                                    StrSql = StrSql & Resultado
                                End If
                                StrSql = StrSql & _
                                " WHERE acunro = " & Arr_Acumulador(Aux_Acu_Actual).acuNro & " AND tconnro = " & Arr_conceptos(Concepto_Actual).tconnro
                                objConn.Execute StrSql, , adExecuteNoRecords
                            End If
                        If CBool(USA_DEBUG) Then
                            'FGZ - 04/06/2012 ---------------------
                            'If rs_WF_impproarg.State = adStateOpen Then rs_WF_impproarg.Close
                            'StrSql = "SELECT * FROM " & TTempWF_impproarg & " WHERE acunro = " & Arr_Acumulador(Aux_Acu_Actual).acuNro
                            StrSql = "SELECT ipacant FROM " & TTempWF_impproarg & " WHERE acunro = " & Arr_Acumulador(Aux_Acu_Actual).acuNro & _
                                     " AND tconnro =" & Arr_conceptos(Concepto_Actual).tconnro
                            OpenRecordset StrSql, rs_WF_impproarg
                                        
                            If Not rs_WF_impproarg.EOF Then
                                Select Case Arr_conceptos(Concepto_Actual).tconnro
                                Case 1:
                                    Texto = "Imponible Monto: " + Format(Arr_Acumulador(Aux_Acu_Actual).acuNro, "000") + "-" + Arr_Acumulador(Aux_Acu_Actual).acudesabr & " "
                                    Flog.writeline Espacios(Tabulador * 3) & Texto & Resultado 'rs_WF_impproarg!ipamonto
                                    Texto = "Imponible Cantidad: " + Format(Arr_Acumulador(Aux_Acu_Actual).acuNro, "000") + "- proporcionado" + Arr_Acumulador(Aux_Acu_Actual).acudesabr & " "
                                    Flog.writeline Espacios(Tabulador * 3) & Texto & rs_WF_impproarg!ipacant
                                Case 2:
                                    Texto = "Imponible Monto: " + Format(Arr_Acumulador(Aux_Acu_Actual).acuNro, "000") + "-" + Arr_Acumulador(Aux_Acu_Actual).acudesabr & " "
                                    Flog.writeline Espacios(Tabulador * 3) & Texto & Resultado 'rs_WF_impproarg!ipamonto
                                    Texto = "Imponible Cantidad: " + Format(Arr_Acumulador(Aux_Acu_Actual).acuNro, "000") + "- proporcionado" + Arr_Acumulador(Aux_Acu_Actual).acudesabr & " "
                                    Flog.writeline Espacios(Tabulador * 3) & Texto & rs_WF_impproarg!ipacant
                                Case 3:
                                    Texto = "Imponible Monto: " + Format(Arr_Acumulador(Aux_Acu_Actual).acuNro, "000") + "-" + Arr_Acumulador(Aux_Acu_Actual).acudesabr & " "
                                    Flog.writeline Espacios(Tabulador * 3) & Texto & Resultado 'rs_WF_impproarg!ipamonto
                                    Texto = "Imponible Cantidad: " + Format(Arr_Acumulador(Aux_Acu_Actual).acuNro, "000") + "- proporcionado" + Arr_Acumulador(Aux_Acu_Actual).acudesabr & " "
                                    Flog.writeline Espacios(Tabulador * 3) & Texto & rs_WF_impproarg!ipacant
                                Case 4:
                                    Texto = "Imponible Monto: " + Format(Arr_Acumulador(Aux_Acu_Actual).acuNro, "000") + "-" + Arr_Acumulador(Aux_Acu_Actual).acudesabr & " "
                                    Flog.writeline Espacios(Tabulador * 3) & Texto & Resultado 'rs_WF_impproarg!ipamonto
                                    Texto = "Imponible Cantidad: " + Format(Arr_Acumulador(Aux_Acu_Actual).acuNro, "000") + "- proporcionado" + Arr_Acumulador(Aux_Acu_Actual).acudesabr & " "
                                    Flog.writeline Espacios(Tabulador * 3) & Texto & rs_WF_impproarg!ipacant
                                Case 5:
                                    Texto = "Imponible Monto: " + Format(Arr_Acumulador(Aux_Acu_Actual).acuNro, "000") + "-" + Arr_Acumulador(Aux_Acu_Actual).acudesabr & " "
                                    Flog.writeline Espacios(Tabulador * 3) & Texto & Resultado 'rs_WF_impproarg!ipamonto
                                    Texto = "Imponible Cantidad: " + Format(Arr_Acumulador(Aux_Acu_Actual).acuNro, "000") + "- proporcionado" + Arr_Acumulador(Aux_Acu_Actual).acudesabr & " "
                                    Flog.writeline Espacios(Tabulador * 3) & Texto & rs_WF_impproarg!ipacant
                                Case Else:
                                    Flog.writeline Espacios(Tabulador * 3) & "El Tipo de concepto no permite que sume a un Imponible. Error de Configuracion."
                                End Select
                            End If
                        End If
                        
                    End If 'If CBool(rs_Con_acum!acutopea) Then
                'FGZ - 08/10/2004
                End If
siguienteCon_Acum:
                Aux_Acu_Actual = Siguiente_Con_Acum(Aux_Acu_Actual)
            Loop
        
        'End If 'If CBool(concepto_retroactivo) Then
        
        
SiguienteConcepto:
        If CBool(USA_DEBUG) Then
            TiempoFinalConcepto = GetTickCount
            Flog.writeline Espacios(Tabulador * 1) & "Tiempo para el concepto: " & (TiempoFinalConcepto - TiempoInicialConcepto)
        End If
        
        If Usa_grossing Then
            Progreso = Progreso + IncPorc / (MaxIteraGross * CConceptosAProc)
        Else
            Progreso = Progreso + IncPorc
        End If
        TiempoAcumulado = GetTickCount
        
        'FGZ - 20/05/2011 ----------------------------------------------------------------
        ' se sacó la actualizacion del progreso de aca y se dejó una vez por cada empleado
        'StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
        '         ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
        '         "' WHERE bpronro = " & NroProcesoBatch
        'objconnProgreso.Execute StrSql, , adExecuteNoRecords
        'FGZ - 20/05/2011 ----------------------------------------------------------------
        
        Concepto_Actual = Concepto_Actual + 1
        'rs_Conceptos.MoveNext
    Loop
    
' -------------------------------------------------------------------------------
' FGZ - 06/02/2004
' Actualizacion de todos los acu_liq generados para este empleado
' -------------------------------------------------------------------------------
'StrSql = "SELECT * FROM acumulador "
'OpenRecordset StrSql, rs_Acumulador

'Ahora uso un vector global de acumuladores
Acumulador_Actual = 0
Acumulador_Actual = Siguiente_Acumulador
Do While Acumulador_Actual <> 0
    Acum = CStr(Arr_Acumulador(Acumulador_Actual).acuNro)
    If objCache_Acu_Liq_Monto.EsSimboloDefinido(CStr(Acum)) Then
        Aux_Acu_Monto = objCache_Acu_Liq_Monto.Valor(CStr(Acum))
        Aux_Acu_MontoReal = objCache_Acu_Liq_MontoReal.Valor(CStr(Acum))
        Aux_Acu_Cant = objCache_Acu_Liq_Cantidad.Valor(CStr(Acum))
    
        'FGZ - 10/09/2004
        If CBool(Arr_Acumulador(Acumulador_Actual).acunoneg) Then
            If Aux_Acu_Monto < 0 Then
                HayAcuNoNeg = True
                'FGZ - 29/04/2008 - Se agregó este mensaje de log para que ayude a encontrar cual es el acumulador negativo
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 2) & "-------------------------------------------------------------------------------"
                    Flog.writeline Espacios(Tabulador * 2) & " ACUMULADORES NEGATIVO ------> : " & Acum & " monto: " & Aux_Acu_Monto
                    Flog.writeline Espacios(Tabulador * 2) & "-------------------------------------------------------------------------------"
                End If
                
            End If
        End If
        ' Crear acu_liq
        'StrSql = "SELECT * FROM acu_liq WHERE acunro = " & Acum & _
        '         " AND cliqnro = " & buliq_cabliq!cliqnro
        StrSql = "SELECT cliqnro FROM acu_liq WHERE acunro = " & Acum & _
                 " AND cliqnro = " & buliq_cabliq!cliqnro
        OpenRecordset StrSql, rs_AcuLiq
        
        If rs_AcuLiq.EOF Then
            ' Inserto
            StrSql = "INSERT INTO acu_liq (" & _
                     "acunro,cliqnro,almonto,almontoreal,alcant" & _
                     ") VALUES (" & Acum & _
                     "," & buliq_cabliq!cliqnro & _
                     ", " & Aux_Acu_Monto & _
                     ", " & Aux_Acu_MontoReal & _
                     ", " & Aux_Acu_Cant & _
                     " )"
        Else
            StrSql = "UPDATE acu_liq SET " & _
                     " alcant = alcant + " & Aux_Acu_Cant & _
                     ", almontoreal = almonto + " & Aux_Acu_Monto & _
                    ", almonto = almonto + " & Aux_Acu_Monto & _
                    " WHERE acunro = " & Acum & _
                    " AND cliqnro = " & buliq_cabliq!cliqnro
        End If
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
    
    Acumulador_Actual = Siguiente_Acumulador
    'rs_Acumulador.MoveNext
Loop

' -------------------------------------------------------------------------------
'SUMAR ACUMULADORES DEL PROCESO A LOS MENSUALES: acu_mes e impmesarg
' -------------------------------------------------------------------------------
'FGZ - 07/06/2012 ------------------
StrSql = "SELECT acunro,almonto,alcant,almontoreal FROM acu_liq WHERE cliqnro = " & buliq_cabliq!cliqnro
OpenRecordset StrSql, rs_AcuLiq

Do While Not rs_AcuLiq.EOF
    If CBool(Arr_Acumulador(rs_AcuLiq!acuNro).acumes) Then
        'el acumulador es mensual
        'FGZ - 07/06/2012 ------------------
        StrSql = "SELECT tconnro,ipamonto,ipacant FROM impproarg WHERE acunro = " & Arr_Acumulador(rs_AcuLiq!acuNro).acuNro & _
                 " AND cliqnro = " & buliq_cabliq!cliqnro
        OpenRecordset StrSql, rs_ImpPro
        
            'FGZ - 07/06/2012 ------------------
            StrSql = "SELECT acunro,ammonto,amcant FROM acu_mes WHERE acunro = " & Arr_Acumulador(rs_AcuLiq!acuNro).acuNro & _
                     " AND amanio = " & buliq_periodo!pliqanio & _
                     " AND ternro = " & buliq_empleado!Ternro & _
                     " AND ammes = " & buliq_periodo!pliqmes
            OpenRecordset StrSql, rs_acumes
                                    
            If rs_acumes.EOF Then
                ' Inserto
                StrSql = "INSERT INTO acu_mes (acunro,amanio, ternro, ammes,ammonto,ammontoreal, amcant " & _
                         ") VALUES (" & _
                         Arr_Acumulador(rs_AcuLiq!acuNro).acuNro & _
                         "," & buliq_periodo!pliqanio & _
                         "," & buliq_empleado!Ternro & _
                         "," & buliq_periodo!pliqmes & _
                         ",0,0,0" & _
                         " )"
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
                        
            ' Actualizo
            StrSql = " UPDATE acu_mes SET ammonto = ammonto + " & rs_AcuLiq!almonto & _
                     ", ammontoreal = ammontoreal + " & rs_AcuLiq!almontoreal & _
                     ", amcant = amcant + " & rs_AcuLiq!alcant & _
                     " WHERE acunro = " & Arr_Acumulador(rs_AcuLiq!acuNro).acuNro & _
                     " AND ternro = " & buliq_cabliq!Empleado & _
                     " AND amanio = " & buliq_periodo!pliqanio & _
                     " AND ammes = " & buliq_periodo!pliqmes
            objConn.Execute StrSql, , adExecuteNoRecords
                    
            If CBool(USA_DEBUG) Then
                Texto = "Mensual " + Format(Arr_Acumulador(rs_AcuLiq!acuNro).acuNro, "000") + "-" + Arr_Acumulador(rs_AcuLiq!acuNro).acudesabr & " Monto: "
                If rs_acumes.EOF Then
                    Flog.writeline Espacios(Tabulador * 1) & Texto & rs_AcuLiq!almonto
                Else
                    If EsNulo(rs_acumes!ammonto) Then
                        Flog.writeline Espacios(Tabulador * 1) & Texto & rs_AcuLiq!almonto
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & Texto & CStr(CDbl(rs_acumes!ammonto + rs_AcuLiq!almonto))
                    End If
                End If
                Texto = "Mensual " + Format(Arr_Acumulador(rs_AcuLiq!acuNro).acuNro, "000") + "- proporcionado" + Arr_Acumulador(rs_AcuLiq!acuNro).acudesabr & " Cantidad: "
                If rs_acumes.EOF Then
                    Flog.writeline Espacios(Tabulador * 1) & Texto & rs_AcuLiq!alcant
                Else
                    If EsNulo(rs_acumes!amcant) Then
                        Flog.writeline Espacios(Tabulador * 1) & Texto & rs_AcuLiq!alcant
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & Texto & CStr(CDbl(rs_acumes!amcant + rs_AcuLiq!alcant))
                    End If
                End If
            End If
            
                                
            If Not rs_ImpPro.EOF Then
                If CBool(Arr_Acumulador(rs_AcuLiq!acuNro).acuimponible) Or CBool(Arr_Acumulador(rs_AcuLiq!acuNro).acuimpcont) Then
                    'Suma en los imponibles del proceso y del mes de acuerdo al tipo de concepto
                    ' No necesita pregunta dentro de lo mensual, porque un acumulador no puede ser imponible sin ser mensual
                    Do While Not rs_ImpPro.EOF
                        'FGZ - 07/06/2012 ------------------
                        StrSql = "SELECT acunro FROM impmesarg " & _
                                 " WHERE acunro = " & Arr_Acumulador(rs_AcuLiq!acuNro).acuNro & _
                                 " AND imaanio = " & buliq_periodo!pliqanio & _
                                 " AND imames = " & buliq_periodo!pliqmes & _
                                 " AND ternro = " & buliq_empleado!Ternro & _
                                 " AND tconnro = " & rs_ImpPro!tconnro
                        OpenRecordset StrSql, rs_ImpMesArg
        
                        If rs_ImpMesArg.EOF Then
                            ' Inserto uno por cada tipo de concepto (tconnro)
                            StrSql = "INSERT INTO impmesarg (acunro,imaanio, imames, ternro, tconnro, imamonto, imacant" & _
                                     ") VALUES (" & _
                                     Arr_Acumulador(rs_AcuLiq!acuNro).acuNro & _
                                     "," & buliq_periodo!pliqanio & _
                                     "," & buliq_periodo!pliqmes & _
                                     "," & buliq_empleado!Ternro & _
                                     "," & rs_ImpPro!tconnro & _
                                     "," & rs_ImpPro!ipamonto & _
                                     "," & rs_ImpPro!ipacant & _
                                     " )"
                            objConn.Execute StrSql, , adExecuteNoRecords
                        Else
                            'Actualizo
                            StrSql = " UPDATE impmesarg SET imamonto = imamonto + " & rs_ImpPro!ipamonto & _
                                     ", imacant = imacant + " & rs_ImpPro!ipacant & _
                                     " WHERE acunro = " & Arr_Acumulador(rs_AcuLiq!acuNro).acuNro & _
                                     " AND imaanio = " & buliq_periodo!pliqanio & _
                                     " AND imames = " & buliq_periodo!pliqmes & _
                                     " AND ternro = " & buliq_empleado!Ternro & _
                                     " AND tconnro = " & rs_ImpPro!tconnro
                            objConn.Execute StrSql, , adExecuteNoRecords
                        End If
                    
                        If CBool(USA_DEBUG) Then
                            Texto = "Imponible " + CStr(rs_ImpPro!tconnro) + ". " + Format(Arr_Acumulador(rs_AcuLiq!acuNro).acuNro, "000") + "-" + Arr_Acumulador(rs_AcuLiq!acuNro).acudesabr & " Monto topeado: "
                            Flog.writeline Espacios(Tabulador * 2) & Texto & rs_ImpPro!ipamonto
                            Texto = "Imponible 1. " + Format(Arr_Acumulador(rs_AcuLiq!acuNro).acuNro, "000") + "-" + Arr_Acumulador(rs_AcuLiq!acuNro).acudesabr & " Cantidad: "
                            Flog.writeline Espacios(Tabulador * 2) & Texto & rs_ImpPro!ipacant
                        End If
'                        If HACE_TRAZA Then
'                            Texto = "Imponible " + CStr(rs_ImpPro!tconnro) + ". Monto topeado: " + Format(Arr_Acumulador(rs_AcuLiq!acuNro).acuNro, "000") + "-" + Arr_Acumulador(rs_AcuLiq!acuNro).acudesabr
'                            Call InsertarTraza(NroCab, 0, 99997, Texto, rs_ImpPro!ipamonto)
'                            Texto = "Imponible 1. Cantidad: " + Format(Arr_Acumulador(rs_AcuLiq!acuNro).acuNro, "000") + "-" + Arr_Acumulador(rs_AcuLiq!acuNro).acudesabr
'                            Call InsertarTraza(NroCab, 0, 99997, Texto, rs_ImpPro!ipacant)
'                        End If
                            
                        rs_ImpPro.MoveNext
                    Loop
                End If
            End If ' If CBool(rs_acumulador!acuimponible) Or CBool(rs_acumulador!acuimpcont) Then
    End If ' If Not rs_acumulador.EOF Then

    rs_AcuLiq.MoveNext
Loop

'FGZ - 14/12/2013 ---------------------------------------------------------------
' -------------------------------------------------------------------------------
'Ajusto montos y porcenyajes de cada concepto con distribucion
' -------------------------------------------------------------------------------
Call AjustarDistribucionNov
'FGZ - 14/12/2013 ---------------------------------------------------------------

Exit Sub

'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
CE:
    'MyRollbackTransliq
    HuboError = True
    EmpleadoSinError = False
    'FGZ - 28/12/2016 ---------------------
    ' Los msg de error se deben mostrar siempre
    'If CBool(USA_DEBUG) Then
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
        Flog.writeline Espacios(Tabulador * 0) & " Empleado abortado: " & buliq_empleado!Empleg
        Flog.writeline Espacios(Tabulador * 0) & " Error: " & Err.Description
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "Ultimo SQL Ejecutado: " & StrSql
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
        Flog.writeline
    'End If
    'FGZ - 28/12/2016 ---------------------
    MyRollbackTransliq
    
    'Actualizo el progreso
    MyBeginTransLiq
        Progreso = Progreso + IncPorc
        TiempoAcumulado = GetTickCount
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
                 ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
                 "' WHERE bpronro = " & NroProcesoBatch
        objConn.Execute StrSql, , adExecuteNoRecords
    MyCommitTransLiq
    
End Sub

Public Sub CargarTablaParametros_Nueva()
' ---------------------------------------------------------------------------------------------
' Descripcion: carga la tabla de simbolos con los parametros en wf_tpa.
' Autor      : FGZ
' Fecha      : 24/05/2011
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim symbols As New CSymbolTable
Dim rs_wf_tpa As New ADODB.Recordset
Dim NParametro As String
Dim Valor As Double
Dim tipo As String
Dim I As Long
'Dim LI As Long
'Dim LS As Long

    Set eval.m_SymbolTable = symbols
     
'    'Aca deberia cargar todas las funciones validas
'    StrSql = "SELECT * FROM funformula "
'    OpenRecordset StrSql, rs_FunFormulas
    'FGZ - 20/08/2004
    If rs_FunFormulas.State = adStateOpen Then
        rs_FunFormulas.MoveFirst
    
        Do While Not rs_FunFormulas.EOF
            NParametro = Trim(rs_FunFormulas!fundesabr)
            tipo = rs_FunFormulas!funprograma
            eval.m_SymbolTable.Add NParametro, tipo
            
            rs_FunFormulas.MoveNext
        Loop
    End If

    'Busco el primer y ultimo parametro de la formula del concepto actual
    'LI_wf_tpa = BuscarPrimer_For_Tpa(Arr_conceptos(Concepto_Actual).fornro)
    'LS_wf_tpa = BuscarUltimo_For_Tpa(Arr_conceptos(Concepto_Actual).fornro)
    
    For I = LI_WF_Tpa To LS_WF_Tpa
        NParametro = Trim(Arr_WF_TPA(I).Nombre)
        Valor = Arr_WF_TPA(I).Valor
        
        eval.m_SymbolTable.Add NParametro, Valor
    Next I

End Sub



Public Sub CargarTablaParametros()
' carga la tabla de simbolos con los parametros en wf_tpa
Dim symbols As New CSymbolTable
Dim rs_wf_tpa As New ADODB.Recordset
Dim NParametro As String
Dim Valor As Double
'Dim rs_FunFormulas As New ADODB.Recordset
Dim tipo As String

    Set eval.m_SymbolTable = symbols
     
'    'Aca deberia cargar todas las funciones validas
'    StrSql = "SELECT * FROM funformula "
'    OpenRecordset StrSql, rs_FunFormulas
    'FGZ - 20/08/2004
    If rs_FunFormulas.State = adStateOpen Then
        rs_FunFormulas.MoveFirst
    
        Do While Not rs_FunFormulas.EOF
            NParametro = Trim(rs_FunFormulas!fundesabr)
            tipo = rs_FunFormulas!funprograma
            eval.m_SymbolTable.Add NParametro, tipo
            
            rs_FunFormulas.MoveNext
        Loop
    End If
  
  
  
    'FGZ - 23/05/2011 -------------------------------------------------------------------
    'se cambió la tabla temporal wf_tpa por una array
    
    ' Cada parametro en wf_tpa se inserta como un simbolo en la tabla
    StrSql = "SELECT * FROM " & TTempWF_tpa
    OpenRecordset StrSql, rs_wf_tpa
    
    Do While Not rs_wf_tpa.EOF
        NParametro = Trim(rs_wf_tpa!Nombre)
        Valor = rs_wf_tpa!Valor
        
        eval.m_SymbolTable.Add NParametro, Valor
        rs_wf_tpa.MoveNext
    Loop
    
    
    
    
End Sub

Public Sub Liqpro15_OLD(ByVal Nro_empleado As Long, ByVal pliq_desde As Integer, ByVal pliq_hasta As Integer, ByVal Monto_total As Double, ByVal Nro_Concepto As Long, ByVal Cabecera_liq As Long, ByRef OK As Boolean)
' -----------------------------------------------------------------------------------
' Descripcion: realiza el proceso de ajuste retroactivo para la liquidacion para un empleado
' Autor: FGZ
' Fecha:
' Ultima Modificacion:
' -----------------------------------------------------------------------------------
Dim rs_buf_periodo_retro As New ADODB.Recordset ' Tabla Periodo
Dim rs_acumulador As New ADODB.Recordset
Dim rs_Acu_Mes As New ADODB.Recordset
Dim rs_his_retro As New ADODB.Recordset
Dim rs_aux_acu_mes As New ADODB.Recordset

Dim Retro_ano_desde As Integer
Dim Retro_mes_desde As Integer
Dim Retro_ano_hasta As Integer
Dim Retro_mes_hasta As Integer
Dim Anio_aux As Integer
Dim Mes_Aux As Integer
Dim Cant_meses_retro As Integer
Dim Monto_prorratear As Double
Dim Monto_ya_prorrateado As Double
Dim Mes_desde As Integer
Dim Mes_hasta As Integer


' Calcular los meses a ajustar
StrSql = "SELECT * FROM periodo WHERE pliqnro = " & pliq_desde
OpenRecordset StrSql, rs_buf_periodo_retro
If Not rs_buf_periodo_retro.EOF Then
    Retro_ano_desde = rs_buf_periodo_retro!pliqanio
    Retro_mes_desde = rs_buf_periodo_retro!pliqmes
End If


StrSql = "SELECT * FROM periodo WHERE pliqnro = " & pliq_hasta
OpenRecordset StrSql, rs_buf_periodo_retro
If Not rs_buf_periodo_retro.EOF Then
    Retro_ano_hasta = rs_buf_periodo_retro!pliqanio
    Retro_mes_hasta = rs_buf_periodo_retro!pliqmes
End If

' Cuenta la cantidad de meses a ajustar
Cant_meses_retro = ((Retro_ano_hasta - Retro_ano_desde - 1) * 12) + (12 - Retro_mes_desde + 1) + Retro_mes_hasta
Monto_prorratear = Monto_total / Cant_meses_retro
Monto_ya_prorrateado = 0

'Recorre los acumuladores del concepto que requiere ser ajustado
StrSql = "SELECT * FROM con_acum " & _
         " INNER JOIN acumulador ON acumulador.acunro = con_acum.acunro " & _
         " WHERE con_acum.concnro = " & Nro_Concepto & _
         " AND acumulador.acuretro = -1"
OpenRecordset StrSql, rs_acumulador

Do While Not rs_acumulador.EOF
    'FGZ - 27/12/2004
    Monto_ya_prorrateado = 0
    'Recorre los años a ajustar
    Anio_aux = Retro_ano_desde
    
    Do While Anio_aux <= Retro_ano_desde
    
        Mes_desde = IIf((Anio_aux = Retro_ano_desde), Retro_mes_desde, 1)
        Mes_hasta = IIf((Anio_aux = Retro_ano_hasta), Retro_mes_hasta, 12)
    
        Mes_Aux = Mes_desde
        Do While Mes_Aux <= Mes_hasta
            StrSql = "SELECT * FROM acu_mes "
            StrSql = StrSql & " WHERE acunro = " & rs_acumulador!acuNro
            StrSql = StrSql & " AND ammes = " & Mes_Aux
            StrSql = StrSql & " AND ternro = " & buliq_empleado!Ternro
            StrSql = StrSql & " AND amanio = " & Anio_aux
            OpenRecordset StrSql, rs_Acu_Mes
            
            If rs_Acu_Mes.EOF Then
                StrSql = "INSERT INTO acu_mes (acunro,amanio,ammes,ternro,ammonto) VALUES (" & rs_acumulador!acuNro
                StrSql = StrSql & "," & Anio_aux
                StrSql = StrSql & "," & Mes_Aux
                StrSql = StrSql & "," & Nro_empleado
                StrSql = StrSql & ",0)"
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
    
            'Log de la Traza
            StrSql = "SELECT * FROM hisretroactivo " & _
                     " WHERE acunro = " & rs_acumulador!acuNro & _
                     " AND amanio = " & Anio_aux & _
                     " AND ammes = " & Mes_Aux & _
                     " AND cliqnro = " & Cabecera_liq & _
                     " AND concnro = " & Nro_Concepto
            OpenRecordset StrSql, rs_his_retro
            
            If rs_his_retro.EOF Then
                StrSql = "INSERT INTO hisretroactivo (acunro,amanio,ammes,cliqnro,concnro,dlimonto,ammonto)" & _
                         " VALUES (" & rs_acumulador!acuNro & _
                         "," & Anio_aux & _
                         "," & Mes_Aux & _
                         "," & Cabecera_liq & _
                         "," & Nro_Concepto & _
                         "," & Monto_prorratear & _
                         ",0)"
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
            Mes_Aux = Mes_Aux + 1
        Loop
        
        ' Actualizo ambos
        Mes_Aux = Mes_desde
        Do While Mes_Aux <= Mes_hasta
            ' Actualizo Acu_mes
            StrSql = "UPDATE acu_mes SET ammonto = ammonto + " & Monto_prorratear & _
                     " WHERE acunro = " & rs_acumulador!acuNro & _
                     " AND ternro = " & buliq_empleado!Ternro & _
                     " AND ammes = " & Mes_Aux & _
                     " AND amanio = " & Anio_aux
            objConn.Execute StrSql, , adExecuteNoRecords
            
            Monto_ya_prorrateado = Monto_ya_prorrateado + Monto_prorratear
            
            ' Actualizo hisretroactivo
            StrSql = "SELECT ammonto FROM acu_mes " & _
                     " WHERE acunro = " & rs_acumulador!acuNro & _
                     " AND ternro = " & buliq_empleado!Ternro & _
                     " AND ammes = " & Mes_Aux & _
                     " AND amanio = " & Anio_aux
            OpenRecordset StrSql, rs_aux_acu_mes
                                 
'            StrSql = "UPDATE hisretroactivo SET dlimonto = dlimonto + " & rs_aux_acu_mes!ammonto & _
'                     " WHERE acunro = " & rs_acumulador!acunro & _
'                     " AND amanio = " & Anio_aux & _
'                     " AND ammes = " & Mes_Aux & _
'                     " AND cliqnro = " & Cabecera_liq & _
'                     " AND concnro = " & Nro_Concepto
                                 
            'FGZ - 20/12/2004
            StrSql = "UPDATE hisretroactivo SET ammonto = " & rs_aux_acu_mes!ammonto & _
                     " , dlimonto =" & Monto_prorratear & _
                     " WHERE acunro = " & rs_acumulador!acuNro & _
                     " AND amanio = " & Anio_aux & _
                     " AND ammes = " & Mes_Aux & _
                     " AND cliqnro = " & Cabecera_liq & _
                     " AND concnro = " & Nro_Concepto
            objConn.Execute StrSql, , adExecuteNoRecords
        
            ' ciero el recordset auxiliar
            If rs_aux_acu_mes.State = adStateOpen Then rs_aux_acu_mes.Close
            
            Mes_Aux = Mes_Aux + 1
        Loop
        
        
        ' Ultimo mes: Para meter el saldo en el ultimo mes, VERIFICAR COMO FUNCIONA
        If (Anio_aux = Retro_ano_desde) And (Monto_ya_prorrateado <> Monto_total) Then
            ' Actualizo Acu_mes
            StrSql = "UPDATE acu_mes SET ammonto = ammonto + " & (Monto_total - Monto_ya_prorrateado) & _
                     " WHERE acunro = " & rs_acumulador!acuNro & _
                     " AND ternro = " & buliq_empleado!Ternro & _
                     " AND ammes = " & Mes_Aux - 1 & _
                     " AND amanio = " & Anio_aux
            objConn.Execute StrSql, , adExecuteNoRecords
            
            StrSql = "UPDATE hisretroactivo SET dlimonto = dlimonto + " & (Monto_total - Monto_ya_prorrateado) & _
                     ", ammonto = ammonto + " & (Monto_total - Monto_ya_prorrateado) & _
                     " WHERE acunro = " & rs_acumulador!acuNro & _
                     " AND amanio = " & Anio_aux & _
                     " AND ammes = " & Mes_Aux - 1 & _
                     " AND cliqnro = " & Cabecera_liq & _
                     " AND concnro = " & Nro_Concepto
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        
        
        Anio_aux = Anio_aux + 1
    Loop
    
    rs_acumulador.MoveNext
Loop

'FGZ - 27/12/2004
OK = True
End Sub


Public Sub Liqpro15(ByVal Nro_empleado As Long, ByVal pliq_desde As Integer, ByVal pliq_hasta As Integer, ByVal Monto_total As Double, ByVal Nro_Concepto As Long, ByVal Cabecera_liq As Long, ByRef OK As Boolean)
' -----------------------------------------------------------------------------------
' Descripcion: realiza el proceso de ajuste retroactivo para la liquidacion para un empleado
' Autor: Martin Ferraro
' Fecha: 18/06/2010
' Ultima Modificacion:
' -----------------------------------------------------------------------------------
Dim rs_acumulador As New ADODB.Recordset
Dim rs_Acu_Mes As New ADODB.Recordset
Dim rs_his_retro As New ADODB.Recordset
Dim rs_wf_Retro As New ADODB.Recordset
Dim SumNov As Double
Dim proporcion As Double
Dim Result As Double

Concepto_Retroactivo = 0

'--------------------------------------------------------------------------------------------------------
'Total de novedad retro para el concepto
'--------------------------------------------------------------------------------------------------------
SumNov = 0
StrSql = "SELECT SUM(monto) total FROM " & TTempWF_Retroactivo
StrSql = StrSql & " WHERE concnro = " & Arr_conceptos(Concepto_Actual).ConcNro
OpenRecordset StrSql, rs_wf_Retro
If Not rs_wf_Retro.EOF Then
    If Not EsNulo(rs_wf_Retro!total) Then SumNov = rs_wf_Retro!total
End If
rs_wf_Retro.Close


'--------------------------------------------------------------------------------------------------------
'Por cada registro retroactivo creado para el concepto
'--------------------------------------------------------------------------------------------------------
'FGZ - 08/06/2012 ----------
StrSql = "SELECT concnro,anio,mes,monto FROM " & TTempWF_Retroactivo
StrSql = StrSql & " WHERE concnro = " & Arr_conceptos(Concepto_Actual).ConcNro
OpenRecordset StrSql, rs_wf_Retro
Do While Not rs_wf_Retro.EOF
    
    Concepto_Retroactivo = -1
    
    proporcion = rs_wf_Retro!Monto / SumNov
    Result = (Monto_total * proporcion)
    
    '--------------------------------------------------------------------------------------------------------
    'Recorre los acumuladores del concepto que requiere ser ajustado
    '--------------------------------------------------------------------------------------------------------
    'FGZ - 08/06/2012 --------------
    StrSql = "SELECT con_acum.acunro FROM con_acum " & _
             " INNER JOIN acumulador ON acumulador.acunro = con_acum.acunro " & _
             " WHERE con_acum.concnro = " & rs_wf_Retro!ConcNro & _
             " AND acumulador.acuretro = -1"
    OpenRecordset StrSql, rs_acumulador
    Do While Not rs_acumulador.EOF
    
        '--------------------------------------------------------------------------------------------------------
        'Actualizo Acumes
        '--------------------------------------------------------------------------------------------------------
        'FGZ - 08/06/2012 --------------
        StrSql = "SELECT acunro FROM acu_mes "
                    StrSql = StrSql & " WHERE acunro = " & rs_acumulador!acuNro
                    StrSql = StrSql & " AND ammes = " & rs_wf_Retro!Mes
                    StrSql = StrSql & " AND ternro = " & buliq_empleado!Ternro
                    StrSql = StrSql & " AND amanio = " & rs_wf_Retro!Anio
        OpenRecordset StrSql, rs_Acu_Mes
        
        If rs_Acu_Mes.EOF Then
            StrSql = "INSERT INTO acu_mes (acunro,amanio,ammes,ternro,ammonto) VALUES (" & rs_acumulador!acuNro
            StrSql = StrSql & "," & rs_wf_Retro!Anio
            StrSql = StrSql & "," & rs_wf_Retro!Mes
            StrSql = StrSql & "," & buliq_empleado!Ternro
            StrSql = StrSql & "," & Result & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        Else
            StrSql = "UPDATE acu_mes SET ammonto = ammonto + " & Result & _
                     " WHERE acunro = " & rs_acumulador!acuNro & _
                     " AND ternro = " & buliq_empleado!Ternro & _
                     " AND ammes = " & rs_wf_Retro!Mes & _
                     " AND amanio = " & rs_wf_Retro!Anio
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        
        '--------------------------------------------------------------------------------------------------------
        'Actualizo HisRetro
        '--------------------------------------------------------------------------------------------------------
        'FGZ - 08/06/2012 --------------
        StrSql = "SELECT acunro FROM hisretroactivo " & _
                 " WHERE acunro = " & rs_acumulador!acuNro & _
                 " AND amanio = " & rs_wf_Retro!Anio & _
                 " AND ammes = " & rs_wf_Retro!Mes & _
                 " AND cliqnro = " & Cabecera_liq & _
                 " AND concnro = " & rs_wf_Retro!ConcNro
        OpenRecordset StrSql, rs_his_retro
        
        If rs_his_retro.EOF Then
            StrSql = "INSERT INTO hisretroactivo (acunro,amanio,ammes,cliqnro,concnro,dlimonto,ammonto)" & _
                     " VALUES (" & rs_acumulador!acuNro & _
                     "," & rs_wf_Retro!Anio & _
                     "," & rs_wf_Retro!Mes & _
                     "," & Cabecera_liq & _
                     "," & rs_wf_Retro!ConcNro & _
                     "," & Result & _
                     ",0)"
            objConn.Execute StrSql, , adExecuteNoRecords
        Else
            StrSql = "UPDATE hisretroactivo SET ammonto = ammonto + " & Result & _
                     " , dlimonto = dlimonto + " & Result & _
                     " WHERE acunro = " & rs_acumulador!acuNro & _
                     " AND amanio = " & rs_wf_Retro!Anio & _
                     " AND ammes = " & rs_wf_Retro!Mes & _
                     " AND cliqnro = " & Cabecera_liq & _
                     " AND concnro = " & rs_wf_Retro!ConcNro
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        
        rs_acumulador.MoveNext
    Loop
    rs_acumulador.Close
    
rs_wf_Retro.MoveNext
OK = True
Loop


End Sub


Public Sub SetearAmpo()
Dim rs_Ampo As New ADODB.Recordset
Dim rs_AmpoAux As New ADODB.Recordset

    'EAM(v6.73) - Se agrego la funcion top porque utiliza el ultimo registro y se resuleve todo en un query
    StrSql = "SELECT ampofecha,valor,contvalor,ampopropo,ampodiario,ampomax,ampomin,ampotconnro FROM ampo  " & _
            " WHERE ampofecha = (SELECT MAX(ampofecha) FROM ampo WHERE ampofecha <= " & ConvFecha(buliq_proceso.Fields.Item("profecfin").Value) & ")"
    OpenRecordset StrSql, rs_Ampo

    If Not rs_Ampo.EOF Then
        'Se quito la sgt linea
        'rs_AmpoAux.MoveLast
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 1) & "AMPO Fecha de Vigencia " & rs_Ampo.Fields.Item("ampofecha").Value
        End If

        
        Do While Not rs_Ampo.EOF
            Valor_Ampo = rs_Ampo.Fields.Item("Valor").Value
            Valor_Ampo_Cont = rs_Ampo.Fields.Item("contvalor").Value
            
            Select Case rs_Ampo.Fields.Item("ampotconnro").Value
            Case 1:
                Cant_Diaria_Ampos_1 = rs_Ampo.Fields.Item("ampodiario").Value
                Cant_Ampo_Proporcionar_1 = 0
                Ampo_Proporciona_1 = CBool(rs_Ampo.Fields.Item("ampopropo").Value)
                Ampo_Max_1 = rs_Ampo.Fields.Item("ampomax").Value
                Ampo_Min_1 = rs_Ampo.Fields.Item("ampomin").Value
            Case 2:
                Cant_Diaria_Ampos_2 = rs_Ampo.Fields.Item("ampodiario").Value
                Cant_Ampo_Proporcionar_2 = 0
                Ampo_Proporciona_2 = CBool(rs_Ampo.Fields.Item("ampopropo").Value)
                Ampo_Max_2 = rs_Ampo.Fields.Item("ampomax").Value
                Ampo_Min_2 = rs_Ampo.Fields.Item("ampomin").Value
            Case 3:
                Cant_Diaria_Ampos_3 = 30 / rs_Ampo!ampomax
                Cant_Ampo_Proporcionar_3 = 0
                Ampo_Proporciona_3 = CBool(rs_Ampo.Fields.Item("ampopropo").Value)
                Ampo_Max_3 = rs_Ampo.Fields.Item("ampomax").Value
                Ampo_Min_3 = rs_Ampo.Fields.Item("ampomin").Value
            Case 4:
                Cant_Diaria_Ampos_4 = rs_Ampo.Fields.Item("ampodiario").Value
                Cant_Ampo_Proporcionar_4 = 0
                Ampo_Proporciona_4 = CBool(rs_Ampo.Fields.Item("ampopropo").Value)
                Ampo_Max_4 = rs_Ampo.Fields.Item("ampomax").Value
                Ampo_Min_4 = rs_Ampo.Fields.Item("ampomin").Value
            Case 5:
                Cant_Diaria_Ampos_5 = rs_Ampo.Fields.Item("ampodiario").Value
                Cant_Ampo_Proporcionar_5 = 0
                Ampo_Proporciona_5 = CBool(rs_Ampo.Fields.Item("ampopropo").Value)
                Ampo_Max_5 = rs_Ampo.Fields.Item("ampomax").Value
                Ampo_Min_5 = rs_Ampo.Fields.Item("ampomin").Value
            End Select
            
            rs_Ampo.MoveNext
        Loop
    Else
        If HACE_TRAZA Then
            Call InsertarTraza(NroCab, 0, 0, "No se encontr¢ la estructura del AMPO para la fecha de pago del Proceso", 0)
        End If
    End If
End Sub


Private Function Existe_Origen(ByRef Posicion As Long, ByVal Origen As Long) As Boolean
Dim Aux_Pos As Integer
Dim Esta As Boolean
Dim Termino As Boolean
    
    Aux_Pos = Posicion
    Termino = False
    Esta = False
    
    Do While Not Esta And Not Termino And Arr_Cge_Segun(Posicion).Nivel = 1 And Arr_conceptos(Concepto_Actual).ConcNro = Arr_Cge_Segun(Posicion).ConcNro
        If Arr_Cge_Segun(Posicion).Origen = Origen Then
            Esta = True
        Else
            Posicion = Posicion + 1
            If Posicion > Max_Cge_Segun Then
                Posicion = Aux_Pos
                Termino = True
            End If
        End If
    Loop
    Existe_Origen = Esta
End Function

Public Sub Desmarcar_BAE()
' -----------------------------------------------------------------------------------
' Descripcion: Desmarca los bae liquidados por el proceso.
' Autor: FGZ
' Fecha: 26/05/2005
' Ultima Modificacion:
' -----------------------------------------------------------------------------------

    On Error GoTo ME_Local
    
'        StrSql = "DELETE FROM wf_bae "
'        StrSql = StrSql & " WHERE ternro = " & buliq_empleado!ternro
'        StrSql = StrSql & " AND pronro = " & buliq_proceso!pronro
'        objConn.Execute StrSql, , adExecuteNoRecords
        
        StrSql = "DELETE FROM ee_parte_bae "
        StrSql = StrSql & " WHERE ternro = " & buliq_empleado!Ternro
        StrSql = StrSql & " AND pronro = " & buliq_proceso!pronro
        objConn.Execute StrSql, , adExecuteNoRecords

Exit Sub

ME_Local:
' FGZ - 17/03/2006 - se desactivó el mensaje de log
'    Flog.writeline
'    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
'    Flog.writeline Espacios(Tabulador * 3) & " Desmarcado de BAE " & Err.Description
'    Flog.writeline
'    Flog.writeline Espacios(Tabulador * 3) & "SQL Ejecutado: " & StrSql
'    Flog.writeline
'    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
'    Flog.writeline
End Sub





Public Sub Desmarcar_Vales()
' -----------------------------------------------------------------------------------
' Descripcion: Desmarca los Vales liquidados por el proceso.
' Autor: FGZ
' Fecha: 29/05/2012
' Ultima Modificacion: FGZ - 01/06/2012 - Le agregué with (rowlock)
' -----------------------------------------------------------------------------------
Dim Intentos As Integer
Dim LimiteIntentos As Integer
Dim TiempoEspera As Integer

    TiempoEspera = 5
    LimiteIntentos = 3
    
    Intentos = 1
    On Error GoTo ME_Local
    
    If TipoBD = 3 Then
        StrSql = "UPDATE vales with (rowlock) SET pronro = null " & _
                 " WHERE pronro = " & buliq_proceso!pronro & _
                 " AND vales.empleado = " & buliq_empleado!Ternro
    Else
        StrSql = "UPDATE vales SET pronro = null " & _
                 " WHERE pronro = " & buliq_proceso!pronro & _
                 " AND vales.empleado = " & buliq_empleado!Ternro
    End If
Retry:
    objConn.Execute StrSql, , adExecuteNoRecords

Fin:
Exit Sub

ME_Local:
    If Err.Number = -2147467259 Then
        'Error de loqueo
        Intentos = Intentos + 1
        If Intentos <= LimiteIntentos Then
            Flog.writeline Espacios(Tabulador * 4) & " Lock ... wait "
            Sleep (TiempoEspera)
            Err.Number = 0
            GoTo Retry
        End If
    End If
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 3) & " Desmarcado de Vales " & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline
    GoTo Fin
End Sub

Public Sub Desmarcar_Licencias()
' -----------------------------------------------------------------------------------
' Descripcion: Desmarca las Licencias liquidadas por el proceso.
' Autor: FGZ
' Fecha: 29/05/2012
' Ultima Modificacion: 01/06/2012 - le agregué with rowlock
' -----------------------------------------------------------------------------------
Dim Intentos As Integer
Dim LimiteIntentos As Integer
Dim TiempoEspera As Integer

    TiempoEspera = 5
    LimiteIntentos = 3
    
    Intentos = 1
    On Error GoTo ME_Local
    
    If TipoBD = 3 Then
        StrSql = "UPDATE emp_lic with (rowlock) SET pronro = null "
        StrSql = StrSql & " WHERE pronro = " & buliq_proceso!pronro
        StrSql = StrSql & " AND empleado = " & buliq_empleado!Ternro
    Else
        StrSql = "UPDATE emp_lic SET pronro = null "
        StrSql = StrSql & " WHERE pronro = " & buliq_proceso!pronro
        StrSql = StrSql & " AND empleado = " & buliq_empleado!Ternro
    End If
Retry:
        objConn.Execute StrSql, , adExecuteNoRecords

Fin:
Exit Sub

ME_Local:
    If Err.Number = -2147467259 Then
        'Error de loqueo
        Intentos = Intentos + 1
        If Intentos <= LimiteIntentos Then
            Flog.writeline Espacios(Tabulador * 4) & " Lock ... wait "
            Sleep (TiempoEspera)
            Err.Number = 0
            GoTo Retry
        End If
    End If
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 3) & " Desmarcado de Licencias " & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline
    GoTo Fin
End Sub


Public Sub Desmarcar_Insalubridad()
' -----------------------------------------------------------------------------------
' Descripcion: Desmarca las Licencias de insalubridad pagadas por el proceso.
' Autor: FGZ
' Fecha: 17/09/2012
' Ultima Modificacion:
' -----------------------------------------------------------------------------------
On Error GoTo ME_Local
        
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 3) & "Desmarcando Insalubridad"
    End If
    
    StrSql = "DELETE FROM lic_pagas "
    StrSql = StrSql & " WHERE liqpronro =" & buliq_proceso!pronro
    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
    objConn.Execute StrSql, , adExecuteNoRecords

Fin:
Exit Sub

ME_Local:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 3) & " Desmarcado de Insalubridad " & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline
    GoTo Fin
End Sub

Public Sub Desmarcar_Comisiones()
' -----------------------------------------------------------------------------------
' Descripcion: Desmarca el detalle de las comisiones pagadas por el proceso.
' Autor: FGZ
' Fecha: 01/10/2012
' Ultima Modificacion:
' -----------------------------------------------------------------------------------
On Error GoTo ME_Local
        
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 3) & "Desmarcando Comisiones"
    End If
    
    StrSql = "DELETE FROM liq_comision "
    StrSql = StrSql & " WHERE pronro =" & buliq_proceso!pronro
    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
    objConn.Execute StrSql, , adExecuteNoRecords

Fin:
Exit Sub

ME_Local:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 3) & " Desmarcado de Comisiones " & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline
    GoTo Fin
End Sub

Public Sub Desmarcar_Gastos()
' -----------------------------------------------------------------------------------
' Descripcion: Desmarca los Gastos pagados por el proceso.
' Autor: FGZ
' Fecha: 14/01/2013
' Ultima Modificacion:
' -----------------------------------------------------------------------------------
On Error GoTo ME_Local
        
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 3) & "Desmarcando Gastos"
    End If
    
    StrSql = "UPDATE gastos SET pronro = null "
    StrSql = StrSql & " WHERE pronro = " & buliq_proceso!pronro
    StrSql = StrSql & " AND ternro = " & buliq_empleado!Ternro
    objConn.Execute StrSql, , adExecuteNoRecords

Fin:
Exit Sub

ME_Local:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 3) & " Desmarcado de Gastos " & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline
    GoTo Fin
End Sub


Public Sub Desmarcar_DiasVacVendidos()
' -----------------------------------------------------------------------------------
' Descripcion: Desmarca los dias de vacaciones vendidos por el proceso.
' Autor: FGZ
' Fecha: 28/02/2013
' Ultima Modificacion:
' -----------------------------------------------------------------------------------
On Error GoTo ME_Local
        
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 3) & "Desmarcando Venta de vacaciones"
    End If
    
    StrSql = "DELETE vacvendidos " & _
            " WHERE pronro = " & buliq_proceso!pronro & " AND ternro = " & buliq_empleado!Ternro & _
            " AND automatico= -1"
    objConn.Execute StrSql, , adExecuteNoRecords
    
    StrSql = "UPDATE vacvendidos SET pronro = 0 WHERE pronro = " & buliq_proceso!pronro & " AND ternro =" & buliq_empleado!Ternro
    objConn.Execute StrSql, , adExecuteNoRecords
    
Fin:
Exit Sub

ME_Local:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 3) & " Desmarcado de Venta de cacaciones " & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline
    GoTo Fin
End Sub


Public Sub Desmarcar_PagosDtos()
' -----------------------------------------------------------------------------------
' Descripcion: Desmarca los Pagos y descuentos liquidados por el proceso.
' Autor: FGZ
' Fecha: 29/05/2012
' Ultima Modificacion: FGZ - 01/06/2012 - Le agregué  with (rowlock)
' -----------------------------------------------------------------------------------
Dim Intentos As Integer
Dim LimiteIntentos As Integer
Dim TiempoEspera As Integer

    TiempoEspera = 5
    LimiteIntentos = 3
    
    Intentos = 1
    On Error GoTo ME_Local
    
    If TipoBD = 3 Then
        StrSql = "UPDATE vacpagdesc with (rowlock) SET pronro = null "
        StrSql = StrSql & " WHERE pronro = " & buliq_proceso!pronro
        StrSql = StrSql & " AND ternro = " & buliq_empleado!Ternro
    Else
        StrSql = "UPDATE vacpagdesc SET pronro = null "
        StrSql = StrSql & " WHERE pronro = " & buliq_proceso!pronro
        StrSql = StrSql & " AND ternro = " & buliq_empleado!Ternro
    End If
Retry:
        objConn.Execute StrSql, , adExecuteNoRecords

Fin:
Exit Sub

ME_Local:
    If Err.Number = -2147467259 Then
        'Error de loqueo
        Intentos = Intentos + 1
        If Intentos <= LimiteIntentos Then
            Flog.writeline Espacios(Tabulador * 4) & " Lock ... wait "
            Sleep (TiempoEspera)
            Err.Number = 0
            GoTo Retry
        End If
    End If
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 3) & " Desmarcado de Pagos/Dtos " & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline
    GoTo Fin
End Sub


Public Sub DesmarcarAnticipo()
' -----------------------------------------------------------------------------------
' Descripcion: Desmarca los Anticipos liquidados por el proceso.
' Autor: Martin
' Fecha: 01/07/2010
' Ultima Modificacion:
' -----------------------------------------------------------------------------------

    On Error GoTo ME_Local
    
        StrSql = "UPDATE anticipos SET pronro = null, antestado = 2 " & _
                 " WHERE pronro = " & buliq_proceso!pronro & _
                 " AND anticipos.empleado = " & buliq_empleado!Ternro & _
                 " AND antestado = 3"
        objConn.Execute StrSql, , adExecuteNoRecords


Fin:
    Exit Sub

ME_Local:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 3) & " Desmarcado de Anticipos " & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline
    GoTo Fin
End Sub

Public Function EsModeloRecalculo(ByVal Proceso As Long) As Boolean
' -----------------------------------------------------------------------------------
' Descripcion: Verifica si el modelo es de recalculo
' Autor: Martin
' Fecha: 29/01/2009
' Ultima Modificacion:
' -----------------------------------------------------------------------------------
Dim rs_Recalc As New ADODB.Recordset
Dim Salida As Boolean
On Error GoTo ME_Local
    
    Salida = False
    
    StrSql = "SELECT *"
    StrSql = StrSql & " FROM proceso INNER JOIN"
    StrSql = StrSql & " tipoproc ON proceso.tprocnro = tipoproc.tprocnro"
    StrSql = StrSql & " AND tprocrecalculo = -1"
    StrSql = StrSql & " Where Proceso.pronro = " & Proceso
    OpenRecordset StrSql, rs_Recalc
    Salida = Not rs_Recalc.EOF
    rs_Recalc.Close
    Set rs_Recalc = Nothing
    
    EsModeloRecalculo = Salida
    
Exit Function
ME_Local:
    
    'Por si no existe el campo tprocrecalculo que es custom de Chile
    EsModeloRecalculo = False
    
End Function

Public Sub Desmarcar_Mov()
' -----------------------------------------------------------------------------------
' Descripcion: Desmarca los Movimientos liquidados por el proceso.
' Autor: Martin Ferraro
' Fecha: 26/09/2008
' Ultima Modificacion:
' -----------------------------------------------------------------------------------

On Error GoTo ME_Local
        
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 3) & "Desmarcando Movimientos"
    End If
    StrSql = "UPDATE gti_movimientos SET pronro = null"
    StrSql = StrSql & " ,movfecliq = null"
    StrSql = StrSql & " ,concnro = null"
    StrSql = StrSql & " WHERE pronro = " & buliq_proceso!pronro
    StrSql = StrSql & " AND ternro = " & buliq_empleado!Ternro
    objConn.Execute StrSql, , adExecuteNoRecords

Fin:
Exit Sub

ME_Local:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 3) & " Desmarcado de Movimientos de gti " & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline
    GoTo Fin
End Sub

Public Sub Desmarcar_Prestamos()
' -----------------------------------------------------------------------------------
' Descripcion: Desmarca los Prestamos liquidados por el proceso.
' Autor: Martin Ferraro
' Fecha: 04/08/2010
' Ultima Modificacion:
' -----------------------------------------------------------------------------------
Dim rs_Prestamos As New ADODB.Recordset

On Error GoTo ME_Local
        
        'BORRA LAS CUOTAS DE PRESTAMOS GENERADOS POR EL PROCESO
        If CBool(USA_DEBUG) Then Flog.writeline Espacios(Tabulador * 1) & "BORRA LAS CUOTAS DE PRESTAMOS GENERADOS POR EL PROCESO"
        
        'Busco las cuotas que fueron generadas por el proceso
        StrSql = "SELECT cuonro, cuototal, cuosaldo, cuogenera , pronrogenera, cuoordenliq "
        StrSql = StrSql & " FROM pre_cuota"
        StrSql = StrSql & " INNER JOIN prestamo ON prestamo.prenro = pre_cuota.prenro"
        StrSql = StrSql & " WHERE pronrogenera = " & buliq_proceso!pronro
        StrSql = StrSql & " AND prestamo.ternro = " & buliq_empleado!Ternro
        StrSql = StrSql & " ORDER BY cuonrocuo, cuoordenliq"
        OpenRecordset StrSql, rs_Prestamos
        Do While Not rs_Prestamos.EOF
            
            'Actualizo la cuota que particiono
            StrSql = "UPDATE pre_cuota"
            'If TodoNada Then
            '    StrSql = StrSql & " SET cuototal = " & rs_Prestamos!cuototal
            'Else
                StrSql = StrSql & " SET cuototal = cuototal + " & rs_Prestamos!cuototal
            'End If
            StrSql = StrSql & " , cuosaldo = " & rs_Prestamos!cuosaldo
            StrSql = StrSql & " WHERE cuonro = " & rs_Prestamos!cuogenera
            objConn.Execute StrSql, , adExecuteNoRecords
            
            'Borro la cuota generada
            StrSql = "DELETE pre_cuota WHERE cuonro = " & rs_Prestamos!cuonro
            objConn.Execute StrSql, , adExecuteNoRecords
            
            rs_Prestamos.MoveNext
        Loop
        rs_Prestamos.Close
        
        'FGZ - 20/05/2011 ------------------------------
        'StrSql = "SELECT * FROM prestamo "
        StrSql = "SELECT prestamo.estnro, prestamo.prenro,cuonro FROM prestamo " & _
                 "INNER JOIN pre_cuota ON prestamo.prenro = pre_cuota.prenro" & _
                 " WHERE prestamo.ternro = " & buliq_empleado!Ternro & _
                 " AND pre_cuota.pronro =" & buliq_proceso!pronro & _
                 " AND pre_cuota.cuocancela = -1"
        OpenRecordset StrSql, rs_Prestamos

        Do While Not rs_Prestamos.EOF
            'FGZ - 22/04/2015 --------------------------------------------
            'Priemro deberia revisar si  hubo alguna cuota con saldo cancelado pero marcada y que no se generó para el siguiente mes
            StrSql = "UPDATE pre_cuota SET cuosaldo = (cuosaldo + cuocancelado), cuototal = cuocancelado, cuocancelado = NULL "
            StrSql = StrSql & " WHERE pronro = " & buliq_proceso!pronro
            StrSql = StrSql & " AND prenro = " & rs_Prestamos!prenro
            StrSql = StrSql & " AND cuonro = " & rs_Prestamos!cuonro
            StrSql = StrSql & " AND cuototal = 0 "
            StrSql = StrSql & " AND cuocancela = -1"
            StrSql = StrSql & " AND cuocancelado IS NOT NULL"
            objConn.Execute StrSql, , adExecuteNoRecords
            'FGZ - 22/04/2015 --------------------------------------------
        
            'Es una cuota cancelada, la marco como no cancelada
            StrSql = "UPDATE pre_cuota SET pronro = null, cuocancela = 0 "
            StrSql = StrSql & " WHERE pronro = " & buliq_proceso!pronro
            StrSql = StrSql & " AND prenro = " & rs_Prestamos!prenro
            StrSql = StrSql & " AND cuonro = " & rs_Prestamos!cuonro
            objConn.Execute StrSql, , adExecuteNoRecords
            

            If rs_Prestamos!estnro = 6 Then
                ' Lo apruebo para re-liq la cuota
                'StrSql = "UPDATE prestamo SET estnro = 3 " & _
                '         " WHERE ternro = " & rs_Prestamos!Ternro & _
                '         " AND prenro =" & rs_Prestamos!prenro
                
                'FGZ - 08/11/2011 --------------------------------
                StrSql = "UPDATE prestamo SET estnro = 3 " & _
                         " WHERE ternro = " & buliq_empleado!Ternro & _
                         " AND prenro =" & rs_Prestamos!prenro
                'FGZ - 08/11/2011 --------------------------------
                 objConn.Execute StrSql, , adExecuteNoRecords
            End If
            
            rs_Prestamos.MoveNext
        Loop

Fin:
If rs_Prestamos.State = adStateOpen Then rs_Prestamos.Close
Set rs_Prestamos = Nothing
Exit Sub

ME_Local:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 3) & " Desmarcado de Prestamos " & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline
    GoTo Fin

End Sub



Public Sub Desmarcar_Embargos(ByVal Nro_Ter As Long, ByVal Nro_Proceso As Long)
' -----------------------------------------------------------------------------------
' Descripcion: Desmarca los embargos liquidados por el proceso.
' Autor: FGZ
' Fecha: 26/05/2005
' Ultima Modificacion:
' -----------------------------------------------------------------------------------
Dim rs_Embargos As New ADODB.Recordset

    On Error GoTo ME_Local
    
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 3) & "BORRA LAS CUOTAS DE EMBARGO GENERADAS POR EL PROCESO"
    End If
    'FGZ - 20/05/2011 ------------------------------------
    'StrSql = "SELECT * FROM embargo " & _

    StrSql = "SELECT tipoemb.tpefordesc, embargo.embnro, embargo.embest, embcuota.embcnro, embcuota.embccancela, embcuota.embcimpreal FROM embargo " & _
             " INNER JOIN tipoemb ON embargo.tpenro = tipoemb.tpenro" & _
             " INNER JOIN embcuota ON embargo.embnro = embcuota.embnro" & _
             " WHERE embargo.ternro =" & Nro_Ter & _
             " AND embcuota.pronro =" & Nro_Proceso & _
             " ORDER BY embargo.embnro, embcuota.embcnro "
    OpenRecordset StrSql, rs_Embargos
    
    If rs_Embargos.EOF Then
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 3) & "No se encontraron cuotas que borrar."
        End If
    End If

    Do While Not rs_Embargos.EOF
    
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 3) & "Procesando Embargo Nro.: " & rs_Embargos!embnro & " Cuota Nro: " & rs_Embargos!embcnro
        End If
    
        If (CInt(rs_Embargos!tpefordesc) < 1) Then ' El embargo es por monto
            If (CDbl(rs_Embargos!embcimpreal) = 0) And (CInt(rs_Embargos!embccancela) = 0) Then
                'Borro la cuota generada
                StrSql = "DELETE FROM embcuota " & _
                        " WHERE embnro =" & rs_Embargos!embnro & _
                        " AND pronro =" & Nro_Proceso & _
                        " AND embcnro =" & rs_Embargos!embcnro
                objConn.Execute StrSql, , adExecuteNoRecords

                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 3) & "Se borra la cuota Nro: " & rs_Embargos!embcnro & " del Embargo Nro.: " & rs_Embargos!embnro
                End If
            Else 'Marco la cuota como no cancelada
                StrSql = "UPDATE embcuota SET embccancela = 0 " & _
                        " , embcimpreal = 0" & _
                        " , pronro = null" & _
                        " WHERE embnro =" & rs_Embargos!embnro & _
                        " AND pronro =" & Nro_Proceso & _
                        " AND embcnro =" & rs_Embargos!embcnro
                objConn.Execute StrSql, , adExecuteNoRecords
                
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 3) & "Se desliquido la cuota Nro: " & rs_Embargos!embcnro & " del Embargo Nro.: " & rs_Embargos!embnro
                End If
            End If
        Else
        ' El embargo es por porcentaje

'            'Marco la cuota como no cancelada
'            If rs_Embargos!embcimpreal <> 0 Then
'                StrSql = "UPDATE embcuota SET embcimpreal = 0 " & _
'                        " , pronro = null" & _
'                        " WHERE embnro =" & rs_Embargos!embnro & _
'                        " AND pronro =" & Nro_Proceso & _
'                        " AND embcnro =" & rs_Embargos!embcnro
'                objConn.Execute StrSql, , adExecuteNoRecords
'            Else
                
                'Borro las cuotas generadas por el proceso
                StrSql = "DELETE FROM embcuota " & _
                    " WHERE embnro =" & rs_Embargos!embnro & _
                    " AND pronro =" & Nro_Proceso & _
                    " AND embcnro =" & rs_Embargos!embcnro
                    
                objConn.Execute StrSql, , adExecuteNoRecords
'            End If
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 3) & "Se borro la cuota " & rs_Embargos!embcnro & " generadas por el proceso actual para el Embargo Nro.: " & rs_Embargos!embnro
            End If
        End If

        If rs_Embargos!embest = "F" Then
            StrSql = "UPDATE embargo SET embest = 'A' " & _
                 " WHERE embnro =" & rs_Embargos!embnro
            objConn.Execute StrSql, , adExecuteNoRecords
            
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 3) & "Se cambio el estado del embargo"
            End If
        End If

        rs_Embargos.MoveNext
    Loop
    
    'FGZ - 04/10/2011 ------------------------------------
    'Ademas hay que restaurar el estado del embargo cuando se activó por una liquidacion
    StrSql = "UPDATE embargo SET embest = embestant"
    StrSql = StrSql & ", pronro = 0"
    StrSql = StrSql & " WHERE embargo.ternro =" & Nro_Ter
    StrSql = StrSql & " AND pronro =" & Nro_Proceso
    StrSql = StrSql & " AND embestant = 'E'"
    objConn.Execute StrSql, , adExecuteNoRecords
    'FGZ - 04/10/2011 ------------------------------------
    
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 1) & "DESLIQUIDACION DE EMBARGOS REALIZADA EXITOSAMENTE"
    End If

Fin:
If rs_Embargos.State = adStateOpen Then rs_Embargos.Close
Exit Sub

ME_Local:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 3) & " Desmarcado de Embargos " & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline
    GoTo Fin
End Sub


Public Sub LimpiarTrazaConcepto(ByVal Cabecera As Long, ByVal Concepto As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Limpia la Traza para un empleado/concepto.
' Autor      : FGZ
' Fecha      : 08/09/2003
' Ultima Mod :
' Descripcion:
' ---------------------------------------------------------------------------------------------
    
    If ReusaTraza Then
        StrSql = "UPDATE traza SET cliqnro = -1, travalor = 0"
        StrSql = StrSql & " ,concnro = -1 ,tpanro = -1, tradesc = NULL"
        StrSql = StrSql & " WHERE cliqnro = " & Cabecera & " AND concnro = " & Concepto
        
    Else
        StrSql = "DELETE FROM traza WHERE cliqnro = " & Cabecera & " AND concnro = " & Concepto
    End If
    objConn.Execute StrSql, , adExecuteNoRecords
End Sub




Public Sub LimpiarTraza(ByVal Cabecera As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Limpia la Traza para un empleado/concepto.
' Autor      : FGZ
' Fecha      : 08/09/2003
' Ultima Mod :
' Descripcion:
' ---------------------------------------------------------------------------------------------
    
    If ReusaTraza Then
    
        StrSql = "UPDATE traza SET cliqnro = -1, travalor = 0"
        StrSql = StrSql & " ,concnro = -1 ,tpanro = -1, tradesc = NULL"
        StrSql = StrSql & " WHERE cliqnro = " & Cabecera
        
    Else
    
        StrSql = "DELETE FROM traza WHERE cliqnro = " & Cabecera
        
    End If
    
    objConn.Execute StrSql, , adExecuteNoRecords

End Sub


Public Sub InsertarTraza(ByVal cliqnro As Long, ByVal Concepto As Long, ByVal tpanro As Long, ByVal Desc As String, ByVal Valor As Double)
' ---------------------------------------------------------------------------------------------
' Descripcion: Graba un registro de traza para un empleado/concepto. {Traza.i}
' Autor      : Lic.Mauricio RHPro
' Fecha      : 27/10/1996
' Traduccion : FGZ
' Fecha      : 05/09/2003
' Ultima Mod :
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_Traza As New ADODB.Recordset
Dim Aux_Texto As String

On Error GoTo MLocal:

    Aux_Texto = Left(Desc, 60)
    
    If ReusaTraza Then
        
        'Busco una traza libre
        'StrSql = "SELECT tranro FROM traza " & NOLOCK & " WHERE cliqnro = -1 AND travalor = 0"
        StrSql = "SELECT " & TOP(1) & " tranro FROM traza " & NOLOCK & " WHERE cliqnro = -1 AND travalor = 0"
        StrSql = StrSql & " AND concnro = -1 AND tpanro = -1 AND tradesc IS NULL"
        OpenRecordset StrSql, rs_Traza
        
        If Not rs_Traza.EOF Then
            'Actualizo el registro encontrado
            StrSql = "UPDATE traza"
            StrSql = StrSql & " SET cliqnro = " & cliqnro
            StrSql = StrSql & " ,concnro = " & Concepto
            StrSql = StrSql & " ,tpanro = " & tpanro
            StrSql = StrSql & " ,tradesc = '" & Aux_Texto & "'"
            StrSql = StrSql & " ,travalor = " & Valor
            StrSql = StrSql & " WHERE tranro = " & rs_Traza!tranro
            objConn.Execute StrSql, , adExecuteNoRecords
        Else
            'Creo un nuevo registro
            StrSql = "INSERT INTO traza (cliqnro,concnro,tpanro,tradesc,travalor,trafrecuencia)"
            StrSql = StrSql & " VALUES (" & cliqnro
            StrSql = StrSql & " ," & Concepto
            StrSql = StrSql & " ," & tpanro
            StrSql = StrSql & " ,'" & Aux_Texto
            StrSql = StrSql & " '," & Valor
            StrSql = StrSql & " ,'0000000')"
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
    
    Else
        
        StrSql = "INSERT INTO traza (cliqnro,concnro,tpanro,tradesc,travalor,trafrecuencia)" & _
                 " VALUES (" & cliqnro & _
                 "," & Concepto & _
                 "," & tpanro & _
                 ",'" & Aux_Texto & _
                 "'," & Valor & _
                 ",'" & Format(ContadorProgreso, "0000000") & _
                 "')"
        objConn.Execute StrSql, , adExecuteNoRecords
        ContadorProgreso = ContadorProgreso + 1
        
    End If

    If rs_Traza.State = adStateOpen Then rs_Traza.Close
    Set rs_Traza = Nothing
    
Exit Sub
MLocal:
'        Flog.Writeline
'        Flog.Writeline Espacios(Tabulador * 0) & "**********************************************************"
'        Flog.Writeline Espacios(Tabulador * 0) & " Error insertando traza "
'        Flog.Writeline Espacios(Tabulador * 0) & " Error: " & Err.Description
'        Flog.Writeline
'        Flog.Writeline Espacios(Tabulador * 0) & "Ultimo SQL Ejecutado: " & StrSql
'        Flog.Writeline
'        Flog.Writeline Espacios(Tabulador * 0) & "**********************************************************"
'        Flog.Writeline
End Sub


Private Sub Estadisticas()
' ---------------------------------------------------------------------------------------------
' Descripcion: Escribe en el log las estadisticas recolectadas durante el procesamiento.
' Autor      : FGZ
' Fecha      : 18/05/2011
' Ultima Mod :
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Segundos As Double

If Usa_Estadisticas Then
    Flog.writeline
    Flog.writeline "================================================================="
    Flog.writeline "= Estadisticas                                                  ="
    Flog.writeline "================================================================="
    Flog.writeline Espacios(Tabulador * 1) & "Cantidad de Empleados: " & Max_Cabeceras
    Flog.writeline Espacios(Tabulador * 1) & "Cantidad de Conceptos evaluados: " & Max_Conceptos
    Flog.writeline Espacios(Tabulador * 1) & "Cantidad de busquedas: " & Max_Programas
    Flog.writeline Espacios(Tabulador * 1) & "Cantidad de acumuladores: " & Max_Acumuladores

    Flog.writeline
    'Flog.writeline "Accesos a BD: "
    Flog.writeline Espacios(Tabulador * 1) & "Cantidad de Lecturas en BD: " & Cantidad_de_OpenRecordset
    Flog.writeline Espacios(Tabulador * 1) & "Cantidad de Lecturas a His_estructura: " & Cantidad_His_Estructura
    Flog.writeline Espacios(Tabulador * 1) & "Cantidad de Lecturas a detliq: " & Cantidad_Detliq
    Flog.writeline Espacios(Tabulador * 1) & "Cantidad de Lecturas a novgral: " & Cantidad_NovG
    Flog.writeline Espacios(Tabulador * 1) & "Cantidad de Lecturas a acumes: " & Cantidad_Acumes
    Flog.writeline Espacios(Tabulador * 1) & "Cantidad de Lecturas a acu_liq: " & Cantidad_Acu_liq
    Flog.writeline Espacios(Tabulador * 1) & "Cantidad de Lecturas a wf_tpa: " & Cantidad_WF_Tpa
    Flog.writeline Espacios(Tabulador * 1) & "Cantidad de Lecturas a wf_impproarg: " & Cantidad_WF_Impproarg
    Flog.writeline Espacios(Tabulador * 1) & "Cantidad de Lecturas a wf_retroactivo: " & Cantidad_WF_Retro

    Flog.writeline Espacios(Tabulador * 1) & "Cantidad de Select * FROM: " & Cantidad_Select
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)

    Flog.writeline
    Segundos = Round((TiempoFinalProceso - TiempoInicialProceso) / 1000, 0)
    Flog.writeline Espacios(Tabulador * 1) & "Tiempo del proceso (segundos): " & Segundos
    If Max_Cabeceras = 0 Then
        Max_Cabeceras = 1
    End If
    If Max_Conceptos = 0 Then
        Max_Conceptos = 1
    End If
    
    Flog.writeline Espacios(Tabulador * 1) & "Cantidad de Segundos por Empleado: " & Round(Segundos / Max_Cabeceras, 4)
    Flog.writeline Espacios(Tabulador * 1) & "Cantidad de Segundos x concepto x Empleado: " & Round(Segundos / (Max_Conceptos * Max_Cabeceras), 4)
    Flog.writeline "================================================================="
Else
    Flog.writeline
    Flog.writeline "================================================================="
    Flog.writeline "= Estadisticas                                                  ="
    Flog.writeline "================================================================="
    Flog.writeline Espacios(Tabulador * 1) & "Cantidad de Lecturas en BD: " & Cantidad_de_OpenRecordset
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    Segundos = Round((TiempoFinalProceso - TiempoInicialProceso) / 1000, 0)
    Flog.writeline Espacios(Tabulador * 1) & "Tiempo del proceso (segundos): " & Segundos
    If Max_Cabeceras = 0 Then
        Max_Cabeceras = 1
    End If
    Flog.writeline Espacios(Tabulador * 1) & "Cantidad de Segundos por Empleado: " & Round(Segundos / Max_Cabeceras, 4)
    Flog.writeline "================================================================="

End If
End Sub



Private Sub PasaraHistorico()
' ---------------------------------------------------------------------------------------------
' Descripcion: Pasa batch_proceso al historico y carga estadisticas.
' Autor      : FGZ
' Fecha      : 03/06/2011
' Ultima Mod :
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_Batch_Proceso As New ADODB.Recordset
Dim rs_His_Batch_Proceso As New ADODB.Recordset
Dim rs_Est As New ADODB.Recordset

Dim SegundosTotal As Double
Dim PromXempleado As Double
Dim PromXempleadoXconc As Double
Dim TInicio 'As Date
Dim TFin 'As Date


MyBeginTransLiq
        StrSql = "SELECT * FROM batch_proceso WHERE bpronro =" & NroProcesoBatch
        OpenRecordset StrSql, rs_Batch_Proceso

        StrSql = "INSERT INTO His_Batch_Proceso (bpronro,btprcnro,bprcfecha,iduser"
        StrSqlDatos = rs_Batch_Proceso!Bpronro & "," & rs_Batch_Proceso!btprcnro & "," & _
        ConvFecha(rs_Batch_Proceso!bprcfecha) & ",'" & rs_Batch_Proceso!iduser & "'"
        If Not IsNull(rs_Batch_Proceso!bprchora) Then
            StrSql = StrSql & ",bprchora"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprchora & "'"
        End If
        If Not IsNull(rs_Batch_Proceso!bprcempleados) Then
            StrSql = StrSql & ",bprcempleados"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprcempleados & "'"
        End If
        If Not IsNull(rs_Batch_Proceso!bprcfecdesde) Then
            StrSql = StrSql & ",bprcfecdesde"
            StrSqlDatos = StrSqlDatos & "," & ConvFecha(rs_Batch_Proceso!bprcfecdesde)
        End If
        If Not IsNull(rs_Batch_Proceso!bprcfechasta) Then
            StrSql = StrSql & ",bprcfechasta"
            StrSqlDatos = StrSqlDatos & "," & ConvFecha(rs_Batch_Proceso!bprcfechasta)
        End If
        If Not IsNull(rs_Batch_Proceso!bprcestado) Then
            StrSql = StrSql & ",bprcestado"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprcestado & "'"
        End If
        If Not IsNull(rs_Batch_Proceso!bprcparam) Then
            StrSql = StrSql & ",bprcparam"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprcparam & "'"
        End If
        If Not IsNull(rs_Batch_Proceso!bprcprogreso) Then
            StrSql = StrSql & ",bprcprogreso"
            StrSqlDatos = StrSqlDatos & "," & rs_Batch_Proceso!bprcprogreso
        End If
        If Not IsNull(rs_Batch_Proceso!bprcfecfin) Then
            StrSql = StrSql & ",bprcfecfin"
            StrSqlDatos = StrSqlDatos & "," & ConvFecha(rs_Batch_Proceso!bprcfecfin)
        End If
        If Not IsNull(rs_Batch_Proceso!bprchorafin) Then
            StrSql = StrSql & ",bprchorafin"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprchorafin & "'"
        End If
        If Not IsNull(rs_Batch_Proceso!bprctiempo) Then
            StrSql = StrSql & ",bprctiempo"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprctiempo & "'"
        End If
        If Not IsNull(rs_Batch_Proceso!Empnro) Then
            StrSql = StrSql & ",empnro"
            StrSqlDatos = StrSqlDatos & "," & rs_Batch_Proceso!Empnro
        End If
        If Not IsNull(rs_Batch_Proceso!bprcPid) Then
            StrSql = StrSql & ",bprcPid"
            StrSqlDatos = StrSqlDatos & "," & rs_Batch_Proceso!bprcPid
        End If
        If Not IsNull(rs_Batch_Proceso!bprcfecInicioEj) Then
            StrSql = StrSql & ",bprcfecInicioEj"
            StrSqlDatos = StrSqlDatos & "," & ConvFecha(rs_Batch_Proceso!bprcfecInicioEj)
        End If
        If Not IsNull(rs_Batch_Proceso!bprcfecFinEj) Then
            StrSql = StrSql & ",bprcfecFinEj"
            StrSqlDatos = StrSqlDatos & "," & ConvFecha(rs_Batch_Proceso!bprcfecFinEj)
        End If
        If Not IsNull(rs_Batch_Proceso!bprcUrgente) Then
            StrSql = StrSql & ",bprcUrgente"
            StrSqlDatos = StrSqlDatos & "," & rs_Batch_Proceso!bprcUrgente
        End If
        If Not IsNull(rs_Batch_Proceso!bprcHoraInicioEj) Then
            StrSql = StrSql & ",bprcHoraInicioEj"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprcHoraInicioEj & "'"
        End If
        If Not IsNull(rs_Batch_Proceso!bprchorafinej) Then
            StrSql = StrSql & ",bprcHoraFinEj"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprchorafinej & "'"
        End If
        'If Not IsNull(rs_Batch_Proceso!bpronroori) Then
        '    StrSql = StrSql & ",bpronroori"
        '    StrSqlDatos = StrSqlDatos & "," & rs_Batch_Proceso!bpronroori
        'End If
        StrSql = StrSql & ") VALUES (" & StrSqlDatos & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        'Reviso que haya copiado
        StrSql = "SELECT * FROM His_batch_proceso WHERE bpronro =" & NroProcesoBatch
        OpenRecordset StrSql, rs_His_Batch_Proceso
        If Not rs_His_Batch_Proceso.EOF Then
            ' Borro de Batch_proceso
            StrSql = "DELETE FROM Batch_Proceso WHERE bpronro = " & NroProcesoBatch
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
    
    
        'Actualizo las estadisticas
        If Usa_Estadisticas Then
        
            'Algunos calculos de promedios en segundos
            TInicio = rs_Batch_Proceso!bprcfecInicioEj & " " & rs_Batch_Proceso!bprcHoraInicioEj
            TFin = rs_Batch_Proceso!bprcfecFinEj & " " & rs_Batch_Proceso!bprchorafinej
            
            'TInicio = rs_Batch_Proceso!bprcHoraInicioEj
            'TFin = rs_Batch_Proceso!bprchorafinej
            
            SegundosTotal = DateDiff("s", TInicio, TFin)
            If SegundosTotal = 0 Then
                SegundosTotal = 1
            End If
            
            'FGZ - 17/07/2015 -------------------------------
            If Max_Cabeceras = 0 Then
                Max_Cabeceras = 1
            End If
            If Max_Conceptos = 0 Then
                Max_Conceptos = 1
            End If
            'FGZ - 17/07/2015 -------------------------------
            
            'FGZ - 09/09/2014 ---------------------------------------
            'PromXempleado = Max_Cabeceras / SegundosTotal
            PromXempleado = SegundosTotal / Max_Cabeceras
            PromXempleadoXconc = PromXempleado / Max_Conceptos
            'FGZ - 09/09/2014 ---------------------------------------
        
            StrSql = "SELECT * FROM His_batch_proceso_est WHERE bpronro =" & NroProcesoBatch
            OpenRecordset StrSql, rs_Est
            If rs_Est.EOF Then
            
                StrSql = "INSERT INTO His_Batch_Proceso_est (bpronro"
                StrSqlDatos = rs_Batch_Proceso!Bpronro
                If Not IsNull(rs_Batch_Proceso!bpronroori) Then
                    StrSql = StrSql & ",bpronroori"
                    StrSqlDatos = StrSqlDatos & "," & rs_Batch_Proceso!bpronroori
                End If
                StrSql = StrSql & ",version"
                StrSqlDatos = StrSqlDatos & ",'" & Version & "'"
                StrSql = StrSql & ",debug"
                StrSqlDatos = StrSqlDatos & "," & CInt(USA_DEBUG)
                StrSql = StrSql & ",andet"
                StrSqlDatos = StrSqlDatos & "," & CInt(HACE_TRAZA)
                StrSql = StrSql & ",bdlocal"
                StrSqlDatos = StrSqlDatos & "," & CInt(True)
                StrSql = StrSql & ",cantlectbd"
                StrSqlDatos = StrSqlDatos & "," & Cantidad_de_OpenRecordset
                StrSql = StrSql & ",cantemp"
                StrSqlDatos = StrSqlDatos & "," & Max_Cabeceras
                StrSql = StrSql & ",cantconc"
                StrSqlDatos = StrSqlDatos & "," & Max_Conceptos
                StrSql = StrSql & ",cantacu"
                StrSqlDatos = StrSqlDatos & "," & Max_Acumuladores
                StrSql = StrSql & ",cantbusq"
                StrSqlDatos = StrSqlDatos & "," & Max_Programas
                StrSql = StrSql & ",segundos"
                StrSqlDatos = StrSqlDatos & "," & SegundosTotal
                StrSql = StrSql & ",promemp"
                StrSqlDatos = StrSqlDatos & "," & PromXempleado
                'FGZ - 17/09/2014 -----------------------------
                StrSql = StrSql & ",cantbusqint"
                StrSqlDatos = StrSqlDatos & "," & Cantidad_BusI
                StrSql = StrSql & ",cantbusqnovg"
                StrSqlDatos = StrSqlDatos & "," & Cantidad_NovG
                StrSql = StrSql & ",cantbusqnove"
                StrSqlDatos = StrSqlDatos & "," & Cantidad_NovE
                StrSql = StrSql & ",cantbusqnovi"
                StrSqlDatos = StrSqlDatos & "," & Cantidad_NovI
                StrSql = StrSql & ",cantconcaju"
                StrSqlDatos = StrSqlDatos & "," & Cantidad_NovA
                'FGZ - 17/09/2014 -----------------------------
                StrSql = StrSql & ") VALUES (" & StrSqlDatos & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
        End If
        ' -----------------------------------------------------------------------------------
MyCommitTransLiq


'Cierro y libero
If rs_Batch_Proceso.State = adStateOpen Then rs_Batch_Proceso.Close
If rs_His_Batch_Proceso.State = adStateOpen Then rs_His_Batch_Proceso.Close
If rs_Est.State = adStateOpen Then rs_Est.Close

Set rs_Batch_Proceso = Nothing
Set rs_His_Batch_Proceso = Nothing
Set rs_Est = Nothing

End Sub


Public Sub Desmarcar_VentaVac(ByVal pronro As Long, ByVal Tercero As Long)
' -----------------------------------------------------------------------------------
' Descripcion: Desmarca las ventas de vacaciones marcadaas por el proceso.
' Autor: EAM
' Fecha: 21/08/2013
' Ultima Modificacion:
' -----------------------------------------------------------------------------------

On Error GoTo ME_Local
        
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 3) & "Desmarcando Venta de Vacaciones"
    End If
    
    StrSql = "DELETE FROM vacvendidos WHERE pronro= " & pronro & " AND ternro= " & Tercero
    objConn.Execute StrSql, , adExecuteNoRecords

Fin:
Exit Sub

ME_Local:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 3) & " Desmarcado de Venta de Vacaciones" & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline
    GoTo Fin
End Sub


Public Sub Desmarcar_Paros(ByVal pronro As Long, ByVal Tercero As Long)
' -----------------------------------------------------------------------------------
' Descripcion: Desmarca llos detalles de paros sindicales marcados por el proceso.
' Autor: FGZ
' Fecha: 15/05/2014
' Ultima Modificacion:
' -----------------------------------------------------------------------------------
On Error GoTo ME_Local
        
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 3) & "Desmarcando detalle de Paros Sindicales"
    End If
    
    StrSql = "UPDATE pardet SET pronro = null "
    StrSql = StrSql & " WHERE pronro = " & pronro
    StrSql = StrSql & " AND ternro = " & Tercero
    objConn.Execute StrSql, , adExecuteNoRecords
    

Fin:
Exit Sub

ME_Local:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 3) & " Desmarcado de Paros Sindicales " & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline
    GoTo Fin
End Sub


Public Sub Eliminar_Distribucion(ByVal Nro_Ter As Long, ByVal Nro_Proceso As Long)
' -----------------------------------------------------------------------------------
' Descripcion: Elimina las distribuciones generadas en el proceso.
' Autor: FGZ
' Fecha: 14/12/2013
' Ultima Modificacion:
' -----------------------------------------------------------------------------------
Dim rs_dist As New ADODB.Recordset

    On Error GoTo ME_Local
    
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 3) & "ELIMINANDO DISTRIBUCIONES GENERADAS POR EL PROCESO"
    End If

    StrSql = "DELETE FROM concepto_dist " & _
            " WHERE ternro = " & Nro_Ter & _
            " AND pronro =" & Nro_Proceso
    objConn.Execute StrSql, , adExecuteNoRecords
    
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 1) & "ELIMINACION DE DISTRIBUCIONES REALIZADA EXITOSAMENTE"
    End If

Fin:
If rs_dist.State = adStateOpen Then rs_dist.Close
Exit Sub

ME_Local:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 3) & " Eliminacion de Distribucion " & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline
    GoTo Fin
End Sub



Public Function CalculoEstadisticoActivo(ByVal tipo As Long) As Boolean
' -----------------------------------------------------------------------------------
' Descripcion: Revisa si esta activa la recoleccion de datos estadisticos.
' Autor: FGZ
' Fecha: 21/11/2013
' Ultima Modificacion:
' -----------------------------------------------------------------------------------
Dim Activo As Boolean
Dim rs As New ADODB.Recordset

On Error GoTo ME_Local
        
    Activo = False
        
    StrSql = "Select estadistica FROM batch_tipproc WHERE btprcnro = 3"
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        Activo = CBool(rs!estadistica)
    End If

Fin:
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 3) & "Calculo Estadístico Activo? " & Activo
    End If
    CalculoEstadisticoActivo = Activo
    Exit Function

ME_Local:
    'No activo el calculo estadistico
    GoTo Fin
End Function

Public Function HayNovDist() As Boolean
' -----------------------------------------------------------------------------------
' Descripcion: Revisa si esta activa la recoleccion de datos estadisticos.
' Autor: FGZ
' Fecha: 21/11/2013
' Ultima Modificacion:
' -----------------------------------------------------------------------------------
Dim Activo As Boolean
Dim rs As New ADODB.Recordset

On Error GoTo ME_Local
        
    Activo = False
        
    StrSql = "Select * FROM nov_dist"
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        Activo = True
    End If

Fin:
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 3) & "Novedades con distribución Individual " & Activo
    End If
    HayNovDist = Activo
    Exit Function

ME_Local:
    'No activo el calculo de novedades con distribucion individual
    GoTo Fin
End Function
