Attribute VB_Name = "GenNovedadesGti"
Option Explicit

'Const Version = 2.01    'Version Inicial
'Const FechaVersion = "14/02/2006"

'Const Version = 2.02    'Modificacion en el sql inicial de la politica 900
'Const FechaVersion = "14/02/2006"

'Const Version = 2.03
'Const FechaVersion = "01/03/2006"

'Const Version = "2.04"
'Const FechaVersion = "17/05/2007"
''Modificaciones: FGZ
''    sub Novedades_GTI_Standar  : Modificacion, pasé la depuración un nivel mas arriba, es decir, en la politica 900
''    sub Novedades_GTI_MAZUL    : Modificacion, pasé la depuración un nivel mas arriba, es decir, en la politica 900
''    sub Novedades_GTI_Expo     : Modificacion, pasé la depuración un nivel mas arriba, es decir, en la politica 900

'===============================================================================
'Const Version = "3.00"
'Const FechaVersion = "04/06/2007"
''Modificaciones: FGZ
''      Mejoras generales de performance.

'Const Version = "3.01"
'Const FechaVersion = "16/04/2008"
''Modificaciones: FGZ
''   Modulo politicas: sub Cargar_DetallePoliticas.
''           Se cambió en el where el <> '' por IS NOT NULL

'Const Version = "3.02"
'Const FechaVersion = "26/06/2008"
''Modificaciones: FGZ
''   Openconnection: Cambio de schema para Oracle.

'Const Version = "3.03"
'Const FechaVersion = "14/10/2008"
''Modificaciones: FGZ
''   Se agregó un parametro para poder generar las novedades aprobadas.

'Const Version = "3.04"
'Const FechaVersion = "21/01/2009"
''Modificaciones: FGZ
''   Encriptacion de string de conexion


'Const Version = "3.05"
'Const FechaVersion = "31/03/2010"
''Modificaciones: FGZ
''    politica 900 - sub Novedades_GTI_Expo     : Modificacion, cuando el parametro era dias
''                           antes dividia siempre por 8 hs horas
''                           horas divide por las horas obligatorias del 1er dia del subturno, del turno del empleado.

'Const Version = "3.06"
'Const FechaVersion = "07/05/2010"
''Modificaciones: FGZ
''    politica 900 - sub Novedades_GTI_Expo: Redondeo general

''----------------------------
'Const Version = "5.00"
'Const FechaVersion = "15/06/2010"
''Modificaciones: FGZ
''    Control por entradas fuera de termino.
''           Antes
''               cuando se queria procesar algo en una fecha que caia en un periodo cerrado no se procesaba.
''           Ahora
''               Se puede reprocear un periodo cerrado solo cuando se aprueba una entrada fuera de termino.
''               Para ese reprocesamiento solo se tendrá en cuenta todas las entradas fuera de termino aprobadas.

'Const Version = "5.01"
'Const FechaVersion = "01/10/2010"
''Modificaciones: FGZ
''    Control por entradas fuera de termino. Estaba levantando mal un parametro

'Const Version = "5.02"
'Const FechaVersion = "03/06/2011"
''Modificaciones: FGZ
''    Se recompiló por un problema en el modelo de politicas.

'Const Version = "5.03"
'Const FechaVersion = "21/06/2011"
''Modificaciones: FGZ
''    Se agregaó el control de firmas a las novedades horarias
''       Se modifico:
''           Buscar_Turno
''           Buscar_Turno_nuevo

'Const Version = "5.04"
'Const FechaVersion = "18/07/2011"
''Modificaciones: FGZ
''    Se cambió el objeto conexion para el progreso



'Const Version = "5.05"
'Const FechaVersion = "14/02/2013"
''Modificaciones: FGZ
''    No estaba levantando la semilla de encriptacion. c_seed = ArrParametros(2)
''    EAM- Se generó el ejecutable y se formalizo la entrega con el caso (18436)
''

Const Version = "5.06"
Const FechaVersion = "19/12/2013"
'Modificaciones: Margiotta, Emanuel (CAS-22808 - SGS - Distribución Contable)
'   Politica 900: Se generó la version 5. Genera novedades por la distribución del AD (tabla->gti_desgldiario)
'               y la diferencia sin desglosar la inserta como una novedad sin desglose
'    EAM- Se generó el ejecutable y se formalizo la entrega con el caso (18436)

'---------------------------------------------------------------------------
'---------------------------------------------------------------------------

Global G_traza As Boolean
Global fec_proc As Integer

Global diatipo As Byte
Global ok As Boolean
Global esFeriado As Boolean
Global hora_desde As String
Global fecha_desde As Date
Global fecha_hasta As Date
Global Hora_desde_aux As String
Global hora_hasta As String
Global Hora_Hasta_aux As String
Global No_Trabaja_just As Boolean
Global nro_jus_ent As Long
Global nro_jus_sal As Long
Global Total_horas As Single
Global Tdias As Integer
Global Thoras As Integer
Global Tmin As Integer
Global Cod_justificacion1 As Long
Global Cod_justificacion2 As Long

Global Horas_Oblig As Single
Global Existe_Reg As Boolean
Global Forma_embudo  As Boolean

Global tiene_turno As Boolean
Global Nro_Turno As Long
Global Tipo_Turno As Integer

Global Tiene_Justif As Boolean
Global Nro_Justif As Long
Global justif_turno As Boolean
Global p_turcomp As Boolean
Global Nro_Grupo As Long
Global Nro_fpgo As Integer
Global Fecha_Inicio As Date
Global P_Asignacion  As Boolean
Global Trabaja     As Boolean ' Indica si trabaja para ese dia
Global Orden_Dia As Integer
Global Nro_Dia As Integer
Global Nro_Subturno As Integer
Global Dia_Libre As Boolean
Global Dias_trabajados As Integer
Global Dias_laborables As Integer

Global Aux_Tipohora As Integer
Global aux_TipoDia As Integer
Global Sigo_Generando As Boolean

Global Hora_Tol As String
Global Fecha_Tol As Date
Global hora_toldto As String
Global fecha_toldto As Date

Global Usa_Conv  As Boolean

Global tol As String

Global Cant_emb As Integer
Global toltemp As String
Global toldto As String
Global acumula As Boolean
Global acumula_dto As Boolean
Global acumula_temp As Boolean
Global convenio As Long

Global tdias_oblig As Single
Global objFeriado As New Feriado



Public Sub Main()
'-----------------------------------------------------------------------
'Procedimiento principal
'
'-----------------------------------------------------------------------
Dim FechaDesde As Date
Dim FechaHasta As Date
Dim Fecha As Date
Dim objRs As New ADODB.Recordset
Dim objrsEmpleado As New ADODB.Recordset
'Dim NroProceso As Long
Dim strcmdLine  As String
Dim Progreso As Single
Dim CEmpleadosAProc As Integer
Dim IncPorc As Single
Dim ListaPar
Dim rs_Batch_Proceso As New ADODB.Recordset
Dim rs_His_Batch_Proceso As New ADODB.Recordset
Dim PID As String
Dim ArrParametros

    
    strcmdLine = Command()
    ArrParametros = Split(strcmdLine, " ", -1)
    If UBound(ArrParametros) > 1 Then
        If IsNumeric(ArrParametros(0)) Then
            NroProceso = ArrParametros(0)
            Etiqueta = ArrParametros(1)
            EncriptStrconexion = CBool(ArrParametros(2))
            c_seed = ArrParametros(2)
        Else
            Exit Sub
        End If
    Else
        If UBound(ArrParametros) > 0 Then
            If IsNumeric(ArrParametros(0)) Then
                NroProceso = ArrParametros(0)
                Etiqueta = ArrParametros(1)
            Else
                Exit Sub
            End If
        Else
            If IsNumeric(strcmdLine) Then
                NroProceso = strcmdLine
            Else
                Exit Sub
            End If
        End If
    End If
    
    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(PathFLog & "ACUNOV" & "-" & NroProceso & ".log", True)
    
    'Abro la conexion
    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    On Error Resume Next
    OpenConnection strconexion, objConnProgreso
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    
    'Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaVersion
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-------------------------------------------------"
    Flog.writeline
    'Flog.writeline "Inicio Proceso :" & NroProceso & " " & Now
    
    'FGZ - 05/08/2009 --------- Control de versiones ------
    Version_Valida = ValidarV(Version, 5, TipoBD)
    If Not Version_Valida Then
        'Actualizo el progreso
        MyBeginTrans
            StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Error de Version', bprcpid = " & PID & " WHERE bpronro = " & NroProceso
            objConnProgreso.Execute StrSql, , adExecuteNoRecords
        MyCommitTrans
        Flog.writeline
        GoTo Final
    End If
    'FGZ - 05/08/2009 --------- Control de versiones ------
    
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprcpid = " & PID & ", bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
    
    On Error GoTo ce
    
    Aprobadas = False
    
    
    StrSql = " SELECT batch_proceso.IdUser, batch_proceso.bprcparam, gti_procacum.gpadesde,gti_procacum.gpahasta  FROM batch_proceso "
    StrSql = StrSql & " INNER JOIN batch_procacum ON batch_procacum.bpronro = batch_proceso.bpronro "
    StrSql = StrSql & " INNER JOIN gti_cab ON  gti_cab.gpanro = batch_procacum.gpanro "
    StrSql = StrSql & " INNER JOIN gti_procacum ON batch_procacum.gpanro = gti_procacum.gpanro "
    StrSql = StrSql & " WHERE batch_proceso.bpronro = " & NroProceso
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        FechaDesde = objRs!gpadesde
        FechaHasta = objRs!gpahasta
        Flog.writeline Espacios(Tabulador * 1) & "Parametros: "
        Flog.writeline Espacios(Tabulador * 1) & "Usuario: " & objRs!IdUser
        Flog.writeline Espacios(Tabulador * 1) & "Desde: " & FechaDesde
        Flog.writeline Espacios(Tabulador * 1) & "Hasta: " & FechaHasta
        Flog.writeline Espacios(Tabulador * 1) & "bprcparam: " & objRs!bprcparam
        If Not EsNulo(objRs!bprcparam) Then
            If InStr(1, objRs!bprcparam, ".") <> 0 Then
                ListaPar = Split(objRs!bprcparam, ".", -1)
                depurar = IIf(IsNumeric(ListaPar(0)), CBool(ListaPar(0)), False)
                If UBound(ListaPar) > 0 Then
                    Aprobadas = CBool(ListaPar(1))
                End If
                If UBound(ListaPar) > 1 Then
                    If IsNumeric(ListaPar(2)) Then
                        ReprocesarFT = IIf(IsNumeric(ListaPar(2)), CBool(ListaPar(2)), False)
                    Else
                        ReprocesarFT = False
                    End If
                Else
                    ReprocesarFT = False
                End If
            Else
                depurar = False
                ReprocesarFT = False
            End If
        Else
            depurar = False
            ReprocesarFT = False
        End If
    Else
        Flog.writeline
        Flog.writeline "No hay nada para procesar."
        Flog.writeline
        Flog.writeline "SQL: " & StrSql
    End If
    
    'FGZ - Mejoras ----------
    Call Inicializar_Globales
    Call Cargar_PoliticasEstructuras(FechaDesde)
    Call Cargar_PoliticasIndividuales
    'FGZ - Mejoras ----------
    
    Flog.writeline Espacios(Tabulador * 0) & "Inicio    :" & Now
    MyBeginTrans
    
    Call Politica(900)
    
    StrSql = "DELETE FROM batch_empleado WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
    StrSql = "DELETE FROM Batch_Procacum WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
    StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
    
    
        ' -----------------------------------------------------------------------------------
    'FGZ - 22/09/2003
    'Copio el proceso en el historico de batch_proceso y lo borro de batch_proceso
        StrSql = "SELECT * FROM batch_proceso WHERE bpronro =" & NroProceso
        OpenRecordset StrSql, rs_Batch_Proceso

        
        StrSql = "INSERT INTO His_Batch_Proceso (bpronro,btprcnro,bprcfecha,iduser"
        StrSqlDatos = rs_Batch_Proceso!bpronro & "," & rs_Batch_Proceso!btprcnro & "," & _
                 ConvFecha(rs_Batch_Proceso!bprcfecha) & ",'" & rs_Batch_Proceso!IdUser & "'"
        
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
        If Not IsNull(rs_Batch_Proceso!empnro) Then
            StrSql = StrSql & ",empnro"
            StrSqlDatos = StrSqlDatos & "," & rs_Batch_Proceso!empnro
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
        If Not IsNull(rs_Batch_Proceso!bprcHoraFinEj) Then
            StrSql = StrSql & ",bprcHoraFinEj"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprcHoraFinEj & "'"
        End If

        StrSql = StrSql & ") VALUES (" & StrSqlDatos & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        
        'Reviso que haya copiado
        StrSql = "SELECT * FROM His_batch_proceso WHERE bpronro =" & NroProceso
        OpenRecordset StrSql, rs_His_Batch_Proceso
        
        If Not rs_His_Batch_Proceso.EOF Then
            ' Borro de Batch_proceso
            StrSql = "DELETE FROM Batch_Proceso WHERE bpronro = " & NroProceso
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
    
        If rs_Batch_Proceso.State = adStateOpen Then rs_Batch_Proceso.Close
        If rs_His_Batch_Proceso.State = adStateOpen Then rs_His_Batch_Proceso.Close
    ' FGZ - 22/09/2003
    ' -----------------------------------------------------------------------------------
    MyCommitTrans

Final:
    If objConn.State = adStateOpen Then objConn.Close
    Set objConn = Nothing
    If objConnProgreso.State = adStateOpen Then objConnProgreso.Close
    Set objConnProgreso = Nothing

    Flog.writeline
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "Fin       :" & Now
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "---------------------------------------------------------------------------------"
    Flog.writeline "Cantidad de Lecturas en BD          : " & Cantidad_de_OpenRecordset
    Flog.writeline "Cantidad de llamadas a politicas    : " & Cantidad_Call_Politicas
'    Flog.writeline "Cantidad de llamadas a EsFeriado    : " & Cantidad_Feriados
'    Flog.writeline "Cantidad de llamadas a BuscarTurno  : " & Cantidad_Turnos
'    Flog.writeline "Cantidad de llamadas a BuscarDia    : " & Cantidad_Dias
'    Flog.writeline
'    Flog.writeline "Cantidad de dias procesados         : " & Cantidad_Empl_Dias_Proc
    Flog.writeline Espacios(Tabulador * 0) & "---------------------------------------------------------------------------------"
    Flog.Close
    
    Exit Sub

ce:
    MyRollbackTrans
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 0) & "Error en Proceso : " & NroProceso & " " & Now
    Flog.writeline Espacios(Tabulador * 0) & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline
    Flog.writeline "Fin Proceso :" & NroProceso & " " & Now
    
    MyBeginTrans
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProceso
        objConn.Execute StrSql, , adExecuteNoRecords
    MyCommitTrans

    If objConn.State = adStateOpen Then objConn.Close
    Set objConn = Nothing
    If objConnProgreso.State = adStateOpen Then objConnProgreso.Close
    Set objConnProgreso = Nothing
End Sub

Public Sub Inicializar_Globales()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que Carga los array globales.
' Autor      : FGZ
' Fecha      : 17/05/2007
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

    'Politicas de alcance global
    Call Cargar_PoliticasGlobales
    Call Cargar_DetallePoliticas
    
    'FGZ - 15/06/2011
    Call ParametrosGlobales
    
End Sub



