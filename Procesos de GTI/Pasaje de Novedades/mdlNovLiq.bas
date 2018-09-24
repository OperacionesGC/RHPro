Attribute VB_Name = "mdlNovLiq"
Option Explicit

'Const Version = 1.01     'Version Inicial
'Const FechaVersion = "15/02/2006"

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
'Modificaciones: FGZ
'   Openconnection: cambio de schema para Oracle.

'Politica 2000: Exportacion de novedades
'       Nueva version

'Const Version = "3.03"
'Const FechaVersion = "31/07/2008"
''Modificaciones: FAF
''   Se modifico Politica2000Expo para que permita leer el string de conexion desde un archivo


'Const Version = "3.04"
'Const FechaVersion = "14/10/2008"
''Modificaciones: FGZ
''   Se modifico Politica2000TXT (CUSTOM CARGIL)
''           Se modificó el nombre del archivo que se genera


'Const Version = "3.05"
'Const FechaVersion = "21/10/2008"
''Modificaciones: FGZ
''   Se modifico Politica2000TXT (CUSTOM CARGIL)
''           Se modificó el primer campo que se exporta

'Const Version = "3.06"
'Const FechaVersion = "27/10/2008"
''Modificaciones: FGZ
''   Se modifico Politica2000TXT (CUSTOM CARGIL)
''           Se modificó el orden en que se exporta (estaba por tipode hora y ahora quieren por legajo)

'Const Version = "3.07"
'Const FechaVersion = "02/12/2008"
''Modificaciones: FGZ
''   Se modifico Politica2000TXT (CUSTOM CARGIL)
''           Se puso el valor en valor absoluto. El signo va al final


'Const Version = "3.08"
'Const FechaVersion = "21/01/2009"
''Modificaciones: FGZ
''   Encriptacion de string de conexion

'Const Version = "5.00"
'Const FechaVersion = "21/06/2011"
''Modificaciones: FGZ
''    Se agregaó el control de firmas a las novedades horarias
''       Se modifico:
''           Buscar_Turno
''           Buscar_Turno_nuevo
''
''       Ademas se le agregó control de firmas sobre las novedades de liquidacon a exportar.
''       El proceso recibe un parametro mas con el iduser que quien autorizaría la novedad a crear (el parametro se pide en la ventana de pasaje de novedades).
''       Si el usuario que disparó el proceso es fin de firma de circuito de autorizacion de novedades de liquidacion ==>
''       La novedad queda autorizada. Sino queda autoizada por quien disparó el proceso y pendiente para el usuario indicado.


'Const Version = "5.01"
'Const FechaVersion = "30/06/2011"
''Modificaciones: FGZ
''    Se agregaó la actualizacion del progreso (solo a versiones 1 y 5 de la pol 2000)

'Const Version = "5.02"
'Const FechaVersion = "06/07/2011"
''Modificaciones: FGZ
''    Se agregaó la actualizacion del progreso (solo a versiones 1 y 5 de la pol 2000)

'Const Version = "5.03"
'Const FechaVersion = "04/10/2012"
'Modificaciones: FGZ
'    Politica 2000 - Se corrigió un problema con el getlastidentity para Oracle
'               CAS-16910- TATA - BUG - pasaje de novedades desde GTI a LIQ

'Const Version = "5.04"
'Const FechaVersion = "11/12/2012"
'Modificaciones: LED
'    Politica 2000 - Se agrego version 7
'               CAS-17445 - Akzo Nobel - Exportacion Novedades a LIQ

'Const Version = "5.05"
'Const FechaVersion = "21/03/2013"
'Modificaciones: LED
'    No estaba levantando la semilla de encriptacion. c_seed = ArrParametros(2)
'               CAS-17445 - Akzo Nobel - Exportacion Novedades a LIQ

'Const Version = "5.06"
'Const FechaVersion = "03/06/2013"
'Modificaciones: LED
'    Modificacion en el nombre del archivo generado en la Politica 2000 - Se agrego version 7, se puso fecha y hora al archivo.
' CAS-17445 - AkzoNobel - Exportacion Novedades a LIQ - Modificaciones

'Const Version = "5.07"
'Const FechaVersion = "06/06/2013"
''Modificaciones: LED
''    Modificacion Politica 2000 - version 7, se corrigieron longitudes de cambios.
'' CAS-17814 - AKZO - QA - Bug Exportacion Novedades a LIQ

'Const Version = "5.08"
'Const FechaVersion = "19/12/2013"
'Modificaciones: EAM (CAS-22808 - SGS - Distribución Contable)
'    Modificacion Politica 2000 - version 8, se corrigieron longitudes de cambios.

'Const Version = "5.09"
'Const FechaVersion = "26/12/2013"
''Modificaciones: LED CAS-22170 - CARGILL - Nueva Exportacion de Novedades GTI
''   Politica 2000 - version 8, Pasaje de Novedades - Politica2000ArchivoTXTV2.


'Const Version = "5.10"
'Const FechaVersion = "30/05/2014"
''Modificaciones: EAM (CAS-22808 - SGS - Distribución Contable)
''    Modificacion Politica 2000 - version 8, se hizo configurable el modelo de asiento  utilizar.
''    Modificacion Politica 2000 - version 8, se agrego la funcionalidad para que controle el alcance de la política ya que el proceso no lo tiene en cuenta


'Const Version = "5.11"
'Const FechaVersion = "02/10/2014"
''Modificaciones: FGZ - LED - CAS-27048 - SGS - Circuito de Autorizacion - Hito 1
''    Modificacion Politica 2000 - version 8, se agrego autofirma a las novedades, segun corresponda.
''                               - version 10, se agregó una nueva version que pasa novedades con la vigencia del proceso de acumulacion parcial. Copia del modelo 1.


'Const Version = "5.12"
'Const FechaVersion = "01/12/2014"
'Modificaciones: Gonzalez Nicolás - CAS-27048 - SGS - Circuito de Autorizacion - Bug de custom en GTI - Versión 8: Cambio del método en resolver la firma de la novedad.
'Modificaciones: FGZ - CAS-27048 - SGS - Circuito de Autorizacion - Hito 1
'    Modificacion Politica 2000 - Se hizo configurable la politica
'                               - version 10, cuando pasa las novedades, el tipo de motivo ahora es configurable. Si no se especifica ==> el default es 4.

'Const Version = "5.13"
'Const FechaVersion = "08/04/2014"
'Modificaciones: Sebastian Stremel - CAS-28352 - Salto Grande - GTI – Pasaje de Novedades - Se creo la version 11 para la politica 2000 dicha version es para CTM
'Modificaciones:

Const Version = "5.14"
Const FechaVersion = "11/02/2016"
'Modificaciones: Sebastian Stremel - CAS-28352 - Salto Grande - GTI – Pasaje de Novedades [Entrega 2] - Se modifico la version 11 de la politica 2000, la misma busca el numero de doc principal del empleado en lugar del legajo
'Modificaciones:

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
Global P_turcomp As Boolean
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



Sub Main()
'-----------------------------------------------------------------------
'Procedimiento principal
'
'-----------------------------------------------------------------------
Dim FechaDesde As Date
Dim FechaHasta As Date
Dim strcmdLine As String
Dim rs_Batch_Proceso As New ADODB.Recordset
Dim rs_His_Batch_Proceso As New ADODB.Recordset
Dim Nombre_Arch As String
Dim ListaPar
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
  
    Nombre_Arch = PathFLog & "NovLiq" & "-" & NroProceso & ".log"
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
  
    'Abro la conexion
    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If
  
    'Abro la conexion
    On Error Resume Next
    OpenConnection strconexion, objConnProgreso
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If
  
    On Error GoTo ce
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaVersion
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-------------------------------------------------"
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "Inicio    :" & Now
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprcpid = " & PID & ", bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objConnProgreso.Execute StrSql, , adExecuteNoRecords
    
    
'    StrSql = " SELECT batch_proceso.IdUser, batch_proceso.bprcparam, gti_procacum.gpadesde,gti_procacum.gpahasta  FROM batch_proceso "
'    StrSql = StrSql & " INNER JOIN batch_procacum ON batch_procacum.bpronro = batch_proceso.bpronro "
'    StrSql = StrSql & " INNER JOIN gti_cab ON  gti_cab.gpanro = batch_procacum.gpanro "
'    StrSql = StrSql & " INNER JOIN gti_procacum ON batch_procacum.gpanro = gti_procacum.gpanro "
'    StrSql = StrSql & " WHERE batch_proceso.bpronro = " & NroProceso
    
    StrSql = " SELECT batch_proceso.IdUser, batch_proceso.bprcparam, gti_procacum.gpadesde,gti_procacum.gpahasta " & _
             " From batch_proceso " & _
             " INNER JOIN Batch_Procacum ON Batch_Procacum.bpronro = batch_proceso.bpronro " & _
             " INNER JOIN gti_Procacum ON Batch_Procacum.gpanro = gti_Procacum.gpanro " & _
             " Where batch_proceso.bpronro = " & NroProceso
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        
        
        User_Proceso = objRs!IdUser
        FechaDesde = objRs!gpadesde
        FechaHasta = objRs!gpahasta
    
        Flog.writeline Espacios(Tabulador * 1) & "Parametros: "
        Flog.writeline Espacios(Tabulador * 1) & "Usuario: " & objRs!IdUser
        Flog.writeline Espacios(Tabulador * 1) & "Desde: " & FechaDesde
        Flog.writeline Espacios(Tabulador * 1) & "Hasta: " & FechaHasta
        Flog.writeline Espacios(Tabulador * 1) & "bprcparam: " & objRs!bprcparam
        If Len(objRs!bprcparam) > 1 Then
            If Not EsNulo(objRs!bprcparam) Then
                If InStr(1, objRs!bprcparam, ".") <> 0 Then
                    ListaPar = Split(objRs!bprcparam, ".", -1)
                    depurar = IIf(IsNumeric(ListaPar(0)), CBool(ListaPar(0)), False)
                    'FGZ - 21/06/2011 -------
                    If UBound(ListaPar) >= 1 Then
                        Firma_User_Destino = ListaPar(1)
                    Else
                        Firma_User_Destino = ""
                    End If
                Else
                    depurar = False
                    Firma_User_Destino = ""
                End If
            Else
                depurar = False
                Firma_User_Destino = ""
            End If
        Else
            depurar = False
            Firma_User_Destino = ""
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
    
    
    Call Politica(2000)
    
    StrSql = "DELETE FROM batch_empleado WHERE bpronro = " & NroProceso
    objConnProgreso.Execute StrSql, , adExecuteNoRecords
    StrSql = "DELETE FROM Batch_Procacum WHERE bpronro = " & NroProceso
    objConnProgreso.Execute StrSql, , adExecuteNoRecords
    StrSql = "UPDATE batch_proceso SET bprcprogreso = 100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
    objConnProgreso.Execute StrSql, , adExecuteNoRecords
        
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
            'objConn.Execute StrSql, , adExecuteNoRecords
        End If
    
        If rs_Batch_Proceso.State = adStateOpen Then rs_Batch_Proceso.Close
        If rs_His_Batch_Proceso.State = adStateOpen Then rs_His_Batch_Proceso.Close
    ' FGZ - 22/09/2003
    ' -----------------------------------------------------------------------------------
        
        
FIN:
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
    Flog.writeline " ------------------------------"
    Flog.writeline Err.Description
    Flog.writeline
    Flog.writeline "Ultimo SQL: " & StrSql
    Flog.writeline
    
    StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProceso
    objConnProgreso.Execute StrSql, , adExecuteNoRecords

    GoTo FIN
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






