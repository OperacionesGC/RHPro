Attribute VB_Name = "RHProAlertas"
Option Explicit


'----------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------

Dim HuboErrores As Boolean
Dim objRs2 As New ADODB.Recordset
Dim l_sql As String
Dim l_schednro As String
Dim l_dia As String
Dim l_hora As String
Dim TituloAlerta As String
Dim dirsalidas As String
Dim Usuario As String
Dim objConn2 As New ADODB.Connection

Global Contador As Integer
Global BProNro As Long
Global Const EncryptionKey = "56238"
Global AleNotiUnicaVez As Boolean
Global Mensaje As String
Global MaxTipoImg As Long
Global dirTemplates As String
Global Estilo As String



Sub Main()

Dim objRs As New ADODB.Recordset
Dim objRs2 As New ADODB.Recordset
Dim strCmdLine As String
Dim Nombre_Arch As String
Dim StrSql As String
Dim PID As String
Dim I As Integer
Dim arr
Dim Manual As String
Dim aleNro As String
Dim AlertaAccion As Integer
Dim AleProcPla As Integer
Dim sqlAle As String
Dim DispararAlerta As Boolean
Dim NroProceso As String
Dim colcount As Integer
Dim auxi As String
Dim tipoMails As Integer
Dim ArrParametros
Dim MailPorResultado As Boolean
Dim multipleAttach As Boolean


    strCmdLine = Command()
    ArrParametros = Split(strCmdLine, " ", -1)
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
            If IsNumeric(strCmdLine) Then
                NroProceso = strCmdLine
            Else
                Exit Sub
            End If
        End If
    End If
    
    On Error GoTo ME_Main

    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas

    TiempoInicialProceso = GetTickCount
    HuboErrores = False
    
    Nombre_Arch = PathFLog & "Alertas" & "-" & NroProceso & ".log"
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Version = " & Version
    Flog.writeline "Fecha   = " & FechaVersion
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline
    Flog.writeline "PID = " & PID
    Flog.writeline
    Flog.writeline "Inicio Alertas : " & Now
    Flog.writeline
    
    On Error Resume Next
    'Abro la conexion
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline "Problemas en la conexion"
        'Exit Sub
        GoTo CierroLog
    End If
    OpenConnection strconexion, objconnProgreso
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline "Problemas en la conexion"
        'Exit Sub
        GoTo CierroLog
    End If
    
    
    'FGZ - 06/08/2012 --------- Control de versiones ------
    Version_Valida = ValidarV(Version, 8, TipoBD)
    If Not Version_Valida Then
        'Actualizo el progreso
        MyBeginTrans
            StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Error de Version', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
        MyCommitTrans
        GoTo CierroLog
    End If
    'FGZ - 06/08/2012 --------- Control de versiones ------
    
    On Error GoTo CE
    
   
    Flog.writeline "Cambio el estado del proceso a Procesando"
    'Cambio el estado del proceso a Procesando '
    'StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProceso
    'FGZ - 06/01/2009 - le agregué el progreso en 1
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & ", bprcprogreso = 1 WHERE bpronro = " & NroProceso
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Obtengo los datos del proceso"
    'Obtengo los datos del proceso
    StrSql = "SELECT bprcparam, iduser FROM batch_proceso WHERE btprcnro = 8 AND bpronro = " & NroProceso
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        'FGZ - 27/09/2013 ----------------------
            If EsNulo(objRs!bprcparam) Then
                Flog.writeline "El proceso de Alertas no tiene parametros (no es posible determinar que alerta ejecutar). "
                aleNro = 0
                Manual = "-1"
            Else
                'Obtengo los parametros del proceso
                arr = Split(objRs!bprcparam, " ")
                aleNro = arr(0)
                Manual = ""
                If (UBound(arr) - LBound(arr) + 1) > 1 Then Manual = arr(1)
            End If
        'FGZ - 27/09/2013 ----------------------
        Usuario = objRs!iduser
        objRs.Close
        
    Else
        Flog.writeline "Error al buscar datos en el proceso: " & NroProceso
        Exit Sub
    End If
    If objRs.State = adStateOpen Then objRs.Close
       
    ' Directorio Salidas
    StrSql = "select sis_dirsalidas, sis_direntradas from sistema"
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        dirsalidas = objRs!sis_dirsalidas & "\attach"
        dirTemplates = objRs!sis_direntradas & "\templates"
        Flog.writeline "Directorio de Salidas: " & dirsalidas
    Else
        Flog.writeline "No se encuentra configurado sis_dirsalidas"
        'FGZ - 15/08/2013 ---------------------------------------------
        Call ActualizarEstado("Error General", PID)
        GoTo CierroLog
        'Exit Sub
        'FGZ - 15/08/2013 ---------------------------------------------
    End If
    If objRs.State = adStateOpen Then objRs.Close
    
    Call CargarEstilo
    
    
    ' Datos de la Alerta
    StrSql = "SELECT ale_cons.aleconsnro, alec_query, cnstring, alertas.aleaccnro, alertas.alenro, alertas.aledesext, alertas.schednro,alertas.aletipomails, alertas.alenotiunica, alertas.aleprocpl, alertas.mailxresul, alertas.alemultattach "
    StrSql = StrSql & " FROM alertas "
    StrSql = StrSql & " INNER JOIN ale_cons ON alertas.alenro = ale_cons.alenro "
    StrSql = StrSql & " INNER JOIN conexion ON ale_cons.cnnro = conexion.cnnro "
    If Manual <> "manual" Then StrSql = StrSql & " AND aleactiva = -1 "
    StrSql = StrSql & "AND alertas.alenro = " & aleNro & " "
    StrSql = StrSql & "ORDER BY ale_cons.aleconsnro"
    Flog.writeline StrSql
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        Flog.writeline "No se encuentra la alerta a ejecutar."
        'FGZ - 15/08/2013 ---------------------------------------------
        Call ActualizarEstado("Error", PID)
        GoTo CierroLog
        'Exit Sub
        'FGZ - 15/08/2013 ---------------------------------------------
    End If
    Flog.writeline "Se Buscan los Queries"

    Do Until objRs.EOF
        'FGZ - 30/03/2009 - se agregaó este campo
        AleNotiUnicaVez = CBool(objRs!alenotiunica)
        l_schednro = objRs!schednro
        TituloAlerta = objRs!aledesext
        AlertaAccion = objRs!aleaccnro
        'FGZ - 09/08/2012 -------------------
        MailPorResultado = CBool(objRs!mailxresul)
        'FGZ - 09/08/2012 -------------------
        
        'LED - 15/12/2014 -------------------
        multipleAttach = CBool(objRs!alemultattach)
        'LED - 15/12/2014 -------------------
        
        If EsNulo(objRs!aleprocpl) Then
            AleProcPla = False
        Else
            AleProcPla = CBool(objRs!aleprocpl)
        End If
        ' tipoMails = objRs!aletipomails
        OpenConnection objRs!cnstring, objConn2
        
        ' si la alerta es para los proceso planificados (en este caso 23-Interfaces), se hace una consulta nueva
        If AleProcPla Then
            sqlAle = ""
            armarSQLProcPlanificado objRs!aleNro, sqlAle
        Else
            sqlAle = ReplaceFields(objRs!alec_query, objRs!aleconsnro, objRs!aleNro)
        End If
        
        Flog.writeline "Se corre la siguiente consulta"
        Flog.writeline sqlAle
        
        OpenRecordsetWithConn sqlAle, objRs2, objConn2
        'FB - 06/06/2014 - Se verifica que el estado del recordset (objRs2) no sea cerrado.
        If objRs2.State <> adStateClosed Then
            If Not objRs2.EOF Then
               Flog.writeline "La consulta configurada en la alerta " & Trim(TituloAlerta) & " arrojó " & objRs2.RecordCount & " resultados."
               Flog.writeline "Accion de la alerta (numero): " & AlertaAccion
              
               Select Case AlertaAccion
                 Case 1
                    Flog.writeline "Accion de la alerta: Enviar Mails"
                    If MailPorResultado Then
                        Call EnviarVariosMails(aleNro, tipoMails, objRs2, multipleAttach)
                    Else
                        Call EnviarMails(aleNro, tipoMails, objRs2)
                    End If
                    If AleProcPla Then ' si es Alerta de procesos planificados
                        cargarAlertaInterfaces aleNro, objRs2
                    End If
               End Select
               
            Else
                Flog.writeline "Advertencia: La consulta configurada en la alerta : " & Trim(TituloAlerta) & " : no arrojó resultados."
            End If
            objRs2.Close
        Else
        'FB - 06/06/2014 - Muestra un error si el objrs2 esta cerrado
            Flog.writeline "Advertencia: La consulta configurada en la alerta : " & Trim(TituloAlerta) & " : no arrojó resultados."
            Flog.writeline "Advertencia: El estado del recordset es cerrado. "
        End If
        objConn2.Close
        objRs.MoveNext
    Loop

    If objRs.State = adStateOpen Then objRs.Close
    
    If Manual = "manual" Then
        l_dia = "BAJA"
    Else
        FechaHora
    End If

    If Not HuboErrores Then
        If l_dia = "BAJA" Then
            Flog.writeline "Setear estado del proceso = 'Procesado'."
            'Actualizo el estado del proceso
            StrSql = "UPDATE batch_proceso SET bprcprogreso =100 , bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
        Else
            Flog.writeline "Alarma recurrente. Setear estado del proceso = 'Pendiente'."
            
            StrSql = "UPDATE batch_proceso SET bprcprogreso =100 , bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date)
            
            StrSql = StrSql & ", bprcestado = 'Pendiente' "
            StrSql = StrSql & ",bprcfecha = " & ConvFecha(l_dia)
            'StrSql = StrSql & ",bprchora = '" & l_hora & "' "
            StrSql = StrSql & " Where bpronro = " & NroProceso
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
        End If
    Else
        Flog.writeline "Se detectaron errores"
        If l_dia = "BAJA" Then
            Flog.writeline "Setear estado del proceso = 'Incompleto'."
            StrSql = "UPDATE batch_proceso SET bprcprogreso =100 , bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProceso
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
        Else
            Flog.writeline "Alarma recurrente. Replanificacion."
            StrSql = "UPDATE batch_proceso SET bprcprogreso =100 , bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date)
            StrSql = StrSql & ", bprcestado = 'Pendiente' "
            StrSql = StrSql & ",bprcfecha = " & ConvFecha(l_dia)
            StrSql = StrSql & ",bprchora = '" & l_hora & "' "
            StrSql = StrSql & "Where bpronro = " & NroProceso
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
        End If
    End If
    
CierroLog:
    Flog.writeline
    Flog.writeline "Finalizado Alertas : " & Now
    Flog.writeline "-----------------------------------------------------------------"
    Flog.Close
    
    If objConn.State = adStateOpen Then objConn.Close
    If objConn2.State = adStateOpen Then objConn2.Close
    If objconnProgreso.State = adStateOpen Then objconnProgreso.Close
    Exit Sub
CE:
    HuboErrores = True
    Flog.writeline " Error: " & Err.Description & Now
    '-------------------------mdf
    If Manual <> "manual" Then
     Flog.writeline "Alarma recurrente. Replanificacion con errores en la query del alerta......."
     FechaHora
     StrSql = "UPDATE batch_proceso SET bprcprogreso =100 , bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date)
     StrSql = StrSql & ", bprcestado = 'Pendiente' "
     StrSql = StrSql & ",bprcfecha = " & ConvFecha(l_dia)
     'StrSql = StrSql & ",bprchora = '" & l_hora & "' "
     StrSql = StrSql & "Where bpronro = " & NroProceso
     objconnProgreso.Execute StrSql, , adExecuteNoRecords
    End If
    '------------------------ mdf
ME_Main:
    HuboErrores = True
   
End Sub


' __________________________________________________________
' __________________________________________________________
Sub armarSQLProcPlanificado(ByVal aleNro, ByRef sqlAle As String)
Dim StrSql As String
Dim objRs As New ADODB.Recordset
Dim l_alertaFecha As Date
Dim l_alertaHora
Dim l_aleenviada
Dim l_procesos
    
sqlAle = ""
l_aleenviada = 0

    
    ' busco la fecha en que se ejecuto la Alerta anterior, Alerta del tipo planificada - por ahora es solo tipo proceso 23- Interface General
    StrSql = "SELECT alenro, alefecha, alehora "
    StrSql = StrSql & " FROM ale_inter "
    StrSql = StrSql & " WHERE aleprocpl=-1 "
    StrSql = StrSql & " ORDER BY alefecha DESC, alehora DESC "
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        l_alertaFecha = objRs!alefecha
        l_alertaHora = objRs!alehora
        l_aleenviada = -1
    End If
    objRs.Close
        
    
    l_procesos = "0"
    
    ' buscar los procesos que estaban pendiente cuando se ejecuto la alerta
    StrSql = "SELECT alenro, bpronro "
    StrSql = StrSql & " FROM ale_inter "
    StrSql = StrSql & " WHERE aleprocpl=-1 AND alebprcest='Pendiente'"
    OpenRecordset StrSql, objRs
    Do While Not objRs.EOF
        l_procesos = l_procesos & "," & objRs!BProNro
    objRs.MoveNext
    Loop
    objRs.Close
    
    
    'OBS: el campo Bpronro debe estar en la posición 4 - NO CAMBIAR
    'arma la consulta sobre las interfaces ejecutadas, según planificador y su estado
    sqlAle = "SELECT bpropladesabr Descripción, scheddesc Planificador, "
    sqlAle = sqlAle & " batch_proceso.bprcfecha + ' ' + batch_proceso.bprchora ""Fecha Planificado"", "
    sqlAle = sqlAle & " batch_proceso.bprcfecInicioEj + ' ' + bprcHoraInicioEj ""Inicio Ejecución"", "
    sqlAle = sqlAle & " batch_proceso.bpronro Proceso, bpromodnro Modelo, batch_proceso.bprcestado Estado,  "
    sqlAle = sqlAle & " crpnregleidos ""Registros Leidos"",crpnregerr ""Registros con Errores"" "  ',crpnregadv ""Registros con Advertencias""
    sqlAle = sqlAle & " FROM batch_interplan "
    sqlAle = sqlAle & " INNER JOIN batch_procplan ON batch_procplan.bproplanro = batch_interplan.bproplanro "
    sqlAle = sqlAle & " INNER JOIN ale_sched ON ale_sched.schednro=batch_procplan.schednro "
    sqlAle = sqlAle & " INNER JOIN batch_proceso ON batch_interplan.bpronro = batch_proceso.bpronro "
    sqlAle = sqlAle & " LEFT JOIN inter_pin ON inter_pin.bpronro = batch_proceso.bpronro "
    If l_aleenviada = -1 Then
        sqlAle = sqlAle & " WHERE "
        sqlAle = sqlAle & "     batch_proceso.bpronro IN (" & l_procesos & ")"
        sqlAle = sqlAle & " OR "
        sqlAle = sqlAle & " ( "
        sqlAle = sqlAle & "   ( batch_proceso.bprcfecha > " & ConvFecha(l_alertaFecha) & ")"
        sqlAle = sqlAle & "     OR "
        sqlAle = sqlAle & "   (  batch_proceso.bprcfecha=" & ConvFecha(l_alertaFecha) & " AND batch_proceso.bprchora >= '" & Format(l_alertaHora, "HH:mm:ss") & "') "
        sqlAle = sqlAle & " )"
    End If
    sqlAle = sqlAle & " ORDER BY batch_proceso.bprcfecha, batch_proceso.bprchora "

    
End Sub

' ___________________________________________________________
' carga información de que proceso de Interface se mando mail
' ___________________________________________________________
Sub cargarAlertaInterfaces(ByVal aleNro As Integer, ByRef resultSet As ADODB.Recordset)
Dim StrSql As String
Dim objRs As New ADODB.Recordset
Dim l_bpronros

l_bpronros = "0"

    resultSet.MoveFirst
    Do While Not resultSet.EOF
        l_bpronros = l_bpronros & "," & resultSet(4)  ' lugar 4, resultSet!Proceso - resultSet!BProNro
    resultSet.MoveNext
    Loop
    
    
    ' buscar los procesos que estaban pendiente cuando se ejecuto la alerta anterior
    StrSql = "SELECT batch_proceso.bpronro, bprcestado "
    StrSql = StrSql & " FROM ale_inter "
    StrSql = StrSql & " INNER JOIN batch_proceso ON batch_proceso.bpronro = ale_inter.bpronro "
    StrSql = StrSql & " WHERE ale_inter.bpronro IN (" & l_bpronros & ")"
    OpenRecordset StrSql, objRs
    
    Do While Not objRs.EOF
        StrSql = " UPDATE ale_inter SET "
        StrSql = StrSql & " alebprcest ='" & objRs!bprcestado & "' ,"
        StrSql = StrSql & " alefecha=" & ConvFecha(Date) & ","
        StrSql = StrSql & " alehora='" & FormatDateTime(Time, 4) & ":00'"
        StrSql = StrSql & " WHERE bpronro = " & objRs!BProNro
        objConn.Execute StrSql, , adExecuteNoRecords
    objRs.MoveNext
    Loop
    objRs.Close
    
    ' buscar los procesos al cual se le envio una alerta
    StrSql = "SELECT batch_proceso.bpronro, bprcestado "
    StrSql = StrSql & " FROM batch_proceso "
    StrSql = StrSql & " WHERE batch_proceso.bpronro IN (" & l_bpronros & ")"
    StrSql = StrSql & "   AND batch_proceso.bpronro NOT IN (SELECT bpronro FROM ale_inter )"
    OpenRecordset StrSql, objRs
    
    Do While Not objRs.EOF
        StrSql = "INSERT INTO ale_inter (alenro,bpronro,alebprcest,alefecha, alehora, aleprocpl) "
        StrSql = StrSql & " VALUES (" & aleNro & "," & objRs!BProNro & ",'" & objRs!bprcestado & "'"
        StrSql = StrSql & "," & ConvFecha(Date) & ",'" & FormatDateTime(Time, 4) & ":00', -1 )"
        objConn.Execute StrSql, , adExecuteNoRecords
    objRs.MoveNext
    Loop
    objRs.Close

    
End Sub
                 



Sub EnviarMails(ByVal aleNro As Integer, ByVal Tipo As Integer, ByRef resultSet As ADODB.Recordset)
'-----------------------------------------------------------------------------------------------
' Se encarga de fijarse el mecanismo que se va a usar para enviar mails
'-----------------------------------------------------------------------------------------------
' Modificado :  21/09/2005 - Fapitalle N. - Ahora se relacionan las alertas
'               con las notificaciones
' Modificado :  09/08/2012 - FGZ - se agregó un parametro a las alertas para ver si se debe unviar un mail por cada resultado de la alerta
'-----------------------------------------------------------------------------------------------
Dim objRs As New ADODB.Recordset

    On Error GoTo CE
    
    ' Datos de las Notificaciones asociadas a la alerta
    
    StrSql = "SELECT noti_ale.notinro, noti_ale.rsenviar, notificacion.tnotinro "
    StrSql = StrSql & "From noti_ale "
    StrSql = StrSql & "INNER JOIN notificacion ON noti_ale.notinro = notificacion.notinro "
    StrSql = StrSql & "WHERE noti_ale.alenro = " & aleNro
    OpenRecordset StrSql, objRs

    Flog.writeline "Comienzo del Envio de Mails a " & StrSql
    
    
    MyBeginTrans
    
    'FGZ - 05/09/2006
    StrSql = "insert into batch_proceso "
    StrSql = StrSql & "(btprcnro, bprcfecha, iduser, bprchora, bprcfecdesde, bprcfechasta, bprcparam, "
    StrSql = StrSql & "bprcestado, bprcprogreso, bprcfecfin, bprchorafin, bprctiempo, empnro, bprcempleados) "
    StrSql = StrSql & "values (25," & ConvFecha(Date) & ",'" & Usuario & "','" & FormatDateTime(Time, 4) & ":00'"
    StrSql = StrSql & ",null,null,'1','Preparando',null,null,null,null,0,null)"
    objConn.Execute StrSql, , adExecuteNoRecords
    
    BProNro = getLastIdentity(objConn, "batch_proceso")
    
    Call BuscarMaxTipoImg
    
    Do Until objRs.EOF
        resultSet.MoveFirst
        Flog.writeline "Tipo" & objRs!TnotiNro
        Mensaje = "" ' 17-03-2011 - inicializar el mensaje para que no arrastre lo mensajes de otors tipos de notificacion
        
            Select Case objRs!TnotiNro
                Case 1  'Envia mails a los empleados definidos en la tabla noti_empleado
                    Call MailsAEmpleado(aleNro, objRs!notinro, objRs!rsenviar, resultSet)
                Case 2  'Envia mails a los empleados de las estructuras definidas en noti_estructura
                    Call MailsAEstructura(aleNro, objRs!notinro, objRs!rsenviar, resultSet)
                Case 3  'Envia mails a los supervisores de los empleados
                    Call MailsASupervisor(aleNro, objRs!notinro, objRs!rsenviar, resultSet)
                Case 4  'Envia mails a los usuarios definidos en la tabla noti_usuario
                    Call MailsAUsuario(aleNro, objRs!notinro, objRs!rsenviar, resultSet)
                Case 5  'Envia mails a los roles definidos en la tabla noti_roles
                    Call MailsARoles(aleNro, objRs!notinro, objRs!rsenviar, resultSet)
                Case 6  'Envia mails a los empleados que cumplan con el query de la notificacion
                    Call MailsAQuery(aleNro, objRs!notinro, objRs!rsenviar, resultSet)
                Case 7  'Envia mails a los Postulantes que cumplan con el query de la notificacion (se ejecuta en conexion)
                    Flog.writeline
                    Flog.writeline "Envia mails a los Postulantes que cumplan con el query de la notificacion (se ejecuta en conexion)"
                    Flog.writeline
                    Call MailsAPostTLP(aleNro, objRs!notinro, objRs!rsenviar, resultSet)
                Case 8 'Envia mails a los terceros que cumplan con el query de la notificacion (Se debe incluir el campo ternro primero en el select)
                    Call MailsAQueryDelAlerta(aleNro, objRs!notinro, objRs!rsenviar, resultSet)
                Case 9 'Envia mails a los empleados que cumplan con el query de la notificacion (Se debe incluir el campo ternro primero en el select)
                    Call MailsAQueryDelAlertaEmp(aleNro, objRs!notinro, objRs!rsenviar, resultSet)
                Case 17
                    Call MailsAGrupoDeUsuario(aleNro, objRs!notinro, objRs!rsenviar, resultSet)
            End Select
        
        objRs.MoveNext
    Loop
    
    'FGZ - 05/09/2006
    StrSql = "UPDATE batch_proceso SET bprcestado = 'Pendiente'"
    StrSql = StrSql & " WHERE bpronro = " & BProNro
    objConn.Execute StrSql, , adExecuteNoRecords
    MyCommitTrans
    
    Exit Sub
CE:
    HuboErrores = True
    Flog.writeline " Error: " & Err.Description & " > " & Now
    MyRollbackTrans
End Sub


'-----------------------------------------------------------------------------------------------
' Se encarga de ejecutar la query correspondiente y mandar el resultado por mail a los distintos destinatarios
'-----------------------------------------------------------------------------------------------
Sub MailsAQuery(ByVal aleNro As Integer, ByVal notinro As Integer, ByVal rsenviar As Integer, ByRef resultSet As ADODB.Recordset)
Dim I As Integer
Dim colcount As Integer
'Dim Mensaje As String
Dim objRs As New ADODB.Recordset
Dim strSQLQuery As String
Dim Param As String
Dim Opcion As Integer
Dim campo As Integer
Dim reemplazo As String
Dim Mails As String
Dim AlertaFileName As String
Dim fs2, AlertaFile
Dim Clave() As String
Dim Enviar As Boolean
Dim rs_Env As New ADODB.Recordset
Dim Indice_rs As Long

    On Error GoTo CE
    Enviar = False
    Indice_rs = 0
    ReDim Preserve Clave(1 To resultSet.RecordCount) As String
    
    colcount = resultSet.Fields.Count
    Do Until resultSet.EOF
        Indice_rs = Indice_rs + 1
        'Escribo el titulo de las columnas
        Mensaje = "<TABLE class=""tabladetalle""><TR>" & vbCrLf
        For I = 0 To (colcount - 1)
            Mensaje = Mensaje & "<TH>" & resultSet.Fields(I).Name & "</TH>" & vbCrLf
        Next
        Mensaje = Mensaje & "</TR>" & vbCrLf
        'Escribo el contenido de la fila
        Mensaje = Mensaje & "<TR>" & vbCrLf
        For I = 0 To (colcount - 1)
            Mensaje = Mensaje & "<TD>" & resultSet.Fields(I).Value & "</TD>" & vbCrLf
            Clave(Indice_rs) = Clave(Indice_rs) & resultSet.Fields(I).Value & " "
        Next
        Mensaje = Mensaje & "</TR></TABLE>" & vbCrLf
        
        'recupero la query correspondiente a la notificacion
        StrSql = "SELECT query FROM noti_query WHERE notinro = " & notinro
        OpenRecordset StrSql, objRs
        If Not objRs.EOF Then
            'reemplazo los parametros de la query por valores verdaderos
            strSQLQuery = objRs!query
            objRs.Close
            Do Until InStr(1, strSQLQuery, "[") = 0
                Param = Mid(strSQLQuery, InStr(1, strSQLQuery, "[") + 1, Len(strSQLQuery))
                Param = Mid(Param, 1, InStr(1, Param, "]") - 1)
                'si tiene forma [campo,opcion] o tiene forma [campo]
                If InStr(1, Param, ",") <> 0 Then
                    campo = Mid(Param, 1, InStr(1, Param, ",") - 1)
                    Opcion = Mid(Param, InStr(1, Param, ",") + 1, Len(Param) - InStr(1, Param, ","))
                Else
                    campo = Param
                End If
                
                Select Case campo
                    Case 1  ' Fecha Actual
                        reemplazo = Date
                    Case 2  ' Columna X, donde X es opcion
                        reemplazo = resultSet.Fields(Opcion).Value
                End Select
                strSQLQuery = Replace(strSQLQuery, "[" & Param & "]", reemplazo)
            Loop ' hasta aca reemplazo la sql
            Flog.writeline "Mail a Query: " & strSQLQuery

            Mails = ""
            OpenRecordset strSQLQuery, objRs
            Enviar = False
            Do Until objRs.EOF
                ' la primera columna del query es una dir de email
                If Not IsNull(objRs.Fields(0).Value) Then
                    If Len(objRs.Fields(0).Value) > 0 Then
                        Mails = Mails & objRs.Fields(0).Value & ";"
                        
                        'FGZ - 30/03/2009 - se le agregó esto para que registre las alertas enviadas ----
                        If AleNotiUnicaVez Then
                            For I = 1 To (Indice_rs)
                                'reviso que no haya sido notificada ya
                                StrSql = "SELECT * FROM ale_enviada "
                                StrSql = StrSql & " WHERE alenro = " & aleNro
                                StrSql = StrSql & " AND notinro = " & notinro
                                StrSql = StrSql & " AND mail = '" & objRs.Fields(0).Value & "'"
                                StrSql = StrSql & " AND clave = '" & Clave(I) & "'"
                                OpenRecordset StrSql, rs_Env
                                If rs_Env.EOF Then
                                    StrSql = "INSERT INTO ale_enviada "
                                    StrSql = StrSql & "(alenro, notinro, fecha, mail, envios, clave) VALUES "
                                    StrSql = StrSql & "("
                                    StrSql = StrSql & aleNro & ","
                                    StrSql = StrSql & notinro & ","
                                    StrSql = StrSql & ConvFecha(Date) & ","
                                    'FGZ - 03/06/2014 -----------------------------
                                    StrSql = StrSql & "'" & Left(objRs.Fields(0).Value, 100) & "',"
                                    StrSql = StrSql & 0 & ","
                                    StrSql = StrSql & "'" & Left(Clave(I), 1000) & "'"
                                    StrSql = StrSql & ")"
                                    objConn.Execute StrSql, , adExecuteNoRecords
                                    
                                    Enviar = Enviar Or True
                                Else
                                    Flog.writeline "Alerta: ya ha sido notificada anteriormente el " & rs_Env!Fecha & ". No se reenviará."
                                    Enviar = Enviar Or False
                                End If
                            Next I
                        Else
                            Enviar = True
                        End If
                        'FGZ - 30/03/2009 - se le agregó esto para que registre las alertas enviadas ----
                        
                    End If
                End If
                objRs.MoveNext
            Loop
            If Mails <> "" And Enviar Then
                Contador = Contador + 1
                'AlertaFileName = dirsalidas & "\ale_" & Replace(FormatDateTime(Date, 2), "/", "-") & "_" & Replace(FormatDateTime(Time, 4), ":", "-") & "-" & String(2 - Len(Second(Now)), "0") & Second(Now) & Contador
                'FGZ - 05/09/2006
                AlertaFileName = dirsalidas & "\msg_" & BProNro & "_ale_" & Replace(FormatDateTime(Date, 2), "/", "-") & "_" & Replace(FormatDateTime(Time, 4), ":", "-") & "-" & String(2 - Len(Second(Now)), "0") & Second(Now) & Contador
                
                'creo el proceso que se encarga de mandar el mail
                If (rsenviar = -1) Then
                    'Creo el archivo a atachar en el mail
                    Set fs2 = CreateObject("Scripting.FileSystemObject")
                    Set AlertaFile = fs2.CreateTextFile(AlertaFileName & ".html", True)
                    Flog.writeline "Fin Alerta: Positiva. Enviado a: " & Mails
                    
                    'AlertaFile.writeline "<html><head>"
                    'AlertaFile.writeline "<STYLE>TABLE{ border : thick solid 1; width : 100%;}TH{ background-color: #333399; COLOR: #ffffff; FONT-FAMILY: 'Arial'; FONT-SIZE: 9pt; FONT-WEIGHT: bold; padding : 2 2 2 5; width : auto;}"
                    'AlertaFile.writeline "TR{ COLOR: black; FONT-FAMILY: Verdana; FONT-SIZE: 08pt; BACKGROUND-COLOR: #E4FEF9; padding : 2; padding-left : 5;}h4{font-family : Verdana, Geneva, Arial, Helvetica, sans-serif;font-size : smaller;font-style : normal;color : Maroon;}</STYLE>"
                    'AlertaFile.writeline "<title>Alertas - RHPro &reg;</title></head><body>"
                    'AlertaFile.writeline "<h4>" & TituloAlerta & "</h4>"
                    'AlertaFile.writeline "<table>" & Mensaje & "</table>"
                    'AlertaFile.writeline "</body></html>"
                    'AlertaFile.Close
                    
                    AlertaFile.writeline "<html><head>"
                    AlertaFile.writeline "<STYLE>" & Estilo & "</STYLE>"
                    AlertaFile.writeline "<title> Alertas - RHPro &reg; </title></head><body>"
                    AlertaFile.writeline "<h4>" & TituloAlerta & "</h4>"
                    AlertaFile.writeline Mensaje
                    AlertaFile.writeline "</body></html>"
                    AlertaFile.Close
                    
                    
                    Call crearProcesosMensajeria(aleNro, Mails, AlertaFileName, AlertaFileName)
                Else
                    Call crearProcesosMensajeria(aleNro, Mails, AlertaFileName, "")
                End If
            Else
                Flog.writeline "Alerta: No se encontraron mails definidos en el resultado de la consulta. No se han enviado mensajes."
            End If
        End If
        objRs.Close
        resultSet.MoveNext
    Loop
    Exit Sub
CE:
    HuboErrores = True
    Flog.writeline " Error: " & Err.Description & " > " & Now
End Sub


'-----------------------------------------------------------------------------------------------
' Se encarga de enviar el resultado de la alerta a los empleados definidos en la tabla noti_empleado
'-----------------------------------------------------------------------------------------------
Sub MailsAEstructura(ByVal aleNro As Integer, ByVal notinro As Integer, ByVal rsenviar As Integer, ByRef resultSet As ADODB.Recordset)

Dim colcount As Long
'Dim Mensaje As String
Dim Mails As String
Dim I As Long
Dim objRs As New ADODB.Recordset
Dim AlertaFileName As String
Dim fs2, MsgFile
Dim AlertaFile
Dim Estructuras As String
Dim tipoEstructura As Integer
Dim Clave() As String
Dim Enviar As Boolean
Dim rs_Env As New ADODB.Recordset
Dim Indice_rs As Long

    
    On Error GoTo CE
    Enviar = False
    Indice_rs = 0
    ReDim Preserve Clave(1 To resultSet.RecordCount) As String
        
    colcount = resultSet.Fields.Count
    'Escribo el titulo de las columnas
    Mensaje = "<TABLE class=""tabladetalle"">" & vbCrLf
    Mensaje = Mensaje & "<TR>" & vbCrLf
    For I = 0 To (colcount - 1)
        Mensaje = Mensaje & "<TH>" & resultSet.Fields(I).Name & "</TH>" & vbCrLf
    Next
    Mensaje = Mensaje & "</TR>" & vbCrLf
    
    'Escribo los datos de la tabla
    Do Until resultSet.EOF
        Indice_rs = Indice_rs + 1
        Mensaje = Mensaje & "<TR>" & vbCrLf
        For I = 0 To (colcount - 1)
            Mensaje = Mensaje & "<TD>" & resultSet.Fields(I).Value & "</TD>" & vbCrLf
            Clave(Indice_rs) = Clave(Indice_rs) & resultSet.Fields(I).Value & " "
        Next
        Mensaje = Mensaje & "</TR>" & vbCrLf
        resultSet.MoveNext
    Loop
    Mensaje = Mensaje & "</TABLE>" & vbCrLf
        
    'Creo el archivo a atachar en el mail
    Contador = Contador + 1
    'AlertaFileName = dirsalidas & "\ale_" & Replace(FormatDateTime(Date, 2), "/", "-") & "_" & Replace(FormatDateTime(Time, 4), ":", "-") & "-" & String(2 - Len(Second(Now)), "0") & Second(Now) & Contador
    'FGZ - 05/09/2006
    AlertaFileName = dirsalidas & "\msg_" & BProNro & "_ale_" & Replace(FormatDateTime(Date, 2), "/", "-") & "_" & Replace(FormatDateTime(Time, 4), ":", "-") & "-" & String(2 - Len(Second(Now)), "0") & Second(Now) & Contador
    
    If (rsenviar = -1) Then
        Set fs2 = CreateObject("Scripting.FileSystemObject")
        Set AlertaFile = fs2.CreateTextFile(AlertaFileName & ".html", True)
        
        'AlertaFile.writeline "<html><head>"
        'AlertaFile.writeline "<STYLE>TABLE{ border : thick solid 1; width : 100%;}TH{ background-color: #333399; COLOR: #ffffff; FONT-FAMILY: 'Arial'; FONT-SIZE: 9pt; FONT-WEIGHT: bold; padding : 2 2 2 5; width : auto;}"
        'AlertaFile.writeline "TR{ COLOR: black; FONT-FAMILY: Verdana; FONT-SIZE: 08pt; BACKGROUND-COLOR: #E4FEF9; padding : 2; padding-left : 5;}h4{font-family : Verdana, Geneva, Arial, Helvetica, sans-serif;font-size : smaller;font-style : normal;color : Maroon;}</STYLE>"
        'AlertaFile.writeline "<title>Alertas - RHPro &reg;</title></head><body>"
        'AlertaFile.writeline "<h4>" & TituloAlerta & "</h4>"
        'AlertaFile.writeline "<table>" & Mensaje & "</table>"
        'AlertaFile.writeline "</body></html>"
        'AlertaFile.Close
        
        AlertaFile.writeline "<html><head>"
        AlertaFile.writeline "<STYLE>" & Estilo & "</STYLE>"
        AlertaFile.writeline "<title> Alertas - RHPro &reg; </title></head><body>"
        AlertaFile.writeline "<h4>" & TituloAlerta & "</h4>"
        AlertaFile.writeline Mensaje
        AlertaFile.writeline "</body></html>"
        AlertaFile.Close
        
        
    End If
            
    'Busco todos los empleados a los cuales les tengo que enviar los mails
    
    ' recupero una lista de estructuras a las cuales mandarle el mail
    StrSql = "SELECT * FROM noti_estructura WHERE notinro = " & notinro
    OpenRecordset StrSql, objRs
    Estructuras = "0"
    If objRs.EOF Then
        Estructuras = ""
        Exit Sub
    Else
        If objRs!estrfija = -1 Then
            Do Until objRs.EOF
                Estructuras = Estructuras & "," & objRs!estrnro
                objRs.MoveNext
            Loop
        Else
            tipoEstructura = objRs!tenro
        End If
    End If
    objRs.Close
    
    If (Estructuras <> "0") Then
    'si es fija es un mismo mail para todos
        StrSql = " SELECT DISTINCT empemail FROM empleado "
        StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.ternro = empleado.ternro "
        StrSql = StrSql & " WHERE his_estructura.htethasta IS NULL "
        StrSql = StrSql & " AND his_estructura.estrnro IN (" & Estructuras & ")"
        OpenRecordset StrSql, objRs
        Mails = ""
        Enviar = False
        Do Until objRs.EOF
            If Not IsNull(objRs!empemail) Then
                If Len(objRs!empemail) > 0 Then
                    Mails = Mails & objRs!empemail & ";"
                    
                    'FGZ - 30/03/2009 - se le agregó esto para que registre las alertas enviadas ----
                    If AleNotiUnicaVez Then
                        For I = 1 To (Indice_rs)
                            'reviso que no haya sido notificada ya
                            StrSql = "SELECT * FROM ale_enviada "
                            StrSql = StrSql & " WHERE alenro = " & aleNro
                            StrSql = StrSql & " AND notinro = " & notinro
                            StrSql = StrSql & " AND mail = '" & objRs!empemail & "'"
                            StrSql = StrSql & " AND clave = '" & Clave(I) & "'"
                            OpenRecordset StrSql, rs_Env
                            If rs_Env.EOF Then
                                StrSql = "INSERT INTO ale_enviada "
                                StrSql = StrSql & "(alenro, notinro, fecha, mail, envios, clave) VALUES "
                                StrSql = StrSql & "("
                                StrSql = StrSql & aleNro & ","
                                StrSql = StrSql & notinro & ","
                                StrSql = StrSql & ConvFecha(Date) & ","
                                'FGZ - 03/06/2014 -----------------------------
                                'StrSql = StrSql & "'" & objRs!usremail & "',"
                                StrSql = StrSql & "'" & Left(objRs!usremail, 100) & "',"
                                StrSql = StrSql & 0 & ","
                                'StrSql = StrSql & "'" & Clave(I) & "'"
                                StrSql = StrSql & "'" & Left(Clave(I), 1000) & "'"
                                StrSql = StrSql & ")"
                                objConn.Execute StrSql, , adExecuteNoRecords
                                
                                Enviar = Enviar Or True
                            Else
                                Flog.writeline "Alerta: ya ha sido notificada anteriormente el " & rs_Env!Fecha & ". No se reenviará."
                                Enviar = Enviar Or False
                            End If
                        Next I
                    Else
                        Enviar = True
                    End If
                    'FGZ - 30/03/2009 - se le agregó esto para que registre las alertas enviadas ----
                    
                End If
            End If
            objRs.MoveNext
        Loop
        objRs.Close
        
        If Mails <> "" And Enviar Then
             Flog.writeline "Fin Alerta: Positiva. Enviado a: " & Mails
            'creo el proceso que se encarga de mandar el mail
            If (rsenviar = -1) Then
                Call crearProcesosMensajeria(aleNro, Mails, AlertaFileName, AlertaFileName)
            Else
                Call crearProcesosMensajeria(aleNro, Mails, AlertaFileName, "")
            End If
        Else
            Flog.writeline "Alerta: No se encontraron mails definidos. Estructuras nro: " & Estructuras
        End If
    Else
    'si no es fija es un mail a los de la estructura de cada empleado
        resultSet.MoveFirst
        Enviar = False
        Do Until resultSet.EOF
            StrSql = "SELECT DISTINCT empemail FROM empleado "
            StrSql = StrSql & "INNER JOIN his_estructura ON his_estructura.ternro = empleado.ternro "
            StrSql = StrSql & "WHERE his_estructura.estrnro = ("
            StrSql = StrSql & "    SELECT his_estructura.estrnro FROM his_estructura"
            StrSql = StrSql & "    INNER JOIN noti_estructura ON his_estructura.tenro = noti_estructura.tenro"
            StrSql = StrSql & "    WHERE his_estructura.htethasta IS NULL"
            StrSql = StrSql & "    AND his_estructura.Ternro = " & resultSet.Fields(0).Value
            StrSql = StrSql & "    AND noti_estructura.notinro = " & notinro & "  )"
            OpenRecordset StrSql, objRs
            
            Mails = ""
            Do Until objRs.EOF
                If Not IsNull(objRs!empemail) Then
                    If Len(objRs!empemail) > 0 Then
                        Mails = Mails & objRs!empemail & ";"
                        
                        'FGZ - 30/03/2009 - se le agregó esto para que registre las alertas enviadas ----
                        If AleNotiUnicaVez Then
                            For I = 1 To (Indice_rs)
                                'reviso que no haya sido notificada ya
                                StrSql = "SELECT * FROM ale_enviada "
                                StrSql = StrSql & " WHERE alenro = " & aleNro
                                StrSql = StrSql & " AND notinro = " & notinro
                                StrSql = StrSql & " AND mail = '" & objRs!empemail & "'"
                                StrSql = StrSql & " AND clave = '" & Clave(I) & "'"
                                OpenRecordset StrSql, rs_Env
                                If rs_Env.EOF Then
                                    StrSql = "INSERT INTO ale_enviada "
                                    StrSql = StrSql & "(alenro, notinro, fecha, mail, envios, clave) VALUES "
                                    StrSql = StrSql & "("
                                    StrSql = StrSql & aleNro & ","
                                    StrSql = StrSql & notinro & ","
                                    StrSql = StrSql & ConvFecha(Date) & ","
                                    'FGZ - 03/06/2014 -----------------------------
                                    'StrSql = StrSql & "'" & objRs!empemail & "',"
                                    StrSql = StrSql & "'" & Left(objRs!empemail, 100) & "',"
                                    StrSql = StrSql & 0 & ","
                                    'StrSql = StrSql & "'" & Clave(I) & "'"
                                    StrSql = StrSql & "'" & Left(Clave(I), 1000) & "'"
                                    StrSql = StrSql & ")"
                                    objConn.Execute StrSql, , adExecuteNoRecords
                                    
                                    Enviar = Enviar Or True
                                Else
                                    Flog.writeline "Alerta: ya ha sido notificada anteriormente el " & rs_Env!Fecha & ". No se reenviará."
                                    Enviar = Enviar Or False
                                End If
                            Next I
                        Else
                            Enviar = True
                        End If
                        'FGZ - 30/03/2009 - se le agregó esto para que registre las alertas enviadas ----
                        
                    End If
                End If
                objRs.MoveNext
            Loop
            
            Flog.writeline "Se han encontrado " & objRs.RecordCount & " registros del tipo de estructura del empleado numero:" & resultSet.Fields(0).Value & ". Mails:<" & Mails & ">"
            
            objRs.Close
            
            If Mails <> "" And Enviar Then
                'creo el proceso que se encarga de mandar el mail
                If (rsenviar = -1) Then
                    Call crearProcesosMensajeria(aleNro, Mails, AlertaFileName, AlertaFileName)
                Else
                    Call crearProcesosMensajeria(aleNro, Mails, AlertaFileName, "")
                End If
            End If
            resultSet.MoveNext
        Loop
    End If
    Exit Sub
CE:
    HuboErrores = True
    Flog.writeline " Error: " & Err.Description & " > " & Now
End Sub


'-----------------------------------------------------------------------------------------------
' Se encarga de enviar el resultado de la alerta a los empleados definidos en la tabla noti_empleado
'-----------------------------------------------------------------------------------------------
Sub MailsAEmpleado(ByVal aleNro As Integer, ByVal notinro As Integer, ByVal rsenviar As Integer, ByRef resultSet As ADODB.Recordset)

Dim colcount As Long
'Dim Mensaje As String
Dim Mails As String
Dim I As Long
Dim objRs As New ADODB.Recordset
Dim AlertaFileName As String
Dim fs2, MsgFile
Dim AlertaFile
Dim Enviar As Boolean
Dim rs_Env As New ADODB.Recordset
Dim Clave() As String
Dim Indice_rs As Long

    On Error GoTo CE
    ReDim Preserve Clave(1 To resultSet.RecordCount) As String
    Indice_rs = 0
    
    colcount = resultSet.Fields.Count
    'Escribo el titulo de las columnas
    
    Mensaje = "<TABLE class=""tabladetalle"">" & vbCrLf
    Mensaje = Mensaje & "<TR>" & vbCrLf
    For I = 0 To (colcount - 1)
        Mensaje = Mensaje & "<TH>" & resultSet.Fields(I).Name & "</TH>" & vbCrLf
    Next
    
    'Escribo los datos de la tabla
    Mensaje = Mensaje & "</TR>" & vbCrLf
    Do Until resultSet.EOF
        Indice_rs = Indice_rs + 1
        Mensaje = Mensaje & "<TR>" & vbCrLf
        For I = 0 To (colcount - 1)
            Mensaje = Mensaje & "<TD>" & resultSet.Fields(I).Value & "</TD>" & vbCrLf
            Clave(Indice_rs) = Clave(Indice_rs) & resultSet.Fields(I).Value & " "
        Next
        Mensaje = Mensaje & "</TR>" & vbCrLf
        resultSet.MoveNext
    Loop
    Mensaje = Mensaje & "</TABLE>" & vbCrLf
            
    'Busco todos los empleados a los cuales les tengo que enviar los mails
    StrSql = "SELECT empemail FROM tercero "
    StrSql = StrSql & "inner join noti_empleado on tercero.ternro = noti_empleado.ternro "
    StrSql = StrSql & "inner join empleado on empleado.ternro = noti_empleado.ternro "
    StrSql = StrSql & "where notinro = " & notinro
    
    OpenRecordset StrSql, objRs
    Mails = ""
    Enviar = False
    Do Until objRs.EOF
        'If Not IsNull(objRs!teremail) Then
        If Not IsNull(objRs!empemail) Then
           'If Len(objRs!teremail) > 0 Then
           If Len(objRs!empemail) > 0 Then
                'Mails = Mails & objRs!teremail & ";"
                Mails = Mails & objRs!empemail & ";"
              
                'FGZ - 30/03/2009 - se le agregó esto para que registre las alertas enviadas ----
                If AleNotiUnicaVez Then
                    For I = 1 To (Indice_rs)
                        'reviso que no haya sido notificada ya
                        StrSql = "SELECT * FROM ale_enviada "
                        StrSql = StrSql & " WHERE alenro = " & aleNro
                        StrSql = StrSql & " AND notinro = " & notinro
                        StrSql = StrSql & " AND mail = '" & objRs!empemail & "'"
                        StrSql = StrSql & " AND clave = '" & Clave(I) & "'"
                        OpenRecordset StrSql, rs_Env
                        If rs_Env.EOF Then
                            StrSql = "INSERT INTO ale_enviada "
                            StrSql = StrSql & "(alenro, notinro, fecha, mail, envios, clave) VALUES "
                            StrSql = StrSql & "("
                            StrSql = StrSql & aleNro & ","
                            StrSql = StrSql & notinro & ","
                            StrSql = StrSql & ConvFecha(Date) & ","
                            'FGZ - 03/06/2014 -----------------------------
                            'StrSql = StrSql & "'" & objRs!empemail & "',"
                            StrSql = StrSql & "'" & Left(objRs!empemail, 100) & "',"
                            'FGZ - 03/06/2014 -----------------------------
                            StrSql = StrSql & 0 & ","
                            'FGZ - 03/06/2014 -----------------------------
                            'StrSql = StrSql & "'" & Clave(I) & "'"
                            StrSql = StrSql & "'" & Left(Clave(I), 1000) & "'"
                            'FGZ - 03/06/2014 -----------------------------
                            StrSql = StrSql & ")"
                            objConn.Execute StrSql, , adExecuteNoRecords
                            
                            Enviar = Enviar Or True
                        Else
                            Flog.writeline "Alerta: ya ha sido notificada anteriormente el " & rs_Env!Fecha & ". No se reenviará."
                            Enviar = Enviar Or False
                        End If
                    Next I
                Else
                    Enviar = True
                End If
                'FGZ - 30/03/2009 - se le agregó esto para que registre las alertas enviadas ----
              
           End If
        End If
        objRs.MoveNext
    Loop
    objRs.Close
    
    If Mails <> "" And Enviar Then
        'creo el proceso que se encarga de mandar el mail
        Contador = Contador + 1
        'AlertaFileName = dirsalidas & "\ale_" & Replace(FormatDateTime(Date, 2), "/", "-") & "_" & Replace(FormatDateTime(Time, 4), ":", "-") & "-" & String(2 - Len(Second(Now)), "0") & Second(Now) & Contador
        'FGZ - 05/09/2006
        'AlertaFileName = dirsalidas & "\msg_" & BProNro & "ale_" & Replace(FormatDateTime(Date, 2), "/", "-") & "_" & Replace(FormatDateTime(Time, 4), ":", "-") & "-" & String(2 - Len(Second(Now)), "0") & Second(Now) & Contador
        'FGZ - 21/06/2007 - estaba faltando el segundo guion bajo y por eso el proceso de mensajeria no lo levanta
        AlertaFileName = dirsalidas & "\msg_" & BProNro & "_ale_" & Replace(FormatDateTime(Date, 2), "/", "-") & "_" & Replace(FormatDateTime(Time, 4), ":", "-") & "-" & String(2 - Len(Second(Now)), "0") & Second(Now) & Contador
        
        If (rsenviar = -1) Then
            'Creo el archivo a atachar en el mail
            Set fs2 = CreateObject("Scripting.FileSystemObject")
            Set AlertaFile = fs2.CreateTextFile(AlertaFileName & ".html", True)
            Flog.writeline "Fin Alerta: Positiva. Enviado a: " & Mails
            
            'AlertaFile.writeline "<html><head>"
            'AlertaFile.writeline "<STYLE>TABLE{ border : thick solid 1; width : 100%;}TH{ background-color: #333399; COLOR: #ffffff; FONT-FAMILY: 'Arial'; FONT-SIZE: 9pt; FONT-WEIGHT: bold; padding : 2 2 2 5; width : auto;}"
            'AlertaFile.writeline "TR{ COLOR: black; FONT-FAMILY: Verdana; FONT-SIZE: 08pt; BACKGROUND-COLOR: #E4FEF9; padding : 2; padding-left : 5;}h4{font-family : Verdana, Geneva, Arial, Helvetica, sans-serif;font-size : smaller;font-style : normal;color : Maroon;}</STYLE>"
            'AlertaFile.writeline "<title>Alertas - RHPro &reg;</title></head><body>"
            'AlertaFile.writeline "<h4>" & TituloAlerta & "</h4>"
            'AlertaFile.writeline "<table>" & Mensaje & "</table>"
            'AlertaFile.writeline "</body></html>"
            'AlertaFile.Close
            
            
            AlertaFile.writeline "<html><head>"
            AlertaFile.writeline "<STYLE>" & Estilo & "</STYLE>"
            AlertaFile.writeline "<title> Alertas - RHPro &reg; </title></head><body>"
            AlertaFile.writeline "<h4>" & TituloAlerta & "</h4>"
            AlertaFile.writeline Mensaje
            AlertaFile.writeline "</body></html>"
            AlertaFile.Close
            
            Call crearProcesosMensajeria(aleNro, Mails, AlertaFileName, AlertaFileName)
        Else
            Call crearProcesosMensajeria(aleNro, Mails, AlertaFileName, "")
        End If
    Else
        Flog.writeline "Alerta: No se encontraron mails definidos en los empleados. No se han enviado mensajes."
    End If
    Exit Sub
CE:
    HuboErrores = True
    Flog.writeline " Error: " & Err.Description & " > " & Now
    Flog.writeline " Ultimo SQL Ejecutado: " & StrSql
End Sub

'-----------------------------------------------------------------------------------------------
' Se encarga de ejecutar la query correspondiente y mandar el resultado por mail a los distintos destinatarios
'-----------------------------------------------------------------------------------------------
Sub MailsAPostTLP(ByVal aleNro As Integer, ByVal notinro As Integer, ByVal rsenviar As Integer, ByRef resultSet As ADODB.Recordset)
Dim I As Integer
Dim colcount As Integer
'Dim Mensaje As String
Dim objRs As New ADODB.Recordset
Dim strSQLQuery As String
Dim Param As String
Dim Opcion As Integer
Dim campo As Integer
Dim reemplazo As String
Dim Mails As String
Dim AlertaFileName As String
Dim fs2, AlertaFile

Dim rs_Link As New ADODB.Recordset
Dim Link As String
Dim Texto_Body As String
'Dim Clave() As String
'Dim Enviar As Boolean
'Dim rs_Env As New ADODB.Recordset
'Dim Indice_rs As Long


    On Error GoTo CE
    'Enviar = False
    'Indice_rs = 0
    'ReDim Preserve Clave(1 To resultSet.RecordCount) As String
    
    StrSql = " SELECT cnnro,cnstring FROM conexion "
    StrSql = StrSql & " WHERE cnnro = 3 "
    OpenRecordset StrSql, rs_Link
    If Not rs_Link.EOF Then
        Link = rs_Link!cnstring
    Else
        Flog.writeline "No se encuentra la conexion para asociar en el mail. Debe existir una conexion nro 3."
    End If
    
    
    colcount = resultSet.Fields.Count
    If resultSet.EOF Then
        Flog.writeline "No hay resultados"
    End If
    Do Until resultSet.EOF
    
        'recupero la query correspondiente a la notificacion
        StrSql = "SELECT query FROM noti_query WHERE notinro = " & notinro
        OpenRecordset StrSql, objRs
        If Not objRs.EOF Then
            'reemplazo los parametros de la query por valores verdaderos
            strSQLQuery = objRs!query
            objRs.Close
            Do Until InStr(1, strSQLQuery, "[") = 0
                Param = Mid(strSQLQuery, InStr(1, strSQLQuery, "[") + 1, Len(strSQLQuery))
                Param = Mid(Param, 1, InStr(1, Param, "]") - 1)
                'si tiene forma [campo,opcion] o tiene forma [campo]
                If InStr(1, Param, ",") <> 0 Then
                    campo = Mid(Param, 1, InStr(1, Param, ",") - 1)
                    Opcion = Mid(Param, InStr(1, Param, ",") + 1, Len(Param) - InStr(1, Param, ","))
                Else
                    campo = Param
                End If
                
                Select Case campo
                    Case 1  ' Fecha Actual
                        reemplazo = Date
                    Case 2  ' Columna X, donde X es opcion
                        reemplazo = resultSet.Fields(Opcion).Value
                End Select
                strSQLQuery = Replace(strSQLQuery, "[" & Param & "]", reemplazo)
            Loop ' hasta aca reemplazo la sql
            Flog.writeline "Mail a Query: " & strSQLQuery

            'Escribo el Mensaje a enviar
            Mensaje = "<TABLE cellpadding=" & Chr(34) & "0" & Chr(34) & "cellspacing=" & Chr(34) & "0" & Chr(34) & ">"
            Mensaje = Mensaje & "<TR><TH> Aviso Importante</TH></TR>"
            Mensaje = Mensaje & "<tr><td>&nbsp;</td></tr>"
            Mensaje = Mensaje & "<tr><td> Recuerda completar tus datos personales para nuestra búsqueda</td></tr>"
            Mensaje = Mensaje & "<tr><td>&nbsp;</td></tr>"
            Mensaje = Mensaje & "<tr><td> Para continuar con el proceso de incorporación te solicitamos ingresar a la siguiente dirección "
            Mensaje = Mensaje & "<a href=" & Chr(34) & Link & Encrypt(EncryptionKey, resultSet!Ternro) & Chr(34) & "> " & Link & resultSet!Ternro & "<a></td></tr>"
            Mensaje = Mensaje & "<tr><td>&nbsp;</td></tr>"
            Mensaje = Mensaje & "<tr><td> Es vital que completes esta información antes del " & resultSet!busfecfin & " a las 24:00 hs, para poder contar con tus datos para el armado del contrato y la "
            Mensaje = Mensaje & "confirmación del día de la firma del mismo.</td></tr>"
            Mensaje = Mensaje & "<tr><td>&nbsp;</td></tr>"
            Mensaje = Mensaje & "<tr><td> En esta página encontraras información de nuestra Empresa y podrás completar tus datos personales.</td></tr>"
            Mensaje = Mensaje & "<tr><td>Te recomendamos que leas con atención ya que esta información que ingreses se utilizara para armar tu legajo.</td></tr>"
            Mensaje = Mensaje & "<tr><td>&nbsp;</td></tr>"
            Mensaje = Mensaje & "<tr><td> Cualquier inconveniente que tengas, por favor comunicate al siguiente tel 5555-55555</td></tr>"
            Mensaje = Mensaje & "<tr><td>&nbsp;</td></tr>"
            Mensaje = Mensaje & "<tr><td>&nbsp;</td></tr>"
            Mensaje = Mensaje & "<tr><td>&nbsp;</td></tr>"
            Mensaje = Mensaje & "<tr><td align=" & Chr(34) & "right" & Chr(34) & ">Departamento de Recruiting</td></tr>"
            Mensaje = Mensaje & "</TABLE>"

            Mails = ""
            strSQLQuery = " SELECT teremail FROM tercero "
            strSQLQuery = strSQLQuery & " WHERE ternro = " & resultSet!Ternro
            OpenRecordset strSQLQuery, objRs
            If Not objRs.EOF Then
                'la primera columna del query es una dir de email
                If Not IsNull(objRs.Fields(0).Value) Then
                    If Len(objRs.Fields(0).Value) > 0 Then
                        Mails = Mails & objRs.Fields(0).Value & ";"
                    End If
                End If
            End If
            
            If Mails <> "" Then
                'Este es el mensaje en html completo. Es lo mismo que lo que se adjunta ------
                Texto_Body = " Este es un aviso Importante !!!" & Chr(13)
                Texto_Body = Texto_Body & Chr(13)
                Texto_Body = Texto_Body & " Recuerda completar tus datos personales para nuestra búsqueda " & Chr(13)
                Texto_Body = Texto_Body & Chr(13)
                Texto_Body = Texto_Body & " Para continuar con el proceso de incorporación te solicitamos ingresar a la siguiente dirección " & Link & Encrypt(EncryptionKey, resultSet!Ternro) & Chr(13)
                Texto_Body = Texto_Body & Chr(13)
                Texto_Body = Texto_Body & " Es vital que completes esta información antes del " & resultSet!busfecfin & " a las 24:00 hs, para poder contar con tus datos para el armado del contrato "
                Texto_Body = Texto_Body & " y la confirmación del día de la firma del mismo." & Chr(13)
                Texto_Body = Texto_Body & Chr(13)
                Texto_Body = Texto_Body & " En esta página encontraras información de nuestra Empresa y podrás completar tus datos personales."
                Texto_Body = Texto_Body & "Te recomendamos que leas con atención ya que esta información que ingreses se utilizara para armar tu legajo." & Chr(13)
                Texto_Body = Texto_Body & Chr(13)
                Texto_Body = Texto_Body & " Cualquier inconveniente que tengas, por favor comunicate al siguiente tel 5555-55555"
                Texto_Body = Texto_Body & Chr(13)
                Texto_Body = Texto_Body & Chr(13)
                Texto_Body = Texto_Body & Chr(13)
                Texto_Body = Texto_Body & " Atte. Departamento de Recruiting."
                'Este es el mensaje en html completo. Es lo mismo que lo que se adjunta ------
                
                Contador = Contador + 1
                'AlertaFileName = dirsalidas & "\ale_" & Replace(FormatDateTime(Date, 2), "/", "-") & "_" & Replace(FormatDateTime(Time, 4), ":", "-") & "-" & String(2 - Len(Second(Now)), "0") & Second(Now) & Contador
                'FGZ - 05/09/2006
                AlertaFileName = dirsalidas & "\msg_" & BProNro & "_ale_" & Replace(FormatDateTime(Date, 2), "/", "-") & "_" & Replace(FormatDateTime(Time, 4), ":", "-") & "-" & String(2 - Len(Second(Now)), "0") & Second(Now) & Contador
                
                'creo el proceso que se encarga de mandar el mail
                If (rsenviar = -1) Then
                    'Creo el archivo a atachar en el mail
                    Set fs2 = CreateObject("Scripting.FileSystemObject")
                    Set AlertaFile = fs2.CreateTextFile(AlertaFileName & ".html", True)
                    Flog.writeline "Fin Alerta: Positiva. Enviado a: " & Mails
                    
                    AlertaFile.writeline "<html><head>"
                    AlertaFile.writeline "<STYLE>TABLE{ border : thick solid 1; width : 100%;}TH{ background-color: #333399; COLOR: #ffffff; FONT-FAMILY: 'Arial'; FONT-SIZE: 9pt; FONT-WEIGHT: bold; padding : 2 2 2 5; width : auto;}"
                    AlertaFile.writeline "TR{ COLOR: black; FONT-FAMILY: Verdana; FONT-SIZE: 08pt; BACKGROUND-COLOR: #E4FEF9; padding : 2; padding-left : 5;}h4{font-family : Verdana, Geneva, Arial, Helvetica, sans-serif;font-size : smaller;font-style : normal;color : Maroon;}</STYLE>"
                    AlertaFile.writeline "<title>Alerta - Busqueda &reg;</title></head><body>"
                    AlertaFile.writeline "<h4>" & TituloAlerta & "</h4>"
                    AlertaFile.writeline "<table>" & Mensaje & "</table>"
                    AlertaFile.writeline "</body></html>"
                    AlertaFile.Close
                    
                    Call crearProcesosMensajeriaTLP(aleNro, Mails, AlertaFileName, AlertaFileName, Texto_Body)
                Else
                    Call crearProcesosMensajeriaTLP(aleNro, Mails, AlertaFileName, "", Texto_Body)
                End If
            Else
                Flog.writeline "Alerta: No se encontraron mails definidos en el resultado de la consulta. No se han enviado mensajes."
            End If
        End If
        objRs.Close
        resultSet.MoveNext
    Loop
    
    If rs_Link.State = adStateOpen Then rs_Link.Close
    Set rs_Link = Nothing
    
    Exit Sub
CE:
    HuboErrores = True
    Flog.writeline " Error: " & Err.Description & " > " & Now
End Sub


Function emailSupervisor(ByVal empleg As Long) As String

Dim StrSql As String
Dim objRs As New ADODB.Recordset

    StrSql = "SELECT jefe.empemail FROM empleado "
    StrSql = StrSql & "INNER JOIN empleado jefe ON empleado.empreporta = jefe.ternro "
    StrSql = StrSql & "WHERE Empleado.empleg = " & empleg
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        emailSupervisor = ""
    Else
        emailSupervisor = objRs!empemail
    End If
    objRs.Close
End Function

'-----------------------------------------------------------------------------------------------
' Se encarga de enviar el resultado de la alerta a los empleados definidos en la tabla noti_empleado
'-----------------------------------------------------------------------------------------------
Sub MailsASupervisor(ByVal aleNro As Integer, ByVal notinro As Integer, ByVal rsenviar As Integer, ByRef resultSet As ADODB.Recordset)

Dim colcount As Long
'Dim Mensaje As String
Dim titulos As String
Dim Mails As String
Dim I As Long
Dim objRs As New ADODB.Recordset
Dim AlertaFileName As String
Dim fs2, MsgFile
Dim AlertaFile
Dim Clave As String
Dim Enviar As Boolean
Dim rs_Env As New ADODB.Recordset

    On Error GoTo CE
    Enviar = False
    
    colcount = resultSet.Fields.Count
    'Escribo el titulo de las columnas
    titulos = "<TR>" & vbCrLf
    For I = 0 To (colcount - 1)
        titulos = titulos & "<TH>" & resultSet.Fields(I).Name & "</TH>" & vbCrLf
    Next
    titulos = titulos & "</TR>" & vbCrLf
    
    ' el primer campo de la consulta es el ternro del empleado
    resultSet.MoveFirst
    Do Until resultSet.EOF
        'Escribo los datos de la tabla
        Mensaje = "<TR>" & vbCrLf
        Clave = ""
        For I = 0 To (colcount - 1)
            Mensaje = Mensaje & "<TD>" & resultSet.Fields(I).Value & "</TD>" & vbCrLf
            Clave = Clave & resultSet.Fields(I).Value & " "
        Next
        Mensaje = Mensaje & "</TR>" & vbCrLf
        
        Mails = emailSupervisor(resultSet.Fields(0).Value)
        
        'FGZ - 30/03/2009 - se le agregó esto para que registre las alertas enviadas ----
        If AleNotiUnicaVez Then
            'reviso que no haya sido notificada ya
            StrSql = "SELECT * FROM ale_enviada "
            StrSql = StrSql & " WHERE alenro = " & aleNro
            StrSql = StrSql & " AND notinro = " & notinro
            StrSql = StrSql & " AND mail = '" & Mails & "'"
            StrSql = StrSql & " AND clave = '" & Clave & "'"
            OpenRecordset StrSql, rs_Env
            If rs_Env.EOF Then
                StrSql = "INSERT INTO ale_enviada "
                StrSql = StrSql & "(alenro, notinro, fecha, mail, envios, clave) VALUES "
                StrSql = StrSql & "("
                StrSql = StrSql & aleNro & ","
                StrSql = StrSql & notinro & ","
                StrSql = StrSql & ConvFecha(Date) & ","
                'FGZ - 03/06/2014 -----------------------------
                StrSql = StrSql & "'" & Left(Mails, 100) & "',"
                StrSql = StrSql & 0 & ","
                StrSql = StrSql & "'" & Left(Clave, 1000) & "'"
                StrSql = StrSql & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
                
                Enviar = True
            Else
                Flog.writeline "Alerta: ya ha sido notificada anteriormente el " & rs_Env!Fecha & ". No se reenviará."
                Enviar = False
            End If
        Else
            Enviar = True
        End If
        'FGZ - 30/03/2009 - se le agregó esto para que registre las alertas enviadas ----
        
        
        If Mails <> "" And Enviar Then
            Contador = Contador + 1
            'AlertaFileName = dirsalidas & "\ale_" & Replace(FormatDateTime(Date, 2), "/", "-") & "_" & Replace(FormatDateTime(Time, 4), ":", "-") & "-" & String(2 - Len(Second(Now)), "0") & Second(Now) & Contador
            'FGZ - 05/09/2006
            AlertaFileName = dirsalidas & "\msg_" & BProNro & "_ale_" & Replace(FormatDateTime(Date, 2), "/", "-") & "_" & Replace(FormatDateTime(Time, 4), ":", "-") & "-" & String(2 - Len(Second(Now)), "0") & Second(Now) & Contador
            AlertaFileName = AlertaFileName & "-" & resultSet.Fields(0).Value
            'creo el proceso que se encarga de mandar el mail
            If (rsenviar = -1) Then
                'Creo el archivo a atachar en el mail
                Set fs2 = CreateObject("Scripting.FileSystemObject")
                Set AlertaFile = fs2.CreateTextFile(AlertaFileName & ".html", True)
                Flog.writeline "Fin Alerta: Positiva. Enviado a: " & Mails
                'AlertaFile.writeline "<html><head>"
                'AlertaFile.writeline "<STYLE>TABLE{ border : thick solid 1; width : 100%;}TH{ background-color: #333399; COLOR: #ffffff; FONT-FAMILY: 'Arial'; FONT-SIZE: 9pt; FONT-WEIGHT: bold; padding : 2 2 2 5; width : auto;}"
                'AlertaFile.writeline "TR{ COLOR: black; FONT-FAMILY: Verdana; FONT-SIZE: 08pt; BACKGROUND-COLOR: #E4FEF9; padding : 2; padding-left : 5;}h4{font-family : Verdana, Geneva, Arial, Helvetica, sans-serif;font-size : smaller;font-style : normal;color : Maroon;}</STYLE>"
                'AlertaFile.writeline "<title>Alertas - RHPro &reg;</title></head><body>"
                'AlertaFile.writeline "<h4>" & TituloAlerta & "</h4>"
                'AlertaFile.writeline "<table class=""tabladetalle"">" & titulos & Mensaje & "</table>"
                'AlertaFile.writeline "</body></html>"
                'AlertaFile.Close
                
                AlertaFile.writeline "<html><head>"
                AlertaFile.writeline "<STYLE>" & Estilo & "</STYLE>"
                AlertaFile.writeline "<title> Alertas - RHPro &reg; </title></head><body>"
                AlertaFile.writeline "<h4>" & TituloAlerta & "</h4>"
                AlertaFile.writeline "<table class=""tabladetalle"">" & titulos & Mensaje & "</table>"
                AlertaFile.writeline "</body></html>"
                AlertaFile.Close
                
                
                Call crearProcesosMensajeria(aleNro, Mails, AlertaFileName, AlertaFileName)
            Else
                Call crearProcesosMensajeria(aleNro, Mails, AlertaFileName, "")
            End If
        End If
        resultSet.MoveNext
    Loop
    Exit Sub
CE:
    HuboErrores = True
    Flog.writeline " Error: " & Err.Description & " > " & Now
End Sub


'-----------------------------------------------------------------------------------------------
' Se encarga de enviar el resultado de la alerta a los usuarios definidos en la tabla ale_usuario
'-----------------------------------------------------------------------------------------------
Sub MailsAUsuario(ByVal aleNro As Integer, ByVal notinro As Integer, ByVal rsenviar As Integer, ByRef resultSet As ADODB.Recordset)

Dim colcount As Long
'Dim Mensaje As String
Dim Mails As String
Dim I As Long
Dim objRs As New ADODB.Recordset
Dim AlertaFileName As String
Dim fs2, MsgFile
Dim AlertaFile
Dim Enviar As Boolean
Dim rs_Env As New ADODB.Recordset
Dim Clave() As String
Dim Indice_rs As Long

    On Error GoTo CE
    Enviar = False
    Indice_rs = 0
    ReDim Preserve Clave(1 To resultSet.RecordCount) As String
    
    colcount = resultSet.Fields.Count
    'Escribo el titulo de las columnas
    Mensaje = "<TABLE class=""tabladetalle"">" & vbCrLf
    Mensaje = Mensaje & "<TR>" & vbCrLf
    For I = 0 To (colcount - 1)
        Mensaje = Mensaje & "<TH>" & resultSet.Fields(I).Name & "</TH>" & vbCrLf
    Next
    
    'Escribo los datos de la tabla
    Mensaje = Mensaje & "</TR>" & vbCrLf
    Do Until resultSet.EOF
        Indice_rs = Indice_rs + 1
        Mensaje = Mensaje & "<TR>" & vbCrLf
        For I = 0 To (colcount - 1)
            Mensaje = Mensaje & "<TD>" & resultSet.Fields(I).Value & "</TD>" & vbCrLf
            Clave(Indice_rs) = Clave(Indice_rs) & resultSet.Fields(I).Value & " "
        Next
        Mensaje = Mensaje & "</TR>" & vbCrLf
        resultSet.MoveNext
    Loop
    
    'MDZ - 06/06/2013 -  se corrigio la que sigue  CAS-18351 -  TIMBO - Alertas - Bug en alertas
    'Mensaje = "</TABLE>" & vbCrLf
    Mensaje = Mensaje + "</TABLE>" & vbCrLf
            
    'Busco todos los usuarios a los cuales les tengo que enviar los mails
    StrSql = "SELECT usremail FROM user_per "
    StrSql = StrSql & "inner join noti_usuario on user_per.iduser = noti_usuario.iduser "
    StrSql = StrSql & "where notinro = " & notinro
    OpenRecordset StrSql, objRs
    Mails = ""
    Do Until objRs.EOF
        If Not IsNull(objRs!usremail) Then
            If Len(objRs!usremail) > 0 Then
                Mails = Mails & objRs!usremail & ";"
                
                'FGZ - 30/03/2009 - se le agregó esto para que registre las alertas enviadas ----
                If AleNotiUnicaVez Then
                    For I = 1 To (Indice_rs)
                        'reviso que no haya sido notificada ya
                        StrSql = "SELECT * FROM ale_enviada "
                        StrSql = StrSql & " WHERE alenro = " & aleNro
                        StrSql = StrSql & " AND notinro = " & notinro
                        StrSql = StrSql & " AND mail = '" & objRs!usremail & "'"
                        StrSql = StrSql & " AND clave = '" & Clave(I) & "'"
                        OpenRecordset StrSql, rs_Env
                        If rs_Env.EOF Then
                            StrSql = "INSERT INTO ale_enviada "
                            StrSql = StrSql & "(alenro, notinro, fecha, mail, envios, clave) VALUES "
                            StrSql = StrSql & "("
                            StrSql = StrSql & aleNro & ","
                            StrSql = StrSql & notinro & ","
                            StrSql = StrSql & ConvFecha(Date) & ","
                            'FGZ - 03/06/2014 -----------------------------
                            StrSql = StrSql & "'" & Left(objRs!usremail, 100) & "',"
                            StrSql = StrSql & 0 & ","
                            StrSql = StrSql & "'" & Left(Clave(I), 1000) & "'"
                            StrSql = StrSql & ")"
                            objConn.Execute StrSql, , adExecuteNoRecords
                            
                            Enviar = Enviar Or True
                        Else
                            Flog.writeline "Alerta: ya ha sido notificada anteriormente el " & rs_Env!Fecha & ". No se reenviará."
                            Enviar = Enviar Or False
                        End If
                    Next I
                Else
                    Enviar = True
                End If
                'FGZ - 30/03/2009 - se le agregó esto para que registre las alertas enviadas ----
           End If
        End If
        objRs.MoveNext
    Loop
    objRs.Close
    
    If Mails <> "" And Enviar Then
        'creo el proceso que se encarga de mandar el mail
        Contador = Contador + 1
        'AlertaFileName = dirsalidas & "\ale_" & Replace(FormatDateTime(Date, 2), "/", "-") & "_" & Replace(FormatDateTime(Time, 4), ":", "-") & "-" & String(2 - Len(Second(Now)), "0") & Second(Now) & Contador
        'FGZ - 05/09/2006
        AlertaFileName = dirsalidas & "\msg_" & BProNro & "_ale_" & Replace(FormatDateTime(Date, 2), "/", "-") & "_" & Replace(FormatDateTime(Time, 4), ":", "-") & "-" & String(2 - Len(Second(Now)), "0") & Second(Now) & Contador
        
        If (rsenviar = -1) Then
            'Creo el archivo a atachar en el mail
            
            Set fs2 = CreateObject("Scripting.FileSystemObject")
            Set AlertaFile = fs2.CreateTextFile(AlertaFileName & ".html", True)
            
            Flog.writeline "Fin Alerta: Positiva."
            'AlertaFile.writeline "<html><head>"
            'AlertaFile.writeline "<STYLE>TABLE{ border : thick solid 1; width : 100%;}TH{ background-color: #333399; COLOR: #ffffff; FONT-FAMILY: 'Arial'; FONT-SIZE: 9pt; FONT-WEIGHT: bold; padding : 2 2 2 5; width : auto;}"
            'AlertaFile.writeline "TR{ COLOR: black; FONT-FAMILY: Verdana; FONT-SIZE: 08pt; BACKGROUND-COLOR: #E4FEF9; padding : 2; padding-left : 5;}h4{font-family : Verdana, Geneva, Arial, Helvetica, sans-serif;font-size : smaller;font-style : normal;color : Maroon;}</STYLE>"
            'AlertaFile.writeline "<title>Alertas - RHPro &reg;</title></head><body>"
            'AlertaFile.writeline "<h4>" & TituloAlerta & "</h4>"
            'AlertaFile.writeline "<table>" & Mensaje & "</table>"
            'AlertaFile.writeline "</body></html>"
            'AlertaFile.Close
            
            AlertaFile.writeline "<html><head>"
            AlertaFile.writeline "<STYLE>" & Estilo & "</STYLE>"
            AlertaFile.writeline "<title> Alertas - RHPro &reg; </title></head><body>"
            AlertaFile.writeline "<h4>" & TituloAlerta & "</h4>"
            AlertaFile.writeline Mensaje
            AlertaFile.writeline "</body></html>"
            AlertaFile.Close
            
            
            Call crearProcesosMensajeria(aleNro, Mails, AlertaFileName, AlertaFileName)
        Else
            Call crearProcesosMensajeria(aleNro, Mails, AlertaFileName, "")
        End If
    Else
        Flog.writeline "Alerta: No se encontraron mails definidos en los supervisores. No se han enviado mensajes."
    End If
    Exit Sub
CE:
    HuboErrores = True
    Flog.writeline " Error: " & Err.Description & " > " & Now
End Sub


'-----------------------------------------------------------------------------------------------
' Se encarga de enviar el resultado de la alerta a los usuarios y grupos definidos en la tabla noti_usuario
'-----------------------------------------------------------------------------------------------
Sub MailsAGrupoDeUsuario(ByVal aleNro As Integer, ByVal notinro As Integer, ByVal rsenviar As Integer, ByRef resultSet As ADODB.Recordset)

    On Error GoTo CE
    
    Dim colcount As Long
    'Dim Mensaje As String
    Dim Mails As String
    Dim I As Long
    Dim objRs As New ADODB.Recordset
    Dim AlertaFileName As String
    Dim fs2, MsgFile
    Dim AlertaFile
    Dim Enviar As Boolean
    Dim rs_Env As New ADODB.Recordset
    Dim Clave() As String
    Dim Indice_rs As Long
    Dim Tabla_p1 As String
    Dim Tabla_p2 As String
    Dim grupo_old As String
    
    If LCase(resultSet.Fields(0).Name) <> "_grupo" Then
        Flog.writeline "El query de la Alerta no esta configurado para el Tipo de notificacion Seleccionado!"
        HuboErrores = True
        Exit Sub
    End If
    
    Enviar = False
    Indice_rs = 0
    
    
    colcount = resultSet.Fields.Count
    'Escribo el titulo de las columnas
    Tabla_p1 = "<TABLE class=""tabladetalle"">" & vbCrLf
    Tabla_p1 = Tabla_p1 & "<TR>" & vbCrLf
    For I = 1 To (colcount - 1)
        Tabla_p1 = Tabla_p1 & "<TH>" & resultSet.Fields(I).Name & "</TH>" & vbCrLf
    Next
    Tabla_p1 = Tabla_p1 & "</TR>" & vbCrLf
        
    resultSet.Sort = "_grupo ASC"
        
Do
    grupo_old = "|-|"
    Tabla_p2 = ""
    Indice_rs = 0
    Enviar = False
    ReDim Clave(1 To resultSet.RecordCount) As String
    
    'Escribo los datos de la tabla
    Do Until resultSet.EOF
        Indice_rs = Indice_rs + 1
        
        Tabla_p2 = Tabla_p2 & "<TR>" & vbCrLf
        For I = 1 To (colcount - 1)
            Tabla_p2 = Tabla_p2 & "<TD>" & resultSet.Fields(I).Value & "</TD>" & vbCrLf
            
        Next
        Tabla_p2 = Tabla_p2 & "</TR>" & vbCrLf
                        
        If Not IsNull(resultSet.Fields(0).Value) Then
            grupo_old = LCase(resultSet.Fields(0).Value)
        End If
        
        resultSet.MoveNext
        
        If Not resultSet.EOF And grupo_old <> "|-|" Then
            
            If grupo_old <> LCase(resultSet.Fields(0).Value) Then
                Exit Do
            End If
            
            For I = 1 To (colcount - 1)
                Clave(Indice_rs) = Clave(Indice_rs) & resultSet.Fields(I).Value & " "
            Next
            Clave(Indice_rs) = grupo_old & Clave(Indice_rs)
            
        End If
        
        'Flog.writeline "grupo =" & grupo_old & "  Clave  : " & Clave(Indice_rs)
    Loop

    If grupo_old = "|-|" Then
        'ningun resulset
        'Flog.writeline "nada que enviar!"
        Exit Sub
    End If
    
    Mensaje = Tabla_p1 & Tabla_p2 & "</TABLE>" & vbCrLf
    
    Flog.writeline "Notificacion para grupo " & grupo_old
    
    'Busco todos los usuarios a los cuales les tengo que enviar los mails
    StrSql = "SELECT usremail FROM user_per "
    StrSql = StrSql & "inner join noti_usuario on user_per.iduser = noti_usuario.iduser "
    StrSql = StrSql & "where notinro = " & notinro & " AND lower(noti_usuario.grupo)='" & grupo_old & "'"
    OpenRecordset StrSql, objRs
    
    'Flog.writeline StrSql
    
    'si no hay grupo definido uso "por defecto"
    If objRs.EOF Then
        objRs.Close
        
        Flog.writeline "No se encontraron usuarios para el grupo. Se envia notificacion a usuarios de grupo por defecto."
        
        StrSql = "SELECT usremail FROM user_per "
        StrSql = StrSql & "inner join noti_usuario on user_per.iduser = noti_usuario.iduser "
        StrSql = StrSql & "where notinro = " & notinro & " AND (noti_usuario.grupo='' OR noti_usuario.grupo IS null)"
        OpenRecordset StrSql, objRs
        'Flog.writeline StrSql
        
    End If
    
    Mails = ""
    Enviar = True
    
    Do Until objRs.EOF
        If Not IsNull(objRs!usremail) Then
            If Len(objRs!usremail) > 0 Then
                Mails = Mails & objRs!usremail & ";"
                
                If AleNotiUnicaVez Then
                   For I = 1 To Indice_rs - 1
                        If Clave(I) <> "" Then
                            'reviso que no haya sido notificada ya
                            StrSql = "SELECT * FROM ale_enviada "
                            StrSql = StrSql & " WHERE alenro = " & aleNro
                            StrSql = StrSql & " AND notinro = " & notinro
                            StrSql = StrSql & " AND mail = '" & objRs!usremail & "'"
                            StrSql = StrSql & " AND clave = '" & Clave(I) & "'"
                            OpenRecordset StrSql, rs_Env
                            If rs_Env.EOF Then
                                StrSql = "INSERT INTO ale_enviada "
                                StrSql = StrSql & "(alenro, notinro, fecha, mail, envios, clave) VALUES "
                                StrSql = StrSql & "("
                                StrSql = StrSql & aleNro & ","
                                StrSql = StrSql & notinro & ","
                                StrSql = StrSql & ConvFecha(Date) & ","
                                'FGZ - 03/06/2014 -----------------------------
                                StrSql = StrSql & "'" & Left(objRs!usremail, 100) & "',"
                                StrSql = StrSql & 0 & ","
                                StrSql = StrSql & "'" & Left(Clave(I), 1000) & "'"
                                StrSql = StrSql & ")"
                                objConn.Execute StrSql, , adExecuteNoRecords
                                Enviar = True
                                
                            Else
                                Flog.writeline "Alerta: ya ha sido notificada anteriormente el " & rs_Env!Fecha & ". No se reenviará."
                                Enviar = False
                                Exit For
                            End If
                        End If
                   Next
                End If
            End If
        End If
        objRs.MoveNext
    Loop
    objRs.Close
    
    
    If Mails <> "" And Enviar Then
        
        Flog.writeline "Se envia email a " & Mails
        
        'creo el proceso que se encarga de mandar el mail
        Contador = Contador + 1
        'AlertaFileName = dirsalidas & "\ale_" & Replace(FormatDateTime(Date, 2), "/", "-") & "_" & Replace(FormatDateTime(Time, 4), ":", "-") & "-" & String(2 - Len(Second(Now)), "0") & Second(Now) & Contador
        'FGZ - 05/09/2006
        AlertaFileName = dirsalidas & "\msg_" & BProNro & "_ale_" & Replace(FormatDateTime(Date, 2), "/", "-") & "_" & Replace(FormatDateTime(Time, 4), ":", "-") & "-" & String(2 - Len(Second(Now)), "0") & Second(Now) & Contador
        Flog.writeline " se genera atach ... " & AlertaFileName
        
        If (rsenviar = -1) Then
                'Creo el archivo a atachar en el mail
                
                Set fs2 = CreateObject("Scripting.FileSystemObject")
                Set AlertaFile = fs2.CreateTextFile(AlertaFileName & ".html", True)
                
                Flog.writeline "Fin Alerta: Positiva."
                AlertaFile.writeline "<html><head>"
                AlertaFile.writeline "<STYLE>" & Estilo & "</STYLE>"
                AlertaFile.writeline "<title> Alertas - RHPro &reg; </title></head><body>"
                AlertaFile.writeline "<h4>" & TituloAlerta & "</h4>"
                AlertaFile.writeline Mensaje
                AlertaFile.writeline "</body></html>"
                AlertaFile.Close
                
                Call crearProcesosMensajeria(aleNro, Mails, AlertaFileName, AlertaFileName)
        Else
                Call crearProcesosMensajeria(aleNro, Mails, AlertaFileName, "")
        End If
    Else
        If Mails = "" Then
            Flog.writeline "Alerta: no se encontraron destinatarios para el grupo " & grupo_old
        End If
    End If

Loop
        
    Exit Sub
CE:
    HuboErrores = True
    Flog.writeline " Error: " & Err.Description & " > " & Now
End Sub


                
                
'-----------------------------------------------------------------------------------------------
' Se encarga de enviar un mail a cada rol de la tabla ale_roles
'-----------------------------------------------------------------------------------------------
Sub MailsARoles(ByVal aleNro As Integer, ByVal notinro As Integer, ByVal rsenviar As Integer, ByRef resultSet As ADODB.Recordset)

Dim colcount As Long
'Dim Mensaje As String
Dim Mails As String
Dim I As Long
Dim objRs As New ADODB.Recordset
Dim AlertaFileName As String
Dim fs2, MsgFile, AlertaFile
Dim hayCambio
Dim evaluadorAnt
Dim rolAnt
Dim titulos
Dim Datos
Dim Clave As String
Dim Enviar As Boolean
Dim rs_Env As New ADODB.Recordset

    On Error GoTo CE
    Enviar = False
    
    colcount = resultSet.Fields.Count
    'Escribo el titulo de las columnas
    titulos = ""
    For I = 2 To (colcount - 1)
        titulos = titulos & "<TH>" & resultSet.Fields(I).Name & "</TH>" & vbCrLf
    Next
    
    Datos = ""
    Do Until resultSet.EOF
        'En el result set la primer columna tiene que ser el ternro y la segunda el rol
        evaluadorAnt = resultSet(0)
        rolAnt = resultSet(1)
        
        'Si hay columna de datos entonces guardo sus valores
        If (colcount - 2) > 0 Then
            Datos = Datos & "<TR>" & vbCrLf
            Clave = ""
            For I = 2 To (colcount - 1)
                Datos = Datos & "<TD>" & resultSet.Fields(I).Value & "</TD>" & vbCrLf
                Clave = Clave & resultSet.Fields(I).Value & " "
            Next
            Datos = Datos & "</TR>" & vbCrLf
        End If
        
        resultSet.MoveNext
        
        'Controlo si hay un cambio de evaluador y rol
        If resultSet.EOF Then
           hayCambio = True
        Else
           hayCambio = (resultSet(0) <> evaluadorAnt) Or (resultSet(1) <> rolAnt)
        End If
        
        'Si hay un cambio le envio un mail al evaluador
        If hayCambio Then
            If Not IsNull(rolAnt) Then
                'Controlo si el rol esta seleccionado en la configuracion de la alerta
                StrSql = "select * from noti_roles "
                StrSql = StrSql & " WHERE notinro = " & notinro & " "
                StrSql = StrSql & " AND evatevnro =  " & rolAnt
                
                OpenRecordset StrSql, objRs
                
                If Not objRs.EOF Then
                    If Not IsNull(evaluadorAnt) Then
                         objRs.Close
                        
                         'Busco el mail del empleado
                         StrSql = "select * from empleado "
                         StrSql = StrSql & " WHERE ternro = " & evaluadorAnt
                         
                         OpenRecordset StrSql, objRs
                         
                         If Not objRs.EOF Then
                            If Not IsNull(objRs!empemail) Then
                               If Len(objRs!empemail) > 0 Then
                                    Mails = objRs!empemail & ";"
                                  
                                    'FGZ - 30/03/2009 - se le agregó esto para que registre las alertas enviadas ----
                                    If AleNotiUnicaVez Then
                                        'reviso que no haya sido notificada ya
                                        StrSql = "SELECT * FROM ale_enviada "
                                        StrSql = StrSql & " WHERE alenro = " & aleNro
                                        StrSql = StrSql & " AND notinro = " & notinro
                                        StrSql = StrSql & " AND mail = '" & objRs!empemail & "'"
                                        StrSql = StrSql & " AND clave = '" & Clave & "'"
                                        OpenRecordset StrSql, rs_Env
                                        If rs_Env.EOF Then
                                            StrSql = "INSERT INTO ale_enviada "
                                            StrSql = StrSql & "(alenro, notinro, fecha, mail, envios, clave) VALUES "
                                            StrSql = StrSql & "("
                                            StrSql = StrSql & aleNro & ","
                                            StrSql = StrSql & notinro & ","
                                            StrSql = StrSql & ConvFecha(Date) & ","
                                            'FGZ - 03/06/2014 -----------------------------
                                            StrSql = StrSql & "'" & Left(objRs!empemail, 100) & "',"
                                            StrSql = StrSql & 0 & ","
                                            StrSql = StrSql & "'" & Left(Clave, 1000) & "'"
                                            StrSql = StrSql & ")"
                                            objConn.Execute StrSql, , adExecuteNoRecords
                                            
                                            Enviar = True
                                        Else
                                            Flog.writeline "Alerta: ya ha sido notificada anteriormente el " & rs_Env!Fecha & ". No se reenviará."
                                            Enviar = False
                                        End If
                                    Else
                                        Enviar = True
                                    End If
                                    'FGZ - 30/03/2009 - se le agregó esto para que registre las alertas enviadas ----
                                    
                                    If Enviar Then
                                        'Creo el archivo a atachar en el mail
                                        'AlertaFileName = dirsalidas & "\ale_" & Replace(FormatDateTime(Date, 2), "/", "-") & "_" & Replace(FormatDateTime(Time, 4), ":", "-") & "-" & String(2 - Len(Second(Now)), "0") & Second(Now)
                                        'FGZ - 05/09/2006
                                        AlertaFileName = dirsalidas & "\msg_" & BProNro & "_ale_" & Replace(FormatDateTime(Date, 2), "/", "-") & "_" & Replace(FormatDateTime(Time, 4), ":", "-") & "-" & String(2 - Len(Second(Now)), "0") & Second(Now) & Contador
    
                                        
                                        Set fs2 = CreateObject("Scripting.FileSystemObject")
                                        Contador = Contador + 1
                                        AlertaFileName = dirsalidas & "\ale_" & Replace(FormatDateTime(Date, 2), "/", "-") & "_" & Replace(FormatDateTime(Time, 4), ":", "-") & "-" & String(2 - Len(Second(Now)), "0") & Second(Now) & Contador
                                        
                                        'AlertaFile.writeline "<html><head>"
                                        'AlertaFile.writeline "<STYLE>TABLE{ border : thick solid 1; width : 100%;}TH{ background-color: #333399; COLOR: #ffffff; FONT-FAMILY: 'Arial'; FONT-SIZE: 9pt; FONT-WEIGHT: bold; padding : 2 2 2 5; width : auto;}"
                                        'AlertaFile.writeline "TR{ COLOR: black; FONT-FAMILY: Verdana; FONT-SIZE: 08pt; BACKGROUND-COLOR: #E4FEF9; padding : 2; padding-left : 5;}h4{font-family : Verdana, Geneva, Arial, Helvetica, sans-serif;font-size : smaller;font-style : normal;color : Maroon;}</STYLE>"
                                        'AlertaFile.writeline "<title>Alertas - RHPro &reg;</title></head><body>"
                                        'AlertaFile.writeline "<table>"
                                        'AlertaFile.writeline "<h4>" & titulos & "</h4>"
                                        'AlertaFile.writeline Datos
                                        'AlertaFile.writeline "</table>"
                                        'AlertaFile.writeline "</body></html>"
                                        'AlertaFile.Close
                                        
                                        AlertaFile.writeline "<html><head>"
                                        AlertaFile.writeline "<STYLE>" & Estilo & "</STYLE>"
                                        AlertaFile.writeline "<title> Alertas - RHPro &reg; </title></head><body>"
                                        AlertaFile.writeline "<h4>" & TituloAlerta & "</h4>"
                                        AlertaFile.writeline "<TABLE class=""tabladetalle"">"
                                        AlertaFile.writeline Datos
                                        AlertaFile.writeline "</TABLE>"
                                        AlertaFile.writeline "</body></html>"
                                        AlertaFile.Close
                                        
                                        
                                        'creo el proceso que se encarga de mandar el mail
                                        'Si hay datos creo un attach con los datos, sino envio un mail sin attach
                                        If (Datos <> "") And (rsenviar = -1) Then
                                            Call crearProcesosMensajeria(aleNro, Mails, AlertaFileName, AlertaFileName)
                                        Else
                                            Call crearProcesosMensajeria(aleNro, Mails, AlertaFileName, "")
                                        End If
                                    Else
                                    
                                    End If
                                Else
                                    Flog.writeline "Alerta: No esta definido el mail para el empleado: " & evaluadorAnt
                                End If
                            End If
                         Else
                            Flog.writeline "No se encontro el empleado con numero " & evaluadorAnt
                         End If
                    End If
                End If
                
                objRs.Close
            End If
        
            If Not resultSet.EOF Then
               evaluadorAnt = resultSet(0)
               rolAnt = resultSet(1)
               Datos = ""
            End If

        End If
        
    Loop
    Exit Sub
CE:
    HuboErrores = True
    Flog.writeline " Error: " & Err.Description & " > " & Now
End Sub

'-----------------------------------------------------------------------------------------------
' Se encarga de ejecutar la query correspondiente y mandar el resultado por mail a los distintos destinatarios
' Se debe incluir el campo tercero en el primer lugar del SELECT
'-----------------------------------------------------------------------------------------------
Sub MailsAQueryDelAlerta(ByVal aleNro As Integer, ByVal notinro As Integer, ByVal rsenviar As Integer, ByRef resultSet As ADODB.Recordset)
Dim I As Integer
Dim colcount As Integer
'Dim Mensaje As String
Dim objRs As New ADODB.Recordset
Dim strSQLQuery As String
Dim Param As String
Dim Opcion As Integer
Dim campo As Integer
Dim reemplazo As String
Dim Mails As String
Dim AlertaFileName As String
Dim fs2, AlertaFile
Dim Clave As String
Dim Enviar As Boolean
Dim rs_Env As New ADODB.Recordset
Dim Indice_rs As Long

    On Error GoTo CE
    Enviar = False
    
    colcount = resultSet.Fields.Count
    Do Until resultSet.EOF
        'Escribo el titulo de las columnas
        Mensaje = "<TABLE class=""tabladetalle""><TR>" & vbCrLf
        For I = 1 To (colcount - 1)
            Mensaje = Mensaje & "<TH>" & resultSet.Fields(I).Name & "</TH>" & vbCrLf
        Next
        Mensaje = Mensaje & "</TR>" & vbCrLf
        'Escribo el contenido de la fila
        Mensaje = Mensaje & "<TR>" & vbCrLf
        For I = 1 To (colcount - 1)
            Mensaje = Mensaje & "<TD>" & resultSet.Fields(I).Value & "</TD>" & vbCrLf
            Clave = Clave & resultSet.Fields(I).Value & " "
        Next
        Mensaje = Mensaje & "</TR></TABLE>" & vbCrLf

        StrSql = " SELECT teremail "
        StrSql = StrSql & " FROM tercero "
        StrSql = StrSql & " WHERE ternro = " & resultSet.Fields("ternro").Value
        Flog.writeline "Mail del tercero: " & StrSql
        Mails = ""
        Enviar = False
        OpenRecordset StrSql, objRs
        Do Until objRs.EOF
            ' la primera columna del query es una dir de email
            If Not IsNull(objRs.Fields(0).Value) Then
                If Len(objRs.Fields(0).Value) > 0 Then
                    Mails = Mails & objRs.Fields(0).Value & ";"
                    
                    'FGZ - 30/03/2009 - se le agregó esto para que registre las alertas enviadas ----
                    If AleNotiUnicaVez Then
                        'reviso que no haya sido notificada ya
                        StrSql = "SELECT * FROM ale_enviada "
                        StrSql = StrSql & " WHERE alenro = " & aleNro
                        StrSql = StrSql & " AND notinro = " & notinro
                        StrSql = StrSql & " AND mail = '" & objRs.Fields(0).Value & "'"
                        StrSql = StrSql & " AND clave = '" & Clave & "'"
                        OpenRecordset StrSql, rs_Env
                        If rs_Env.EOF Then
                            StrSql = "INSERT INTO ale_enviada "
                            StrSql = StrSql & "(alenro, notinro, fecha, mail, envios, clave) VALUES "
                            StrSql = StrSql & "("
                            StrSql = StrSql & aleNro & ","
                            StrSql = StrSql & notinro & ","
                            StrSql = StrSql & ConvFecha(Date) & ","
                            'FGZ - 03/06/2014 -----------------------------
                            StrSql = StrSql & "'" & Left(objRs.Fields(0).Value, 100) & "',"
                            StrSql = StrSql & 0 & ","
                            StrSql = StrSql & "'" & Left(Clave, 1000) & "'"
                            StrSql = StrSql & ")"
                            objConn.Execute StrSql, , adExecuteNoRecords
                            
                            Enviar = Enviar Or True
                        Else
                            Flog.writeline "Alerta: ya ha sido notificada anteriormente el " & rs_Env!Fecha & ". No se reenviará."
                            Enviar = Enviar Or False
                        End If
                    Else
                        Enviar = True
                    End If
                    'FGZ - 30/03/2009 - se le agregó esto para que registre las alertas enviadas ----
                    
                    
                End If
            End If
            objRs.MoveNext
        Loop
        
        If Mails <> "" And Enviar Then
            Contador = Contador + 1
            
            AlertaFileName = dirsalidas & "\msg_" & BProNro & "_ale_" & Replace(FormatDateTime(Date, 2), "/", "-") & "_" & Replace(FormatDateTime(Time, 4), ":", "-") & "-" & String(2 - Len(Second(Now)), "0") & Second(Now) & Contador
            
            'creo el proceso que se encarga de mandar el mail
            If (rsenviar = -1) Then
                'Creo el archivo a atachar en el mail
                Set fs2 = CreateObject("Scripting.FileSystemObject")
                Set AlertaFile = fs2.CreateTextFile(AlertaFileName & ".html", True)
                Flog.writeline "Fin Alerta: Positiva. Enviado a: " & Mails
                'AlertaFile.writeline "<html><head>"
                'AlertaFile.writeline "<STYLE>TABLE{ border : thick solid 1; width : 100%;}TH{ background-color: #333399; COLOR: #ffffff; FONT-FAMILY: 'Arial'; FONT-SIZE: 9pt; FONT-WEIGHT: bold; padding : 2 2 2 5; width : auto;}"
                'AlertaFile.writeline "TR{ COLOR: black; FONT-FAMILY: Verdana; FONT-SIZE: 08pt; BACKGROUND-COLOR: #E4FEF9; padding : 2; padding-left : 5;}h4{font-family : Verdana, Geneva, Arial, Helvetica, sans-serif;font-size : smaller;font-style : normal;color : Maroon;}</STYLE>"
                'AlertaFile.writeline "<title>Alertas - RHPro &reg;</title></head><body>"
                'AlertaFile.writeline "<h4>" & TituloAlerta & "</h4>"
                'AlertaFile.writeline "<table>" & Mensaje & "</table>"
                'AlertaFile.writeline "</body></html>"
                'AlertaFile.Close
                
                AlertaFile.writeline "<html><head>"
                AlertaFile.writeline "<STYLE>" & Estilo & "</STYLE>"
                AlertaFile.writeline "<title> Alertas - RHPro &reg; </title></head><body>"
                AlertaFile.writeline "<h4>" & TituloAlerta & "</h4>"
                AlertaFile.writeline Mensaje
                AlertaFile.writeline "</body></html>"
                AlertaFile.Close
                
                
                Call crearProcesosMensajeria(aleNro, Mails, AlertaFileName, AlertaFileName)
            Else
                Call crearProcesosMensajeria(aleNro, Mails, AlertaFileName, "")
            End If
        Else
            Flog.writeline "Alerta: No se encontraron mails definidos en el resultado de la consulta. No se han enviado mensajes."
        End If
        resultSet.MoveNext
    Loop
    Exit Sub
CE:
    HuboErrores = True
    Flog.writeline " Error: " & Err.Description & " > " & Now
End Sub

'-----------------------------------------------------------------------------------------------
' Se encarga de ejecutar la query correspondiente y mandar el resultado por mail a los distintos destinatarios
' Se debe incluir el campo tercero en el primer lugar del SELECT
'-----------------------------------------------------------------------------------------------
Sub MailsAQueryDelAlertaEmp(ByVal aleNro As Integer, ByVal notinro As Integer, ByVal rsenviar As Integer, ByRef resultSet As ADODB.Recordset)
Dim I As Integer
Dim colcount As Integer
'Dim Mensaje As String
Dim objRs As New ADODB.Recordset
Dim strSQLQuery As String
Dim Param As String
Dim Opcion As Integer
Dim campo As Integer
Dim reemplazo As String
Dim Mails As String
Dim AlertaFileName As String
Dim fs2, AlertaFile
Dim Clave As String
Dim Enviar As Boolean
Dim rs_Env As New ADODB.Recordset
Dim Indice_rs As Long

    On Error GoTo CE
    Enviar = False
    
    colcount = resultSet.Fields.Count
    Do Until resultSet.EOF
        'Escribo el titulo de las columnas
        Mensaje = "<TABLE class=""tabladetalle""><TR>" & vbCrLf
        For I = 1 To (colcount - 1)
            Mensaje = Mensaje & "<TH>" & resultSet.Fields(I).Name & "</TH>" & vbCrLf
        Next
        Mensaje = Mensaje & "</TR>" & vbCrLf
        'Escribo el contenido de la fila
        Mensaje = Mensaje & "<TR>" & vbCrLf
        For I = 1 To (colcount - 1)
            Mensaje = Mensaje & "<TD>" & resultSet.Fields(I).Value & "</TD>" & vbCrLf
            Clave = Clave & resultSet.Fields(I).Value & " "
        Next
        Mensaje = Mensaje & "</TR></TABLE>" & vbCrLf

        StrSql = " SELECT empemail "
        StrSql = StrSql & " FROM empleado "
        StrSql = StrSql & " WHERE ternro = " & resultSet.Fields("ternro").Value
        Flog.writeline "Mail del Empleado: " & StrSql
        Mails = ""
        Enviar = False
        OpenRecordset StrSql, objRs
        Do Until objRs.EOF
            ' la primera columna del query es una dir de email
            If Not IsNull(objRs.Fields(0).Value) Then
                If Len(objRs.Fields(0).Value) > 0 Then
                    Mails = Mails & objRs.Fields(0).Value & ";"
                    
                    'FGZ - 30/03/2009 - se le agregó esto para que registre las alertas enviadas ----
                    If AleNotiUnicaVez Then
                        'reviso que no haya sido notificada ya
                        StrSql = "SELECT * FROM ale_enviada "
                        StrSql = StrSql & " WHERE alenro = " & aleNro
                        StrSql = StrSql & " AND notinro = " & notinro
                        StrSql = StrSql & " AND mail = '" & objRs.Fields(0).Value & "'"
                        StrSql = StrSql & " AND clave = '" & Clave & "'"
                        OpenRecordset StrSql, rs_Env
                        If rs_Env.EOF Then
                            StrSql = "INSERT INTO ale_enviada "
                            StrSql = StrSql & "(alenro, notinro, fecha, mail, envios, clave) VALUES "
                            StrSql = StrSql & "("
                            StrSql = StrSql & aleNro & ","
                            StrSql = StrSql & notinro & ","
                            StrSql = StrSql & ConvFecha(Date) & ","
                            'FGZ - 03/06/2014 -----------------------------
                            StrSql = StrSql & "'" & Left(objRs.Fields(0).Value, 100) & "',"
                            StrSql = StrSql & 0 & ","
                            StrSql = StrSql & "'" & Left(Clave, 1000) & "'"
                            StrSql = StrSql & ")"
                            objConn.Execute StrSql, , adExecuteNoRecords
                            
                            Enviar = Enviar Or True
                        Else
                            Flog.writeline "Alerta: ya ha sido notificada anteriormente el " & rs_Env!Fecha & ". No se reenviará."
                            Enviar = Enviar Or False
                        End If
                    Else
                        Enviar = True
                    End If
                    'FGZ - 30/03/2009 - se le agregó esto para que registre las alertas enviadas ----
                    
                    
                End If
            End If
            objRs.MoveNext
        Loop
        
        If Mails <> "" And Enviar Then
            Contador = Contador + 1
            
            AlertaFileName = dirsalidas & "\msg_" & BProNro & "_ale_" & Replace(FormatDateTime(Date, 2), "/", "-") & "_" & Replace(FormatDateTime(Time, 4), ":", "-") & "-" & String(2 - Len(Second(Now)), "0") & Second(Now) & Contador
            
            'creo el proceso que se encarga de mandar el mail
            If (rsenviar = -1) Then
                'Creo el archivo a atachar en el mail
                Set fs2 = CreateObject("Scripting.FileSystemObject")
                Set AlertaFile = fs2.CreateTextFile(AlertaFileName & ".html", True)
                Flog.writeline "Fin Alerta: Positiva. Enviado a: " & Mails
                'AlertaFile.writeline "<html><head>"
                'AlertaFile.writeline "<STYLE>TABLE{ border : thick solid 1; width : 100%;}TH{ background-color: #333399; COLOR: #ffffff; FONT-FAMILY: 'Arial'; FONT-SIZE: 9pt; FONT-WEIGHT: bold; padding : 2 2 2 5; width : auto;}"
                'AlertaFile.writeline "TR{ COLOR: black; FONT-FAMILY: Verdana; FONT-SIZE: 08pt; BACKGROUND-COLOR: #E4FEF9; padding : 2; padding-left : 5;}h4{font-family : Verdana, Geneva, Arial, Helvetica, sans-serif;font-size : smaller;font-style : normal;color : Maroon;}</STYLE>"
                'AlertaFile.writeline "<title>Alertas - RHPro &reg;</title></head><body>"
                'AlertaFile.writeline "<h4>" & TituloAlerta & "</h4>"
                'AlertaFile.writeline "<table>" & Mensaje & "</table>"
                'AlertaFile.writeline "</body></html>"
                'AlertaFile.Close
                
                AlertaFile.writeline "<html><head>"
                AlertaFile.writeline "<STYLE>" & Estilo & "</STYLE>"
                AlertaFile.writeline "<title> Alertas - RHPro &reg; </title></head><body>"
                AlertaFile.writeline "<h4>" & TituloAlerta & "</h4>"
                AlertaFile.writeline Mensaje
                AlertaFile.writeline "</body></html>"
                AlertaFile.Close
                
                Call crearProcesosMensajeria(aleNro, Mails, AlertaFileName, AlertaFileName)
            Else
                Call crearProcesosMensajeria(aleNro, Mails, AlertaFileName, "")
            End If
        Else
            Flog.writeline "Alerta: No se encontraron mails definidos en el resultado de la consulta. No se han enviado mensajes."
        End If
        resultSet.MoveNext
    Loop
    Exit Sub
CE:
    HuboErrores = True
    Flog.writeline " Error: " & Err.Description & " > " & Now
End Sub


Sub MailsA_X_1Xresultado(ByVal aleNro As Integer, ByVal notinro As Integer, ByVal rsenviar As Integer, ByRef resultSet As ADODB.Recordset, ByVal TnotiNro As Long, ByVal multipleAttach As Boolean)
'-----------------------------------------------------------------------------------------------
' Se encarga de enviar un mail por cada resultado de la alerta
'-----------------------------------------------------------------------------------------------
Dim I As Integer
Dim Desde As Integer
Dim colcount As Integer
Dim objRs As New ADODB.Recordset
Dim strSQLQuery As String
Dim Param As String
Dim Opcion As Integer
Dim campo As Integer
Dim reemplazo As String
Dim Mails As String
Dim AlertaFileName As String
Dim fs2, AlertaFile
Dim Clave() As String
Dim Enviar As Boolean
Dim rs_Env As New ADODB.Recordset
Dim Indice_rs As Long
Dim Resultado As String
Dim Col1 As Long
Dim ArrMails
Dim colcountAux As Long
Dim campoClave As Long
Dim docAttachs As String
Dim rutaAttach As String
Dim evenroAttach As Long
Dim salir As Boolean
Dim corte As Integer
Dim corteAnt As Integer
Dim agrupa As Integer
Dim notidesde As Integer
Dim notihasta As Integer


    On Error GoTo CE
    Enviar = False
    
    ReDim Preserve Clave(1 To resultSet.RecordCount) As String
    Indice_rs = 0
    
    
    'El primer campo de resultados debe ser el nro de tercero
    
    colcount = resultSet.Fields.Count
    campoClave = 0
    
    If multipleAttach Then
        'busco la ruta de accesible de attach
        StrSql = " SELECT sis_dirmail FROM sistema WHERE sisnro = 1 "
        OpenRecordset StrSql, objRs
        If Not objRs.EOF Then
            Flog.writeline "Ruta accesible desde el mail encontrada: " & objRs!sis_dirmail
            rutaAttach = objRs!sis_dirmail
            If Right(rutaAttach, 1) <> "\" Then
                rutaAttach = rutaAttach & "\"
            End If
        Else
            Flog.writeline "Ruta accesible desde el mail NO encontrada: " & objRs!sis_dirmail
        End If
    End If
    
    Do Until resultSet.EOF
        docAttachs = ""
        '-----------------------------------------------------------------------------------------------------------
        'FGZ 09/08/2012 --------------------------------------------------------------------------------------------
        'Creo el proceso de mensajeria
        MyBeginTrans
        StrSql = "insert into batch_proceso "
        StrSql = StrSql & "(btprcnro, bprcfecha, iduser, bprchora, bprcfecdesde, bprcfechasta, bprcparam, "
        StrSql = StrSql & "bprcestado, bprcprogreso, bprcfecfin, bprchorafin, bprctiempo, empnro, bprcempleados) "
        StrSql = StrSql & "values (25," & ConvFecha(Date) & ",'" & Usuario & "','" & FormatDateTime(Time, 4) & ":00'"
        StrSql = StrSql & ",null,null,'1','Preparando',null,null,null,null,0,null)"
        objConn.Execute StrSql, , adExecuteNoRecords
        BProNro = getLastIdentity(objConn, "batch_proceso")
        
        
        Indice_rs = Indice_rs + 1
        Resultado = ""
        'Escribo el titulo de las columnas
        Mensaje = "<TABLE class=""tabladetalle""><TR>" & vbCrLf
        
        'LED - 15/12/2014
        If TnotiNro = 18 Then
            Call obtenerConfigNotificacion(notinro, corte, agrupa, notidesde, notihasta)
            For I = notidesde To (colcount - 1)
                Mensaje = Mensaje & "<TH>" & resultSet.Fields(I).Name & "</TH>" & vbCrLf
            Next
        Else
            If multipleAttach Then
                If campoClave <> resultSet.Fields(1).Value Then
                    For I = 3 To (colcount - 1)
                        Mensaje = Mensaje & "<TH>" & resultSet.Fields(I).Name & "</TH>" & vbCrLf
                    Next
                End If
            Else
                For I = 1 To (colcount - 1)
                    Mensaje = Mensaje & "<TH>" & resultSet.Fields(I).Name & "</TH>" & vbCrLf
                Next
            End If
        End If
        'LED - 15/12/2014
        
'        For I = 1 To (colcount - 1)
'            Mensaje = Mensaje & "<TH>" & resultSet.Fields(I).Name & "</TH>" & vbCrLf
'        Next
        
        Mensaje = Mensaje & "</TR>" & vbCrLf
        'Escribo el contenido de la fila
        If TnotiNro = 18 Then
            corteAnt = 0
            salir = False
            campoClave = 0
            Do While Not resultSet.EOF And Not salir
                campoClave = resultSet.Fields(0)
            
                'If corteAnt <> resultSet.Fields(corte) Then
                    Mensaje = Mensaje & "<TR>" & vbCrLf
                    corteAnt = resultSet.Fields(corte)
                'End If
                
                'If campoClave <> resultSet.Fields(agrupa).Value Then
                    'Mensaje = Mensaje & "<TD>" & resultSet.Fields(agrupa).Value & "</TD>" & vbCrLf
                
                    'For I = notidesde To (notihasta)
                    For I = notidesde To (colcount - 1)
                        Mensaje = Mensaje & "<TD>" & resultSet.Fields(I).Value & "</TD>" & vbCrLf
                        Clave(Indice_rs) = Clave(Indice_rs) & resultSet.Fields(I).Value & " "
                        If EsNulo(Resultado) Then
                            Resultado = Clave(Indice_rs)
                        Else
                            Resultado = Resultado & septptepar & resultSet.Fields(I).Value
                        End If
                    
                    Next
                'End If
                
                resultSet.MoveNext
                If Not resultSet.EOF Then
                    Mensaje = Mensaje & "</TR>" & vbCrLf
                    If campoClave <> resultSet.Fields(0) Or corteAnt <> resultSet.Fields(corte) Then
                        salir = True
                        resultSet.MovePrevious
                    End If
                Else
                    resultSet.MovePrevious
                    salir = True
                End If
            Loop
            
        Else
            Mensaje = Mensaje & "<TR>" & vbCrLf
            If multipleAttach Then
                
                salir = False
                evenroAttach = 0
                Do While Not resultSet.EOF And Not salir
                    campoClave = resultSet.Fields(0)
                    If Not EsNulo(resultSet.Fields(2).Value) Then
                        docAttachs = docAttachs & rutaAttach & resultSet.Fields(2).Value & ";"
                    End If
                    
                    evenroAttach = resultSet.Fields(1).Value
                    resultSet.MoveNext
                    If Not resultSet.EOF Then
                        If campoClave <> resultSet.Fields(0).Value Or evenroAttach <> resultSet.Fields(1).Value Then
                            salir = True
                        End If
                    Else
                        salir = True
                    End If
                    
                Loop
                resultSet.MovePrevious
                evenroAttach = 0
                If evenroAttach <> resultSet.Fields(1).Value Then
                    For I = 3 To (colcount - 1)
                        Mensaje = Mensaje & "<TD>" & resultSet.Fields(I).Value & "</TD>" & vbCrLf
                        'Clave = Clave & resultSet.Fields(I).Value & " "
                        Clave(Indice_rs) = Clave(Indice_rs) & resultSet.Fields(I).Value & " "
                        If EsNulo(Resultado) Then
                            Resultado = Clave(Indice_rs)
                        Else
                            Resultado = Resultado & septptepar & resultSet.Fields(I).Value
                        End If
                    Next
                    evenroAttach = resultSet.Fields(1).Value
                End If
            Else
                For I = 1 To (colcount - 1)
                    Mensaje = Mensaje & "<TD>" & resultSet.Fields(I).Value & "</TD>" & vbCrLf
                    'Clave = Clave & resultSet.Fields(I).Value & " "
                    Clave(Indice_rs) = Clave(Indice_rs) & resultSet.Fields(I).Value & " "
                    If EsNulo(Resultado) Then
                        Resultado = Clave(Indice_rs)
                    Else
                        Resultado = Resultado & septptepar & resultSet.Fields(I).Value
                    End If
                Next
            End If
        End If
        Mensaje = Mensaje & "</TR></TABLE>" & vbCrLf

        'Se asume que la primer columna es el nro de legajo
        Col1 = 0
        If IsNumeric(resultSet.Fields(0).Value) Then
            StrSql = "SELECT empleado.ternro FROM empleado"
            StrSql = StrSql & " WHERE empleg = " & resultSet.Fields(0).Value
            OpenRecordset StrSql, objRs
            If Not objRs.EOF Then
                Col1 = objRs!Ternro
            End If
        End If

        'FGZ - 09/08/2012 -------------------------------------------------------
        'Procedimiento general que devuelve una lista de mails a notificar
        Call ListaNotificacion(aleNro, notinro, TnotiNro, resultSet, Mails)
        ArrMails = Split(Mails, ";")
        For I = 0 To UBound(ArrMails)
            If Not EsNulo(ArrMails(I)) Then
                    If AleNotiUnicaVez Then
                        'reviso que no haya sido notificada ya
                        StrSql = "SELECT * FROM ale_enviada "
                        StrSql = StrSql & " WHERE alenro = " & aleNro
                        StrSql = StrSql & " AND notinro = " & notinro
                        StrSql = StrSql & " AND mail = '" & ArrMails(I) & "'"
                        StrSql = StrSql & " AND clave = '" & Clave(Indice_rs) & "'"
                        OpenRecordset StrSql, rs_Env
                        If rs_Env.EOF Then
                            StrSql = "INSERT INTO ale_enviada "
                            StrSql = StrSql & "(alenro, notinro, fecha, mail, envios, clave) VALUES "
                            StrSql = StrSql & "("
                            StrSql = StrSql & aleNro & ","
                            StrSql = StrSql & notinro & ","
                            StrSql = StrSql & ConvFecha(Date) & ","
                            'FGZ - 03/06/2014 ------------------------
                            StrSql = StrSql & "'" & Left(ArrMails(I), 100) & "',"
                            StrSql = StrSql & 0 & ","
                            StrSql = StrSql & "'" & Left(Clave(Indice_rs), 1000) & "'"
                            StrSql = StrSql & ")"
                            objConn.Execute StrSql, , adExecuteNoRecords
                            
                            Enviar = Enviar Or True
                        Else
                            Flog.writeline "Alerta: ya ha sido notificada anteriormente el " & rs_Env!Fecha & ". No se reenviará."
                            Enviar = Enviar Or False
                        End If
                    Else
                        Enviar = True
                    End If
            
            
            End If
        Next I
        'FGZ - 09/08/2012 -------------------------------------------------------
        
        If Mails <> "" And Enviar Then
            Contador = Contador + 1
            
            AlertaFileName = dirsalidas & "\msg_" & BProNro & "_ale_" & Replace(FormatDateTime(Date, 2), "/", "-") & "_" & Replace(FormatDateTime(Time, 4), ":", "-") & "-" & String(2 - Len(Second(Now)), "0") & Second(Now) & Contador
            
            'creo el proceso que se encarga de mandar el mail
            If (rsenviar = -1) Then
                'Creo el archivo a atachar en el mail
                Set fs2 = CreateObject("Scripting.FileSystemObject")
                Set AlertaFile = fs2.CreateTextFile(AlertaFileName & ".html", True)
                Flog.writeline "Fin Alerta: Positiva. Enviado a: " & Mails
                
                'AlertaFile.writeline "<html><head>"
                'AlertaFile.writeline "<STYLE> TABLE{ border : thick solid 1; width : 100%;}TH{ background-color: #333399; COLOR: #ffffff; FONT-FAMILY: 'Arial'; FONT-SIZE: 9pt; FONT-WEIGHT: bold; padding : 2 2 2 5; width : auto;}"
                'AlertaFile.writeline "TR{ COLOR: black; FONT-FAMILY: Verdana; FONT-SIZE: 08pt; BACKGROUND-COLOR: #E4FEF9; padding : 2; padding-left : 5;}h4{font-family : Verdana, Geneva, Arial, Helvetica, sans-serif;font-size : smaller;font-style : normal;color : Maroon;}</STYLE>"
                'AlertaFile.writeline "<title> Alertas - RHPro &reg; </title></head><body>"
                'AlertaFile.writeline "<h4>" & TituloAlerta & "</h4>"
                'AlertaFile.writeline "<table>" & Mensaje & "</table>"
                'AlertaFile.writeline "</body></html>"
                
                AlertaFile.writeline "<html><head>"
                AlertaFile.writeline "<STYLE>" & Estilo & "</STYLE>"
                AlertaFile.writeline "<title> Alertas - RHPro &reg; </title></head><body>"
                AlertaFile.writeline "<h4>" & TituloAlerta & "</h4>"
                AlertaFile.writeline Mensaje
                AlertaFile.writeline "</body></html>"
                AlertaFile.Close
                
                'Call crearProcesosMensajeria(aleNro, Mails, AlertaFileName, AlertaFileName)
                Call crearProcesosMensajeriaNuevo(aleNro, Mails, AlertaFileName, AlertaFileName, Resultado, Col1, docAttachs)
                'Clave(Indice_rs)
            Else
                'Call crearProcesosMensajeria(aleNro, Mails, AlertaFileName, "")
                Call crearProcesosMensajeriaNuevo(aleNro, Mails, AlertaFileName, "", Resultado, Col1, docAttachs)
            End If
        Else
            Flog.writeline "Alerta: No se encontraron mails definidos en el resultado de la consulta. No se han enviado mensajes."
        End If
        
        '---------------------------------------------------------------------------
        'Actualizo el proceso de mensajeria-----------------------------------------
        'FGZ - 05/09/2006
        StrSql = "UPDATE batch_proceso SET bprcestado = 'Pendiente'"
        StrSql = StrSql & " WHERE bpronro = " & BProNro
        objConn.Execute StrSql, , adExecuteNoRecords
        MyCommitTrans
        '---------------------------------------------------------------------------
        
        resultSet.MoveNext
    Loop
    Exit Sub
CE:
    HuboErrores = True
    Flog.writeline " Error: " & Err.Description & " > " & Now
    MyRollbackTrans
End Sub


Public Sub ListaNotificacion(ByVal aleNro As Long, ByVal notinro As Long, ByVal TipoNoti As Long, ByVal resultSet As ADODB.Recordset, ByRef Mails As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: sub que carga una lista de mails segun el tipo de notificacion
' Autor      : FGZ
' Fecha      : 09/08/2012
' Ultima Mod :
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset

Dim Estructuras As String
Dim strSQLQuery As String
Dim Clave
Dim Param As String
Dim Opcion As Integer
Dim campo As Integer
Dim reemplazo As String

On Error GoTo Me_Listamails

Mails = ""

Select Case TipoNoti
Case 1, 10: 'Envia mails a los empleados definidos en la tabla noti_empleado
        'Busco todos los empleados a los cuales les tengo que enviar los mails
        StrSql = "SELECT empemail email FROM tercero "
        StrSql = StrSql & "inner join noti_empleado on tercero.ternro = noti_empleado.ternro "
        StrSql = StrSql & "inner join empleado on empleado.ternro = noti_empleado.ternro "
        StrSql = StrSql & "where notinro = " & notinro
        OpenRecordset StrSql, rs
        Do Until rs.EOF
            If Not EsNulo(rs!email) Then
                Mails = Mails & rs!email & ";"
            End If
            rs.MoveNext
        Loop
Case 2  'Envia mails a los empleados de las estructuras definidas en noti_estructura
    ' recupero una lista de estructuras a las cuales mandarle el mail
    StrSql = "SELECT tenro, estrnro, estrfija FROM noti_estructura WHERE notinro = " & notinro
    OpenRecordset StrSql, rs2
    Estructuras = "0"
    If Not rs2.EOF Then
        If CBool(rs2!estrfija) Then
            Do Until rs2.EOF
                Estructuras = Estructuras & "," & rs2!estrnro
                rs2.MoveNext
            Loop
        End If
        If (Estructuras <> "0") Then
            'si es fija es un mismo mail para todos
            StrSql = " SELECT DISTINCT empemail email FROM empleado "
            StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.ternro = empleado.ternro "
            StrSql = StrSql & " WHERE his_estructura.htethasta IS NULL "
            StrSql = StrSql & " AND his_estructura.estrnro IN (" & Estructuras & ")"
            OpenRecordset StrSql, rs
            Do Until rs.EOF
                If Not EsNulo(rs!email) Then
                    Mails = Mails & rs!email & ";"
                End If
                rs.MoveNext
            Loop
        Else
            StrSql = "SELECT DISTINCT empemail email FROM empleado "
            StrSql = StrSql & "INNER JOIN his_estructura ON his_estructura.ternro = empleado.ternro "
            StrSql = StrSql & "WHERE his_estructura.estrnro = ("
            StrSql = StrSql & "    SELECT his_estructura.estrnro FROM his_estructura"
            StrSql = StrSql & "    INNER JOIN noti_estructura ON his_estructura.tenro = noti_estructura.tenro"
            StrSql = StrSql & "    WHERE his_estructura.htethasta IS NULL"
            StrSql = StrSql & "    AND his_estructura.Ternro = " & resultSet.Fields(0).Value
            StrSql = StrSql & "    AND noti_estructura.notinro = " & notinro & "  )"
            OpenRecordset StrSql, rs
            Do Until rs.EOF
                If Not EsNulo(rs!email) Then
                    Mails = Mails & rs!email & ";"
                End If
                rs.MoveNext
            Loop
        End If
    End If
Case 3  'Envia mails a los supervisores de los empleados
    Clave = resultSet.Fields(0).Value
    
    StrSql = "SELECT jefe.empemail email FROM empleado "
    StrSql = StrSql & "INNER JOIN empleado jefe ON empleado.empreporta = jefe.ternro "
    StrSql = StrSql & "WHERE Empleado.empleg = " & Clave
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        Mails = rs!email
    End If
Case 4  'Envia mails a los usuarios definidos en la tabla noti_usuario
    StrSql = "SELECT usremail email FROM user_per "
    StrSql = StrSql & "inner join noti_usuario on user_per.iduser = noti_usuario.iduser "
    StrSql = StrSql & "where notinro = " & notinro
    OpenRecordset StrSql, rs
    Do Until rs.EOF
        If Not EsNulo(rs!email) Then
            Mails = Mails & rs!email & ";"
        End If
        rs.MoveNext
    Loop
Case 5  'Envia mails a los roles definidos en la tabla noti_roles
    Flog.writeline "Tipo de Notificacion no se puede aplicar por cada resultado"
Case 6  'Envia mails a los empleados que cumplan con el query de la notificacion
    'recupero la query correspondiente a la notificacion
    StrSql = "SELECT query FROM noti_query WHERE notinro = " & notinro
    OpenRecordset StrSql, rs2
    If Not rs2.EOF Then
        'reemplazo los parametros de la query por valores verdaderos
        strSQLQuery = rs2!query
        Do Until InStr(1, strSQLQuery, "[") = 0
            Param = Mid(strSQLQuery, InStr(1, strSQLQuery, "[") + 1, Len(strSQLQuery))
            Param = Mid(Param, 1, InStr(1, Param, "]") - 1)
            'si tiene forma [campo,opcion] o tiene forma [campo]
            If InStr(1, Param, ",") <> 0 Then
                campo = Mid(Param, 1, InStr(1, Param, ",") - 1)
                Opcion = Mid(Param, InStr(1, Param, ",") + 1, Len(Param) - InStr(1, Param, ","))
            Else
                campo = Param
            End If
            
            Select Case campo
                Case 1  ' Fecha Actual
                    reemplazo = Date
                Case 2  ' Columna X, donde X es opcion
                    reemplazo = resultSet.Fields(Opcion).Value
            End Select
            strSQLQuery = Replace(strSQLQuery, "[" & Param & "]", reemplazo)
        Loop ' hasta aca reemplazo la sql
        OpenRecordset strSQLQuery, rs
        Do Until rs.EOF
            ' la primera columna del query es una dir de email
            If Not IsNull(rs.Fields(0).Value) Then
                If Len(rs.Fields(0).Value) > 0 Then
                    Mails = Mails & rs.Fields(0).Value & ";"
                End If
            End If
            rs.MoveNext
        Loop
    End If
Case 7  'Envia mails a los Postulantes que cumplan con el query de la notificacion (se ejecuta en conexion)
    Flog.writeline "Tipo de Notificacion no se puede aplicar por cada resultado"
Case 8 'Envia mails a los terceros que cumplan con el query de la notificacion (Se debe incluir el campo ternro primero en el select)
    StrSql = " SELECT teremail "
    StrSql = StrSql & " FROM tercero "
    StrSql = StrSql & " WHERE ternro = " & resultSet.Fields("ternro").Value
    OpenRecordset StrSql, rs
    Do Until rs.EOF
        ' la primera columna del query es una dir de email
        If Not IsNull(rs.Fields(0).Value) Then
            If Len(rs.Fields(0).Value) > 0 Then
                Mails = Mails & rs.Fields(0).Value & ";"
            End If
        End If
        rs.MoveNext
    Loop

Case 9 'Envia mails a los empleados que cumplan con el query de la notificacion (Se debe incluir el campo ternro primero en el select)
    StrSql = " SELECT empemail "
    StrSql = StrSql & " FROM empleado "
    StrSql = StrSql & " WHERE ternro = " & resultSet.Fields("ternro").Value
    OpenRecordset StrSql, rs
    Do Until rs.EOF
        ' la primera columna del query es una dir de email
        If Not IsNull(rs.Fields(0).Value) Then
            If Len(rs.Fields(0).Value) > 0 Then
                Mails = Mails & rs.Fields(0).Value & ";"
            End If
        End If
        rs.MoveNext
    Loop
    
Case 18 'Envia mails a los terceros que cumplan con el query de la notificacion (Se debe incluir el campo ternro primero en el select)
    StrSql = " SELECT empemail FROM empleado " & _
             " WHERE ternro = " & resultSet.Fields("ternro").Value
    OpenRecordset StrSql, rs
    Do Until rs.EOF
        ' la primera columna del query es una dir de email
        If Not IsNull(rs.Fields(0).Value) Then
            If Len(rs.Fields(0).Value) > 0 Then
                Mails = Mails & rs.Fields(0).Value & ";"
            End If
        End If
        rs.MoveNext
    Loop
Case Else
    Flog.writeline "Tipo de Notificacion Inexistente"
End Select


If rs.State = adStateOpen Then rs.Close
If rs2.State = adStateOpen Then rs2.Close
Set rs = Nothing
Set rs2 = Nothing
Exit Sub

Me_Listamails:
    Flog.writeline "Error obteniendo lista de mails a notificar"
    Mails = ""
    
End Sub





'-----------------------------------------------------------------------------------------------
' Crear el proceso de mensajeria y le agrega el archivo a enviar
'-----------------------------------------------------------------------------------------------
Sub crearProcesosMensajeria(aleNro As Integer, ByVal mailBoxs As String, ByVal AlertaFileName As String, ByVal AlertaAttachFileName As String)

Dim objRs As New ADODB.Recordset
Dim fs2, MsgFile
Dim titulo As String
Dim DescExt As String
    
    Set fs2 = CreateObject("Scripting.FileSystemObject")
    
    'FGZ - 05/09/2006
    'Los nombres de los archivos para los mails de esta alerta empiezan con el bpronro de este proceso
    Set MsgFile = fs2.CreateTextFile(AlertaFileName & ".msg", True)
    'Set MsgFile = fs2.CreateTextFile("MSG_" & CStr(BProNro) & "_" & AlertaFileName & ".msg", True)
    
    StrSql = "select * from alertas "
    StrSql = StrSql & " WHERE alertas.alenro = " & aleNro & " "
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        Flog.writeline "No se encuentra la alerta a ejecutar."
        Exit Sub
    End If
    
    MsgFile.writeline "[MailMessage]"
    MsgFile.writeline "FromName=RHPro Servicio de Alertas"
    'FGZ - 23/11/2015 ---------------------------------------------
    'MsgFile.writeline "Subject=Alerta - " & Trim(objRs!aledes)
    MsgFile.writeline "Subject= " & Trim(objRs!aledes)
    'FGZ - 23/11/2015 ---------------------------------------------
    MsgFile.writeline "Body1=" & Trim(objRs!aledesext)
    If Len(AlertaAttachFileName) > 0 Then
       MsgFile.writeline "Attachment=" & AlertaAttachFileName & ".html"
    Else
       MsgFile.writeline "Attachment="
    End If
    
    MsgFile.writeline "Recipients=" & mailBoxs
    titulo = Trim(objRs!aledes)
    'FGZ - 03/06/2014 -----------------
    DescExt = Trim(objRs!aledesext)
    'FGZ - 03/06/2014 -----------------
    
    If objRs.State = adStateOpen Then objRs.Close
    
    StrSql = "select cfgemailfrom,cfgemailhost,cfgemailport,cfgemailuser,cfgemailpassword,cfgssl from conf_email where cfgemailest = -1"
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        MsgFile.writeline "FromAddress=" & objRs!cfgemailfrom
        MsgFile.writeline "Host=" & objRs!cfgemailhost
        MsgFile.writeline "Port=" & objRs!cfgemailport
        MsgFile.writeline "User=" & objRs!cfgemailuser
        MsgFile.writeline "Password=" & objRs!cfgemailpassword
    Else
        Flog.writeline "No existen datos configurados para el envio de emails, o no existe configuracion activa"
        Exit Sub
    End If
    MsgFile.writeline "CCO="
    MsgFile.writeline "CC="
    
    Dim ArrTags As String
    ArrTags = "{ññA001}" & septptepar & Mensaje & septptegrp
    ArrTags = ArrTags & "{ññA002}" & septptepar & titulo & septptegrp
    'FGZ - 03/06/2014 -----------------
    ArrTags = ArrTags & "{ññA003}" & septptepar & DescExt & septptegrp
    'FGZ - 03/06/2014 -----------------
    
    MsgFile.writeline "HTMLBody=" & generarTemplate(6, aleNro, 1, ArrTags, dirsalidas, BProNro, 0, AlertaFileName) 'BodyHtml
    MsgFile.writeline "HTMLMailHeader=" & generarMailHeader(6, aleNro, 1)
    MsgFile.writeline "SSL=" & objRs!cfgssl
    
'    'FGZ - 04/09/2006 - Saco esto y lo pongo afuera
'    StrSql = "insert into batch_proceso "
'    StrSql = StrSql & "(btprcnro, bprcfecha, iduser, bprchora, bprcfecdesde, bprcfechasta, bprcparam, "
'    StrSql = StrSql & "bprcestado, bprcprogreso, bprcfecfin, bprchorafin, bprctiempo, empnro, bprcempleados) "
'    StrSql = StrSql & "values (25," & ConvFecha(Date) & ",'" & Usuario & "','" & FormatDateTime(Time, 4) & ":00'"
'    StrSql = StrSql & ",null,null,'1','Pendiente',null,null,null,null,0,null)"
'
'    objConn.Execute StrSql, , adExecuteNoRecords
    
    If objRs.State = adStateOpen Then objRs.Close

End Sub


Sub crearProcesosMensajeriaNuevo(aleNro As Integer, ByVal mailBoxs As String, ByVal AlertaFileName As String, ByVal AlertaAttachFileName As String, ByVal Registro As String, ByVal Col1 As Long, ByVal docsAttachs As String)
'-----------------------------------------------------------------------------------------------
' Crear el proceso de mensajeria y le agrega el archivo a enviar
'-----------------------------------------------------------------------------------------------
Dim objRs As New ADODB.Recordset
Dim fs2, MsgFile
Dim titulo As String
Dim Clave() As String
Dim I As Long
Dim MaxCol As Long
'Dim Imagen As String
'Dim TieneImagen As Boolean

    Clave = Split(Registro, septptepar)
    MaxCol = UBound(Clave)
    
    Set fs2 = CreateObject("Scripting.FileSystemObject")
    
    'FGZ - 05/09/2006
    'Los nombres de los archivos para los mails de esta alerta empiezan con el bpronro de este proceso
    Set MsgFile = fs2.CreateTextFile(AlertaFileName & ".msg", True)
    'Set MsgFile = fs2.CreateTextFile("MSG_" & CStr(BProNro) & "_" & AlertaFileName & ".msg", True)
    
    StrSql = "select * from alertas "
    StrSql = StrSql & " WHERE alertas.alenro = " & aleNro & " "
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        Flog.writeline "No se encuentra la alerta a ejecutar."
        Exit Sub
    End If
    
    MsgFile.writeline "[MailMessage]"
    MsgFile.writeline "FromName=RHPro Servicio de Alertas"
   'FGZ - 23/11/2015 ---------------------------------------------
    'MsgFile.writeline "Subject=Alerta - " & Trim(objRs!aledes)
    MsgFile.writeline "Subject= " & Trim(objRs!aledes)
    'FGZ - 23/11/2015 ---------------------------------------------
    MsgFile.writeline "Body1=" & Trim(objRs!aledesext)
    If Len(AlertaAttachFileName) > 0 Then
        MsgFile.writeline "Attachment=" & AlertaAttachFileName & ".html"
    Else
        MsgFile.writeline "Attachment="
    End If
    
    MsgFile.writeline "Recipients=" & mailBoxs
    titulo = Trim(objRs!aledes)
    
    If objRs.State = adStateOpen Then objRs.Close
    
    StrSql = "select cfgemailfrom,cfgemailhost,cfgemailport,cfgemailuser,cfgemailpassword from conf_email where cfgemailest = -1"
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        MsgFile.writeline "FromAddress=" & objRs!cfgemailfrom
        MsgFile.writeline "Host=" & objRs!cfgemailhost
        MsgFile.writeline "Port=" & objRs!cfgemailport
        MsgFile.writeline "User=" & objRs!cfgemailuser
        MsgFile.writeline "Password=" & objRs!cfgemailpassword
    Else
        Flog.writeline "No existen datos configurados para el envio de emails, o no existe configuracion activa"
        Exit Sub
    End If
    MsgFile.writeline "CCO="
    MsgFile.writeline "CC="
    
    Dim ArrTags As String
    'ArrTags = "{ññA001}" & septptepar & Mensaje & septptegrp
    'ArrTags = ArrTags & "{ññA002}" & septptepar & titulo & septptegrp
    
    ArrTags = "{ññA002}" & septptepar & titulo & septptegrp
    
    
    'FGZ - 31/07/2012 -------------------------------
    'Agrego los tags de las columnas de los resultados de la alerta
    For I = 0 To MaxCol
        ArrTags = ArrTags & "{ññC" & Format(I, "000") & "}" & septptepar & Clave(I) & septptegrp
    Next I
    
    MsgFile.writeline "HTMLBody=" & generarTemplate(6, aleNro, 1, ArrTags, dirsalidas, BProNro, Col1, AlertaFileName) 'BodyHtml
    MsgFile.writeline "HTMLMailHeader=" & generarMailHeader(6, aleNro, 1) & docsAttachs
    
'    'FGZ - 04/09/2006 - Saco esto y lo pongo afuera
'    StrSql = "insert into batch_proceso "
'    StrSql = StrSql & "(btprcnro, bprcfecha, iduser, bprchora, bprcfecdesde, bprcfechasta, bprcparam, "
'    StrSql = StrSql & "bprcestado, bprcprogreso, bprcfecfin, bprchorafin, bprctiempo, empnro, bprcempleados) "
'    StrSql = StrSql & "values (25," & ConvFecha(Date) & ",'" & Usuario & "','" & FormatDateTime(Time, 4) & ":00'"
'    StrSql = StrSql & ",null,null,'1','Pendiente',null,null,null,null,0,null)"
'
'    objConn.Execute StrSql, , adExecuteNoRecords
    
    If objRs.State = adStateOpen Then objRs.Close

End Sub


Sub crearProcesosMensajeriaTLP(aleNro As Integer, ByVal mailBoxs As String, ByVal AlertaFileName As String, ByVal AlertaAttachFileName As String, ByVal Mensaje As String)
Dim objRs As New ADODB.Recordset
Dim fs2, MsgFile

    Set fs2 = CreateObject("Scripting.FileSystemObject")
    
    'FGZ - 05/09/2006
    'Los nombres de los archivos para los mails de esta alerta empiezan con el bpronro de este proceso
    Set MsgFile = fs2.CreateTextFile(AlertaFileName & ".msg", True)
    'Set MsgFile = fs2.CreateTextFile("MSG_" & CStr(BProNro) & "_" & AlertaFileName & ".msg", True)
    
    StrSql = "select * from alertas "
    StrSql = StrSql & " WHERE alertas.alenro = " & aleNro & " "
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        Flog.writeline "No se encuentra la alerta a ejecutar."
        Exit Sub
    End If
    
    MsgFile.writeline "[MailMessage]"
    MsgFile.writeline "FromName=Teleperformance. Dpto Recruiting."
    MsgFile.writeline "Subject=Alerta - Busqueda"
    
    'MsgFile.writeline "Body1=" & Trim(objRs!aledesext)
    MsgFile.writeline "Body1=" & Mensaje
    
    If Len(AlertaAttachFileName) > 0 Then
       MsgFile.writeline "Attachment=" & AlertaAttachFileName & ".html"
    Else
       MsgFile.writeline "Attachment="
    End If
    
    MsgFile.writeline "Recipients=" & mailBoxs
    
    If objRs.State = adStateOpen Then objRs.Close
    
    StrSql = "select cfgemailfrom,cfgemailhost,cfgemailport,cfgemailuser,cfgemailpassword from conf_email where cfgemailest = -1"
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        MsgFile.writeline "FromAddress=" & objRs!cfgemailfrom
        MsgFile.writeline "Host=" & objRs!cfgemailhost
        MsgFile.writeline "Port=" & objRs!cfgemailport
        MsgFile.writeline "User=" & objRs!cfgemailuser
        MsgFile.writeline "Password=" & objRs!cfgemailpassword
    Else
        Flog.writeline "No existen datos configurados para el envio de emails, o no existe configuracion activa"
        Exit Sub
    End If
    If objRs.State = adStateOpen Then objRs.Close
End Sub

Public Sub BuscarImg(ByVal Tipo As Long, ByVal Ternro As Long, ByRef Img As String, ByRef ImgFull As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Funcion que busca la imagen del tipo y tercero pasado por parametros.
' Autor      : FGZ
' Fecha      : 03/08/2012
' Ultima Mod :
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs As New ADODB.Recordset
        
        StrSql = "SELECT tipimdire, tipimanchodef, tipimaltodef, terimnombre, ter_imag.terimfecha "
        StrSql = StrSql & " FROM ter_imag "
        StrSql = StrSql & " LEFT JOIN tipoimag ON tipoimag.tipimnro = ter_imag.tipimnro "
        StrSql = StrSql & " WHERE ter_imag.ternro = " & Ternro
        StrSql = StrSql & " AND ter_imag.tipimnro = " & Tipo
        StrSql = StrSql & " ORDER BY ter_imag.terimfecha DESC "
        OpenRecordset StrSql, rs
        If Not rs.EOF Then
            If Not EsNulo(rs("terimnombre")) Then
                ImgFull = Trim(rs("tipimdire")) & Trim(rs("terimnombre"))
                Img = Trim(rs("terimnombre"))
               'resultado = <img src="<%=l_tipimdire%><%=l_terimnombre%>" height="<%=l_tipimaltodef%>" alt="" border="0" style="border:2px solid #333;overflow:hidden">"
            Else
                ImgFull = "nofoto.png"
                Img = "nofoto.png"
                'resultado = <img src="/rhprox2/shared/fotos/nofoto.png" alt="" border="0" style="margin-right:2px;margin-top:2px;">
            End If
        Else
            ImgFull = "nofoto.png"
            Img = "nofoto.png"
        End If
End Sub


Public Sub OpenRecordsetWithConn(strSQLQuery As String, ByRef objRs As ADODB.Recordset, ByRef Conn As ADODB.Connection, Optional lockType As LockTypeEnum = adLockReadOnly)
Dim pos1 As Integer
Dim pos2 As Integer
Dim aux As String

    'Abre un recordset con la consulta strSQLQuery
    If objRs.State <> adStateClosed Then
        If objRs.lockType <> adLockReadOnly Then objRs.UpdateBatch
        objRs.Close
    End If
    
    objRs.CacheSize = 500

    objRs.Open strSQLQuery, Conn, adOpenDynamic, lockType, adCmdText
    
End Sub


Public Sub BuscarMaxTipoImg()
' ---------------------------------------------------------------------------------------------
' Descripcion: Funcion que busca el maximo tipo de imagen existente en la bd
' Autor      : FGZ
' Fecha      : 03/08/2012
' Ultima Mod :
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs As New ADODB.Recordset

    MaxTipoImg = 0
        
    StrSql = "SELECT Max(tipimnro) maximo "
    StrSql = StrSql & " FROM tipoimag "
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        MaxTipoImg = rs("maximo")
    End If
    
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing

End Sub


Sub FechaHora()
'MDF: cambie los if que manipulaban las horas usando arreglos
    Dim l_day
    Dim arr_fech1
    Dim arr_fech2
    Dim l_frectipnro
    Dim l_schedhora
    Dim l_alesch_fecini
    Dim l_alesch_fecfin
    Dim l_alesch_frecrep
    l_sql = "SELECT frectipnro, alesch_fecini, schedhora, alesch_frecrep, alesch_fecfin "
    l_sql = l_sql & "FROM ale_sched "
    l_sql = l_sql & "WHERE schednro = " & l_schednro
    OpenRecordset l_sql, objRs2
    l_frectipnro = objRs2!frectipnro
    l_schedhora = objRs2!schedhora
    l_alesch_fecini = objRs2!alesch_fecini
    l_alesch_fecfin = objRs2!alesch_fecfin
    l_alesch_frecrep = objRs2!alesch_frecrep
    If (DateValue(Date) >= DateValue(l_alesch_fecini)) And (DateValue(Date) <= DateValue(l_alesch_fecfin)) Then
        ' Diariamiente
        If l_frectipnro = 1 Then
           ' If Int(Left(FormatDateTime(Time, 4), 2) & Mid(Time, 4, 2)) + 1 > Int(Left(FormatDateTime(l_schedhora, 4), 2) & Mid(l_schedhora, 4, 2)) Then
            arr_fech1 = Split(CStr(FormatDateTime(Time, 4)), ":") 'mdf
            If Len(arr_fech1(0)) = 1 Then
              arr_fech1(0) = "0" & arr_fech1(0)
            End If
            arr_fech2 = Split(CStr(l_schedhora), ":") 'mdf
            If Len(arr_fech2(0)) = 1 Then
              arr_fech2(0) = "0" & arr_fech2(0)
            End If
            If Int(arr_fech1(0) & arr_fech1(1)) + 1 > Int(arr_fech2(0) & arr_fech2(1)) Then
                'l_dia = DateValue(Date + 1)
                l_dia = DateAdd("d", 1, Date)
                l_hora = l_schedhora & ":00"
            Else
                l_dia = Date 'mdf
                l_hora = l_schedhora & ":00"
            End If
        Else
        ' Mensualmente
            If l_frectipnro = 3 Then
                  If Int(Day(Date)) < Int(l_alesch_frecrep) Then
                     l_dia = DateValue(l_alesch_frecrep & "/" & Month(Date) & "/" & Year(Date))
                     'l_dia = l_alesch_frecrep & "/" & Month(Date) & "/" & Year(Date)
                     l_hora = l_schedhora & ":00"
                  Else
                     If Int(Day(Date)) = Int(l_alesch_frecrep) Then
                    'If Int(Left(FormatDateTime(Time, 4), 2) & Mid(Time, 4, 2)) + 1 > Int(Left(FormatDateTime(l_schedhora, 4), 2) & Mid(l_schedhora, 4, 2)) Then
                        arr_fech1 = Split(CStr(FormatDateTime(Time, 4)), ":") 'mdf
                        If Len(arr_fech1(0)) = 1 Then
                          arr_fech1(0) = "0" & arr_fech1(0)
                        End If
                        arr_fech2 = Split(CStr(l_schedhora), ":") 'mdf
                        If Len(arr_fech2(0)) = 1 Then
                           arr_fech2(0) = "0" & arr_fech2(0)
                        End If
                        If Int(arr_fech1(0) & arr_fech1(1)) + 1 > Int(arr_fech2(0) & arr_fech2(1)) Then
                        'l_dia = DateValue(l_alesch_frecrep & "/" & Int(Month(Date)) + 1 & "/" & Year(Date))
                        'l_hora = l_schedhora & ":00"
                          l_dia = DateAdd("m", 1, Date) 'un mes al dia de la ejecucion
                          l_hora = l_schedhora & ":00"
                         Else
                           l_dia = DateValue(l_alesch_frecrep & "/" & Month(Date) & "/" & Year(Date))
                           l_hora = l_schedhora & ":00"
                         End If
                      End If
                 End If
            Else
            ' Semanalmente
               If l_frectipnro = 2 Then
                    If Weekday(Date) < Int(l_alesch_frecrep) Then
                        l_dia = DateValue(Date + l_alesch_frecrep - Weekday(Date))
                        l_hora = l_schedhora & ":00"
                    Else
                        If Weekday(Date) > Int(l_alesch_frecrep) Then
                            l_dia = DateValue(Date + 7 - (Weekday(Date) - l_alesch_frecrep))
                            l_hora = l_schedhora & ":00"
                        Else
                            'If Int(Left(FormatDateTime(Time, 4), 2) & Mid(Time, 4, 2)) + 1 > Int(Left(FormatDateTime(l_schedhora, 4), 2) & Mid(l_schedhora, 4, 2)) Then
                            arr_fech1 = Split(CStr(FormatDateTime(Time, 4)), ":") 'mdf
                            If Len(arr_fech1(0)) = 1 Then
                              arr_fech1(0) = "0" & arr_fech1(0)
                            End If
                            arr_fech2 = Split(CStr(l_schedhora), ":") 'mdf
                            If Len(arr_fech2(0)) = 1 Then
                              arr_fech2(0) = "0" & arr_fech2(0)
                            End If
                            If Int(arr_fech1(0) & arr_fech1(1)) + 1 > Int(arr_fech2(0) & arr_fech2(1)) Then
                               ' l_dia = DateValue(Date + 7 - (Weekday(Date) - l_alesch_frecrep))
                                l_dia = DateAdd("d", 7, Date)
                                l_hora = l_schedhora
                            Else
                                l_dia = Date
                                l_hora = l_schedhora & ":00"
                            End If
                         End If
                     End If
            Else
                ' Temporal
                If l_frectipnro = 4 Then
                    l_dia = DateValue(Now + l_alesch_frecrep)
                    l_hora = FormatDateTime(Now + l_alesch_frecrep, 4) & ":00"
                End If
            End If
        End If
       End If
        ' Fecha siguiente fuera del tope maximo
        If DateValue(l_dia) > DateValue(l_alesch_fecfin) Then
            Flog.writeline "ATENCION:No se puede activar la alerta, ya que la fecha de la proxima ejecucion esta fuera del rango de vigencia del schedule asociado. Verifique el mismo."
            HuboErrores = True
            Exit Sub
        End If
    Else
    ' Fecha fuera de rango
    End If
    objRs2.Close
End Sub


Function ReplaceFields(ByVal query As String, ByVal aleconsnro As String, ByVal aleNro As String) As String
    Dim sql As String
    Dim campo As String
    
    Do Until InStr(1, query, "@@") = 0
        campo = Mid(query, InStr(1, query, "@@") + 2, Len(query))
        campo = Mid(campo, 1, InStr(1, campo, "@@") - 1)
        sql = "select alepa_valor, alepa_tipo from ale_param where alenro = " & aleNro
        sql = sql & " and aleconsnro = " & aleconsnro
        sql = sql & " and upper(alepa_nombre) = '" & UCase(campo) & "'"
        OpenRecordset sql, objRs2
        If Not objRs2.EOF Then
            If UCase(objRs2!alepa_tipo) = "D" Or UCase(objRs2!alepa_tipo) = "S" Then
              query = Replace(query, "@@" & campo & "@@", "'" & objRs2!alepa_valor & "'")
            Else
              query = Replace(query, "@@" & campo & "@@", objRs2!alepa_valor)
            End If
        Else
            Flog.writeline "Error en campo parametro '" & campo & "'"
            'FGZ - 27/09/2013 --------------------------------------------------------
            query = Replace(query, "@@" & campo & "@@", 0)
            'Exit Function
            'FGZ - 27/09/2013 --------------------------------------------------------
        End If
        If objRs2.State = adStateOpen Then objRs2.Close
    Loop
    ReplaceFields = query
End Function


Sub EnviarVariosMails(ByVal aleNro As Integer, ByVal Tipo As Integer, ByRef resultSet As ADODB.Recordset, ByVal multipleAttach As Boolean)
'-----------------------------------------------------------------------------------------------
' Se encarga de enviar 1 mail por cada resultado de la alerta
'-----------------------------------------------------------------------------------------------
' Modificado :  09/08/2012 - FGZ - se agregó un parametro a las alertas para ver si se debe unviar un mail por cada resultado de la alerta
'-----------------------------------------------------------------------------------------------
Dim objRs As New ADODB.Recordset
'Dim MailPorResultado As Boolean

    On Error GoTo CE
    
    ' Datos de las Notificaciones asociadas a la alerta
    
    StrSql = "SELECT noti_ale.notinro, noti_ale.rsenviar, notificacion.tnotinro "
    StrSql = StrSql & "From noti_ale "
    StrSql = StrSql & "INNER JOIN notificacion ON noti_ale.notinro = notificacion.notinro "
    StrSql = StrSql & "WHERE noti_ale.alenro = " & aleNro
    OpenRecordset StrSql, objRs

    Flog.writeline "Comienzo del Envio de Mails a " & StrSql
    
    
'        MyBeginTrans
'
'        'FGZ - 05/09/2006
'        StrSql = "insert into batch_proceso "
'        StrSql = StrSql & "(btprcnro, bprcfecha, iduser, bprchora, bprcfecdesde, bprcfechasta, bprcparam, "
'        StrSql = StrSql & "bprcestado, bprcprogreso, bprcfecfin, bprchorafin, bprctiempo, empnro, bprcempleados) "
'        StrSql = StrSql & "values (25," & ConvFecha(Date) & ",'" & Usuario & "','" & FormatDateTime(Time, 4) & ":00'"
'        StrSql = StrSql & ",null,null,'1','Preparando',null,null,null,null,0,null)"
'        objConn.Execute StrSql, , adExecuteNoRecords
'
'        BProNro = getLastIdentity(objConn, "batch_proceso")


    Call BuscarMaxTipoImg
    
    Do Until objRs.EOF
        resultSet.MoveFirst
        Flog.writeline "Tipo" & objRs!TnotiNro
        
        Mensaje = "" ' 17-03-2011 - inicializar el mensaje para que no arrastre lo mensajes de otors tipos de notificacion
        
        'Envia 1 mail a quien corresponda (segun tipo notificacion) por cada resultado de la alerta
        Call MailsA_X_1Xresultado(aleNro, objRs!notinro, objRs!rsenviar, resultSet, objRs!TnotiNro, multipleAttach)
        objRs.MoveNext
    Loop
    
'    'FGZ - 05/09/2006
'    StrSql = "UPDATE batch_proceso SET bprcestado = 'Pendiente'"
'    StrSql = StrSql & " WHERE bpronro = " & BProNro
'    objConn.Execute StrSql, , adExecuteNoRecords
'    MyCommitTrans
    
    Exit Sub
CE:
    HuboErrores = True
    Flog.writeline " Error: " & Err.Description & " > " & Now
    MyRollbackTrans
End Sub



Public Sub CargarEstilo()
Const ForReading = 1
Dim f, fs
Dim strline As String
Dim Encontro As Boolean


    Estilo = ""
    On Error Resume Next
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.OpenTextFile(dirTemplates & "\estilo.css", ForReading, 0)
    If Err.Number <> 0 Then
        Estilo = "body {border: none;overflow:auto;} Table{border : 0px;width : 100%;}TH{background-color: #D13528;COLOR: #ffffff;FONT-FAMILY: Arial;FONT-SIZE: 9pt;FONT-WEIGHT: bold;padding : 2 2 2 5;width : auto;}"
        Estilo = Estilo & "TD{COLOR: black;FONT-FAMILY: Verdana;FONT-SIZE: 08pt;padding : 2;padding-left : 5;}"
        GoTo Fin
    End If
    Encontro = False
    Do While Not f.AtEndOfStream
        Estilo = Estilo & f.ReadLine()
    Loop

Fin:
    f.Close
End Sub

Public Sub ActualizarEstado(ByVal Estado As String, ByVal PID As Long)

        'Actualizo el progreso
        MyBeginTrans
            StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = '" & Estado & "', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
        MyCommitTrans

End Sub

Public Sub obtenerConfigNotificacion(ByVal notinro As Integer, ByRef corte As Integer, ByRef agrupa As Integer, ByRef notidesde As Integer, ByRef notihasta As Integer)
Dim rsAux As New ADODB.Recordset
    
    StrSql = " SELECT colcorte, colagrupa, coldesde, colhasta FROM noti_agrupa WHERE notinro = " & notinro
    OpenRecordset StrSql, rsAux
    If Not rsAux.EOF Then
        corte = rsAux!colcorte
        agrupa = rsAux!colagrupa
        notidesde = rsAux!coldesde
        notihasta = rsAux!colhasta
    End If
    
    If rsAux.State = adStateOpen Then rsAux.Close
    Set rsAux = Nothing
End Sub
