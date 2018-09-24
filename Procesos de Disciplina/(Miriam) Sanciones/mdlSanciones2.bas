Attribute VB_Name = "MdlSanciones"
Option Explicit

'Const Version = "1.00"
'Const FechaVersion = "29/09/2011"
''Modificaciones: Margiotta , Emanuel


'Const Version = 1.01
'Const FechaVersion = "09/11/2011"
'Modificaciones: Margiotta, Emanuel
    'Se corrigio la consulta de la funcion CantEmpXCausal que busca el conjuto de empleado

Const Version = 1.02
Const FechaVersion = "18/07/2014"
'Modificaciones: Ruiz Miriam - CAS-11070 - Monresa - Sanciones
    'Se corrigio cuando procesaba mas de un empleado


'---------------------------------------------------------------------------
'       Version no liberada
'---------------------------------------------------------------------------
'Const Version = 1.xx
'Const FechaVersion = "24/02/2015"
'Modificaciones: LED - Ruiz Miriam - Monresa - Sanciones
    ' si el criterio tiene mas de una condición descuenta solo cuando las dos se cumplen
    ' se corrigió la cantidad acumulada

'---------------------------------------------------------------------------
'---------------------------------------------------------------------------
Dim CEmpleadosAProc As Integer
Dim CDiasAProc As Integer
Dim IncPorc As Single
Dim Progreso As Single
Dim TiempoAcumulado As Single
Dim IncPorcEmpleado As Single
Dim HuboErrores As Boolean
Dim ProgresoEmpleado As Single
Dim TiempoInicialProceso As Single





Public Sub Main()

Dim FechaDesde As Date
Dim FechaHasta As Date
Dim FechaProc As Date
Dim strcmdLine As String
Dim fs
Dim objRs As New ADODB.Recordset
Dim myrs As New ADODB.Recordset

Dim rsCausal As New ADODB.Recordset
Dim rsEmpCausal As New ADODB.Recordset
Dim rsCriterio As New ADODB.Recordset
Dim rsCondicion As New ADODB.Recordset
Dim rsCantEmpleados As New ADODB.Recordset
Dim rsAux As New ADODB.Recordset
Dim cantEmpleados As Long
Dim cantEmpleadosAux As Long
Dim cantCausales As Long
Dim procTodosEmp As Integer
ReDim arrSqlEmpCausales(0) As String
Dim sqlAux As String
Dim i As Long
'Dim reprocesar As Integer
Dim arrProcParam
Dim Hora As String
Dim condtipocantagrup As String
Dim aplicoReduccion As Boolean
'MR - 17/07/2014
Dim CantidadEmpleados As Integer

Dim rs_Batch_Proceso As New ADODB.Recordset
Dim rs_His_Batch_Proceso As New ADODB.Recordset
Dim rs_Per As New ADODB.Recordset
Dim PID As String
Dim arrParametros


    strcmdLine = Command()
    arrParametros = Split(strcmdLine, " ", -1)
    If UBound(arrParametros) > 1 Then
        If IsNumeric(arrParametros(0)) Then
            NroProceso = arrParametros(0)
            Etiqueta = arrParametros(1)
            EncriptStrconexion = CBool(arrParametros(2))
            c_seed = arrParametros(2)
        Else
            Exit Sub
        End If
    Else
        If UBound(arrParametros) > 0 Then
            If IsNumeric(arrParametros(0)) Then
                NroProceso = arrParametros(0)
                Etiqueta = arrParametros(1)
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

    HuboErrores = False
    
    
    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(PathFLog & "Sanciones" & "-" & NroProceso & ".log", True)
    
    Cantidad_de_OpenRecordset = 0
    Cantidad_Call_Politicas = 0
    
    'Abro la conexion
    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If
    
    Nivel_Tab_Log = 0
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-------------------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion(No liberada): " & FechaVersion
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-------------------------------------------------------------"
    Flog.writeline
    
    '--------- Control de versiones ------
    Version_Valida = ValidarV(Version, 2, TipoBD)
    If Not Version_Valida Then
        'Actualizo el progreso
        MyBeginTrans
            StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Error de Version', bprcpid = " & PID & " WHERE bpronro = " & NroProceso
            objConn.Execute StrSql, , adExecuteNoRecords
        MyCommitTrans
        
        Flog.writeline
        GoTo Final
    End If
    
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprcpid = " & PID & ", bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & _
             ", bprcestado = 'Procesando', bprcprogreso = 1 , bprctiempo = 1 WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords


    
    'EAM- Obtiene los datos de Bach_Proceso
    StrSql = "SELECT IdUser,bprcfecha,bprcfecdesde,bprcfechasta, bprcparam, bprcempleados FROM batch_proceso WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        Flog.writeline Espacios(Tabulador * 1) & "Parametros: "
        Flog.writeline Espacios(Tabulador * 1) & "Usuario: " & objRs!IdUser
        Flog.writeline Espacios(Tabulador * 1) & "Fecha: " & objRs!bprcfecha
        Flog.writeline Espacios(Tabulador * 1) & "Desde: " & objRs!bprcfecdesde
        Flog.writeline Espacios(Tabulador * 1) & "bprcparam: " & objRs!bprcparam
        
        arrProcParam = Split(objRs!bprcparam, "@")
        'EAM- Verifica si en el proceso se selecciono todos los empleados o no
        If EsNulo(arrProcParam(0)) Then
            procTodosEmp = 0
        Else
            If (arrProcParam(0) = -1) Then
                procTodosEmp = -1
            Else
                procTodosEmp = 0
            End If
        End If
        
        'reprocesar = arrProcParam(1)
        FechaProc = objRs!bprcfecdesde
        CantidadEmpleados = objRs!bprcempleados
        
        
    Else
        Exit Sub
    End If
    OpenConnection strconexion, CnTraza
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If
    
    Flog.writeline "-------------------------------------------------------------"
    Flog.writeline "-------------------------------------------------------------"
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "Inicio :" & Now
    
    TiempoInicialProceso = GetTickCount
    'EAM- Si no se selecciono todos los empleados en el proceso, obtener el conjuto a procesar
    If (procTodosEmp = 0) Then
        'Obtiene los empleados del proceso
        StrSql = "SELECT empleado.ternro, empleado.empleg FROM batch_empleado " & _
                " INNER JOIN empleado ON batch_empleado.ternro = empleado.Ternro " & _
                " WHERE batch_empleado.bpronro = " & NroProceso
        If myrs.State = adStateOpen Then myrs.Close
        OpenRecordset StrSql, myrs
    Else
       StrSql = "SELECT empleado.ternro, empleado.empleg FROM empleado " & _
                "WHERE empest= -1"
      
    End If
      If myrs.State = adStateOpen Then myrs.Close
        OpenRecordset StrSql, myrs
    
    'EAM- Obtiene los causales que se van a analizar
    StrSql = "SELECT * FROM dis_causal WHERE causestado=-1"
    OpenRecordset StrSql, rsCausal
    
    cantCausales = rsCausal.RecordCount
    cantEmpleados = 0
    
    'EAM- Recorre todos los causales activos para calcular el total de empleados
    Do While Not rsCausal.EOF
        'EAM- Obtiene el total de empleados a procesar
        cantEmpleadosAux = CantEmpXCausal(rsCausal("causalcnivel"), rsCausal("causnro"), procTodosEmp, sqlAux, FechaProc)
        If cantEmpleadosAux > 0 Then
            cantEmpleados = cantEmpleados + cantEmpleadosAux
        End If
        rsCausal.MoveNext
        
        'EAM- Arma un array con las sql que obtiene los empleados.
        ReDim Preserve arrSqlEmpCausales(UBound(arrSqlEmpCausales()) + 1)
        arrSqlEmpCausales(UBound(arrSqlEmpCausales)) = sqlAux
    Loop
    
    Flog.writeline Espacios(Tabulador * 0) & "Cantidad de empleados alcanzados por el causal: " & cantEmpleados
    'Obtiene el avance del progreso
    If cantEmpleados > 0 Then
        'IncPorc = (100 / ((cantEmpleados * cantCausales) * 2))
        IncPorc = (50 / cantEmpleados)
    End If
    
          
    'Activo el manejador de errores
    On Error GoTo CE
    
    rsCausal.MoveFirst
    Progreso = 0
    i = 1
    
    'Arma la fecha para registrar todas las sanciones y poder identificar todo lo generado por el proceso
    Hora = Format(CStr(Hour(Now) & Minute(Now)), "0000")
    
    
    '_____________________________________________________________________________________________________
    'EAM- Borra todas las Sanciones dentro de de los períodos de configurados
    Do While Not rsCausal.EOF
    
        'EAM- Obtiene todos los empleados del causal
        OpenRecordset arrSqlEmpCausales(i), rsEmpCausal
        
        'Cicla todos los empleado alcanzados por el causal
        Do While Not rsEmpCausal.EOF
            'Chequea que el empleado este activo
            If empleadoActivo(rsEmpCausal!ternro) And faseActiva(rsEmpCausal!ternro, FechaProc) Then
            
            'EAM- Busca todos los criterios activos
            StrSql = "SELECT * FROM dis_criterio WHERE critestado=-1 and causnro = " & rsCausal("causnro")
            OpenRecordset StrSql, rsCriterio
            Do While Not rsCriterio.EOF
                
                'EAM- Busca todas las condiciones de los criterios
                StrSql = "SELECT * FROM dis_condicion WHERE critnro= " & rsCriterio("critnro") & " ORDER BY perievnro ASC"
                OpenRecordset StrSql, rsCondicion
                Do While Not rsCondicion.EOF
                                    
                    'EAM- Obtiene las fechas desde y hasta del período
                    Call buscarFechaPeriodo(rsCondicion!perievnro, FechaProc, FechaDesde, FechaHasta, rsEmpCausal!ternro)
                    'EAM- Le asigna como fecha de corte la seleccionada en el proceso.
                    If (FechaHasta = Empty) Then
                        'rsCondicion.MoveNext
                        Flog.writeline Espacios(Tabulador * 1) & "El empleado tiene la fase cerrada a la fecha: " & FechaProc
                    Else
                        FechaHasta = CDate(FechaProc)
                        Flog.writeline Espacios(Tabulador * 1) & "Período de Evaluacion desde: " & FechaDesde & " hasta: " & FechaHasta
                
                        Call BorrarSancionExistentes(rsEmpCausal!ternro, rsCondicion!condnro, FechaDesde, FechaHasta, rsCausal!causnro, rsCausal!gradnro, 1)
                        Flog.writeline Espacios(Tabulador * 0) & "Se eliminaron todas las sanciones desde " & FechaDesde & " a la fecha " & FechaHasta
                        Cantidad_de_OpenRecordset = Cantidad_de_OpenRecordset + 1
                    End If
                    
                    rsCondicion.MoveNext
                Loop
                rsCriterio.MoveNext
            Loop
            End If
            rsEmpCausal.MoveNext
        Loop
        rsCausal.MoveNext
        i = i + 1
        Progreso = Progreso + IncPorc
        TiempoAcumulado = GetTickCount
        If CantidadEmpleados > 0 Then
                 CantidadEmpleados = CantidadEmpleados - 1
            End If
         'StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & ", bprcempleados = " & CantidadEmpleados & " WHERE bpronro = " & NroProceso
         StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
                    ", bprcempleados = " & CantidadEmpleados & _
                    ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
                    "' WHERE bpronro = " & NroProceso


            objConn.Execute StrSql, , adExecuteNoRecords

    Loop
    'Flog.writeline Espacios(Tabulador * 0) & "Se eliminaron todas las sanciones hasta la fecha " & FechaHasta
    Flog.writeline "-------------------------------------------------------------"
    Flog.writeline
    
    
    '_____________________________________________________________________________________________________
    'EAM- Genera las Sanciones
    i = 1
    rsCausal.MoveFirst
        
    
    Do While Not rsCausal.EOF
        'EAM- Obtiene todos los empleados del causal
        OpenRecordset arrSqlEmpCausales(i), rsEmpCausal
        
        Flog.writeline Espacios(Tabulador * 0) & "Causal: " & rsCausal("causnro")
        'Setea la variable en falso para indicar que todavía no se aplico reducción de graduacion.
        aplicoReduccion = False
        
        'Cicla todos los empleado alcanzados por el causal
        Do While Not rsEmpCausal.EOF
            Flog.writeline Espacios(Tabulador * 0) & "  Tercero: " & rsEmpCausal("ternro")
            
            'Chequea que el empleado este activo
            If empleadoActivo(rsEmpCausal!ternro) And faseActiva(rsEmpCausal!ternro, FechaProc) Then
'            If Not empleadoActivo(rsEmpCausal!ternro) Then
'                Flog.writeline Espacios(Tabulador * 1) & "El empleado se encuentra inactivo. Ternro: " & rsEmpCausal!ternro
'            Else
            
            'EAM- Busca todos los criterios activos
            StrSql = "SELECT * FROM dis_criterio WHERE critestado=-1 and causnro = " & rsCausal("causnro")
            OpenRecordset StrSql, rsCriterio
            Do While Not rsCriterio.EOF
                
                'EAM- Busca todas las condiciones de los criterios
                StrSql = "SELECT * FROM dis_condicion WHERE critnro= " & rsCriterio("critnro") & " ORDER BY perievnro ASC"
                OpenRecordset StrSql, rsCondicion
                Do While Not rsCondicion.EOF
                    Flog.writeline Espacios(Tabulador * 0) & "    Criterio: " & rsCriterio("critnro") & " Condición: " & rsCondicion("condnro")
                    
                    'EAM- Obtiene las fechas desde y hasta del período
                    Call buscarFechaPeriodo(rsCondicion!perievnro, FechaProc, FechaDesde, FechaHasta, rsEmpCausal!ternro)
                    
                    'EAM- Le asigna como fecha de corte la seleccionada en el proceso.
                    If (FechaHasta = Empty) Then
                        'rsCondicion.MoveNext
                        Flog.writeline Espacios(Tabulador * 1) & "El empleado tiene la fase cerrada a la fecha: " & FechaProc
                    Else
                        'EAM- Le asigna como fecha de corte la seleccionada en el proceso.
                        FechaHasta = CDate(FechaProc)
                        Flog.writeline Espacios(Tabulador * 1) & "Período de Evaluacion desde: " & FechaDesde & " hasta: " & FechaHasta
                                            
                        'EAM- Verifica si existen sanciones posteriores
                        StrSql = "SELECT hsannro FROM dis_his_sancion WHERE hsandesde >=" & ConvFecha(FechaHasta) & " AND ternro= " & rsEmpCausal!ternro
                        OpenRecordset StrSql, rsAux
                        
                        If rsAux.EOF Then
                            'EAM- Si no existen sanciones analiza si hay sanciones para el rango de fecha
                            If IsNull(rsCondicion!condtipocantagrup) Or (Trim(rsCondicion!condtipocantagrup) = "") Then
                                condtipocantagrup = 0
                            Else
                                condtipocantagrup = rsCondicion!condtipocantagrup
                            End If
                            
                            Call GenerarSancion(Hora, rsCondicion!condnro, FechaDesde, FechaHasta, rsEmpCausal!ternro, rsCondicion!condactagup, rsCondicion!condcantagrup, condtipocantagrup, rsCondicion!conduniagrup, aplicoReduccion)
                            Cantidad_de_OpenRecordset = Cantidad_de_OpenRecordset + 1
                        Else
                            Flog.writeline Espacios(Tabulador * 1) & "Hay Sanciones postriores a la fecha : " & FechaDesde & " y no se porcesará."
                        End If
                    End If
                    rsCondicion.MoveNext
                Loop
                rsCriterio.MoveNext
            Loop
            End If
            'Actualizo el progreso del proceso
            
            If CantidadEmpleados > 0 Then
                 CantidadEmpleados = CantidadEmpleados - 1
            End If
          ' StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & ", bprcempleados = " & CantidadEmpleados & " WHERE bpronro = " & NroProceso
           
             Progreso = Progreso + IncPorc
             TiempoAcumulado = GetTickCount
         'StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & ", bprcempleados = " & CantidadEmpleados & " WHERE bpronro = " & NroProceso
         StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
                    ", bprcempleados = " & CantidadEmpleados & _
                    ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
                    "' WHERE bpronro = " & NroProceso
                     objConn.Execute StrSql, , adExecuteNoRecords
            rsEmpCausal.MoveNext
        Loop
    
        
        Flog.writeline Espacios(Tabulador * 0) & "---------------------------------------------------------------------------"
SiguienteEmpleado:
        rsCausal.MoveNext
        i = i + 1
    Loop
    
'Borra los empleados de la tabla batch_empleado
StrSql = "DELETE FROM batch_empleado WHERE bpronro = " & NroProceso
objConn.Execute StrSql, , adExecuteNoRecords



'Deshabilito el manejador de errores
On Error GoTo 0

 'Habilito manejador gral
 On Error GoTo ME_Main
 
 StrSql = "DELETE FROM Batch_Procacum WHERE bpronro = " & NroProceso
 objConn.Execute StrSql, , adExecuteNoRecords
 
 ' Actualizo el Btach_Proceso
 If Not HuboErrores Then
     StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcempleados = 0 , bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
 Else
     StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ",bprcempleados = 0 , bprcestado = 'Incompleto' WHERE bpronro = " & NroProceso
 End If
 objConn.Execute StrSql, , adExecuteNoRecords

 ' -----------------------------------------------------------------------------------
 'Copio el proceso en el historico de batch_proceso y lo borro de batch_proceso
 If Not HuboErrores Then
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
 End If
 ' -----------------------------------------------------------------------------------



Final:
    If depurar Then
        If CnTraza.State = adStateOpen Then CnTraza.Close
    End If
       
    If myrs.State = adStateOpen Then myrs.Close
    Set myrs = Nothing
    objConn.Close
    Set objConn = Nothing
    
    If objRs.State = adStateOpen Then objRs.Close
    Set objRs = Nothing
    If rs_Per.State = adStateOpen Then rs_Per.Close
    Set rs_Per = Nothing
    
    Flog.writeline
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "Fin :" & Now
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "---------------------------------------------------------------------------------"
    Flog.writeline "Cantidad de Lecturas en BD          : " & Cantidad_de_OpenRecordset
    Flog.writeline "Cantidad de llamadas a politicas    : " & Cantidad_Call_Politicas
    Flog.writeline Espacios(Tabulador * 0) & "---------------------------------------------------------------------------------"
    Flog.Close
    
    
Exit Sub
    
CE:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 0) & "Error. Empleado abortado " & " " & FechaProc
    Flog.writeline Espacios(Tabulador * 0) & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline

    HuboErrores = True
    GoTo SiguienteEmpleado

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
    MyBeginTrans
       
        If CantidadEmpleados > 0 Then
           CantidadEmpleados = CantidadEmpleados - 1
         End If
        StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcempleados = " & CantidadEmpleados & ", bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Error General', bprcpid = " & PID & " WHERE bpronro = " & NroProceso
        objConn.Execute StrSql, , adExecuteNoRecords
    MyCommitTrans
End Sub


'EAM- Obtiene la cantidad de empleados alcanczado por el causal segun el proceso
Private Function CantEmpXCausal(ByVal causalcnivel, ByVal causnro As Long, ByVal TodosEmp, ByRef sql, ByVal FechaProc As Date) As Long
 Dim Cantidad As Long
 Dim rsCantEmp As New ADODB.Recordset

    Select Case causalcnivel
        Case 1: 'global
            If (TodosEmp = 0) Then
                StrSql = "SELECT count(ternro) cant FROM batch_empleado WHERE batch_empleado.bpronro = " & NroProceso
                OpenRecordset StrSql, rsCantEmp
            Else
                StrSql = "SELECT count(ternro) cant FROM empleado WHERE empest=-1"
                OpenRecordset StrSql, rsCantEmp
            End If
            sql = "SELECT * " & Mid(StrSql, InStrRev(StrSql, "FROM"))

        Case 2: 'Estructura
            If (TodosEmp = 0) Then
                StrSql = " SELECT count(his_estructura.ternro) cant FROM dis_causal" & _
                        " INNER JOIN dis_causal_alcance ON dis_causal.causnro = dis_causal_alcance.causnro" & _
                        " INNER JOIN his_estructura ON dis_causal_alcance.causalcorigen = his_estructura.estrnro " & _
                        " WHERE dis_causal.causnro = " & causnro & " And (his_estructura.htetdesde <= " & ConvFecha(FechaProc) & ") AND " & _
                        " ((" & ConvFecha(FechaProc) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))" & _
                        " AND his_estructura.ternro IN (SELECT ternro FROM batch_empleado WHERE batch_empleado.bpronro = " & NroProceso & ")"
                OpenRecordset StrSql, rsCantEmp
                
                sql = "SELECT ternro ternro " & Mid(StrSql, InStr(StrSql, "FROM"))
            Else
                StrSql = " SELECT count(his_estructura.ternro) cant FROM dis_causal" & _
                        " INNER JOIN dis_causal_alcance ON dis_causal.causnro = dis_causal_alcance.causnro" & _
                        " INNER JOIN his_estructura ON dis_causal_alcance.causalcorigen = his_estructura.estrnro " & _
                        " WHERE dis_causal.causnro = " & causnro & " And (his_estructura.htetdesde <= " & ConvFecha(FechaProc) & ") AND " & _
                        " ((" & ConvFecha(FechaProc) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
                        '& " AND htetdesde >= " & ConvFecha(FechaProc) & " AND htethasta <= " & ConvFecha(FechaProc)
                        
        
                OpenRecordset StrSql, rsCantEmp
                
                sql = "SELECT ternro ternro " & Mid(StrSql, InStrRev(StrSql, "FROM"))
                'sql = Mid(sql, 1, Len(sql) - 1)
            End If
           
        Case 3 'Individual
            If (TodosEmp = 0) Then
                StrSql = "SELECT count(causalcorigen) cant FROM dis_causal " & _
                        " INNER JOIN dis_causal_alcance ON dis_causal.causnro = dis_causal_alcance.causnro " & _
                        " WHERE dis_causal.causnro = " & causnro & _
                        " AND causalcorigen IN (SELECT ternro FROM batch_empleado WHERE batch_empleado.bpronro = " & NroProceso & ")"
                OpenRecordset StrSql, rsCantEmp
            Else
                StrSql = "SELECT count(causalcorigen) cant FROM dis_causal_alcance WHERE dis_causal_alcance.causnro= " & causnro
                OpenRecordset StrSql, rsCantEmp
            End If
            sql = "SELECT causalcorigen,causalcorigen ternro " & Mid(StrSql, InStr(StrSql, "FROM"))
    End Select
    
    
    If Not rsCantEmp.EOF Then
        Cantidad = rsCantEmp("cant")
    Else
        Cantidad = 0
    End If
    
    rsCantEmp.Close
    CantEmpXCausal = Cantidad

End Function

'EAM- Obtiene la fecha desde y hasta según el periodo. El cálculo es en funcion del dia en el que esta parado el proceso
Private Sub buscarFechaPeriodo(ByVal perievnro, ByVal fdesdeProc As Date, ByRef fdesde As Date, ByRef fhasta As Date, ByVal ternro As Long)
 Dim rsPeriEv As New ADODB.Recordset
 Dim auxDia As Date
    
    'Obtiene la cant de dias del período de Evaluación
    StrSql = "SELECT perievdias FROM dis_peri_ev WHERE perievnro= " & perievnro
    OpenRecordset StrSql, rsPeriEv
    
    If Not rsPeriEv.EOF Then
        Select Case rsPeriEv("perievdias")
        
            Case Is = -1: 'Zafra
                    rsPeriEv.Close
                    StrSql = "SELECT altfec,bajfec FROM empleado " & _
                            " INNER JOIN fases ON empleado.ternro = fases.empleado " & _
                            " WHERE ternro = " & ternro & " And empleado.empest = -1 And fases.estado = -1 AND altfec<= " & ConvFecha(fdesdeProc) & _
                            " AND (bajfec>= " & ConvFecha(fdesdeProc) & " OR bajfec IS NULL) ORDER BY altfec DESC"
                    OpenRecordset StrSql, rsPeriEv
                    
                    'EAM- Si tiene fase activa toma la fecha desde como inicio
                    If Not rsPeriEv.EOF Then
                        If Not IsNull(rsPeriEv!bajfec) Then
                            fdesde = rsPeriEv!altfec
                            fhasta = rsPeriEv!bajfec
                            
                            If DateDiff("d", fdesde, fhasta) > 365 Then
                                Flog.writeline Espacios(Tabulador * 0) & "La fase del empleado: " & ternro & " supera los 365 días."
                            End If
                        Else
                            fdesde = rsPeriEv!altfec
                            fhasta = fdesdeProc
                                                
                            If DateDiff("d", fdesde, fhasta) > 365 Then
                                Flog.writeline Espacios(Tabulador * 0) & "La fase del empleado: " & ternro & " supera los 365 días."
                            End If
                        End If
                    Else
                        rsPeriEv.Close
                        StrSql = "SELECT altfec,bajfec FROM empleado " & _
                                "INNER JOIN  fases ON empleado.ternro = fases.empleado " & _
                                "WHERE empleado =" & ternro & "  And empleado.empest = -1 AND altfec<= " & ConvFecha(fdesdeProc) & _
                                " AND (bajfec>= " & ConvFecha(fdesdeProc) & " OR bajfec IS NULL) ORDER BY altfec DESC"
                        OpenRecordset StrSql, rsPeriEv
                        
                        If Not rsPeriEv.EOF Then
                            If Not IsNull(rsPeriEv!bajfec) Then
                                Flog.writeline Espacios(Tabulador * 0) & "El empleado: " & ternro & " tiene la fase activa actualmente."
                                fdesde = rsPeriEv!altfec
                                fhasta = rsPeriEv!bajfec
                                
                                If DateDiff("d", fdesde, fhasta) > 365 Then
                                    Flog.writeline Espacios(Tabulador * 0) & "La fase del empleado: " & ternro & " supera los 365 días."
                                End If
                            Else
                                Flog.writeline Espacios(Tabulador * 0) & "El empleado: " & ternro & " tiene la fase cerrada a la fecha: " & rsPeriEv!bajfec
                                fdesde = rsPeriEv!altfec
                                fhasta = fdesdeProc
                                                    
                                If DateDiff("d", fdesde, fhasta) > 365 Then
                                    Flog.writeline Espacios(Tabulador * 0) & "La fase del empleado: " & ternro & " supera los 365 días."
                                End If
                            End If
                        Else
                            fhasta = Empty
                        End If
                    End If
            Case Is <= 7: 'Semanal
                'EAM- Obtiene la fecha desde y hasta de la semana en que se esta procesando
                    fdesde = DateAdd("d", -(Weekday(fdesdeProc) - 1), fdesdeProc)
                    fhasta = DateAdd("d", 6, fdesde)
                
            Case Is < 30: 'Quincenal
                'EAM- Si es menor, es la primer Quincena
                If (Day(fdesdeProc) <= rsPeriEv("perievdias")) Then
                    fdesde = DateAdd("d", -(Day(fdesdeProc) - 1), fdesdeProc)
                    fhasta = DateAdd("d", 14, fdesde)
                Else
                    '2° Quincena
                    fdesde = DateAdd("d", -(Day(fdesdeProc) - 16), fdesdeProc)
                    auxDia = DateAdd("m", 1, fdesdeProc)
                    fhasta = DateAdd("d", -Day(fdesdeProc), auxDia)
                End If
            
            Case 30: 'Mensual
                    'Calcula la fecha desde
                    fdesde = CDate("01/" & Month(fdesdeProc) & "/" & Year(fdesdeProc))
                    'Calcula la fecha hasta
                    auxDia = DateAdd("m", 1, fdesde)
                    fhasta = DateAdd("d", -1, auxDia)
                                                    
            Case 365: 'Anual
                    'Calcula la fecha desde y hasta
                    fdesde = CDate("01/01/" & Year(fdesdeProc))
                    fhasta = CDate("31/12/" & Year(fdesdeProc))
        End Select
    Else
        fdesdeProc = Empty
    End If
    
    'Flog.writeline Espacios(Tabulador * 1) & "Período de Evaluacion desde: " & fdesde & " hasta: " & fhasta
    rsPeriEv.Close
End Sub


'EAM- Genera la Sancion
'Private Sub GenerarSancion(ByVal condnro As Long, ByVal fdesde As Date, ByVal fhasta As Date, ByVal cantmax As Double, _
'                            ByVal valmax As Double, ByVal causnro As Long, ByVal ternro As Long)
Private Sub GenerarSancion(ByVal Hora As String, ByVal condnro As Long, ByVal fdesde As Date, ByVal fhasta As Date, _
                ByVal ternro As Long, ByVal agruparActivo As Integer, ByVal cantagrup As Long, ByVal tipocantagrup As String, _
                ByVal uniagrup As Integer, ByRef aplicoReduccion As Boolean)
 Dim rsDatos As New ADODB.Recordset
 Dim rsAD As New ADODB.Recordset
 Dim rsAux As New ADODB.Recordset
 Dim codEstado As Integer
 Dim cantAcumulada As Double
 Dim nroSancion As Long
 Dim cantSanciones As Integer
 Dim fecInfraccion As Date
 Dim i As Integer
 Dim sancionesNotif As Long
 Dim gradorden As Integer
 Dim fechaFinAux As Date
 Dim contoInfraccion As Boolean
 Dim cantInfraccion As Long
 Dim ConteplarInfraccion As Boolean
 Dim fecInfraccSancion As String 'Guarda las fechas de las infracciones que disparan las sanciones
 Dim fecInfraccSancionAUX
     
     'EAM- Obtiene los datos para dar de alta la sancion
     StrSql = "SELECT cantmax,valmax,dis_causal.causnro, dis_causal.gradnro,condorigennro FROM dis_condicion " & _
            " INNER JOIN dis_criterio ON dis_condicion.critnro = dis_criterio.critnro " & _
            " INNER JOIN dis_causal ON dis_causal.causnro = dis_criterio.causnro " & _
            " WHERE dis_condicion.condnro = " & condnro
    OpenRecordset StrSql, rsDatos
     
    'EAM- Obtiene el primer estado del de la sancion
     StrSql = "SELECT estsancodigo FROM dis_estado_sancion ORDER BY estsancodigo ASC"
     OpenRecordset StrSql, rsAux
     If Not rsAux.EOF Then
        codEstado = rsAux!estsancodigo
     End If
     rsAux.Close
     
 
    'EAM- Obtiene todas las horas infraccionables en el rango de fecha y no estan como dias excluidos
    StrSql = "SELECT adcanthoras,adfecha FROM gti_acumdiario WHERE infraccion = -1 AND adfecha >= " & ConvFecha(fdesde) & " AND adfecha <= " & ConvFecha(fhasta) & _
            " AND gti_acumdiario.ternro= " & ternro & " AND thnro= " & rsDatos!condorigennro & _
            " AND adfecha NOT IN (SELECT dexcdia FROM dis_dias_excluidos WHERE dexcdia >=" & ConvFecha(fdesde) & " AND dexcdia <= " & ConvFecha(fhasta) & ")" & _
            " ORDER BY adfecha ASC "
    OpenRecordset StrSql, rsAD
    
    'EAM- Si es EOF, es porque no tiene horas con infracción.
    If rsAD.EOF Then
        Exit Sub
    End If

    
    'Inicializa las variables
    cantSanciones = 0
    cantInfraccion = 0
    cantAcumulada = 0
    ConteplarInfraccion = True
    
    'EAM- Si tiene activa la agrupacion setea las fechas para comenzar a analizar.
    If agruparActivo = -1 Then
        contoInfraccion = False
        fechaFinAux = DateAdd(tipocantagrup, cantagrup, fdesde)
        
        'EAM- controla para el caso de los zafrales que pueden tener la fase de alta mitad de mes
        fechaFinAux = CDate("01" & "/" & Month(fechaFinAux) & "/" & Year(fechaFinAux))
    End If
                    
   
    
    'EAM- Recorre todas las horas de infracción y las analiza par ver si hay que generar Sanción.
    Do While Not rsAD.EOF
            
        cantAcumulada = cantAcumulada + rsAD!adcanthoras
        'aplicoReduccion = False
        
        
        'EAM- Si es menor o igual que 0 no se Evalúa el criterio por Repetición
        If rsDatos!cantmax > 0 Then
        
        
            'EAM- si tiene la agrupacion activa, calcula las infracciones segun la configuración
            If agruparActivo = -1 Then
                If rsAD!adfecha < fechaFinAux Then
                    If Not contoInfraccion Then
                        cantInfraccion = cantInfraccion + uniagrup
                        contoInfraccion = True
                        ConteplarInfraccion = True
                    Else
                        ConteplarInfraccion = False
                    End If
                Else
                    contoInfraccion = True
                    fechaFinAux = DateAdd(tipocantagrup, cantagrup, rsAD!adfecha)
                    fechaFinAux = CDate("01" & "/" & Month(fechaFinAux) & "/" & Year(fechaFinAux))
                    cantInfraccion = cantInfraccion + uniagrup
                    ConteplarInfraccion = True
                End If
            Else
                'EAM- Sin agrupacion
                ConteplarInfraccion = True
                cantInfraccion = cantInfraccion + 1
            End If
        

            'EAM- Si hay mas infracciones que lo permitido y se contempla los analaliza.
            If (cantInfraccion > rsDatos!cantmax) And (ConteplarInfraccion) Then
                                                        
                If (Not aplicoReduccion) Then
                    Call ReducirGraduacionSancion(rsDatos!causnro, rsDatos!condorigennro, ternro, rsAD!adfecha)
                    aplicoReduccion = True
                End If
                Flog.writeline Espacios(Tabulador * 2) & "Fecha de Infraccion : " & rsAD!adfecha
                
                'EAM- Actualiza el grado de sancion para generar la nueva sancion
                gradorden = ActualizarGradoSancion(ternro, rsDatos!causnro, rsDatos!gradnro, 1, 0)
                Flog.writeline Espacios(Tabulador * 1) & "Cantidad de Infracciones : " & cantInfraccion & " y se generarán una sanción por Repetición de orden" & gradorden & "."
                
                'Inserto la sancion para la condicion.
                StrSql = "INSERT INTO dis_his_sancion (ternro,hsananio,fasnro,causnro,condnro,hsandesde,hsanhasta,estsancodigo " & _
                        ",hsanfecinfra,hsanfechagen,gradorden,hsanhora,hsantipocond )" & _
                        " VALUES (" & ternro & "," & Year(fdesde) & ",''," & rsDatos!causnro & "," & condnro & "," & ConvFecha(fdesde) & _
                        "," & ConvFecha(fhasta) & "," & codEstado & "," & ConvFecha(rsAD!adfecha) & "," & ConvFecha(Now) & "," & gradorden & "," & Hora & ",1)"
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
        End If
    
    
        'EAM- si es menor o igual que 0 no se Evalúa el criterio por Acumulación
        If rsDatos!valmax > 0 Then
    
            'EAM- Evalua la infracciones por Cantidad Acumulada
            If ((cantAcumulada * 60) > rsDatos!valmax) Then
                'Calcula la cantidad de sanciones a generar
                cantSanciones = Int((cantAcumulada * 60) / rsDatos!valmax)
                Flog.writeline Espacios(Tabulador * 1) & "Cantidad Acumulada de Infracciones : " & (cantAcumulada * 60) & " y se generarán " & cantSanciones & " sanciones por acumulacion."
                
                'EAM- Descuenta en la cantidad acumulada lo que va a generar para no volver a tenerlo en cuenta
                'cantAcumulada = cantAcumulada - (cantSanciones * 60)
                 cantAcumulada = cantAcumulada - (cantSanciones)
                'Verifica si hay que reducir la escala por no tener infracciones
                If (Not aplicoReduccion) And (cantSanciones > 0) Then
                    Call ReducirGraduacionSancion(rsDatos!causnro, rsDatos!condorigennro, ternro, rsAD!adfecha)
                    aplicoReduccion = True
                End If
                
                                    
                For i = 1 To cantSanciones
                    'EAM- Actualiza el grado de sancion para generar la nueva sancion
                    gradorden = ActualizarGradoSancion(ternro, rsDatos!causnro, rsDatos!gradnro, 2, 0)
                    
                    'Inserto la sancion para la condicion.
                    StrSql = "INSERT INTO dis_his_sancion (ternro,hsananio,fasnro,causnro,condnro,hsandesde,hsanhasta,estsancodigo " & _
                            ",hsanfecinfra,hsanfechagen,gradorden,hsanhora,hsantipocond )" & _
                            " VALUES (" & ternro & "," & Year(fdesde) & ",''," & rsDatos!causnro & "," & condnro & "," & ConvFecha(fdesde) & _
                            "," & ConvFecha(fhasta) & "," & codEstado & "," & ConvFecha(rsAD!adfecha) & "," & ConvFecha(Now) & "," & gradorden & "," & Hora & ",2)"
                    objConn.Execute StrSql, , adExecuteNoRecords
                Next
                
            End If
        End If
        
        rsAD.MoveNext
    Loop
End Sub

'Restaura los grados de la sancion antes de haberse generado las sanciones borradas
Private Function RestaurarGradoSancion(ByVal ternro As Long, ByVal causnro As Long, ByVal gradnro As Long, ByVal hsantipocond As Integer, ByVal reducirEscala As Integer)
 Dim rsDatos As New ADODB.Recordset
 
    'EAM- Obtiene el orden de la última sanción
    StrSql = "SELECT gradorden FROM dis_grad_san_emp WHERE causnro=" & causnro & " AND ternro= " & ternro
    OpenRecordset StrSql, rsDatos
    
    If Not rsDatos.EOF Then
        rsDatos.Close
        
        'Si es -1 esta reprocesando o eliminado una sancion y reduce la escala.
        If (reducirEscala = -1) Then
            'Obtiene el grado de la último sancion generada
            StrSql = "SELECT gradorden,hsantipocond FROM dis_his_sancion dhs WHERE causnro=" & causnro & " AND ternro=" & ternro & " ORDER BY hsanfecinfra DESC"
            OpenRecordset StrSql, rsDatos
            
            If Not rsDatos.EOF Then
                'EAM- Actualiza la nueva gradaución para el empleado en la causal analizada
                StrSql = "UPDATE dis_grad_san_emp SET gradorden= " & rsDatos("gradorden") & ",hsantipocond = " & rsDatos("hsantipocond") & _
                        " WHERE ternro= " & ternro & " AND causnro= " & causnro
                objConn.Execute StrSql, , adExecuteNoRecords
                Flog.writeline Espacios(Tabulador * 1) & "Se restauro despues de la eliminación al grado: " & rsDatos("gradorden")
             Else
                'EAM- Si el orden es 0 quiere decir que es el primer nivel y hay que eliminarlo
                StrSql = "DELETE FROM dis_grad_san_emp WHERE ternro= " & ternro & " AND causnro= " & causnro '& " AND hsantipocond= " & hsantipocond
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
        End If
    End If
    
    rsDatos.Close
    
        
End Function

'Actualiza la tabla el grado de la sancion y devuele el grado
Private Function ActualizarGradoSancion(ByVal ternro As Long, ByVal causnro As Long, ByVal gradnro As Long, ByVal hsantipocond As Integer, ByVal reducirEscala As Integer)
 Dim rsDatos As New ADODB.Recordset
 Dim gradorden As Integer
 
 
    'EAM- Obtiene el orden de la última sanción
    StrSql = "SELECT gradorden FROM dis_grad_san_emp WHERE causnro=" & causnro & " AND ternro= " & ternro
    OpenRecordset StrSql, rsDatos
    
    If rsDatos.EOF Then
        'EAM- Inserta la nueva gradaucion para el empleado en la causal analizada
        StrSql = "INSERT INTO dis_grad_san_emp (ternro,causnro,gradorden,fecha,hsantipocond) " & _
                " VALUES(" & ternro & "," & causnro & "," & 0 & "," & ConvFecha(Now) & "," & hsantipocond & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        Flog.writeline Espacios(Tabulador * 1) & "El orden de la graduacion siguiente es : 0"
    Else
        gradorden = rsDatos("gradorden")
        rsDatos.Close
        
        'Si es -1 esta reprocesando o eliminado una sancion y reduce la escala.
        If (reducirEscala = -1) Then
            'Busca el siguiente nivel menor de la escala del que tiene
            StrSql = "SELECT gradorden FROM dis_graduacion_det WHERE gradnro =" & gradnro & " AND gradorden <" & gradorden & _
                    " ORDER BY gradorden DESC"
            OpenRecordset StrSql, rsDatos
            
            If Not rsDatos.EOF Then
                'EAM- Actualiza la nueva gradaución para el empleado en la causal analizada
                StrSql = "UPDATE dis_grad_san_emp SET gradorden= " & rsDatos("gradorden") & ",hsantipocond = " & hsantipocond & _
                        " WHERE ternro= " & ternro & " AND causnro= " & causnro
                objConn.Execute StrSql, , adExecuteNoRecords
                
                gradorden = rsDatos("gradorden")
                
                'Flog.writeline Espacios(Tabulador * 1) & "El orden de la graduacion es : " & rsDatos("gradorden")
            Else
                StrSql = "DELETE FROM dis_grad_san_emp WHERE ternro=" & ternro & " AND causnro= " & causnro
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
                                                           
        Else
            
            'Busca el siguiente nivel de la escala
            StrSql = "SELECT gradorden FROM dis_graduacion_det WHERE gradnro =" & gradnro & " AND gradorden >" & gradorden & " ORDER BY gradorden ASC"
            OpenRecordset StrSql, rsDatos
            
            If Not rsDatos.EOF Then
                'EAM- Actualiza la nueva gradaucion para el empleado en la causal analizada
                StrSql = "UPDATE dis_grad_san_emp SET gradorden= " & rsDatos("gradorden") & ",hsantipocond = " & hsantipocond & _
                        " WHERE ternro= " & ternro & " AND causnro= " & causnro
                objConn.Execute StrSql, , adExecuteNoRecords
                Flog.writeline Espacios(Tabulador * 1) & "El orden de la graduacion siguiente es : " & rsDatos("gradorden")
                
                gradorden = rsDatos("gradorden")
            Else
                'EAM- Actualiza la nueva gradaucion para el empleado en la causal analizada
                StrSql = "UPDATE dis_grad_san_emp SET gradorden= " & gradorden & ",hsantipocond = " & hsantipocond & _
                        " WHERE ternro= " & ternro & " AND causnro= " & causnro
                objConn.Execute StrSql, , adExecuteNoRecords
                
                Flog.writeline Espacios(Tabulador * 1) & "El orden de la graduacion siguiente a :" & gradorden & " no existe."
            End If
        End If
        
        ActualizarGradoSancion = gradorden
        
    End If
End Function

'EAM- Verifica si ya hay sanciones generadas para el rango de fechas y las elimina y actualiza el grado de sancion
Private Sub BorrarSancionExistentes(ByVal ternro As Long, ByVal condnro As Long, ByVal fdesde As Date, ByVal fhasta As Date, ByVal causnro As Long, ByVal gradnro As Integer, ByVal hsantipocond As Integer)
 Dim rsAux As New ADODB.Recordset
 
    'Verifica si ya hay sanciones generadas para el rango de fechas y las elimina y actualiza el grado de sancion
    StrSql = "SELECT hsannro,hsannotifecha,hsanlicfecha FROM dis_his_sancion WHERE condnro= " & condnro & " AND hsandesde >= " & _
            ConvFecha(fdesde) & " AND hsanhasta <= " & ConvFecha(fhasta) & " AND ternro= " & ternro & " ORDER BY hsanfecinfra DESC "
    OpenRecordset StrSql, rsAux
    
    
    Do While Not rsAux.EOF
        If (IsNull(rsAux("hsannotifecha")) And IsNull(rsAux("hsanlicfecha"))) Then
            'Reprocesa la sancion porque no esta notificada
            StrSql = "DELETE FROM dis_his_sancion WHERE hsannro= " & rsAux!hsannro
            objConn.Execute StrSql, , adExecuteNoRecords
            
            Call RestaurarGradoSancion(ternro, causnro, gradnro, hsantipocond, -1)
        End If
            
        rsAux.MoveNext
    Loop
End Sub

''EAM- Verifica si ya hay sanciones generadas para el rango de fechas y las elimina y actualiza el grado de sancion
'Private Sub BorrarSancionExistentes(ByVal ternro As Long, ByVal condnro As Long, ByVal fdesde As Date, ByVal fhasta As Date, ByVal causnro As Long, ByVal gradnro As Integer, ByVal hsantipocond As Integer)
' Dim rsAux As New ADODB.Recordset
'
'    'Verifica si ya hay sanciones generadas para el rango de fechas y las elimina y actualiza el grado de sancion
'    StrSql = "SELECT hsannro,hsannotifecha,hsanlicfecha FROM dis_his_sancion WHERE condnro= " & condnro & " AND hsandesde= " & _
'            ConvFecha(fdesde) & " AND hsanhasta= " & ConvFecha(fhasta) & " AND hsantipocond= " & hsantipocond & " and ternro= " & ternro
'    OpenRecordset StrSql, rsAux
'
'
'    Do While Not rsAux.EOF
'        If (IsNull(rsAux("hsannotifecha")) And IsNull(rsAux("hsanlicfecha"))) Then
'            Call ActualizarGradoSancion(ternro, causnro, gradnro, hsantipocond, -1)
'
'            'Reprocesa la sancion porque no esta notificada
'            StrSql = "DELETE FROM dis_his_sancion WHERE hsannro= " & rsAux!hsannro
'            objConn.Execute StrSql, , adExecuteNoRecords
'        End If
'
'        rsAux.MoveNext
'    Loop
'End Sub

'EAM- Obtiene las sanciones notificadas en el rango de fechas del período
Private Function BuscarSancionesNotif(ByVal ternro As Long, ByVal condnro As Long, ByVal fdesde As Date, ByVal fhasta As Date, ByVal hsantipocond As Integer)
 Dim rsAux As New ADODB.Recordset
 
    'Obtiene las sanciones notificadas en el rango de fechas
    StrSql = "SELECT SUM(hsannro) cant FROM dis_his_sancion WHERE condnro= " & condnro & " AND hsandesde= " & _
            ConvFecha(fdesde) & " AND hsanhasta= " & ConvFecha(fhasta) & " AND hsantipocond= 2 AND ternro= " & ternro & _
            "AND NOT hsannotifecha IS NULL"
    OpenRecordset StrSql, rsAux
    
    If Not IsNull(rsAux("cant")) Then
        BuscarSancionesNotif = rsAux("cant")
    Else
        BuscarSancionesNotif = 0
    End If
    
End Function



Private Sub ReducirGraduacionSancion(ByVal causnro As Long, ByVal condorigennro As Long, ByVal ternro As Long, ByVal fechasta As Date)
 Dim rsAux As New ADODB.Recordset
 Dim i As Integer
 Dim fecUltSancion As Date
 Dim cantdias As Long
 Dim cantNivelBajar As Integer
 Dim causdiasescala As Long
 Dim causreducniv As Long
    
    'EAM- Obtiene la fecha de la ultima infracción registrada
  StrSql = "SELECT hsanfecinfra FROM  dis_his_sancion WHERE causnro=" & causnro & "  AND ternro= " & ternro & " ORDER BY hsanfecinfra DESC"
    OpenRecordset StrSql, rsAux
    If Not rsAux.EOF Then
        'Calcula la diferencia entre la ultima sancion y la actual
        fecUltSancion = rsAux("hsanfecinfra")
    Else
        Flog.writeline Espacios(Tabulador * 1) & "No se realizo el cálculo de reduccion de sanciones porque no posee infracciones posteriores:"
        Exit Sub
    End If

     'Calcula la diferencia entre fechas para ver si se reduce la sancion
        If (CDate(fecUltSancion) <> CDate(fechasta)) Then
            cantdias = DateDiff("d", CDate(fecUltSancion), CDate(fechasta))
            fecUltSancion = fechasta
        End If
        Flog.writeline Espacios(Tabulador * 1) & "Reduccion. Cantidad de días: " & cantdias & ". Fecha de la últimas Infracciones " & DateAdd("d", -cantdias, fecUltSancion) & " al " & fecUltSancion
                
        'EAM- Busca la configuración de reducción de sanciones
        rsAux.Close
        StrSql = "SELECT causdiasescala,causreducniv FROM dis_causal WHERE causnro = " & causnro
        OpenRecordset StrSql, rsAux
            
        cantNivelBajar = Int(cantdias / rsAux!causdiasescala)
        causdiasescala = rsAux("causdiasescala")
        causreducniv = rsAux("causreducniv")
                            
        'EAM- Si es mayr a 0 hay que reducir la graduación de la escala
        For i = 1 To cantNivelBajar
            rsAux.Close
            StrSql = "SELECT gradorden FROM dis_grad_san_emp WHERE causnro=" & causnro & " AND ternro= " & ternro
            OpenRecordset StrSql, rsAux
            
            If Not rsAux.EOF Then
                'Si la configuracion de la reducción es 0, O la reduccion da 0, elimino el registro y comienza el calculo de nuevo
                If (causreducniv = 0) Or ((CLng(rsAux("gradorden")) - causreducniv) < 0) Then
                    StrSql = "DELETE FROM dis_grad_san_emp WHERE ternro=" & ternro & " AND causnro= " & causnro
                    objConn.Execute StrSql, , adExecuteNoRecords
                Else
                    StrSql = "UPDATE dis_grad_san_emp SET gradorden= " & (CLng(rsAux("gradorden")) - causreducniv) & _
                            " WHERE ternro= " & ternro & " AND causnro= " & causnro
                    objConn.Execute StrSql, , adExecuteNoRecords
                    
                    Flog.writeline Espacios(Tabulador * 2) & "Se redujo el orden le la graduacion en: " & causreducniv & " niveles."
                End If
            Else
                Flog.writeline Espacios(Tabulador * 2) & "No se redujo la graduación del tercero: " & ternro & " porque no posee Sanciones anteriores."
            End If
            
            
        Next
End Sub

'EAM- Retorna si el empleado está activo o no
Function empleadoActivo(ternro) As Boolean
 Dim rsAux As New ADODB.Recordset

    StrSql = "SELECT empleg FROM empleado WHERE ternro= " & ternro & " AND empest=-1"
    OpenRecordset StrSql, rsAux
    
    If Not rsAux.EOF Then
        empleadoActivo = True
    Else
        Flog.writeline Espacios(Tabulador * 1) & "El empleado se encuentra inactivo. Ternro: " & ternro
        empleadoActivo = False
    End If
 rsAux.Close
End Function

Function faseActiva(ByVal ternro As Long, ByVal fdesdeProc As Date) As Boolean
 Dim rsAux As New ADODB.Recordset

    'Verifica si tiene la fase activa el empleado
    StrSql = "SELECT * FROM fases WHERE empleado = " & ternro & " And fases.estado = -1 AND altfec<= " & ConvFecha(fdesdeProc) & _
            " AND (bajfec>= " & ConvFecha(fdesdeProc) & " OR bajfec IS NULL) ORDER BY altfec DESC"
    OpenRecordset StrSql, rsAux
    
    If rsAux.EOF Then
        Flog.writeline Espacios(Tabulador * 1) & "El empleado: " & ternro & " tiene la fase cerrada a la fecha : " & fdesdeProc
        faseActiva = False
    Else
        faseActiva = True
    End If
    rsAux.Close
    
End Function
