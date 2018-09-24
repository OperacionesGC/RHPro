Attribute VB_Name = "PedidodeCambio"
Option Explicit

Global Const Version = "1.00"
Global Const FechaModificacion = "16/11/2010" 'Martin Ferraro
Global Const UltimaModificacion = "Version Inicial"



Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial del Generador de Exportacion de Cash Management.
' Autor      : Martin Ferraro
' Fecha      : 16/11/2010
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim objconnMain As New ADODB.Connection
Dim strCmdLine
Dim Nombre_Arch As String
Dim rs_batch_proceso As New ADODB.Recordset
Dim PID As String
Dim bprcparam As String
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
    'Abro la conexion
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Then
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If
    OpenConnection strconexion, objconnProgreso
    If Err.Number <> 0 Then
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If
    On Error GoTo 0
    
    On Error GoTo ME_Main
    
    Nombre_Arch = PathFLog & "PedidoCambio-" & NroProcesoBatch & ".log"
    
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
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0 ,bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Flog.writeline "Pone el estado en procesando"
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 278 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        bprcparam = rs_batch_proceso!bprcparam
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        Call PedidoDeCambio(NroProcesoBatch, bprcparam)
    Else
        Flog.writeline "no encontró el proceso"
    End If
    
    TiempoFinalProceso = GetTickCount
    Flog.writeline "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    
    If Not HuboError Then
        StrSql = "UPDATE batch_proceso SET bprcprogreso = 100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
    Else
        StrSql = "UPDATE batch_proceso SET bprcprogreso = 100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
    End If
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
Fin:
    Flog.Close
    'If objConn.State = adStateOpen Then objConn.Close
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
    MyBeginTrans
        StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Error General', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
    MyCommitTrans
End Sub


Public Sub PedidoDeCambio(ByVal BproNro As Long, ByVal Parametros As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de Pedidos de Cambio.
' Autor      : Martin Ferraro
' Fecha      : 16/11/2010
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
'Arreglo que contiene los parametros
Dim arrParam

'Parametros desde ASP
Dim PedCambNro As Long

'RecordSet
Dim rs_Empleados As New ADODB.Recordset
Dim rs_PedCambio As New ADODB.Recordset
Dim rs_Aux As New ADODB.Recordset

'Variables
Dim montoNov As Double
Dim montoNovInsert As Double

    
' Inicio codigo ejecutable
On Error GoTo CE

' Levanto cada parametro por separado, el separador de parametros es "@"
Flog.writeline Espacios(Tabulador * 0) & "levantando parametros" & Parametros
If Not IsNull(Parametros) Then
    
    PedCambNro = Parametros
    Flog.writeline Espacios(Tabulador * 1) & "Parametros " & Parametros
    
Else
    Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encuentran los paramentros."
    HuboError = True
    Exit Sub
End If
Flog.writeline


'---------------------------------------------------------------------------------
'Busqueda del pedido de cambio
'---------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 0) & "Buscando el pedido de Cambio " & PedCambNro
StrSql = "SELECT pedcambnro,tipoalcance,codalcance,tipocambio,"
StrSql = StrSql & " tipoorigen,origen,operacion,valor,valorant,estado,"
StrSql = StrSql & " fechavigencia,solicitante,usuario,fechasol,pedcamdesext"
StrSql = StrSql & " FROM pedidocambio"
StrSql = StrSql & " WHERE pedcambnro = " & PedCambNro
OpenRecordset StrSql, rs_PedCambio
If rs_PedCambio.EOF Then
    Flog.writeline Espacios(Tabulador * 1) & "No se encontro el pedido de Cambio " & PedCambNro
    Exit Sub
Else
    If CInt(rs_PedCambio!estado) <> 2 Then
        Flog.writeline Espacios(Tabulador * 1) & "El pedido de Cambio " & PedCambNro & " no se encuentra en estado Autorizado."
        Exit Sub
    End If
    
    Select Case CInt(rs_PedCambio!tipoalcance)
        Case 1:
            'Individual
            Flog.writeline Espacios(Tabulador * 1) & "Alcance Individual"
            
            StrSql = "SELECT empleado.empleg, empleado.ternro, empleado.terape, empleado.ternom"
            StrSql = StrSql & " FROM empleado"
            StrSql = StrSql & " WHERE empleado.ternro = " & rs_PedCambio!codalcance
            
        Case 2:
            'Estructura
            Flog.writeline Espacios(Tabulador * 1) & "Alcance Estructura"
            
            StrSql = "SELECT estructura.estrnro, estructura.estrdabr, tipoestructura.tedabr"
            StrSql = StrSql & " FROM estructura"
            StrSql = StrSql & " INNER JOIN tipoestructura ON tipoestructura.tenro = estructura.tenro"
            StrSql = StrSql & " WHERE estructura.estrnro = " & rs_PedCambio!codalcance
            OpenRecordset StrSql, rs_Empleados
            
            If rs_Empleados.EOF Then
                Flog.writeline Espacios(Tabulador * 2) & "No se encontro la Estructura del alcance"
                rs_Empleados.Close
                Exit Sub
            Else
                Flog.writeline Espacios(Tabulador * 2) & "Buscando Empleados de " & rs_Empleados!tedabr & " " & rs_Empleados!estrdabr & " a la fecha " & rs_PedCambio!fechavigencia
                rs_Empleados.Close
                StrSql = "SELECT empleado.empleg, empleado.ternro, empleado.terape, empleado.ternom"
                StrSql = StrSql & " FROM his_estructura"
                StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = his_estructura.ternro"
                StrSql = StrSql & " AND empleado.empest = -1"
                StrSql = StrSql & " WHERE his_estructura.estrnro = " & rs_PedCambio!codalcance
                StrSql = StrSql & " AND his_estructura.htetdesde <= " & ConvFecha(rs_PedCambio!fechavigencia)
                StrSql = StrSql & " AND ((his_estructura.htethasta IS NULL) OR (his_estructura.htethasta >= " & ConvFecha(rs_PedCambio!fechavigencia) & "))"
                StrSql = StrSql & " ORDER BY empleado.empleg"
            End If
            
        Case 3:
            'Global
            Flog.writeline Espacios(Tabulador * 1) & "Alcance Global"
            
            StrSql = "SELECT empleado.empleg, empleado.ternro, empleado.terape, empleado.ternom"
            StrSql = StrSql & " FROM empleado"
            StrSql = StrSql & " WHERE empleado.empest = -1"
            StrSql = StrSql & " ORDER BY empleado.empleg"
    
    End Select
    
    OpenRecordset StrSql, rs_Empleados
    Progreso = 0
    CEmpleadosAProc = rs_Empleados.RecordCount
    If CEmpleadosAProc = 0 Then
       Flog.writeline Espacios(Tabulador * 1) & "No se encontraron empleados que satisfacen el alcance."
       Exit Sub
       CEmpleadosAProc = 1
    End If
    IncPorc = (99 / CEmpleadosAProc)
    
    If CInt(rs_PedCambio!tipocambio) = 1 Then
        'CAMBIO DE NOVEDAD
        Flog.writeline Espacios(Tabulador * 0) & "Cambio de Novedad"
        
        StrSql = "SELECT conccod, concabr FROM concepto WHERE concnro = " & rs_PedCambio!tipoorigen
        OpenRecordset StrSql, rs_Aux
        If Not rs_Aux.EOF Then
            Flog.writeline Espacios(Tabulador * 1) & "Concepto: " & rs_Aux!Conccod & " " & rs_Aux!concabr
        Else
            Flog.writeline Espacios(Tabulador * 1) & "No se encontro el concepto " & rs_PedCambio!tipoorigen
            Exit Sub
        End If
        rs_Aux.Close
        
        StrSql = "SELECT tpanro, tpadabr FROM tipopar WHERE tpanro = " & rs_PedCambio!Origen
        OpenRecordset StrSql, rs_Aux
        If Not rs_Aux.EOF Then
            Flog.writeline Espacios(Tabulador * 1) & "Parametro: " & rs_Aux!tpanro & " " & rs_Aux!tpadabr
        Else
            Flog.writeline Espacios(Tabulador * 1) & "No se encontro el Parametro " & rs_PedCambio!Origen
            Exit Sub
        End If
        rs_Aux.Close
        
        Flog.writeline Espacios(Tabulador * 1) & "Fecha de Vigencia: " & rs_PedCambio!fechavigencia
        Select Case CInt(rs_PedCambio!operacion)
            Case 1:
                Flog.writeline Espacios(Tabulador * 1) & "Operacion: Monto Fijo"
            Case 2:
                Flog.writeline Espacios(Tabulador * 1) & "Operacion: Aumento"
            Case 3:
                Flog.writeline Espacios(Tabulador * 1) & "Operacion: Porcentaje"
        End Select
        Flog.writeline Espacios(Tabulador * 1) & "Monto: " & rs_PedCambio!Valor
        
    Else
        'CAMBIO DE ESTRUCTURA
        Flog.writeline Espacios(Tabulador * 0) & "Cambio de Estructura"
        
        StrSql = "SELECT tenro, tedabr FROM tipoestructura WHERE tenro = " & rs_PedCambio!tipoorigen
        OpenRecordset StrSql, rs_Aux
        If Not rs_Aux.EOF Then
            Flog.writeline Espacios(Tabulador * 1) & "Tipo Estructura: " & rs_Aux!tenro & " " & rs_Aux!tedabr
        Else
            Flog.writeline Espacios(Tabulador * 1) & "No se encontro el Tipo Estructura " & rs_PedCambio!tipoorigen
            Exit Sub
        End If
        rs_Aux.Close
        
        StrSql = "SELECT estrnro, estrdabr FROM estructura WHERE estrnro = " & CLng(rs_PedCambio!Valor)
        OpenRecordset StrSql, rs_Aux
        If Not rs_Aux.EOF Then
            Flog.writeline Espacios(Tabulador * 1) & "Estructura Nueva: " & rs_Aux!estrnro & " " & rs_Aux!estrdabr
        Else
            Flog.writeline Espacios(Tabulador * 1) & "No se encontro la Nueva Estructura " & rs_PedCambio!Valor
            Exit Sub
        End If
        rs_Aux.Close
        
        Flog.writeline Espacios(Tabulador * 1) & "Fecha de Vigencia: " & rs_PedCambio!fechavigencia
        
    End If
    
End If

Flog.writeline
Flog.writeline Espacios(Tabulador * 0) & "------------------------------------------------------------------------"
Flog.writeline Espacios(Tabulador * 0) & "Comienza el procesamiento de Registros"
Flog.writeline Espacios(Tabulador * 0) & "------------------------------------------------------------------------"

Do While Not rs_Empleados.EOF
        
    Flog.writeline Espacios(Tabulador * 1) & "Procesando Empleado " & rs_Empleados!empleg & " - " & rs_Empleados!terape & " " & rs_Empleados!ternom
    
    If CInt(rs_PedCambio!tipocambio) = 1 Then 'NOVEDAD ---------------------------------------------------------------------------------
        
        montoNov = 0
        
        'Busco Novedad sin vigencia
        StrSql = "SELECT nenro, nevalor"
        StrSql = StrSql & " FROM novemp"
        StrSql = StrSql & " WHERE nevigencia <> -1"
        StrSql = StrSql & " AND tpanro = " & rs_PedCambio!Origen
        StrSql = StrSql & " AND empleado = " & rs_Empleados!ternro
        StrSql = StrSql & " AND concnro = " & rs_PedCambio!tipoorigen
        OpenRecordset StrSql, rs_Aux
        If Not rs_Aux.EOF Then
            
            If CInt(rs_PedCambio!operacion) <> 1 Then
                montoNov = rs_Aux!nevalor
            End If
            
            Flog.writeline Espacios(Tabulador * 2) & "Cerrando la novedad SIN Vigencia " & rs_Aux!nenro & " al dia anterior de la fecha vigencia."
            StrSql = "UPDATE novemp SET"
            StrSql = StrSql & " nevigencia = -1"
            StrSql = StrSql & " ,nehasta = " & ConvFecha(DateAdd("d", -1, rs_PedCambio!fechavigencia))
            StrSql = StrSql & " WHERE nenro = " & rs_Aux!nenro
            objConn.Execute StrSql, , adExecuteNoRecords
        
        Else
        
            'No encontro novedad sin vigencia, 'Busco si tiene una novedad con vigencia superior a la vigencia
            StrSql = "SELECT nenro, nevalor, nedesde"
            StrSql = StrSql & " FROM novemp"
            StrSql = StrSql & " WHERE nevigencia = -1"
            StrSql = StrSql & " AND tpanro = " & rs_PedCambio!Origen
            StrSql = StrSql & " AND empleado = " & rs_Empleados!ternro
            StrSql = StrSql & " AND concnro = " & rs_PedCambio!tipoorigen
            StrSql = StrSql & " AND("
            StrSql = StrSql & " (nedesde > " & ConvFecha(DateAdd("d", -1, rs_PedCambio!fechavigencia))
            StrSql = StrSql & " ))"
            OpenRecordset StrSql, rs_Aux
            If Not rs_Aux.EOF Then
                Flog.writeline Espacios(Tabulador * 2) & "No se puede realizar el cambio porque existe superposicion con la novedad " & rs_Aux!nenro & " " & rs_Aux!nedesde
                GoTo SgtEmpl
            End If
            rs_Aux.Close
            
            'Verifico si hay una novedad con vigencia para cerrar
            StrSql = "SELECT nenro, nevalor"
            StrSql = StrSql & " FROM novemp"
            StrSql = StrSql & " WHERE nevigencia = -1"
            StrSql = StrSql & " AND tpanro = " & rs_PedCambio!Origen
            StrSql = StrSql & " AND empleado = " & rs_Empleados!ternro
            StrSql = StrSql & " AND concnro = " & rs_PedCambio!tipoorigen
            StrSql = StrSql & " AND ((nedesde < " & ConvFecha(rs_PedCambio!fechavigencia) & ") AND ((nehasta is NULL) or (nehasta >= " & ConvFecha(rs_PedCambio!fechavigencia) & ")))"
            OpenRecordset StrSql, rs_Aux
            If Not rs_Aux.EOF Then
            
                If CInt(rs_PedCambio!operacion) <> 1 Then
                    montoNov = rs_Aux!nevalor
                End If
                
                Flog.writeline Espacios(Tabulador * 2) & "Cerrando novedad " & rs_Aux!nenro & " al dia anterior de la fecha vigencia."
                StrSql = "UPDATE novemp SET"
                StrSql = StrSql & " nehasta = " & ConvFecha(DateAdd("d", -1, rs_PedCambio!fechavigencia))
                StrSql = StrSql & " WHERE nenro = " & rs_Aux!nenro
                objConn.Execute StrSql, , adExecuteNoRecords
                
            End If
        
        End If
        
        Select Case CInt(rs_PedCambio!operacion)
            Case 1:
                montoNovInsert = rs_PedCambio!Valor
            Case 2:
                'Aumento
                If montoNov = 0 Then
                    Flog.writeline Espacios(Tabulador * 2) & "La operacion es Aumento y no se encontro Novedad para la base del calculo"
                    GoTo SgtEmpl
                Else
                    montoNovInsert = montoNov + rs_PedCambio!Valor
                    Flog.writeline Espacios(Tabulador * 2) & rs_PedCambio!Valor & " Aumento de " & montoNov & " = " & montoNovInsert
                End If
            Case 3:
                'Porcentaje
                If montoNov = 0 Then
                    Flog.writeline Espacios(Tabulador * 2) & "La operacion es Porcentaje y no se encontro Novedad para la base del calculo"
                    GoTo SgtEmpl
                Else
                    montoNovInsert = ((montoNov * rs_PedCambio!Valor) / 100) + montoNov
                    Flog.writeline Espacios(Tabulador * 2) & rs_PedCambio!Valor & " % de aumento de " & montoNov & " = " & montoNovInsert
                End If
        End Select
        
        'Inserto la novedad de Cambio
        StrSql = "INSERT INTO novemp"
        StrSql = StrSql & "(concnro,tpanro,empleado,"
        StrSql = StrSql & "nevalor,nevigencia,nedesde)"
        StrSql = StrSql & "VALUES"
        StrSql = StrSql & "( " & rs_PedCambio!tipoorigen
        StrSql = StrSql & ", " & rs_PedCambio!Origen
        StrSql = StrSql & ", " & rs_Empleados!ternro
        StrSql = StrSql & ", " & montoNovInsert
        StrSql = StrSql & ", -1"
        StrSql = StrSql & ", " & ConvFecha(rs_PedCambio!fechavigencia)
        StrSql = StrSql & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        Flog.writeline Espacios(Tabulador * 2) & "Novedad insertada"
        
        
    Else 'ESTRUCTURA -------------------------------------------------------------------------------------------------------------------
    
        'Busco si tiene una estructura del mismo tipo superior a la vigencia
        StrSql = "SELECT his_estructura.estrnro, estructura.estrdabr, his_estructura.htetdesde, his_estructura.htethasta"
        StrSql = StrSql & " FROM his_estructura"
        StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro"
        StrSql = StrSql & " WHERE his_estructura.ternro = " & rs_Empleados!ternro
        StrSql = StrSql & " AND his_estructura.tenro = " & rs_PedCambio!tipoorigen
        StrSql = StrSql & " AND("
        StrSql = StrSql & " (his_estructura.htetdesde > " & ConvFecha(DateAdd("d", -1, rs_PedCambio!fechavigencia))
        StrSql = StrSql & " ))"
        OpenRecordset StrSql, rs_Aux
        If Not rs_Aux.EOF Then
            Flog.writeline Espacios(Tabulador * 2) & "No se puede realizar el cambio porque existe superposicion con la estructura " & rs_Aux!estrdabr & " " & rs_Aux!htetdesde
            GoTo SgtEmpl
        End If
        rs_Aux.Close
        
        'Verifico si hay una estructura abierta para cerrarla
        StrSql = "SELECT his_estructura.estrnro, his_estructura.htetdesde"
        StrSql = StrSql & " FROM his_estructura"
        StrSql = StrSql & " WHERE his_estructura.ternro = " & rs_Empleados!ternro
        StrSql = StrSql & " AND his_estructura.tenro = " & rs_PedCambio!tipoorigen
        StrSql = StrSql & " AND ((his_estructura.htetdesde < " & ConvFecha(rs_PedCambio!fechavigencia) & ") AND ((his_estructura.htethasta is NULL) or (his_estructura.htethasta >= " & ConvFecha(rs_PedCambio!fechavigencia) & ")))"
        OpenRecordset StrSql, rs_Aux
        If Not rs_Aux.EOF Then
            Flog.writeline Espacios(Tabulador * 2) & "Cerrando estructura al dia anterior de la fecha vigencia."
            
            StrSql = "UPDATE his_estructura"
            StrSql = StrSql & " SET  htethasta = " & ConvFecha(DateAdd("d", -1, rs_PedCambio!fechavigencia))
            StrSql = StrSql & " WHERE estrnro = " & rs_Aux!estrnro
            StrSql = StrSql & " AND htetdesde = " & ConvFecha(rs_Aux!htetdesde)
            StrSql = StrSql & " AND tenro = " & rs_PedCambio!tipoorigen
            StrSql = StrSql & " AND ternro = " & rs_Empleados!ternro
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        rs_Aux.Close
        
        'Inserto la estructura del Cambio
        StrSql = "INSERT INTO his_estructura"
        StrSql = StrSql & " (tenro,ternro,estrnro,htetdesde)"
        StrSql = StrSql & " VALUES"
        StrSql = StrSql & "(" & rs_PedCambio!tipoorigen
        StrSql = StrSql & "," & rs_Empleados!ternro
        StrSql = StrSql & "," & rs_PedCambio!Valor
        StrSql = StrSql & "," & ConvFecha(rs_PedCambio!fechavigencia)
        StrSql = StrSql & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        Flog.writeline Espacios(Tabulador * 2) & "Estructura insertada"
        
    End If
    
    Flog.writeline
    
    '---------------------------------------------------------------------------------
    'Actualizo el progreso
    '---------------------------------------------------------------------------------
SgtEmpl:
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
             "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    rs_Empleados.MoveNext
    
Loop

If rs_Empleados.State = adStateOpen Then rs_Empleados.Close
If rs_PedCambio.State = adStateOpen Then rs_PedCambio.Close
If rs_Aux.State = adStateOpen Then rs_Aux.Close


Set rs_Empleados = Nothing
Set rs_PedCambio = Nothing
Set rs_Aux = Nothing


Exit Sub

CE:
    Flog.writeline "=================================================================="
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="
    
    
    MyBeginTrans
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
             "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    MyCommitTrans
    
    HuboError = True

End Sub
