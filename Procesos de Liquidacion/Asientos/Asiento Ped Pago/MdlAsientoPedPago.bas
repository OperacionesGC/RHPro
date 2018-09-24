Attribute VB_Name = "MdlAsientoPedPago"
Option Explicit

'Const Version = 1.01
'Const FechaVersion = "21/02/2008"   'Martin Ferraro - Version Inicial

'Const Version = 1.02
'Const FechaVersion = "03/08/2009"   'Manuel Lopez - Encriptacion de string connection. Se perdio la version 1.02 creada en el año 2009

'Const Version = 1.03
'Const FechaVersion = "02/03/2015"   'Dimatz Rafael - CAS 13764 - Se estandarizo para que desencripte el string de conexion

Const Version = 1.04
Const FechaVersion = "20/04/2015"   'Dimatz Rafael - CAS 13764 - Se corrigio el NroVol para que obtenga el primer parametro

'--------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------


Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial del Generador de Asientos Contables de pedidos de pago.
' Autor      : Martin Ferraro
' Fecha      : 21/02/2008
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim objconnMain As New ADODB.Connection
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
            EncriptStrconexion = CBool(ArrParametros(2))
            c_seed = ArrParametros(2)
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
    
    Nombre_Arch = PathFLog & "Asiento_Contable Pedido Pago" & "-" & NroProcesoBatch & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaVersion
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "-------------------------------------------------"
    Flog.writeline
    Flog.writeline

    TiempoInicialProceso = GetTickCount
    
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
    
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "'"
    StrSql = StrSql & " , bprcfecinicioej = " & ConvFecha(Date)
    StrSql = StrSql & " , bprcestado = 'Procesando'"
    StrSql = StrSql & " , bprcpid = " & PID
    StrSql = StrSql & " , bprctiempo = 0 "
    StrSql = StrSql & " , bprcprogreso = 0 "
    StrSql = StrSql & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 215 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    If Not rs_batch_proceso.EOF Then
        bprcparam = rs_batch_proceso!bprcparam
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        Call GenerarAsiento(NroProcesoBatch, bprcparam)
    Else
        Flog.writeline "NO SE ENCONTRO EL PROCESO " & NroProcesoBatch
    End If
    
    TiempoFinalProceso = GetTickCount
    If Not HuboError Then
        StrSql = "UPDATE batch_proceso SET bprctiempo = " & (TiempoFinalProceso - TiempoInicialProceso) & ", bprcprogreso = 100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
    Else
        StrSql = "UPDATE batch_proceso SET bprctiempo = " & (TiempoFinalProceso - TiempoInicialProceso) & ", bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProcesoBatch
    End If
    objconnProgreso.Execute StrSql, , adExecuteNoRecords


Fin:
    Flog.Close
    If objConn.State = adStateOpen Then objConn.Close
    If objconnProgreso.State = adStateOpen Then objconnProgreso.Close
Exit Sub

ME_Main:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 0) & " Error : " & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline
End Sub


Public Sub GenerarAsiento(ByVal bpronro As Long, ByVal Parametros As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Programa que se ejecuta para generar Asiento Contable
'              Configurado en el tipo de proceso batch
' Autor      : Martin Ferraro
' Fecha      : 21/02/2008
' --------------------------------------------------------------------------------------------
Dim rs_ProcVol As New ADODB.Recordset
Dim rs_Proc_V_modasi As New ADODB.Recordset
Dim rs_modLinea As New ADODB.Recordset
Dim rs_Pedidos As New ADODB.Recordset
Dim rs_Pagos As New ADODB.Recordset
Dim rs_Aux As New ADODB.Recordset
Dim rs_Estructura As New ADODB.Recordset
Dim rs_Consulta As New ADODB.Recordset

Dim NroVol
Dim NroAsi As Long
Dim NroLin As Long
Dim cuenta As String
Dim cuentaRes As String
Dim Monto As Double
Dim nroCheque As String
Dim lineaNivTernro1 As Long
Dim lineaNivTernro2 As Long
Dim lineaNivTernro3 As Long
Dim Estrnro1 As Long
Dim Estrnro2 As Long
Dim Estrnro3 As Long
Dim vol_Fec_Asiento As Date
Dim Legajo As Long
Dim cantidadPagos As Long
Dim cantidadLineas As Long
Dim NVol

On Error GoTo ME_GenerarAsiento

' Levanto cada parametro por separado, el separador de parametros es "."
Flog.writeline Espacios(Tabulador * 0) & "Inicio del proceso de volcado."
If Not IsNull(Parametros) Then
        NVol = Split(Parametros, ".")
        NroVol = NVol(0)
End If

Flog.writeline Espacios(Tabulador * 0) & "Parametros: "
Flog.writeline Espacios(Tabulador * 0) & "            Numero de Proceso de Volcado = " & NroVol
Flog.writeline
Flog.writeline


Flog.writeline Espacios(Tabulador * 0) & "Buscando el proceso de volcado."
'Buscando el proceso de volcado
StrSql = "SELECT * FROM proc_vol WHERE proc_vol.vol_cod = " & NroVol
OpenRecordset StrSql, rs_ProcVol

If rs_ProcVol.EOF Then
    Flog.writeline Espacios(Tabulador * 0) & "ERROR. Proceso de Volcado no encontrado."
    Exit Sub
End If


Flog.writeline Espacios(Tabulador * 0) & "Buscando modelo proceso de volcado."
'Buscando el modelos asociados al proceso
StrSql = "SELECT * FROM proc_v_modasi WHERE proc_v_modasi.vol_cod = " & NroVol
OpenRecordset StrSql, rs_Proc_V_modasi

If rs_Proc_V_modasi.EOF Then
    Flog.writeline Espacios(Tabulador * 0) & "ERROR. El proceso de volcado no tiene modelo asociado."
    Exit Sub
End If


Flog.writeline Espacios(Tabulador * 0) & "Buscando las lineas del modelo."
'Buscando las lineas del modelo
StrSql = "SELECT * From mod_linea WHERE masinro = " & rs_Proc_V_modasi!asi_cod
StrSql = StrSql & "ORDER BY linaorden"
OpenRecordset StrSql, rs_modLinea
If rs_modLinea.EOF Then
    Flog.writeline Espacios(Tabulador * 0) & "ERROR. El modelo del proceso de volcado no tiene lineas asociadas."
    Exit Sub
End If

'Buscando Pedidos de pago del proceso
Flog.writeline Espacios(Tabulador * 0) & "Buscando Pedidos de Pago a porcesar del proceso de volcado."
StrSql = "SELECT pedidopago.* FROM proc_vol_pl"
StrSql = StrSql & " INNER JOIN pedidopago ON pedidopago.ppagnro = proc_vol_pl.pronro"
StrSql = StrSql & " WHERE proc_vol_pl.vol_cod = " & NroVol
OpenRecordset StrSql, rs_Pedidos
If rs_Pedidos.EOF Then
    Flog.writeline Espacios(Tabulador * 0) & "ERROR. El proceso de volcado no tiene Pedidos de Pago asociados."
    Exit Sub
End If

'Contando la cantidad de pagos de los pedidos
cantidadPagos = 0
Do While Not rs_Pedidos.EOF
    StrSql = "SELECT count(ternro) cant"
    StrSql = StrSql & " From pago"
    StrSql = StrSql & " Where ppagnro = " & rs_Pedidos!ppagnro
    OpenRecordset StrSql, rs_Pagos
    If Not rs_Pagos.EOF Then
        If Not IsNull(rs_Pagos!cant) Then
            cantidadPagos = cantidadPagos + rs_Pagos!cant
        End If
    End If
    rs_Pagos.Close
    
    rs_Pedidos.MoveNext
Loop



'seteo las variables de progreso
Progreso = 0
cantidadLineas = rs_modLinea.RecordCount
IncPorc = 95 / (cantidadLineas * cantidadPagos)


'variable de fecha de asiento
vol_Fec_Asiento = rs_ProcVol!vol_Fec_Asiento


Flog.writeline Espacios(Tabulador * 0) & "Cantidad de lineas del modelo del proceso de volcado = " & cantidadLineas
Flog.writeline Espacios(Tabulador * 0) & "Cantidad de Pagos a procesar del proceso de volcado = " & cantidadPagos
Flog.writeline


'Por cada Pedido de pago
rs_Pedidos.MoveFirst
Do While Not rs_Pedidos.EOF
    
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "----------------------------------------------------"
    Flog.writeline Espacios(Tabulador * 0) & "Analizando Pedido de Pago " & rs_Pedidos!ppagnro & " - " & rs_Pedidos!ppagdesabr
    Flog.writeline Espacios(Tabulador * 0) & "----------------------------------------------------"
    
    '-----------------------------------------------------------------------------
    'Inicializaciones
    '-----------------------------------------------------------------------------
    cuentaRes = ""
    cuenta = ""
    
    '-----------------------------------------------------------------------------
    '-----------------------------------------------------------------------------
    'Resuelvo el Haber del pedido de pago
    '-----------------------------------------------------------------------------
    '-----------------------------------------------------------------------------
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "RESOLVIENDO HABER:"
    
    '-----------------------------------------------------------------------------
    'Monto del pedido
    '-----------------------------------------------------------------------------
    If EsNulo(rs_Pedidos!ppagimporte) Then
        Flog.writeline Espacios(Tabulador * 1) & "Monto del pedido nulo. No se crea cuenta Haber."
        GoTo ResolverPagos
    Else
        If rs_Pedidos!ppagimporte = 0 Then
            Flog.writeline Espacios(Tabulador * 1) & "Monto del pedido 0. No se crea cuenta Haber."
            GoTo ResolverPagos
        Else
            Flog.writeline Espacios(Tabulador * 1) & "Monto del Pedido = " & rs_Pedidos!ppagimporte
            Monto = rs_Pedidos!ppagimporte
        End If
    End If
    
    
    '-----------------------------------------------------------------------------
    'Busco la linea del haber en la cuenta bancaria del pedido
    '-----------------------------------------------------------------------------
    StrSql = "SELECT linacuenta "
    StrSql = StrSql & " FROM ctabancaria"
    StrSql = StrSql & " WHERE ctabancaria.cbnro = " & rs_Pedidos!cbnro
    OpenRecordset StrSql, rs_Aux
    If rs_Aux.EOF Then
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontro la cuenta bancaria del Pedido de Pago. No se crea la cuenta Haber del Pedido de Pago."
        GoTo ResolverPagos
    Else
        If EsNulo(rs_Aux!linacuenta) Then
            Flog.writeline Espacios(Tabulador * 1) & "ERROR. Linea contable Nula de la cuenta bancaria del Pedido de Pago. No se crea la cuenta Haber del Pedido de Pago."
            GoTo ResolverPagos
        Else
            cuenta = Trim(rs_Aux!linacuenta)
            If Len(cuenta) = 0 Then
                Flog.writeline Espacios(Tabulador * 1) & "ERROR. Linea contable Nula de la cuenta bancaria del Pedido de Pago. No se crea la cuenta Haber del Pedido de Pago."
                GoTo ResolverPagos
            Else
                Flog.writeline Espacios(Tabulador * 1) & "Cuenta Haber " & cuenta
            End If
        End If
    End If
    rs_Aux.Close
    
    
    '-----------------------------------------------------------------------------
    'Armo la linea teniendo en cuenta cheques
    '-----------------------------------------------------------------------------
    If EsNulo(rs_Pedidos!nroCheque) Then
        nroCheque = ""
    Else
        nroCheque = rs_Pedidos!nroCheque
    End If
    
    cuentaRes = cuenta
    Call ArmarCuentaHaber(cuentaRes, nroCheque)
    Flog.writeline Espacios(Tabulador * 1) & "ARMADO DE CUENTA: " & cuenta & " ----------> " & cuentaRes
    
    
    '-----------------------------------------------------------------------------
    'Guardo la linea
    '-----------------------------------------------------------------------------
    Flog.writeline Espacios(Tabulador * 1) & "GUARDANDO CUENTA HABER " & cuentaRes & " MONTO = " & Monto
    Call GuardarLineaAsi(NroVol, rs_Proc_V_modasi!asi_cod, 0, "Linea Haber PedPago " & rs_Pedidos!ppagdesabr, 0, cuentaRes, Monto)
    
    
    '-----------------------------------------------------------------------------
    '-----------------------------------------------------------------------------
    'Resuelvo el Debe con los pagos del pedido
    '-----------------------------------------------------------------------------
    '-----------------------------------------------------------------------------
ResolverPagos:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "RESOLVIENDO DEBE:"
    
    
    'Por cada lineas del modelo analizo cada pago
    rs_modLinea.MoveFirst
    Do While Not rs_modLinea.EOF
        
        'Procesando linea x del modelo
        cuenta = rs_modLinea!linacuenta
        Flog.writeline Espacios(Tabulador * 1) & "Procesando linea " & cuenta & " " & rs_modLinea!linadesc
        Flog.writeline Espacios(Tabulador * 1) & "----------------------------------------------------"
         
        '-----------------------------------------------------------------------
        'Configuracion de las estrcturas de la linea
        '-----------------------------------------------------------------------
        lineaNivTernro1 = IIf(Not EsNulo(rs_modLinea!lineaNivTernro1), rs_modLinea!lineaNivTernro1, 0)
        lineaNivTernro2 = IIf(Not EsNulo(rs_modLinea!lineaNivTernro2), rs_modLinea!lineaNivTernro2, 0)
        lineaNivTernro3 = IIf(Not EsNulo(rs_modLinea!lineaNivTernro3), rs_modLinea!lineaNivTernro3, 0)
        Flog.writeline Espacios(Tabulador * 1) & "Nivel 1 de estructura de la linea = " & lineaNivTernro1
        Flog.writeline Espacios(Tabulador * 1) & "Nivel 2 de estructura de la linea = " & lineaNivTernro2
        Flog.writeline Espacios(Tabulador * 1) & "Nivel 3 de estructura de la linea = " & lineaNivTernro3
        Flog.writeline
        
        '-----------------------------------------------------------------------
        'Busco los pagos del pedido
        '-----------------------------------------------------------------------
        StrSql = "SELECT *"
        StrSql = StrSql & " From pago"
        StrSql = StrSql & " Where ppagnro = " & rs_Pedidos!ppagnro
        OpenRecordset StrSql, rs_Pagos
        Do While Not rs_Pagos.EOF
            
            cuentaRes = ""
            Estrnro1 = 0
            Estrnro2 = 0
            Estrnro3 = 0
            Monto = 0
            Legajo = 0
            
            StrSql = "SELECT ternro, empleg, terape, ternom"
            StrSql = StrSql & " From empleado"
            StrSql = StrSql & " Where ternro = " & rs_Pagos!Ternro
            OpenRecordset StrSql, rs_Aux
            If Not rs_Aux Then
                Legajo = rs_Aux!empleg
            End If
            
            Flog.writeline Espacios(Tabulador * 2) & "Resolviendo Pago del empleado " & Legajo & " monto " & rs_Pagos!pagomonto
            
            If EsNulo(rs_Pagos!pagomonto) Then
                Flog.writeline Espacios(Tabulador * 3) & "Monto del pedido nulo. No se crea cuenta Haber."
                GoTo SgtPago
            Else
                If rs_Pagos!pagomonto = 0 Then
                    Flog.writeline Espacios(Tabulador * 3) & "Monto del pedido 0. No se crea cuenta Haber."
                    GoTo SgtPago
                Else
                    'Flog.writeline Espacios(Tabulador * 3) & "Monto del Pedido = " & rs_Pedidos!ppagimporte
                    Monto = rs_Pagos!pagomonto
                End If
            End If
            
            '-------------------------------------------------------------------------------------------
            'Control de las estructuras del empleado se correspondan con los filtros de la linea Nivel 1
            '-------------------------------------------------------------------------------------------
            If lineaNivTernro1 <> 0 Then
                StrSql = " SELECT estrnro FROM his_estructura "
                StrSql = StrSql & " WHERE ternro = " & rs_Pagos!Ternro & " AND "
                StrSql = StrSql & " tenro =" & lineaNivTernro1 & " AND "
                StrSql = StrSql & " (htetdesde <= " & ConvFecha(vol_Fec_Asiento) & ") AND "
                StrSql = StrSql & " ((" & ConvFecha(vol_Fec_Asiento) & " <= htethasta) or (htethasta is null))"
                OpenRecordset StrSql, rs_Estructura
                If Not rs_Estructura.EOF Then
                    Estrnro1 = rs_Estructura!Estrnro
                    'Verifico que si tiene filtro en la linea y si coincide
                    StrSql = "SELECT estrnro FROM mod_lin_filtro "
                    StrSql = StrSql & " WHERE masinro = " & rs_modLinea!masinro
                    StrSql = StrSql & " AND linaorden = " & rs_modLinea!LinaOrden
                    StrSql = StrSql & " AND tenro = " & lineaNivTernro1
                    OpenRecordset StrSql, rs_Aux
                    If Not rs_Aux.EOF Then
                        'Existe filtro para la linea con ese tipo de estuctura, controlo si existe la estructura
                        StrSql = "SELECT estrnro FROM mod_lin_filtro "
                        StrSql = StrSql & " WHERE masinro = " & rs_modLinea!masinro
                        StrSql = StrSql & " AND linaorden = " & rs_modLinea!LinaOrden
                        StrSql = StrSql & " AND tenro = " & lineaNivTernro1
                        StrSql = StrSql & " AND estrnro = " & Estrnro1
                        OpenRecordset StrSql, rs_Consulta
                        If rs_Consulta.EOF Then
                            'La linea no supera el filtro
                            Flog.writeline Espacios(Tabulador * 3) & "ERROR. El empleado NO supera el filtro del PRIMER nivel de estructura de la linea a la fecha " & vol_Fec_Asiento
                            GoTo SgtPago
                        End If
                        rs_Consulta.Close
                    End If
                    rs_Aux.Close
                Else
                    Flog.writeline Espacios(Tabulador * 3) & "ERROR. El empleado NO pertenece al PRIMER nivel de estructura de la linea a la fecha " & vol_Fec_Asiento
                    GoTo SgtPago
                End If
                rs_Estructura.Close
            End If
            
            
            '-------------------------------------------------------------------------------------------
            'Control de las estructuras del empleado se correspondan con los filtros de la linea Nivel 2
            '-------------------------------------------------------------------------------------------
            If lineaNivTernro2 <> 0 Then
                StrSql = " SELECT estrnro FROM his_estructura "
                StrSql = StrSql & " WHERE ternro = " & rs_Pagos!Ternro & " AND "
                StrSql = StrSql & " tenro =" & lineaNivTernro2 & " AND "
                StrSql = StrSql & " (htetdesde <= " & ConvFecha(vol_Fec_Asiento) & ") AND "
                StrSql = StrSql & " ((" & ConvFecha(vol_Fec_Asiento) & " <= htethasta) or (htethasta is null))"
                OpenRecordset StrSql, rs_Estructura
                If Not rs_Estructura.EOF Then
                    Estrnro2 = rs_Estructura!Estrnro
                    'Verifico que si tiene filtro en la linea y si coincide
                    StrSql = "SELECT estrnro FROM mod_lin_filtro "
                    StrSql = StrSql & " WHERE masinro = " & rs_modLinea!masinro
                    StrSql = StrSql & " AND linaorden = " & rs_modLinea!LinaOrden
                    StrSql = StrSql & " AND tenro = " & lineaNivTernro2
                    OpenRecordset StrSql, rs_Aux
                    If Not rs_Aux.EOF Then
                        'Existe filtro para la linea con ese tipo de estuctura, controlo si existe la estructura
                        StrSql = "SELECT estrnro FROM mod_lin_filtro "
                        StrSql = StrSql & " WHERE masinro = " & rs_modLinea!masinro
                        StrSql = StrSql & " AND linaorden = " & rs_modLinea!LinaOrden
                        StrSql = StrSql & " AND tenro = " & lineaNivTernro2
                        StrSql = StrSql & " AND estrnro = " & Estrnro2
                        OpenRecordset StrSql, rs_Consulta
                        If rs_Consulta.EOF Then
                            Flog.writeline Espacios(Tabulador * 3) & "ERROR. El empleado NO supera el filtro del SEGUNDO nivel de estructura de la linea a la fecha " & vol_Fec_Asiento
                            GoTo SgtPago
                        End If
                        rs_Consulta.Close
                    End If
                    rs_Aux.Close
                Else
                    Flog.writeline Espacios(Tabulador * 3) & "ERROR. El empleado NO pertenece al SEGUNDO nivel de estructura de la linea a la fecha " & vol_Fec_Asiento
                    GoTo SgtPago
                End If
                rs_Estructura.Close
            End If
            
            
            '-------------------------------------------------------------------------------------------
            'Control de las estructuras del empleado se correspondan con los filtros de la linea Nivel 3
            '-------------------------------------------------------------------------------------------
            If lineaNivTernro3 <> 0 Then
                StrSql = " SELECT estrnro FROM his_estructura "
                StrSql = StrSql & " WHERE ternro = " & rs_Pagos!Ternro & " AND "
                StrSql = StrSql & " tenro =" & lineaNivTernro3 & " AND "
                StrSql = StrSql & " (htetdesde <= " & ConvFecha(vol_Fec_Asiento) & ") AND "
                StrSql = StrSql & " ((" & ConvFecha(vol_Fec_Asiento) & " <= htethasta) or (htethasta is null))"
                OpenRecordset StrSql, rs_Estructura
                If Not rs_Estructura.EOF Then
                    Estrnro3 = rs_Estructura!Estrnro
                    'Verifico que si tiene filtro en la linea y si coincide
                    StrSql = "SELECT estrnro FROM mod_lin_filtro "
                    StrSql = StrSql & " WHERE masinro = " & rs_modLinea!masinro
                    StrSql = StrSql & " AND linaorden = " & rs_modLinea!LinaOrden
                    StrSql = StrSql & " AND tenro = " & lineaNivTernro3
                    OpenRecordset StrSql, rs_Aux
                    If Not rs_Aux.EOF Then
                        'Existe filtro para la linea con ese tipo de estuctura, controlo si existe la estructura
                        StrSql = "SELECT estrnro FROM mod_lin_filtro "
                        StrSql = StrSql & " WHERE masinro = " & rs_modLinea!masinro
                        StrSql = StrSql & " AND linaorden = " & rs_modLinea!LinaOrden
                        StrSql = StrSql & " AND tenro = " & lineaNivTernro3
                        StrSql = StrSql & " AND estrnro = " & Estrnro3
                        OpenRecordset StrSql, rs_Consulta
                        If rs_Consulta.EOF Then
                            Flog.writeline Espacios(Tabulador * 3) & "ERROR. El empleado NO supera el filtro del TERCER nivel de estructura de la linea a la fecha " & vol_Fec_Asiento
                            GoTo SgtPago
                        End If
                        rs_Consulta.Close
                    End If
                    rs_Aux.Close
                Else
                    Flog.writeline Espacios(Tabulador * 3) & "ERROR. El empleado NO pertenece al TERCER nivel de estructura de la linea a la fecha " & vol_Fec_Asiento
                    GoTo SgtPago
                End If
                rs_Estructura.Close
            End If
            
            
            '-------------------------------------------------------------------------------------------
            'Armo la cuenta debe
            '-------------------------------------------------------------------------------------------
            cuentaRes = cuenta
            Call ArmarCuentaDebe(cuentaRes, rs_Pagos!Ternro, Legajo, lineaNivTernro1, lineaNivTernro2, lineaNivTernro3, Estrnro1, Estrnro2, Estrnro3)
            
            
            '-----------------------------------------------------------------------------
            'Guardo la linea
            '-----------------------------------------------------------------------------
            Flog.writeline Espacios(Tabulador * 3) & "GUARDANDO CUENTA DEBE " & cuentaRes & " MONTO = " & Monto
            Call GuardarLineaAsi(NroVol, rs_modLinea!masinro, rs_modLinea!LinaOrden, "Pago", -1, cuentaRes, Monto)
            
SgtPago:    rs_Pagos.MoveNext

            'Actualizar el progreso
            TiempoFinalProceso = GetTickCount
            Progreso = Progreso + IncPorc
            StrSql = "UPDATE batch_proceso SET bprctiempo = " & (TiempoFinalProceso - TiempoInicialProceso) & ", bprcprogreso = " & Progreso & " WHERE bpronro = " & NroProcesoBatch
            objconnProgreso.Execute StrSql, , adExecuteNoRecords

        Loop
        rs_Pagos.Close
        
        rs_modLinea.MoveNext
        Flog.writeline
    Loop
    
    rs_Pedidos.MoveNext
Loop
rs_Pedidos.Close

'-----------------------------------------------------------------------------
'Creo el asiento
'-----------------------------------------------------------------------------
Call CrearAsiento(NroVol, rs_Proc_V_modasi!asi_cod)


'Cuento la cantidad de lineas generadas
StrSql = "SELECT count(*) Lineas FROM linea_asi "
StrSql = StrSql & " WHERE vol_cod =" & NroVol
If rs_Aux.State = adStateOpen Then rs_Aux.Close
OpenRecordset StrSql, rs_Aux
If Not rs_Aux.EOF Then
    NroLin = rs_Aux!Lineas
End If


'Cuento la cantidad de asientos generados
StrSql = "SELECT COUNT(DISTINCT masinro) Asientos FROM linea_asi "
StrSql = StrSql & " WHERE vol_cod =" & NroVol
If rs_Aux.State = adStateOpen Then rs_Aux.Close
OpenRecordset StrSql, rs_Aux
If Not rs_Aux.EOF Then
    NroAsi = rs_Aux!Asientos
End If

StrSql = "UPDATE proc_vol SET vol_reg_cab = " & NroAsi & _
             ", vol_reg_det =" & NroLin & _
             " WHERE proc_vol.vol_cod =" & NroVol
objConn.Execute StrSql, , adExecuteNoRecords


If rs_ProcVol.State = adStateOpen Then rs_ProcVol.Close
If rs_Proc_V_modasi.State = adStateOpen Then rs_Proc_V_modasi.Close
If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
If rs_Consulta.State = adStateOpen Then rs_Consulta.Close
If rs_modLinea.State = adStateOpen Then rs_modLinea.Close
If rs_Pedidos.State = adStateOpen Then rs_Pedidos.Close
If rs_Pagos.State = adStateOpen Then rs_Pagos.Close
If rs_Aux.State = adStateOpen Then rs_Aux.Close

Set rs_ProcVol = Nothing
Set rs_Proc_V_modasi = Nothing
Set rs_Estructura = Nothing
Set rs_Consulta = Nothing
Set rs_modLinea = Nothing
Set rs_Pedidos = Nothing
Set rs_Pagos = Nothing
Set rs_Aux = Nothing


Exit Sub

'Manejador de Errores del procedimiento
ME_GenerarAsiento:
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Procedimiento: GenerarAsiento"
    Flog.writeline "Ultimo SQL Ejecutado: " & StrSql
End Sub



Public Sub GuardarLineaAsi(ByVal vol_cod As Long, ByVal masinro As Long, ByVal Linea As Long, ByVal desclinea As String, ByVal dh As Integer, ByVal cuenta As String, ByVal Monto As Double)
' --------------------------------------------------------------------------------------------
' Descripcion: Inserta las cuentas en la base de datos en linea_asi
' Autor      : Martin Ferraro
' Fecha      : 28/02/2008
' --------------------------------------------------------------------------------------------
Dim rs_Linea_asi As New ADODB.Recordset
    
On Error GoTo ME_GuardarLineaAsi
            
    'Miro si la linea ya esta en la base para el proceso y modelo
    StrSql = "SELECT * FROM linea_asi " & _
             " WHERE linea_asi.vol_cod = " & vol_cod & _
             " AND linea_asi.cuenta  = '" & Mid(cuenta, 1, 50) & "'" & _
             " AND linea_asi.masinro = " & masinro & _
             " AND linea_asi.dh = " & dh
    OpenRecordset StrSql, rs_Linea_asi
    
    If rs_Linea_asi.EOF Then
    
        'No existe una linea con esa cuenta, entonces la inserto
        StrSql = "INSERT INTO linea_asi (cuenta,vol_cod,masinro,linea,desclinea,dh,monto)" & _
                 " VALUES ('" & Mid(cuenta, 1, 50) & _
                 "'," & vol_cod & _
                 "," & masinro & _
                 "," & Linea & _
                 ",'" & Mid(desclinea, 1, 60) & _
                 "'," & dh & _
                 "," & Round(Monto, 4) & _
                 ")"
        objConn.Execute StrSql, , adExecuteNoRecords
    Else
    
        'la linea existe, debo modificar el monto
        StrSql = "UPDATE linea_asi SET monto = monto + " & Monto & _
                 " WHERE linea_asi.vol_cod =" & vol_cod & _
                 " AND linea_asi.cuenta  ='" & Mid(cuenta, 1, 50) & "'" & _
                 " AND linea_asi.masinro =" & masinro
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
    
    rs_Linea_asi.Close
    
'cierro todo
If rs_Linea_asi.State = adStateOpen Then rs_Linea_asi.Close
Set rs_Linea_asi = Nothing

Exit Sub
'Manejador de Errores del procedimiento
ME_GuardarLineaAsi:
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Procedimiento: GuardarLineaAsi"
    Flog.writeline "Ultimo SQL Ejecutado: " & StrSql


End Sub




Public Sub ArmarCuentaDebe(ByRef NroCuenta As String, ByVal Ternro As Long, ByVal Legajo As String, ByVal Masinivternro1 As Long, ByVal Masinivternro2 As Long, ByVal Masinivternro3 As Long, ByVal Estructura1 As Long, ByVal Estructura2 As Long, ByVal Estructura3 As Long)
' --------------------------------------------------------------------------------------------
' Descripcion: Arma la cuenta Debe de acuerdo a la configuracion de la misma
' Autor      : Martin Ferraro
' Fecha      : 04/08/2006
' --------------------------------------------------------------------------------------------
Dim Aux_Cuenta As String
Dim Aux_Legajo As String

Dim DescEstructura1 As String
Dim DescEstructura2 As String
Dim DescEstructura3 As String

Dim I As Integer
Dim ch As String
Dim CantL As Integer
Dim CantE As Integer
Dim TipoE As String
Dim TipoE_Actual As String
Dim EsEstructura As Boolean
Dim Termino As Boolean

Dim PosE1 As Integer
Dim PosE2 As Integer
Dim PosE3 As Integer

Dim EsDocumento As Boolean
Dim CantD As Long
Dim TipoD As String
Dim TipoD_Actual As String
Dim DescDocumento As String

Dim rs_Estructura As New ADODB.Recordset
Dim rs_Documento As New ADODB.Recordset

On Error GoTo ME_ArmarCuentaDebe

Aux_Cuenta = NroCuenta
Aux_Legajo = Legajo

'Descripcion de estructura 1
If Masinivternro1 = 0 Then
    'Modelo sin apertura
    DescEstructura1 = "00000000000000000000"
Else
    'Modelo con apertura, busco la descripcion de la estructura
    StrSql = " SELECT estrcodext, estrnro, tenro FROM estructura " & _
             " WHERE estrnro = " & Estructura1
    OpenRecordset StrSql, rs_Estructura
    
    If Not rs_Estructura.EOF Then
            DescEstructura1 = IIf(IsNull(rs_Estructura!estrcodext), "00000000000000000000", rs_Estructura!estrcodext & "00000000000000000000")
    Else
        DescEstructura1 = "00000000000000000000"
    End If
    rs_Estructura.Close
End If
DescEstructura1 = Left(DescEstructura1, 20)

'Descripcion de estructura 2
If Masinivternro2 = 0 Then
    'Modelo sin apertura
    DescEstructura2 = "00000000000000000000"
Else
    'Modelo con apertura, busco la descripcion de la estructura
    StrSql = " SELECT estrcodext, estrnro, tenro FROM estructura " & _
             " WHERE estrnro = " & Estructura2
    OpenRecordset StrSql, rs_Estructura
    
    If Not rs_Estructura.EOF Then
            DescEstructura2 = IIf(IsNull(rs_Estructura!estrcodext), "00000000000000000000", rs_Estructura!estrcodext & "00000000000000000000")
    Else
        DescEstructura2 = "00000000000000000000"
    End If
    rs_Estructura.Close
End If
DescEstructura2 = Left(DescEstructura2, 20)

'Descripcion de estructura 3
If Masinivternro3 = 0 Then
    'Modelo sin apertura
    DescEstructura3 = "00000000000000000000"
Else
    'Modelo con apertura, busco la descripcion de la estructura
    StrSql = " SELECT estrcodext, estrnro, tenro FROM estructura " & _
             " WHERE estrnro = " & Estructura3
    OpenRecordset StrSql, rs_Estructura
    
    If Not rs_Estructura.EOF Then
            DescEstructura3 = IIf(IsNull(rs_Estructura!estrcodext), "00000000000000000000", rs_Estructura!estrcodext & "00000000000000000000")
    Else
        DescEstructura3 = "00000000000000000000"
    End If
    rs_Estructura.Close
End If
DescEstructura3 = Left(DescEstructura3, 20)


PosE1 = 1
PosE2 = 1
PosE3 = 1


'Voy recorriendo de Izquierda a Derecha el aux_cuenta y voy generando el NroCuenta
I = 1
NroCuenta = ""
CantL = 0
CantE = 0
CantD = 0
Termino = False

Do While Not (I > Len(Aux_Cuenta))
    ch = UCase(Mid(Aux_Cuenta, I, 1))

    Select Case ch
    Case "_", "-", ".", "*":
        NroCuenta = NroCuenta & ch
        I = I + 1
    Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9":
        NroCuenta = NroCuenta & ch
        I = I + 1
    Case "E": 'Estrcutura
        EsEstructura = True
        CantE = 1
        'leo el nro de la estructura
        I = I + 1
        ch = UCase(Mid(Aux_Cuenta, I, 1))
        TipoE = ch
        Termino = False
        
        Do While EsEstructura And Not Termino
            'leo el siguiente
            I = I + 1
            If Not (I > Len(Aux_Cuenta)) Then
                ch = UCase(Mid(Aux_Cuenta, I, 1))
            Else
                Termino = True
            End If
            
            If ch = "E" And Not Termino Then
                'leo lel nro de la estructura
                I = I + 1
                ch = UCase(Mid(Aux_Cuenta, I, 1))
                TipoE_Actual = ch
                
                Do While TipoE = TipoE_Actual And EsEstructura And Not Termino
                    CantE = CantE + 1
    
                    I = I + 1
                    If Not (I > Len(Aux_Cuenta)) Then
                        ch = UCase(Mid(Aux_Cuenta, I, 1))
                    Else
                        Termino = True
                    End If
                    
                    If ch = "E" Then
                        'leo el nro de la estructura
                        I = I + 1
                        ch = UCase(Mid(Aux_Cuenta, I, 1))
                        TipoE_Actual = ch
                    Else
                        Termino = True
                    End If
                Loop
                
            Else
                EsEstructura = False
            End If
            
            'reemplazo por la estructura correspondiente
            Select Case TipoE
            Case 1:
                'NroCuenta = NroCuenta & Right(Estructura1, CantE)
                NroCuenta = NroCuenta & Mid(DescEstructura1, PosE1, CantE)
                PosE1 = PosE1 + CantE
                If PosE1 >= 20 Then PosE1 = 1
            Case 2:
                'NroCuenta = NroCuenta & Right(Estructura2, CantE)
                NroCuenta = NroCuenta & Mid(DescEstructura2, PosE2, CantE)
                PosE2 = PosE2 + CantE
                If PosE2 >= 20 Then PosE2 = 1
            Case 3:
                'NroCuenta = NroCuenta & Right(Estructura3, CantE)
                NroCuenta = NroCuenta & Mid(DescEstructura3, PosE3, CantE)
                PosE3 = PosE3 + CantE
                If PosE3 >= 20 Then PosE3 = 1
            End Select
            
            TipoE = TipoE_Actual
            CantE = 1
        Loop
        
    Case "L": 'Legajo
        Termino = False
        CantL = 1
        I = I + 1
        If I <= Len(Aux_Cuenta) Then
            ch = UCase(Mid(Aux_Cuenta, I, 1))
        End If
        
        Do While ch = "L" And Not Termino
            CantL = CantL + 1
            I = I + 1
            If I <= Len(Aux_Cuenta) Then
                ch = UCase(Mid(Aux_Cuenta, I, 1))
            Else
                Termino = True
            End If
        Loop
        
        NroCuenta = NroCuenta & Right(Format(Aux_Legajo, "0000000000"), CantL)
     
  

    Case "D": 'DXX NUMERO DE TIPO DE DOC, DONDE XX ES EL TIPO DE DOC
        EsDocumento = True
        CantD = 1
        'leo el nro de documento
        I = I + 1
        ch = UCase(Mid(Aux_Cuenta, I, 2))
        'Avanzo otro porque lei 2
        I = I + 1
        TipoD = ch
        Termino = False

        Do While EsDocumento And Not Termino
            'leo el siguiente
            I = I + 1
            If Not (I > Len(Aux_Cuenta)) Then
                ch = UCase(Mid(Aux_Cuenta, I, 1))
            Else
                Termino = True
            End If

            If ch = "D" And Not Termino Then
                'leo lel nro de documento
                I = I + 1
                ch = UCase(Mid(Aux_Cuenta, I, 2))
                'Avanzo otro porque lei 2
                I = I + 1
                TipoD_Actual = ch

                Do While TipoD = TipoD_Actual And EsDocumento And Not Termino
                    CantD = CantD + 1

                    I = I + 1
                    If Not (I > Len(Aux_Cuenta)) Then
                        ch = UCase(Mid(Aux_Cuenta, I, 1))
                    Else
                        Termino = True
                    End If

                    If ch = "D" Then
                        'leo el nro de documento
                        I = I + 1
                        ch = UCase(Mid(Aux_Cuenta, I, 2))
                        'Avanzo otro porque lei 2
                        I = I + 1
                        TipoD_Actual = ch
                    Else
                        Termino = True
                    End If
                Loop

            Else
                EsDocumento = False
            End If
            
            

            'Busco el documento para reemplazar
            DescDocumento = "00000000000000000000"
            If IsNumeric(TipoD) Then
                StrSql = " SELECT nrodoc " & _
                         " From ter_doc " & _
                         " Where ter_doc.ternro = " & Ternro & " And ter_doc.tidnro = " & CLng(TipoD)
                OpenRecordset StrSql, rs_Documento
                If Not rs_Documento.EOF Then
                        DescDocumento = IIf(IsNull(rs_Documento!NroDoc), "00000000000000000000", rs_Documento!NroDoc & "00000000000000000000")
                Else
                    DescDocumento = "00000000000000000000"
                End If
                rs_Documento.Close
            End If
            DescDocumento = Left(DescDocumento, 20)
            
            'Reemplazo en la cuenta
            NroCuenta = NroCuenta & Mid(DescDocumento, 1, CantD)

            TipoD = TipoD_Actual
            CantD = 1
        
        Loop
    
     
    Case "a" To "z", "A" To "Z":
        NroCuenta = NroCuenta & ch
        I = I + 1
    Case Else:
        I = I + 1
    End Select
Loop


'cierro todo
If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
If rs_Documento.State = adStateOpen Then rs_Documento.Close
Set rs_Estructura = Nothing
Set rs_Documento = Nothing

Exit Sub
'Manejador de Errores del procedimiento
ME_ArmarCuentaDebe:
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Procedimiento: ArmarCuentaDebe"
    Flog.writeline "Ultimo SQL Ejecutado: " & StrSql

End Sub


Public Sub ArmarCuentaHaber(ByRef NroCuenta As String, ByVal cheque As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Arma la cuenta de acuerdo a la configuracion de la misma
' Autor      : Martin Ferraro
' Fecha      : 28/02/2008
' --------------------------------------------------------------------------------------------
Dim Aux_Cuenta As String
Dim Aux_Legajo As String

Dim DescEstructura1 As String
Dim DescEstructura2 As String
Dim DescEstructura3 As String

Dim I As Integer
Dim ch As String
Dim CantL As Integer
Dim CantE As Integer
Dim TipoE As String
Dim TipoE_Actual As String
Dim EsEstructura As Boolean
Dim Termino As Boolean

Dim PosE1 As Integer
Dim PosE2 As Integer
Dim PosE3 As Integer

Dim Desc_cheque As String
Dim CantH As Long


On Error GoTo ME_ArmarCuentaHaber

Aux_Cuenta = NroCuenta
Aux_Legajo = "00000000000000000000"

'Descripcion de estructura 1
DescEstructura1 = "00000000000000000000"
'Descripcion de estructura 2
DescEstructura2 = "00000000000000000000"
'Descripcion de estructura 3
DescEstructura3 = "00000000000000000000"
'Descripcion de cheque
'Desc_cheque = cheque
Desc_cheque = Left(cheque, 20)

PosE1 = 1
PosE2 = 1
PosE3 = 1


'Voy recorriendo de Izquierda a Derecha el aux_cuenta y voy generando el NroCuenta
I = 1
NroCuenta = ""
CantL = 0
CantH = 0
CantE = 0
Termino = False

Do While Not (I > Len(Aux_Cuenta))
    ch = UCase(Mid(Aux_Cuenta, I, 1))

    Select Case ch
    Case "_", "-", ".", "*":
        NroCuenta = NroCuenta & ch
        I = I + 1
    Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9":
        NroCuenta = NroCuenta & ch
        I = I + 1
    Case "E": 'Estrcutura
        EsEstructura = True
        CantE = 1
        'leo el nro de la estructura
        I = I + 1
        ch = UCase(Mid(Aux_Cuenta, I, 1))
        TipoE = ch
        Termino = False
        
        Do While EsEstructura And Not Termino
            'leo el siguiente
            I = I + 1
            If Not (I > Len(Aux_Cuenta)) Then
                ch = UCase(Mid(Aux_Cuenta, I, 1))
            Else
                Termino = True
            End If
            
            If ch = "E" And Not Termino Then
                'leo lel nro de la estructura
                I = I + 1
                ch = UCase(Mid(Aux_Cuenta, I, 1))
                TipoE_Actual = ch
                
                Do While TipoE = TipoE_Actual And EsEstructura And Not Termino
                    CantE = CantE + 1
    
                    I = I + 1
                    If Not (I > Len(Aux_Cuenta)) Then
                        ch = UCase(Mid(Aux_Cuenta, I, 1))
                    Else
                        Termino = True
                    End If
                    
                    If ch = "E" Then
                        'leo el nro de la estructura
                        I = I + 1
                        ch = UCase(Mid(Aux_Cuenta, I, 1))
                        TipoE_Actual = ch
                    Else
                        Termino = True
                    End If
                Loop
                
            Else
                EsEstructura = False
            End If
            
            'reemplazo por la estructura correspondiente
            Select Case TipoE
            Case 1:
                NroCuenta = NroCuenta & Mid(DescEstructura1, PosE1, CantE)
                PosE1 = PosE1 + CantE
                If PosE1 >= 20 Then PosE1 = 1
            Case 2:
                NroCuenta = NroCuenta & Mid(DescEstructura2, PosE2, CantE)
                PosE2 = PosE2 + CantE
                If PosE2 >= 20 Then PosE2 = 1
            Case 3:
                NroCuenta = NroCuenta & Mid(DescEstructura3, PosE3, CantE)
                PosE3 = PosE3 + CantE
                If PosE3 >= 20 Then PosE3 = 1
            End Select
            
            TipoE = TipoE_Actual
            CantE = 1
        Loop
        
    Case "L": 'Legajo
        Termino = False
        CantL = 1
        I = I + 1
        If I <= Len(Aux_Cuenta) Then
            ch = UCase(Mid(Aux_Cuenta, I, 1))
        End If
        
        Do While ch = "L" And Not Termino
            CantL = CantL + 1
            I = I + 1
            If I <= Len(Aux_Cuenta) Then
                ch = UCase(Mid(Aux_Cuenta, I, 1))
            Else
                Termino = True
            End If
        Loop
        
        NroCuenta = NroCuenta & Right(Format(Aux_Legajo, "0000000000"), CantL)
     
    Case "H": 'Cheque
        Termino = False
        CantH = 1
        I = I + 1
        If I <= Len(Aux_Cuenta) Then
            ch = UCase(Mid(Aux_Cuenta, I, 1))
        End If
        
        Do While ch = "H" And Not Termino
            CantH = CantH + 1
            I = I + 1
            If I <= Len(Aux_Cuenta) Then
                ch = UCase(Mid(Aux_Cuenta, I, 1))
            Else
                Termino = True
            End If
        Loop
        
        NroCuenta = NroCuenta & Right(Format(Desc_cheque, "00000000000000000000"), CantH)
     
     
    Case "a" To "z", "A" To "Z":
        NroCuenta = NroCuenta & ch
        I = I + 1
    Case Else:
        I = I + 1
    End Select
Loop


Exit Sub
'Manejador de Errores del procedimiento
ME_ArmarCuentaHaber:
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Procedimiento: ArmarCuentaHaber"
    Flog.writeline "Ultimo SQL Ejecutado: " & StrSql

End Sub

    

Public Sub CrearAsiento(ByVal vol_cod As Long, ByVal masinro As Long)

Dim rs_lineaAsi As New ADODB.Recordset

Dim montoDebe As Double
Dim montoHaber As Double
    
On Error GoTo ME_CrearAsiento
    
    montoDebe = 0
    montoHaber = 0
    
    
    '-------------------------------------------------------------------------------
    'Busco monto del debe
    '-------------------------------------------------------------------------------
    StrSql = " SELECT SUM(monto) debe FROM linea_asi " & _
             " WHERE linea_asi.masinro = " & masinro & _
             " AND linea_asi.vol_cod = " & vol_cod & _
             " AND linea_asi.dh = -1"
    OpenRecordset StrSql, rs_lineaAsi
    If Not rs_lineaAsi.EOF Then
        montoDebe = IIf(EsNulo(rs_lineaAsi!debe), 0, rs_lineaAsi!debe)
    End If
    rs_lineaAsi.Close
    
    
    '-------------------------------------------------------------------------------
    'Busco monto del Haber
    '-------------------------------------------------------------------------------
    StrSql = " SELECT SUM(monto) haber FROM linea_asi " & _
             " WHERE linea_asi.masinro = " & masinro & _
             " AND linea_asi.vol_cod = " & vol_cod & _
             " AND linea_asi.dh = 0"
    OpenRecordset StrSql, rs_lineaAsi
    If Not rs_lineaAsi.EOF Then
        montoHaber = IIf(EsNulo(rs_lineaAsi!haber), 0, rs_lineaAsi!haber)
    End If
    rs_lineaAsi.Close
    
    Flog.writeline
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "----------------------------------------------------"
    Flog.writeline Espacios(Tabulador * 0) & "MONTO DEBE = " & Round(montoDebe, 4) & " MONTO HABER = " & Round(montoHaber, 4)
    Flog.writeline Espacios(Tabulador * 0) & "----------------------------------------------------"
   
    '-------------------------------------------------------------------------------
    'Creo el asiento
    '-------------------------------------------------------------------------------
    StrSql = "SELECT * FROM asiento " & _
             " WHERE masinro = " & masinro & _
             " AND vol_cod = " & vol_cod
    OpenRecordset StrSql, rs_lineaAsi
    
    If rs_lineaAsi.EOF Then
        StrSql = "INSERT INTO asiento (masinro,asidebe,asihaber,vol_cod) " & _
                 " VALUES (" & masinro & _
                 "," & Round(montoDebe, 4) & _
                 "," & Round(montoHaber, 4) & _
                 "," & vol_cod & _
                 ")"
        objConn.Execute StrSql, , adExecuteNoRecords
    Else
        StrSql = "UPDATE asiento SET asidebe = " & Round(montoDebe, 4) & _
                 ",asihaber =" & Round(montoHaber, 4) & _
                 " WHERE masinro = " & masinro & _
                 " AND vol_cod =" & vol_cod
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
    rs_lineaAsi.Close
    
If rs_lineaAsi.State = adStateOpen Then rs_lineaAsi.Close
Set rs_lineaAsi = Nothing

Exit Sub
'Manejador de Errores del procedimiento
ME_CrearAsiento:
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Procedimiento: CrearAsiento"
    Flog.writeline "Ultimo SQL Ejecutado: " & StrSql


End Sub



