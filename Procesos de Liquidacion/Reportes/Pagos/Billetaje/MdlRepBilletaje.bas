Attribute VB_Name = "MdlRepBilletaje"
Option Explicit

Global IdUser As String
Global Fecha As Date
Global Hora As String

Global Fecha_Inicio_periodo As Date
Global Fecha_Fin_Periodo As Date
Global StrSql2 As String



Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial del Generador de Reportes.
' Autor      : FGZ
' Fecha      : 02/03/2004
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

'    strCmdLine = Command()
'    ArrParametros = Split(strCmdLine, " ", -1)
'    If UBound(ArrParametros) > 0 Then
'        If IsNumeric(ArrParametros(0)) Then
'            NroProcesoBatch = ArrParametros(0)
'            Etiqueta = ArrParametros(1)
'        Else
'            Exit Sub
'        End If
'    Else
'        If IsNumeric(strCmdLine) Then
'            NroProcesoBatch = strCmdLine
'        Else
'            Exit Sub
'        End If
'    End If
    
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

    'Abro la conexion
    Nombre_Arch = PathFLog & "Reporte_Billetaje" & "-" & NroProcesoBatch & ".log"
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    
    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    OpenConnection strconexion, objconnProgreso
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "PID = " & PID
    'Cambio el estado del proceso a Procesando y el PID
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 48 AND bpronro =" & NroProcesoBatch
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
    Flog.writeline "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    
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

Public Sub Suecam01(ByVal bpronro As Long, ByVal Nroliq As Long, ByVal Todos_Pro As Boolean, ByVal Proc_Aprob As Integer, ByVal Empresa As Long, ByVal Todos_Pedidos As Boolean, ByVal Nro_Pedido As Long)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del Reporte de Billetaje
' Autor      : FGZ
' Fecha      : 03/05/2004
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim Total_General As Single
Dim Acumulador_pago As Long
Dim T As Single
Dim Monto As Single
Dim Resto As Single
Dim entero As Integer
Dim Total_cobrado As Single
Dim OK As Boolean
Dim Continua As Boolean

Dim I As Integer
Dim Ultimo_Empleado As Long
Dim Ultima_Moneda As Long

Dim rs_Periodo As New ADODB.Recordset
Dim rs_Procesos As New ADODB.Recordset
Dim rs_Rep41 As New ADODB.Recordset
Dim rs_Billetes As New ADODB.Recordset
Dim rs_Confppago As New ADODB.Recordset

'cargo el periodo
StrSql = "SELECT * FROM periodo WHERE pliqnro = " & CStr(Nroliq)
OpenRecordset StrSql, rs_Periodo
If rs_Periodo.EOF Then
    Flog.writeline "No se encontró el Periodo"
    Exit Sub
End If

Fecha_Inicio_periodo = rs_Periodo!pliqdesde
Fecha_Fin_Periodo = rs_Periodo!pliqhasta

' Comienzo la transaccion
MyBeginTrans

'Depuracion del Temporario
StrSql = "DELETE FROM rep41 "
StrSql = StrSql & " WHERE pliqnro = " & Nroliq
If Not Todos_Pro Then
    StrSql = StrSql & " AND pronro = " & NroProc
Else
    StrSql = StrSql & " AND pronro = 0"
    StrSql = StrSql & " AND proaprob = " & CInt(Proc_Aprob)
End If
If Todos_Pedidos Then
    StrSql = StrSql & " AND todos_pedidos = -1"
Else
    StrSql = StrSql & " AND todos_pedidos = 0"
    StrSql = StrSql & " AND ppagnro = " & CInt(Nro_Pedido)
End If
StrSql = StrSql & " AND empresa = " & Empresa
objConn.Execute StrSql, , adExecuteNoRecords



'DETERMINAR EL ACUMULADOR DE PAGO
StrSql = "SELECT * FROM confppag "
OpenRecordset StrSql, rs_Confppago
If Not rs_Confppago.EOF Then
    Acumulador_pago = rs_Confppago!acuNro
End If

Total_General = 0
Total_cobrado = 0

'Busco los procesos a evaluar
StrSql = "SELECT * FROM  periodo "
StrSql = StrSql & " INNER JOIN proceso ON periodo.pliqnro = proceso.pliqnro "
StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
StrSql = StrSql & " INNER JOIN pago ON pago.pagorigen = cabliq.cliqnro "
StrSql = StrSql & " INNER JOIN pedidopago ON pago.ppagnro = pedidopago.ppagnro AND pedidopago.tppanro = 1 "
StrSql = StrSql & " INNER JOIN formapago ON pedidopago.fpagnro = formapago.fpagnro"
'StrSql = StrSql & " INNER JOIN his_estructura  empresa ON empresa.ternro = cabliq.empleado and empresa.tenro = 10 "
'StrSql = StrSql & " INNER JOIN estructura emp ON emp.estrnro = empresa.estrnro AND emp.empnro =" & Empresa
StrSql = StrSql & " INNER JOIN his_estructura  ON his_estructura.ternro = cabliq.empleado and his_estructura.tenro = 10 "
StrSql = StrSql & " INNER JOIN empresa ON empresa.estrnro = his_estructura.estrnro AND empresa.empnro =" & Empresa
StrSql = StrSql & " INNER JOIN empleado ON pago.ternro = empleado.ternro"
StrSql = StrSql & " INNER JOIN moneda ON moneda.monnro = formapago.monnro"
StrSql = StrSql & " WHERE (his_estructura.htetdesde <= " & ConvFecha(Fecha_Fin_Periodo) & ") AND "
StrSql = StrSql & " ((" & ConvFecha(Fecha_Fin_Periodo) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
StrSql = StrSql & " AND periodo.pliqnro =" & Nroliq
If Not Todos_Pedidos Then
    StrSql = StrSql & " AND pago.ppagnro =" & Nro_Pedido
End If
If Not Todos_Pro Then
    StrSql = StrSql & " AND proceso.pronro =" & NroProc
Else
    StrSql = StrSql & " AND proaprob = " & CInt(Proc_Aprob)
    StrSql = StrSql & " AND proceso.empnro = " & Empresa
End If
OpenRecordset StrSql, rs_Procesos

'seteo de las variables de progreso
Progreso = 0
CConceptosAProc = rs_Procesos.RecordCount
If CConceptosAProc = 0 Then
    CConceptosAProc = 1
End If
IncPorc = (100 / CConceptosAProc)

Ultimo_Empleado = -1
Ultima_Moneda = -1
Do While Not rs_Procesos.EOF

    If Ultima_Moneda <> rs_Procesos!monnro Then
        StrSql = "SELECT * FROM billete WHERE monnro =" & rs_Procesos!monnro
        StrSql = StrSql & " ORDER BY billvalor DESC "
        OpenRecordset StrSql, rs_Billetes
        
        If rs_Billetes.EOF Then
            Flog.writeline "Falta la configuraci¢n de los Billetes para la moneda: " & rs_Procesos!mondesabr
            OK = False
        Else
            OK = True
            Ultima_Moneda = rs_Procesos!monnro
        End If
    End If
    
    Total_cobrado = Total_cobrado + rs_Procesos!pagomonto
    
    If EsElUltimoEmpleado(rs_Procesos, Ultimo_Empleado) Then
        Total_General = Total_General + Total_cobrado
        
        If Not OK Then
            Total_cobrado = 0
        Else
            If Total_cobrado < 0 Then
                Flog.writeline "Existen montos negativos, no podrán ser tomados en cuenta para el Proceso. "
            Else
                Continua = True
                Do While Not rs_Billetes.EOF And Continua
                    Monto = Total_cobrado / rs_Billetes!billvalor
                    entero = Fix(Monto)
                
                    If entero <> 0 Then
                        Resto = Total_cobrado - (entero * rs_Billetes!billvalor)
                        
                        'Si no existe el rep41
                        StrSql = "SELECT * FROM rep41 "
                        StrSql = StrSql & " WHERE ternro = " & rs_Procesos!ternro
                        StrSql = StrSql & " AND bpronro = " & bpronro
                        StrSql = StrSql & " AND pliqnro = " & Nroliq
                        StrSql = StrSql & " AND empresa = " & Empresa
                        If Not Todos_Pro Then
                            StrSql = StrSql & " AND pronro = " & NroProc
                        Else
                            StrSql = StrSql & " AND pronro = 0"
                            StrSql = StrSql & " AND proaprob= " & CInt(Proc_Aprob)
                        End If
                        StrSql = StrSql & " AND monnro =" & rs_Procesos!monnro
                        StrSql = StrSql & " AND billcod =" & rs_Billetes!billvalor
                        OpenRecordset StrSql, rs_Rep41
                    
                        If rs_Rep41.EOF Then
                            'Inserto
                            StrSql = "INSERT INTO rep41 (bpronro,pliqnro,pronro,proaprob,empresa,iduser,fecha,hora,"
                            StrSql = StrSql & "ternro,monnro,billcod,billdes,billcan,"
                            StrSql = StrSql & "moncod,moncan,mondes "
                            StrSql = StrSql & ") VALUES ("
                            StrSql = StrSql & bpronro & ","
                            StrSql = StrSql & Nroliq & ","
                            If Not Todos_Pro Then
                                StrSql = StrSql & rs_Procesos!pronro & ","
                                StrSql = StrSql & rs_Procesos!proaprob & ","
                            Else
                                StrSql = StrSql & "0" & ","
                                StrSql = StrSql & CInt(Proc_Aprob) & ","
                            End If
                            StrSql = StrSql & Empresa & ","
                            StrSql = StrSql & "'" & IdUser & "',"
                            StrSql = StrSql & ConvFecha(Fecha) & ","
                            StrSql = StrSql & "'" & Hora & "',"
                            StrSql = StrSql & rs_Procesos!ternro & ","
                            StrSql = StrSql & rs_Billetes!monnro & ","
                            StrSql = StrSql & rs_Billetes!billvalor & ","
                            StrSql = StrSql & "'" & rs_Billetes!billdesabr & "',"
                            StrSql = StrSql & "0,"
                            StrSql = StrSql & "0,"
                            StrSql = StrSql & "0,"
                            StrSql = StrSql & "'" & rs_Procesos!mondesabr & "'"
                            'StrSql = StrSql & rs_Procesos! & ","
                            'StrSql = StrSql & rs_Procesos!
                            StrSql = StrSql & ")"
                            objConn.Execute StrSql, , adExecuteNoRecords
                        End If
                        
                        'Actualizo
                        StrSql = "UPDATE rep41 SET billcan = billcan + " & entero
                        StrSql = StrSql & " WHERE ternro = " & rs_Procesos!ternro
                        StrSql = StrSql & " AND bpronro = " & bpronro
                        StrSql = StrSql & " AND pliqnro = " & Nroliq
                        StrSql = StrSql & " AND empresa = " & Empresa
                        If Not Todos_Pro Then
                            StrSql = StrSql & " AND pronro = " & NroProc
                        Else
                            StrSql = StrSql & " AND pronro = 0"
                            StrSql = StrSql & " AND proaprob= " & CInt(Proc_Aprob)
                        End If
                        StrSql = StrSql & " AND monnro =" & rs_Procesos!monnro
                        StrSql = StrSql & " AND billcod =" & rs_Billetes!billvalor
                        objConn.Execute StrSql, , adExecuteNoRecords
                        
                        
                        If Resto = 0 Then
                            Total_cobrado = 0
                            Continua = False
                        Else
                            Total_cobrado = Resto
                        End If
                    End If 'If entero <> 0 Then
                    
                    'siguiente billete
                    rs_Billetes.MoveNext
                Loop
                    'Actualizo el progreso del Proceso
                    Progreso = Progreso + IncPorc
                    TiempoAcumulado = GetTickCount
                    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
                             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
                             "' WHERE bpronro = " & NroProcesoBatch
                    objconnProgreso.Execute StrSql, , adExecuteNoRecords
                            
                    Ultimo_Empleado = rs_Procesos!ternro
            End If 'If Total_cobrado < 0 Then
        End If 'If Not Ok Then
    End If 'If EsElUltimoEmpleado(rs_Procesos, Ultimo_Empleado) Then
    
    'Siguiente proceso
    rs_Procesos.MoveNext
Loop

'Fin de la transaccion
MyCommitTrans


If rs_Procesos.State = adStateOpen Then rs_Procesos.Close
If rs_Periodo.State = adStateOpen Then rs_Periodo.Close
If rs_Rep41.State = adStateOpen Then rs_Rep41.Close
If rs_Billetes.State = adStateOpen Then rs_Billetes.Close
If rs_Confppago.State = adStateOpen Then rs_Confppago.Close

Set rs_Procesos = Nothing
Set rs_Billetes = Nothing
Set rs_Rep41 = Nothing
Set rs_Periodo = Nothing
Set rs_Confppago = Nothing

Exit Sub
CE:
    HuboError = True
    MyRollbackTrans

End Sub



Public Sub LevantarParamteros(ByVal bpronro As Long, ByVal parametros As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento para levantar los parametros pasados en batch_proceso en bprcparam
' Autor      : FGZ
' Fecha      :
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim pos1 As Integer
Dim pos2 As Integer

Dim Nroliq As Long
Dim Todos_Pro As Boolean
Dim Proc_Aprob As Integer
Dim Empresa As Long
Dim Todos_Pedidos As Boolean
Dim NroPedido As Long

' Levanto cada parametro por separado, el separador de parametros es "."
If Not IsNull(parametros) Then
    If Len(parametros) >= 1 Then
        pos1 = 1
        pos2 = InStr(pos1, parametros, ".") - 1
        Nroliq = CLng(Mid(parametros, pos1, pos2))
    
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, ".") - 1
        Proc_Aprob = Mid(parametros, pos1, pos2 - pos1 + 1)
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, ".") - 1
        Todos_Pro = CBool(Mid(parametros, pos1, pos2 - pos1 + 1))
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, ".") - 1
        NroProc = CLng(Mid(parametros, pos1, pos2 - pos1 + 1))
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, ".") - 1
        Todos_Pedidos = CBool(Mid(parametros, pos1, pos2 - pos1 + 1))
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, ".") - 1
        NroPedido = CLng(Mid(parametros, pos1, pos2 - pos1 + 1))
        
        pos1 = pos2 + 2
        pos2 = Len(parametros)
        Empresa = Mid(parametros, pos1, pos2 - pos1 + 1)
    End If
End If


Call Suecam01(bpronro, Nroliq, Todos_Pro, Proc_Aprob, Empresa, Todos_Pedidos, NroPedido)

End Sub


Public Function EsElUltimoEmpleado(ByVal rs As ADODB.Recordset, ByVal Anterior As Long) As Boolean
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento para saber si es el ultimo empleado de la secuencia
' Autor      : FGZ
' Fecha      :
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
    
    rs.MoveNext
    If rs.EOF Then
        EsElUltimoEmpleado = True
    Else
        If rs!Empleado <> Anterior Then
            EsElUltimoEmpleado = True
        Else
            EsElUltimoEmpleado = False
        End If
    End If
    rs.MovePrevious
End Function

