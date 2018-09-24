Attribute VB_Name = "MdlTickets"
Option Explicit


'Const Version = 2.01   'Version Inicial
'Const FechaVersion = "23/05/2006"
'Modificacion :
'   FGZ - cuando no haybia empleados para procesar en batch_proceso daba un error.

'Const Version = 2.02
'Const FechaVersion = "07/08/2006"
'Modificacion :
'   Mariano Capriz - Se agrego el PRONRO para ke muestre a ke proceso pertenece

'Const Version = 2.03
'Const FechaVersion = "08/08/2006"
'Modificacion :
'   Mariano Capriz - Se agrego para que contemple el caso de ke el usuario seleccione Todos los Procesos

'Const Version = 2.04
'Const FechaVersion = "19/09/2006"
'Modificacion : Se agregaron lineas al flog. y se modifico un IF

'Const Version = 2.05
'Const FechaVersion = "14/06/2007"
'Modificacion : Diego Rosso - Se generaron las versiones 2.04 y 2.05 para nivelar el fuente.

Const Version = 2.06
Const FechaVersion = "31/07/2009" 'Martin Ferraro - Encriptacion de string connection


'----------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------

Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial Tickets.
' Autor      : FGZ
' Fecha      : 28/06/2004
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

    Nombre_Arch = PathFLog & "Tickets" & "-" & NroProcesoBatch & ".log"
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaVersion
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline
    
    'Abro la conexion
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
    
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 54 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        bprcparam = rs_batch_proceso!bprcparam
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        Call Generar_Tickets(NroProcesoBatch, bprcparam)
    End If
    
    TiempoFinalProceso = GetTickCount
    Flog.writeline "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    
    If Not HuboError Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcprogreso = 100, bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcprogreso = 100, bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
    End If
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Flog.Close
    objConn.Close
    objconnProgreso.Close
End Sub


Public Sub Generar_Tickets(ByVal bpronro As Long, ByVal Parametros As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion de Tickets
' Autor      : FGZ
' Fecha      : 28/06/2004
' Ult. Mod   : Mariano Capriz - 07/08/2006 - Se agrego la descripcion del Proceso
'              Mariano Capriz - 08/08/2006 - Se agrego codigo para ke contemple el caso en ke el usuario seleccione Todos los Procesos
' --------------------------------------------------------------------------------------------
Dim TikPedNro As Long
Dim Lista_Pro As String

Dim Separador As String
Dim Pliqnro As Long
Dim Todos_Procesos As Boolean
Dim Todos_Empleados As Boolean

Dim Monto As Single
Dim Cantidad As Single
Dim Primera_vez As Boolean
Dim EtikNro As Long
Dim I As Integer
Dim CMarcasAProc

Dim pos1 As Integer
Dim pos2 As Integer

Dim rs_TikPedido As New ADODB.Recordset
Dim rs_Emprtik As New ADODB.Recordset
Dim rs_Ticket As New ADODB.Recordset
Dim rs_Empleados As New ADODB.Recordset
Dim rs_Detliq As New ADODB.Recordset
Dim rs_EMP_Ticket As New ADODB.Recordset
Dim rs_TikValor As New ADODB.Recordset
Dim rs_Distrib As New ADODB.Recordset
'-------------------------------------------------
'MDC - 07/08/2006
Dim rs_TodosProcesos As New ADODB.Recordset
'-------------------------------------------------

On Error GoTo CE
    
' El formato del mismo es (tikpednro, Lista_procesos )
Separador = "."
' Levanto cada parametro por separado, el separador de parametros es "."
If Not IsNull(Parametros) Then
    If Len(Parametros) >= 1 Then
    
        Flog.writeline "Obteniendo los parametros"
        
        pos1 = 1
        pos2 = InStr(pos1, Parametros, Separador) - 1
        TikPedNro = CLng(Mid(Parametros, pos1, pos2))
        Flog.writeline "Nro de Pedido de Ticket: " & TikPedNro
    
        pos1 = pos2 + 2
        pos2 = InStr(pos1, Parametros, Separador) - 1
        Pliqnro = Mid(Parametros, pos1, pos2 - pos1 + 1)
        Flog.writeline "Periodo: " & Pliqnro
    
        pos1 = pos2 + 2
        pos2 = InStr(pos1, Parametros, Separador) - 1
        Todos_Procesos = CBool(Mid(Parametros, pos1, pos2 - pos1 + 1))
        Flog.writeline "Todos los procesos: " & Todos_Procesos
    
        pos1 = pos2 + 2
        pos2 = InStr(pos1, Parametros, Separador) - 1
        Lista_Pro = Mid(Parametros, pos1, pos2 - pos1 + 1)
        If Lista_Pro <> "" Then Flog.writeline "Lista de procesos: " & Lista_Pro
        ' esta lista tiene los nro de procesos separados por comas
        
        pos1 = pos2 + 2
        pos2 = Len(Parametros)
        Todos_Empleados = CBool(Mid(Parametros, pos1, pos2 - pos1 + 1))
        Flog.writeline "Todos los empleados: " & Todos_Empleados
        
    End If
End If


' -----------------------------------------------------------

'Comienzo la transaccion
MyBeginTrans

'MDC - 08/08/2006 -------------------------------------------------------------------------------
'Si se selecciono Todos los Procesos los busco y armo la lista
If Todos_Procesos = True Then
    StrSql = "select * from proceso where pliqnro=" & Pliqnro
    OpenRecordset StrSql, rs_TodosProcesos
    
    Lista_Pro = rs_TodosProcesos!pronro
    rs_TodosProcesos.MoveNext
    
    Do Until rs_TodosProcesos.EOF = True
        Lista_Pro = Lista_Pro & "," & rs_TodosProcesos!pronro
        rs_TodosProcesos.MoveNext
    Loop
    Flog.writeline "Lista de procesos: " & Lista_Pro
End If
'-----------------------------------------------------------------------------------------------
StrSql = "SELECT * FROM tikpedido "
StrSql = StrSql & " WHERE tikpednro =" & TikPedNro

Flog.writeline "Buscando el pedido de tickets"

OpenRecordset StrSql, rs_TikPedido

If Not rs_TikPedido.EOF Then

    Pliqnro = rs_TikPedido!Pliqnro
    
    If Not Todos_Empleados Then
        StrSql = "SELECT * FROM batch_empleado "
        StrSql = StrSql & " WHERE bpronro =" & NroProcesoBatch
    Else
        StrSql = "SELECT * FROM empleado ORDER BY empleg "
    End If
    
    'Flog.writeline "Buscando los empleados"
    
    OpenRecordset StrSql, rs_Empleados
    
    'seteo de las variables de progreso
    Progreso = 0
'    CConceptosAProc = rs_Ticket.RecordCount
'    If CConceptosAProc = 0 Then
'        CConceptosAProc = 1
'    End If
    CEmpleadosAProc = rs_Empleados.RecordCount
    If CEmpleadosAProc = 0 Then
       CEmpleadosAProc = 1
    End If
    Flog.writeline "Buscando los empleados, cantidad=" & CEmpleadosAProc
                
    StrSql = "SELECT * FROM emprtik "
    StrSql = StrSql & " WHERE ternro =" & rs_TikPedido!emprtik
    
    Flog.writeline "Buscando la empresa de tickets"
    
    OpenRecordset StrSql, rs_Emprtik
    
    'FGZ - 23/05/2006
    If rs_Empleados.EOF Then
        Flog.writeline "No hay empleados para procesar."
        GoTo Fin
    End If
    
    If Not rs_Emprtik.EOF Then
        
        StrSql = "SELECT * FROM ticket "
        StrSql = StrSql & " WHERE emprtik =" & rs_Emprtik!ternro
        
        OpenRecordset StrSql, rs_Ticket
        
        
        CMarcasAProc = rs_Ticket.RecordCount
        If CMarcasAProc = 0 Then
           CMarcasAProc = 1
        End If
        
        IncPorc = ((100 / CEmpleadosAProc) * (100 / CMarcasAProc)) / 100
        'IncPorc = (100 / CEmpleadosAProc)
        
        Do While Not rs_Ticket.EOF
        
            Flog.writeline "Buscando las marcas asociadas a la empresa: " & rs_Ticket!tikdesc
            
            Flog.writeline "Me posiciono en el primer empleado"
            'FGZ - 23/05/2006
            rs_Empleados.MoveFirst
            
            '-----------------------------------------------------------------------------------------------
            'MDC-07/08/2006
            '-----------------------------------------------------------------------------------------------
            Do While Not rs_Empleados.EOF
                Primera_vez = True
                Monto = 0
                Cantidad = 0
                'Busco los detliq de los conceptos asociados al ticket en el periodo del pedido
                ' para los procesos seleccionados para el empleado
                StrSql = "SELECT detliq.dlimonto Monto, detliq.dlicant Cantidad, proceso.pronro PROCESO FROM proceso "
                StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
                StrSql = StrSql & " INNER JOIN detliq ON cabliq.cliqnro = detliq.cliqnro "
                StrSql = StrSql & " INNER JOIN ticket_conc ON detliq.concnro = ticket_conc.concnro AND ticket_conc.tiknro =" & rs_Ticket!tiknro
                StrSql = StrSql & " WHERE proceso.pliqnro =" & Pliqnro
                StrSql = StrSql & " AND proceso.pronro IN (" & Lista_Pro & ")"
                StrSql = StrSql & " AND cabliq.empleado =" & rs_Empleados!ternro
                
                'Flog.writeline "StrSql= " & StrSql
                
                Flog.writeline "Buscando los conceptos de tickets para el empleado, ternro=" & rs_Empleados!ternro
                
                OpenRecordset StrSql, rs_Detliq
                'Flog.writeline "Pronro=" & ProNro
                
                Do While Not rs_Detliq.EOF
                    Monto = IIf(IsNull(rs_Detliq!Monto), 0, rs_Detliq!Monto)
                    Flog.writeline "Monto=" & Monto
                    
                    Cantidad = IIf(IsNull(rs_Detliq!Cantidad), 0, rs_Detliq!Cantidad)
                    
                    If Cantidad = 0 Then Cantidad = 1
                    Flog.writeline "Cantidad=" & Cantidad
                    
                    If Monto <> 0 Then
                        'Busco si existe el EMP_TICKET
                        StrSql = "SELECT * FROM emp_ticket WHERE " & _
                                 " tikpednro = " & TikPedNro & _
                                 " AND tiknro = " & rs_Ticket!tiknro & _
                                 " AND empleado = " & rs_Empleados!ternro
                                 
                        Flog.writeline "Controlo si el empleado ya tiene una ticket de esta marca"
                        
                        OpenRecordset StrSql, rs_EMP_Ticket
                        
                        If rs_EMP_Ticket.EOF Then
                            '-----------------------------------------------------------------------------
                            'MDC - 07/08/2006
                            StrSql = "INSERT INTO emp_ticket ("
                            StrSql = StrSql & "empleado,tiknro,tikpednro,etikfecha,etikmonto,etikcant,etikmanual,pronro"
                            StrSql = StrSql & ") VALUES (" & rs_Empleados!ternro
                            StrSql = StrSql & "," & rs_Ticket!tiknro
                            StrSql = StrSql & "," & TikPedNro
                            StrSql = StrSql & "," & ConvFecha(Date)
                            StrSql = StrSql & "," & Monto
                            StrSql = StrSql & "," & Cantidad
                            StrSql = StrSql & ",0"
                            StrSql = StrSql & "," & rs_Detliq!Proceso 'Proceso
                            StrSql = StrSql & ")"
                            
                            
                            Flog.writeline "Inserto un ticket para el empleado " & rs_Empleados!ternro
                            '------------------------------------------------------------------------------
                            
                            objConn.Execute StrSql, , adExecuteNoRecords
                            EtikNro = getLastIdentity(objConn, "emp_ticket")
                            If Primera_vez Then
                                'Borro la distribucion
                                StrSql = "DELETE emp_tikdist "
                                StrSql = StrSql & " WHERE etiknro = " & EtikNro ' ,tikvalnro,tiknro,etikdmonto,etikdmontouni,etikdcant"
                                'StrSql = StrSql & " AND tikvalnro = " & rs_TikValor!tikvalnro
                                StrSql = StrSql & " AND tiknro = " & rs_Ticket!tiknro
                                
                                Flog.writeline "Borro la distribucion de tickets"
                                
                                objConn.Execute StrSql, , adExecuteNoRecords
                                Primera_vez = False
                            End If
                        Else
                            If Not CBool(rs_EMP_Ticket!etikmanual) Then 'And IsNull(rs_EMP_Ticket!pronro) Then  12-09-2006
                                '-----------------------------------------------------------------------------
                                'MDC - 07/08/2006
                                StrSql = "UPDATE emp_ticket SET etikmonto = " & IIf(Not Primera_vez, "etikmonto + ", "") & Monto
                                StrSql = StrSql & " , etikfecha = " & ConvFecha(Date)
                                StrSql = StrSql & " , etikcant = " & Cantidad
                                StrSql = StrSql & ", pronro = " & rs_Detliq!Proceso 'Proceso
                                StrSql = StrSql & " WHERE tikpednro = " & TikPedNro
                                StrSql = StrSql & " AND tiknro = " & rs_Ticket!tiknro
                                StrSql = StrSql & " AND empleado = " & rs_Empleados!ternro
                                
                                Flog.writeline "Actualizo el ticket para el empleado " & rs_Empleados!ternro
                                '------------------------------------------------------------------------------
                                
                                objConn.Execute StrSql, , adExecuteNoRecords
                                EtikNro = rs_EMP_Ticket!EtikNro
                                If Primera_vez Then
                                    'Borro la distribucion
                                    StrSql = "DELETE emp_tikdist "
                                    StrSql = StrSql & " WHERE etiknro = " & EtikNro ' ,tikvalnro,tiknro,etikdmonto,etikdmontouni,etikdcant"
                                    StrSql = StrSql & " AND tiknro = " & rs_Ticket!tiknro
                                    
                                    Flog.writeline "Borro la distribucion de tickets " & rs_Ticket!tiknro
                                    
                                    objConn.Execute StrSql, , adExecuteNoRecords
                                    Primera_vez = False
                                End If
                            End If
                        End If
                        
                        'Inserto el detalle de los tickets
                        Valor = Monto / Cantidad
                        
                        StrSql = "SELECT * FROM tikvalor "
                        StrSql = StrSql & " INNER JOIN ticket_valor ON ticket_valor.tvalnro = tikvalor.tvalnro AND ticket_valor.tiknro = " & rs_Ticket!tiknro
                        StrSql = StrSql & " WHERE tvalmonto = " & Valor
                        
                        Flog.writeline "Busco si existe algun valor de ticket para el monto actual " & Valor
                        
                        OpenRecordset StrSql, rs_TikValor
                        
                        If Not rs_TikValor.EOF Then
                            If Cantidad <> 0 Then 'inserto los EMP_TIKDIST
                            
                                StrSql = "SELECT * FROM emp_tikdist "
                                StrSql = StrSql & " WHERE "
                                StrSql = StrSql & " etiknro = " & EtikNro
                                StrSql = StrSql & " AND tikvalnro = " & rs_TikValor!tikvalnro
                                StrSql = StrSql & " AND tiknro = " & rs_Ticket!tiknro
                                
                                Flog.writeline "Controlo si existe alguna distribucion para el valor actual"
                                 
                                OpenRecordset StrSql, rs_Distrib
                                
                                If rs_Distrib.EOF Then
                            
                                    StrSql = "INSERT INTO emp_tikdist ("
                                    StrSql = StrSql & "etiknro,tikvalnro,tiknro,etikdmonto,etikdmontouni,etikdcant"
                                    StrSql = StrSql & ") VALUES (" & EtikNro
                                    StrSql = StrSql & "," & rs_TikValor!tikvalnro
                                    StrSql = StrSql & "," & rs_Ticket!tiknro
                                    StrSql = StrSql & "," & Monto
                                    StrSql = StrSql & "," & Valor
                                    StrSql = StrSql & "," & Cantidad
                                    StrSql = StrSql & " )"
                                    
                                    Flog.writeline "Guardo la distribucion"
                                    
                                    objConn.Execute StrSql, , adExecuteNoRecords
                                
                                Else
                                
                                    StrSql = "UPDATE emp_tikdist SET "
                                    StrSql = StrSql & " etikdmonto = etikdmonto + " & Monto
                                    StrSql = StrSql & ",etikdcant = etikdcant + " & Cantidad
                                    StrSql = StrSql & " WHERE "
                                    StrSql = StrSql & " etiknro = " & EtikNro
                                    StrSql = StrSql & " AND tikvalnro = " & rs_TikValor!tikvalnro
                                    StrSql = StrSql & " AND tiknro = " & rs_Ticket!tiknro
                                    
                                    Flog.writeline "Actualizo la distribucion"
                                    
                                    objConn.Execute StrSql, , adExecuteNoRecords
                                
                                End If
                             End If
                        Else
                            Flog.writeline " No se encuentró ticket de valor " & Valor
                            Flog.writeline " Empleado abortado "
                            'Fuerzo el eof
                            rs_Detliq.MoveLast
                            rs_Detliq.MoveNext
                        End If
                        
                    End If
                    
                    If Not rs_Detliq.EOF Then
                       rs_Detliq.MoveNext
                    End If
                Loop
                
                'Actualizo el progreso
                Progreso = Progreso + IncPorc
                TiempoAcumulado = GetTickCount
                StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
                         ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
                         "' WHERE bpronro = " & NroProcesoBatch
                objconnProgreso.Execute StrSql, , adExecuteNoRecords
                
                rs_Empleados.MoveNext
            Loop
            rs_Ticket.MoveNext
        Loop
    Else
        Flog.writeline " No se encuentra la Empresa del ticket (EMPRTICK) " & rs_TikPedido!emprtik
    End If
Else
    Flog.writeline " No se encontró el pedido " & TikPedNro
End If

Fin:
'Fin de la transaccion
MyCommitTrans


'Cierro todo y libero
If rs_TikPedido.State = adStateOpen Then rs_TikPedido.Close
If rs_Emprtik.State = adStateOpen Then rs_Emprtik.Close
If rs_Ticket.State = adStateOpen Then rs_Ticket.Close
If rs_Empleados.State = adStateOpen Then rs_Empleados.Close
If rs_Detliq.State = adStateOpen Then rs_Detliq.Close
If rs_EMP_Ticket.State = adStateOpen Then rs_EMP_Ticket.Close
If rs_TikValor.State = adStateOpen Then rs_TikValor.Close

Set rs_EMP_Ticket = Nothing
Set rs_TikPedido = Nothing
Set rs_Emprtik = Nothing
Set rs_Ticket = Nothing
Set rs_Empleados = Nothing
Set rs_Detliq = Nothing
Set rs_TikValor = Nothing
Exit Sub

CE:
    MyRollbackTrans
    
    MyBeginTrans
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
             "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    MyCommitTrans
    
    HuboError = True
    Flog.writeline
    Flog.writeline "**********************************************************"
    Flog.writeline " Empleado abortado: "
    Flog.writeline " Error: " & Err.Description
    Flog.writeline "**********************************************************"
    Flog.writeline
End Sub


Public Sub Generar_Tickets_old(ByVal bpronro As Long, ByVal Parametros As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion de Tickets
' Autor      : FGZ
' Fecha      : 28/06/2004
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim TikPedNro As Long
Dim Lista_Pro As String
Dim Separador As String
Dim Pliqnro As Long
Dim Todos_Procesos As Boolean
Dim Todos_Empleados As Boolean

Dim Monto As Single
Dim Cantidad As Single

Dim pos1 As Integer
Dim pos2 As Integer

Dim rs_TikPedido As New ADODB.Recordset
Dim rs_Emprtik As New ADODB.Recordset
Dim rs_Ticket As New ADODB.Recordset
Dim rs_Empleados As New ADODB.Recordset
Dim rs_Detliq As New ADODB.Recordset
Dim rs_EMP_Ticket As New ADODB.Recordset

On Error GoTo CE
    
' El formato del mismo es (tikpednro, Lista_procesos )
Separador = "."
' Levanto cada parametro por separado, el separador de parametros es "."
If Not IsNull(Parametros) Then
    If Len(Parametros) >= 1 Then
    
        Flog.writeline "Obteniendo parametros"
        
        pos1 = 1
        pos2 = InStr(pos1, Parametros, Separador) - 1
        TikPedNro = CLng(Mid(Parametros, pos1, pos2))
    
        pos1 = pos2 + 2
        pos2 = InStr(pos1, Parametros, Separador) - 1
        Pliqnro = Mid(Parametros, pos1, pos2 - pos1 + 1)
    
        pos1 = pos2 + 2
        pos2 = InStr(pos1, Parametros, Separador) - 1
        Todos_Procesos = CBool(Mid(Parametros, pos1, pos2 - pos1 + 1))
    
        pos1 = pos2 + 2
        pos2 = InStr(pos1, Parametros, Separador) - 1
        Lista_Pro = Mid(Parametros, pos1, pos2 - pos1 + 1)
        ' esta lista tiene los nro de procesos separados por comas
        
        pos1 = pos2 + 2
        pos2 = Len(Parametros)
        Todos_Empleados = CBool(Mid(Parametros, pos1, pos2 - pos1 + 1))
        
    End If
End If


' -----------------------------------------------------------

'Comienzo la transaccion
MyBeginTrans

StrSql = "SELECT * FROM tikpedido "
StrSql = StrSql & " WHERE tikpednro =" & TikPedNro

Flog.writeline "Buscando el pedido de tickets"

OpenRecordset StrSql, rs_TikPedido

If Not rs_TikPedido.EOF Then
    
    Pliqnro = rs_TikPedido!Pliqnro
    
    StrSql = "SELECT * FROM emprtik "
    StrSql = StrSql & " WHERE ternro =" & rs_TikPedido!emprtik
    
    Flog.writeline "Buscando la empresa de tickets"
    
    OpenRecordset StrSql, rs_Emprtik
    
    If Not rs_Emprtik.EOF Then
        
        StrSql = "SELECT * FROM ticket "
        StrSql = StrSql & " WHERE emprtik =" & rs_Emprtik!ternro
        
        Flog.writeline "Buscando la marca de tickets"
        
        OpenRecordset StrSql, rs_Ticket
        
        Do While Not rs_Ticket.EOF
        
            If Not Todos_Empleados Then
                StrSql = "SELECT * FROM batch_empleado "
                StrSql = StrSql & " WHERE bpronro =" & NroProcesoBatch
            Else
                StrSql = "SELECT * FROM empleado ORDER BY empleg "
            End If
            
            Flog.writeline "Buscando los empleados"
            
            OpenRecordset StrSql, rs_Empleados
            
            'seteo de las variables de progreso
            Progreso = 0
            CConceptosAProc = rs_Ticket.RecordCount
            If CConceptosAProc = 0 Then
                CConceptosAProc = 1
            End If
            
            CEmpleadosAProc = rs_Empleados.RecordCount
            If CEmpleadosAProc = 0 Then
               CEmpleadosAProc = 1
            End If
            
            IncPorc = ((100 / CEmpleadosAProc) * (100 / CConceptosAProc)) / 100
            
            Do While Not rs_Empleados.EOF
                Monto = 0
                Cantidad = 0
                'Busco los detliq de los conceptos asociados al ticket en el periodo del pedido
                ' para los procesos seleccionados para el empleado
                StrSql = "SELECT sum(detliq.dlimonto) Monto, sum(detliq.dlicant) Cantidad FROM proceso "
                StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
                StrSql = StrSql & " INNER JOIN detliq ON cabliq.cliqnro = detliq.cliqnro "
                StrSql = StrSql & " INNER JOIN ticket_conc ON detliq.concnro = ticket_conc.concnro AND ticket_conc.tiknro =" & rs_Ticket!tiknro
                StrSql = StrSql & " WHERE proceso.pliqnro =" & Pliqnro
                StrSql = StrSql & " AND proceso.pronro IN (" & Lista_Pro & ")"
                StrSql = StrSql & " AND cabliq.empleado =" & rs_Empleados!ternro
                
                Flog.writeline "Buscnado los conceptos para el empleado"
                
                OpenRecordset StrSql, rs_Detliq
            
                If Not rs_Detliq.EOF Then
                    Monto = IIf(IsNull(rs_Detliq!Monto), 0, rs_Detliq!Monto)
                    Cantidad = IIf(IsNull(rs_Detliq!Cantidad), 0, rs_Detliq!Cantidad)
                End If
                            
                If Monto <> 0 Then
                    'Busco si existe el EMP_TICKET
                    StrSql = "SELECT * FROM emp_ticket WHERE " & _
                             " tikpednro = " & TikPedNro & _
                             " AND tiknro = " & rs_Ticket!tiknro & _
                             " AND empleado = " & rs_Empleados!ternro
                             
                    Flog.writeline "Buscando si el empleado ya tiene un ticket de la marca actual"
                    
                    OpenRecordset StrSql, rs_EMP_Ticket
                    
                    If rs_EMP_Ticket.EOF Then
                        StrSql = "INSERT INTO emp_ticket ("
                        StrSql = StrSql & "empleado,tiknro,tikpednro,etikfecha,etikmonto,etikcant,etikmanual"
                        StrSql = StrSql & ") VALUES (" & rs_Empleados!ternro
                        StrSql = StrSql & "," & rs_Ticket!tiknro
                        StrSql = StrSql & "," & TikPedNro
                        StrSql = StrSql & "," & ConvFecha(Date)
                        StrSql = StrSql & "," & Monto
                        StrSql = StrSql & "," & Cantidad
                        StrSql = StrSql & ",0"
                        StrSql = StrSql & " )"
                        
                        Flog.writeline "Insertando el ticket de la marca actual"
                        
                        objConn.Execute StrSql, , adExecuteNoRecords
                        
                    Else
                        If Not CBool(rs_EMP_Ticket!etikmanual) And IsNull(rs_EMP_Ticket!pronro) Then
                            StrSql = "UPDATE emp_ticket SET etikmonto = " & Monto
                            StrSql = StrSql & " , etikfecha = " & ConvFecha(Date)
                            StrSql = StrSql & " , etikcant = " & Cantidad
                            StrSql = StrSql & " WHERE tikpednro = " & TikPedNro
                            StrSql = StrSql & " AND tiknro = " & rs_Ticket!tiknro
                            StrSql = StrSql & " AND empleado = " & rs_Empleados!ternro
                            
                            Flog.writeline "Actualizando el ticket de la marca actual"
                            
                            objConn.Execute StrSql, , adExecuteNoRecords
                        End If
                    End If
                End If
                
                'Actualizo el progreso
                Progreso = Progreso + IncPorc
                TiempoAcumulado = GetTickCount
                
                Flog.writeline "Insertando el ticket de la marca actual"
                
                StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
                         ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
                         "' WHERE bpronro = " & NroProcesoBatch
                objconnProgreso.Execute StrSql, , adExecuteNoRecords
                
                rs_Empleados.MoveNext
            Loop
            rs_Ticket.MoveNext
        Loop
    Else
        Flog.writeline " No se encuentra la Empresa del ticket (EMPRTICK) " & rs_TikPedido!emprtik
    End If
Else
    Flog.writeline " No se encontró el pedido " & TikPedNro
End If

'Fin de la transaccion
MyCommitTrans


'Cierro todo y libero
If rs_TikPedido.State = adStateOpen Then rs_TikPedido.Close
If rs_Emprtik.State = adStateOpen Then rs_Emprtik.Close
If rs_Ticket.State = adStateOpen Then rs_Ticket.Close
If rs_Empleados.State = adStateOpen Then rs_Empleados.Close
If rs_Detliq.State = adStateOpen Then rs_Detliq.Close
If rs_EMP_Ticket.State = adStateOpen Then rs_EMP_Ticket.Close

Set rs_EMP_Ticket = Nothing
Set rs_TikPedido = Nothing
Set rs_Emprtik = Nothing
Set rs_Ticket = Nothing
Set rs_Empleados = Nothing
Set rs_Detliq = Nothing

Exit Sub

CE:
    MyRollbackTrans
    
    MyBeginTrans
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
             "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    MyCommitTrans
    
    HuboError = True
    Flog.writeline
    Flog.writeline "**********************************************************"
    Flog.writeline " Empleado abortado: "
    Flog.writeline " Error: " & Err.Description
    Flog.writeline "**********************************************************"
    Flog.writeline
End Sub


