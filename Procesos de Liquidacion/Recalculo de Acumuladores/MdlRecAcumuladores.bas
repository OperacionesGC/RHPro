Attribute VB_Name = "MdlRecAcumuladores"
Option Explicit

'Const Version = 1.01 'Se agrego para que imprima el nro de version
'Const FechaVersion = "15/12/2005"

'Const Version = 1.02 'Se agrego nuevos mensajes en el log."
'Const FechaVersion = "07/03/2006"

'Const Version = "1.03" '
'Const FechaVersion = "13/02/2009"
'Const UltimaModificacion = " FGZ " 'Encriptacion de string de conexion y SCHEMA de ORACLE
'                                   le agregué que consulte por eof
'
Const Version = "1.04" '
Const FechaVersion = "24/04/2014"
Const UltimaModificacion = "" ' Carmen Quintero - CAS-23925 - SANTANA TEXTIL - BUG EN RECALCULO DE ACUMULADORES - Se modificó para actualice el progreso mas seguido.




Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial del Recalculo de Acumuladores.
' Autor      : FGZ
' Fecha      : 02/02/2004
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
    
    Nombre_Arch = PathFLog & "Recalculo_Acumuladores" & "-" & NroProcesoBatch & ".log"
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    'Abro la conexion
    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    
    'Abro la conexion para actualizar el progreso
    On Error Resume Next
    OpenConnection strconexion, objconnProgreso
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    
    On Error GoTo ME_Main
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Version = " & Version
    Flog.writeline "Modificacion = " & UltimaModificacion
    Flog.writeline "Fecha   = " & FechaVersion
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline
    Flog.writeline "PID = " & PID
    
    TiempoInicialProceso = GetTickCount
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprcprogreso = 1, bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 30 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    If Not rs_batch_proceso.EOF Then
        Flog.writeline "Inicio Proceso " & Now
        bprcparam = rs_batch_proceso!bprcparam
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        Call Rec_acu(NroProcesoBatch, bprcparam)
    Else
        Flog.writeline "No se encontró el proceso " & NroProcesoBatch
    End If
    
    If Not HuboError Then
        Flog.writeline "Fin Proceso sin errores " & Now
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado', bprcprogreso = 100 WHERE bpronro = " & NroProcesoBatch
    Else
        Flog.writeline "Fin Proceso con errores " & Now
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error', bprcprogreso = 100 WHERE bpronro = " & NroProcesoBatch
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
        Flog.writeline Espacios(Tabulador * 0) & " Error General " & Err.Description
        If InStr(1, Err.Description, "ODBC") > 0 Then
            'Fue error de Consulta de SQL
            Flog.writeline
            Flog.writeline Espacios(Tabulador * 0) & "SQL Ejecutado: " & StrSql
            Flog.writeline
        End If
        Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
        Flog.writeline
        
    'Actualizo el progreso
    MyBeginTrans
        StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Error General', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
    MyCommitTrans
    GoTo Fin:
    

End Sub


Public Sub Rec_acu(ByVal NroProceso As Long, ByVal Parametros As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Realiza el proceso de recalculo del acumulador para el/los
'              periodos/procesos indicado por parámetro. (acu_liq y/o acu_mes).
' Autor      : FGZ
' Fecha      : 02/02/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim RecalculaAcumulado As Boolean
Dim RecalculaAcumuladoMensual As Boolean
Dim Nro_Acu As Long
Dim Todos As Boolean
Dim PeriodoDesde_pliqnro As Long
Dim PeriodoHasta_pliqnro As Long
Dim Nro_Pro As Long

Dim pos1 As Integer
Dim pos2 As Integer
Dim FechaDesde As Date
Dim FechaHasta As Date

Dim No_Hay_Ningun As Boolean
Dim Cargo_Algo As Boolean
Dim Monto As Single
Dim Cantidad As Long
Dim Ultimo_pliqanio As Integer
Dim Ultimo_pliqmes As Integer
Dim Ultimo_empleado As Long
Dim Ultimo_acumulador As Long
Dim Entro As Boolean
Dim cantProcesos As Long
Dim IncPorcProceso As Single

Dim rs_Procesos As New ADODB.Recordset
Dim rs_Acu_liq As New ADODB.Recordset
Dim rs_Con_Acum As New ADODB.Recordset
Dim rs_Detliq As New ADODB.Recordset
Dim rs_Periodo As New ADODB.Recordset
Dim rs_Acu_Mes As New ADODB.Recordset

On Error GoTo Manejador:

' Parametros    : Recalculo de Acumulados (acu_liq)
'               : Recalculo de Acumulados Mensuales (acumes)
'               : nro-acu    : nro del acumulador.
'               : Todos      : Todos los procesos (-1 / 0)
'                          -1:  nro-per desde    : nro de periodo.
'                               nro-per hasta    : nro de periodo.
'                           0:  nro-pro    : nro de proceso .
            

' Levanto cada parametro por separado, el separador de parametros es "."
If Not IsNull(Parametros) Then
    If Len(Parametros) >= 1 Then
    
        Flog.writeline Espacios(Tabulador * 1) & "Lista de Parámetros: " & Parametros
        
        Flog.writeline Espacios(Tabulador * 1) & "Obtengo el parámetro 1"
        pos1 = 1
        pos2 = InStr(pos1, Parametros, ".") - 1
        RecalculaAcumulado = CBool(Mid(Parametros, pos1, pos2))
    
        Flog.writeline Espacios(Tabulador * 1) & "Obtengo el parámetro 2"
        pos1 = pos2 + 2
        pos2 = InStr(pos1, Parametros, ".") - 1
        RecalculaAcumuladoMensual = CBool(Mid(Parametros, pos1, pos2 - pos1 + 1))
    
        Flog.writeline Espacios(Tabulador * 1) & "Obtengo el parámetro 3"
        pos1 = pos2 + 2
        pos2 = InStr(pos1, Parametros, ".") - 1
        Nro_Acu = CLng(Mid(Parametros, pos1, pos2 - pos1 + 1))
    
        Flog.writeline Espacios(Tabulador * 1) & "Obtengo el parámetro 4"
        pos1 = pos2 + 2
        pos2 = InStr(pos1, Parametros, ".") - 1
        Todos = CBool(Mid(Parametros, pos1, pos2 - pos1 + 1))
        
        If Not Todos Then
        
            Flog.writeline Espacios(Tabulador * 1) & "Obtengo el parámetro 5 - Un Proceso"
            pos1 = pos2 + 2
            pos2 = Len(Parametros)
            Nro_Pro = CLng(Mid(Parametros, pos1, pos2 - pos1 + 1))
        Else
            Flog.writeline Espacios(Tabulador * 1) & "Obtengo el parámetro 5 - Todos"
            pos1 = pos2 + 2
            pos2 = InStr(pos1, Parametros, ".") - 1
            PeriodoDesde_pliqnro = CLng(Mid(Parametros, pos1, pos2 - pos1 + 1))

            Flog.writeline Espacios(Tabulador * 1) & "Obtengo el parámetro 6"
            pos1 = pos2 + 2
            pos2 = Len(Parametros)
            PeriodoHasta_pliqnro = CLng(Mid(Parametros, pos1, pos2 - pos1 + 1))
        End If
    End If
End If

Flog.writeline Espacios(Tabulador * 1) & "------------------------------------------------------------------"
Flog.writeline Espacios(Tabulador * 1) & "--------------------------- Parametros ---------------------------"
Flog.writeline Espacios(Tabulador * 1) & "Recalcula Acumulados (acu_liq) " & RecalculaAcumulado
Flog.writeline Espacios(Tabulador * 1) & "Recalcula Acumulados Mensuales (acumes) " & RecalculaAcumuladoMensual
Flog.writeline Espacios(Tabulador * 1) & "Acumulador " & Nro_Acu
Flog.writeline Espacios(Tabulador * 1) & "Todos los procesos (-1 / 0) " & Todos
Flog.writeline Espacios(Tabulador * 1) & "Periodo desde " & PeriodoDesde_pliqnro
Flog.writeline Espacios(Tabulador * 1) & "Periodo Hasta " & PeriodoHasta_pliqnro
Flog.writeline Espacios(Tabulador * 1) & "Nro de Proceso " & Nro_Pro
Flog.writeline Espacios(Tabulador * 1) & "------------------------------------------------------------------"
Flog.writeline

If Todos Then
    ' Busco las fechas desde y hasta de los periodos
    StrSql = "SELECT * FROM periodo " & _
            " WHERE periodo.pliqnro =" & PeriodoDesde_pliqnro
    OpenRecordset StrSql, rs_Periodo
    If Not rs_Periodo.EOF Then
        FechaDesde = rs_Periodo!pliqdesde
    End If

    StrSql = "SELECT * FROM periodo " & _
            " WHERE periodo.pliqnro =" & PeriodoHasta_pliqnro
    OpenRecordset StrSql, rs_Periodo
    If Not rs_Periodo.EOF Then
        FechaHasta = rs_Periodo!pliqhasta
    End If
End If


If Todos Then
    StrSql = "SELECT * FROM periodo " & _
            " INNER JOIN proceso ON proceso.pliqnro = periodo.pliqnro " & _
            " INNER JOIN cabliq ON cabliq.pronro = proceso.pronro " & _
            " WHERE periodo.pliqdesde >=" & ConvFecha(FechaDesde) & _
            " AND periodo.pliqhasta <=" & ConvFecha(FechaHasta) & _
            " ORDER BY periodo.pliqnro,proceso.pronro,cabliq.empleado"
Else
    StrSql = "SELECT * FROM periodo " & _
            " INNER JOIN proceso ON proceso.pliqnro = periodo.pliqnro " & _
            " INNER JOIN cabliq ON cabliq.pronro = proceso.pronro " & _
            " WHERE proceso.pronro = " & Nro_Pro & _
            " ORDER BY periodo.pliqnro,proceso.pronro,cabliq.empleado"
End If
OpenRecordset StrSql, rs_Procesos


'FGZ - 12/02/2009 - le agregué que consulte por eof --------
If rs_Procesos.EOF Then
    cantProcesos = 1
Else
    cantProcesos = rs_Procesos.RecordCount
End If
'cantProcesos = rs_Procesos.RecordCount
'FGZ - 12/02/2009 - le agregué que consulte por eof --------

If cantProcesos = 0 Then
   cantProcesos = 1
End If
    
IncPorcProceso = (10 / cantProcesos)
Progreso = 0

'Comienzo la transaccion
MyBeginTrans

Flog.writeline Espacios(Tabulador * 1) & "Actualizar los acumuladores del concepto"
Flog.writeline
'Actualizar los acumuladores del concepto
If RecalculaAcumulado Then
    Flog.writeline Espacios(Tabulador * 1) & "Aculiq ===================>>>>>"
    Flog.writeline
    
    Do While Not rs_Procesos.EOF
        'caso especial para los acumuladores que no tiene ningun concepto para sumar
        ' en el acumulador, pero en algun momento tuvieron, se debe ocultar */
        IncPorc = (IncPorcProceso / cantProcesos)
        Flog.writeline Espacios(Tabulador * 2) & "caso especial para los acumuladores que no tiene ningun concepto para sumar"
        
        StrSql = "SELECT * FROM acu_liq " & _
                " WHERE acunro = " & Nro_Acu & _
                " AND cliqnro =" & rs_Procesos!cliqnro
        OpenRecordset StrSql, rs_Acu_liq
        If Not rs_Acu_liq.EOF Then
            'si ya existía, ver si seguirá existiendo
            'Validar que al menos exista un concepto en Detliq que sume en el acumulador
            Flog.writeline Espacios(Tabulador * 3) & "si ya existía, ver si seguirá existiendo"
            
            No_Hay_Ningun = True
            Cargo_Algo = False
           'verifica :
'            StrSql = "SELECT * FROM con_acum " & _
'                    " WHERE acunro = " & Nro_Acu
'            OpenRecordset StrSql, rs_Con_Acum
'
'            Do While Not rs_Con_Acum.EOF And No_Hay_Ningun
'                StrSql = "SELECT * FROM detliq " & _
'                        " WHERE concnro = " & rs_Con_Acum!concnro & _
'                        " AND cliqnro =" & rs_Procesos!cliqnro
'                OpenRecordset StrSql, rs_Detliq
'                If Not rs_Detliq.EOF Then
'                    'ocultar el ACU_LIQ, NO HAY NADA PARA SUMAR
'                    No_Hay_Ningun = False
'                End If
'
'                rs_Con_Acum.MoveNext
'            Loop
           
            StrSql = "SELECT * FROM con_acum " & _
                     " INNER JOIN detliq ON con_acum.concnro = detliq.concnro " & _
                     " WHERE acunro = " & Nro_Acu & _
                     " AND cliqnro =" & rs_Procesos!cliqnro
            OpenRecordset StrSql, rs_Con_Acum
            If Not rs_Con_Acum.EOF Then
                'ocultar el ACU_LIQ, NO HAY NADA PARA SUMAR
                No_Hay_Ningun = False
                Flog.writeline Espacios(Tabulador * 3) & "ocultar el ACU_LIQ, NO HAY NADA PARA SUMAR " & No_Hay_Ningun
            Else
                No_Hay_Ningun = True
            End If
           
            If No_Hay_Ningun Then
                'no quedo ningun concepto que sume en el acumulador en los detalles de liquidacion
                StrSql = "DELETE acu_liq WHERE acunro = " & Nro_Acu & _
                         " AND cliqnro =" & rs_Procesos!cliqnro
                Flog.writeline Espacios(Tabulador * 3) & "no quedo ningun concepto que sume en el acumulador en los detalles de liquidacion "
                Flog.writeline Espacios(Tabulador * 3) & " SQL: " & StrSql
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
        End If
        
        StrSql = "SELECT * FROM detliq " & _
                " INNER JOIN cabliq ON cabliq.cliqnro = detliq.cliqnro " & _
                " INNER JOIN con_acum ON con_acum.concnro = detliq.concnro " & _
                " WHERE con_acum.acunro = " & Nro_Acu & _
                " AND cabliq.cliqnro =" & rs_Procesos!cliqnro & _
                " ORDER BY con_acum.acunro"
        OpenRecordset StrSql, rs_Detliq
        
        'FGZ - 12/02/2009 - le agregué que consulte por eof --------
        If rs_Detliq.EOF Then
            CConceptosAProc = 1
        Else
            CConceptosAProc = rs_Detliq.RecordCount
        End If
        'CConceptosAProc = rs_Detliq.RecordCount
        'FGZ - 12/02/2009 - le agregué que consulte por eof --------
        
        If CConceptosAProc = 0 Then
            CConceptosAProc = 1
        End If
        
        'Calculo el incremento de cada elemento dentro del proceso
        Flog.writeline Espacios(Tabulador * 2) & "Calculo el incremento de cada elemento dentro del proceso "
        Flog.writeline
        IncPorc = (IncPorcProceso / CConceptosAProc)
        
        Ultimo_acumulador = -1
        Do While Not rs_Detliq.EOF
            Flog.writeline Espacios(Tabulador * 3) & "Acumulador " & rs_Detliq!acuNro
            If rs_Detliq!acuNro <> Ultimo_acumulador Then
                Flog.writeline Espacios(Tabulador * 3) & "Cambio de acumulador "
                Ultimo_acumulador = rs_Detliq!acuNro
                
                StrSql = "SELECT * FROM acu_liq " & _
                        " WHERE acunro = " & Nro_Acu & _
                        " AND cliqnro =" & rs_Procesos!cliqnro
                OpenRecordset StrSql, rs_Acu_liq
                If rs_Acu_liq.EOF Then
                    StrSql = "INSERT INTO acu_liq (" & _
                         "acunro,cliqnro,almonto,almontoreal,alcant" & _
                         ") VALUES (" & Nro_Acu & _
                         "," & rs_Procesos!cliqnro & _
                         ", 0 " & _
                         ", 0" & _
                         ", 0" & _
                         " )"
                    Flog.writeline Espacios(Tabulador * 4) & "Inserto en cero"
                    objConn.Execute StrSql, , adExecuteNoRecords
                    Cargo_Algo = False
                Else
                    If Not Cargo_Algo Then
                        Monto = 0
                        Cantidad = 0
                    Else
                        Monto = rs_Acu_liq!almonto
                        Cantidad = rs_Acu_liq!alcant
                    End If
                    Cargo_Algo = True
                    
                    StrSql = "UPDATE acu_liq SET " & _
                             " alcant = " & Cantidad & _
                             ", almonto =" & Monto & _
                            " WHERE acunro = " & Nro_Acu & _
                            " AND cliqnro = " & rs_Procesos!cliqnro
                    Flog.writeline Espacios(Tabulador * 4) & "Actualizo a " & Monto
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
            End If
            StrSql = "UPDATE acu_liq SET " & _
                     " alcant = alcant + " & rs_Detliq!dlicant & _
                    ", almonto = almonto + " & rs_Detliq!dlimonto & _
                    " WHERE acunro = " & Nro_Acu & _
                    " AND cliqnro = " & rs_Procesos!cliqnro
            Flog.writeline Espacios(Tabulador * 3) & "Actualizo acumualdor + " & rs_Detliq!dlimonto
            objConn.Execute StrSql, , adExecuteNoRecords
        
            rs_Detliq.MoveNext
            
            'Actualizar el progreso
            TiempoFinalProceso = GetTickCount
            Progreso = Progreso + IncPorc
            StrSql = "UPDATE batch_proceso SET bprctiempo = " & (TiempoFinalProceso - TiempoInicialProceso) & ", bprcprogreso = " & Progreso & " WHERE bpronro = " & NroProcesoBatch
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
            
            Flog.writeline
            Flog.writeline Espacios(Tabulador * 4) & "Progreso " & Progreso
            Flog.writeline Espacios(Tabulador * 4) & "Hora " & Now
            Flog.writeline
        Loop 'periodo,proceso,cabliq
    
        rs_Procesos.MoveNext
        'Actualizar el progreso
        TiempoFinalProceso = GetTickCount
        Progreso = Progreso + IncPorc
        StrSql = "UPDATE batch_proceso SET bprctiempo = " & (TiempoFinalProceso - TiempoInicialProceso) & ", bprcprogreso = " & Progreso & " WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Loop
End If


If RecalculaAcumuladoMensual Then
    Flog.writeline Espacios(Tabulador * 1) & "Acu_mes ===================>>>>>"
    Flog.writeline
    If Todos Then
        StrSql = "SELECT * FROM periodo " & _
                " INNER JOIN proceso ON proceso.pliqnro = periodo.pliqnro " & _
                " INNER JOIN cabliq ON cabliq.pronro = proceso.pronro " & _
                " INNER JOIN acu_liq ON cabliq.cliqnro = acu_liq.cliqnro " & _
                " WHERE periodo.pliqdesde >=" & ConvFecha(FechaDesde) & _
                " AND periodo.pliqhasta <=" & ConvFecha(FechaHasta) & _
                " AND acu_liq.acunro =" & Nro_Acu & _
                " ORDER BY periodo.pliqanio,cabliq.empleado, acu_liq.acunro"
    Else
        StrSql = "SELECT * FROM periodo " & _
                " INNER JOIN proceso ON proceso.pliqnro = periodo.pliqnro " & _
                " INNER JOIN cabliq ON cabliq.pronro = proceso.pronro " & _
                " INNER JOIN acu_liq ON cabliq.cliqnro = acu_liq.cliqnro " & _
                " WHERE proceso.pronro = " & Nro_Pro & _
                " AND acu_liq.acunro =" & Nro_Acu & _
                " ORDER BY periodo.pliqanio,periodo.pliqmes,cabliq.empleado, acu_liq.acunro"
    End If
    If rs_Procesos.State = adStateOpen Then rs_Procesos.Close
    OpenRecordset StrSql, rs_Procesos
    
    'FGZ - 12/02/2009 - le agregué que consulte por eof --------
    If rs_Procesos.EOF Then
        CConceptosAProc = 1
    Else
        CConceptosAProc = rs_Procesos.RecordCount
    End If
    'CConceptosAProc = rs_Procesos.RecordCount
    'FGZ - 12/02/2009 - le agregué que consulte por eof --------
    
    If CConceptosAProc = 0 Then
        CConceptosAProc = 1
    End If
    IncPorc = (10 / CConceptosAProc)
    
    If Progreso = 0 Then
        Progreso = 10
    Else
        Progreso = Progreso + IncPorc
    End If
    
    Ultimo_pliqanio = -1
    Ultimo_pliqmes = -1
    Ultimo_empleado = -1
    Ultimo_acumulador = -1
    Do While Not rs_Procesos.EOF
        IncPorc = (5 / CConceptosAProc)
        If rs_Procesos!pliqanio <> Ultimo_pliqanio Or (rs_Procesos!pliqanio = Ultimo_pliqanio And rs_Procesos!pliqmes <> Ultimo_pliqmes) Or rs_Procesos!Empleado <> Ultimo_empleado Or rs_Procesos!acuNro <> Ultimo_acumulador Then
            Ultimo_acumulador = rs_Procesos!acuNro
            Ultimo_empleado = rs_Procesos!Empleado
            Ultimo_pliqanio = rs_Procesos!pliqanio
            Ultimo_pliqmes = rs_Procesos!pliqmes
            
            StrSql = "SELECT * FROM acu_mes " & _
                    " WHERE acunro = " & Nro_Acu & _
                    " AND amanio =" & rs_Procesos!pliqanio & _
                    " AND ammes =" & rs_Procesos!pliqmes & _
                    " AND ternro =" & rs_Procesos!Empleado
            OpenRecordset StrSql, rs_Acu_Mes
            
            If rs_Acu_Mes.EOF Then
                StrSql = "INSERT INTO acu_mes (" & _
                     "acunro,amanio,ammes,ammonto,amcant,ternro" & _
                     ") VALUES (" & Nro_Acu & _
                     "," & rs_Procesos!pliqanio & _
                     "," & rs_Procesos!pliqmes & _
                     ",0,0" & _
                     "," & rs_Procesos!Empleado & _
                     " )"
                Flog.writeline Espacios(Tabulador * 3) & "Inserto en cero"
                objConn.Execute StrSql, , adExecuteNoRecords
            Else
                StrSql = "UPDATE acu_mes SET " & _
                         " amcant = 0" & _
                        ", ammonto = 0" & _
                        " WHERE acunro = " & Nro_Acu & _
                        " AND amanio =" & rs_Procesos!pliqanio & _
                        " AND ammes =" & rs_Procesos!pliqmes & _
                        " AND ternro =" & rs_Procesos!Empleado
                Flog.writeline Espacios(Tabulador * 3) & "Actualizo a cero"
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
        End If
        StrSql = "UPDATE acu_mes SET " & _
                 " amcant = amcant + " & rs_Procesos!alcant & _
                ", ammonto = ammonto + " & rs_Procesos!almonto & _
                " WHERE acunro = " & Nro_Acu & _
                " AND amanio =" & rs_Procesos!pliqanio & _
                " AND ammes =" & rs_Procesos!pliqmes & _
                " AND ternro =" & rs_Procesos!Empleado
        Flog.writeline Espacios(Tabulador * 2) & "Actualizo acumulador " & Nro_Acu & " + " & rs_Procesos!almonto
        objConn.Execute StrSql, , adExecuteNoRecords


        'Actualizar el progreso
        TiempoFinalProceso = GetTickCount
        Progreso = Progreso + IncPorc
        StrSql = "UPDATE batch_proceso SET bprctiempo = " & (TiempoFinalProceso - TiempoInicialProceso) & ", bprcprogreso = " & Progreso & " WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords

        rs_Procesos.MoveNext
    Loop
    
    If rs_Procesos.State = adStateOpen Then rs_Procesos.Close
    
    
    Flog.writeline Espacios(Tabulador * 1) & "LIMPIEZA DE ACUMES CUANDO NO EXISTEN ACULIQ EN ESE MES"
    'LIMPIEZA DE ACUMES CUANDO NO EXISTEN ACULIQ EN ESE MES.
    If Todos Then
        StrSql = "SELECT * FROM periodo " & _
                " INNER JOIN proceso ON proceso.pliqnro = periodo.pliqnro " & _
                " WHERE periodo.pliqdesde >=" & ConvFecha(FechaDesde) & _
                " AND periodo.pliqhasta <=" & ConvFecha(FechaHasta) & _
                " ORDER BY periodo.pliqnro"
    Else
        StrSql = "SELECT * FROM periodo " & _
                " INNER JOIN proceso ON proceso.pliqnro = periodo.pliqnro " & _
                " WHERE proceso.pronro = " & Nro_Pro & _
                " ORDER BY periodo.pliqnro"
    End If
    
    If rs_Periodo.State = adStateOpen Then rs_Periodo.Close
    OpenRecordset StrSql, rs_Periodo

    'FGZ - 12/02/2009 - le agregué que consulte por eof --------
    If rs_Periodo.EOF Then
        cantProcesos = 1
    Else
        cantProcesos = rs_Periodo.RecordCount
    End If
    'cantProcesos = rs_Periodo.RecordCount
    'FGZ - 12/02/2009 - le agregué que consulte por eof --------
    
    If cantProcesos = 0 Then
        cantProcesos = 1
    End If
    IncPorcProceso = (20 / cantProcesos)
    
    If Progreso = 0 Then
        Progreso = 20
    Else
        Progreso = Progreso + IncPorc
    End If
    
    Do While Not rs_Periodo.EOF
        StrSql = "SELECT * FROM acu_mes " & _
                " WHERE acunro = " & Nro_Acu & _
                " AND amanio =" & rs_Periodo!pliqanio & _
                " AND ammes =" & rs_Periodo!pliqmes
        OpenRecordset StrSql, rs_Acu_Mes
        
        'FGZ - 12/02/2009 - le agregué que consulte por eof --------
        If rs_Acu_Mes.EOF Then
            CConceptosAProc = 1
        Else
            CConceptosAProc = rs_Acu_Mes.RecordCount
        End If
        'CConceptosAProc = rs_Acu_Mes.RecordCount
        'FGZ - 12/02/2009 - le agregué que consulte por eof --------
        
        If CConceptosAProc = 0 Then
            CConceptosAProc = 1
        End If
        
        'Calculo el incremento de cada elemento dentro del proceso
        IncPorc = (IncPorcProceso / CConceptosAProc)
        
        Do While Not rs_Acu_Mes.EOF
            Entro = False
            StrSql = "SELECT * FROM proceso " & _
                    " INNER JOIN cabliq ON cabliq.pronro = proceso.pronro " & _
                    " INNER JOIN acu_liq ON cabliq.cliqnro = acu_liq.cliqnro " & _
                    " WHERE proceso.pliqnro =" & rs_Periodo!pliqnro & _
                    " AND acu_liq.acunro = " & Nro_Acu & _
                    " AND cabliq.empleado =" & rs_Acu_Mes!Ternro
            If rs_Procesos.State = adStateOpen Then rs_Procesos.Close
            OpenRecordset StrSql, rs_Procesos
            
            If Not rs_Procesos.EOF Then
                Entro = True
                Flog.writeline Espacios(Tabulador * 2) & "Entró"
            End If
            
            If Not Entro Then
                Flog.writeline Espacios(Tabulador * 2) & "No Entró"
                StrSql = "UPDATE acu_mes SET " & _
                         " amcant = 0 " & _
                        ", ammonto = 0" & _
                        " WHERE acunro = " & Nro_Acu & _
                        " AND amanio =" & rs_Periodo!pliqanio & _
                        " AND ammes =" & rs_Periodo!pliqmes & _
                        " AND ternro =" & rs_Acu_Mes!Ternro
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
             
            rs_Acu_Mes.MoveNext
            
            'Actualizar el progreso
            TiempoFinalProceso = GetTickCount
            Progreso = Progreso + IncPorc
            If Progreso > 100 Then
               Progreso = 100
            End If
            StrSql = "UPDATE batch_proceso SET bprctiempo = " & (TiempoFinalProceso - TiempoInicialProceso) & ", bprcprogreso = " & Progreso & " WHERE bpronro = " & NroProcesoBatch
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
            
        Loop
        
        rs_Periodo.MoveNext
        'Actualizar el progreso
        TiempoFinalProceso = GetTickCount
        Progreso = Progreso + IncPorc
        If Progreso > 100 Then
            Progreso = 100
        End If
        StrSql = "UPDATE batch_proceso SET bprctiempo = " & (TiempoFinalProceso - TiempoInicialProceso) & ", bprcprogreso = " & Progreso & " WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Loop
    
    If rs_Periodo.State = adStateOpen Then rs_Periodo.Close
    
End If
MyCommitTrans

Exit Sub

Manejador:
    MyRollbackTrans
    
    HuboError = True
    EmpleadoSinError = False
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
        Flog.writeline Espacios(Tabulador * 0) & " Error: " & Err.Description
        If InStr(1, Err.Description, "ODBC") > 0 Then
            'Fue error de Consulta de SQL
            Flog.writeline
            Flog.writeline Espacios(Tabulador * 0) & "SQL Ejecutado: " & StrSql
            Flog.writeline
        End If
        Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
        Flog.writeline
        
    'Actualizo el progreso
    'MyBeginTrans
        Progreso = Progreso + IncPorc
        TiempoAcumulado = GetTickCount
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
                 ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
                 "' WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
    'MyCommitTrans

End Sub
