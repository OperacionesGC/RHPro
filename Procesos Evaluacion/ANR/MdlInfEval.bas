Attribute VB_Name = "MdlRepInfEvaluaciones"
Option Explicit

Global IdUser As String
Global Fecha As Date
Global Hora As String

Global Fecha_Inicio_periodo As Date
Global Fecha_Fin_Periodo As Date
Global StrSql2 As String


Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial del Proceso.
' Autor      : FGZ
' Fecha      : 07/09/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
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

    'Abro la conexion
    OpenConnection strconexion, objConn
    OpenConnection strconexion, objconnProgreso
    
    Nombre_Arch = PathFLog & "Informe_evaluaciones" & "-" & NroProcesoBatch & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "PID = " & PID
    'Cambio el estado del proceso a Procesando y el PID
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 36 AND bpronro =" & NroProcesoBatch
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
Public Sub Generacion(ByVal FiltroEmpleado As String, ByVal bpronro As Long, ByVal Nroliq As Long, ByVal Todos_Pro As Boolean, ByVal Proc_Aprob As Integer, ByVal Empresa As Long, ByVal Agrupado As Boolean, _
    ByVal AgrupaTE1 As Boolean, ByVal Tenro1 As Long, Estrnro1 As Long, _
    ByVal AgrupaTE2 As Boolean, ByVal Tenro2 As Long, Estrnro2 As Long, _
    ByVal AgrupaTE3 As Boolean, ByVal Tenro3 As Long, Estrnro3 As Long)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del informe de Evaluaciones
' Autor      : FGZ
' Fecha      : 13/09/2004
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim Aux_Total_Importe
Dim strLinea As String
Dim Aux_Linea As String

Dim FechaDesde As Date
Dim FechaHasta As Date
Dim Nro_Reporte As Integer
Dim Conf_Ok As Boolean
Dim concnro As Long
Dim Nro_Concepto As Long

Dim Estructura1 As Long
Dim Estructura2 As Long

Dim rs_Confrep As New ADODB.Recordset
Dim rs_Periodo As New ADODB.Recordset
Dim rs_Procesos As New ADODB.Recordset
Dim rs_Estructura As New ADODB.Recordset
Dim rs_Reporte As New ADODB.Recordset

'cargo el periodo
StrSql = "SELECT * FROM periodo WHERE pliqnro = " & CStr(Nroliq)
OpenRecordset StrSql, rs_Periodo
If rs_Periodo.EOF Then
    Flog.writeline "No se encontró el Periodo"
    Exit Sub
End If
Fecha_Inicio_periodo = rs_Periodo!pliqdesde
Fecha_Fin_Periodo = rs_Periodo!pliqhasta

'Configuracion del Reporte
Nro_Reporte = 0
Conf_Ok = False
StrSql = "SELECT * FROM confrep WHERE repnro = " & Nro_Reporte
OpenRecordset StrSql, rs_Confrep
If rs_Confrep.EOF Then
    Flog.writeline "No se encontró la configuración del Reporte"
    Exit Sub
Else
    Do While Not rs_Confrep.EOF
        Select Case rs_Confrep!confnrocol
        Case 1:
        Case 2:
        Case 3:

        End Select
        rs_Confrep.MoveNext
    Loop
End If
If Not Conf_Ok Then
    Flog.writeline "Columna 3. El concepto no esta configurado"
    Exit Sub
End If

' Comienzo la transaccion
MyBeginTrans

'Depuracion del Temporario
StrSql = "DELETE FROM rep_jub_mov "
objConn.Execute StrSql, , adExecuteNoRecords

'Busco los procesos a evaluar
StrSql = "SELECT cabliq.*, proceso.*, periodo.*, empleado.* "
If AgrupaTE1 Then
    StrSql = StrSql & ", te1.tenro tenro1, te1.estrnro estrnro1"
End If
If AgrupaTE2 Then
    StrSql = StrSql & ", te2.tenro tenro2, te2.estrnro estrnro2"
End If
If AgrupaTE3 Then
    StrSql = StrSql & ", te3.tenro tenro3, te3.estrnro estrnro3"
End If
StrSql = StrSql & "  FROM  periodo "
StrSql = StrSql & " INNER JOIN proceso ON periodo.pliqnro = proceso.pliqnro "
StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
StrSql = StrSql & " INNER JOIN empleado ON cabliq.empleado = empleado.ternro "
StrSql = StrSql & " INNER JOIN his_estructura  ON his_estructura.ternro = empleado.ternro and his_estructura.tenro = 10 "
StrSql = StrSql & " INNER JOIN empresa ON empresa.estrnro = his_estructura.estrnro AND empresa.empnro =" & Empresa
If AgrupaTE1 Then
    StrSql = StrSql & " INNER JOIN his_estructura TE1 ON te1.ternro = empleado.ternro "
End If
If AgrupaTE2 Then
    StrSql = StrSql & " INNER JOIN his_estructura TE2 ON te2.ternro = empleado.ternro "
End If
If AgrupaTE3 Then
    StrSql = StrSql & " INNER JOIN his_estructura TE3 ON te3.ternro = empleado.ternro "
End If

StrSql = StrSql & " WHERE periodo.pliqnro =" & Nroliq
StrSql = StrSql & " AND " & FiltroEmpleado
StrSql = StrSql & " AND empresa.empnro =" & Empresa
StrSql = StrSql & " AND (his_estructura.htetdesde <= " & ConvFecha(Fecha_Fin_Periodo) & ") AND "
StrSql = StrSql & " ((" & ConvFecha(Fecha_Fin_Periodo) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
If Not Todos_Pro Then
    StrSql = StrSql & " AND proceso.pronro IN (" & NroProc & ")"
Else
    StrSql = StrSql & " AND proaprob = " & CInt(Proc_Aprob)
End If
If AgrupaTE1 Then
    StrSql = StrSql & " AND  te1.tenro = " & Tenro1 & " AND "
    If Estrnro1 <> 0 Then
        StrSql = StrSql & " te1.estrnro = " & Estrnro1 & " AND "
    End If
    StrSql = StrSql & " (te1.htetdesde <= " & ConvFecha(Fecha_Fin_Periodo) & ") AND "
    StrSql = StrSql & " ((" & ConvFecha(Fecha_Fin_Periodo) & " <= te1.htethasta) or (te1.htethasta is null)) "
End If
If AgrupaTE2 Then
    StrSql = StrSql & " AND te2.tenro = " & Tenro2 & " AND "
    If Estrnro2 <> 0 Then
        StrSql = StrSql & " te2.estrnro = " & Estrnro2 & " AND "
    End If
    StrSql = StrSql & " (te2.htetdesde <= " & ConvFecha(Fecha_Fin_Periodo) & ") AND "
    StrSql = StrSql & " ((" & ConvFecha(Fecha_Fin_Periodo) & " <= te2.htethasta) or (te2.htethasta is null))  "
End If
If AgrupaTE3 Then
    StrSql = StrSql & " AND te3.tenro = " & Tenro3 & " AND "
    If Estrnro3 <> 0 Then
        StrSql = StrSql & " te3.estrnro = " & Estrnro3 & " AND "
    End If
    StrSql = StrSql & " (te3.htetdesde <= " & ConvFecha(Fecha_Fin_Periodo) & ") AND "
    StrSql = StrSql & " ((" & ConvFecha(Fecha_Fin_Periodo) & " <= te3.htethasta) or (te3.htethasta is null))"
End If
StrSql = StrSql & " ORDER BY empleado.ternro"
OpenRecordset StrSql, rs_Procesos


'seteo de las variables de progreso
Progreso = 0
CConceptosAProc = rs_Procesos.RecordCount
If CConceptosAProc = 0 Then
    CConceptosAProc = 1
End If
IncPorc = (100 / CConceptosAProc)

'inicializo

Do While Not rs_Procesos.EOF
        
        
        
        
        'Si no existe el rep_juv_mov
        StrSql = "SELECT * FROM rep_jub_mov "
        StrSql = StrSql & " WHERE ternro = " & rs_Procesos!ternro
        StrSql = StrSql & " AND bpronro = " & bpronro
        StrSql = StrSql & " AND pliqnro = " & Nroliq
        StrSql = StrSql & " AND empresa = " & Empresa
        If Not Todos_Pro Then
            StrSql = StrSql & " AND pronro = '" & ListaNroProc & "'"
        Else
            StrSql = StrSql & " AND pronro = '0'"
            StrSql = StrSql & " AND proaprob= " & CInt(Proc_Aprob)
        End If
        OpenRecordset StrSql, rs_Reporte
    
        If rs_Reporte.EOF Then
            'Inserto
            StrSql = "INSERT INTO rep_jub_mov (bpronro,pliqnro,pronro,proaprob,empresa,iduser,fecha,hora,"
            StrSql = StrSql & "tiporegistro,nroidentificador,tidnro,nrodoc,importe,"
            StrSql = StrSql & "ternro,empleg,apeynom,"
            StrSql = StrSql & "tenro1,estrnro1,tedesc1,estrdesc1,tenro2,estrnro2,tedesc2,estrdesc2,tenro3,estrnro3,tedesc3,estrdesc3 "
            StrSql = StrSql & ") VALUES ("
            StrSql = StrSql & bpronro & ","
            StrSql = StrSql & Nroliq & ","
            If Not Todos_Pro Then
                StrSql = StrSql & "'" & ListaNroProc & "',"
                StrSql = StrSql & rs_Procesos!proaprob & ","
            Else
                StrSql = StrSql & "0" & ","
                StrSql = StrSql & CInt(Proc_Aprob) & ","
            End If
            StrSql = StrSql & Empresa & ","
            StrSql = StrSql & "'" & IdUser & "',"
            StrSql = StrSql & ConvFecha(Fecha) & ","
            StrSql = StrSql & "'" & Hora & "',"
            
            StrSql = StrSql & "'" & Left(Reg3.Tipo_Reg, 1) & "',"
            StrSql = StrSql & "'" & Left(Reg3.Nro_ID, 15) & "',"
            StrSql = StrSql & "'" & Left(Reg3.Tipo_Doc, 1) & "',"
            StrSql = StrSql & "'" & Left(Reg3.Nro_Doc, 8) & "',"
            StrSql = StrSql & CSng(Reg3.Importe) & ","
            
            StrSql = StrSql & rs_Procesos!ternro & ","
            StrSql = StrSql & rs_Procesos!empleg & ","
            StrSql = StrSql & "'" & rs_Procesos!terape
            If Not IsNull(rs_Procesos!terape2) Then
                StrSql = StrSql & " " & rs_Procesos!terape2
            End If
            StrSql = StrSql & ", " & rs_Procesos!ternom
            If Not IsNull(rs_Procesos!ternom2) Then
                StrSql = StrSql & " " & rs_Procesos!ternom2
            End If
            StrSql = StrSql & "'" & ","
            
            'Estructuras
            If AgrupaTE1 Then
                StrSql = StrSql & Tenro1 & ","
            Else
                StrSql = StrSql & "null" & ","
            End If
            StrSql = StrSql & Estrnro1 & ","
            
            'Descripcion tipo estructura
            If AgrupaTE1 Then
                StrSql2 = "SELECT * FROM tipoestructura WHERE tenro =" & rs_Procesos!Tenro1
                If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
                OpenRecordset StrSql2, rs_Estructura
                If Not rs_Estructura.EOF Then
                    StrSql = StrSql & "'" & rs_Estructura!tedabr & "'" & ","
                Else
                    StrSql = StrSql & "' '" & ","
                End If
                'Descripcion Estructura
                StrSql2 = "SELECT * FROM estructura WHERE estrnro =" & rs_Procesos!Estrnro1
                If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
                OpenRecordset StrSql2, rs_Estructura
                If Not rs_Estructura.EOF Then
                    StrSql = StrSql & "'" & rs_Estructura!estrdabr & "'" & ","
                Else
                    StrSql = StrSql & "' '" & ","
                End If
            Else
                StrSql = StrSql & "' '" & ","
                StrSql = StrSql & "' '" & ","
            End If
            
            If AgrupaTE2 Then
                StrSql = StrSql & Tenro2 & ","
            Else
                StrSql = StrSql & "null" & ","
            End If
            StrSql = StrSql & Estrnro2 & ","
            
            If AgrupaTE2 Then
                'Descripcion tipo estructura
                StrSql2 = "SELECT * FROM tipoestructura WHERE tenro =" & rs_Procesos!Tenro2
                If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
                OpenRecordset StrSql2, rs_Estructura
                If Not rs_Estructura.EOF Then
                    StrSql = StrSql & "'" & rs_Estructura!tedabr & "'" & ","
                Else
                    StrSql = StrSql & "' '" & ","
                End If
                'Descripcion Estructura
                StrSql2 = "SELECT * FROM estructura WHERE estrnro =" & rs_Procesos!Estrnro2
                If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
                OpenRecordset StrSql2, rs_Estructura
                If Not rs_Estructura.EOF Then
                    StrSql = StrSql & "'" & rs_Estructura!estrdabr & "'" & ","
                Else
                    StrSql = StrSql & "' '" & ","
                End If
            Else
                StrSql = StrSql & "' '" & ","
                StrSql = StrSql & "' '" & ","
            End If
            
            If AgrupaTE3 Then
                StrSql = StrSql & Tenro3 & ","
            Else
                StrSql = StrSql & "null" & ","
            End If
            StrSql = StrSql & Estrnro3 & ","
            
            'Descripcion tipo estructura
            If AgrupaTE3 Then
                StrSql2 = "SELECT * FROM tipoestructura WHERE tenro =" & rs_Procesos!Tenro3
                If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
                OpenRecordset StrSql2, rs_Estructura
                If Not rs_Estructura.EOF Then
                    StrSql = StrSql & "'" & rs_Estructura!tedabr & "'" & ","
                Else
                    StrSql = StrSql & "' '" & ","
                End If
                'Descripcion Estructura
                StrSql2 = "SELECT * FROM estructura WHERE estrnro =" & rs_Procesos!Estrnro3
                If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
                OpenRecordset StrSql2, rs_Estructura
                If Not rs_Estructura.EOF Then
                    StrSql = StrSql & "'" & rs_Estructura!estrdabr & "'"
                Else
                    StrSql = StrSql & "' '"
                End If
            Else
                StrSql = StrSql & "' '" & ","
                StrSql = StrSql & "' '"
            End If
            
            StrSql = StrSql & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        Else
            'Actualizo
            StrSql = "UPDATE rep_jub_mov SET importe = importe + " & Reg3.Importe
            StrSql = StrSql & " WHERE ternro = " & rs_Procesos!ternro
            StrSql = StrSql & " AND bpronro = " & bpronro
            StrSql = StrSql & " AND pliqnro = " & Nroliq
            StrSql = StrSql & " AND empresa = " & Empresa
            If Not Todos_Pro Then
                StrSql = StrSql & " AND pronro = '" & NroProc & "'"
            Else
                StrSql = StrSql & " AND pronro = '0'"
                StrSql = StrSql & " AND proaprob= " & CInt(Proc_Aprob)
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
    End If
    'Actualizo el progreso del Proceso
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
             "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
                
    'Siguiente proceso
    rs_Procesos.MoveNext
Loop
fAuxiliar.Close

   
'Fin de la transaccion
MyCommitTrans

If rs_Confrep.State = adStateOpen Then rs_Confrep.Close
If rs_Procesos.State = adStateOpen Then rs_Procesos.Close
If rs_Periodo.State = adStateOpen Then rs_Periodo.Close
If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
If rs_Reporte.State = adStateOpen Then rs_Reporte.Close

Set rs_Confrep = Nothing
Set rs_Procesos = Nothing
Set rs_Periodo = Nothing
Set rs_Estructura = Nothing
Set rs_Reporte = Nothing

Exit Sub
CE:
    HuboError = True
    MyRollbackTrans

    If rs_Confrep.State = adStateOpen Then rs_Confrep.Close
    If rs_Procesos.State = adStateOpen Then rs_Procesos.Close
    If rs_Periodo.State = adStateOpen Then rs_Periodo.Close
    If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
    If rs_Reporte.State = adStateOpen Then rs_Reporte.Close
    
    Set rs_Confrep = Nothing
    Set rs_Procesos = Nothing
    Set rs_Periodo = Nothing
    Set rs_Estructura = Nothing
    Set rs_Reporte = Nothing
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
Dim Separador As String

Dim FechaDesde As Date
Dim FechaHasta As Date
Dim pliqdesde As Long
Dim pliqhasta As Long
Dim Todos_Pro As Boolean
Dim Proc_Aprob As Integer
Dim Empresa As Long
Dim FiltroEmpleados As String

Dim Tenro1 As Long
Dim Tenro2 As Long
Dim Tenro3 As Long
Dim Estrnro1 As Long
Dim Estrnro2 As Long
Dim Estrnro3 As Long

Dim AgrupaTE1 As Boolean
Dim AgrupaTE2 As Boolean
Dim AgrupaTE3 As Boolean
Dim Agrupado As Boolean

'Inicializacion
Agrupado = False
Tenro1 = 0
Tenro2 = 0
Tenro3 = 0
AgrupaTE1 = False
AgrupaTE2 = False
AgrupaTE3 = False

'Orden de los parametros

Separador = "@"
' Levanto cada parametro por separado
If Not IsNull(parametros) Then
    If Len(parametros) >= 1 Then
        pos1 = 1
        pos2 = InStr(pos1, parametros, Separador) - 1
        FiltroEmpleados = Mid(parametros, pos1, pos2 - pos1 + 1)
    
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        pliqdesde = CLng(Mid(parametros, pos1, pos2 - pos1 + 1))
    
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        pliqhasta = CLng(Mid(parametros, pos1, pos2 - pos1 + 1))
    
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        FechaDesde = CDate(Mid(parametros, pos1, pos2 - pos1 + 1))
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        FechaHasta = CDate(Mid(parametros, pos1, pos2 - pos1 + 1))
    
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        Proc_Aprob = Mid(parametros, pos1, pos2 - pos1 + 1)
    
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        NroProc = Mid(parametros, pos1, pos2 - pos1 + 1)
        If NroProc = "0" Then
            Todos_Pro = True
        Else
            Todos_Pro = False
        End If
        ListaNroProc = Replace(NroProc, ",", "-")
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        Tenro1 = Mid(parametros, pos1, pos2 - pos1 + 1)
        If Not Tenro1 = 0 Then
            Agrupado = True
            AgrupaTE1 = True
        End If
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        Estrnro1 = Mid(parametros, pos1, pos2 - pos1 + 1)
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        Tenro2 = Mid(parametros, pos1, pos2 - pos1 + 1)
        If Not Tenro2 = 0 Then
            AgrupaTE2 = True
        End If
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        Estrnro2 = Mid(parametros, pos1, pos2 - pos1 + 1)
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        Tenro3 = Mid(parametros, pos1, pos2 - pos1 + 1)
        If Not Tenro3 = 0 Then
            AgrupaTE3 = True
        End If
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        Estrnro3 = Mid(parametros, pos1, pos2 - pos1 + 1)
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, Separador) - 1
        Empresa = Mid(parametros, pos1, pos2 - pos1 + 1)
        Empresa = 1
    End If
End If

Call Generacion(FiltroEmpleados, bpronro, pliqdesde, Todos_Pro, Proc_Aprob, Empresa, Agrupado, AgrupaTE1, Tenro1, Estrnro1, AgrupaTE2, Tenro2, Estrnro2, AgrupaTE3, Tenro3, Estrnro3)

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



Private Sub Reprdp5()
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del informe de Evaluaciones
' Autor      : FGZ
' Fecha      : 13/09/2004
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------

'Generaci¢n de Datos para Reporte de STULICH
'El Evento debe ser RDP, mitad de ciclo o fin de ciclo.
'
'se filtra por:
'Departamento: Todos o uno seleccionado.
'Categoria: Todas o una seleccionada.
'Consejero: Todos o uno seleccionado.
'Evaluado: Todos o uno seleccionado.
'


End Sub

