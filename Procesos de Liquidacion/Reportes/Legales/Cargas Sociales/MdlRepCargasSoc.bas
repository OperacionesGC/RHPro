Attribute VB_Name = "MdlRepCargasSoc"
Global Const Version = "1.01" ' Cesar Stankunas
Global Const FechaModificacion = "06/08/2009"
Global Const UltimaModificacion = ""    'Encriptacion de string connection

Option Explicit

Global IdUser As String
Global Fecha As Date
Global Hora As String

Global Fecha_Inicio_periodo As Date
Global Fecha_Fin_Periodo As Date
Global StrSql2 As String



Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial del Generador de Reporte de Aportes y Contribuciones.
' Autor      : FGZ
' Fecha      : 17/02/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim objconnMain As New ADODB.Connection
Dim strCmdLine
Dim Nombre_Arch As String
Dim HuboError As Boolean
Dim rs_batch_proceso As New ADODB.Recordset
Dim bprcparam As String
Dim PID As String
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
    
    Nombre_Arch = PathFLog & "Reporte_CargasSociales" & "-" & NroProcesoBatch & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "PID = " & PID
    
    'Abro la conexion
    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    
    On Error Resume Next
    OpenConnection strconexion, objconnProgreso
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    
    'Cambio el estado del proceso a Procesando y el PID
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 32 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        bprcparam = rs_batch_proceso!bprcparam
        IdUser = rs_batch_proceso!IdUser
        Fecha = rs_batch_proceso!bprcfecha
        Hora = rs_batch_proceso!bprchora
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        Call Sueayc02(NroProcesoBatch, bprcparam)
    End If
    
    TiempoFinalProceso = GetTickCount
    Flog.writeline "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    
    If Not HuboError Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
    End If
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    objConn.Close
    objconnProgreso.Close
    
    Set objConn = Nothing
    Set objconnProgreso = Nothing
    
Flog.Close

End Sub


Public Sub Sueayc02(ByVal bpronro As Long, ByVal Parametros As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del Reporte de Aportes y Contribuciones
' Autor      : FGZ
' Fecha      : 17/02/2004
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim pos1, pos2 As Integer

Dim Nroliq As Long
Dim Todos_Pro As Boolean
Dim Proc_Aprob As Integer
Dim Empresa As Long

Dim fechadesde As Date
Dim fechahasta As Date

Dim selector  As Integer
Dim selector2 As Integer
Dim selector3 As Integer
Dim X As Boolean

Dim rs_Reporte As New ADODB.Recordset
Dim rs_Periodo As New ADODB.Recordset
Dim rs_Procesos As New ADODB.Recordset
Dim rs_Confrep As New ADODB.Recordset
Dim rs_Detliq As New ADODB.Recordset
Dim rs_Rep08 As New ADODB.Recordset
Dim rs_Empleados As New ADODB.Recordset

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

Dim rs_Estructura As New ADODB.Recordset

'Inicializacion
Agrupado = False
Tenro1 = 0
Tenro2 = 0
Tenro3 = 0
AgrupaTE1 = False
AgrupaTE2 = False
AgrupaTE3 = False

' Levanto cada parametro por separado, el separador de parametros es "."
If Not IsNull(Parametros) Then
    If Len(Parametros) >= 1 Then
        pos1 = 1
        pos2 = InStr(pos1, Parametros, ".") - 1
        Nroliq = CLng(Mid(Parametros, pos1, pos2))
    
        pos1 = pos2 + 2
        pos2 = InStr(pos1, Parametros, ".") - 1
        Todos_Pro = CBool(Mid(Parametros, pos1, pos2 - pos1 + 1))
        
        If Not Todos_Pro Then
            pos1 = pos2 + 2
            pos2 = InStr(pos1, Parametros, ".") - 1
            'NroProc = CLng(Mid(Parametros, pos1, pos2 - pos1 + 1))
            NroProc = Mid(Parametros, pos1, pos2 - pos1 + 1)
            ListaNroProc = Replace(NroProc, "-", ",")
        Else
            'NroProc = 0
            NroProc = "0"
            ListaNroProc = "0"
        End If
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, Parametros, ".") - 1
        Proc_Aprob = Mid(Parametros, pos1, pos2 - pos1 + 1)
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, Parametros, ".") - 1
        If pos2 > 0 Then
            Empresa = Mid(Parametros, pos1, pos2 - pos1 + 1)
        Else
            pos2 = Len(Parametros)
            Empresa = Mid(Parametros, pos1, pos2 - pos1 + 1)
        End If
        
       
        'A continuacion pueden venir hasta tres niveles de agrupamiento
        ' cero,uno, dos o tres niveles
        pos1 = pos2 + 2
        pos2 = InStr(pos1, Parametros, ".") - 1
        If pos2 > 0 Then
            Agrupado = True
            AgrupaTE1 = True
            Tenro1 = Mid(Parametros, pos1, pos2 - pos1 + 1)
        
            pos1 = pos2 + 2
            pos2 = InStr(pos1, Parametros, ".") - 1
            If Not (pos2 > 0) Then
                pos2 = Len(Parametros)
            End If
            Estrnro1 = Mid(Parametros, pos1, pos2 - pos1 + 1)

            pos1 = pos2 + 2
            pos2 = InStr(pos1, Parametros, ".") - 1
            If pos2 > 0 Then
                AgrupaTE2 = True
                Tenro2 = Mid(Parametros, pos1, pos2 - pos1 + 1)
            
                pos1 = pos2 + 2
                pos2 = InStr(pos1, Parametros, ".") - 1
                If Not (pos2 > 0) Then
                    pos2 = Len(Parametros)
                End If
                Estrnro2 = Mid(Parametros, pos1, pos2 - pos1 + 1)
                
                pos1 = pos2 + 2
                pos2 = InStr(pos1, Parametros, ".") - 1
                If pos2 > 0 Then
                    AgrupaTE3 = True
                    Tenro3 = Mid(Parametros, pos1, pos2 - pos1 + 1)
                
                    pos1 = pos2 + 2
                    pos2 = Len(Parametros)
                    Estrnro3 = Mid(Parametros, pos1, pos2 - pos1 + 1)
                End If
            End If
        End If
    End If
End If


StrSql = "Select * FROM reporte where reporte.repnro = 66 "
OpenRecordset StrSql, rs_Reporte
If rs_Reporte.EOF Then
    Flog.writeln "El Reporte Numero 66 no ha sido Configurado"
    Exit Sub
End If
If rs_Reporte.State = adStateOpen Then rs_Reporte.Close

'Configuracion del Reporte
StrSql = "SELECT * FROM confrep WHERE repnro = 66 AND conftipo = 'TCO'"
OpenRecordset StrSql, rs_Confrep

If rs_Confrep.EOF Then
    Flog.writeline "No se encontró la configuración del Reporte"
    Exit Sub
End If

'cargo el periodo
StrSql = "SELECT * FROM periodo WHERE pliqnro = " & CStr(Nroliq)
OpenRecordset StrSql, rs_Periodo
If rs_Periodo.EOF Then
    Flog.writeline "No se encontró el Periodo"
    Exit Sub
Else
    Fecha_Inicio_periodo = rs_Periodo!pliqdesde
    Fecha_Fin_Periodo = rs_Periodo!pliqhasta
End If
If rs_Periodo.State = adStateOpen Then rs_Periodo.Close

' Comienzo la transaccion
MyBeginTrans

'Depuracion del Temporario
StrSql = "DELETE FROM rep08 "
StrSql = StrSql & " WHERE pliqnro = " & Nroliq
If Not Todos_Pro Then
    StrSql = StrSql & " AND pronro = '" & NroProc & "'"
Else
    StrSql = StrSql & " AND pronro = '0'"
    StrSql = StrSql & " AND proaprob = " & CInt(Proc_Aprob)
End If
StrSql = StrSql & " AND empresa = " & Empresa
If Agrupado Then
    StrSql = StrSql & " AND tenro1 = " & Tenro1 & " AND estrnro1 = " & Estrnro1
    If Tenro2 <> 0 Then
        StrSql = StrSql & " AND tenro2 = " & Tenro2 & " AND estrnro2 = " & Estrnro2
        If Tenro3 <> 0 Then
            StrSql = StrSql & " AND tenro3 = " & Tenro3 & " AND estrnro3 = " & Estrnro3
        End If
    End If
Else
    StrSql = StrSql & " AND tenro1 is null AND estrnro1 is null"
    StrSql = StrSql & " AND tenro2 is null AND estrnro2 is null"
    StrSql = StrSql & " AND tenro3 is null AND estrnro3 is null"
End If
objConn.Execute StrSql, , adExecuteNoRecords


'Busco los procesos a evaluar
StrSql = "SELECT cabliq.*, proceso.*, periodo.*, empleado.*"
If AgrupaTE1 Then
    StrSql = StrSql & ", te1.tenro tenro1, te1.estrnro estrnro1"
End If
If AgrupaTE2 Then
    StrSql = StrSql & ", te2.tenro tenro2, te2.estrnro estrnro2"
End If
If AgrupaTE3 Then
    StrSql = StrSql & ", te3.tenro tenro3, te3.estrnro estrnro3"
End If
StrSql = StrSql & "  FROM  empleado "
If AgrupaTE1 Then
    StrSql = StrSql & " INNER JOIN his_estructura TE1 ON te1.ternro = empleado.ternro "
End If
If AgrupaTE2 Then
    StrSql = StrSql & " INNER JOIN his_estructura TE2 ON te2.ternro = empleado.ternro "
End If
If AgrupaTE3 Then
    StrSql = StrSql & " INNER JOIN his_estructura TE3 ON te3.ternro = empleado.ternro "
End If
StrSql = StrSql & " INNER JOIN cabliq ON cabliq.empleado = empleado.ternro "
StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro "
StrSql = StrSql & " INNER JOIN his_estructura  empresa ON empresa.ternro = empleado.ternro and empresa.tenro = 10 "
StrSql = StrSql & " INNER JOIN estructura emp ON emp.estrnro = empresa.estrnro AND emp.empnro =" & Empresa
StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro WHERE "
If AgrupaTE1 Then
    StrSql = StrSql & " te1.tenro = " & Tenro1 & " AND "
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
If Not Agrupado Then
    StrSql = StrSql & " periodo.pliqnro =" & Nroliq
Else
    StrSql = StrSql & " AND periodo.pliqnro =" & Nroliq
End If
'StrSql = StrSql & " AND periodo.empnro =" & Empresa
If Not Todos_Pro Then
    StrSql = StrSql & " AND proceso.pronro IN (" & ListaNroProc & ")"
Else
    StrSql = StrSql & " AND proaprob = " & CInt(Proc_Aprob)
    StrSql = StrSql & " AND proceso.empnro = " & Empresa
End If
StrSql = StrSql & " ORDER BY empleado.ternro"
OpenRecordset StrSql, rs_Procesos


'seteo de las variables de progreso
Progreso = 0
CConceptosAProc = rs_Procesos.RecordCount
CEmpleadosAProc = rs_Confrep.RecordCount
If CEmpleadosAProc = 0 Then
   CEmpleadosAProc = 1
End If
If CEmpleadosAProc = 0 Then
    CEmpleadosAProc = 1
End If
If CConceptosAProc = 0 Then
    CConceptosAProc = 1
End If
IncPorc = ((100 / CEmpleadosAProc) * (100 / CConceptosAProc)) / 100
    
Do While Not rs_Procesos.EOF
    rs_Confrep.MoveFirst
    
    Do While Not rs_Confrep.EOF
        StrSql = "SELECT * FROM detliq " & _
                 " INNER JOIN cabliq ON detliq.cliqnro = cabliq.cliqnro " & _
                 " INNER JOIN concepto ON detliq.concnro = concepto.concnro " & _
                 " WHERE concepto.concpuente = -1 " & _
                 " AND concepto.tconnro =" & rs_Confrep!confval & _
                 " AND cabliq.cliqnro =" & rs_Procesos!cliqnro
        OpenRecordset StrSql, rs_Detliq
    
        If Not rs_Detliq.EOF Then
            'Si no existe el rep08
            StrSql = "SELECT * FROM rep08 "
            StrSql = StrSql & " WHERE concnro = " & rs_Detliq!concnro
            StrSql = StrSql & " AND bpronro = " & bpronro
'                StrSql = StrSql & " AND pliqnro = " & Nroliq
'                StrSql = StrSql & " AND empresa = " & Empresa
'                StrSql = StrSql & " AND pronro = " & NroProc
'                StrSql = StrSql & " AND proaprob= " & CInt(Proc_Aprob)
            OpenRecordset StrSql, rs_Rep08
        
            If rs_Rep08.EOF Then
                'Inserto
                StrSql = "INSERT INTO rep08 (bpronro,pliqnro,pronro,proaprob,tenro1,estrnro1,tedesc1,estrdesc1,"
                StrSql = StrSql & "tenro2,estrnro2,tedesc2,estrdesc2,tenro3,estrnro3,tedesc3,estrdesc3,empresa,iduser,fecha,hora,"
                StrSql = StrSql & "concnro,total_liquidado,cant_liquidado,"
                StrSql = StrSql & "emp_liquidado,asigfam) VALUES ("
                StrSql = StrSql & bpronro & ","
                StrSql = StrSql & Nroliq & ","
                If Not Todos_Pro Then
                    'StrSql = StrSql & rs_Procesos!pronro & ","
                    StrSql = StrSql & "'" & NroProc & "',"
                    StrSql = StrSql & rs_Procesos!proaprob & ","
                Else
                    StrSql = StrSql & "'0'" & ","
                    StrSql = StrSql & CInt(Proc_Aprob) & ","
                End If
                
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
                        StrSql = StrSql & "'" & rs_Estructura!estrdabr & "'" '& ","
                    Else
                        StrSql = StrSql & "' '" & ","
                    End If
                Else
                    StrSql = StrSql & "' '" & ","
                    StrSql = StrSql & "' '"
                End If
                StrSql = StrSql & ","
                
                StrSql = StrSql & Empresa & ","
                StrSql = StrSql & "'" & IdUser & "',"
                StrSql = StrSql & ConvFecha(Fecha) & ","
                StrSql = StrSql & "'" & Format(Hora, "hh:mm:ss") & "',"
                StrSql = StrSql & rs_Detliq!concnro & ","
                StrSql = StrSql & "0,0,0,0)"
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
            
            'Actualizo
            StrSql = "UPDATE rep08 SET total_liquidado = total_liquidado + " & rs_Detliq!dlimonto
            StrSql = StrSql & ", cant_liquidado = cant_liquidado + " & rs_Detliq!dlicant
            StrSql = StrSql & ", emp_liquidado = emp_liquidado + 1 "
            StrSql = StrSql & " WHERE concnro = " & rs_Detliq!concnro
            StrSql = StrSql & " AND bpronro = " & bpronro
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
            
        'Actualizo el progreso del Proceso
        Progreso = Progreso + IncPorc
        TiempoAcumulado = GetTickCount
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
                 ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
                 "' WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
            
        'Siguiente confrep
        rs_Confrep.MoveNext
    Loop
       
    'Siguiente proceso
    rs_Procesos.MoveNext
Loop


'Actualizo las asignaciones familiares
StrSql = "SELECT * FROM confrep WHERE repnro = 66 AND conftipo = 'TCO' AND confnrocol = 1"
OpenRecordset StrSql, rs_Confrep

If Not rs_Confrep.EOF Then
    StrSql = "SELECT * FROM rep08 "
    StrSql = StrSql & " INNER JOIN concepto ON rep08.concnro = concepto.concnro "
    StrSql = StrSql & " WHERE rep08.bpronro = " & bpronro
    StrSql = StrSql & " AND concepto.tconnro = " & rs_Confrep!confval
    OpenRecordset StrSql, rs_Rep08
    
    Do While Not rs_Rep08.EOF
        StrSql = "UPDATE rep08 SET asigfam = " & rs_Rep08!total_liquidado
        StrSql = StrSql & " WHERE concnro = " & rs_Rep08!concnro
        StrSql = StrSql & " AND bpronro = " & rs_Rep08!bpronro
        objConn.Execute StrSql, , adExecuteNoRecords
    
        rs_Rep08.MoveNext
    Loop
End If


'Fin de la transaccion
MyCommitTrans


If rs_Empleados.State = adStateOpen Then rs_Empleados.Close
If rs_Procesos.State = adStateOpen Then rs_Procesos.Close
If rs_Periodo.State = adStateOpen Then rs_Periodo.Close
If rs_Confrep.State = adStateOpen Then rs_Confrep.Close
If rs_Detliq.State = adStateOpen Then rs_Detliq.Close
If rs_Rep08.State = adStateOpen Then rs_Rep08.Close
If rs_Reporte.State = adStateOpen Then rs_Reporte.Close

Set rs_Empleados = Nothing
Set rs_Procesos = Nothing
Set rs_Confrep = Nothing
Set rs_Detliq = Nothing
Set rs_Rep08 = Nothing
Set rs_Periodo = Nothing
Set rs_Reporte = Nothing


Exit Sub
CE:
    HuboError = True
    MyRollbackTrans

End Sub


Public Function EsElUltimoEmpleado(ByVal rs As ADODB.Recordset, ByVal Anterior As Long) As Boolean
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

