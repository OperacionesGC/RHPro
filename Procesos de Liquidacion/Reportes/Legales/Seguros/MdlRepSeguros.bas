Attribute VB_Name = "MdlRepSeguros"
Option Explicit

Const Version = "1.01"
Const FechaVersion = "21/10/2009"
'Modificaciones: Manuel Lopez - Encriptacion de string connection


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
    
    Nombre_Arch = PathFLog & "Reporte_Seguros" & "-" & NroProcesoBatch & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "PID = " & PID
    'Cambio el estado del proceso a Procesando y el PID
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 44 AND bpronro =" & NroProcesoBatch
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

Public Sub Sueseg02(ByVal bpronro As Long, ByVal Nroliq As Long, ByVal Todos_Pro As Boolean, ByVal Proc_Aprob As Integer, ByVal Empresa As Long, ByVal Agrup As Integer, ByVal Agrupado As Boolean, _
    ByVal AgrupaTE1 As Boolean, ByVal Tenro1 As Long, Estrnro1 As Long, _
    ByVal AgrupaTE2 As Boolean, ByVal Tenro2 As Long, Estrnro2 As Long, _
    ByVal AgrupaTE3 As Boolean, ByVal Tenro3 As Long, Estrnro3 As Long)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del Reporte de Seguros
' Autor      : FGZ
' Fecha      : 01/04/2004
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------

Dim fechadesde As Date
Dim fechahasta As Date

Dim Arreglo(5) As Single
Dim I As Integer
Dim Ultimo_Empleado As Long
Dim Estructura As Long
Dim PrimeraVez As Boolean
Dim ColumnaConfiguracion As Boolean
Dim EncontroValor As Boolean
Dim Nro_Reporte As Integer

Dim col1 As Single
Dim col2 As Single
Dim col3 As Single

Dim nro_concnro As Long
Dim nro_tpanro As Long
Dim nro_con As Integer

Dim rs_Reporte As New ADODB.Recordset
Dim rs_Periodo As New ADODB.Recordset
Dim rs_Procesos As New ADODB.Recordset
Dim rs_Confrep As New ADODB.Recordset
Dim rs_Rep27 As New ADODB.Recordset
Dim rs_Detliq As New ADODB.Recordset
Dim rs_Concepto As New ADODB.Recordset
Dim rs_Parametro As New ADODB.Recordset
Dim rs_traza As New ADODB.Recordset
Dim rs_NovEmp As New ADODB.Recordset
Dim rs_Estructura As New ADODB.Recordset

'Inicializacion
col1 = 0
col2 = 0
col3 = 0
For I = 1 To 5
    Arreglo(I) = 0
Next I

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
Nro_Reporte = 24
'Concepto
StrSql = "SELECT * FROM confrep WHERE repnro = " & Nro_Reporte & " AND confnrocol = 1"
OpenRecordset StrSql, rs_Confrep
If rs_Confrep.EOF Then
    Flog.writeline "No se encontró la configuración del Reporte"
    Exit Sub
Else
    nro_concnro = rs_Confrep!confval
    StrSql = "SELECT * FROM concepto WHERE conccod = " & nro_concnro
    OpenRecordset StrSql, rs_Concepto
    If rs_Concepto.EOF Then
        Flog.writeline "Columna 1. El concepto no existe"
        Exit Sub
    Else
        nro_con = rs_Concepto!concnro
    End If
End If

'Parametros
StrSql = "SELECT * FROM confrep WHERE repnro = " & Nro_Reporte & " AND confnrocol = 2"
OpenRecordset StrSql, rs_Confrep
If rs_Confrep.EOF Then
    Flog.writeline "Columna 2. No se encontró la configuración del Reporte"
    Exit Sub
Else
    nro_tpanro = rs_Confrep!confval
    StrSql = "SELECT * FROM tipopar WHERE tpanro = " & nro_tpanro
    OpenRecordset StrSql, rs_Parametro
    If rs_Parametro.EOF Then
        Flog.writeline "Columna 2. El parametro no existe"
        Exit Sub
    End If
End If


' Comienzo la transaccion
MyBeginTrans

'Depuracion del Temporario
StrSql = "DELETE FROM rep27 "
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
    StrSql = StrSql & " AND tenro1 is null AND estrnro1 = 0"
    StrSql = StrSql & " AND tenro2 is null AND estrnro2 = 0"
    StrSql = StrSql & " AND tenro3 is null AND estrnro3 = 0"
End If
objConn.Execute StrSql, , adExecuteNoRecords

'Busco los procesos a evaluar
'StrSql = "SELECT * FROM  empleado "
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
'StrSql = StrSql & " INNER JOIN estructura emp ON emp.estrnro = empresa.estrnro AND emp.empnro =" & Empresa
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
StrSql = StrSql & " AND empresa.empnro =" & Empresa
StrSql = StrSql & " AND (his_estructura.htetdesde <= " & ConvFecha(Fecha_Fin_Periodo) & ") AND "
StrSql = StrSql & " ((" & ConvFecha(Fecha_Fin_Periodo) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"

If Not Todos_Pro Then
    'StrSql = StrSql & " AND proceso.pronro =" & NroProc
    StrSql = StrSql & " AND proceso.pronro IN (" & ListaNroProc & ")"
Else
    StrSql = StrSql & " AND proaprob = " & CInt(Proc_Aprob)
    StrSql = StrSql & " AND proceso.empnro = " & Empresa
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

Do While Not rs_Procesos.EOF
    
    'Col1
    StrSql = "SELECT * FROM confrep WHERE repnro = " & Nro_Reporte
    StrSql = StrSql & " AND conftipo = '" & "CO'"
    If rs_Confrep.State = adStateOpen Then rs_Confrep.Close
    OpenRecordset StrSql, rs_Confrep
    Do While Not rs_Confrep.EOF
        StrSql = "SELECT * FROM detliq " & _
                 " INNER JOIN cabliq ON detliq.cliqnro = cabliq.cliqnro " & _
                 " INNER JOIN concepto ON detliq.concnro = concepto.concnro " & _
                 " WHERE concepto.concnro = " & nro_con & _
                 " AND cabliq.cliqnro =" & rs_Procesos!cliqnro & _
                 " AND (concepto.concimp = -1" & _
                 " OR concepto.concpuente = 0)"
        OpenRecordset StrSql, rs_Detliq
        Do While Not rs_Detliq.EOF
            col1 = col1 + rs_Detliq!dlimonto
            
            rs_Detliq.MoveNext
        Loop
        
        'Siguiente confrep
        rs_Confrep.MoveNext
    Loop
        
    'Col2
    StrSql = "SELECT * FROM novemp WHERE concnro = " & nro_con & _
             " AND tpanro = " & nro_tpanro & " AND empleado = " & rs_Procesos!ternro
    OpenRecordset StrSql, rs_NovEmp
    If Not rs_NovEmp.EOF Then
        col2 = col2 + rs_NovEmp!nevalor
    End If

    'Col3
    StrSql = "SELECT * FROM traza " & _
             " WHERE cliqnro = " & rs_Procesos!cliqnro & _
             " AND concnro = " & nro_con & _
             " AND tpanro = " & nro_tpanro
    OpenRecordset StrSql, rs_traza
    If Not rs_traza.EOF Then
        col3 = col3 + rs_traza!travalor
    End If

    'Si no existe el rep27
    StrSql = "SELECT * FROM rep27 "
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
    OpenRecordset StrSql, rs_Rep27

    If rs_Rep27.EOF Then
        'Inserto
        StrSql = "INSERT INTO rep27 (bpronro,pliqnro,pronro,proaprob,empresa,iduser,fecha,hora,"
        StrSql = StrSql & "ternro,empleg,apeynom,dlimonto,tpanro1,tpanro2,"
        StrSql = StrSql & "tenro1,estrnro1,tedesc1,estrdesc1,tenro2,estrnro2,tedesc2,estrdesc2,tenro3,estrnro3,tedesc3,estrdesc3 "
        StrSql = StrSql & ") VALUES ("
        StrSql = StrSql & bpronro & ","
        StrSql = StrSql & Nroliq & ","
        If Not Todos_Pro Then
            'StrSql = StrSql & rs_Procesos!pronro & ","
            StrSql = StrSql & "'" & NroProc & "',"
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
        
        'columnas
        StrSql = StrSql & col1 & ","
        StrSql = StrSql & col2 & ","
        StrSql = StrSql & col3 & ","
        
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
        StrSql = "UPDATE rep27 SET dlimonto = dlimonto + " & col1
        StrSql = StrSql & ",tpanro1 = tpanro1 + " & col2
        StrSql = StrSql & ",tpanro2 = tpanro2 + " & col3
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
        
    'Actualizo el progreso del Proceso
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
             "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
            
    'Limpio
    col1 = 0
    col2 = 0
    col3 = 0
    
    'Siguiente proceso
    rs_Procesos.MoveNext
Loop

'Fin de la transaccion
MyCommitTrans

If rs_Procesos.State = adStateOpen Then rs_Procesos.Close
If rs_Periodo.State = adStateOpen Then rs_Periodo.Close
If rs_Confrep.State = adStateOpen Then rs_Confrep.Close
If rs_Detliq.State = adStateOpen Then rs_Detliq.Close
If rs_Rep27.State = adStateOpen Then rs_Rep27.Close
If rs_Reporte.State = adStateOpen Then rs_Reporte.Close
If rs_Concepto.State = adStateOpen Then rs_Concepto.Close
If rs_Parametro.State = adStateOpen Then rs_Parametro.Close
If rs_traza.State = adStateOpen Then rs_traza.Close
If rs_NovEmp.State = adStateOpen Then rs_NovEmp.Close
If rs_Estructura.State = adStateOpen Then rs_Estructura.Close

Set rs_Procesos = Nothing
Set rs_Confrep = Nothing
Set rs_Detliq = Nothing
Set rs_Rep27 = Nothing
Set rs_Periodo = Nothing
Set rs_Reporte = Nothing
Set rs_Concepto = Nothing
Set rs_Parametro = Nothing
Set rs_traza = Nothing
Set rs_NovEmp = Nothing
Set rs_Estructura = Nothing

Exit Sub
CE:
    HuboError = True
    MyRollbackTrans

    If rs_Procesos.State = adStateOpen Then rs_Procesos.Close
    If rs_Periodo.State = adStateOpen Then rs_Periodo.Close
    If rs_Confrep.State = adStateOpen Then rs_Confrep.Close
    If rs_Detliq.State = adStateOpen Then rs_Detliq.Close
    If rs_Rep27.State = adStateOpen Then rs_Rep27.Close
    If rs_Reporte.State = adStateOpen Then rs_Reporte.Close
    If rs_Concepto.State = adStateOpen Then rs_Concepto.Close
    If rs_Parametro.State = adStateOpen Then rs_Parametro.Close
    If rs_traza.State = adStateOpen Then rs_traza.Close
    If rs_NovEmp.State = adStateOpen Then rs_NovEmp.Close
    If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
    
    Set rs_Procesos = Nothing
    Set rs_Confrep = Nothing
    Set rs_Detliq = Nothing
    Set rs_Rep27 = Nothing
    Set rs_Periodo = Nothing
    Set rs_Reporte = Nothing
    Set rs_Concepto = Nothing
    Set rs_Parametro = Nothing
    Set rs_traza = Nothing
    Set rs_NovEmp = Nothing
    Set rs_Estructura = Nothing
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
Dim Todos_Sindicatos As Boolean
Dim Nro_Sindicato As Long
Dim Agrup As Integer
Dim TextoAgrupado As String

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

' Levanto cada parametro por separado, el separador de parametros es "."
If Not IsNull(parametros) Then
    If Len(parametros) >= 1 Then
        
        pos1 = 1
        pos2 = InStr(pos1, parametros, ".") - 1
        Nroliq = CLng(Mid(parametros, pos1, pos2 - pos1 + 1))
    
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, ".") - 1
        Todos_Pro = CBool(Mid(parametros, pos1, pos2 - pos1 + 1))
        
        If Not Todos_Pro Then
            pos1 = pos2 + 2
            pos2 = InStr(pos1, parametros, ".") - 1
            NroProc = Mid(parametros, pos1, pos2 - pos1 + 1)
            ListaNroProc = Replace(NroProc, "-", ",")
        Else
            NroProc = "0"
            ListaNroProc = "0"
        End If
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, ".") - 1
        Proc_Aprob = Mid(parametros, pos1, pos2 - pos1 + 1)
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, ".") - 1
        If Not pos2 > 0 Then
            pos2 = Len(parametros)
            Empresa = Mid(parametros, pos1, pos2 - pos1 + 1)
            Agrup = 0
        Else
        
            Empresa = Mid(parametros, pos1, pos2 - pos1 + 1)
            
'            pos1 = pos2 + 2
'            pos2 = InStr(pos1, parametros, ".") - 1
'            Agrup = Mid(parametros, pos1, pos2 - pos1 + 1)
            
'            Select Case Agrup
'            Case 1:
                'A continuacion pueden venir hasta tres niveles de agrupamiento
                ' cero,uno, dos o tres niveles
                pos1 = pos2 + 2
                pos2 = InStr(pos1, parametros, ".") - 1
                If pos2 > 0 Then
                    Agrupado = True
                    AgrupaTE1 = True
                    Tenro1 = Mid(parametros, pos1, pos2 - pos1 + 1)
                
                    pos1 = pos2 + 2
                    pos2 = InStr(pos1, parametros, ".") - 1
                    If Not (pos2 > 0) Then
                        pos2 = Len(parametros)
                    End If
                    Estrnro1 = Mid(parametros, pos1, pos2 - pos1 + 1)
        
                    pos1 = pos2 + 2
                    pos2 = InStr(pos1, parametros, ".") - 1
                    If pos2 > 0 Then
                        AgrupaTE2 = True
                        Tenro2 = Mid(parametros, pos1, pos2 - pos1 + 1)
                    
                        pos1 = pos2 + 2
                        pos2 = InStr(pos1, parametros, ".") - 1
                        If Not (pos2 > 0) Then
                            pos2 = Len(parametros)
                        End If
                        Estrnro2 = Mid(parametros, pos1, pos2 - pos1 + 1)
                        
                        pos1 = pos2 + 2
                        pos2 = InStr(pos1, parametros, ".") - 1
                        If pos2 > 0 Then
                            AgrupaTE3 = True
                            Tenro3 = Mid(parametros, pos1, pos2 - pos1 + 1)
                        
                            pos1 = pos2 + 2
                            pos2 = Len(parametros)
                            Estrnro3 = Mid(parametros, pos1, pos2 - pos1 + 1)
                        End If
                    End If
                End If
'            End Select
        End If
    End If
End If


Call Sueseg02(bpronro, Nroliq, Todos_Pro, Proc_Aprob, Empresa, Agrup, Agrupado, AgrupaTE1, Tenro1, Estrnro1, AgrupaTE2, Tenro2, Estrnro2, AgrupaTE3, Tenro3, Estrnro3)

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

