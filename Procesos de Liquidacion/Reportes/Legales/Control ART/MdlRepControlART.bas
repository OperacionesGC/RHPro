Attribute VB_Name = "MdlRepControlART"
Option Explicit

''Global Const Version = "1.01" ' Cesar Stankunas
''Global Const FechaModificacion = "06/08/2009"
''Global Const UltimaModificacion = ""    'Encriptacion de string connection

Global Const Version = "1.02"
Global Const FechaModificacion = "28/04/2011"
Global Const UltimaModificacion = "FGZ" 'Agregados de log y correccion del query que busca los empleados de la empresa en el periodo
Global Const UltimaModificacion1 = " "
Global Const UltimaModificacion2 = " "


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
    
    Nombre_Arch = PathFLog & "Reporte_Control_ART" & "-" & NroProcesoBatch & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Version = " & Version
    Flog.writeline "Modificacion = " & UltimaModificacion
    Flog.writeline "Modificacion = " & UltimaModificacion1
    Flog.writeline "Modificacion = " & UltimaModificacion2
    Flog.writeline "Fecha        = " & FechaModificacion
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline
    Flog.writeline "PID = " & PID
    
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
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 42 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        IdUser = rs_batch_proceso!IdUser
        Fecha = rs_batch_proceso!bprcfecha
        Hora = rs_batch_proceso!bprchora
        bprcparam = rs_batch_proceso!bprcparam
        rs_batch_proceso.Close
        
        Call LevantarParamteros(NroProcesoBatch, bprcparam)
    End If
    Set rs_batch_proceso = Nothing
    
    TiempoFinalProceso = GetTickCount
    Flog.writeline "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    
    If Not HuboError Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
    End If
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Set objConn = Nothing
    Set objconnProgreso = Nothing
    
    Flog.Close

End Sub


Public Sub Sueacu11(ByVal bpronro As Long, ByVal Nroliq As Long, ByVal Todos_Pro As Boolean, ByVal Proc_Aprob As Integer, ByVal Empresa As Long, ByVal Agrupado As Boolean, _
    ByVal AgrupaTE1 As Boolean, ByVal Tenro1 As Long, Estrnro1 As Long, _
    ByVal AgrupaTE2 As Boolean, ByVal Tenro2 As Long, Estrnro2 As Long, _
    ByVal AgrupaTE3 As Boolean, ByVal Tenro3 As Long, Estrnro3 As Long)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del Reporte de Sindicato
' Autor      : FGZ
' Fecha      : 02/03/2004
' Ult. Mod   : 02/09/2005 - Fapitalle N. - Cambio en la indexacion de las columnas
' Fecha      :
' --------------------------------------------------------------------------------------------

Dim fechadesde As Date
Dim FechaHasta As Date

Dim Arreglo(20) As Single
Dim I As Integer
Dim Columna As Integer
Dim Ultimo_Empleado As Long
Dim Estructura As Long
Dim PrimeraVez As Boolean

Dim rs_Reporte As New ADODB.Recordset
Dim rs_Periodo As New ADODB.Recordset
Dim rs_Procesos As New ADODB.Recordset
Dim rs_Confrep As New ADODB.Recordset
Dim rs_Rep94 As New ADODB.Recordset
Dim rs_Detliq As New ADODB.Recordset
Dim rs_Acu_Liq As New ADODB.Recordset
Dim rs_Categoria As New ADODB.Recordset
Dim rs_Ter_Doc As New ADODB.Recordset
Dim rs_Estructura As New ADODB.Recordset
Dim EncontroAlgo As Boolean
'Inicializacion
For I = 1 To 20
    Arreglo(I) = 0
Next I

EncontroAlgo = False
'cargo el periodo
StrSql = "SELECT * FROM periodo WHERE pliqnro = " & CStr(Nroliq)
OpenRecordset StrSql, rs_Periodo
If rs_Periodo.EOF Then
    Flog.writeline "No se encontró el Periodo"
    Exit Sub
End If
Fecha_Inicio_periodo = rs_Periodo!pliqdesde
Fecha_Fin_Periodo = rs_Periodo!pliqhasta

StrSql = "Select * FROM reporte where reporte.repnro = 68"
OpenRecordset StrSql, rs_Reporte
If rs_Reporte.EOF Then
    Flog.writeln "El Reporte Numero 68 no ha sido Configurado"
    Exit Sub
End If
If rs_Reporte.State = adStateOpen Then rs_Reporte.Close


' Comienzo la transaccion
MyBeginTrans

'Depuracion del Temporario
StrSql = "DELETE FROM rep94 "
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
'FGZ - 28/04/2011 -------------
'StrSql = StrSql & " INNER JOIN his_estructura  empresa ON empresa.ternro = empleado.ternro and empresa.tenro = 10 AND ((" & ConvFecha(Fecha_Fin_Periodo) & " <= empresa.htethasta) or (empresa.htethasta is null)) AND (empresa.htetdesde <= " & ConvFecha(Fecha_Fin_Periodo) & ") "
StrSql = StrSql & " INNER JOIN his_estructura  hempresa ON hempresa.ternro = empleado.ternro and hempresa.tenro = 10 AND ((" & ConvFecha(Fecha_Fin_Periodo) & " <= hempresa.htethasta) or (hempresa.htethasta is null)) AND (hempresa.htetdesde <= " & ConvFecha(Fecha_Fin_Periodo) & ") "
'StrSql = StrSql & " INNER JOIN estructura emp ON emp.estrnro = empresa.estrnro AND emp.empnro =" & Empresa
StrSql = StrSql & " INNER JOIN empresa emp ON emp.estrnro = hempresa.estrnro AND emp.empnro =" & Empresa
'FGZ - 28/04/2011 -------------
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
If Not Todos_Pro Then
    'StrSql = StrSql & " AND proceso.pronro =" & NroProc
    StrSql = StrSql & " AND proceso.pronro IN (" & ListaNroProc & ")"
Else
    StrSql = StrSql & " AND proaprob = " & CInt(Proc_Aprob)
    StrSql = StrSql & " AND proceso.empnro = " & Empresa
End If
StrSql = StrSql & " ORDER BY empleado.ternro"
OpenRecordset StrSql, rs_Procesos
If rs_Procesos.EOF Then
    Flog.writeline "No se encontró empleados a procesar."
    Flog.writeline StrSql
End If


'Configuracion del Reporte
StrSql = "SELECT * FROM confrep WHERE repnro = 68 AND ((conftipo = 'CO') OR (conftipo = 'AC'))"
OpenRecordset StrSql, rs_Confrep

If rs_Confrep.EOF Then
    Flog.writeline "No se encontró la configuración del Reporte"
    Exit Sub
End If

'seteo de las variables de progreso
Progreso = 0
CConceptosAProc = rs_Procesos.RecordCount
CEmpleadosAProc = rs_Confrep.RecordCount
If CConceptosAProc = 0 Then
    CConceptosAProc = 1
End If
If CEmpleadosAProc = 0 Then
   CEmpleadosAProc = 1
End If
IncPorc = ((100 / CEmpleadosAProc) * (100 / CConceptosAProc)) / 100

Ultimo_Empleado = -1
Do While Not rs_Procesos.EOF
    If Ultimo_Empleado = -1 Then 'Es el primer ternro
        For I = 1 To 20
            Arreglo(I) = 0
        Next I
    End If
    
    Ultimo_Empleado = rs_Procesos!Ternro
    
        rs_Confrep.MoveFirst
        Columna = 1 'para indexar las columnas AC y CO nada mas
        Do While Not rs_Confrep.EOF
            Select Case UCase(rs_Confrep!conftipo)
            Case "AC":
                StrSql = "SELECT * FROM acu_liq " & _
                         " INNER JOIN cabliq ON acu_liq.cliqnro = cabliq.cliqnro " & _
                         " WHERE acu_liq.acunro = " & rs_Confrep!confval & _
                         " AND cabliq.cliqnro =" & rs_Procesos!cliqnro
                OpenRecordset StrSql, rs_Acu_Liq
                Do While Not rs_Acu_Liq.EOF
'                    Arreglo(rs_Confrep!confnrocol) = Arreglo(rs_Confrep!confnrocol) + rs_Acu_Liq!almonto
'                    Arreglo(Columna) = Arreglo(Columna) + rs_Acu_Liq!almonto
                    If rs_Acu_Liq!almonto <> 0 Then
                        Arreglo(Columna) = Arreglo(Columna) + rs_Acu_Liq!almonto
                        EncontroAlgo = True
                        Columna = Columna + 1
                    End If
                    rs_Acu_Liq.MoveNext
                Loop
                
            Case "CO":
                StrSql = "SELECT * FROM detliq " & _
                         " INNER JOIN cabliq ON detliq.cliqnro = cabliq.cliqnro " & _
                         " INNER JOIN concepto ON detliq.concnro = concepto.concnro " & _
                         " WHERE concepto.conccod = " & rs_Confrep!confval & _
                         " AND cabliq.cliqnro =" & rs_Procesos!cliqnro
                OpenRecordset StrSql, rs_Detliq
                Do While Not rs_Detliq.EOF
'                    Arreglo(rs_Confrep!confnrocol) = Arreglo(rs_Confrep!confnrocol) + rs_Detliq!dlimonto
'                    Arreglo(Columna) = Arreglo(Columna) + rs_Detliq!dlimonto
                    If rs_Detliq!dlimonto <> 0 Then
                        Arreglo(Columna) = Arreglo(Columna) + rs_Detliq!dlimonto
                        EncontroAlgo = True
                        Columna = Columna + 1
                    End If
                    
                    rs_Detliq.MoveNext
                Loop
            Case Else
            End Select
            
                
            'Siguiente confrep
            rs_Confrep.MoveNext
        Loop
                
        If EsElUltimoEmpleado(rs_Procesos, Ultimo_Empleado) Then
            
                'Si no existe el rep94
                StrSql = "SELECT * FROM rep94 "
                StrSql = StrSql & " WHERE ternro = " & rs_Procesos!Ternro
                StrSql = StrSql & " AND bpronro = " & bpronro
                StrSql = StrSql & " AND pliqnro = " & Nroliq
                StrSql = StrSql & " AND empresa = " & Empresa
                If Not Todos_Pro Then
                    StrSql = StrSql & " AND pronro = '" & NroProc & "'"
                Else
                    StrSql = StrSql & " AND pronro = '0'"
                    StrSql = StrSql & " AND proaprob= " & CInt(Proc_Aprob)
                End If
                OpenRecordset StrSql, rs_Rep94
            
                If rs_Rep94.EOF Then
                
                    'Inserto
                    StrSql = "INSERT INTO rep94 (bpronro,pliqnro,pronro,proaprob,empresa,iduser,fecha,hora,"
                    StrSql = StrSql & "ternro,empleg,apenom,catdesc,cuil,tenro1,estrnro1,tedesc1,estrdesc1,tenro2,estrnro2,tedesc2,estrdesc2,tenro3,estrnro3,tedesc3,estrdesc3"
                    For I = 1 To 20
                        If Arreglo(I) <> 0 And Not IsNull(Arreglo(I)) Then
                            StrSql = StrSql & ",col" & CStr(I)
                        End If
                    Next I
                    StrSql = StrSql & ") VALUES ("
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
                    StrSql = StrSql & Empresa & ","
                    StrSql = StrSql & "'" & IdUser & "',"
                    StrSql = StrSql & ConvFecha(Fecha) & ","
                    StrSql = StrSql & "'" & Format(Hora, "hh:mm:ss") & "',"
                    
                    ' busco los datos del empleado
                    'Tercero
                        StrSql = StrSql & rs_Procesos!Ternro & ","
                    'Legajo
                        StrSql = StrSql & rs_Procesos!empleg & ","
                    'Apellido y nombre
                        StrSql = StrSql & "'" & rs_Procesos!terape
                        If Not IsNull(rs_Procesos!terape2) Then
                            StrSql = StrSql & " " & rs_Procesos!terape2
                        End If
                        StrSql = StrSql & ", " & rs_Procesos!ternom
                        If Not IsNull(rs_Procesos!ternom2) Then
                            StrSql = StrSql & " " & rs_Procesos!ternom2
                        End If
                        StrSql = StrSql & "'"
                        
                    'Categoria
                    StrSql2 = "SELECT * FROM His_estructura "
                    StrSql2 = StrSql2 & " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro "
                    StrSql2 = StrSql2 & " WHERE his_estructura.tenro = 3 AND "
                    StrSql2 = StrSql2 & " (his_estructura.htetdesde <= " & ConvFecha(Fecha_Fin_Periodo) & ") AND "
                    StrSql2 = StrSql2 & " ((" & ConvFecha(Fecha_Fin_Periodo) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
                    StrSql2 = StrSql2 & " AND his_estructura.ternro =" & rs_Procesos!Ternro
                    OpenRecordset StrSql2, rs_Categoria
                    If Not rs_Categoria.EOF Then
                        StrSql = StrSql & "," & "'" & rs_Categoria!estrdabr & "'"
                    Else
                        StrSql = StrSql & "," & "' '"
                    End If
                    StrSql = StrSql & ","
                    
                    'CUIL
                    StrSql2 = "SELECT * FROM ter_doc WHERE ternro =" & rs_Procesos!Ternro & _
                             " AND tidnro = 10"
                    OpenRecordset StrSql2, rs_Ter_Doc
                    If Not rs_Ter_Doc.EOF Then
                        StrSql = StrSql & "'" & rs_Ter_Doc!nrodoc & "',"
                    Else
                        StrSql = StrSql & "' '" & ","
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
                            StrSql = StrSql & "'" & rs_Estructura!estrdabr & "'"
                        Else
                            StrSql = StrSql & "' '"
                        End If
                    Else
                        StrSql = StrSql & "' '" & ","
                        StrSql = StrSql & "' '" '& ","
                    End If
                    
'                    If Not EncontroAlgo Then
'                        StrSql = Mid(StrSql, 1, Len(StrSql) - 1)
'                    End If
                    'Los demas datos
                    For I = 1 To 20
                        If Arreglo(I) <> 0 And Not IsNull(Arreglo(I)) Then
                            StrSql = StrSql & "," & Arreglo(I)
                        End If
                    Next I
                    StrSql = StrSql & ")"
                Else
                    'Actualizo
                    StrSql = "UPDATE rep94 "
                    PrimeraVez = True
                    For I = 1 To 20
                        If Arreglo(I) <> 0 And Not IsNull(Arreglo(I)) Then
                            If PrimeraVez Then
                                StrSql = StrSql & "SET "
                                PrimeraVez = False
                            Else
                                StrSql = StrSql & ", "
                            End If
                            StrSql = StrSql & "col" & CStr(I) & "= col" & CStr(I) & "+ " & Arreglo(I)
                        End If
                    Next I
                    StrSql = StrSql & " WHERE ternro = " & rs_Procesos!Ternro
                    StrSql = StrSql & " AND bpronro = " & bpronro
                    StrSql = StrSql & " AND pliqnro = " & Nroliq
                    StrSql = StrSql & " AND empresa = " & Empresa
                    If Not Todos_Pro Then
                        StrSql = StrSql & " AND pronro = '" & NroProc & "'"
                    Else
                        StrSql = StrSql & " AND pronro = '0'"
                        StrSql = StrSql & " AND proaprob= " & CInt(Proc_Aprob)
                    End If
                End If
                objConn.Execute StrSql, , adExecuteNoRecords
                
                Ultimo_Empleado = -1
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

'Fin de la transaccion
MyCommitTrans

If rs_Procesos.State = adStateOpen Then rs_Procesos.Close
If rs_Periodo.State = adStateOpen Then rs_Periodo.Close
If rs_Confrep.State = adStateOpen Then rs_Confrep.Close
If rs_Detliq.State = adStateOpen Then rs_Detliq.Close
If rs_Acu_Liq.State = adStateOpen Then rs_Acu_Liq.Close
If rs_Rep94.State = adStateOpen Then rs_Rep94.Close
If rs_Reporte.State = adStateOpen Then rs_Reporte.Close
If rs_Ter_Doc.State = adStateOpen Then rs_Ter_Doc.Close
If rs_Categoria.State = adStateOpen Then rs_Categoria.Close
 
 
Set rs_Procesos = Nothing
Set rs_Confrep = Nothing
Set rs_Detliq = Nothing
Set rs_Acu_Liq = Nothing
Set rs_Rep94 = Nothing
Set rs_Periodo = Nothing
Set rs_Reporte = Nothing
Set rs_Ter_Doc = Nothing
Set rs_Categoria = Nothing

Exit Sub
CE:
    HuboError = True
    MyRollbackTrans
    Flog.writeline Err.Description
    
    Flog.writeline "Ultimo SQL ejecutado: " & StrSql
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
        Nroliq = CLng(Mid(parametros, pos1, pos2))
    
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, ".") - 1
        Todos_Pro = CBool(Mid(parametros, pos1, pos2 - pos1 + 1))
        
        If Not Todos_Pro Then
            pos1 = pos2 + 2
            pos2 = InStr(pos1, parametros, ".") - 1
            'NroProc = CLng(Mid(parametros, pos1, pos2 - pos1 + 1))
            NroProc = Mid(parametros, pos1, pos2 - pos1 + 1)
            ListaNroProc = Replace(NroProc, "-", ",")
        Else
            'NroProc = 0
            NroProc = "0"
            ListaNroProc = "0"
        End If
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, ".") - 1
        Proc_Aprob = Mid(parametros, pos1, pos2 - pos1 + 1)
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, ".") - 1
        If pos2 > 0 Then
            Empresa = Mid(parametros, pos1, pos2 - pos1 + 1)
        Else
            pos2 = Len(parametros)
            Empresa = Mid(parametros, pos1, pos2 - pos1 + 1)
        End If
        
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
    End If
End If


Call Sueacu11(bpronro, Nroliq, Todos_Pro, Proc_Aprob, Empresa, Agrupado, AgrupaTE1, Tenro1, Estrnro1, AgrupaTE2, Tenro2, Estrnro2, AgrupaTE3, Tenro3, Estrnro3)

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

