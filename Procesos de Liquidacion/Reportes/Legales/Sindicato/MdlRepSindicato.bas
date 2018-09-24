Attribute VB_Name = "MdlRepSindicato"
Option Explicit

'Const Version = 1.01    'Se agregaron datos al log
'Const FechaVersion = "09/12/2005"

'Const Version = 1.02
'Const FechaVersion = "14/07/2006" ' FAF
        ' Si se eligen varios procesos y se da el caso de que en alguno no tenga liquidado ningun
        ' CO o AC de los configurados en el confrep, puede darse el caso que no muestre al empleado.
        ' Se soluciono de manera que muestra al empleado que tenga al menos un CO o AC en cualquiera
        ' de los procesos seleccionados.
        
'Const Version = 1.03
'Const FechaVersion = "14/07/2006" 'Martin Ferraro - Encriptacion de string connection
        
'Const Version = 1.04
'Const FechaVersion = "22/06/2011" 'Matias Dallegro - Se cambio la dimension del array arreglo(20) a arreglo(30)

'Const Version = 1.05
'Const FechaVersion = "19/12/2014" 'Fernandez, Matias -CAS-28549  - MEGATLON - Error en reporte de sindicatos-
                                  ' Se acomodaron los and en consulta por estructura y se agrego el modulo de clase
                                  'feriado
                 

'Const Version = 1.06
'Const FechaVersion = "09/01/2015" 'Fernandez, Matias -CAS-28549  - MEGATLON - Error en reporte de sindicatos-
                                  'Se corrigio consulta sql y mensajes de log

Const Version = 1.07
Const FechaVersion = "16/01/2015" 'Fernandez, Matias -CAS-28549  - MEGATLON - Error en reporte de sindicatos-
                                  'Se corrigen mas errores al armar la sql


Global IdUser As String
Global Fecha As Date
Global hora As String

Global Fecha_Inicio_periodo As Date
Global Fecha_Fin_Periodo As Date
Global StrSql2 As String
Global TipoEstructura As Long



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
    
On Error GoTo MAINE

    Nombre_Arch = PathFLog & "Reporte_Sindicato" & "-" & NroProcesoBatch & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Version = " & Version
    Flog.writeline "Fecha   = " & FechaVersion
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline
    Flog.writeline "PID = " & PID
    
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
    
On Error GoTo MAINE

    'Cambio el estado del proceso a Procesando y el PID
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 34 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        IdUser = rs_batch_proceso!IdUser
        Fecha = rs_batch_proceso!bprcfecha
        hora = rs_batch_proceso!bprchora
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
    
    objConn.Close
    objconnProgreso.Close
    
    Exit Sub
    
MAINE:
    Flog.writeline "Error Main: " & Err.Description
    Flog.writeline "Ultimo sql Ejecutado: " & StrSql
    
End Sub


Public Sub Suesin03_Provincia(ByVal bpronro As Long, ByVal Nroliq As Long, ByVal Todos_Pro As Boolean, ByVal Proc_Aprob As Integer, ByVal Empresa As Long, ByVal Todos_Sindicatos As Boolean, ByVal Nro_Sindicato As Long, ByVal Agrupado As Integer)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del Reporte de Sindicato
' Autor      : FGZ
' Fecha      : 02/03/2004
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------

Dim fechadesde As Date
Dim fechahasta As Date

Dim Arreglo(20) As Single
Dim I As Integer
Dim Ultimo_Empleado As Long
Dim Estructura As Long
Dim PrimeraVez As Boolean
Dim ColumnaConfiguracion As Boolean
Dim EncontroValor As Boolean

Dim rs_Reporte As New ADODB.Recordset
Dim rs_Periodo As New ADODB.Recordset
Dim rs_Procesos As New ADODB.Recordset
Dim rs_Confrep As New ADODB.Recordset
Dim rs_Rep06 As New ADODB.Recordset
Dim rs_Detliq As New ADODB.Recordset
Dim rs_Acu_Liq As New ADODB.Recordset
Dim rs_Domicilio As New ADODB.Recordset
Dim rs_Tercero As New ADODB.Recordset

On Error GoTo CE

'Inicializacion
For I = 1 To 20
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

' Comienzo la transaccion
MyBeginTrans

'Depuracion del Temporario
StrSql = "DELETE FROM rep06 "
StrSql = StrSql & " WHERE pliqnro = " & Nroliq
If Not Todos_Pro Then
    StrSql = StrSql & " AND pronro = '" & NroProc & "'"
Else
    StrSql = StrSql & " AND pronro = '0'"
    StrSql = StrSql & " AND proaprob = " & CInt(Proc_Aprob)
End If
If Todos_Sindicatos Then
    StrSql = StrSql & " AND todos_sind = -1"
Else
    StrSql = StrSql & " AND todos_sind = 0"
    StrSql = StrSql & " AND sindicato = " & CInt(Nro_Sindicato)
End If
StrSql = StrSql & " AND empresa = " & Empresa
StrSql = StrSql & " AND agrup = " & Agrupado
objConn.Execute StrSql, , adExecuteNoRecords

'Configuracion del Reporte
StrSql = "SELECT * FROM confrep WHERE repnro = 7 "
OpenRecordset StrSql, rs_Confrep

If rs_Confrep.EOF Then
    Flog.writeline "No se encontró la configuración del Reporte"
    Exit Sub
End If

'FGZ ' 25/01/2005
' Busco el tipo de estructura para el join
StrSql = "SELECT * FROM confrep WHERE repnro = 7 "
StrSql = StrSql & " AND upper(conftipo) = 'TE'"
If rs_Confrep.State = adStateOpen Then rs_Confrep.Close
OpenRecordset StrSql, rs_Confrep
If Not rs_Confrep.EOF Then
    TipoEstructura = rs_Confrep!confval
    Flog.writeline "Hay configurado un tipo de estructura para el JOIN " & TipoEstructura
End If

'Configuracion del Reporte
StrSql = "SELECT * FROM confrep WHERE repnro = 7 "
OpenRecordset StrSql, rs_Confrep
If rs_Confrep.EOF Then
    Flog.writeline "No se encontró la configuración del Reporte"
    Exit Sub
End If

'Busco los procesos a evaluar
StrSql = "SELECT cabliq.*, proceso.*, periodo.*, empleado.*, Sindicato.estrnro sindnro FROM  empleado " '3)
StrSql = StrSql & " INNER JOIN his_estructura Sindicato ON Sindicato.ternro = empleado.ternro "
StrSql = StrSql & " INNER JOIN cabliq ON cabliq.empleado = empleado.ternro "
StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro "
StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro "
StrSql = StrSql & " INNER JOIN his_estructura empresa ON empresa.ternro = empleado.ternro and empresa.tenro = 10 "
'StrSql = StrSql & " INNER JOIN estructura emp ON emp.estrnro = empresa.estrnro AND emp.empnro =" & Empresa
StrSql = StrSql & " INNER JOIN estructura emp ON emp.estrnro = empresa.estrnro "
'StrSql = StrSql & " WHERE Sindicato.tenro = 16 AND "
StrSql = StrSql & " WHERE Sindicato.tenro = " & TipoEstructura & " AND "
StrSql = StrSql & " (Sindicato.htetdesde <= " & ConvFecha(Fecha_Fin_Periodo) & ") AND "
StrSql = StrSql & " ((" & ConvFecha(Fecha_Fin_Periodo) & " <= Sindicato.htethasta) or (Sindicato.htethasta is null))"
StrSql = StrSql & " AND (empresa.htetdesde <= " & ConvFecha(Fecha_Fin_Periodo) & ") AND "
StrSql = StrSql & " ((" & ConvFecha(Fecha_Fin_Periodo) & " <= empresa.htethasta) or (empresa.htethasta is null))"
StrSql = StrSql & " AND periodo.pliqnro =" & Nroliq
If Not Todos_Sindicatos Then
    StrSql = StrSql & " AND Sindicato.estrnro = " & Nro_Sindicato
End If
If Not Todos_Pro Then
    StrSql = StrSql & " AND proceso.pronro IN (" & ListaNroProc & ")"
Else
    StrSql = StrSql & " AND proaprob = " & CInt(Proc_Aprob)
    StrSql = StrSql & " AND proceso.empnro = " & Empresa
End If
StrSql = StrSql & " ORDER BY empleado.ternro"

Flog.writeline "SQL Empleados: " & StrSql

OpenRecordset StrSql, rs_Procesos

'seteo de las variables de progreso
Progreso = 0
CConceptosAProc = rs_Procesos.RecordCount
If CConceptosAProc = 0 Then
    CConceptosAProc = 1
End If
IncPorc = (100 / CConceptosAProc)


Ultimo_Empleado = -1
Do While Not rs_Procesos.EOF
    If Ultimo_Empleado <> rs_Procesos!Ternro Then
        For I = 1 To 20
            Arreglo(I) = 0
        Next I
        EncontroValor = False
    End If
    Ultimo_Empleado = rs_Procesos!Ternro
    
    rs_Confrep.MoveFirst
    Do While Not rs_Confrep.EOF
        Select Case UCase(rs_Confrep!conftipo)
        Case "AC":
            StrSql = "SELECT * FROM acu_liq " & _
                     " INNER JOIN cabliq ON acu_liq.cliqnro = cabliq.cliqnro " & _
                     " WHERE acu_liq.acunro = " & rs_Confrep!confval & _
                     " AND cabliq.cliqnro =" & rs_Procesos!cliqnro
            OpenRecordset StrSql, rs_Acu_Liq
            Do While Not rs_Acu_Liq.EOF
                Arreglo(rs_Confrep!confnrocol) = Arreglo(rs_Confrep!confnrocol) + rs_Acu_Liq!almonto
                EncontroValor = True
                rs_Acu_Liq.MoveNext
            Loop
            
        Case "CO", "EDR":
            StrSql = "SELECT * FROM detliq " & _
                     " INNER JOIN cabliq ON detliq.cliqnro = cabliq.cliqnro " & _
                     " INNER JOIN concepto ON detliq.concnro = concepto.concnro " & _
                     " WHERE (concepto.conccod = " & rs_Confrep!confval & _
                     " OR concepto.conccod = '" & rs_Confrep!confval2 & "')" & _
                     " AND cabliq.cliqnro =" & rs_Procesos!cliqnro
            OpenRecordset StrSql, rs_Detliq
            Do While Not rs_Detliq.EOF
                Arreglo(rs_Confrep!confnrocol) = Arreglo(rs_Confrep!confnrocol) + rs_Detliq!dlimonto
                EncontroValor = True
                rs_Detliq.MoveNext
            Loop
        Case Else
            'estamos en la columna 1
            If rs_Confrep!confnrocol = 1 Then
                'si el valor es 0 ==> la provincia es del domicilio del empleado
                ' sino el valor es el tenro y la provincia es la del dom de la estructura de ese tipo
                If rs_Confrep!confval = -1 Then
                    Estructura = -1
                Else
                    Estructura = rs_Confrep!confval
                End If
            Else    'FGZ - 25/01/2005
                'Es el tipo de estructura contra el cual se mapea el join inicial de empleados a procesar
                TipoEstructura = rs_Confrep!confval
            End If
        End Select
        
        'Siguiente confrep
        rs_Confrep.MoveNext
    Loop
                
    If EsElUltimoEmpleado(rs_Procesos, Ultimo_Empleado) And EncontroValor Then
    
        'Si no existe el rep03
        StrSql = "SELECT * FROM rep06 "
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
        StrSql = StrSql & " AND agrup =" & Agrupado
        OpenRecordset StrSql, rs_Rep06
    
        If rs_Rep06.EOF Then
        
            'Inserto
            StrSql = "INSERT INTO rep06 (bpronro,pliqnro,pronro,proaprob,empresa,iduser,fecha,hora,"
            StrSql = StrSql & "ternro,empleg,apenom,todos_sind,sindicato,provdesc,agrup"
            For I = 2 To 20
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
            StrSql = StrSql & "'" & hora & "',"
            StrSql = StrSql & rs_Procesos!Ternro & ","
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
            
            If Todos_Sindicatos Then
                StrSql = StrSql & "-1," & rs_Procesos!sindnro & ","
            Else
                StrSql = StrSql & "0," & rs_Procesos!sindnro & ","
            End If

            ' Carga la provincia para cortar por provincia
            If Estructura = -1 Then
                StrSql2 = " SELECT * FROM detdom " & _
                         " INNER JOIN cabdom ON detdom.domnro = cabdom.domnro " & _
                         " INNER JOIN provincia ON provincia.provnro = detdom.provnro " & _
                         " WHERE cabdom.ternro = " & rs_Procesos!Ternro & " AND " & _
                         " cabdom.domdefault = -1"
                OpenRecordset StrSql2, rs_Domicilio
                If Not rs_Domicilio.EOF Then
                    'StrSql = StrSql & rs_Domicilio!provnro & ","
                    StrSql = StrSql & "'" & rs_Domicilio!provdesc & "',"
                Else
                    'StrSql = StrSql & "0,"
                    StrSql = StrSql & "'Sin definir ',"
                    Flog.writeline "No se encontró la provincia para el empleado: " & rs_Procesos!Ternro
                End If
            Else
                StrSql2 = "SELECT * FROM His_estructura "
                StrSql2 = StrSql2 & " INNER JOIN tercero ON tercero.ternro = his_estructura.estrnro "
                StrSql2 = StrSql2 & " WHERE his_estructura.tenro = " & Estructura & " AND "
                StrSql2 = StrSql2 & " (his_estructura.htetdesde <= " & ConvFecha(Fecha_Fin_Periodo) & ") AND "
                StrSql2 = StrSql2 & " ((" & ConvFecha(Fecha_Fin_Periodo) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
                StrSql2 = StrSql2 & " AND his_estructura.ternro =" & rs_Procesos!Ternro
                OpenRecordset StrSql2, rs_Tercero
            
                If Not rs_Tercero.EOF Then
                    StrSql2 = " SELECT * FROM detdom " & _
                             " INNER JOIN cabdom ON detdom.domnro = cabdom.domnro " & _
                             " INNER JOIN provincia ON provincia.provnro = detdom.provnro " & _
                             " WHERE cabdom.ternro = " & rs_Tercero!Ternro & " AND cabdom.domdefault = -1"
                    OpenRecordset StrSql2, rs_Domicilio
                    If Not rs_Domicilio.EOF Then
                        'StrSql = StrSql & rs_Domicilio!provnro & ","
                        StrSql = StrSql & "'" & rs_Domicilio!provdesc & "',"
                    Else
                        'StrSql = StrSql & "0,"
                        StrSql = StrSql & "'Sin definir ',"
                        Flog.writeline "No se encontró la provincia para el empleado: " & rs_Procesos!Ternro
                    End If
                Else
                    'StrSql = StrSql & "0,"
                    StrSql = StrSql & "' ',"
                    Flog.writeline "Tipo de estructura no Configurada"
                End If
            End If
            StrSql = StrSql & Agrupado
            
            For I = 2 To 20
                If Arreglo(I) <> 0 And Not IsNull(Arreglo(I)) Then
                    StrSql = StrSql & "," & Arreglo(I)
                End If
            Next I
            StrSql = StrSql & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        Else
            'Actualizo
            StrSql = "UPDATE rep06 "
            PrimeraVez = True
            For I = 2 To 20
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
            StrSql = StrSql & " AND agrup =" & Agrupado
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
            
        
    Ultimo_Empleado = rs_Procesos!Ternro
        
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
If rs_Rep06.State = adStateOpen Then rs_Rep06.Close
If rs_Reporte.State = adStateOpen Then rs_Reporte.Close
If rs_Domicilio.State = adStateOpen Then rs_Domicilio.Close

Set rs_Procesos = Nothing
Set rs_Confrep = Nothing
Set rs_Detliq = Nothing
Set rs_Acu_Liq = Nothing
Set rs_Rep06 = Nothing
Set rs_Periodo = Nothing
Set rs_Reporte = Nothing
Set rs_Domicilio = Nothing


Exit Sub
CE:
    HuboError = True
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo sql Ejecutado: " & StrSql
    MyRollbackTrans

End Sub

Public Sub Suesin03(ByVal bpronro As Long, ByVal Nroliq As Long, ByVal Todos_Pro As Boolean, ByVal Proc_Aprob As Integer, ByVal Empresa As Long, ByVal Todos_Sindicatos As Boolean, ByVal Nro_Sindicato As Long, ByVal Agrupado As Integer)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del Reporte de Sindicato
' Autor      : FGZ
' Fecha      : 02/03/2004
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------

Dim fechadesde As Date
Dim fechahasta As Date

Dim Arreglo(30) As Single
Dim I As Integer
Dim Ultimo_Empleado As Long
Dim Estructura As Long
Dim PrimeraVez As Boolean
Dim ColumnaConfiguracion As Boolean
Dim EncontroValor As Boolean

Dim rs_Reporte As New ADODB.Recordset
Dim rs_Periodo As New ADODB.Recordset
Dim rs_Procesos As New ADODB.Recordset
Dim rs_Confrep As New ADODB.Recordset
Dim rs_Rep06 As New ADODB.Recordset
Dim rs_Detliq As New ADODB.Recordset
Dim rs_Acu_Liq As New ADODB.Recordset
Dim rs_Domicilio As New ADODB.Recordset
Dim rs_Tercero As New ADODB.Recordset

On Error GoTo CE

'Inicializacion
For I = 1 To 30
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

' Comienzo la transaccion
MyBeginTrans

'Depuracion del Temporario
StrSql = "DELETE FROM rep06 "
StrSql = StrSql & " WHERE pliqnro = " & Nroliq
If Not Todos_Pro Then
    StrSql = StrSql & " AND pronro = '" & NroProc & "'"
Else
    StrSql = StrSql & " AND pronro = '0'"
    StrSql = StrSql & " AND proaprob = " & CInt(Proc_Aprob)
End If
If Todos_Sindicatos Then
    StrSql = StrSql & " AND todos_sind = -1"
Else
    StrSql = StrSql & " AND todos_sind = 0"
    StrSql = StrSql & " AND sindicato = " & CInt(Nro_Sindicato)
End If
StrSql = StrSql & " AND empresa = " & Empresa
StrSql = StrSql & " AND agrup = " & Agrupado
objConn.Execute StrSql, , adExecuteNoRecords

'Configuracion del Reporte
StrSql = "SELECT * FROM confrep WHERE repnro = 7 "
OpenRecordset StrSql, rs_Confrep
If rs_Confrep.EOF Then
    Flog.writeline "No se encontró la configuración del Reporte"
    Exit Sub
End If

'FGZ ' 25/01/2005
' Busco el tipo de estructura para el join
StrSql = "SELECT * FROM confrep WHERE repnro = 7 "
StrSql = StrSql & " AND upper(conftipo) = 'TE'"
If rs_Confrep.State = adStateOpen Then rs_Confrep.Close
OpenRecordset StrSql, rs_Confrep
If Not rs_Confrep.EOF Then
    TipoEstructura = rs_Confrep!confval
    Flog.writeline "Hay configurado un tipo de estructura para el JOIN " & TipoEstructura
End If

'Configuracion del Reporte
StrSql = "SELECT * FROM confrep WHERE repnro = 7 "
OpenRecordset StrSql, rs_Confrep
If rs_Confrep.EOF Then
    Flog.writeline "No se encontró la configuración del Reporte"
    Exit Sub
End If

'Busco los procesos a evaluar
StrSql = "SELECT  cabliq.*, proceso.*, periodo.*, empleado.*, Sindicato.estrnro sindnro FROM  empleado "
StrSql = StrSql & " INNER JOIN his_estructura Sindicato ON Sindicato.ternro = empleado.ternro "
StrSql = StrSql & " INNER JOIN cabliq ON cabliq.empleado = empleado.ternro "
StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro "
StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro "
StrSql = StrSql & " INNER JOIN his_estructura empresa ON empresa.ternro = empleado.ternro and empresa.tenro = 10 "
'StrSql = StrSql & " INNER JOIN estructura emp ON emp.estrnro = empresa.estrnro AND emp.empnro =" & Empresa
StrSql = StrSql & " INNER JOIN estructura emp ON emp.estrnro = empresa.estrnro "
'StrSql = StrSql & " WHERE Sindicato.tenro = 16 AND "
StrSql = StrSql & " WHERE Sindicato.tenro = " & TipoEstructura & " AND "
StrSql = StrSql & " (Sindicato.htetdesde <= " & ConvFecha(Fecha_Fin_Periodo) & ") AND "
StrSql = StrSql & " ((" & ConvFecha(Fecha_Fin_Periodo) & " <= Sindicato.htethasta) or (Sindicato.htethasta is null))"
StrSql = StrSql & " AND (empresa.htetdesde <= " & ConvFecha(Fecha_Fin_Periodo) & ") AND "
StrSql = StrSql & " ((" & ConvFecha(Fecha_Fin_Periodo) & " <= empresa.htethasta) or (empresa.htethasta is null))"
StrSql = StrSql & " AND periodo.pliqnro =" & Nroliq
'StrSql = StrSql & " AND periodo.empnro =" & Empresa
If Not Todos_Sindicatos Then
    StrSql = StrSql & " AND Sindicato.estrnro = " & Nro_Sindicato
End If
If Not Todos_Pro Then
     StrSql = StrSql & " AND proceso.pronro IN (" & ListaNroProc & ")"
'    StrSql = StrSql & " AND proceso.pronro =" & NroProc
Else
    StrSql = StrSql & " AND proaprob = " & CInt(Proc_Aprob)
    StrSql = StrSql & " AND proceso.empnro = " & Empresa
End If
StrSql = StrSql & " ORDER BY empleado.ternro"

Flog.writeline "SQL Empleados: " & StrSql

OpenRecordset StrSql, rs_Procesos

'seteo de las variables de progreso
Progreso = 0
CConceptosAProc = rs_Procesos.RecordCount
If CConceptosAProc = 0 Then
    CConceptosAProc = 1
End If
IncPorc = (100 / CConceptosAProc)

Ultimo_Empleado = -1
Do While Not rs_Procesos.EOF
    If Ultimo_Empleado <> rs_Procesos!Ternro Then
        For I = 1 To 30
            Arreglo(I) = 0
        Next I
        EncontroValor = False
    End If
    Ultimo_Empleado = rs_Procesos!Ternro
    
    rs_Confrep.MoveFirst
    Do While Not rs_Confrep.EOF
        Select Case UCase(rs_Confrep!conftipo)
        Case "AC":
            StrSql = "SELECT * FROM acu_liq " & _
                     " INNER JOIN cabliq ON acu_liq.cliqnro = cabliq.cliqnro " & _
                     " WHERE acu_liq.acunro = " & rs_Confrep!confval & _
                     " AND cabliq.cliqnro =" & rs_Procesos!cliqnro
            OpenRecordset StrSql, rs_Acu_Liq
            Do While Not rs_Acu_Liq.EOF
                Arreglo(rs_Confrep!confnrocol) = Arreglo(rs_Confrep!confnrocol) + rs_Acu_Liq!almonto
                EncontroValor = True
                rs_Acu_Liq.MoveNext
            Loop
            
        Case "CO", "EDR":
            StrSql = "SELECT * FROM detliq " & _
                     " INNER JOIN cabliq ON detliq.cliqnro = cabliq.cliqnro " & _
                     " INNER JOIN concepto ON detliq.concnro = concepto.concnro " & _
                     " WHERE (concepto.conccod = " & rs_Confrep!confval & _
                     " OR concepto.conccod = '" & rs_Confrep!confval2 & "')" & _
                     " AND cabliq.cliqnro =" & rs_Procesos!cliqnro
            OpenRecordset StrSql, rs_Detliq
            Do While Not rs_Detliq.EOF
                Arreglo(rs_Confrep!confnrocol) = Arreglo(rs_Confrep!confnrocol) + rs_Detliq!dlimonto
                EncontroValor = True
                rs_Detliq.MoveNext
            Loop
        Case Else
            'estamos en la columna 1
            If rs_Confrep!confnrocol = 1 Then
                'si el valor es 0 ==> la provincia es del domicilio del empleado
                ' sino el valor es el tenro y la provincia es la del dom de la estructura de ese tipo
                If rs_Confrep!confval = -1 Then
                    Estructura = -1
                Else
                    Estructura = rs_Confrep!confval
                End If
            End If
        End Select
        
        'Siguiente confrep
        rs_Confrep.MoveNext
    Loop
        
    If EsElUltimoEmpleado(rs_Procesos, Ultimo_Empleado) And EncontroValor Then
    
        'Si no existe el rep03
        StrSql = "SELECT * FROM rep06 "
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
        StrSql = StrSql & " AND agrup =" & Agrupado
        OpenRecordset StrSql, rs_Rep06
    
        If rs_Rep06.EOF Then
        
            'Inserto
            StrSql = "INSERT INTO rep06 (bpronro,pliqnro,pronro,proaprob,empresa,iduser,fecha,hora,"
            StrSql = StrSql & "ternro,empleg,apenom,todos_sind,sindicato,agrup"
            For I = 2 To 30
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
            StrSql = StrSql & "'" & hora & "',"
            StrSql = StrSql & rs_Procesos!Ternro & ","
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
            
            If Todos_Sindicatos Then
                StrSql = StrSql & "-1," & rs_Procesos!sindnro & ","
            Else
                StrSql = StrSql & "0," & rs_Procesos!sindnro & ","
            End If
            StrSql = StrSql & Agrupado
            
            For I = 2 To 30
                If Arreglo(I) <> 0 And Not IsNull(Arreglo(I)) Then
                    StrSql = StrSql & "," & Arreglo(I)
                End If
            Next I
            StrSql = StrSql & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        Else
            'Actualizo
            StrSql = "UPDATE rep06 "
            PrimeraVez = True
            For I = 2 To 30
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
            StrSql = StrSql & " AND agrup =" & Agrupado
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
            
    Ultimo_Empleado = rs_Procesos!Ternro
    
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
If rs_Rep06.State = adStateOpen Then rs_Rep06.Close
If rs_Reporte.State = adStateOpen Then rs_Reporte.Close
If rs_Domicilio.State = adStateOpen Then rs_Domicilio.Close

Set rs_Procesos = Nothing
Set rs_Confrep = Nothing
Set rs_Detliq = Nothing
Set rs_Acu_Liq = Nothing
Set rs_Rep06 = Nothing
Set rs_Periodo = Nothing
Set rs_Reporte = Nothing
Set rs_Domicilio = Nothing


Exit Sub
CE:
    HuboError = True
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo sql Ejecutado: " & StrSql
    MyRollbackTrans

End Sub

Public Sub Suesin03_Estructuras(ByVal bpronro As Long, ByVal Nroliq As Long, ByVal Todos_Pro As Boolean, ByVal Proc_Aprob As Integer, ByVal Empresa As Long, ByVal Todos_Sindicatos As Boolean, ByVal Nro_Sindicato As Long, ByVal Agrup As Integer, ByVal Agrupado As Boolean, _
    ByVal AgrupaTE1 As Boolean, ByVal Tenro1 As Long, Estrnro1 As Long, _
    ByVal AgrupaTE2 As Boolean, ByVal Tenro2 As Long, Estrnro2 As Long, _
    ByVal AgrupaTE3 As Boolean, ByVal Tenro3 As Long, Estrnro3 As Long)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del Reporte de Sindicato
' Autor      : FGZ
' Fecha      : 02/03/2004
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------

Dim fechadesde As Date
Dim fechahasta As Date

Dim Arreglo(30) As Single
Dim I As Integer
Dim Ultimo_Empleado As Long
Dim Estructura As Long
Dim PrimeraVez As Boolean
Dim ColumnaConfiguracion As Boolean
Dim EncontroValor As Boolean

Dim rs_Reporte As New ADODB.Recordset
Dim rs_Periodo As New ADODB.Recordset
Dim rs_Procesos As New ADODB.Recordset
Dim rs_Confrep As New ADODB.Recordset
Dim rs_Rep06 As New ADODB.Recordset
Dim rs_Detliq As New ADODB.Recordset
Dim rs_Acu_Liq As New ADODB.Recordset
Dim rs_Domicilio As New ADODB.Recordset
Dim rs_Tercero As New ADODB.Recordset
Dim rs_Estructura As New ADODB.Recordset

On Error GoTo CE

'Inicializacion
For I = 1 To 30
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

' Comienzo la transaccion
MyBeginTrans

'Depuracion del Temporario
StrSql = "DELETE FROM rep06 "
StrSql = StrSql & " WHERE pliqnro = " & Nroliq
If Not Todos_Pro Then
    StrSql = StrSql & " AND pronro = '" & NroProc & "'"
Else
    StrSql = StrSql & " AND pronro = '0'"
    StrSql = StrSql & " AND proaprob = " & CInt(Proc_Aprob)
End If
If Todos_Sindicatos Then
    StrSql = StrSql & " AND todos_sind = -1"
Else
    StrSql = StrSql & " AND todos_sind = 0"
    StrSql = StrSql & " AND sindicato = " & CInt(Nro_Sindicato)
End If
StrSql = StrSql & " AND empresa = " & Empresa
StrSql = StrSql & " AND agrup = " & Agrup
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


'Configuracion del Reporte
StrSql = "SELECT * FROM confrep WHERE repnro = 7 "
OpenRecordset StrSql, rs_Confrep
If rs_Confrep.EOF Then
    Flog.writeline "No se encontró la configuración del Reporte"
    Exit Sub
End If

'FGZ ' 25/01/2005
' Busco el tipo de estructura para el join
StrSql = "SELECT * FROM confrep WHERE repnro = 7 "
StrSql = StrSql & " AND upper(conftipo) = 'TE'"
If rs_Confrep.State = adStateOpen Then rs_Confrep.Close
OpenRecordset StrSql, rs_Confrep
If Not rs_Confrep.EOF Then
    TipoEstructura = rs_Confrep!confval
    Flog.writeline "Hay configurado un tipo de estructura para el JOIN " & TipoEstructura
End If

'Configuracion del Reporte
StrSql = "SELECT * FROM confrep WHERE repnro = 7 "
OpenRecordset StrSql, rs_Confrep
If rs_Confrep.EOF Then
    Flog.writeline "No se encontró la configuración del Reporte"
    Exit Sub
End If


'Busco los procesos a evaluar
StrSql = "SELECT cabliq.*, proceso.*, periodo.*, empleado.*,Sindicato.estrnro sindnro "   ' 1)
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
StrSql = StrSql & " INNER JOIN his_estructura Sindicato ON sindicato.ternro = empleado.ternro "
StrSql = StrSql & " INNER JOIN his_estructura empresa ON empresa.ternro = empleado.ternro and empresa.tenro = 10 "

'StrSql = StrSql & " INNER JOIN estructura emp ON emp.estrnro = empresa.estrnro AND emp.empnro =" & Empresa
StrSql = StrSql & " INNER JOIN estructura emp ON emp.estrnro = empresa.estrnro "
If AgrupaTE1 Then
    StrSql = StrSql & " INNER JOIN his_estructura TE1 ON te1.ternro = empleado.ternro "
End If
If AgrupaTE2 Then
    StrSql = StrSql & " INNER JOIN his_estructura TE2 ON te2.ternro = empleado.ternro "
End If
If AgrupaTE3 Then
    StrSql = StrSql & " INNER JOIN his_estructura TE3 ON te3.ternro = empleado.ternro "
End If
StrSql = StrSql & " WHERE (periodo.pliqnro =" & Nroliq
If Not Todos_Pro Then
    StrSql = StrSql & " AND proceso.pronro IN (" & ListaNroProc & ")) "
'    StrSql = StrSql & " AND proceso.pronro =" & NroProc & ") "
Else
    StrSql = StrSql & " AND proceso.proaprob = " & CInt(Proc_Aprob) & ") "
    StrSql = StrSql & " AND proceso.empnro = " & Empresa
End If
'StrSql = StrSql & " AND sindicato.tenro = 16 AND "
StrSql = StrSql & " AND Sindicato.tenro = " & TipoEstructura & " AND "
If Not Todos_Sindicatos Then
    StrSql = StrSql & " sindicato.estrnro = " & Nro_Sindicato & "AND "
End If
StrSql = StrSql & " (sindicato.htetdesde <= " & ConvFecha(Fecha_Fin_Periodo) & ") AND "
StrSql = StrSql & " ((" & ConvFecha(Fecha_Fin_Periodo) & " <= sindicato.htethasta) or (sindicato.htethasta is null)) "
StrSql = StrSql & " AND (empresa.htetdesde <= " & ConvFecha(Fecha_Fin_Periodo) & ") AND "
StrSql = StrSql & " ((" & ConvFecha(Fecha_Fin_Periodo) & " <= empresa.htethasta) or (empresa.htethasta is null)) " 'mdf
If AgrupaTE1 Then
    StrSql = StrSql & "  AND te1.tenro = " & Tenro1 '& " AND " ---->mdf
    If Estrnro1 <> 0 Then
        StrSql = StrSql & " AND te1.estrnro = " & Estrnro1 '& " AND " -----> mdf
    End If
    StrSql = StrSql & " And (te1.htetdesde <= " & ConvFecha(Fecha_Fin_Periodo) & ") AND "
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

Flog.writeline "SQL Empleados ---> " & StrSql

OpenRecordset StrSql, rs_Procesos


'seteo de las variables de progreso
Progreso = 0
CConceptosAProc = rs_Procesos.RecordCount
If CConceptosAProc = 0 Then
    CConceptosAProc = 1
End If
IncPorc = (100 / CConceptosAProc)
Ultimo_Empleado = -1
Do While Not rs_Procesos.EOF
    If Ultimo_Empleado <> rs_Procesos!Ternro Then
        For I = 1 To 30
            Arreglo(I) = 0
        Next I
        EncontroValor = False
    End If
    Ultimo_Empleado = rs_Procesos!Ternro
    
    rs_Confrep.MoveFirst
    
    Do While Not rs_Confrep.EOF
        'ColumnaConfiguracion = False
        Select Case UCase(rs_Confrep!conftipo)
        Case "AC":
            StrSql = "SELECT * FROM acu_liq " & _
                     " INNER JOIN cabliq ON acu_liq.cliqnro = cabliq.cliqnro " & _
                     " WHERE acu_liq.acunro = " & rs_Confrep!confval & _
                     " AND cabliq.cliqnro =" & rs_Procesos!cliqnro
            OpenRecordset StrSql, rs_Acu_Liq
            Do While Not rs_Acu_Liq.EOF
                Arreglo(rs_Confrep!confnrocol) = Arreglo(rs_Confrep!confnrocol) + rs_Acu_Liq!almonto
                EncontroValor = True
                rs_Acu_Liq.MoveNext
            Loop
            
        Case "CO", "EDR":
            StrSql = "SELECT * FROM detliq " & _
                     " INNER JOIN cabliq ON detliq.cliqnro = cabliq.cliqnro " & _
                     " INNER JOIN concepto ON detliq.concnro = concepto.concnro " & _
                     " WHERE (concepto.conccod = " & rs_Confrep!confval & _
                     " OR concepto.conccod = '" & rs_Confrep!confval2 & "')" & _
                     " AND cabliq.cliqnro =" & rs_Procesos!cliqnro
            OpenRecordset StrSql, rs_Detliq
            Do While Not rs_Detliq.EOF
                Arreglo(rs_Confrep!confnrocol) = Arreglo(rs_Confrep!confnrocol) + rs_Detliq!dlimonto
                EncontroValor = True
                rs_Detliq.MoveNext
            Loop
        Case Else
            'ColumnaConfiguracion = True
            'estamos en la columna 1
            If rs_Confrep!confnrocol = 1 Then
                'si el valor es 0 ==> la provincia es del domicilio del empleado
                ' sino el valor es el tenro y la provincia es la del dom de la estructura de ese tipo
                If rs_Confrep!confval = -1 Then
                    Estructura = -1
                Else
                    Estructura = rs_Confrep!confval
                End If
            End If
        End Select
        
        'Siguiente confrep
        rs_Confrep.MoveNext
    Loop
                
    If EsElUltimoEmpleado(rs_Procesos, Ultimo_Empleado) And EncontroValor Then
    
        'Si no existe el rep03
        StrSql = "SELECT * FROM rep06 "
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
        StrSql = StrSql & " AND agrup =" & Agrup
        OpenRecordset StrSql, rs_Rep06
    
        If rs_Rep06.EOF Then
            'Inserto
            StrSql = "INSERT INTO rep06 (bpronro,pliqnro,pronro,proaprob,empresa,iduser,fecha,hora,"
            StrSql = StrSql & "ternro,empleg,apenom,todos_sind,sindicato,agrup,tenro1,estrnro1,tedesc1,estrdesc1,tenro2,estrnro2,tedesc2,estrdesc2,tenro3,estrnro3,tedesc3,estrdesc3"
            For I = 2 To 30
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
            StrSql = StrSql & "'" & hora & "',"
            StrSql = StrSql & rs_Procesos!Ternro & ","
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
            
            
            If Todos_Sindicatos Then
                StrSql = StrSql & "-1," & rs_Procesos!sindnro & ","
            Else
                StrSql = StrSql & "0," & rs_Procesos!sindnro & ","
            End If
            StrSql = StrSql & Agrup & ","
            
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
            
            For I = 2 To 30
                If Arreglo(I) <> 0 And Not IsNull(Arreglo(I)) Then
                    StrSql = StrSql & "," & Arreglo(I)
                End If
            Next I
            StrSql = StrSql & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        Else
            'Actualizo
            StrSql = "UPDATE rep06 "
            PrimeraVez = True
            For I = 2 To 30
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
            StrSql = StrSql & " AND agrup =" & Agrup
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
            
            
    Ultimo_Empleado = rs_Procesos!Ternro
        
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
If rs_Rep06.State = adStateOpen Then rs_Rep06.Close
If rs_Reporte.State = adStateOpen Then rs_Reporte.Close
If rs_Domicilio.State = adStateOpen Then rs_Domicilio.Close

Set rs_Procesos = Nothing
Set rs_Confrep = Nothing
Set rs_Detliq = Nothing
Set rs_Acu_Liq = Nothing
Set rs_Rep06 = Nothing
Set rs_Periodo = Nothing
Set rs_Reporte = Nothing
Set rs_Domicilio = Nothing


Exit Sub
CE:
    HuboError = True
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo sql Ejecutado: " & StrSql
    MyRollbackTrans

End Sub


Public Sub Suesin04(ByVal bpronro As Long, ByVal Nroliq As Long, ByVal Todos_Pro As Boolean, ByVal Proc_Aprob As Integer, ByVal Empresa As Long, ByVal Todos_Sindicatos As Boolean, ByVal Nro_Sindicato As Long, ByVal Agrupado As Integer, ByVal TextoAgrupado As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del Reporte de Sindicato
' Autor      : FGZ
' Fecha      : 02/03/2004
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------

Dim fechadesde As Date
Dim fechahasta As Date

Dim Arreglo(30) As Single
Dim I As Integer
Dim Ultimo_Empleado As Long
Dim Estructura As Long
Dim PrimeraVez As Boolean
Dim ColumnaConfiguracion As Boolean
Dim EncontroValor As Boolean

Dim rs_Reporte As New ADODB.Recordset
Dim rs_Periodo As New ADODB.Recordset
Dim rs_Procesos As New ADODB.Recordset
Dim rs_Confrep As New ADODB.Recordset
Dim rs_Rep06 As New ADODB.Recordset
Dim rs_Detliq As New ADODB.Recordset
Dim rs_Acu_Liq As New ADODB.Recordset
Dim rs_Domicilio As New ADODB.Recordset
Dim rs_Tercero As New ADODB.Recordset
Dim Rep As Long

On Error GoTo CE

'Inicializacion
For I = 1 To 30
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

' Comienzo la transaccion
MyBeginTrans

'Depuracion del Temporario
StrSql = "DELETE FROM rep06 "
StrSql = StrSql & " WHERE pliqnro = " & Nroliq
If Not Todos_Pro Then
    StrSql = StrSql & " AND pronro = '" & NroProc & "'"
Else
    StrSql = StrSql & " AND pronro = '0'"
    StrSql = StrSql & " AND proaprob = " & CInt(Proc_Aprob)
End If
If Todos_Sindicatos Then
    StrSql = StrSql & " AND todos_sind = -1"
Else
    StrSql = StrSql & " AND todos_sind = 0"
    StrSql = StrSql & " AND sindicato = " & CInt(Nro_Sindicato)
End If
StrSql = StrSql & " AND empresa = " & Empresa
StrSql = StrSql & " AND agrup = " & Agrupado
objConn.Execute StrSql, , adExecuteNoRecords

'Configuracion del Reporte
Select Case UCase(TextoAgrupado)
Case "UOCRA":
    StrSql = "SELECT * FROM confrep WHERE repnro = 45 "
    Rep = 45
Case "UOMRA":
    StrSql = "SELECT * FROM confrep WHERE repnro = 46 "
    Rep = 46
Case Else
    Flog.writeline "Parametros de agrupacion incorrecto"
    Exit Sub
End Select
OpenRecordset StrSql, rs_Confrep

If rs_Confrep.EOF Then
    Flog.writeline "No se encontró la configuración del Reporte"
    Exit Sub
End If


'FGZ ' 25/01/2005
' Busco el tipo de estructura para el join
StrSql = "SELECT * FROM confrep WHERE repnro = " & Rep
StrSql = StrSql & " AND upper(conftipo) = 'TE'"
If rs_Confrep.State = adStateOpen Then rs_Confrep.Close
OpenRecordset StrSql, rs_Confrep
If Not rs_Confrep.EOF Then
    TipoEstructura = rs_Confrep!confval
    Flog.writeline "Hay configurado un tipo de estructura para el JOIN " & TipoEstructura
End If

'Busco los procesos a evaluar
StrSql = "SELECT cabliq.*, proceso.*, periodo.*, empleado.*, Sindicato.estrnro sindnro FROM  empleado " '2)
StrSql = StrSql & " INNER JOIN his_estructura Sindicato ON Sindicato.ternro = empleado.ternro "
StrSql = StrSql & " INNER JOIN cabliq ON cabliq.empleado = empleado.ternro "
StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro "
StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro "
StrSql = StrSql & " INNER JOIN his_estructura  empresa ON empresa.ternro = empleado.ternro and empresa.tenro = 10 "
'StrSql = StrSql & " INNER JOIN estructura emp ON emp.estrnro = empresa.estrnro AND emp.empnro =" & Empresa
StrSql = StrSql & " INNER JOIN estructura emp ON emp.estrnro = empresa.estrnro "
'StrSql = StrSql & " WHERE Sindicato.tenro = 16 AND "
StrSql = StrSql & " WHERE Sindicato.tenro = " & TipoEstructura & " AND "
StrSql = StrSql & " (Sindicato.htetdesde <= " & ConvFecha(Fecha_Fin_Periodo) & ") AND "
StrSql = StrSql & " ((" & ConvFecha(Fecha_Fin_Periodo) & " <= Sindicato.htethasta) or (Sindicato.htethasta is null))"
StrSql = StrSql & " AND (empresa.htetdesde <= " & ConvFecha(Fecha_Fin_Periodo) & ") AND "
StrSql = StrSql & " ((" & ConvFecha(Fecha_Fin_Periodo) & " <= empresa.htethasta) or (empresa.htethasta is null))"
StrSql = StrSql & " AND periodo.pliqnro =" & Nroliq
If Not Todos_Sindicatos Then
    StrSql = StrSql & " AND Sindicato.estrnro = " & Nro_Sindicato
End If
If Not Todos_Pro Then
    StrSql = StrSql & " AND proceso.pronro IN (" & ListaNroProc & ")"
'    StrSql = StrSql & " AND proceso.pronro =" & NroProc
Else
    StrSql = StrSql & " AND proaprob = " & CInt(Proc_Aprob)
    StrSql = StrSql & " AND proceso.empnro = " & Empresa
End If
StrSql = StrSql & " ORDER BY empleado.ternro"

Flog.writeline "SQL Empleados: " & StrSql

OpenRecordset StrSql, rs_Procesos

'seteo de las variables de progreso
Progreso = 0
CConceptosAProc = rs_Procesos.RecordCount
If CConceptosAProc = 0 Then
    CConceptosAProc = 1
End If
IncPorc = (100 / CConceptosAProc)


Ultimo_Empleado = -1
Do While Not rs_Procesos.EOF
    If Ultimo_Empleado <> rs_Procesos!Ternro Then
        For I = 1 To 30
            Arreglo(I) = 0
        Next I
        EncontroValor = False
    End If
    Ultimo_Empleado = rs_Procesos!Ternro
    
    rs_Confrep.MoveFirst
    Do While Not rs_Confrep.EOF
        Select Case UCase(rs_Confrep!conftipo)
        Case "AC":
            StrSql = "SELECT * FROM acu_liq " & _
                     " INNER JOIN cabliq ON acu_liq.cliqnro = cabliq.cliqnro " & _
                     " WHERE acu_liq.acunro = " & rs_Confrep!confval & _
                     " AND cabliq.cliqnro =" & rs_Procesos!cliqnro
            OpenRecordset StrSql, rs_Acu_Liq
            Do While Not rs_Acu_Liq.EOF
                Arreglo(rs_Confrep!confnrocol) = Arreglo(rs_Confrep!confnrocol) + rs_Acu_Liq!almonto
                EncontroValor = True
                rs_Acu_Liq.MoveNext
            Loop
            
        Case "CO", "EDR":
            StrSql = "SELECT * FROM detliq " & _
                     " INNER JOIN cabliq ON detliq.cliqnro = cabliq.cliqnro " & _
                     " INNER JOIN concepto ON detliq.concnro = concepto.concnro " & _
                     " WHERE (concepto.conccod = " & rs_Confrep!confval & _
                     " OR concepto.conccod = '" & rs_Confrep!confval2 & "')" & _
                     " AND cabliq.cliqnro =" & rs_Procesos!cliqnro
            OpenRecordset StrSql, rs_Detliq
            Do While Not rs_Detliq.EOF
                Arreglo(rs_Confrep!confnrocol) = Arreglo(rs_Confrep!confnrocol) + rs_Detliq!dlimonto
                EncontroValor = True
                rs_Detliq.MoveNext
            Loop
        Case Else
            'estamos en la columna 1
            If rs_Confrep!confnrocol = 1 Then
                'si el valor es 0 ==> la provincia es del domicilio del empleado
                ' sino el valor es el tenro y la provincia es la del dom de la estructura de ese tipo
                If rs_Confrep!confval = -1 Then
                    Estructura = -1
                Else
                    Estructura = rs_Confrep!confval
                End If
            End If
        End Select
        
        'Siguiente confrep
        rs_Confrep.MoveNext
    Loop
            
                
    If EsElUltimoEmpleado(rs_Procesos, Ultimo_Empleado) And EncontroValor Then
        'Si no existe el rep03
        StrSql = "SELECT * FROM rep06 "
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
        StrSql = StrSql & " AND agrup =" & Agrupado
        OpenRecordset StrSql, rs_Rep06
    
        If rs_Rep06.EOF Then
            'Inserto
            StrSql = "INSERT INTO rep06 (bpronro,pliqnro,pronro,proaprob,empresa,iduser,fecha,hora,"
            StrSql = StrSql & "ternro,empleg,apenom,todos_sind,sindicato,agrup"
            For I = 1 To 30
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
            StrSql = StrSql & "'" & hora & "',"
            StrSql = StrSql & rs_Procesos!Ternro & ","
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
            
            If Todos_Sindicatos Then
                StrSql = StrSql & "-1," & rs_Procesos!sindnro & ","
            Else
                StrSql = StrSql & "0," & rs_Procesos!sindnro & ","
            End If
            StrSql = StrSql & Agrupado

            For I = 1 To 30
                If Arreglo(I) <> 0 And Not IsNull(Arreglo(I)) Then
                    StrSql = StrSql & "," & Arreglo(I)
                End If
            Next I
            StrSql = StrSql & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        Else
            'Actualizo
            PrimeraVez = True
            For I = 2 To 30
                If Arreglo(I) <> 0 And Not IsNull(Arreglo(I)) Then
                    If PrimeraVez Then
                        StrSql = "UPDATE rep06 "
                        StrSql = StrSql & "SET "
                        PrimeraVez = False
                    Else
                        StrSql = StrSql & ", "
                    End If
                    StrSql = StrSql & "col" & CStr(I) & "= col" & CStr(I) & "+ " & Arreglo(I)
                End If
            Next I
            If Not PrimeraVez Then
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
                StrSql = StrSql & " AND agrup =" & Agrupado
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
        End If
    End If
    
    'Actualizo el progreso del Proceso
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
             "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
            
            
    Ultimo_Empleado = rs_Procesos!Ternro
        
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
If rs_Rep06.State = adStateOpen Then rs_Rep06.Close
If rs_Reporte.State = adStateOpen Then rs_Reporte.Close
If rs_Domicilio.State = adStateOpen Then rs_Domicilio.Close

Set rs_Procesos = Nothing
Set rs_Confrep = Nothing
Set rs_Detliq = Nothing
Set rs_Acu_Liq = Nothing
Set rs_Rep06 = Nothing
Set rs_Periodo = Nothing
Set rs_Reporte = Nothing
Set rs_Domicilio = Nothing


Exit Sub
CE:
    HuboError = True
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo sql Ejecutado: " & StrSql
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
    
        Flog.writeline "Parametros: " & parametros
    
        pos1 = 1
        pos2 = InStr(pos1, parametros, ".") - 1
        Nroliq = CLng(Mid(parametros, pos1, pos2))
    
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, ".") - 1
        Todos_Pro = CBool(Mid(parametros, pos1, pos2 - pos1 + 1))
        
        If Not Todos_Pro Then
            pos1 = pos2 + 2
            pos2 = InStr(pos1, parametros, ".") - 1
            NroProc = Mid(parametros, pos1, pos2 - pos1 + 1)
            ListaNroProc = Replace(NroProc, "-", ",")
'            NroProc = CLng(Mid(parametros, pos1, pos2 - pos1 + 1))
        Else
            NroProc = 0
        End If
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, ".") - 1
        Proc_Aprob = Mid(parametros, pos1, pos2 - pos1 + 1)
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, ".") - 1
        Empresa = Mid(parametros, pos1, pos2 - pos1 + 1)
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, ".") - 1
        Todos_Sindicatos = CBool(Mid(parametros, pos1, pos2 - pos1 + 1))
        
        If Not Todos_Sindicatos Then
            pos1 = pos2 + 2
            pos2 = InStr(pos1, parametros, ".") - 1
            Nro_Sindicato = Mid(parametros, pos1, pos2 - pos1 + 1)
        End If
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, ".") - 1
        Agrup = Mid(parametros, pos1, pos2 - pos1 + 1)
        
        Select Case Agrup
        Case 1:
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
            TextoAgrupado = ""
        Case Else
            pos1 = pos2 + 2
            pos2 = Len(parametros)
            TextoAgrupado = Mid(parametros, pos1, pos2 - pos1 + 1)
        End Select
    End If
End If

'seteo el tipo de estructura por default para hacer el join con empleados
TipoEstructura = 16

Select Case Agrup
Case 1: 'Por Estructuras
    Call Suesin03_Estructuras(bpronro, Nroliq, Todos_Pro, Proc_Aprob, Empresa, Todos_Sindicatos, Nro_Sindicato, Agrup, True, AgrupaTE1, Tenro1, Estrnro1, AgrupaTE2, Tenro2, Estrnro2, AgrupaTE3, Tenro3, Estrnro3)
Case 2: 'FAECYS
    Call Suesin03(bpronro, Nroliq, Todos_Pro, Proc_Aprob, Empresa, Todos_Sindicatos, Nro_Sindicato, Agrup)
Case 3: 'Provincia
    Call Suesin03_Provincia(bpronro, Nroliq, Todos_Pro, Proc_Aprob, Empresa, Todos_Sindicatos, Nro_Sindicato, Agrup)
Case 4, 5: '"UOCRA", "UOCRA"
    Call Suesin04(bpronro, Nroliq, Todos_Pro, Proc_Aprob, Empresa, Todos_Sindicatos, Nro_Sindicato, Agrup, TextoAgrupado)
End Select

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

