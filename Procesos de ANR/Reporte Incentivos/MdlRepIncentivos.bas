Attribute VB_Name = "MdlRepIncentivos"
Option Explicit

'Const Version = "1.01"
'Const FechaVersion = "22/07/2009"
''Autor = FGZ
''   Version Inicial


'Const Version = "1.02"
'Const FechaVersion = "11/08/2009"
''Autor = FGZ
''   Se le agregó unos controles sobre los nombres y apellidos dado que tienen comillas simples cargadas en los nombres

'Const Version = "1.03"
'Const FechaVersion = "02/09/2009"
''Autor = FGZ
''   Se Hicieron algunos cambios
''       La columna de S. Bruto
''               1ero se busca el acumulador mensual y si no lo encuentra busca el item de remuneracion
''
''       En la columna PAGOS AL AÑO deberá mostrarse cuantas veces x año se paga el item,
''       por lo tanto debera al informar periodicidad
''       trimestral,  4 veces al año (4 trimestres)
''       mensual, 12 veces al año
''       semestral, 2 veces al año
''       anual, 1 vez al año
''
''       en las columnas target debe mostrarse el resultado del calculo de:
''       MONTO x CANTI DE VECES QUE SE PAGO X 1.0833 con excepción de los ítems que se abonan 1 vez al año
''       en cuyo caso el cálculo es: MONTO * 1


'Const Version = "1.04"
'Const FechaVersion = "25/09/2009"
''Autor = FGZ
''   Estaba calculando mal la cantidad de columans configurables y eso hacia que se desfasara el reporte


Const Version = "1.05"
Const FechaVersion = "07/10/2009"
'Autor = FGZ
'       en las columnas target debe mostrarse solo el valor del ITEM de remuneracion:



'------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------
Public Type TipoTarget
    Item As Long
    Descripcion As String
    Periodicidad As Integer
    Valor As Double
End Type

Public Type TipoDet
    tipo As String
    Origen As String
    Valor As Double
End Type


Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial del Generador de Reporte de Hs liquidadas.
' Autor      :
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim objconnMain As New ADODB.Connection
Dim strCmdLine
Dim Nombre_Arch As String
Dim rs_batch_proceso As New ADODB.Recordset
Dim PID As String
Dim bprcparam As String
Dim iduser As String
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

    Nombre_Arch = PathFLog & "Reporte_Incentivos" & "-" & NroProcesoBatch & ".log"
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaVersion
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-------------------------------------------------"
    Flog.writeline


    On Error Resume Next
    'Abro la conexion
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
    On Error GoTo 0
    
    On Error GoTo ME_Main
    
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0 ,bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Flog.writeline "Pone el estado en procesando"
    
    'Obtengo los datos del proceso
    'StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 236 AND bpronro =" & NroProcesoBatch
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        bprcparam = rs_batch_proceso!bprcparam
        iduser = rs_batch_proceso!iduser
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        Call Calcular(NroProcesoBatch, bprcparam, iduser)
    Else
        Flog.writeline Espacios(Tabulador * 0) & "No encontró el proceso " & NroProcesoBatch
    End If
    
    TiempoFinalProceso = GetTickCount
    Flog.writeline Espacios(Tabulador * 0) & "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    
    If Not HuboError Then
        Progreso = 100
        'UpdateProgreso (Progreso)
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
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


Public Sub Calcular(ByVal bpronro As Long, ByVal Parametros As String, ByVal iduser As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del reporte
' Autor      : FGZ
' Fecha      : 20/01/2004
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Const multiplicador = 1.08333

Dim ArrParametros
Dim ArrTargets

Dim Legdesde As Long
Dim Leghasta As Long
Dim Estado As Integer
Dim Tenro1 As String
Dim Estrnro1 As String
Dim Tenro2 As String
Dim Estrnro2 As String
Dim Tenro3 As String
Dim Estrnro3 As String
Dim ListadoASP As String

Dim ColConfrep As Integer
Dim ColumnasConfrep() As Variant
Dim Fecha As Date
Dim FechaInicial As Date
Dim FechaComp As Date
Dim TipoRep As Integer
Dim Orden As String
Dim Ordenado As String

Dim I      As Integer
Dim J      As Integer
Dim K      As Integer
Dim Y      As Integer
Dim CantRegistros
Dim Fila As Long
Dim MaxCols As Integer
Dim MaxColsF As Integer
'Dim Aux_MaxCols As Integer
Dim MaxItems As Integer
Dim T1 As Double
Dim T2 As Double
Dim T3 As Double
Dim Fijo As Double
Dim FijoAux As Double
Dim Incremento As Double
Dim FechaInc As Date
Dim Motivo As String


Dim Target(1 To 5) As TipoTarget
Dim Columna(1 To 10, 1 To 3) As TipoDet

Dim rs_Empleados As New ADODB.Recordset
Dim rs_Confrep As New ADODB.Recordset
Dim rs_Items As New ADODB.Recordset
Dim rs_Val As New ADODB.Recordset
Dim rs_Remu As New ADODB.Recordset


    On Error GoTo CE

MyBeginTrans

    'Parametros = "1@1@10@-1@0@0@0@0@0@0@01/07/2008@01/07/2009@L@A@1,2"

    'Levanto cada parametro por separado, el separador de parametros es "."
    Flog.writeline Espacios(Tabulador * 1) & "Parametros " & Parametros
    ArrParametros = Split(Parametros, "@")
          
    TipoRep = CLng(ArrParametros(0))
    Legdesde = CLng(ArrParametros(1))
    Leghasta = CLng(ArrParametros(2))
    Estado = CInt(ArrParametros(3))
    
    If Not EsNulo(ArrParametros(4)) Then
        Tenro1 = CLng(ArrParametros(4))
    Else
        Tenro1 = 0
    End If
    If Not EsNulo(ArrParametros(5)) Then
        Estrnro1 = CLng(ArrParametros(5))
    Else
        Estrnro1 = 0
    End If
    
    If Not EsNulo(ArrParametros(6)) Then
        Tenro2 = CLng(ArrParametros(6))
    Else
        Tenro2 = 0
    End If
    If Not EsNulo(ArrParametros(7)) Then
        Estrnro2 = CLng(ArrParametros(7))
    Else
        Estrnro2 = 0
    End If
    
    If Not EsNulo(ArrParametros(8)) Then
        Tenro3 = CLng(ArrParametros(8))
    Else
        Tenro3 = 0
    End If
    If Not EsNulo(ArrParametros(9)) Then
        Estrnro3 = CLng(ArrParametros(9))
    Else
        Estrnro3 = 0
    End If
    
    FechaInicial = CDate(ArrParametros(10))
    Fecha = FechaInicial
    FechaComp = CDate(ArrParametros(11))
    Orden = CStr(ArrParametros(12))
    Ordenado = CStr(ArrParametros(13))
    If Ordenado = "A" Then
        Ordenado = " "
    Else
        Ordenado = " DESC"
    End If
    ListadoASP = CStr(ArrParametros(14))
    
    Flog.writeline Espacios(Tabulador * 1) & "Terminó de levantar los parametros"


ArrTargets = Split(ListadoASP, ",")
MaxItems = UBound(ArrTargets) + 1




Target(1).Item = CLng(ArrTargets(0))
If UBound(ArrTargets) > 0 Then
    Target(2).Item = CLng(ArrTargets(1))
Else
    Target(2).Item = 0
End If
If UBound(ArrTargets) > 1 Then
    Target(3).Item = CLng(ArrTargets(2))
Else
    Target(3).Item = 0
End If
If UBound(ArrTargets) > 2 Then
    Target(4).Item = CLng(ArrTargets(3))
Else
    Target(4).Item = 0
End If
If UBound(ArrTargets) > 3 Then
    Target(5).Item = CLng(ArrTargets(4))
Else
    Target(5).Item = 0
End If

For I = 1 To 10
    For J = 1 To 3
        Columna(I, J).tipo = ""
        Columna(I, J).Origen = "0"
        Columna(I, J).Valor = 0
    Next J
Next I

Flog.writeline
Flog.writeline "Busco la configuracion del reporte."
'Armo el sector de detalle de un reporte
StrSql = "SELECT distinct confnrocol,confval, confval2, confetiq, conftipo FROM confrep where repnro = 255 order by confnrocol, confval2"
OpenRecordset StrSql, rs_Confrep
If Not rs_Confrep.EOF Then
    rs_Confrep.MoveFirst
    Do
        If rs_Confrep!confnrocol > 0 And rs_Confrep!confnrocol <= 10 Then
            If Not EsNulo(rs_Confrep!confval2) Then
                Columna(rs_Confrep!confnrocol, CLng(rs_Confrep!confval2)).tipo = rs_Confrep!conftipo
                Columna(rs_Confrep!confnrocol, rs_Confrep!confval2).Origen = Columna(rs_Confrep!confnrocol, rs_Confrep!confval2).Origen & "," & rs_Confrep!confval
            Else
                Columna(rs_Confrep!confnrocol, 1).tipo = rs_Confrep!conftipo
                Columna(rs_Confrep!confnrocol, 1).Origen = Columna(rs_Confrep!confnrocol, 1).Origen & "," & rs_Confrep!confval
            End If
        End If
        rs_Confrep.MoveNext
    Loop Until rs_Confrep.EOF
    MaxCols = rs_Confrep.RecordCount
    If MaxCols > 10 Then
        MaxCols = 10
    End If
Else
    Flog.writeline "No se encontraron acumuladores configurados para el sector de detalles del reporte. Abortando"
    Exit Sub
End If
'Aux_MaxCols = MaxCols

For I = 1 To MaxItems
    'Busco las descripciones
    StrSql = "SELECT remitedesabr FROM remu_items "
    StrSql = StrSql & " WHERE remitenro = " & Target(I).Item
    OpenRecordset StrSql, rs_Items
    If Not rs_Items.EOF Then
        Target(I).Descripcion = rs_Items!remitedesabr
    Else
        Target(I).Descripcion = ""
    End If
Next I



'Calculo la cantidad de columans dinamicas
StrSql = "SELECT distinct confnrocol FROM confrep where repnro = 255"
OpenRecordset StrSql, rs_Confrep
If Not rs_Confrep.EOF Then
    MaxColsF = rs_Confrep.RecordCount
    If MaxColsF > 10 Then
        MaxColsF = 10
    End If
Else
    MaxColsF = 1
End If


Flog.writeline
Flog.writeline "Busco los empleados afectados."

'Me guardo la consulta de empleados en StrEmpleado
If Tenro3 <> "" And Tenro3 <> "0" Then ' esto ocurre solo cuando se seleccionan los tres niveles
    StrSql = " SELECT DISTINCT empleado.ternro, empleado.empleg, empleado.terape, empleado.ternom "
    StrSql = StrSql & " FROM empleado  "
    StrSql = StrSql & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro  AND estact1.tenro  = " & Tenro1
    StrSql = StrSql & " AND (estact1.htetdesde<=" & ConvFecha(Fecha) & " AND (estact1.htethasta is null or estact1.htethasta>=" & ConvFecha(Fecha) & "))"
    If Estrnro1 <> "" And Estrnro1 <> "0" And Estrnro1 <> "-1" Then 'cuando se le asigna un valor al nivel 1
        StrSql = StrSql & " AND estact1.estrnro =" & Estrnro1
    End If
    StrSql = StrSql & " INNER JOIN his_estructura estact2 ON empleado.ternro = estact2.ternro  AND estact2.tenro  = " & Tenro2
    StrSql = StrSql & " AND (estact2.htetdesde<=" & ConvFecha(Fecha) & " AND (estact2.htethasta is null or estact2.htethasta>=" & ConvFecha(Fecha) & "))"
    If Estrnro2 <> "" And Estrnro2 <> "0" And Estrnro2 <> "-1" Then 'cuando se le asigna un valor al nivel 2
        StrSql = StrSql & " AND estact2.estrnro =" & Estrnro2
    End If
    StrSql = StrSql & " INNER JOIN his_estructura estact3 ON empleado.ternro = estact3.ternro  AND estact3.tenro  = " & Tenro3 & _
    " AND (estact3.htetdesde<=" & ConvFecha(Fecha) & " AND (estact3.htethasta is null or estact3.htethasta>=" & ConvFecha(Fecha) & "))"
    If Estrnro3 <> "" And Estrnro3 <> "0" And Estrnro3 <> "-1" Then 'cuando se le asigna un valor al nivel 3
        StrSql = StrSql & " AND estact3.estrnro =" & Estrnro3
    End If
    If Estado = 1 Then
        StrSql = StrSql & " WHERE " & "(empleg >= " & Legdesde & ") AND (empleg <= " & Leghasta & ")"
    Else
        StrSql = StrSql & " WHERE " & "(empleg >= " & Legdesde & ") AND (empleg <= " & Leghasta & ") AND (empest = " & Estado & ")"
    End If
Else
    If Tenro2 <> "" And Tenro2 <> "0" Then ' ocurre cuando se selecciono hasta el segundo nivel
        StrSql = "SELECT DISTINCT empleado.ternro, empleado.empleg, empleado.terape, empleado.ternom"
        StrSql = StrSql & " FROM empleado  "
        StrSql = StrSql & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro  AND estact1.tenro  = " & Tenro1
        StrSql = StrSql & " AND (estact1.htetdesde<=" & ConvFecha(Fecha) & " AND (estact1.htethasta is null or estact1.htethasta>=" & ConvFecha(Fecha) & "))"
        If Estrnro1 <> "" And Estrnro1 <> "0" And Estrnro1 <> "-1" Then
            StrSql = StrSql & " AND estact1.estrnro =" & Estrnro1
        End If
        StrSql = StrSql & " INNER JOIN his_estructura estact2 ON empleado.ternro = estact2.ternro  AND estact2.tenro  = " & Tenro2 & _
        " AND (estact2.htetdesde<=" & ConvFecha(Fecha) & " AND (estact2.htethasta is null or estact2.htethasta>=" & ConvFecha(Fecha) & "))"
        If Estrnro2 <> "" And Estrnro2 <> "0" And Estrnro2 <> "-1" Then
            StrSql = StrSql & " AND estact2.estrnro =" & Estrnro2
        End If
        If Estado = 1 Then
            StrSql = StrSql & " WHERE " & "(empleg >= " & Legdesde & ") AND (empleg <= " & Leghasta & ")"
        Else
            StrSql = StrSql & " WHERE " & "(empleg >= " & Legdesde & ") AND (empleg <= " & Leghasta & ") AND (empest = " & Estado & ")"
        End If
    Else
        If Tenro1 <> "" And Tenro1 <> "0" Then ' Cuando solo selecionamos el primer nivel
            StrSql = "SELECT DISTINCT empleado.ternro, empleado.empleg, empleado.terape, empleado.ternom "
            StrSql = StrSql & " FROM empleado  "
            StrSql = StrSql & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro  AND estact1.tenro  = " & Tenro1
            StrSql = StrSql & " AND (estact1.htetdesde<=" & ConvFecha(Fecha) & " AND (estact1.htethasta is null or estact1.htethasta>=" & ConvFecha(Fecha) & "))"
            If Estrnro1 <> "" And Estrnro1 <> "0" And Estrnro1 <> "-1" Then
                StrSql = StrSql & " AND estact1.estrnro =" & Estrnro1
            End If
            StrSql = StrSql & " WHERE " & "(empleg >= " & Legdesde & ") AND (empleg <= " & Leghasta & ")"
        Else ' cuando no hay nivel de estructura seleccionado
            StrSql = " SELECT DISTINCT empleado.ternro, empleado.empleg, empleado.terape, empleado.ternom "
            StrSql = StrSql & " FROM empleado  "
            If Estado = 1 Then
                StrSql = StrSql & " WHERE " & "(empleg >= " & Legdesde & ") AND (empleg <= " & Leghasta & ")"
            Else
                StrSql = StrSql & " WHERE " & "(empleg >= " & Legdesde & ") AND (empleg <= " & Leghasta & ") AND (empest = " & Estado & ")"
            End If
        End If
    End If
End If
If Orden = "L" Then
    StrSql = StrSql & " ORDER BY empleg " & Ordenado
Else
    StrSql = StrSql & " ORDER BY terape " & Ordenado & ", ternom " & Ordenado
End If

StrSql = StrSql & ""
'Una vez q tengo los empleados filtro en el intervalo de período seleccionado por el usuario
OpenRecordset StrSql, rs_Empleados

'Calculo el incremento del progreso
If Not rs_Empleados.EOF Then
    If rs_Empleados.RecordCount = 0 Then
        CantRegistros = 1
    Else
        CantRegistros = rs_Empleados.RecordCount
    End If
Else
    CantRegistros = 1
End If
IncPorc = 99 / (CantRegistros * TipoRep)
Fila = 0

Flog.writeline
Flog.writeline "Procesando " & CantRegistros & " Empleados."
Flog.writeline
Do While Not rs_Empleados.EOF
    Flog.writeline "Procesando empleado " & rs_Empleados!empleg
    
    
    'MaxCols = Aux_MaxCols
    Fila = Fila + 1
    For Y = 1 To TipoRep
        
        'Inicializo columnas
        For I = 1 To MaxCols
            For J = 1 To 3
                Columna(I, J).Valor = 0
            Next J
        Next I
        
        For I = 1 To MaxItems
            Target(I).Periodicidad = 0
            Target(I).Valor = 0
        Next I
        
        'Rresuelvo las columnas
        For I = 1 To MaxCols
            For J = 1 To 3
                Select Case UCase(Columna(I, J).tipo)
                Case "AC":  'Acumulador mensual
                    StrSql = "SELECT ammonto FROM acu_mes "
                    StrSql = StrSql & " WHERE ternro = " & rs_Empleados!ternro
                    StrSql = StrSql & " AND acu_mes.acunro IN (" & Columna(I, J).Origen & ")"
                    StrSql = StrSql & " AND ( ammes = " & Month(Fecha) & " AND amanio = " & Year(Fecha) & ")"
                    OpenRecordset StrSql, rs_Val
                    If Not rs_Val.EOF Then
                        Columna(I, J).Valor = rs_Val!ammonto
                        
                        'No sigo buscando
                        J = 3
                        I = MaxCols
                    End If
                Case "ITE": 'Item de remuneracion
                    StrSql = "SELECT vpactado, remcant FROM remu_emp "
                    StrSql = StrSql & " INNER JOIN remu_per ON remu_emp.remperiod = remu_per.rempernro "
                    StrSql = StrSql & " WHERE ternro = " & rs_Empleados!ternro
                    StrSql = StrSql & " AND remitenro IN (" & Columna(I, J).Origen & ")"
                    StrSql = StrSql & " AND remdesde <= " & ConvFecha(Fecha)
                    StrSql = StrSql & " AND (remhasta >= " & ConvFecha(Fecha) & " OR remhasta IS NULL)"
                    OpenRecordset StrSql, rs_Remu
                    If Not rs_Remu.EOF Then
                        Columna(I, J).Valor = IIf(Not EsNulo(rs_Remu!vpactado), rs_Remu!vpactado, 0)
                        'No sigo buscando
                        J = 3
                        I = MaxCols
                    End If
                Case Else
                End Select
            Next J
        Next I
    
    
        'Resuelvo los targets
        For I = 1 To MaxItems
            StrSql = "SELECT vpactado, remcant FROM remu_emp "
            StrSql = StrSql & " INNER JOIN remu_per ON remu_emp.remperiod = remu_per.rempernro "
            StrSql = StrSql & " WHERE ternro = " & rs_Empleados!ternro
            StrSql = StrSql & " AND remitenro = " & Target(I).Item
            StrSql = StrSql & " AND remdesde <= " & ConvFecha(Fecha)
            StrSql = StrSql & " AND (remhasta >= " & ConvFecha(Fecha) & " OR remhasta IS NULL)"
            OpenRecordset StrSql, rs_Remu
            If Not rs_Remu.EOF Then
                Target(I).Periodicidad = IIf(Not EsNulo(rs_Remu!remcant), rs_Remu!remcant, 0)
                'Target(I).Valor = IIf(Not EsNulo(rs_Remu!vpactado), rs_Remu!vpactado, 0)
                'Target(I).Valor = Target(I).Periodicidad * IIf(Not EsNulo(rs_Remu!vpactado), rs_Remu!vpactado, 0)
                
                'FGZ - 07/10/2009 - otra vez se cambió el calculo. Ahora el cliente lo quiere como estaba originalmente
'                If Target(I).Periodicidad = 1 Then
'                    Target(I).Valor = Target(I).Periodicidad * IIf(Not EsNulo(rs_Remu!vpactado), rs_Remu!vpactado, 0)
'                Else
'                    Target(I).Valor = multiplicador * Target(I).Periodicidad * IIf(Not EsNulo(rs_Remu!vpactado), rs_Remu!vpactado, 0)
'                End If
                
                'FGZ - 07/10/2009 - otra vez se cambió el calculo. Ahora el cliente lo quiere como estaba originalmente
                Target(I).Valor = IIf(Not EsNulo(rs_Remu!vpactado), rs_Remu!vpactado, 0)
                
            End If
        Next I
       
        
        'Resuelvo los totales
        T1 = (Columna(1, 1).Valor + Columna(1, 2).Valor + Columna(1, 3).Valor) * 13
        'T2 = (Target(1).Valor * Target(1).Periodicidad * multiplicador) + (Target(2).Valor * Target(2).Periodicidad * multiplicador)
        'FGZ - 07/10/20009 - cambié la forma de calculo ademas de agregar hasta el item 5
        
        'Target 1
        I = 1
        If Target(I).Periodicidad = 1 Then
            T2 = Target(I).Valor
        Else
            T2 = (Target(I).Valor * Target(I).Periodicidad * multiplicador)
        End If
        
        'Item 2
        I = 2
        If Target(I).Periodicidad = 1 Then
            T2 = T2 + Target(I).Valor
        Else
            T2 = T2 + (Target(I).Valor * Target(I).Periodicidad * multiplicador)
        End If
        
        'Item 3
        I = 3
        If Target(I).Periodicidad = 1 Then
            T2 = T2 + Target(I).Valor
        Else
            T2 = T2 + (Target(I).Valor * Target(I).Periodicidad * multiplicador)
        End If
        
        
        'Item 4
        I = 4
        If Target(I).Periodicidad = 1 Then
            T2 = T2 + Target(I).Valor
        Else
            T2 = T2 + (Target(I).Valor * Target(I).Periodicidad * multiplicador)
        End If
        
        
        'Item 5
        I = 5
        If Target(I).Periodicidad = 1 Then
            T2 = T2 + Target(I).Valor
        Else
            T2 = T2 + (Target(I).Valor * Target(I).Periodicidad * multiplicador)
        End If
        
        
        
        
        
        
        T3 = T1 + T2
        Fijo = Round(T3 / 13)
        
        'Calculo el incremento
        FechaInc = Fecha
        If Y = 2 Then
            If FijoAux = 0 Or Fijo = 0 Then
                If Fijo = 0 Then
                    Incremento = 0
                Else
                    Incremento = 100
                End If
            Else
                If Fijo > FijoAux Then
                    Incremento = (Fijo * 100 / FijoAux) - 100
                Else
                    Incremento = 100 - (Fijo * 100 / FijoAux)
                End If
            End If
            Incremento = Round(Incremento)
            
            'Busco la fecha del incremento
            FechaInc = Fecha
            
            'Busco el motivo del incremento
            Motivo = "mmmm"
            
            
        Else
            FijoAux = Fijo
        End If
      
        Flog.writeline " Insertando.. "
        
        'inserto en la tabla del reporte
        StrSql = "INSERT INTO rep_incentivos (bpronro,fecha,hora,iduser,filanro"
        StrSql = StrSql & ",ternro,legajo,apellido,nombre"
        StrSql = StrSql & ",fechaAn,detalle, tiporep"
        'Columnas variables
        For I = 1 To MaxColsF
            StrSql = StrSql & ",col" & I
        Next I
        'Targets
        For I = 1 To MaxItems
            StrSql = StrSql & ",ite_desc_" & I
            StrSql = StrSql & ",ite_pag_" & I
            StrSql = StrSql & ",ite_val_" & I
        Next I
        'Totales
        StrSql = StrSql & ",rem1, rem2, rem3"
        StrSql = StrSql & ",fijo, increm, fechainc, motivoinc"
        
        StrSql = StrSql & ") VALUES ("
        StrSql = StrSql & NroProcesoBatch
        StrSql = StrSql & "," & ConvFecha(Date)
        StrSql = StrSql & ",'" & Mid(Time, 1, 8) & "'"
        StrSql = StrSql & ",'" & iduser & "'"
        StrSql = StrSql & "," & Fila
        
        StrSql = StrSql & "," & rs_Empleados!ternro
        StrSql = StrSql & "," & rs_Empleados!empleg
        StrSql = StrSql & ",'" & Replace(rs_Empleados!terape, "'", " ") & "'"
        StrSql = StrSql & ",'" & Replace(rs_Empleados!ternom, "'", " ") & "'"
        
        StrSql = StrSql & "," & ConvFecha(Fecha)
        StrSql = StrSql & "," & Y
        StrSql = StrSql & "," & TipoRep
        
        'Columnas variables
        'For I = 1 To MaxCols
        For I = 1 To MaxColsF
            'StrSql = StrSql & "," & Columna(I).Valor
            StrSql = StrSql & "," & (Columna(I, 1).Valor + Columna(I, 2).Valor + Columna(I, 3).Valor)
        Next I
        'Targets
        For I = 1 To MaxItems
            StrSql = StrSql & ",'" & Target(I).Descripcion & "'"
            StrSql = StrSql & "," & Target(I).Periodicidad
            StrSql = StrSql & "," & Target(I).Valor
        Next I
        'Totales
        StrSql = StrSql & "," & T1
        StrSql = StrSql & "," & T2
        StrSql = StrSql & "," & T3
        
        'Fijos e Incrementos
        StrSql = StrSql & "," & Fijo
        StrSql = StrSql & "," & Incremento
        StrSql = StrSql & "," & ConvFecha(FechaInc)
        StrSql = StrSql & ",'" & Motivo & "'"
        StrSql = StrSql & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
    
        '------------------------
        'Actualizo el progreso
        Progreso = Progreso + IncPorc
        TiempoFinalProceso = GetTickCount
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
                 ", bprctiempo ='" & Left(CStr((TiempoAcumulado - TiempoInicialProceso)), 10) & _
                 "' WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
        
        Fecha = FechaComp
    Next Y
    Fecha = FechaInicial
    rs_Empleados.MoveNext
Loop

'------------------------
'------------------------

'Fin de la transaccion
If Not HuboError Then
    MyCommitTrans
Else
    MyRollbackTrans
End If


If rs_Empleados.State = adStateOpen Then rs_Empleados.Close
If rs_Confrep.State = adStateOpen Then rs_Confrep.Close

Set rs_Empleados = Nothing
Set rs_Confrep = Nothing

Exit Sub
CE:
    Flog.writeline "=================================================================="
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="
    MyRollbackTrans
    MyBeginTrans
    Progreso = Round(Progreso + IncPorc, 2)
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
             "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    MyCommitTrans
    
    HuboError = True
    Flog.writeline " Error: " & Err.Description

End Sub


