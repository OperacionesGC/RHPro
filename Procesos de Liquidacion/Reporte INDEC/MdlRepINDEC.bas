Attribute VB_Name = "MdlRepINDEC"
Option Explicit

'Const Version = "1.00"
'Const FechaVersion = "30/05/2008"
'Autor = Martin Ferraro
'Reporte INDEC

Const Version = "1.01"
Const FechaVersion = "31/07/2009" 'Martin Ferraro - Encriptacion de string connection


Public Type TConfRep
    tipo As String
    Cod As String
End Type

'Arreglos
Dim ArrConfRep(16, 11) As TConfRep
Dim ArrEtiq(16) As String

'Ind Arr
Dim indArrFila As Long
Dim indArrcol As Long

'Globales
Dim listaJornal As String
Dim listaMensual As String
Dim listaResto As String
Dim EtiqJornal As String
Dim EtiqMensual As String
Dim EtiqResto As String
Dim PeriodoDesde As Date
Dim PeriodoHasta As Date
Dim periodoAntDesde As Date
Dim periodoAntHasta As Date
Dim encontroPerAnt As Boolean
Dim teDesc1 As String
Dim teDesc2 As String
Dim teDesc3 As String
Dim FiltroSql As String
Dim titulo As String
Dim EmpreNomb As String
Dim EmpresaTernro As Long
Dim Empresa As Long
Dim Periodo As Long
Dim PeriodoDesc As String
Dim ordenCab As Long


Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial del Generador de Reporte.
' Autor      : Martin Ferraro
' Fecha      : 30/05/2008
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim objconnMain As New ADODB.Connection
Dim strCmdLine
Dim Nombre_Arch As String
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

    On Error GoTo ME_Main
    Nombre_Arch = PathFLog & "Rep_INDEC" & "-" & NroProcesoBatch & ".log"
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
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If
    OpenConnection strconexion, objconnProgreso
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If
    On Error GoTo 0
    
    On Error GoTo ME_Main
    
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0 ,bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Flog.writeline "Pone el estado en procesando"
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 217 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        bprcparam = rs_batch_proceso!bprcparam
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        Call RepINDEC(NroProcesoBatch, bprcparam)
    Else
        Flog.writeline "no encontró el proceso"
    End If
    
    TiempoFinalProceso = GetTickCount
    Flog.writeline "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    
    If Not HuboError Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
    End If
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
Fin:
    Flog.Close
    'If objConn.State = adStateOpen Then objConn.Close
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


Public Sub RepINDEC(ByVal bpronro As Long, ByVal Parametros As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del reporte INDEC
' Autor      : Martin Ferraro
' Fecha      : 02/04/2008
' --------------------------------------------------------------------------------------------


'Arreglo que contiene los parametros
Dim arrParam
Dim I As Long

'Parametros desde ASP
Dim Tenro1 As Long
Dim Estrnro1 As Long
Dim Tenro2 As Long
Dim Estrnro2 As Long
Dim Tenro3 As Long
Dim Estrnro3 As Long

'RecordSet
Dim rs_Consult As New ADODB.Recordset
Dim rs_Tenro1 As New ADODB.Recordset
Dim rs_Tenro2 As New ADODB.Recordset
Dim rs_Tenro3 As New ADODB.Recordset

'Variables
Dim hayTenro1 As Boolean
Dim hayTenro2 As Boolean
Dim hayTenro3 As Boolean

Dim guardarUlt As Boolean

'Inicio codigo ejecutable
On Error GoTo CE


'------------------------------------------------------------------------------------
' Levanto cada parametro por separado, el separador de parametros es "@"
'------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 0) & "levantando parametros" & Parametros
If Not IsNull(Parametros) Then
    
    arrParam = Split(Parametros, "@")
    
    If UBound(arrParam) = 9 Then
        
        FiltroSql = arrParam(0)
        Periodo = CLng(arrParam(1))
        Tenro1 = CLng(arrParam(2))
        Estrnro1 = CLng(arrParam(3))
        Tenro2 = CLng(arrParam(4))
        Estrnro2 = CLng(arrParam(5))
        Tenro3 = CLng(arrParam(6))
        Estrnro3 = CLng(arrParam(7))
        Empresa = CLng(arrParam(8))
        titulo = arrParam(9)
        
        Flog.writeline Espacios(Tabulador * 1) & "Filtro = " & FiltroSql
        Flog.writeline Espacios(Tabulador * 1) & "Empresa = " & Empresa
        Flog.writeline Espacios(Tabulador * 1) & "Periodo = " & Periodo
        Flog.writeline Espacios(Tabulador * 1) & "TE1 = " & Tenro1
        Flog.writeline Espacios(Tabulador * 1) & "Estr1 = " & Estrnro1
        Flog.writeline Espacios(Tabulador * 1) & "TE2 = " & Tenro2
        Flog.writeline Espacios(Tabulador * 1) & "Estr2 = " & Estrnro2
        Flog.writeline Espacios(Tabulador * 1) & "TE3 = " & Tenro3
        Flog.writeline Espacios(Tabulador * 1) & "Estr3 = " & Estrnro3
        Flog.writeline Espacios(Tabulador * 1) & "Titulo = " & titulo
    Else
        Flog.writeline Espacios(Tabulador * 0) & "ERROR. La cantidad de parametros no es la esperada."
        HuboError = True
        Exit Sub
        
    End If
Else
    Flog.writeline Espacios(Tabulador * 0) & "ERROR. No se encuentran los paramentros."
    HuboError = True
    Exit Sub
End If

Flog.writeline


'------------------------------------------------------------------------------------
'Configuracion del Reporte
'------------------------------------------------------------------------------------
Flog.writeline
Flog.writeline Espacios(Tabulador * 0) & "Buscando configuración de Reporte 229."
listaJornal = "0"
listaMensual = "0"
listaResto = "0"
EtiqJornal = ""
EtiqMensual = ""
EtiqResto = ""
indArrFila = 1
indArrcol = 0

StrSql = "SELECT *"
StrSql = StrSql & " FROM confrep"
StrSql = StrSql & " WHERE repnro = 229"
StrSql = StrSql & " ORDER BY confnrocol"
OpenRecordset StrSql, rs_Consult

If rs_Consult.EOF Then
    Flog.writeline Espacios(Tabulador * 1) & "No se encontró la configuración del Reporte"
    Exit Sub
End If
    
guardarUlt = False
Do While (Not rs_Consult.EOF)
        
    Select Case rs_Consult!confnrocol
        Case 1
            listaJornal = listaJornal & "," & rs_Consult!confval
            EtiqJornal = IIf(EsNulo(rs_Consult!confetiq), "", rs_Consult!confetiq)
        Case 2
            listaMensual = listaMensual & "," & rs_Consult!confval
            EtiqMensual = IIf(EsNulo(rs_Consult!confetiq), "", rs_Consult!confetiq)
        Case 3
            listaResto = listaResto & "," & rs_Consult!confval
            EtiqResto = IIf(EsNulo(rs_Consult!confetiq), "", rs_Consult!confetiq)
        Case 4 To 19
            'Armo la matriz de conceptos y acumuladores a buscar
            'Filas = distintas filas del reporte.
            'Columnas = distintos conceptos o acum de la fila
             
             If indArrFila <> (rs_Consult!confnrocol - 3) Then
                'Cambio de fila, guardo el tope de conceptos y acum config en la componente 0 de la fila
                ArrConfRep(indArrFila, 0).tipo = indArrcol
                
                'Cambio de fila de la matriz
                indArrFila = indArrFila + 1
                indArrcol = 1
                guardarUlt = False
             Else
                'Agrego una nueva columna a la misma fila
                indArrcol = indArrcol + 1
                guardarUlt = True
             End If
             
             If (indArrcol <= 10) And (indArrFila <= 15) Then
                'Guardo los valores en la matriz
                ArrConfRep(indArrFila, indArrcol).tipo = rs_Consult!conftipo
                ArrConfRep(indArrFila, indArrcol).Cod = rs_Consult!confval2
                ArrEtiq(indArrFila) = rs_Consult!confetiq
             End If
             
    End Select
    
    rs_Consult.MoveNext

Loop

If guardarUlt Then
    If (indArrcol <= 10) And (indArrFila <= 15) Then ArrConfRep(indArrFila, 0).tipo = indArrcol
End If

rs_Consult.Close
Flog.writeline
Flog.writeline Espacios(Tabulador * 0) & "Matriz de configuraion del confrep"
Flog.writeline
For indArrFila = 1 To 15
    For indArrcol = 1 To 10
        Flog.Write ArrConfRep(indArrFila, indArrcol).tipo & " " & ArrConfRep(indArrFila, indArrcol).Cod & " | "
    Next
    Flog.writeline
Next

Flog.writeline

For indArrFila = 1 To 15
    Flog.Write ArrConfRep(indArrFila, 0).tipo & " | "
Next
Flog.writeline


'Validacion de datos obligatorios del confrep
If listaJornal = "0" Then
    Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se configuraron las estructuras Jornalizados de la columna 1 de la configuracion del reporte 229"
    Exit Sub
End If
If listaMensual = "0" Then
    Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se configuraron las estructuras mensualizados de la columna 2 de la configuracion del reporte 229"
    Exit Sub
End If
If listaResto = "0" Then
    Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se configuraron las estructuras Resto de la columna 3 de la configuracion del reporte 229"
    Exit Sub
End If


'------------------------------------------------------------------------------------
'Datos de la empresa
'------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 0) & "Buscando Empresa."
StrSql = "SELECT estructura.estrdabr, empresa.ternro FROM estructura"
StrSql = StrSql & " INNER JOIN empresa ON empresa.estrnro = estructura.estrnro"
StrSql = StrSql & " WHERE estructura.tenro = 10 AND estructura.estrnro = " & Empresa
OpenRecordset StrSql, rs_Consult

If rs_Consult.EOF Then
    Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontró la empresa."
    Exit Sub
Else
    EmpreNomb = IIf(EsNulo(rs_Consult!estrdabr), "", rs_Consult!estrdabr)
    EmpresaTernro = rs_Consult!ternro
    Flog.writeline Espacios(Tabulador * 1) & "Empresa: " & EmpreNomb
End If
rs_Consult.Close


'------------------------------------------------------------------------------------
'Periodo actual
'------------------------------------------------------------------------------------
Flog.writeline
Flog.writeline Espacios(Tabulador * 0) & "Buscando periodo filtro"
StrSql = "SELECT pliqnro, pliqdesc, pliqdesde, pliqhasta"
StrSql = StrSql & " FROM Periodo"
StrSql = StrSql & " WHERE pliqnro = " & Periodo
OpenRecordset StrSql, rs_Consult
If Not rs_Consult.EOF Then
    PeriodoDesde = rs_Consult!pliqdesde
    PeriodoHasta = rs_Consult!pliqhasta
    PeriodoDesc = rs_Consult!pliqdesc
Else
    Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontro el periodo del filtro."
    Exit Sub
End If
rs_Consult.Close


'------------------------------------------------------------------------------------
'Periodo anterior
'------------------------------------------------------------------------------------
Flog.writeline
Flog.writeline Espacios(Tabulador * 0) & "Buscando periodo anterior"
StrSql = "SELECT pliqnro, pliqdesc, pliqdesde, pliqhasta"
StrSql = StrSql & " From Periodo"
StrSql = StrSql & " WHERE pliqhasta <= " & ConvFecha(PeriodoDesde)
StrSql = StrSql & " ORDER BY pliqhasta DESC"
OpenRecordset StrSql, rs_Consult
If Not rs_Consult.EOF Then
    periodoAntDesde = rs_Consult!pliqdesde
    periodoAntHasta = rs_Consult!pliqhasta
    encontroPerAnt = True
Else
    encontroPerAnt = False
    Flog.writeline Espacios(Tabulador * 1) & "No se encontro el periodo periodo anterior."
End If
rs_Consult.Close


'------------------------------------------------------------------------------------
'Armo la estructura de los ciclo por los tres niveles de estructuras del filtro
'------------------------------------------------------------------------------------
Flog.writeline
Flog.writeline Espacios(Tabulador * 0) & "Buscando los tres niveles de estructura"
hayTenro1 = False
hayTenro2 = False
hayTenro3 = False
teDesc1 = ""
teDesc2 = ""
teDesc3 = ""

'Primer nivel
If Tenro1 <> 0 Then
    StrSql = "SELECT estructura.estrnro, estructura.estrdabr, estructura.tenro, tipoestructura.tedabr"
    StrSql = StrSql & " From estructura"
    StrSql = StrSql & " INNER JOIN tipoestructura ON tipoestructura.tenro = estructura.tenro"
    StrSql = StrSql & " Where tipoestructura.tenro = " & Tenro1
    If Estrnro1 <> 0 Then
        StrSql = StrSql & " AND estructura.estrnro = " & Estrnro1
    End If
    StrSql = StrSql & " ORDER BY estrdabr"
    OpenRecordset StrSql, rs_Tenro1
    If rs_Tenro1.EOF Then
        hayTenro1 = False
        Flog.writeline Espacios(Tabulador * 1) & "No se encontraron estructuras de primer nivel de tipo = " & Tenro1
    Else
        teDesc1 = rs_Tenro1!tedabr
        hayTenro1 = True
    End If
    
    'Segundo nivel
    If Tenro2 <> 0 Then
        StrSql = "SELECT estructura.estrnro, estructura.estrdabr, estructura.tenro, tipoestructura.tedabr"
        StrSql = StrSql & " From estructura"
        StrSql = StrSql & " INNER JOIN tipoestructura ON tipoestructura.tenro = estructura.tenro"
        StrSql = StrSql & " Where tipoestructura.tenro = " & Tenro2
        If Estrnro1 <> 0 Then
            StrSql = StrSql & " AND estructura.estrnro = " & Estrnro2
        End If
        OpenRecordset StrSql, rs_Tenro2
        StrSql = StrSql & " ORDER BY estrdabr"
        If rs_Tenro2.EOF Then
            hayTenro2 = False
            Flog.writeline Espacios(Tabulador * 1) & "No se encontraron estructuras de segundo nivel de tipo = " & Tenro2
        Else
            teDesc2 = rs_Tenro1!tedabr
            hayTenro2 = True
        End If
        
        'Tercer nivel
        If Tenro3 <> 0 Then
            StrSql = "SELECT estructura.estrnro, estructura.estrdabr, estructura.tenro, tipoestructura.tedabr"
            StrSql = StrSql & " From estructura"
            StrSql = StrSql & " INNER JOIN tipoestructura ON tipoestructura.tenro = estructura.tenro"
            StrSql = StrSql & " Where tipoestructura.tenro = " & Tenro3
            If Estrnro1 <> 0 Then
                StrSql = StrSql & " AND estructura.estrnro = " & Estrnro3
            End If
            StrSql = StrSql & " ORDER BY estrdabr"
            OpenRecordset StrSql, rs_Tenro3
            If rs_Tenro3.EOF Then
                hayTenro3 = False
                Flog.writeline Espacios(Tabulador * 1) & "No se encontraron estructuras de tercer nivel de tipo = " & Tenro3
            Else
                teDesc1 = rs_Tenro3!tedabr
                hayTenro3 = True
            End If
            
        End If
        
    End If
    
End If


'------------------------------------------------------------------------------------
'configuracion de las variables de progreso
'------------------------------------------------------------------------------------
Flog.writeline
Flog.writeline Espacios(Tabulador * 0) & "Configurando progreso"
Progreso = 0
CEmpleadosAProc = 1

If hayTenro1 Then
    CEmpleadosAProc = rs_Tenro1.RecordCount
    If hayTenro2 Then
        CEmpleadosAProc = CEmpleadosAProc * rs_Tenro2.RecordCount
        If hayTenro3 Then
            CEmpleadosAProc = CEmpleadosAProc * rs_Tenro3.RecordCount
        End If
    End If
End If

If CEmpleadosAProc = 0 Then
   CEmpleadosAProc = 1
End If
IncPorc = (100 / CEmpleadosAProc)
Flog.writeline Espacios(Tabulador * 1) & "Hojas a analizar: " & CEmpleadosAProc


'------------------------------------------------------------------------------------
'Ciclos principales
'------------------------------------------------------------------------------------
ordenCab = 1
Flog.writeline
Flog.writeline
Flog.writeline
Flog.writeline Espacios(Tabulador * 0) & "--------------------------------------------------------"
Flog.writeline Espacios(Tabulador * 0) & "Comienza el procesamiento de Hojas."
Flog.writeline Espacios(Tabulador * 0) & "--------------------------------------------------------"
Flog.writeline
If (hayTenro1 And hayTenro2 And hayTenro3) Then
    
    'Ciclo por 3 niveles de estructura, genera una hoja por cada conbinacion de estr
    Do While Not rs_Tenro1.EOF
        rs_Tenro2.MoveFirst
        Do While Not rs_Tenro2.EOF
            rs_Tenro3.MoveFirst
            Do While Not rs_Tenro3.EOF
                Flog.writeline
                
                Flog.writeline Espacios(Tabulador * 0) & " TE1 = " & rs_Tenro1!tenro & " Estrnro1 = " & rs_Tenro1!estrnro & " " & IIf(EsNulo(rs_Tenro1!estrdabr), "", rs_Tenro1!estrdabr) & " TE2 = " & rs_Tenro2!tenro & " Estrnro2 = " & rs_Tenro2!estrnro & " " & IIf(EsNulo(rs_Tenro2!estrdabr), "", rs_Tenro2!estrdabr) & " TE3 = " & rs_Tenro3!tenro & " Estrnro3 = " & rs_Tenro3!estrnro & " " & IIf(EsNulo(rs_Tenro3!estrdabr), "", rs_Tenro3!estrdabr)
                Call GenerarHoja(rs_Tenro1!tenro, rs_Tenro1!estrnro, IIf(EsNulo(rs_Tenro1!estrdabr), "", rs_Tenro1!estrdabr), rs_Tenro2!tenro, rs_Tenro2!estrnro, IIf(EsNulo(rs_Tenro2!estrdabr), "", rs_Tenro2!estrdabr), rs_Tenro3!tenro, rs_Tenro3!estrnro, IIf(EsNulo(rs_Tenro3!estrdabr), "", rs_Tenro3!estrdabr))
                
                'Actualizo el progreso-----------------------------------------------------------
                Progreso = Progreso + IncPorc
                CEmpleadosAProc = CEmpleadosAProc - 1
                TiempoAcumulado = GetTickCount
                StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
                         ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
                         "' WHERE bpronro = " & NroProcesoBatch
                objconnProgreso.Execute StrSql, , adExecuteNoRecords
                '--------------------------------------------------------------------------------
                
                rs_Tenro3.MoveNext
            Loop
            rs_Tenro2.MoveNext
        Loop
        rs_Tenro1.MoveNext
    Loop

Else
    If (hayTenro1 And hayTenro2) Then
        
        'Ciclo por 2 niveles de estructura, genera una hoja por cada conbinacion de estr
        Do While Not rs_Tenro1.EOF
            rs_Tenro2.MoveFirst
            Do While Not rs_Tenro2.EOF
                Flog.writeline
                
                Flog.writeline Espacios(Tabulador * 0) & " TE1 = " & rs_Tenro1!tenro & " Estrnro1 = " & rs_Tenro1!estrnro & " " & IIf(EsNulo(rs_Tenro1!estrdabr), "", rs_Tenro1!estrdabr) & " TE2 = " & rs_Tenro2!tenro & " Estrnro2 = " & rs_Tenro2!estrnro & " " & IIf(EsNulo(rs_Tenro2!estrdabr), "", rs_Tenro2!estrdabr)
                Call GenerarHoja(rs_Tenro1!tenro, rs_Tenro1!estrnro, IIf(EsNulo(rs_Tenro1!estrdabr), "", rs_Tenro1!estrdabr), rs_Tenro2!tenro, rs_Tenro2!estrnro, IIf(EsNulo(rs_Tenro2!estrdabr), "", rs_Tenro2!estrdabr), 0, 0, "")
                
                'Actualizo el progreso-----------------------------------------------------------
                Progreso = Progreso + IncPorc
                CEmpleadosAProc = CEmpleadosAProc - 1
                TiempoAcumulado = GetTickCount
                StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
                         ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
                         "' WHERE bpronro = " & NroProcesoBatch
                objconnProgreso.Execute StrSql, , adExecuteNoRecords
                '--------------------------------------------------------------------------------
                
                rs_Tenro2.MoveNext
            Loop
            rs_Tenro1.MoveNext
        Loop
    
    Else
        If (hayTenro1) Then
            
            'Ciclo por un nivel de estructura, genera una hoja por cada conbinacion de estr
            Do While Not rs_Tenro1.EOF
                Flog.writeline
                
                Flog.writeline Espacios(Tabulador * 0) & " TE1 = " & rs_Tenro1!tenro & " Estrnro1 = " & rs_Tenro1!estrnro & " " & IIf(EsNulo(rs_Tenro1!estrdabr), "", rs_Tenro1!estrdabr)
                Call GenerarHoja(rs_Tenro1!tenro, rs_Tenro1!estrnro, IIf(EsNulo(rs_Tenro1!estrdabr), "", rs_Tenro1!estrdabr), 0, 0, "", 0, 0, "")
                
                'Actualizo el progreso-----------------------------------------------------------
                Progreso = Progreso + IncPorc
                CEmpleadosAProc = CEmpleadosAProc - 1
                TiempoAcumulado = GetTickCount
                StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
                         ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
                         "' WHERE bpronro = " & NroProcesoBatch
                objconnProgreso.Execute StrSql, , adExecuteNoRecords
                '--------------------------------------------------------------------------------
                
                rs_Tenro1.MoveNext
            Loop
        
        Else
            'Ciclo sin corte por estructuras, genera una sola hoja
            
            Call GenerarHoja(0, 0, "", 0, 0, "", 0, 0, "")
                    
            'Actualizo el progreso-----------------------------------------------------------
            Progreso = Progreso + IncPorc
            CEmpleadosAProc = CEmpleadosAProc - 1
            TiempoAcumulado = GetTickCount
            StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
                     ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
                     "' WHERE bpronro = " & NroProcesoBatch
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
            '--------------------------------------------------------------------------------
        End If 'Ciclo 1 nivel
    End If 'Ciclo 2 niveles
End If 'Ciclo 3 niveles


If rs_Consult.State = adStateOpen Then rs_Consult.Close
If rs_Tenro1.State = adStateOpen Then rs_Tenro1.Close
If rs_Tenro2.State = adStateOpen Then rs_Tenro2.Close
If rs_Tenro3.State = adStateOpen Then rs_Tenro3.Close

Set rs_Consult = Nothing
Set rs_Tenro1 = Nothing
Set rs_Tenro2 = Nothing
Set rs_Tenro3 = Nothing

Exit Sub

CE:
    Flog.writeline "=================================================================="
    Flog.writeline "Error en RepINDEC"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="
    
        
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
             "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
        
    HuboError = True

End Sub



Public Sub GenerarHoja(ByVal te1 As Long, ByVal estr1 As Long, ByVal estrDesc1 As String, ByVal te2 As Long, ByVal estr2 As Long, ByVal estrDesc2 As String, ByVal te3 As Long, ByVal estr3 As Long, ByVal estrDesc3 As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Graba los datos para las tres estructuras
' Autor      : Martin Ferraro
' Fecha      : 02/04/2008
' --------------------------------------------------------------------------------------------

Dim rs_datos As New ADODB.Recordset
Dim valor1 As Double
Dim valor2 As Double
Dim valor3 As Double
Dim ultimaCabecera As Long
Dim nroTipo As Long

'Inicio codigo ejecutable
On Error GoTo E_GenerarHoja


Flog.writeline
'------------------------------------------------------------------------------------
'Creo la cabecera
'------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 1) & "Creando cabecera."
StrSql = "INSERT INTO rep_indec"
StrSql = StrSql & " (bpronro"
StrSql = StrSql & " ,empnro"
StrSql = StrSql & " ,pliqnro"
StrSql = StrSql & " ,empnomb"
StrSql = StrSql & " ,pliqdesc"
StrSql = StrSql & " ,tenro1"
StrSql = StrSql & " ,tenro2"
StrSql = StrSql & " ,tenro3"
StrSql = StrSql & " ,estrnro1"
StrSql = StrSql & " ,estrnro2"
StrSql = StrSql & " ,estrnro3"
StrSql = StrSql & " ,estrdesc1"
StrSql = StrSql & " ,estrdesc2"
StrSql = StrSql & " ,estrdesc3"
StrSql = StrSql & " ,etiq1"
StrSql = StrSql & " ,etiq2"
StrSql = StrSql & " ,etiq3"
StrSql = StrSql & " ,titulo"
StrSql = StrSql & " ,orden)"
StrSql = StrSql & " Values"
StrSql = StrSql & " (" & NroProcesoBatch
StrSql = StrSql & " ," & Empresa
StrSql = StrSql & " ," & Periodo
StrSql = StrSql & " ,'" & Mid(EmpreNomb, 1, 100) & "'"
StrSql = StrSql & " ,'" & Mid(PeriodoDesc, 1, 100) & "'"
StrSql = StrSql & " ," & te1
StrSql = StrSql & " ," & te2
StrSql = StrSql & " ," & te3
StrSql = StrSql & " ," & estr1
StrSql = StrSql & " ," & estr2
StrSql = StrSql & " ," & estr3
StrSql = StrSql & " ,'" & Mid(estrDesc1, 1, 100) & "'"
StrSql = StrSql & " ,'" & Mid(estrDesc2, 1, 100) & "'"
StrSql = StrSql & " ,'" & Mid(estrDesc3, 1, 100) & "'"
StrSql = StrSql & " ,'" & Mid(EtiqJornal, 1, 100) & "'"
StrSql = StrSql & " ,'" & Mid(EtiqMensual, 1, 100) & "'"
StrSql = StrSql & " ,'" & Mid(EtiqResto, 1, 100) & "'"
StrSql = StrSql & " ,'" & Mid(CStr(NroProcesoBatch) & " - " & EmpreNomb & " - " & PeriodoDesc & " - " & titulo, 1, 400) & "'"
StrSql = StrSql & " ," & ordenCab
StrSql = StrSql & " )"
objConn.Execute StrSql, , adExecuteNoRecords

ordenCab = ordenCab + 1
ultimaCabecera = getLastIdentity(objConn, "rep_indec")


'------------------------------------------------------------------------------------
'Personal mes anterior
'------------------------------------------------------------------------------------
valor1 = 0
valor2 = 0
valor3 = 0
If encontroPerAnt Then
    
    Flog.writeline Espacios(Tabulador * 2) & "Procesando Personal mes anterior."
    
    'Jornalizados
    '------------------------------------------------------------------------------------
    Flog.writeline Espacios(Tabulador * 3) & "Jornalizados."
    StrSql = "SELECT count (distinct empleado.ternro) cant "
    StrSql = StrSql & " FROM empleado"
    'Filtro Empresa
    StrSql = StrSql & " INNER JOIN his_estructura empresa ON empleado.ternro = empresa.ternro"
    StrSql = StrSql & " AND empresa.estrnro = " & Empresa & " AND empresa.tenro = 10"
    StrSql = StrSql & " AND empresa.htetdesde <= " & ConvFecha(periodoAntHasta) & " AND (empresa.htethasta IS NULL OR empresa.htethasta >= " & ConvFecha(periodoAntHasta) & ")"
    'Filtros de niveles de estructura
    If te1 <> 0 Then
        StrSql = StrSql & " INNER JOIN his_estructura tenro1 ON empleado.ternro = tenro1.ternro"
        StrSql = StrSql & " AND tenro1.tenro = " & te1
        StrSql = StrSql & " AND tenro1.htetdesde <= " & ConvFecha(periodoAntHasta) & " AND (tenro1.htethasta IS NULL OR tenro1.htethasta >= " & ConvFecha(periodoAntHasta) & ")"
        StrSql = StrSql & " AND tenro1.estrnro = " & estr1
    End If
    If te2 <> 0 Then
        StrSql = StrSql & " INNER JOIN his_estructura tenro2 ON empleado.ternro = tenro2.ternro"
        StrSql = StrSql & " AND tenro2.tenro = " & te2
        StrSql = StrSql & " AND tenro2.htetdesde <= " & ConvFecha(periodoAntHasta) & " AND (tenro2.htethasta IS NULL OR tenro2.htethasta >= " & ConvFecha(periodoAntHasta) & ")"
        StrSql = StrSql & " AND tenro2.estrnro = " & estr2
    End If
    If te3 <> 0 Then
        StrSql = StrSql & " INNER JOIN his_estructura tenro3 ON empleado.ternro = tenro3.ternro"
        StrSql = StrSql & " AND tenro3.tenro = " & te3
        StrSql = StrSql & " AND tenro3.htetdesde <= " & ConvFecha(periodoAntHasta) & " AND (tenro3.htethasta IS NULL OR tenro3.htethasta >= " & ConvFecha(periodoAntHasta) & ")"
        StrSql = StrSql & " AND tenro3.estrnro = " & estr3
    End If
    'Filtro Jornalizados
    StrSql = StrSql & " INNER JOIN his_estructura jornal ON empleado.ternro = jornal.ternro"
    StrSql = StrSql & " AND jornal.htetdesde <= " & ConvFecha(periodoAntHasta) & " AND (jornal.htethasta IS NULL OR jornal.htethasta >= " & ConvFecha(periodoAntHasta) & ")"
    StrSql = StrSql & " AND jornal.estrnro IN (" & listaJornal & ")"
    'Filtro Empleados
    StrSql = StrSql & " WHERE " & FiltroSql
    OpenRecordset StrSql, rs_datos
    
    If Not rs_datos.EOF Then
        valor1 = IIf(EsNulo(rs_datos!cant), 0, rs_datos!cant)
    End If
    rs_datos.Close
    

    'Mensualizados
    '------------------------------------------------------------------------------------
    Flog.writeline Espacios(Tabulador * 3) & "Mensualizados."
    StrSql = "SELECT count (distinct empleado.ternro) cant "
    StrSql = StrSql & " FROM empleado"
    'Filtro Empresa
    StrSql = StrSql & " INNER JOIN his_estructura empresa ON empleado.ternro = empresa.ternro"
    StrSql = StrSql & " AND empresa.estrnro = " & Empresa & " AND empresa.tenro=10"
    StrSql = StrSql & " AND empresa.htetdesde <= " & ConvFecha(periodoAntHasta) & " AND (empresa.htethasta IS NULL OR empresa.htethasta >= " & ConvFecha(periodoAntHasta) & ")"
    'Filtros de niveles de estructura
    If te1 <> 0 Then
        StrSql = StrSql & " INNER JOIN his_estructura tenro1 ON empleado.ternro = tenro1.ternro"
        StrSql = StrSql & " AND tenro1.tenro = " & te1
        StrSql = StrSql & " AND tenro1.htetdesde <= " & ConvFecha(periodoAntHasta) & " AND (tenro1.htethasta IS NULL OR tenro1.htethasta >= " & ConvFecha(periodoAntHasta) & ")"
        StrSql = StrSql & " AND tenro1.estrnro = " & estr1
    End If
    If te2 <> 0 Then
        StrSql = StrSql & " INNER JOIN his_estructura tenro2 ON empleado.ternro = tenro2.ternro"
        StrSql = StrSql & " AND tenro2.tenro = " & te2
        StrSql = StrSql & " AND tenro2.htetdesde <= " & ConvFecha(periodoAntHasta) & " AND (tenro2.htethasta IS NULL OR tenro2.htethasta >= " & ConvFecha(periodoAntHasta) & ")"
        StrSql = StrSql & " AND tenro2.estrnro = " & estr2
    End If
    If te3 <> 0 Then
        StrSql = StrSql & " INNER JOIN his_estructura tenro3 ON empleado.ternro = tenro3.ternro"
        StrSql = StrSql & " AND tenro3.tenro = " & te3
        StrSql = StrSql & " AND tenro3.htetdesde <= " & ConvFecha(periodoAntHasta) & " AND (tenro3.htethasta IS NULL OR tenro3.htethasta >= " & ConvFecha(periodoAntHasta) & ")"
        StrSql = StrSql & " AND tenro3.estrnro = " & estr3
    End If
    'Filtro Jornalizados
    StrSql = StrSql & " INNER JOIN his_estructura mensual ON empleado.ternro = mensual.ternro"
    StrSql = StrSql & " AND mensual.htetdesde <= " & ConvFecha(periodoAntHasta) & " AND (mensual.htethasta IS NULL OR mensual.htethasta >= " & ConvFecha(periodoAntHasta) & ")"
    StrSql = StrSql & " AND mensual.estrnro IN (" & listaMensual & ")"
    'Filtro Empleados
    StrSql = StrSql & " WHERE " & FiltroSql
    OpenRecordset StrSql, rs_datos
    
    If Not rs_datos.EOF Then
        valor2 = IIf(EsNulo(rs_datos!cant), 0, rs_datos!cant)
    End If
    rs_datos.Close
    

    'Resto
    '------------------------------------------------------------------------------------
    Flog.writeline Espacios(Tabulador * 3) & "Resto."
    StrSql = "SELECT count (distinct empleado.ternro) cant "
    StrSql = StrSql & " FROM empleado"
    'Filtro Empresa
    StrSql = StrSql & " INNER JOIN his_estructura empresa ON empleado.ternro = empresa.ternro"
    StrSql = StrSql & " AND empresa.estrnro = " & Empresa & " AND empresa.tenro=10"
    StrSql = StrSql & " AND empresa.htetdesde <= " & ConvFecha(periodoAntHasta) & " AND (empresa.htethasta IS NULL OR empresa.htethasta >= " & ConvFecha(periodoAntHasta) & ")"
    'Filtros de niveles de estructura
    If te1 <> 0 Then
        StrSql = StrSql & " INNER JOIN his_estructura tenro1 ON empleado.ternro = tenro1.ternro"
        StrSql = StrSql & " AND tenro1.tenro = " & te1
        StrSql = StrSql & " AND tenro1.htetdesde <= " & ConvFecha(periodoAntHasta) & " AND (tenro1.htethasta IS NULL OR tenro1.htethasta >= " & ConvFecha(periodoAntHasta) & ")"
        StrSql = StrSql & " AND tenro1.estrnro = " & estr1
    End If
    If te2 <> 0 Then
        StrSql = StrSql & " INNER JOIN his_estructura tenro2 ON empleado.ternro = tenro2.ternro"
        StrSql = StrSql & " AND tenro2.tenro = " & te2
        StrSql = StrSql & " AND tenro2.htetdesde <= " & ConvFecha(periodoAntHasta) & " AND (tenro2.htethasta IS NULL OR tenro2.htethasta >= " & ConvFecha(periodoAntHasta) & ")"
        StrSql = StrSql & " AND tenro2.estrnro = " & estr2
    End If
    If te3 <> 0 Then
        StrSql = StrSql & " INNER JOIN his_estructura tenro3 ON empleado.ternro = tenro3.ternro"
        StrSql = StrSql & " AND tenro3.tenro = " & te3
        StrSql = StrSql & " AND tenro3.htetdesde <= " & ConvFecha(periodoAntHasta) & " AND (tenro3.htethasta IS NULL OR tenro3.htethasta >= " & ConvFecha(periodoAntHasta) & ")"
        StrSql = StrSql & " AND tenro3.estrnro = " & estr3
    End If
    'Filtro Jornalizados
    StrSql = StrSql & " INNER JOIN his_estructura resto ON empleado.ternro = resto.ternro"
    StrSql = StrSql & " AND resto.htetdesde <= " & ConvFecha(periodoAntHasta) & " AND (resto.htethasta IS NULL OR resto.htethasta >= " & ConvFecha(periodoAntHasta) & ")"
    StrSql = StrSql & " AND resto.estrnro IN (" & listaResto & ")"
    'Filtro Empleados
    StrSql = StrSql & " WHERE " & FiltroSql
    OpenRecordset StrSql, rs_datos
    
    If Not rs_datos.EOF Then
        valor3 = IIf(EsNulo(rs_datos!cant), 0, rs_datos!cant)
    End If
    rs_datos.Close
    
End If

'Guardo los valores de la linea
'------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 2) & "Guardando valores Personal mes anterior."
Call GuardarFila(ultimaCabecera, 1, valor1, valor2, valor3, "Personal mes anterior")



'------------------------------------------------------------------------------------
'Altas
'------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 2) & "Procesando Altas."
valor1 = 0
valor2 = 0
valor3 = 0

'Jornalizados
'------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 3) & "Jornalizados."
StrSql = "SELECT count (distinct fases.fasnro) cant "
StrSql = StrSql & " FROM empleado"
'Que tenga fases que se abren en el periodo
StrSql = StrSql & " INNER JOIN fases ON fases.empleado = empleado.ternro"
StrSql = StrSql & " AND fases.altfec <= " & ConvFecha(PeriodoHasta)
StrSql = StrSql & " AND " & ConvFecha(PeriodoDesde) & " <= fases.altfec"
'Filtro Empresa
StrSql = StrSql & " INNER JOIN his_estructura empresa ON empleado.ternro = empresa.ternro"
StrSql = StrSql & " AND empresa.estrnro = " & Empresa & " AND empresa.tenro=10"
StrSql = StrSql & " AND empresa.htetdesde <= " & ConvFecha(PeriodoHasta) & " AND (empresa.htethasta IS NULL OR empresa.htethasta >= " & ConvFecha(PeriodoHasta) & ")"
'Filtros de niveles de estructura
If te1 <> 0 Then
    StrSql = StrSql & " INNER JOIN his_estructura tenro1 ON empleado.ternro = tenro1.ternro"
    StrSql = StrSql & " AND tenro1.tenro = " & te1
    StrSql = StrSql & " AND tenro1.htetdesde <= " & ConvFecha(PeriodoHasta) & " AND (tenro1.htethasta IS NULL OR tenro1.htethasta >= " & ConvFecha(PeriodoHasta) & ")"
    StrSql = StrSql & " AND tenro1.estrnro = " & estr1
End If
If te2 <> 0 Then
    StrSql = StrSql & " INNER JOIN his_estructura tenro2 ON empleado.ternro = tenro2.ternro"
    StrSql = StrSql & " AND tenro2.tenro = " & te2
    StrSql = StrSql & " AND tenro2.htetdesde <= " & ConvFecha(PeriodoHasta) & " AND (tenro2.htethasta IS NULL OR tenro2.htethasta >= " & ConvFecha(PeriodoHasta) & ")"
    StrSql = StrSql & " AND tenro2.estrnro = " & estr2
End If
If te3 <> 0 Then
    StrSql = StrSql & " INNER JOIN his_estructura tenro3 ON empleado.ternro = tenro3.ternro"
    StrSql = StrSql & " AND tenro3.tenro = " & te3
    StrSql = StrSql & " AND tenro3.htetdesde <= " & ConvFecha(PeriodoHasta) & " AND (tenro3.htethasta IS NULL OR tenro3.htethasta >= " & ConvFecha(PeriodoHasta) & ")"
    StrSql = StrSql & " AND tenro3.estrnro = " & estr3
End If
'Filtro Jornalizados
StrSql = StrSql & " INNER JOIN his_estructura jornal ON empleado.ternro = jornal.ternro"
StrSql = StrSql & " AND jornal.htetdesde <= " & ConvFecha(PeriodoHasta) & " AND (jornal.htethasta IS NULL OR jornal.htethasta >= " & ConvFecha(PeriodoHasta) & ")"
StrSql = StrSql & " AND jornal.estrnro IN (" & listaJornal & ")"
'Filtro Empleados
StrSql = StrSql & " WHERE " & FiltroSql
OpenRecordset StrSql, rs_datos

If Not rs_datos.EOF Then
    valor1 = IIf(EsNulo(rs_datos!cant), 0, rs_datos!cant)
End If
rs_datos.Close


'Mensualizados
'------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 3) & "Mensualizados."
StrSql = "SELECT count (distinct fases.fasnro) cant "
StrSql = StrSql & " FROM empleado"
'Que tenga fases que se abren en el periodo
StrSql = StrSql & " INNER JOIN fases ON fases.empleado = empleado.ternro"
StrSql = StrSql & " AND fases.altfec <= " & ConvFecha(PeriodoHasta)
StrSql = StrSql & " AND " & ConvFecha(PeriodoDesde) & " <= fases.altfec"
'Filtro Empresa
StrSql = StrSql & " INNER JOIN his_estructura empresa ON empleado.ternro = empresa.ternro"
StrSql = StrSql & " AND empresa.estrnro = " & Empresa & " AND empresa.tenro=10"
StrSql = StrSql & " AND empresa.htetdesde <= " & ConvFecha(PeriodoHasta) & " AND (empresa.htethasta IS NULL OR empresa.htethasta >= " & ConvFecha(PeriodoHasta) & ")"
'Filtros de niveles de estructura
If te1 <> 0 Then
    StrSql = StrSql & " INNER JOIN his_estructura tenro1 ON empleado.ternro = tenro1.ternro"
    StrSql = StrSql & " AND tenro1.tenro = " & te1
    StrSql = StrSql & " AND tenro1.htetdesde <= " & ConvFecha(PeriodoHasta) & " AND (tenro1.htethasta IS NULL OR tenro1.htethasta >= " & ConvFecha(PeriodoHasta) & ")"
    StrSql = StrSql & " AND tenro1.estrnro = " & estr1
End If
If te2 <> 0 Then
    StrSql = StrSql & " INNER JOIN his_estructura tenro2 ON empleado.ternro = tenro2.ternro"
    StrSql = StrSql & " AND tenro2.tenro = " & te2
    StrSql = StrSql & " AND tenro2.htetdesde <= " & ConvFecha(PeriodoHasta) & " AND (tenro2.htethasta IS NULL OR tenro2.htethasta >= " & ConvFecha(PeriodoHasta) & ")"
    StrSql = StrSql & " AND tenro2.estrnro = " & estr2
End If
If te3 <> 0 Then
    StrSql = StrSql & " INNER JOIN his_estructura tenro3 ON empleado.ternro = tenro3.ternro"
    StrSql = StrSql & " AND tenro3.tenro = " & te3
    StrSql = StrSql & " AND tenro3.htetdesde <= " & ConvFecha(PeriodoHasta) & " AND (tenro3.htethasta IS NULL OR tenro3.htethasta >= " & ConvFecha(PeriodoHasta) & ")"
    StrSql = StrSql & " AND tenro3.estrnro = " & estr3
End If
'Filtro Jornalizados
StrSql = StrSql & " INNER JOIN his_estructura mensual ON empleado.ternro = mensual.ternro"
StrSql = StrSql & " AND mensual.htetdesde <= " & ConvFecha(PeriodoHasta) & " AND (mensual.htethasta IS NULL OR mensual.htethasta >= " & ConvFecha(PeriodoHasta) & ")"
StrSql = StrSql & " AND mensual.estrnro IN (" & listaMensual & ")"
'Filtro Empleados
StrSql = StrSql & " WHERE " & FiltroSql
OpenRecordset StrSql, rs_datos

If Not rs_datos.EOF Then
    valor2 = IIf(EsNulo(rs_datos!cant), 0, rs_datos!cant)
End If
rs_datos.Close


'Resto
'------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 3) & "Resto."
StrSql = "SELECT count (distinct fases.fasnro) cant "
StrSql = StrSql & " FROM empleado"
'Que tenga fases que se abren en el periodo
StrSql = StrSql & " INNER JOIN fases ON fases.empleado = empleado.ternro"
StrSql = StrSql & " AND fases.altfec <= " & ConvFecha(PeriodoHasta)
StrSql = StrSql & " AND " & ConvFecha(PeriodoDesde) & " <= fases.altfec"
'Filtro Empresa
StrSql = StrSql & " INNER JOIN his_estructura empresa ON empleado.ternro = empresa.ternro"
StrSql = StrSql & " AND empresa.estrnro = " & Empresa & " AND empresa.tenro=10"
StrSql = StrSql & " AND empresa.htetdesde <= " & ConvFecha(PeriodoHasta) & " AND (empresa.htethasta IS NULL OR empresa.htethasta >= " & ConvFecha(PeriodoHasta) & ")"
'Filtros de niveles de estructura
If te1 <> 0 Then
    StrSql = StrSql & " INNER JOIN his_estructura tenro1 ON empleado.ternro = tenro1.ternro"
    StrSql = StrSql & " AND tenro1.tenro = " & te1
    StrSql = StrSql & " AND tenro1.htetdesde <= " & ConvFecha(PeriodoHasta) & " AND (tenro1.htethasta IS NULL OR tenro1.htethasta >= " & ConvFecha(PeriodoHasta) & ")"
    StrSql = StrSql & " AND tenro1.estrnro = " & estr1
End If
If te2 <> 0 Then
    StrSql = StrSql & " INNER JOIN his_estructura tenro2 ON empleado.ternro = tenro2.ternro"
    StrSql = StrSql & " AND tenro2.tenro = " & te2
    StrSql = StrSql & " AND tenro2.htetdesde <= " & ConvFecha(PeriodoHasta) & " AND (tenro2.htethasta IS NULL OR tenro2.htethasta >= " & ConvFecha(PeriodoHasta) & ")"
    StrSql = StrSql & " AND tenro2.estrnro = " & estr2
End If
If te3 <> 0 Then
    StrSql = StrSql & " INNER JOIN his_estructura tenro3 ON empleado.ternro = tenro3.ternro"
    StrSql = StrSql & " AND tenro3.tenro = " & te3
    StrSql = StrSql & " AND tenro3.htetdesde <= " & ConvFecha(PeriodoHasta) & " AND (tenro3.htethasta IS NULL OR tenro3.htethasta >= " & ConvFecha(PeriodoHasta) & ")"
    StrSql = StrSql & " AND tenro3.estrnro = " & estr3
End If
'Filtro Jornalizados
StrSql = StrSql & " INNER JOIN his_estructura resto ON empleado.ternro = resto.ternro"
StrSql = StrSql & " AND resto.htetdesde <= " & ConvFecha(PeriodoHasta) & " AND (resto.htethasta IS NULL OR resto.htethasta >= " & ConvFecha(PeriodoHasta) & ")"
StrSql = StrSql & " AND resto.estrnro IN (" & listaResto & ")"
'Filtro Empleados
StrSql = StrSql & " WHERE " & FiltroSql
OpenRecordset StrSql, rs_datos

If Not rs_datos.EOF Then
    valor3 = IIf(EsNulo(rs_datos!cant), 0, rs_datos!cant)
End If
rs_datos.Close

'Guardo los valores de la linea
'------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 2) & "Guardando valores Altas."
Call GuardarFila(ultimaCabecera, 1, valor1, valor2, valor3, "Altas")



'------------------------------------------------------------------------------------
'Bajas
'------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 2) & "Procesando Bajas."
valor1 = 0
valor2 = 0
valor3 = 0

'Jornalizados
'------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 3) & "Jornalizados."
StrSql = "SELECT count (distinct fases.fasnro) cant "
StrSql = StrSql & " FROM empleado"
'Que tenga fases que se abren en el periodo
StrSql = StrSql & " INNER JOIN fases ON fases.empleado = empleado.ternro"
StrSql = StrSql & " AND fases.bajfec <= " & ConvFecha(PeriodoHasta)
StrSql = StrSql & " AND " & ConvFecha(PeriodoDesde) & " <= fases.bajfec"
'Filtro Empresa
StrSql = StrSql & " INNER JOIN his_estructura empresa ON empleado.ternro = empresa.ternro"
StrSql = StrSql & " AND empresa.estrnro = " & Empresa & " AND empresa.tenro=10"
StrSql = StrSql & " AND empresa.htetdesde <= " & ConvFecha(PeriodoHasta) & " AND (empresa.htethasta IS NULL OR empresa.htethasta >= " & ConvFecha(PeriodoHasta) & ")"
'Filtros de niveles de estructura
If te1 <> 0 Then
    StrSql = StrSql & " INNER JOIN his_estructura tenro1 ON empleado.ternro = tenro1.ternro"
    StrSql = StrSql & " AND tenro1.tenro = " & te1
    StrSql = StrSql & " AND tenro1.htetdesde <= " & ConvFecha(PeriodoHasta) & " AND (tenro1.htethasta IS NULL OR tenro1.htethasta >= " & ConvFecha(PeriodoHasta) & ")"
    StrSql = StrSql & " AND tenro1.estrnro = " & estr1
End If
If te2 <> 0 Then
    StrSql = StrSql & " INNER JOIN his_estructura tenro2 ON empleado.ternro = tenro2.ternro"
    StrSql = StrSql & " AND tenro2.tenro = " & te2
    StrSql = StrSql & " AND tenro2.htetdesde <= " & ConvFecha(PeriodoHasta) & " AND (tenro2.htethasta IS NULL OR tenro2.htethasta >= " & ConvFecha(PeriodoHasta) & ")"
    StrSql = StrSql & " AND tenro2.estrnro = " & estr2
End If
If te3 <> 0 Then
    StrSql = StrSql & " INNER JOIN his_estructura tenro3 ON empleado.ternro = tenro3.ternro"
    StrSql = StrSql & " AND tenro3.tenro = " & te3
    StrSql = StrSql & " AND tenro3.htetdesde <= " & ConvFecha(PeriodoHasta) & " AND (tenro3.htethasta IS NULL OR tenro3.htethasta >= " & ConvFecha(PeriodoHasta) & ")"
    StrSql = StrSql & " AND tenro3.estrnro = " & estr3
End If
'Filtro Jornalizados
StrSql = StrSql & " INNER JOIN his_estructura jornal ON empleado.ternro = jornal.ternro"
StrSql = StrSql & " AND jornal.htetdesde <= " & ConvFecha(PeriodoHasta) & " AND (jornal.htethasta IS NULL OR jornal.htethasta >= " & ConvFecha(PeriodoHasta) & ")"
StrSql = StrSql & " AND jornal.estrnro IN (" & listaJornal & ")"
'Filtro Empleados
StrSql = StrSql & " WHERE " & FiltroSql
OpenRecordset StrSql, rs_datos

If Not rs_datos.EOF Then
    valor1 = IIf(EsNulo(rs_datos!cant), 0, rs_datos!cant)
End If
rs_datos.Close


'Mensualizados
'------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 3) & "Mensualizados."
StrSql = "SELECT count (distinct fases.fasnro) cant "
StrSql = StrSql & " FROM empleado"
'Que tenga fases que se abren en el periodo
StrSql = StrSql & " INNER JOIN fases ON fases.empleado = empleado.ternro"
StrSql = StrSql & " AND fases.bajfec <= " & ConvFecha(PeriodoHasta)
StrSql = StrSql & " AND " & ConvFecha(PeriodoDesde) & " <= fases.bajfec"
'Filtro Empresa
StrSql = StrSql & " INNER JOIN his_estructura empresa ON empleado.ternro = empresa.ternro"
StrSql = StrSql & " AND empresa.estrnro = " & Empresa & " AND empresa.tenro=10"
StrSql = StrSql & " AND empresa.htetdesde <= " & ConvFecha(PeriodoHasta) & " AND (empresa.htethasta IS NULL OR empresa.htethasta >= " & ConvFecha(PeriodoHasta) & ")"
'Filtros de niveles de estructura
If te1 <> 0 Then
    StrSql = StrSql & " INNER JOIN his_estructura tenro1 ON empleado.ternro = tenro1.ternro"
    StrSql = StrSql & " AND tenro1.tenro = " & te1
    StrSql = StrSql & " AND tenro1.htetdesde <= " & ConvFecha(PeriodoHasta) & " AND (tenro1.htethasta IS NULL OR tenro1.htethasta >= " & ConvFecha(PeriodoHasta) & ")"
    StrSql = StrSql & " AND tenro1.estrnro = " & estr1
End If
If te2 <> 0 Then
    StrSql = StrSql & " INNER JOIN his_estructura tenro2 ON empleado.ternro = tenro2.ternro"
    StrSql = StrSql & " AND tenro2.tenro = " & te2
    StrSql = StrSql & " AND tenro2.htetdesde <= " & ConvFecha(PeriodoHasta) & " AND (tenro2.htethasta IS NULL OR tenro2.htethasta >= " & ConvFecha(PeriodoHasta) & ")"
    StrSql = StrSql & " AND tenro2.estrnro = " & estr2
End If
If te3 <> 0 Then
    StrSql = StrSql & " INNER JOIN his_estructura tenro3 ON empleado.ternro = tenro3.ternro"
    StrSql = StrSql & " AND tenro3.tenro = " & te3
    StrSql = StrSql & " AND tenro3.htetdesde <= " & ConvFecha(PeriodoHasta) & " AND (tenro3.htethasta IS NULL OR tenro3.htethasta >= " & ConvFecha(PeriodoHasta) & ")"
    StrSql = StrSql & " AND tenro3.estrnro = " & estr3
End If
'Filtro Jornalizados
StrSql = StrSql & " INNER JOIN his_estructura mensual ON empleado.ternro = mensual.ternro"
StrSql = StrSql & " AND mensual.htetdesde <= " & ConvFecha(PeriodoHasta) & " AND (mensual.htethasta IS NULL OR mensual.htethasta >= " & ConvFecha(PeriodoHasta) & ")"
StrSql = StrSql & " AND mensual.estrnro IN (" & listaMensual & ")"
'Filtro Empleados
StrSql = StrSql & " WHERE " & FiltroSql
OpenRecordset StrSql, rs_datos

If Not rs_datos.EOF Then
    valor2 = IIf(EsNulo(rs_datos!cant), 0, rs_datos!cant)
End If
rs_datos.Close


'Resto
'------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 3) & "Resto."
StrSql = "SELECT count (distinct fases.fasnro) cant "
StrSql = StrSql & " FROM empleado"
'Que tenga fases que se abren en el periodo
StrSql = StrSql & " INNER JOIN fases ON fases.empleado = empleado.ternro"
StrSql = StrSql & " AND fases.bajfec <= " & ConvFecha(PeriodoHasta)
StrSql = StrSql & " AND " & ConvFecha(PeriodoDesde) & " <= fases.bajfec"
'Filtro Empresa
StrSql = StrSql & " INNER JOIN his_estructura empresa ON empleado.ternro = empresa.ternro"
StrSql = StrSql & " AND empresa.estrnro = " & Empresa & " AND empresa.tenro=10"
StrSql = StrSql & " AND empresa.htetdesde <= " & ConvFecha(PeriodoHasta) & " AND (empresa.htethasta IS NULL OR empresa.htethasta >= " & ConvFecha(PeriodoHasta) & ")"
'Filtros de niveles de estructura
If te1 <> 0 Then
    StrSql = StrSql & " INNER JOIN his_estructura tenro1 ON empleado.ternro = tenro1.ternro"
    StrSql = StrSql & " AND tenro1.tenro = " & te1
    StrSql = StrSql & " AND tenro1.htetdesde <= " & ConvFecha(PeriodoHasta) & " AND (tenro1.htethasta IS NULL OR tenro1.htethasta >= " & ConvFecha(PeriodoHasta) & ")"
    StrSql = StrSql & " AND tenro1.estrnro = " & estr1
End If
If te2 <> 0 Then
    StrSql = StrSql & " INNER JOIN his_estructura tenro2 ON empleado.ternro = tenro2.ternro"
    StrSql = StrSql & " AND tenro2.tenro = " & te2
    StrSql = StrSql & " AND tenro2.htetdesde <= " & ConvFecha(PeriodoHasta) & " AND (tenro2.htethasta IS NULL OR tenro2.htethasta >= " & ConvFecha(PeriodoHasta) & ")"
    StrSql = StrSql & " AND tenro2.estrnro = " & estr2
End If
If te3 <> 0 Then
    StrSql = StrSql & " INNER JOIN his_estructura tenro3 ON empleado.ternro = tenro3.ternro"
    StrSql = StrSql & " AND tenro3.tenro = " & te3
    StrSql = StrSql & " AND tenro3.htetdesde <= " & ConvFecha(PeriodoHasta) & " AND (tenro3.htethasta IS NULL OR tenro3.htethasta >= " & ConvFecha(PeriodoHasta) & ")"
    StrSql = StrSql & " AND tenro3.estrnro = " & estr3
End If
'Filtro Jornalizados
StrSql = StrSql & " INNER JOIN his_estructura resto ON empleado.ternro = resto.ternro"
StrSql = StrSql & " AND resto.htetdesde <= " & ConvFecha(PeriodoHasta) & " AND (resto.htethasta IS NULL OR resto.htethasta >= " & ConvFecha(PeriodoHasta) & ")"
StrSql = StrSql & " AND resto.estrnro IN (" & listaResto & ")"
'Filtro Empleados
StrSql = StrSql & " WHERE " & FiltroSql
OpenRecordset StrSql, rs_datos

If Not rs_datos.EOF Then
    valor3 = IIf(EsNulo(rs_datos!cant), 0, rs_datos!cant)
End If
rs_datos.Close

'Guardo los valores de la linea
'------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 2) & "Guardando valores Bajas."
Call GuardarFila(ultimaCabecera, 1, valor1, valor2, valor3, "Bajas")


'Dim ArrConfRep(16, 11) As TConfRep
'Dim ArrEtiq(16) As String
'Dim indArrFila As Long
'Dim indArrcol As Long

Dim listaConc As String
Dim listaAcum As String

For indArrFila = 1 To 8
    
    
    
    If (indArrFila = 1) Or (indArrFila = 2) Then
        nroTipo = 2
        Flog.writeline Espacios(Tabulador * 2) & "Procesando Horas Trabajadas Fila " & indArrFila & "."
    Else
        nroTipo = 3
        Flog.writeline Espacios(Tabulador * 2) & "Procesando Salarios Devengados Fila " & indArrFila & "."
    End If
    
    If (ArrConfRep(indArrFila, 0).tipo <> "0") And (ArrConfRep(indArrFila, 0).tipo <> "") Then
        
        listaConc = "'0'"
        listaAcum = "0"
        'Recorro todos los componentes de los conceptos y acumuladores configurados por confrep
        For indArrcol = 1 To ArrConfRep(indArrFila, 0).tipo
            If ArrConfRep(indArrFila, indArrcol).tipo = "AC" Then
                listaAcum = listaAcum & "," & IIf(EsNulo(ArrConfRep(indArrFila, indArrcol).Cod), 0, ArrConfRep(indArrFila, indArrcol).Cod)
            Else
                listaConc = listaConc & ",'" & IIf(EsNulo(ArrConfRep(indArrFila, indArrcol).Cod), 0, ArrConfRep(indArrFila, indArrcol).Cod) & "'"
            End If
        Next indArrcol

        valor1 = 0
        valor2 = 0
        valor3 = 0
        
        'Busco los acumuladores configurados
        If listaAcum <> "0" Then
            Call BuscarAcum(te1, estr1, te2, estr2, te3, estr3, 1, listaAcum, Periodo, valor1, nroTipo)
            Call BuscarAcum(te1, estr1, te2, estr2, te3, estr3, 2, listaAcum, Periodo, valor2, nroTipo)
            Call BuscarAcum(te1, estr1, te2, estr2, te3, estr3, 3, listaAcum, Periodo, valor3, nroTipo)
        End If
        
        'Busco los conceptos configurados
        If listaConc <> "'0'" Then
            Call BuscarConc(te1, estr1, te2, estr2, te3, estr3, 1, listaConc, Periodo, valor1, nroTipo)
            Call BuscarConc(te1, estr1, te2, estr2, te3, estr3, 2, listaConc, Periodo, valor2, nroTipo)
            Call BuscarConc(te1, estr1, te2, estr2, te3, estr3, 3, listaConc, Periodo, valor3, nroTipo)
        End If
        Flog.writeline Espacios(Tabulador * 2) & "Guardando valores."
        Call GuardarFila(ultimaCabecera, nroTipo, valor1, valor2, valor3, ArrEtiq(indArrFila))
    Else
        'No se configuro la columna
        Flog.writeline Espacios(Tabulador * 3) & "No se encontraron valores configurados en el confrep para la fila."
        Flog.writeline Espacios(Tabulador * 2) & "Guardando valores nulos."
        Call GuardarFila(ultimaCabecera, nroTipo, 0, 0, 0, "")
    End If
    
Next indArrFila

'Cierro recordsets
If rs_datos.State = adStateOpen Then rs_datos.Close
Set rs_datos = Nothing

Exit Sub

E_GenerarHoja:
    Flog.writeline "=================================================================="
    Flog.writeline "Error en GenerarHoja"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="
    

End Sub



Public Sub GuardarFila(ByVal Cabecera As Long, ByVal tipo As Long, ByVal Cant1 As Double, ByVal Cant2 As Double, ByVal Cant3 As Double, ByVal Texto As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Guarda los 3 valores de una fila
' Autor      : Martin Ferraro
' Fecha      : 02/04/2008
' --------------------------------------------------------------------------------------------

On Error GoTo E_GuardarFila:

StrSql = "INSERT INTO rep_indec_det"
StrSql = StrSql & " (indecnro"
StrSql = StrSql & " ,tipo"
StrSql = StrSql & " ,valor1"
StrSql = StrSql & " ,valor2"
StrSql = StrSql & " ,valor3"
StrSql = StrSql & " ,etiqueta)"
StrSql = StrSql & " Values"
StrSql = StrSql & " (" & Cabecera
StrSql = StrSql & " ," & tipo
StrSql = StrSql & " ," & Cant1
StrSql = StrSql & " ," & Cant2
StrSql = StrSql & " ," & Cant3
StrSql = StrSql & " ,'" & Mid(Texto, 1, 100) & "')"
objConn.Execute StrSql, , adExecuteNoRecords

Exit Sub

E_GuardarFila:
    Flog.writeline "=================================================================="
    Flog.writeline "Error en GuardarFila"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="

End Sub



Public Sub BuscarAcum(ByVal te1 As Long, ByVal estr1 As Long, ByVal te2 As Long, ByVal estr2 As Long, ByVal te3 As Long, ByVal estr3 As Long, ByVal tipo As Integer, ByVal Acumuladores As String, ByVal nroPeriodo As Long, ByRef valorSalida As Double, ByVal buscar As Integer)
' --------------------------------------------------------------------------------------------
' Descripcion: Busca el monto de los acumuladores de listaAcum liquidados en el periodo para jornales, mens. o resto segun tipo
' Autor      : Martin Ferraro
' Fecha      : 02/04/2008
' --------------------------------------------------------------------------------------------
Dim ValorAux As Double
Dim rs_Cabliq As New ADODB.Recordset

On Error GoTo E_BuscarAcum

ValorAux = 0

'Busco todos los cabliq a procesar
Select Case buscar:
    Case 2
        StrSql = " SELECT distinct cabliq.cliqnro, acu_liq.acunro, acu_liq.alcant valor"
    Case 3
        StrSql = " SELECT distinct cabliq.cliqnro, acu_liq.acunro, acu_liq.almonto valor"
    Case Else
        'No contemplado
        Exit Sub
End Select
StrSql = StrSql & " FROM cabliq "
StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro AND proceso.pliqnro = " & nroPeriodo
StrSql = StrSql & " INNER JOIN acu_liq ON cabliq.cliqnro = acu_liq.cliqnro AND acu_liq.acunro IN (" & Acumuladores & ")"
'Filtro Jornal (tipo = 1) , Mensual (tipo = 2), Resto (tipo = 3)
StrSql = StrSql & " INNER JOIN his_estructura ON cabliq.empleado = his_estructura.ternro"
StrSql = StrSql & " AND his_estructura.htetdesde <= " & ConvFecha(PeriodoHasta) & " AND (his_estructura.htethasta IS NULL OR his_estructura.htethasta >= " & ConvFecha(PeriodoHasta) & ")"
Select Case tipo
    Case 1
        StrSql = StrSql & " AND his_estructura.estrnro IN (" & listaJornal & ")"
    Case 2
        StrSql = StrSql & " AND his_estructura.estrnro IN (" & listaMensual & ")"
    Case 3
        StrSql = StrSql & " AND his_estructura.estrnro IN (" & listaResto & ")"
End Select
'Filtro Empresa
StrSql = StrSql & " INNER JOIN his_estructura empresa ON cabliq.empleado = empresa.ternro"
StrSql = StrSql & " AND empresa.estrnro = " & Empresa & " AND empresa.tenro = 10"
StrSql = StrSql & " AND empresa.htetdesde <= " & ConvFecha(PeriodoHasta) & " AND (empresa.htethasta IS NULL OR empresa.htethasta >= " & ConvFecha(PeriodoHasta) & ")"
'Filtros de niveles de estructura
If te1 <> 0 Then
    StrSql = StrSql & " INNER JOIN his_estructura tenro1 ON cabliq.empleado = tenro1.ternro"
    StrSql = StrSql & " AND tenro1.tenro = " & te1
    StrSql = StrSql & " AND tenro1.htetdesde <= " & ConvFecha(PeriodoHasta) & " AND (tenro1.htethasta IS NULL OR tenro1.htethasta >= " & ConvFecha(PeriodoHasta) & ")"
    StrSql = StrSql & " AND tenro1.estrnro = " & estr1
End If
If te2 <> 0 Then
    StrSql = StrSql & " INNER JOIN his_estructura tenro2 ON cabliq.empleado = tenro2.ternro"
    StrSql = StrSql & " AND tenro2.tenro = " & te2
    StrSql = StrSql & " AND tenro2.htetdesde <= " & ConvFecha(PeriodoHasta) & " AND (tenro2.htethasta IS NULL OR tenro2.htethasta >= " & ConvFecha(PeriodoHasta) & ")"
    StrSql = StrSql & " AND tenro2.estrnro = " & estr2
End If
If te3 <> 0 Then
    StrSql = StrSql & " INNER JOIN his_estructura tenro3 ON cabliq.empleado = tenro3.ternro"
    StrSql = StrSql & " AND tenro3.tenro = " & te3
    StrSql = StrSql & " AND tenro3.htetdesde <= " & ConvFecha(PeriodoHasta) & " AND (tenro3.htethasta IS NULL OR tenro3.htethasta >= " & ConvFecha(PeriodoHasta) & ")"
    StrSql = StrSql & " AND tenro3.estrnro = " & estr3
End If
OpenRecordset StrSql, rs_Cabliq

Do While Not rs_Cabliq.EOF
    ValorAux = ValorAux + rs_Cabliq!Valor
    rs_Cabliq.MoveNext
Loop

rs_Cabliq.Close

'Escribo la variable de salida
valorSalida = valorSalida + ValorAux

'Cierro recordsets
If rs_Cabliq.State = adStateOpen Then rs_Cabliq.Close
Set rs_Cabliq = Nothing


Exit Sub

E_BuscarAcum:
    Flog.writeline "=================================================================="
    Flog.writeline "Error en BuscarAcum"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="

End Sub



Public Sub BuscarConc(ByVal te1 As Long, ByVal estr1 As Long, ByVal te2 As Long, ByVal estr2 As Long, ByVal te3 As Long, ByVal estr3 As Long, ByVal tipo As Integer, ByVal Conceptos As String, ByVal nroPeriodo As Long, ByRef valorSalida As Double, ByVal buscar As Integer)
' --------------------------------------------------------------------------------------------
' Descripcion: Busca el monto de los conceptos de listaAcum liquidados en el periodo para jornales, mens. o resto segun tipo
' Autor      : Martin Ferraro
' Fecha      : 02/04/2008
' --------------------------------------------------------------------------------------------
Dim ValorAux As Double
Dim rs_Cabliq As New ADODB.Recordset

On Error GoTo E_BuscarConc

ValorAux = 0

'Busco todos los cabliq a procesar
Select Case buscar:
    Case 2
        StrSql = " SELECT distinct cabliq.cliqnro, detliq.concnro, detliq.dlicant valor"
    Case 3
        StrSql = " SELECT distinct cabliq.cliqnro, detliq.concnro, detliq.dlimonto valor"
    Case Else
        'No contemplado
        Exit Sub
End Select
StrSql = StrSql & " FROM cabliq "
StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro AND proceso.pliqnro = " & nroPeriodo
StrSql = StrSql & " INNER JOIN detliq ON cabliq.cliqnro = detliq.cliqnro"
StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro AND concepto.conccod IN (" & Conceptos & ")"
'Filtro Jornal (tipo = 1) , Mensual (tipo = 2), Resto (tipo = 3)
StrSql = StrSql & " INNER JOIN his_estructura ON cabliq.empleado = his_estructura.ternro"
StrSql = StrSql & " AND his_estructura.htetdesde <= " & ConvFecha(PeriodoHasta) & " AND (his_estructura.htethasta IS NULL OR his_estructura.htethasta >= " & ConvFecha(PeriodoHasta) & ")"
Select Case tipo
    Case 1
        StrSql = StrSql & " AND his_estructura.estrnro IN (" & listaJornal & ")"
    Case 2
        StrSql = StrSql & " AND his_estructura.estrnro IN (" & listaMensual & ")"
    Case 3
        StrSql = StrSql & " AND his_estructura.estrnro IN (" & listaResto & ")"
End Select
'Filtro Empresa
StrSql = StrSql & " INNER JOIN his_estructura empresa ON cabliq.empleado = empresa.ternro"
StrSql = StrSql & " AND empresa.estrnro = " & Empresa & " AND empresa.tenro = 10"
StrSql = StrSql & " AND empresa.htetdesde <= " & ConvFecha(PeriodoHasta) & " AND (empresa.htethasta IS NULL OR empresa.htethasta >= " & ConvFecha(PeriodoHasta) & ")"
'Filtros de niveles de estructura
If te1 <> 0 Then
    StrSql = StrSql & " INNER JOIN his_estructura tenro1 ON cabliq.empleado = tenro1.ternro"
    StrSql = StrSql & " AND tenro1.tenro = " & te1
    StrSql = StrSql & " AND tenro1.htetdesde <= " & ConvFecha(PeriodoHasta) & " AND (tenro1.htethasta IS NULL OR tenro1.htethasta >= " & ConvFecha(PeriodoHasta) & ")"
    StrSql = StrSql & " AND tenro1.estrnro = " & estr1
End If
If te2 <> 0 Then
    StrSql = StrSql & " INNER JOIN his_estructura tenro2 ON cabliq.empleado = tenro2.ternro"
    StrSql = StrSql & " AND tenro2.tenro = " & te2
    StrSql = StrSql & " AND tenro2.htetdesde <= " & ConvFecha(PeriodoHasta) & " AND (tenro2.htethasta IS NULL OR tenro2.htethasta >= " & ConvFecha(PeriodoHasta) & ")"
    StrSql = StrSql & " AND tenro2.estrnro = " & estr2
End If
If te3 <> 0 Then
    StrSql = StrSql & " INNER JOIN his_estructura tenro3 ON cabliq.empleado = tenro3.ternro"
    StrSql = StrSql & " AND tenro3.tenro = " & te3
    StrSql = StrSql & " AND tenro3.htetdesde <= " & ConvFecha(PeriodoHasta) & " AND (tenro3.htethasta IS NULL OR tenro3.htethasta >= " & ConvFecha(PeriodoHasta) & ")"
    StrSql = StrSql & " AND tenro3.estrnro = " & estr3
End If
OpenRecordset StrSql, rs_Cabliq

Do While Not rs_Cabliq.EOF
    ValorAux = ValorAux + rs_Cabliq!Valor
    rs_Cabliq.MoveNext
Loop

rs_Cabliq.Close

'Escribo la variable de salida
valorSalida = valorSalida + ValorAux

'Cierro recordsets
If rs_Cabliq.State = adStateOpen Then rs_Cabliq.Close
Set rs_Cabliq = Nothing


Exit Sub

E_BuscarConc:
    Flog.writeline "=================================================================="
    Flog.writeline "Error en BuscarConc"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="

End Sub

