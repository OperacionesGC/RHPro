Attribute VB_Name = "MdlRepBandasVsSueldos"
Option Explicit
'----------------------------------------------------------------------------------------
'Const Version = "1.00"
'Const FechaVersion = "06/08/2009"


'Const Version = "1.01"
'Const FechaVersion = "07/10/2009"   'FGZ
'                                   Correccion de errrores varios

Const Version = "1.02"
Const FechaVersion = "05/12/2013"  'Carmen Quintero
'                                   CAS-19674 - H&A - Error al procesar reporte variación de bandas -
'                                   Se agregó condicion en la consulta principal


' ---------------------------------------------------------------------------------------------
' Descripcion: Reporte de bandas vs sueldos
' Autor      : Gustavo Ring
' Fecha      : 06/08/2009
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

' Variables

Global idusuario As String
Global Fecha As Date
Global hora As String




Public Sub Main()

' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial del Generador de Reporte.
' Autor      : Gustavo Ring
' Fecha      : 06/08/2009
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim objconnMain As New ADODB.Connection
Dim strCmdLine As String
Dim Nombre_Arch As String
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

    On Error Resume Next
    'Abro la conexion
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Then
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If
    OpenConnection strconexion, objconnProgreso
    If Err.Number <> 0 Then
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If
    On Error GoTo 0
    
    On Error GoTo ME_Main
    
    Nombre_Arch = PathFLog & "Rep_Banda_vs_sueldo" & "-" & NroProcesoBatch & ".log"
    
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
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0 ,bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Flog.writeline "Pone el estado en procesando"
    
    'Obtengo los datos del proceso
    'StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 251 AND bpronro =" & NroProcesoBatch
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        bprcparam = rs_batch_proceso!bprcparam
        idusuario = rs_batch_proceso!iduser
        Fecha = rs_batch_proceso!bprcfecha
        hora = rs_batch_proceso!bprchora
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        Call RepBandaVersusSueldo(NroProcesoBatch, bprcparam)
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


Public Sub RepBandaVersusSueldo(ByVal bpronro As Long, ByVal Parametros As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del reporte Banda Versus Sueldo
' Autor      : Gustavo Ring
' Fecha      : 06/08/2009
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
Dim Orden As String
Dim mes As Integer
Dim anio As Integer
Dim titulofiltro As String
Dim bsinterna As Integer
Dim obnro As Integer
Dim empzona As Integer
Dim filtro As String
'Dim fecestr As String
Dim fecestr As Date
Dim ultimodia As Integer
Dim lista As String
'RecordSet
Dim rs_Consult As New ADODB.Recordset
Dim rs_insert As New ADODB.Recordset
Dim siglazona As String
Dim col1 As Double
Dim col2 As Double
Dim col3 As Double
Dim puesto As String
Dim grado As String
Dim zona As String
Dim referencia As String
Dim mediana As Double
Dim dif As Double
Dim porcdif As Double
Dim filanro As Integer
Dim guardarUlt As Boolean
Dim StrSql As String
Dim zonacampo As String

'Inicio codigo ejecutable
On Error GoTo CE


'------------------------------------------------------------------------------------
' Levanto cada parametro por separado, el separador de parametros es "@"
'------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 0) & "levantando parametros" & Parametros
If Not IsNull(Parametros) Then
    
    arrParam = Split(Parametros, "@")
    
    If UBound(arrParam) = 13 Then
                
        filtro = CStr(arrParam(0))
        Tenro1 = CLng(arrParam(1))
        Estrnro1 = CLng(arrParam(2))
        Tenro2 = CLng(arrParam(3))
        Estrnro2 = CLng(arrParam(4))
        Tenro3 = CLng(arrParam(5))
        Estrnro3 = CLng(arrParam(6))
        
        
        Orden = arrParam(7)
        mes = CInt(arrParam(8))
        anio = CInt(arrParam(9))
        bsinterna = CInt(arrParam(10))
        obnro = CInt(arrParam(11))
        empzona = CInt(arrParam(12))
        titulofiltro = arrParam(13)
        
        Flog.writeline Espacios(Tabulador * 1) & "Filtro = " & filtro
        Flog.writeline Espacios(Tabulador * 1) & "TE1 = " & Tenro1
        Flog.writeline Espacios(Tabulador * 1) & "Estr1 = " & Estrnro1
        Flog.writeline Espacios(Tabulador * 1) & "TE2 = " & Tenro2
        Flog.writeline Espacios(Tabulador * 1) & "Estr2 = " & Estrnro2
        Flog.writeline Espacios(Tabulador * 1) & "TE3 = " & Tenro3
        Flog.writeline Espacios(Tabulador * 1) & "Estr3 = " & Estrnro3
        Flog.writeline Espacios(Tabulador * 1) & "Mes = " & mes
        Flog.writeline Espacios(Tabulador * 1) & "Año = " & anio
        Flog.writeline Espacios(Tabulador * 1) & "Orden = " & Orden
        Flog.writeline Espacios(Tabulador * 1) & "Banda interna = " & bsinterna
        Flog.writeline Espacios(Tabulador * 1) & "Cod.Origen = " & obnro
        Flog.writeline Espacios(Tabulador * 1) & "Cod.Zona.Banda = " & empzona
        Flog.writeline Espacios(Tabulador * 1) & "Título Filtro = " & titulofiltro
    Else
        Flog.writeline Espacios(Tabulador * 0) & "ERROR. La cantidad de parámetros del filtro no es la esperada."
        HuboError = True
        Exit Sub
        
    End If
Else
    Flog.writeline Espacios(Tabulador * 0) & "ERROR. No se encuentran los parámentros del filtro."
    HuboError = True
    Exit Sub
End If

Flog.writeline

Select Case empzona
   Case 0
        siglazona = "A"
        zonacampo = "bszonaa"
   Case 1
        siglazona = "AB"
        zonacampo = "bszonaab"
   Case 2
        siglazona = "B"
        zonacampo = "bszonab"
   Case 3
        siglazona = "BC"
        zonacampo = "bszonabc"
   Case 4
        siglazona = "C"
        zonacampo = "bszonac"
End Select

'------------------------------------------------------------------------------------
'Configuracion del Reporte
'------------------------------------------------------------------------------------
Flog.writeline
Flog.writeline Espacios(Tabulador * 0) & "Buscando configuración de Reporte 265."

Dim ac1 As Integer
Dim ac2 As Integer
Dim ac3 As Integer

Dim e1 As String
Dim e2 As String
Dim e3 As String

Dim teDesc1 As String
Dim teDesc2 As String
Dim teDesc3 As String
Dim ordenEst As String
Dim total As Double

Dim ordenCab As Integer

StrSql = "SELECT * "
StrSql = StrSql & " FROM confrep"
StrSql = StrSql & " WHERE repnro = 265"
StrSql = StrSql & " ORDER BY confnrocol"
OpenRecordset StrSql, rs_Consult

If rs_Consult.EOF Then
    Flog.writeline Espacios(Tabulador * 1) & "No se encontró la configuración del Reporte"
    Exit Sub
End If
    
Do While (Not rs_Consult.EOF)
        
    Select Case UCase(rs_Consult!conftipo)
        Case "AC1"
            ac1 = rs_Consult!confval
            e1 = IIf(EsNulo(rs_Consult!confetiq), "", rs_Consult!confetiq)
            Flog.writeline Espacios(Tabulador * 1) & "AC1:" & ac1 & " - " & e1
        Case "AC2"
            ac2 = rs_Consult!confval
            e2 = IIf(EsNulo(rs_Consult!confetiq), 0, rs_Consult!confetiq)
            Flog.writeline Espacios(Tabulador * 1) & "AC2:" & ac2 & " - " & e2
        Case "AC3"
            ac3 = rs_Consult!confval
            e3 = IIf(EsNulo(rs_Consult!confetiq), 0, rs_Consult!confetiq)
            Flog.writeline Espacios(Tabulador * 1) & "AC3:" & ac3 & " - " & e3
    End Select
    
    rs_Consult.MoveNext

Loop

'---------------------------------------------------------------------------------------
' Guardo el registro cabecera del proceso
'--------------------------------4-------------------------------------------------------
'FGZ - cambié esto.
'ultimodia = Day(DateSerial(anio, mes, 0))

'ultimodia = Day(DateSerial(anio, mes, 0))
'fecestr = ultimodia & "/" & mes & "/" & anio

If mes = 12 Or mes = 1 Then
    fecestr = CDate("31" & "/" & Format(mes, "00") & "/" & Format(anio, "0000"))
Else
    fecestr = CDate("01" & "/" & Format(mes + 1, "00") & "/" & Format(anio, "0000")) - 1
End If

StrSql = " INSERT INTO rep_bandaext (bpronro,repdes,repfec,fecha,hora,iduser,etiq1,etiq2,etiq3) VALUES "
StrSql = StrSql & " (" & bpronro & ",'" & titulofiltro & "'," & ConvFecha(fecestr) & "," & ConvFecha(Fecha) & ",'" & hora & "','" & idusuario & "','" & e1 & "','" & e2 & "','" & e3 & "')"
objconnProgreso.Execute StrSql, , adExecuteNoRecords


'------------------------------------------------------------------------------------
' Filtro los empleados
'------------------------------------------------------------------------------------

Flog.writeline
Flog.writeline Espacios(Tabulador * 0) & " Calculo los empleado que hacen match con el filtro."
teDesc1 = ""
teDesc2 = ""
teDesc3 = ""
StrSql = " SELECT DISTINCT empleg, grado.gradesabr, puesto.puedesc,ternom,terape, acu_mes1.ammonto monto1,empleado.ternro, "
StrSql = StrSql & zonacampo & " mediana "
If Tenro3 <> 0 Then
    StrSql = StrSql & ",tenro1.tenro t1, tenro1.estrnro e1,tenro2.tenro t2, tenro2.estrnro e2,tenro3.tenro t3, tenro3.estrnro e3 "
Else
    If Tenro2 <> 0 Then
        StrSql = StrSql & ",tenro1.tenro t1, tenro1.estrnro e1,tenro2.tenro t2, tenro2.estrnro e2 "
    Else
        If Tenro1 <> 0 Then
            StrSql = StrSql & ",tenro1.tenro t1, tenro1.estrnro e1 "
        End If
    End If
End If


If ac2 <> 0 Then
    StrSql = StrSql & ",acu_mes2.ammonto monto2"
End If

If ac3 <> 0 Then
    StrSql = StrSql & ",acu_mes3.ammonto monto3"
End If

StrSql = StrSql & " FROM empleado "

StrSql = StrSql & " INNER JOIN acu_mes acu_mes1 ON acu_mes1.ternro = empleado.ternro AND acu_mes1.ammes = " & mes & " AND acu_mes1.amanio = " & anio & " AND acu_mes1.acunro = " & ac1
 
If ac2 <> 0 Then
    StrSql = StrSql & " LEFT JOIN acu_mes acu_mes2 ON acu_mes2.ternro = empleado.ternro AND acu_mes2.ammes = " & mes & " AND acu_mes2.amanio = " & anio & " AND acu_mes2.acunro = " & ac2
End If

If ac3 <> 0 Then
    StrSql = StrSql & " LEFT JOIN acu_mes acu_mes3 ON acu_mes3.ternro = empleado.ternro AND acu_mes3.ammes = " & mes & " AND acu_mes3.amanio = " & anio & " AND acu_mes3.acunro = " & ac3
End If

StrSql = StrSql & " INNER JOIN his_estructura estpuesto ON empleado.ternro = estpuesto.ternro"
StrSql = StrSql & " AND estpuesto.tenro = 4 "
StrSql = StrSql & " AND estpuesto.htetdesde <= " & ConvFecha(fecestr) & " AND (estpuesto.htethasta IS NULL OR estpuesto.htethasta >= " & ConvFecha(fecestr) & ")"
StrSql = StrSql & " INNER JOIN puesto ON estpuesto.estrnro = puesto.estrnro "
'StrSql = StrSql & " INNER JOIN puesto_grado ON puesto_grado.puenro = puesto.puenro "
'Agregado 05/12/2013
StrSql = StrSql & " INNER JOIN puesto_grado ON puesto_grado.puenro = puesto.puenro and puesto_grado.granro = empleado.granro "
'fin
StrSql = StrSql & " INNER JOIN grado ON grado.granro = puesto_grado.granro "
StrSql = StrSql & " INNER JOIN banda_salarial ON banda_salarial.granro = grado.granro"
If bsinterna = 0 Then
    StrSql = StrSql & " AND banda_salarial.obnro = " & obnro
Else
    StrSql = StrSql & " AND banda_salarial.bsinterna = -1 "
End If
If Tenro1 <> 0 Then
    StrSql = StrSql & " INNER JOIN his_estructura tenro1 ON empleado.ternro = tenro1.ternro"
    StrSql = StrSql & " AND tenro1.tenro = " & Tenro1
    StrSql = StrSql & " AND tenro1.htetdesde <= " & ConvFecha(fecestr) & " AND (tenro1.htethasta IS NULL OR tenro1.htethasta >= " & ConvFecha(fecestr) & ")"
    If Estrnro1 <> 0 Then
        StrSql = StrSql & " AND tenro1.estrnro = " & Estrnro1
    End If
End If

If Tenro2 <> 0 Then
    StrSql = StrSql & " INNER JOIN his_estructura tenro2 ON empleado.ternro = tenro2.ternro"
    StrSql = StrSql & " AND tenro2.tenro = " & Tenro2
    StrSql = StrSql & " AND tenro2.htetdesde <= " & ConvFecha(fecestr) & " AND (tenro2.htethasta IS NULL OR tenro2.htethasta >= " & ConvFecha(fecestr) & ")"
    If Estrnro2 <> 0 Then
        StrSql = StrSql & " AND tenro2.estrnro = " & Estrnro2
    End If
End If
If Tenro3 <> 0 Then
    StrSql = StrSql & " INNER JOIN his_estructura tenro3 ON empleado.ternro = tenro3.ternro"
    StrSql = StrSql & " AND tenro3.tenro = " & Tenro3
    StrSql = StrSql & " AND tenro3.htetdesde <= " & ConvFecha(fecestr) & " AND (tenro3.htethasta IS NULL OR tenro3.htethasta >= " & ConvFecha(fecestr) & ")"
    If Estrnro3 <> 0 Then
        StrSql = StrSql & " AND tenro3.estrnro = " & Estrnro3
    End If
End If

StrSql = StrSql & " WHERE " & filtro & " AND empzona = '" & siglazona & "'"

If Tenro3 <> 0 Then
    StrSql = StrSql & " ORDER BY t1,e1,t2,e2,t3,e3," & Orden
Else
    If Tenro2 <> 0 Then
        StrSql = StrSql & " ORDER BY t1,e1,t2,e2," & Orden
    Else
        If Tenro1 <> 0 Then
             StrSql = StrSql & " ORDER BY t1,e1," & Orden
        
        Else
            StrSql = StrSql & " ORDER BY " & Orden
        End If
    End If
End If


OpenRecordset StrSql, rs_Consult

'------------------------------------------------------------------------------------
'configuracion de las variables de progreso
'------------------------------------------------------------------------------------

Flog.writeline
Flog.writeline Espacios(Tabulador * 0) & "Configurando progreso"
Progreso = 0
CEmpleadosAProc = 1

CEmpleadosAProc = rs_Consult.RecordCount

If CEmpleadosAProc = 0 Then
   CEmpleadosAProc = 1
End If

IncPorc = (100 / CEmpleadosAProc)
Flog.writeline Espacios(Tabulador * 1) & "Cantidad de empleados: " & CEmpleadosAProc
           
Do While Not rs_Consult.EOF
                                                                                                                                                        
    If Not IsNull(rs_Consult!monto1) Then
        col1 = rs_Consult!monto1
        filanro = 1
    Else
        col1 = 0
    End If
    
    'FGZ - 07/10/2009 - le agregué este control
    If ac2 <> 0 Then
        If Not IsNull(rs_Consult!monto2) Then
            filanro = 2
            col2 = rs_Consult!monto2
        Else
            col2 = 0
        End If
    Else
        col2 = 0
    End If

    'FGZ - 07/10/2009 - le agregué este control
    If ac3 <> 0 Then
        If Not IsNull(rs_Consult!monto3) Then
            filanro = 3
            col3 = rs_Consult!monto3
        Else
            col3 = 0
        End If
    Else
        col3 = 0
    End If
    
    mediana = rs_Consult!mediana
    
    total = col1 + col2 + col3
    dif = mediana - total
    
    If dif < 0 Then
        referencia = "Supera"
    Else
        If dif > 0 Then
            referencia = "Por Debajo"
        Else
            referencia = "Iguala"
        End If
    End If
    
    porcdif = (dif * 100) / total
    
    
    StrSql = " INSERT INTO rep_bandaext_det (bpronro,mes,anio,tenro1,estrnro1,tenro2,estrnro2,tenro3,estrnro3,filanro,ternro,legajo,apellido,nombre,col1,col2,col3,puesto,grado,zona,mediana,dif,porcdif,referencia) "
    StrSql = StrSql & " VALUES (" & NroProcesoBatch & "," & mes & "," & anio & "," & Tenro1 & "," & Estrnro1 & "," & Tenro2 & "," & Estrnro2 & "," & Tenro3 & "," & Estrnro3 & "," & filanro & "," & rs_Consult!ternro & "," & rs_Consult!empleg & ",'" & rs_Consult!terape & "','" & rs_Consult!ternom & "'," & col1 & "," & col2 & "," & col3 & ",'" & rs_Consult!puedesc & "','" & rs_Consult!gradesabr & "','" & siglazona & " '," & mediana & "," & dif & "," & porcdif & ",'" & referencia & "')"
    OpenRecordset StrSql, rs_insert
    
    'Actualizo el progreso-----------------------------------------------------------
    Progreso = Progreso + IncPorc
    CEmpleadosAProc = CEmpleadosAProc - 1
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
    ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
    "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    rs_Consult.MoveNext
    
Loop

'--------------------------------------------------------------------------------

If rs_Consult.State = adStateOpen Then rs_Consult.Close
Set rs_Consult = Nothing

Exit Sub

CE:
    Flog.writeline "=================================================================="
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

 
