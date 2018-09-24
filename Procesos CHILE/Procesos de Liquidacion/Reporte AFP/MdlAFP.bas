Attribute VB_Name = "MdlAFP"
Option Explicit

'Const Version = 1.09
'Const FechaVersion = "31/01/2007"
'Autor = Martin Ferraro
'Reporte AFP para Chile

Const Version = "1.10"
Const FechaVersion = "12/11/2015"
'Autor = Dimatz Rafael - CAS 32780 - Se corrigio el Representante Legal, para que muestre el configurado por sistema

Public Type TMovimiento
    CodMov As Integer
    FecIniMov As String
    FecFinMov As String
End Type

'Arreglo de movimientos de empleados
Dim ArrMov(80) As TMovimiento
Dim TopeArrMov As Integer



Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial del Generador de Reporte AFP.
' Autor      : Martin Ferraro
' Fecha      : 31/01/2007
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

    strCmdLine = Command()
    ArrParametros = Split(strCmdLine, " ", -1)
    If UBound(ArrParametros) > 0 Then
        If IsNumeric(ArrParametros(0)) Then
            NroProcesoBatch = ArrParametros(0)
            Etiqueta = ArrParametros(1)
            EncriptStrconexion = CBool(ArrParametros(2))
            c_seed = ArrParametros(2)
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
    
    Nombre_Arch = PathFLog & "Generacion_Reporte_AFP" & "-" & NroProcesoBatch & ".log"
    
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
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 143 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        bprcparam = rs_batch_proceso!bprcparam
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        Call Afp(NroProcesoBatch, bprcparam)
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


Public Sub Afp(ByVal bpronro As Long, ByVal Parametros As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del reporte del Afps
' Autor      : Martin Ferraro
' Fecha      : 31/01/2007
' --------------------------------------------------------------------------------------------


Dim CantMov As Integer

'Arreglo que contiene los parametros
Dim arrParam
Dim I As Long

'Parametros desde ASP
Dim FiltroSql As String
Dim EstrnroAFP As Long
Dim Empresa As Long
Dim Periodo As Long
Dim ListaProc As String
Dim Tenro1 As Long
Dim Estrnro1 As Long
Dim Tenro2 As Long
Dim Estrnro2 As Long
Dim Tenro3 As Long
Dim Estrnro3 As Long
Dim FecEstr As Date
Dim TituloFiltro As String
Dim OrdenSql As String

'RecordSet
Dim rs_Empleados As New ADODB.Recordset
Dim rs_Confrep As New ADODB.Recordset
Dim rs_Consult As New ADODB.Recordset
Dim rs_Monto As New ADODB.Recordset
Dim rs_Estructura As New ADODB.Recordset

'Variables
Dim FecDesdePer As Date
Dim FecHastaPer As Date
Dim UltimoEmpleado As Long
Dim UltimaAFP As Long
Dim EmpreNomb As String
Dim EmpreRUT As String
Dim topeArreglo As Integer
Dim arreglo(80) As Double
Dim contador As Integer
Dim Rut As String
Dim PliqDesc As String
Dim terape As String
Dim ternom As String
Dim Orden As Long
Dim Tipo_Doc

' Inicio codigo ejecutable
On Error GoTo CE

' Levanto cada parametro por separado, el separador de parametros es "@"
Flog.writeline Espacios(Tabulador * 0) & "levantando parametros" & Parametros
If Not IsNull(Parametros) Then
    
    arrParam = Split(Parametros, "@")
    
    If UBound(arrParam) = 13 Then
    
        FiltroSql = arrParam(0)
        EstrnroAFP = CLng(arrParam(1))
        Empresa = arrParam(2)
        Periodo = CLng(arrParam(3))
        ListaProc = arrParam(4)
        ListaProc = Replace(ListaProc, "U", ",")
        Tenro1 = CLng(arrParam(5))
        Estrnro1 = CLng(arrParam(6))
        Tenro2 = CLng(arrParam(7))
        Estrnro2 = CLng(arrParam(8))
        Tenro3 = CLng(arrParam(9))
        Estrnro3 = CLng(arrParam(10))
        FecEstr = CDate(arrParam(11))
        TituloFiltro = arrParam(12)
        OrdenSql = arrParam(13)
    
        Flog.writeline Espacios(Tabulador * 1) & "Filtro = " & FiltroSql
        Flog.writeline Espacios(Tabulador * 1) & "Afp = " & EstrnroAFP
        Flog.writeline Espacios(Tabulador * 1) & "Empresa = " & Empresa
        Flog.writeline Espacios(Tabulador * 1) & "Periodo = " & Periodo
        Flog.writeline Espacios(Tabulador * 1) & "Procesos = " & ListaProc
        Flog.writeline Espacios(Tabulador * 1) & "TE1 = " & Tenro1
        Flog.writeline Espacios(Tabulador * 1) & "Estr1 = " & Estrnro1
        Flog.writeline Espacios(Tabulador * 1) & "TE2 = " & Tenro2
        Flog.writeline Espacios(Tabulador * 1) & "Estr2 = " & Estrnro2
        Flog.writeline Espacios(Tabulador * 1) & "TE3 = " & Tenro3
        Flog.writeline Espacios(Tabulador * 1) & "Estr3 = " & Estrnro3
        Flog.writeline Espacios(Tabulador * 1) & "Fecha Estr =" & FecEstr
        Flog.writeline Espacios(Tabulador * 1) & "Titulo = " & TituloFiltro
        Flog.writeline Espacios(Tabulador * 1) & "Orden = " & OrdenSql
        
    Else
        Flog.writeline Espacios(Tabulador * 0) & "ERROR. La cantidad de parametros no es la esperada."
        Exit Sub
        
    End If
Else
    Flog.writeline Espacios(Tabulador * 0) & "ERROR. No se encuentran los paramentros."
    Exit Sub
End If


Flog.writeline


'cargo el periodo
Flog.writeline Espacios(Tabulador * 0) & "Buscando Periodo."
StrSql = "SELECT * FROM periodo WHERE pliqnro = " & Periodo
OpenRecordset StrSql, rs_Consult

If rs_Consult.EOF Then
    Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontró el Periodo."
    Exit Sub
Else
    FecDesdePer = rs_Consult!pliqdesde
    FecHastaPer = rs_Consult!pliqhasta
    PliqDesc = IIf(EsNulo(rs_Consult!PliqDesc), "", rs_Consult!PliqDesc)
    Flog.writeline Espacios(Tabulador * 1) & "Periodo: " & PliqDesc
End If
rs_Consult.Close


'Cargo el nombre de la empresa
Flog.writeline Espacios(Tabulador * 0) & "Buscando Empresa."
StrSql = "SELECT * FROM estructura WHERE tenro = 10 AND estrnro = " & Empresa
OpenRecordset StrSql, rs_Consult

If rs_Consult.EOF Then
    Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontró la empresa."
    Exit Sub
Else
    EmpreNomb = IIf(EsNulo(rs_Consult!estrdabr), "", rs_Consult!estrdabr)
    Flog.writeline Espacios(Tabulador * 1) & "Empresa: " & EmpreNomb
End If
rs_Consult.Close

'Configuracion del Confrep Tipo de Documento
StrSql = "SELECT *"
StrSql = StrSql & " FROM confrepadv"
StrSql = StrSql & " WHERE repnro = 49 AND confnrocol = 7 "
OpenRecordset StrSql, rs_Confrep

If rs_Confrep.EOF Then
    Flog.writeline "No se encontró configurado el Documento de la Empresa"
    Exit Sub
Else
    Tipo_Doc = rs_Confrep!confval
End If

'Cargo el RUT de la empresa
Flog.writeline Espacios(Tabulador * 0) & "Buscando RUT de la Empresa."
StrSql = "SELECT * FROM empresa "
StrSql = StrSql & "INNER JOIN ter_doc ON ter_doc.ternro = empresa.ternro "
StrSql = StrSql & "AND ter_doc.tidnro = " & Tipo_Doc
StrSql = StrSql & "WHERE empresa.estrnro = " & Empresa
OpenRecordset StrSql, rs_Consult

If rs_Consult.EOF Then
    Flog.writeline Espacios(Tabulador * 1) & "No se encontró RUT de la empresa."
Else
    EmpreRUT = IIf(EsNulo(rs_Consult!nrodoc), "", rs_Consult!nrodoc)
    Flog.writeline Espacios(Tabulador * 1) & "RUT Empresa: " & EmpreRUT
End If
rs_Consult.Close


'Configuracion del Reporte
StrSql = "SELECT *"
StrSql = StrSql & " FROM confrepadv"
StrSql = StrSql & " WHERE repnro = 49"
StrSql = StrSql & " AND (conftipo = 'AC' OR conftipo = 'CO')"
StrSql = StrSql & " AND (conftipo = 'AC' OR conftipo = 'CO')"
StrSql = StrSql & " AND confnrocol <= 6 "
OpenRecordset StrSql, rs_Confrep

If rs_Confrep.EOF Then
    Flog.writeline "No se encontró la configuración del Reporte"
    Exit Sub
End If
  
  
'Comienzo la transaccion
MyBeginTrans

'Inicializacion de variables
topeArreglo = 10 'Valor maximo de columnas del confrep
TopeArrMov = 20
UltimoEmpleado = -1
UltimaAFP = -1
Orden = 1

'---------------------------------------------------------------------------------
'Consulta Principal
'---------------------------------------------------------------------------------
StrSql = "SELECT cabliq.pronro, cabliq.cliqnro, empleado.ternro, empleado.empleg, empleado.terape, empleado.terape2, "
StrSql = StrSql & " empleado.ternom, empleado.ternom2, proceso.prodesc, afp.estrnro cajubnro, afpestr.estrdabr cajubdesc "
StrSql = StrSql & " FROM proceso "
StrSql = StrSql & " INNER JOIN cabliq ON cabliq.pronro = proceso.pronro "
StrSql = StrSql & " INNER JOIN empleado ON cabliq.empleado = empleado.ternro "
'Filtro Empresa
StrSql = StrSql & " INNER JOIN his_estructura empresa ON empleado.ternro = empresa.ternro "
StrSql = StrSql & " AND empresa.estrnro = " & Empresa & " AND empresa.tenro=10 "
StrSql = StrSql & " AND empresa.htetdesde <= " & ConvFecha(FecDesdePer) & " AND (empresa.htethasta IS NULL OR empresa.htethasta >= " & ConvFecha(FecDesdePer) & ") "
'Filtro Caja Jubilacion (AFP)
StrSql = StrSql & " INNER JOIN his_estructura afp ON empleado.ternro = afp.ternro "
StrSql = StrSql & " AND afp.tenro = 15 "
StrSql = StrSql & " AND afp.htetdesde <= " & ConvFecha(FecDesdePer) & " AND (afp.htethasta IS NULL OR afp.htethasta >= " & ConvFecha(FecDesdePer) & ") "
If EstrnroAFP <> 0 Then
    StrSql = StrSql & " AND afp.estrnro = " & EstrnroAFP
End If
StrSql = StrSql & " INNER JOIN estructura afpestr ON afpestr.estrnro = afp.estrnro "
'Filtros de niveles de estructura
If Tenro1 <> 0 Then
    StrSql = StrSql & " INNER JOIN his_estructura tenro1 ON empleado.ternro = tenro1.ternro "
    StrSql = StrSql & " AND tenro1.tenro = " & Tenro1
    StrSql = StrSql & " AND tenro1.htetdesde <= " & ConvFecha(FecEstr) & " AND (tenro1.htethasta IS NULL OR tenro1.htethasta >= " & ConvFecha(FecEstr) & ") "
    If Estrnro1 <> 0 Then
        StrSql = StrSql & " AND tenro1.estrnro = " & Estrnro1
    End If
End If
If Tenro2 <> 0 Then
    StrSql = StrSql & " INNER JOIN his_estructura tenro2 ON empleado.ternro = tenro2.ternro "
    StrSql = StrSql & " AND tenro2.tenro = " & Tenro2
    StrSql = StrSql & " AND tenro2.htetdesde <= " & ConvFecha(FecEstr) & " AND (tenro2.htethasta IS NULL OR tenro2.htethasta >= " & ConvFecha(FecEstr) & ") "
    If Estrnro2 <> 0 Then
        StrSql = StrSql & " AND tenro2.estrnro = " & Estrnro2
    End If
End If
If Tenro3 <> 0 Then
    StrSql = StrSql & " INNER JOIN his_estructura tenro3 ON empleado.ternro = tenro3.ternro "
    StrSql = StrSql & " AND tenro3.tenro = " & Tenro3
    StrSql = StrSql & " AND tenro3.htetdesde <= " & ConvFecha(FecEstr) & " AND (tenro3.htethasta IS NULL OR tenro3.htethasta >= " & ConvFecha(FecEstr) & ") "
    If Estrnro3 <> 0 Then
        StrSql = StrSql & " AND tenro3.estrnro = " & Estrnro3
    End If
End If
'Filtro Empleados
StrSql = StrSql & " WHERE " & FiltroSql
'Filtro Periodo
StrSql = StrSql & " AND proceso.pliqnro =" & Periodo
'Filtro Procesos
StrSql = StrSql & " AND proceso.pronro IN (" & ListaProc & ")"
StrSql = StrSql & " ORDER BY cajubdesc , " & OrdenSql & " , proceso.pronro"
OpenRecordset StrSql, rs_Empleados


'seteo de las variables de progreso
Progreso = 0
CEmpleadosAProc = rs_Empleados.RecordCount
If CEmpleadosAProc = 0 Then
   Flog.writeline "no hay empleados"
   CEmpleadosAProc = 1
End If
IncPorc = (100 / CEmpleadosAProc)
        
Flog.writeline
Flog.writeline
Flog.writeline
Flog.writeline Espacios(Tabulador * 0) & "--------------------------------------------------------"
Flog.writeline Espacios(Tabulador * 0) & "Comienza el procesamiento de empleados."
Flog.writeline Espacios(Tabulador * 0) & "--------------------------------------------------------"
Flog.writeline


'Comienzo a procesar los empleados
Do While Not rs_Empleados.EOF
    
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "PROCESANDO: " & rs_Empleados!empleg & "  - " & rs_Empleados!terape & " " & rs_Empleados!ternom & " PROCESO: " & rs_Empleados!pronro & " " & rs_Empleados!prodesc
    Flog.writeline Espacios(Tabulador * 0) & "-------------------------------------------------------------"
    
    
    rs_Confrep.MoveFirst
    
    
    'Verificacion de cambio de empleado
    If rs_Empleados!ternro <> UltimoEmpleado Then
    
        UltimoEmpleado = rs_Empleados!ternro
        
        'Inicializo arreglo de montos
        For contador = 1 To Max(topeArreglo, TopeArrMov)
            arreglo(contador) = 0
            ArrMov(contador).CodMov = 0
            ArrMov(contador).FecIniMov = ""
            ArrMov(contador).FecFinMov = ""
        Next contador
        
    End If
        
            
    Flog.writeline Espacios(Tabulador * 1) & "Buscando los CO y AC configurados en confrep"
    Do While Not rs_Confrep.EOF
        Select Case UCase(rs_Confrep!conftipo)
        Case "AC":
            StrSql = "SELECT * FROM acu_liq WHERE cliqnro = " & rs_Empleados!cliqnro
            StrSql = StrSql & " AND acunro =" & rs_Confrep!confval
            OpenRecordset StrSql, rs_Monto
            If Not rs_Monto.EOF Then
                If rs_Monto!almonto <> 0 Then
                    arreglo(rs_Confrep!confnrocol) = arreglo(rs_Confrep!confnrocol) + rs_Monto!almonto
                End If
            End If
            rs_Monto.Close
            
        Case "CO":
            StrSql = "SELECT * FROM concepto "
            StrSql = StrSql & " WHERE concepto.conccod = '" & rs_Confrep!confval2 & "'"
            OpenRecordset StrSql, rs_Consult
            If Not rs_Consult.EOF Then
                StrSql = "SELECT * FROM detliq WHERE concnro = " & rs_Consult!concnro
                StrSql = StrSql & " AND cliqnro =" & rs_Empleados!cliqnro
                OpenRecordset StrSql, rs_Monto
                If Not rs_Monto.EOF Then
                    If rs_Monto!dlimonto <> 0 Then
                        arreglo(rs_Confrep!confnrocol) = arreglo(rs_Confrep!confnrocol) + rs_Monto!dlimonto
                    End If
                End If
                rs_Monto.Close
            End If
            rs_Consult.Close
            
        Case Else
        
        End Select
    
        rs_Confrep.MoveNext
        
    Loop
    
    'Reviso si es el ultimo empleado para calcular datos faltantes y guardar
    If EsElUltimoEmpleado(rs_Empleados, UltimoEmpleado) Then
        
        Flog.writeline Espacios(Tabulador * 1) & "Buscando Rut."
        StrSql = " SELECT ter_doc.nrodoc FROM ter_doc "
        StrSql = StrSql & " WHERE ter_doc.ternro= " & rs_Empleados!ternro
        StrSql = StrSql & " AND ter_doc.tidnro = 1 "
        OpenRecordset StrSql, rs_Consult
        
        If Not rs_Consult.EOF Then
            Rut = IIf(EsNulo(rs_Consult!nrodoc), "", rs_Consult!nrodoc)
            Flog.writeline Espacios(Tabulador * 2) & "Rut: " & rs_Consult!nrodoc
        Else
            Flog.writeline Espacios(Tabulador * 2) & "No se encontro Rut."
        End If

        
        Flog.writeline Espacios(Tabulador * 1) & "Buscando Movimientos."
        CantMov = 0
        Call BuscarMov(rs_Empleados!ternro, FecDesdePer, FecHastaPer, CantMov)
        
        'Verifico si cambio de Obra social para crear la cabecera
        If UltimaAFP <> rs_Empleados!cajubnro Then
            
            'Inserto la cabecera
            Flog.writeline Espacios(Tabulador * 1) & "Insertando AFP " & rs_Empleados!cajubnro & " - " & rs_Empleados!cajubdesc
            StrSql = " INSERT INTO rep_afp "
            StrSql = StrSql & " ("
            StrSql = StrSql & " bpronro,"
            StrSql = StrSql & " rut,"
            StrSql = StrSql & " afpnro,"
            StrSql = StrSql & " afpnom,"
            StrSql = StrSql & " periodo,"
            StrSql = StrSql & " empresa,"
            StrSql = StrSql & " descripcion"
            StrSql = StrSql & " )"
            StrSql = StrSql & " VALUES ("
            StrSql = StrSql & NroProcesoBatch
            StrSql = StrSql & " , '" & Mid(EmpreRUT, 1, 15) & "'"
            StrSql = StrSql & " , " & rs_Empleados!cajubnro
            StrSql = StrSql & " , '" & Mid(rs_Empleados!cajubdesc, 1, 50) & "'"
            StrSql = StrSql & " , '" & Mid(PliqDesc, 1, 50) & "'"
            StrSql = StrSql & " , '" & Mid(EmpreNomb, 1, 50) & "'"
            StrSql = StrSql & " , '" & Mid(TituloFiltro, 1, 100) & "'"
            StrSql = StrSql & " )"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            UltimaAFP = rs_Empleados!cajubnro
            Orden = 1
        
        End If
        
        terape = IIf(EsNulo(rs_Empleados!terape), "", rs_Empleados!terape) & IIf(EsNulo(rs_Empleados!terape2), "", " " & rs_Empleados!terape2)
        ternom = IIf(EsNulo(rs_Empleados!ternom), "", rs_Empleados!ternom) & IIf(EsNulo(rs_Empleados!ternom2), "", " " & rs_Empleados!ternom2)
        'Inserto la detalle
        Flog.writeline Espacios(Tabulador * 1) & "Insertando Empleado "
        StrSql = " INSERT INTO rep_afp_det "
        StrSql = StrSql & " ("
        StrSql = StrSql & " bpronro,"
        StrSql = StrSql & " ternro,"
        StrSql = StrSql & " orden,"
        StrSql = StrSql & " rut,"
        StrSql = StrSql & " afpnro,"
        StrSql = StrSql & " legajo,"
        StrSql = StrSql & " apellido,"
        StrSql = StrSql & " nombre,"
        StrSql = StrSql & " codigo,"
        StrSql = StrSql & " fecini,"
        StrSql = StrSql & " fecfin,"
        StrSql = StrSql & " val_col1,"
        StrSql = StrSql & " val_col2,"
        StrSql = StrSql & " val_col3,"
        StrSql = StrSql & " val_col4,"
        StrSql = StrSql & " val_col5,"
        StrSql = StrSql & " val_col6"
        StrSql = StrSql & " )"
        StrSql = StrSql & " VALUES ("
        StrSql = StrSql & NroProcesoBatch
        StrSql = StrSql & " , " & rs_Empleados!ternro
        StrSql = StrSql & " , " & Orden
        StrSql = StrSql & " , '" & Mid(Rut, 1, 12) & "'"
        StrSql = StrSql & " , " & rs_Empleados!cajubnro
        StrSql = StrSql & " , " & rs_Empleados!empleg
        StrSql = StrSql & " , '" & Mid(terape, 1, 50) & "'"
        StrSql = StrSql & " , '" & Mid(ternom, 1, 50) & "'"
        StrSql = StrSql & " , " & ArrMov(1).CodMov
        If ArrMov(1).FecIniMov = "" Then
            StrSql = StrSql & ", null"
        Else
            StrSql = StrSql & ", " & ConvFecha(CDate(ArrMov(1).FecIniMov))
        End If
        If ArrMov(1).FecFinMov = "" Then
            StrSql = StrSql & ", null"
        Else
            StrSql = StrSql & ", " & ConvFecha(CDate(ArrMov(1).FecFinMov))
        End If
        StrSql = StrSql & " , " & arreglo(1)
        StrSql = StrSql & " , " & arreglo(2)
        StrSql = StrSql & " , " & arreglo(3)
        StrSql = StrSql & " , " & arreglo(4)
        StrSql = StrSql & " , " & arreglo(5)
        StrSql = StrSql & " , " & arreglo(6)
        StrSql = StrSql & " )"
        objConn.Execute StrSql, , adExecuteNoRecords
            
        Orden = Orden + 1
        
        'Si hay mas de un movimiento repito los registros
        contador = 2
        Do While contador <= CantMov
            StrSql = " INSERT INTO rep_afp_det "
            StrSql = StrSql & " ("
            StrSql = StrSql & " bpronro,"
            StrSql = StrSql & " ternro,"
            StrSql = StrSql & " orden,"
            StrSql = StrSql & " rut,"
            StrSql = StrSql & " afpnro,"
            StrSql = StrSql & " legajo,"
            StrSql = StrSql & " apellido,"
            StrSql = StrSql & " nombre,"
            StrSql = StrSql & " codigo,"
            StrSql = StrSql & " fecini,"
            StrSql = StrSql & " fecfin,"
            StrSql = StrSql & " val_col1,"
            StrSql = StrSql & " val_col2,"
            StrSql = StrSql & " val_col3,"
            StrSql = StrSql & " val_col4,"
            StrSql = StrSql & " val_col5,"
            StrSql = StrSql & " val_col6"
            StrSql = StrSql & " )"
            StrSql = StrSql & " VALUES ("
            StrSql = StrSql & NroProcesoBatch
            StrSql = StrSql & " , " & rs_Empleados!ternro
            StrSql = StrSql & " , " & Orden
            StrSql = StrSql & " , '" & Mid(Rut, 1, 12) & "'"
            StrSql = StrSql & " , " & rs_Empleados!cajubnro
            StrSql = StrSql & " , " & rs_Empleados!empleg
            StrSql = StrSql & " , '" & Mid(terape, 1, 50) & "'"
            StrSql = StrSql & " , '" & Mid(ternom, 1, 50) & "'"
            StrSql = StrSql & " , " & ArrMov(contador).CodMov
            If ArrMov(contador).FecIniMov = "" Then
                StrSql = StrSql & ", null"
            Else
                StrSql = StrSql & ", " & ConvFecha(CDate(ArrMov(contador).FecIniMov))
            End If
            If ArrMov(contador).FecFinMov = "" Then
                StrSql = StrSql & ", null"
            Else
                StrSql = StrSql & ", " & ConvFecha(CDate(ArrMov(contador).FecFinMov))
            End If
            StrSql = StrSql & " , 0 "
            StrSql = StrSql & " , 0 "
            StrSql = StrSql & " , 0 "
            StrSql = StrSql & " , 0 "
            StrSql = StrSql & " , 0 "
            StrSql = StrSql & " , 0 "
            StrSql = StrSql & " )"
            objConn.Execute StrSql, , adExecuteNoRecords

            Orden = Orden + 1
            contador = contador + 1
        Loop
    End If
      
    'Actualizo el progreso
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
             "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    'Paso a siguiente cabliq
    rs_Empleados.MoveNext
    
Loop

'Fin de la transaccion
If Not HuboError Then
    MyCommitTrans
End If


If rs_Empleados.State = adStateOpen Then rs_Empleados.Close
If rs_Confrep.State = adStateOpen Then rs_Confrep.Close
If rs_Monto.State = adStateOpen Then rs_Monto.Close
If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
If rs_Consult.State = adStateOpen Then rs_Consult.Close


Set rs_Empleados = Nothing
Set rs_Monto = Nothing
Set rs_Confrep = Nothing
Set rs_Estructura = Nothing
Set rs_Consult = Nothing

Exit Sub

CE:
    Flog.writeline "=================================================================="
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="
    
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

End Sub


Public Function EsElUltimoEmpleado(ByVal rs As ADODB.Recordset, ByVal Anterior As Long) As Boolean
    rs.MoveNext
    If rs.EOF Then
        EsElUltimoEmpleado = True
    Else
        If rs!ternro <> Anterior Then
            EsElUltimoEmpleado = True
        Else
            EsElUltimoEmpleado = False
        End If
    End If
    rs.MovePrevious
End Function


Public Sub BuscarMov(ByVal ternro As Long, ByVal Desde As Date, ByVal Hasta As Date, ByRef Cantidad As Integer)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que se encarga de buscar las altas y bajas dentro del rango de fechas
' Autor      : Martin Ferraro
' Fecha      : 23/01/2007
' --------------------------------------------------------------------------------------------
Dim rs_Fases As New ADODB.Recordset
Dim Reg0 As TMovimiento
Dim Reg1 As TMovimiento

On Error GoTo ErrorBuscarMov
    
    'Inicializo las salidas
    Reg0.CodMov = 0
    Reg0.FecIniMov = ""
    Reg0.FecFinMov = ""
    Reg1.CodMov = 0
    Reg1.FecIniMov = ""
    Reg1.FecFinMov = ""
    Cantidad = 0
    
    'Busco primer fase de alta dentro del rango
    StrSql = " SELECT * FROM fases"
    StrSql = StrSql & " WHERE fases.empleado = " & ternro
    StrSql = StrSql & " AND fases.altfec >= " & ConvFecha(Desde)
    StrSql = StrSql & " AND fases.altfec <= " & ConvFecha(Hasta)
    StrSql = StrSql & " AND fases.real = -1"
    StrSql = StrSql & " ORDER BY fases.altfec"
    OpenRecordset StrSql, rs_Fases
    
    If Not rs_Fases.EOF Then
        Reg1.CodMov = 1
        Reg1.FecIniMov = IIf(EsNulo(rs_Fases!altfec), "", rs_Fases!altfec)
        
        'Si hay mas de una fase me paro en la ultima
        If rs_Fases.RecordCount > 1 Then
            rs_Fases.MoveLast
        End If
        
        'Miro si la baja esta en el periodo
        If Not IsNull(rs_Fases!bajfec) Then
            If CDate(rs_Fases!bajfec) <= CDate(Hasta) Then
                Reg1.FecFinMov = rs_Fases!bajfec
            End If
        End If
        
        rs_Fases.Close
        
        'Busco baja en el periodo menor al alta encontrado
        StrSql = " SELECT * FROM fases"
        StrSql = StrSql & " WHERE fases.empleado = " & ternro
        StrSql = StrSql & " AND fases.bajfec >= " & ConvFecha(Desde)
        StrSql = StrSql & " AND fases.bajfec < " & ConvFecha(Reg1.FecIniMov)
        StrSql = StrSql & " AND fases.real = -1"
        StrSql = StrSql & " ORDER BY fases.bajfec DESC"
        OpenRecordset StrSql, rs_Fases
    
        If Not rs_Fases.EOF Then
            Reg0.CodMov = 2
            Reg0.FecFinMov = rs_Fases!bajfec
            rs_Fases.Close
            
            'Guardo en el arreglo reg 0
            Cantidad = Cantidad + 1
            ArrMov(Cantidad) = Reg0
            
        End If
        
        'Guardo en el arreglo reg 1
        Cantidad = Cantidad + 1
        ArrMov(Cantidad) = Reg1
        
        
    Else
        
        rs_Fases.Close
        
        'Busco baja en el periodo
        StrSql = " SELECT * FROM fases"
        StrSql = StrSql & " WHERE fases.empleado = " & ternro
        StrSql = StrSql & " AND fases.bajfec >= " & ConvFecha(Desde)
        StrSql = StrSql & " AND fases.bajfec <= " & ConvFecha(Hasta)
        StrSql = StrSql & " AND fases.real = -1"
        StrSql = StrSql & " ORDER BY fases.bajfec DESC"
        OpenRecordset StrSql, rs_Fases
    
        If Not rs_Fases.EOF Then
            Reg0.CodMov = 2
            Reg0.FecFinMov = rs_Fases!bajfec
        
            'Guardo en el arreglo reg 0
            Cantidad = Cantidad + 1
            ArrMov(Cantidad) = Reg0
        End If
        
        rs_Fases.Close
        
    End If
    
    
    
    'Busco todas las licencias <> vacaciones que comienzan en el periodo
    StrSql = " SELECT *"
    StrSql = StrSql & " FROM emp_lic"
    StrSql = StrSql & " WHERE empleado = " & ternro
    StrSql = StrSql & " AND elfechadesde >= " & ConvFecha(Desde)
    StrSql = StrSql & " AND elfechadesde <= " & ConvFecha(Hasta)
    StrSql = StrSql & " AND tdnro <> 2"
    StrSql = StrSql & " ORDER BY elfechadesde"
    
    OpenRecordset StrSql, rs_Fases
    
    Do While Not rs_Fases.EOF
    
        'Guardo registros hasta el tope
        If Cantidad < TopeArrMov Then
            Cantidad = Cantidad + 1
            ArrMov(Cantidad).CodMov = 3
            ArrMov(Cantidad).FecIniMov = IIf(EsNulo(rs_Fases!elfechadesde), "", rs_Fases!elfechadesde)
            ArrMov(Cantidad).FecFinMov = IIf(EsNulo(rs_Fases!elfechahasta), "", rs_Fases!elfechahasta)
        End If
        
        rs_Fases.MoveNext
        
    Loop
    
    rs_Fases.Close
    
    

If rs_Fases.State = adStateOpen Then rs_Fases.Close
Set rs_Fases = Nothing

Exit Sub

ErrorBuscarMov:
Flog.writeline "Error en BuscarMov: " & Err.Description
Flog.writeline "Ultimo SQl Ejecutado: " & StrSql

End Sub


Public Function Max(ByRef valor1 As Integer, ByRef valor2 As Integer) As Integer
    
    If valor1 >= valor2 Then
        Max = valor1
    Else
        Max = valor2
    End If
    
End Function

