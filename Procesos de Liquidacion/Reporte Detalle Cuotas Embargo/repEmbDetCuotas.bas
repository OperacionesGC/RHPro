Attribute VB_Name = "repEmbDetCuotas"
Option Explicit

'Global Const Version = "1.01"
'Global Const FechaModificacion = "26/04/2007"
'Global Const UltimaModificacion = " " ' Martin Ferraro - Version Inicial

Global Const Version = "1.02" ' Cesar Stankunas
Global Const FechaModificacion = "05/08/2009"
Global Const UltimaModificacion = ""    'Encriptacion de string connection

'------------------------------------------------------------------------------------------
Dim fs, f
'Global Flog

Dim NroProceso As Long

Global Path As String
Global HuboErrores As Boolean
Global EmpErrores As Boolean

Global pronro2 As String
Global listaacunro As String

Global titulofiltro As String
Global filtro As String
Global fecestr As String
Global tenro1  As Long
Global estrnro1  As Long
Global tenro2  As Long
Global estrnro2 As Long
Global tenro3 As Long
Global estrnro3 As Long
Global Orden As String
Global fec_desde As String
Global fec_hasta As String
Global terape As String
Global terape2 As String
Global ternom As String
Global ternom2 As String
Global embnro As Long
Global tpenro As Long
Global tpedesc As String
Global desc As Double
Global empleg As Long
Global juzdesabr As String
Global secdesabr As String
Global banco As String
Global EmpNombre As String
Global EmpDire As String
Global EmpLogo As String
Global EmpLogoAlto As Integer
Global EmpLogoAncho As Integer
Global profecpago As String
Global importe As Double
Global embdesext As String
Global embcnro As Integer
Global IdUser As String
Global Fecha As Date
Global Hora As String

Global pliqdesde As Long
Global pDesde As String
Global pHasta As String
Global EmpresaEmpnro As Long


Private Sub Main()

Dim strCmdLine As String
Dim Nombre_Arch As String

Dim Cantidad As Integer
Dim cantidadProcesada As Integer

Dim StrSql As String
Dim StrSql2 As String
Dim objRs As New ADODB.Recordset
Dim objRs2 As New ADODB.Recordset
Dim rsCuotas As New ADODB.Recordset
Dim rsJuzgado As New ADODB.Recordset
Dim rsProceso As New ADODB.Recordset
Dim rsSecretaria As New ADODB.Recordset
Dim rsBanco As New ADODB.Recordset
Dim rs_estructura As New ADODB.Recordset
Dim rs_Domicilio As New ADODB.Recordset
Dim rs_logo As New ADODB.Recordset
Dim I
Dim totalAcum
Dim cantRegistros
Dim PID As String
Dim parametros As String
Dim ArrParametros
Dim fecAuxHasta
Dim fecAuxDesde
Dim auxDesc As Double
Dim listaPeriodo As String
Dim listaProceso As String

    strCmdLine = Command()
    ArrParametros = Split(strCmdLine, " ", -1)
    If UBound(ArrParametros) > 1 Then
        If IsNumeric(ArrParametros(0)) Then
            NroProceso = ArrParametros(0)
            Etiqueta = ArrParametros(1)
            EncriptStrconexion = CBool(ArrParametros(2))
            c_seed = ArrParametros(2)
        Else
            Exit Sub
        End If
    Else
        If UBound(ArrParametros) > 0 Then
            If IsNumeric(ArrParametros(0)) Then
                NroProceso = ArrParametros(0)
                Etiqueta = ArrParametros(1)
            Else
                Exit Sub
            End If
        Else
            If IsNumeric(strCmdLine) Then
                NroProceso = strCmdLine
            Else
                Exit Sub
            End If
        End If
    End If
    
    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas

    Nombre_Arch = PathFLog & "ReporteDetCuotas" & "-" & NroProceso & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    Flog.writeline "_________________________________________________________________"
    Flog.writeline "Version = " & Version
    Flog.writeline "Modificacion = " & UltimaModificacion
    Flog.writeline "Fecha = " & FechaModificacion
    Flog.writeline "_________________________________________________________________"
    Flog.writeline
    
    On Error GoTo CE
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "PID = " & PID
    Flog.writeline
    
    TiempoInicialProceso = GetTickCount
    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    HuboErrores = False
    
    Flog.writeline "Inicio Proceso de Reporte de Detalles de Cuotas de Embargos: " & Now
    Flog.writeline "Cambio el estado del proceso a Procesando"
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & ", bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Obtengo los datos del proceso"
    
    TiempoAcumulado = GetTickCount
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, objRs2
    
    If Not objRs2.EOF Then
       IdUser = objRs2!IdUser
       Fecha = objRs2!bprcfecha
       Hora = objRs2!bprchora
       
       Flog.writeline Espacios(Tabulador * 0) & "Recuperando Parametros."
       'Obtengo los parametros del proceso
       parametros = objRs2!bprcparam
       
        If Not IsNull(parametros) Then
        
            ArrParametros = Split(parametros, "@")
            If UBound(ArrParametros) = 12 Then
                  
               'Obtengo el Titulo
                titulofiltro = ArrParametros(0)
               'Obtengo el filtro utilizado como restricciones a la busqueda de embargos
                filtro = ArrParametros(1)
                If InStr(filtro, "embargo.embest") > 0 Then
                    filtro = Replace(filtro, "embargo.embest = A", "embargo.embest = 'A'")
                    filtro = Replace(filtro, "embargo.embest = E", "embargo.embest = 'E'")
                    filtro = Replace(filtro, "embargo.embest = F", "embargo.embest = 'F'")
                    filtro = Replace(filtro, "embargo.embest = I", "embargo.embest = 'I'")
                End If
                ' Fecha a considerar las estructuras
                fecestr = CDate(ArrParametros(2))
                ' Nro tipo de la primer estructura
                tenro1 = CLng(ArrParametros(3))
                'Codigo de la primer estructura
                estrnro1 = CLng(ArrParametros(4))
                ' Nro tipo de la segunda estructura
                tenro2 = CLng(ArrParametros(5))
                ' Codigo de la segunda estructura
                estrnro2 = CLng(ArrParametros(6))
                ' Nro tipo de la tercera estructura
                tenro3 = CLng(ArrParametros(7))
                ' Codigo de la tercer estructura
                estrnro3 = CLng(ArrParametros(8))
                ' String conteniendo el orden en el cual se debe realizar la busqueda de embargos
                Orden = ArrParametros(9)
                ' Periodo desde
                pliqdesde = CLng(ArrParametros(10))
                ' Periodo Hasta
                pliqhasta = CLng(ArrParametros(11))
                ' empresa
                EmpresaEmpnro = CLng(ArrParametros(12))
                
                Flog.writeline
                Flog.writeline Espacios(Tabulador * 1) & " Datos del Filtro: "
                Flog.writeline Espacios(Tabulador * 1) & "    Filtro: " & ArrParametros(1)
                Flog.writeline Espacios(Tabulador * 1) & "    Fecha Estr: " & ArrParametros(2)
                Flog.writeline Espacios(Tabulador * 1) & "    Estructuras: " & ArrParametros(3) & " - " & ArrParametros(4) & " - " & ArrParametros(5) & " - " & ArrParametros(6) & " - " & ArrParametros(7) & " - " & ArrParametros(8)
                Flog.writeline Espacios(Tabulador * 1) & "    Orden: " & ArrParametros(9)
                Flog.writeline Espacios(Tabulador * 1) & "    Período de Emb.: " & ArrParametros(10) & " - " & ArrParametros(11)
                Flog.writeline Espacios(Tabulador * 1) & "    EmpresaEmpnro: " & ArrParametros(12)
                Flog.writeline
                
            Else
                Flog.writeline Espacios(Tabulador * 1) & "ERROR. La cantidad de parametros no es la esperada."
                Exit Sub
            End If
        
        Else
            Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encuentran los paramentros."
            Exit Sub
        End If
             
        'EMPIEZA EL PROCESO

        '------------------------------------------------------------------------------------------------------------------------
        'BUSCANDO DATOS DE LOS PERIODOS
        '------------------------------------------------------------------------------------------------------------------------
        Flog.writeline Espacios(Tabulador * 0) & "Buscando datos periodos."
        
        'Busco la fecha del periodo desde
        StrSql = "SELECT pliqdesde, pliqdesc, pliqnro FROM periodo WHERE periodo.pliqnro = " & pliqdesde
        OpenRecordset StrSql, objRs
        If Not objRs.EOF Then
            fec_desde = objRs!pliqdesde
            pDesde = objRs!pliqnro & " - " & objRs!pliqdesc
        Else
            Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encuentran periodo desde."
            Exit Sub
        End If
        objRs.Close
        
        'Busco la fecha del periodo hasta
        StrSql = "SELECT pliqdesde, pliqdesc, pliqnro FROM periodo WHERE periodo.pliqnro = " & pliqhasta
        OpenRecordset StrSql, objRs
        If Not objRs.EOF Then
            fec_hasta = objRs!pliqdesde
            pHasta = objRs!pliqnro & " - " & objRs!pliqdesc
        Else
            Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encuentran periodo hasta."
            Exit Sub
        End If
        objRs.Close
        
        'Armo lista de periodos con fecha desde entre fechas desde hasta
        If ((Not EsNulo(fec_desde)) And (Not EsNulo(fec_hasta))) Then
            listaPeriodo = ""
            StrSql = "SELECT * FROM periodo"
            StrSql = StrSql & " WHERE periodo.pliqdesde <= " & ConvFecha(fec_hasta)
            StrSql = StrSql & " AND " & ConvFecha(fec_desde) & " <= periodo.pliqdesde"
            OpenRecordset StrSql, objRs
            
            Do While Not objRs.EOF
                If Len(listaPeriodo) = 0 Then
                    listaPeriodo = objRs!pliqnro
                Else
                    listaPeriodo = listaPeriodo & "," & objRs!pliqnro
                End If
                objRs.MoveNext
            Loop
            objRs.Close
        Else
            Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encuentran las fechas de los periodos."
            Exit Sub
        End If
        
        
        Flog.writeline Espacios(Tabulador * 0) & "Buscando los procesos de los periodos."
        If Len(listaPeriodo) <> 0 Then
            StrSql = "SELECT * FROM Proceso"
            StrSql = StrSql & " WHERE Proceso.pliqnro IN (" & listaPeriodo & ")"
            StrSql = StrSql & " AND Proceso.empnro = " & EmpresaEmpnro
            OpenRecordset StrSql, objRs
            
            Do While Not objRs.EOF
                If Len(listaProceso) = 0 Then
                    listaProceso = objRs!pronro
                Else
                    listaProceso = listaProceso & "," & objRs!pronro
                End If
                objRs.MoveNext
            Loop
            objRs.Close
        
        Else
            Flog.writeline Espacios(Tabulador * 1) & "ERROR. Lista de periodos vacia."
            Exit Sub
        End If
        
        
        If Len(listaProceso) = 0 Then
            Flog.writeline Espacios(Tabulador * 1) & "ERROR. Lista de proceso vacia."
            listaProceso = "0"
        End If
        
        StrSql2 = " embcuota.pronro IN ( " & listaProceso & ")"
        '------------------------------------------------------------------------------------------------------------------------
        'GENERO LA StrSql QUE BUSCA LOS EMPLEADOS
        '------------------------------------------------------------------------------------------------------------------------

        If tenro3 <> 0 Then ' esto ocurre solo cuando se seleccionan los tres niveles
                    StrSql = " SELECT DISTINCT tipoemb.*, embargo.*,empleado.ternro,empleg, terape, terape2, ternom, ternom2 "
                    StrSql = StrSql & ", estact1.tenro tenro1, estact1.estrnro estrnro1 "
                    StrSql = StrSql & ", estact2.tenro tenro2, estact2.estrnro estrnro2 "
                    StrSql = StrSql & ", estact3.tenro tenro3, estact3.estrnro estrnro3 "
                    StrSql = StrSql & " FROM empleado  "
                    StrSql = StrSql & " INNER JOIN embargo ON empleado.ternro = embargo.ternro "
                    StrSql = StrSql & " INNER JOIN embcuota ON embargo.embnro = embcuota.embnro "
                    StrSql = StrSql & " AND embcuota.embccancela = -1 "
                    StrSql = StrSql & " INNER JOIN tipoemb ON tipoemb.tpenro = embargo.tpenro "
                    StrSql = StrSql & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro  AND estact1.tenro  = " & tenro1
                    StrSql = StrSql & " AND (estact1.htetdesde<=" & ConvFecha(fecestr) & " AND (estact1.htethasta is null or estact1.htethasta>=" & ConvFecha(fecestr) & "))"
                    If estrnro1 <> 0 Then 'cuando se le asigna un valor al nivel 1
                            StrSql = StrSql & " AND estact1.estrnro =" & estrnro1
                    End If
                    StrSql = StrSql & " INNER JOIN his_estructura estact2 ON empleado.ternro = estact2.ternro  AND estact2.tenro  = " & tenro2
                    StrSql = StrSql & " AND (estact2.htetdesde<=" & ConvFecha(fecestr) & " AND (estact2.htethasta is null or estact2.htethasta>=" & ConvFecha(fecestr) & "))"
                    If estrnro2 <> 0 Then 'cuando se le asigna un valor al nivel 2
                        StrSql = StrSql & " AND estact2.estrnro =" & estrnro2
                    End If
                    StrSql = StrSql & " INNER JOIN his_estructura estact3 ON empleado.ternro = estact3.ternro  AND estact3.tenro  = " & tenro3
                    StrSql = StrSql & " AND (estact3.htetdesde<=" & ConvFecha(fecestr) & " AND (estact3.htethasta is null or estact3.htethasta>=" & ConvFecha(fecestr) & "))"
                    If estrnro3 <> 0 Then ' cuando se le asigna un valor al nivel 3
                            StrSql = StrSql & " AND estact3.estrnro =" & estrnro3
                    End If
                    StrSql = StrSql & " WHERE " & filtro
                    StrSql = StrSql & " AND (" & StrSql2 & ") "
                    StrSql = StrSql & " ORDER BY tenro1,estrnro1,tenro2,estrnro2,tenro3,estrnro3," & Orden

        ElseIf tenro2 <> 0 Then  'ocurre cuando se selecciono hasta el segundo nivel
                    StrSql = "SELECT DISTINCT tipoemb.*, embargo.*,empleado.ternro,empleg, terape, terape2, ternom, ternom2"
                    StrSql = StrSql & ", estact1.tenro tenro1, estact1.estrnro estrnro1 "
                    StrSql = StrSql & ", estact2.tenro tenro2, estact2.estrnro estrnro2 "
                    StrSql = StrSql & " FROM empleado  "
                    StrSql = StrSql & " INNER JOIN embargo ON empleado.ternro = embargo.ternro "
                    StrSql = StrSql & " INNER JOIN embcuota ON embargo.embnro = embcuota.embnro "
                    StrSql = StrSql & " AND embcuota.embccancela = -1 "
                    StrSql = StrSql & " INNER JOIN tipoemb ON tipoemb.tpenro = embargo.tpenro "
                    StrSql = StrSql & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro  AND estact1.tenro  = " & tenro1
                    StrSql = StrSql & " AND (estact1.htetdesde<=" & ConvFecha(fecestr) & " AND (estact1.htethasta is null or estact1.htethasta>=" & ConvFecha(fecestr) & "))"
                    If estrnro1 <> 0 Then
                        StrSql = StrSql & " AND estact1.estrnro =" & estrnro1
                    End If
                    StrSql = StrSql & " INNER JOIN his_estructura estact2 ON empleado.ternro = estact2.ternro  AND estact2.tenro  = " & tenro2
                    StrSql = StrSql & " AND (estact2.htetdesde<=" & ConvFecha(fecestr) & " AND (estact2.htethasta is null or estact2.htethasta>=" & ConvFecha(fecestr) & "))"
                    If estrnro2 <> 0 Then
                        StrSql = StrSql & " AND estact2.estrnro =" & estrnro2
                    End If
                    StrSql = StrSql & " WHERE " & filtro
                    StrSql = StrSql & " AND (" & StrSql2 & ") "
                    StrSql = StrSql & " ORDER BY tenro1,estrnro1,tenro2,estrnro2," & Orden
           
        ElseIf tenro1 <> 0 Then  ' Cuando solo selecionamos el primer nivel
                    StrSql = "SELECT DISTINCT tipoemb.*, embargo.*,empleado.ternro,empleg, terape, terape2, ternom, ternom2 "
                    StrSql = StrSql & ", estact1.tenro tenro1, estact1.estrnro estrnro1 "
                    StrSql = StrSql & " FROM empleado  "
                    StrSql = StrSql & " INNER JOIN embargo ON empleado.ternro = embargo.ternro "
                    StrSql = StrSql & " INNER JOIN embcuota ON embargo.embnro = embcuota.embnro "
                    StrSql = StrSql & " AND embcuota.embccancela = -1 "
                    StrSql = StrSql & " INNER JOIN tipoemb ON tipoemb.tpenro = embargo.tpenro "
                    StrSql = StrSql & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro  AND estact1.tenro  = " & tenro1
                    StrSql = StrSql & " AND (estact1.htetdesde<=" & ConvFecha(fecestr) & " AND (estact1.htethasta is null or estact1.htethasta>=" & ConvFecha(fecestr) & "))"
                    If estrnro1 <> 0 Then
                        StrSql = StrSql & " AND estact1.estrnro =" & estrnro1
                    End If
                    StrSql = StrSql & " WHERE " & filtro
                    StrSql = StrSql & " AND (" & StrSql2 & ") "
                    StrSql = StrSql & " ORDER BY tenro1,estrnro1," & Orden
        
        Else  ' cuando no hay nivel de estructura seleccionado
                    StrSql = " SELECT DISTINCT tipoemb.*, embargo.*,empleado.ternro,empleg, terape, terape2, ternom, ternom2 "
                    StrSql = StrSql & " FROM empleado  "
                    StrSql = StrSql & " INNER JOIN embargo ON empleado.ternro = embargo.ternro "
                    StrSql = StrSql & " INNER JOIN embcuota ON embargo.embnro = embcuota.embnro "
                    StrSql = StrSql & " AND embcuota.embccancela = -1 "
                    StrSql = StrSql & " INNER JOIN tipoemb ON tipoemb.tpenro = embargo.tpenro "
                    StrSql = StrSql & " WHERE " & filtro
                    StrSql = StrSql & " AND (" & StrSql2 & ") "
                    StrSql = StrSql & " ORDER BY " & Orden
        End If
                      
       'Busco el periodo desde
       OpenRecordset StrSql, objRs
        
       If objRs.EOF Then
          Flog.writeline "No se encontraron embargos para el Reporte."
          'Exit Sub

       Else
                If objRs.RecordCount <> 0 Then
                    Cantidad = objRs.RecordCount
                    cantRegistros = Cantidad
                Else
                    Cantidad = 1
                    cantRegistros = 0
                End If
                
                IncPorc = 99 / Cantidad
                cantidadProcesada = Cantidad
               
                ' Genero los datos
                Do Until objRs.EOF
        
                    EmpErrores = False
                    embnro = objRs!embnro
                    embdesext = objRs!embdesext
                    
                    ' Genero los datos del embargo
                    Flog.writeline "Generando datos del embargo " & embnro
                    tpedesc = IIf(Not EsNulo(objRs!tpedesabr), CStr(objRs!tpedesabr), "")
                    '---------------------------------------------------
                    ' Descripcion del juzgado
                    StrSql = " SELECT juzdesabr FROM juzgado "
                    StrSql = StrSql & " WHERE juznro = " & IIf(Not EsNulo(objRs!embjuz), objRs!embjuz, 0)
                    OpenRecordset StrSql, rsJuzgado
                    If Not rsJuzgado.EOF Then
                        juzdesabr = IIf(Not EsNulo(rsJuzgado!juzdesabr), rsJuzgado!juzdesabr, "")
                    Else
                        juzdesabr = ""
                        Flog.writeline "Juzgado Nulo " & embnro
                    End If
                    rsJuzgado.Close
                    
                    '---------------------------------------------------
                    ' Descripcion de la secretaria
                    StrSql = " SELECT secdesabr FROM secretaria "
                    StrSql = StrSql & " WHERE secnro = " & IIf(Not EsNulo(objRs!embsec), objRs!embsec, 0)
                    OpenRecordset StrSql, rsSecretaria
                    If Not rsSecretaria.EOF Then
                        secdesabr = IIf(Not EsNulo(rsSecretaria!secdesabr), rsSecretaria!secdesabr, "")
                    Else
                        secdesabr = ""
                    End If
                    rsSecretaria.Close
                    
                    '---------------------------------------------------
                    ' Descripcion del banco
                    StrSql = "SELECT bandesc, bansucdesc FROM banco "
                    StrSql = StrSql & "WHERE estrnro = " & IIf(Not EsNulo(objRs!benbanco), objRs!benbanco, 0)
                    OpenRecordset StrSql, rsBanco
                    If Not rsBanco.EOF Then
                        banco = IIf(Not EsNulo(rsBanco!bandesc), Left(rsBanco!bandesc, 40), "")
                        banco = banco & " - " & IIf(Not EsNulo(rsBanco!bansucdesc), rsBanco!bansucdesc, "")
                    Else
                        banco = ""
                    End If
                    rsBanco.Close
                    
                    '---------------------------------------------------
                    ' Datos del empleado
                    empleg = CLng(objRs!empleg)
                    terape = CStr(objRs!terape)
                    terape2 = IIf(Not EsNulo(objRs!terape2), objRs!terape2, "")
                    ternom = CStr(objRs!ternom)
                    ternom2 = IIf(Not EsNulo(objRs!ternom2), objRs!ternom2, "")
                    If tenro1 <> 0 Then
                        estrnro1 = objRs!estrnro1
                    End If
                    If tenro2 <> 0 Then
                        estrnro2 = objRs!estrnro2
                    End If
                    If tenro3 <> 0 Then
                        estrnro3 = objRs!estrnro3
                    End If
                
                
                    
                    '---------------------------------------------------
                    ' Busco las cuotas asociadas al Embargo
                    StrSql = " SELECT embcuota.*, profecpago FROM embcuota "
                    StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = embcuota.pronro "
                    StrSql = StrSql & " AND embcuota.embccancela = -1 "
                    StrSql = StrSql & " WHERE embnro = " & embnro
                    StrSql = StrSql & " AND (" & StrSql2 & ") "
                    OpenRecordset StrSql, rsCuotas
                
                    Do While Not rsCuotas.EOF
                        ' Fecha del proceso asociado a la cuota
                        profecpago = IIf(Not EsNulo(rsCuotas!profecpago), rsCuotas!profecpago, "")
                        
                        embcnro = rsCuotas!embcnro
                        
                        importe = IIf(Not EsNulo(rsCuotas!embcimpreal), rsCuotas!embcimpreal, 0)
                        importe = FormatNumber(CDbl(importe), 2)
                        
                        
                        Flog.writeline "Insertando datos en la tabla "
                    
                        ' Inserto los datos del detalle en la tabla
                        Call InsertarDatosDet
                                        
                        rsCuotas.MoveNext
                    Loop
                    rsCuotas.Close
                    
                   'Actualizo el progreso
                    TiempoAcumulado = GetTickCount
                    Progreso = Progreso + IncPorc
                    cantidadProcesada = cantidadProcesada - 1
                    
                    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & numberForSQL(Progreso)
                    StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'"
                    StrSql = StrSql & ", bprcempleados ='" & CStr(cantidadProcesada) & "' WHERE bpronro = " & NroProceso
                    objConn.Execute StrSql, , adExecuteNoRecords
                    
                    objRs.MoveNext
               Loop
           
       End If
    
    Else

       Exit Sub

    End If
    
    ' Insertar Datos Comunes de los embargos
    Call InsertarDatos(cantRegistros)

    'Actualizo el estado del proceso
    If Not HuboErrores Then
       StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
       Flog.writeline "Proceso Finalizado Correctamente"
    Else
       StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProceso
       Flog.writeline "Proceso Incompleto"
    End If
    
    objConn.Execute StrSql, , adExecuteNoRecords

    Flog.writeline "Fin :" & Now
    Flog.Close
    If objRs.State = adStateOpen Then objRs.Close

    Exit Sub
    
CE:
    HuboErrores = True
    Flog.writeline " Error: " & Err.Description & Now

End Sub

'------------------------------------------------------------------------------------
' Se encarga de Insertar los datos comunes de la consulta en la tabla de Resultados
'------------------------------------------------------------------------------------
Sub InsertarDatos(ByVal Cantidad As Integer)

    Dim StrSql As String
    
    On Error GoTo MError
    
    '-------------------------------------------------------------------------------
    'Inserto los datos en la BD
    '-------------------------------------------------------------------------------
    StrSql = "INSERT INTO rep_det_cuota_emb (bpronro,pdesde,phasta,cant,titrep,fecharep,horarep,iduser) "
    StrSql = StrSql & "VALUES (" & _
             NroProceso & ",'" & pDesde & "','" & pHasta & "'," & Cantidad & _
             ",'" & titulofiltro & "'," & ConvFecha(Fecha) & "," & "'" & Hora & "','" & IdUser & "')"
    objConn.Execute StrSql, , adExecuteNoRecords
        
    Exit Sub
                
MError:
    Flog.writeline " Error: " & Err.Description
    HuboErrores = True
    EmpErrores = True
    Exit Sub

End Sub

'------------------------------------------------------------------------------------
' Se encarga de Insertar los datos de la consulta en la tabla de Resultados
'------------------------------------------------------------------------------------
Sub InsertarDatosDet()

    Dim StrSql As String
    
    On Error GoTo MError2
    
    '-------------------------------------------------------------------------------
    'Inserto los datos en la BD
    '-------------------------------------------------------------------------------
    StrSql = "INSERT INTO rep_det_cuota_emb_det " & _
             "(bpronro,empleg,terape,ternom2,terape2,ternom,tenro1,estrnro1,tenro2,estrnro2,tenro3," & _
             "estrnro3,embnro,embdesext,tpedesabr,embcnro,juzdesabr,secdesabr,banco,importe,fecha)"
    If profecpago = "" Then
        profecpago = "NULL"
    Else
        profecpago = ConvFecha(profecpago)
    End If
    StrSql = StrSql & "VALUES (" & _
             NroProceso & "," & empleg & ",'" & terape & "','" & ternom2 & "','" & terape2 & "','" & ternom & "'," & _
             tenro1 & "," & estrnro1 & "," & tenro2 & "," & estrnro2 & "," & tenro3 & "," & _
             estrnro3 & "," & embnro & ",'" & embdesext & "','" & tpedesc & "'," & embcnro & ",'" & _
             juzdesabr & "','" & secdesabr & "','" & banco & "'," & importe & "," & profecpago & ")"
             
    objConn.Execute StrSql, , adExecuteNoRecords
        
    Exit Sub
                
MError2:
    Flog.writeline " Error: " & Err.Description
    HuboErrores = True
    EmpErrores = True
    Exit Sub

End Sub

Function numberForSQL(Str)
   
  numberForSQL = Replace(Str, ",", ".")

End Function

