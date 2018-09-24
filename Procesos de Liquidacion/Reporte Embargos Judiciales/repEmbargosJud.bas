Attribute VB_Name = "repEmbargosJud"
Option Explicit

'Global Const Version = "1.01"
'Global Const FechaModificacion = "10/03/2007" ' Gustavo Ring
'Global Const UltimaModificacion = " " 'Sacar las vistas - agregar versión y comentarios

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
Global orden As String
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
Global EmpEstrnro As Long
Global EmpTernro As Long
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

    Nombre_Arch = PathFLog & "ReporteEmbargos" & "-" & NroProceso & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    Flog.writeline "_________________________________________________________________"
    Flog.writeline "Version = " & Version
    Flog.writeline "Modificacion = " & UltimaModificacion
    Flog.writeline "Fecha = " & FechaModificacion
    Flog.writeline "_________________________________________________________________"
    Flog.writeline
       
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "PID = " & PID
    
    TiempoInicialProceso = GetTickCount
    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    On Error GoTo CE
    HuboErrores = False
    
    Flog.writeline "Inicio Proceso de Reporte de Embargos Judiciales: " & Now
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
       
       'Obtengo los parametros del proceso
       parametros = objRs2!bprcparam
       ArrParametros = Split(parametros, "@")
              
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
        orden = ArrParametros(9)

        ' Fecha inicial del periodo en el cual se deben buscar los embargos
        fec_desde = ArrParametros(10)
        
        ' Fecha final del periodo en el cual se deben buscar los embargos
        fec_hasta = ArrParametros(11)
       

   '______________________________________________________
        Flog.writeline " Datos del Filtro: "
        Flog.writeline "    Filtro: " & ArrParametros(1)
        Flog.writeline "    Fecha Estr: " & ArrParametros(2)
        Flog.writeline "    Estructuras: " & ArrParametros(3) & " - " & ArrParametros(4) & " - " & ArrParametros(5) & " - " & ArrParametros(6) & " - " & ArrParametros(7) & " - " & ArrParametros(8)
        Flog.writeline "    Orden: " & ArrParametros(9)
        Flog.writeline "    Período de Emb.: " & ArrParametros(10) & " - " & ArrParametros(11)
        Flog.writeline " "
             
        'EMPIEZA EL PROCESO

        '------------------------------------------------------------------------------------------------------------------------
        'GENERO LA SQL QUE BUSCA EL PERIODOS
        '------------------------------------------------------------------------------------------------------------------------

        fecAuxHasta = Split(fec_hasta, "/", -1, 1)
        fecAuxDesde = Split(fec_desde, "/", -1, 1)

        ' Controlo que el periodo tenga cuotas generadas en el rango dado
        StrSql2 = " ( ( (embcuota.embcanio > " & Int(fecAuxDesde(2)) & ") AND (embcuota.embcanio < " & Int(fecAuxHasta(2)) & ") ) "

        StrSql2 = StrSql2 & " OR "

        StrSql2 = StrSql2 & " ( (embcuota.embcanio = " & Int(fecAuxDesde(2)) & ") AND (embcuota.embcanio < " & Int(fecAuxHasta(2)) & ") AND (embcuota.embcmes >= " & Int(fecAuxDesde(1)) & ") ) "
        
        StrSql2 = StrSql2 & " OR "
        
        StrSql2 = StrSql2 & " ( (embcuota.embcanio > " & Int(fecAuxDesde(2)) & ") AND (embcuota.embcanio = " & Int(fecAuxHasta(2)) & ") AND (embcuota.embcmes <= " & Int(fecAuxHasta(1)) & ") ) "
        
        StrSql2 = StrSql2 & " OR "
        
        StrSql2 = StrSql2 & " ( (embcuota.embcanio = " & Int(fecAuxDesde(2)) & ") AND (embcuota.embcanio = " & Int(fecAuxHasta(2)) & ") AND (embcuota.embcmes >= " & Int(fecAuxDesde(1)) & ") AND (embcuota.embcmes <= " & Int(fecAuxHasta(1)) & ") ) )"
        
        
        ' ____________________________________________________________
        Flog.writeline "  SQL para control del periodo de las cuotas. "
        Flog.writeline "    " & StrSql2
        Flog.writeline " "
'
       '------------------------------------------------------------------------------------------------------------------------
        'GENERO LA StrSql QUE BUSCA LOS EMPLEADOS
        '------------------------------------------------------------------------------------------------------------------------

        If tenro3 <> 0 Then ' esto ocurre solo cuando se seleccionan los tres niveles
                    StrSql = " SELECT DISTINCT tipoemb.*, embargo.*,empleado.ternro,empleg, terape, terape2, ternom, ternom2 "
                    StrSql = StrSql & ", estact1.tenro AS tenro1, estact1.estrnro AS estrnro1 "
                    StrSql = StrSql & ", estact2.tenro AS tenro2, estact2.estrnro AS estrnro2 "
                    StrSql = StrSql & ", estact3.tenro AS tenro3, estact3.estrnro AS estrnro3 "
                    StrSql = StrSql & " FROM empleado  "
                    StrSql = StrSql & " INNER JOIN embargo ON empleado.ternro = embargo.ternro "
                    StrSql = StrSql & " INNER JOIN embcuota ON embargo.embnro = embcuota.embnro "
                    StrSql = StrSql & " AND embcuota.pronro is not null "
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
                    StrSql = StrSql & " ORDER BY tenro1,estrnro1,tenro2,estrnro2,tenro3,estrnro3," & orden

        ElseIf tenro2 <> 0 Then  'ocurre cuando se selecciono hasta el segundo nivel
                    StrSql = "SELECT DISTINCT tipoemb.*, embargo.*,empleado.ternro,empleg, terape, terape2, ternom, ternom2"
                    StrSql = StrSql & ", estact1.tenro AS tenro1, estact1.estrnro AS estrnro1 "
                    StrSql = StrSql & ", estact2.tenro AS tenro2, estact2.estrnro AS estrnro2 "
                    StrSql = StrSql & " FROM empleado  "
                    StrSql = StrSql & " INNER JOIN embargo ON empleado.ternro = embargo.ternro "
                    StrSql = StrSql & " INNER JOIN embcuota ON embargo.embnro = embcuota.embnro "
                    StrSql = StrSql & " AND embcuota.pronro is not null "
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
                    StrSql = StrSql & " ORDER BY tenro1,estrnro1,tenro2,estrnro2," & orden
           
        ElseIf tenro1 <> 0 Then  ' Cuando solo selecionamos el primer nivel
                    StrSql = "SELECT DISTINCT tipoemb.*, embargo.*,empleado.ternro,empleg, terape, terape2, ternom, ternom2 "
                    StrSql = StrSql & ", estact1.tenro AS tenro1, estact1.estrnro AS estrnro1 "
                    StrSql = StrSql & " FROM empleado  "
                    StrSql = StrSql & " INNER JOIN embargo ON empleado.ternro = embargo.ternro "
                    StrSql = StrSql & " INNER JOIN embcuota ON embargo.embnro = embcuota.embnro "
                    StrSql = StrSql & " AND embcuota.pronro is not null "
                    StrSql = StrSql & " INNER JOIN tipoemb ON tipoemb.tpenro = embargo.tpenro "
                    StrSql = StrSql & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro  AND estact1.tenro  = " & tenro1
                    StrSql = StrSql & " AND (estact1.htetdesde<=" & ConvFecha(fecestr) & " AND (estact1.htethasta is null or estact1.htethasta>=" & ConvFecha(fecestr) & "))"
                    If estrnro1 <> 0 Then
                        StrSql = StrSql & " AND estact1.estrnro =" & estrnro1
                    End If
                    StrSql = StrSql & " WHERE " & filtro
                    StrSql = StrSql & " AND (" & StrSql2 & ") "
                    StrSql = StrSql & " ORDER BY tenro1,estrnro1," & orden
        
        Else  ' cuando no hay nivel de estructura seleccionado
                    StrSql = " SELECT DISTINCT tipoemb.*, embargo.*,empleado.ternro,empleg, terape, terape2, ternom, ternom2 "
                    StrSql = StrSql & " FROM empleado  "
                    StrSql = StrSql & " INNER JOIN embargo ON empleado.ternro = embargo.ternro "
                    StrSql = StrSql & " INNER JOIN embcuota ON embargo.embnro = embcuota.embnro "
                    StrSql = StrSql & " AND embcuota.pronro is not null "
                    StrSql = StrSql & " INNER JOIN tipoemb ON tipoemb.tpenro = embargo.tpenro "
                    StrSql = StrSql & " WHERE " & filtro
                    StrSql = StrSql & " AND (" & StrSql2 & ") "
                    StrSql = StrSql & " ORDER BY " & orden
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
                    ' Busco los datos de la empresa
'                    StrSql = "SELECT his_estructura.estrnro, empresa.ternro, empresa.empnom " & _
'                        " From his_estructura" & _
'                        " INNER JOIN empresa ON empresa.estrnro = his_estructura.estrnro" & _
'                        " WHERE his_estructura.htethasta IS NULL" & _
'                        " AND his_estructura.ternro = " & objRs!ternro & _
'                        " AND his_estructura.tenro  = 10"
'                        " WHERE his_estructura.htetdesde <=" & ConvFecha(fec_hasta) & " AND " & _
'                        " (his_estructura.htethasta >= " & ConvFecha(fec_hasta) & " OR his_estructura.htethasta IS NULL)"
'                    OpenRecordset StrSql, rs_estructura
                    
                    EmpEstrnro = 0
                    EmpNombre = ""
                    EmpTernro = 0
                    EmpDire = "   "
                    EmpLogo = ""
                    EmpLogoAlto = 0
                    EmpLogoAncho = 0

                    
'                    If rs_estructura.EOF Then
'                        Flog.Writeline "No se encontró la empresa"
                        'Exit Sub
'                    Else
'                        EmpNombre = rs_estructura!empnom
'                        EmpEstrnro = rs_estructura!estrnro
'                        EmpTernro = rs_estructura!ternro
'                    End If
'                    rs_estructura.Close
                    
                    'Consulta para obtener la direccion de la empresa
'                    StrSql = "SELECT detdom.calle,detdom.nro,localidad.locdesc, codigopostal From cabdom " & _
'                        " INNER JOIN detdom ON detdom.domnro = cabdom.domnro AND cabdom.ternro =" & EmpTernro & _
'                        " INNER JOIN localidad ON detdom.locnro = localidad.locnro "
'                    OpenRecordset StrSql, rs_Domicilio
'                    If rs_Domicilio.EOF Then
'                        Flog.Writeline "No se encontró el domicilio de la empresa"
                        'Exit Sub
'                        EmpDire = "   "
'                    Else
'                        EmpDire = IIf(Not EsNulo(rs_Domicilio!calle), rs_Domicilio!calle, "") & " " & IIf(Not EsNulo(rs_Domicilio!Nro), rs_Domicilio!Nro, "") & "<br>" & IIf(Not EsNulo(rs_Domicilio!codigopostal), rs_Domicilio!codigopostal, "") & " " & IIf(Not EsNulo(rs_Domicilio!locdesc), rs_Domicilio!locdesc, "")
'                    End If
'                    rs_Domicilio.Close
                    
                    'Consulta para buscar el logo de la empresa
'                    StrSql = "SELECT ter_imag.terimnombre, tipoimag.tipimdire, tipoimag.tipimanchodef, tipoimag.tipimaltodef" & _
'                        " From ter_imag " & _
'                        " INNER JOIN tipoimag ON tipoimag.tipimnro = 1 AND tipoimag.tipimnro = ter_imag.tipimnro" & _
'                        " AND ter_imag.ternro = " & EmpTernro
'                    OpenRecordset StrSql, rs_logo
'                    If rs_logo.EOF Then
'                        Flog.Writeline "No se encontró el Logo de la Empresa"
                        'Exit Sub
'                        EmpLogo = ""
'                        EmpLogoAlto = 0
'                        EmpLogoAncho = 0
'                    Else
'                        EmpLogo = rs_logo!tipimdire & rs_logo!terimnombre
'                        EmpLogoAlto = rs_logo!tipimaltodef
'                        EmpLogoAncho = rs_logo!tipimanchodef
'                    End If
'                    rs_logo.Close
                    
                    '---------------------------------------------------
                    ' Busco las cuotas asociadas al Embargo
                    StrSql = " SELECT embcuota.*, profecpago FROM embcuota "
                    StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = embcuota.pronro "
                    StrSql = StrSql & " AND embcuota.pronro is not null  "
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
    StrSql = "INSERT INTO rep_emb_jud (bpronro,fecdesde,fechasta,cant,titrep,fecharep,horarep,iduser) "
    StrSql = StrSql & "VALUES (" & _
             NroProceso & "," & ConvFecha(fec_desde) & "," & ConvFecha(fec_hasta) & "," & Cantidad & _
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
    StrSql = "INSERT INTO rep_emb_jud_det " & _
             "(bpronro,empleg,terape,ternom2,terape2,ternom,tenro1,estrnro1,tenro2,estrnro2,tenro3," & _
             "estrnro3,embnro,embdesext,tpedesabr,embcnro,juzdesabr,secdesabr,banco,importe,fecha," & _
             "empnombre,empdire,emplogo,emplogoalto,emplogoancho)"
    If profecpago = "" Then
        profecpago = "NULL"
    Else
        profecpago = ConvFecha(profecpago)
    End If
    StrSql = StrSql & "VALUES (" & _
             NroProceso & "," & empleg & ",'" & terape & "','" & ternom2 & "','" & terape2 & "','" & ternom & "'," & _
             tenro1 & "," & estrnro1 & "," & tenro2 & "," & estrnro2 & "," & tenro3 & "," & _
             estrnro3 & "," & embnro & ",'" & embdesext & "','" & tpedesc & "'," & embcnro & ",'" & _
             juzdesabr & "','" & secdesabr & "','" & banco & "'," & importe & "," & profecpago & ",'" & _
             EmpNombre & "','" & EmpDire & "','" & EmpLogo & "'," & EmpLogoAlto & "," & EmpLogoAncho & ")"
             
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

