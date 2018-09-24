Attribute VB_Name = "MdlRepDeclaracionJuradaAnses"
Option Explicit

'Version: 0.01
'Primera Version realizada, todavia en etapa de desarrollo
Const Version = 1.1
Const FechaVersion = "05/05/2006"
Const tiprep_nro = 165

'------------------------------------------------------------
'------------------------------------------------------------
'------------------------------------------------------------
Global IdUser As String
Global Fecha As Date
Global Hora As String

Global Aux_Autoriz_Apenom As String
Global Aux_Autoriz_Docu As String
Global Aux_Autoriz_Prov_Emis As String

Global Aux_Certifi_Corresponde As String
Global Aux_Certifi_Doc_Tipo As String
Global Aux_Certifi_Doc_Nro As String
Global Aux_Certifi_Expedida As String



Public Sub Main()
    ' ---------------------------------------------------------------------------------------------
    ' Descripcion: Procedimiento inicial del Generador de Reportes.
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
    'strCmdLine = "10310"
    ArrParametros = Split(strCmdLine, " ", -1)
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
    
    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas

    'Abro la conexion
    OpenConnection strconexion, objConn
    
    Nombre_Arch = PathFLog & "DeclaracionJuradaAnses_PS53" & "-" & NroProcesoBatch & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.Writeline "-----------------------------------------------------------------"
    Flog.Writeline "Version = " & Version
    Flog.Writeline "Fecha   = " & FechaVersion
    Flog.Writeline "-----------------------------------------------------------------"
    Flog.Writeline
    Flog.Writeline "PID = " & PID
    
    
    'Cambio el estado del proceso a Procesando y el PID
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objConn.Execute StrSql, , adExecuteNoRecords
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE (btprcnro = 129 ) AND bpronro =" & NroProcesoBatch
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
    Flog.Writeline "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    
    If Not HuboError Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
    End If
    objConn.Execute StrSql, , adExecuteNoRecords

    Flog.Close
    objConn.Close
    
End Sub

Public Sub Generar_Reporte(ByVal bpronro As Long, ByVal HFecha As Date)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del Reporte de Certificado Anses de Servicios
' Autor      : FGZ
' Fecha      : 15/07/2004
' Ult. Mod   : FGZ - 10/11/2005
' Desc       : agrupa en hojas de hasta 5 movimientos de empresas por c/u.
' --------------------------------------------------------------------------------------------
'Variables auxiliares

'Empresa
Dim Aux_Empresa_ternro As String
Dim Aux_Empresa_Cuit As String
Dim Aux_Empresa_RazonSocial As String
Dim Aux_Empresa_DMTRCalle As String
Dim Aux_Empresa_DMTRNum As String
Dim Aux_Empresa_DMTRPiso As String
Dim Aux_Empresa_DMTRDepto As String
Dim Aux_Empresa_DMTRCodPost As String
Dim Aux_Empresa_DMTRLoc As String
Dim Aux_Empresa_DMTRProv As String
Dim Aux_Empresa_TelefonoLab As String
Dim Aux_Empresa_provcod

'Inicializo variables de Empresa
Aux_Empresa_Cuit = " "
Aux_Empresa_RazonSocial = " "
Aux_Empresa_DMTRCalle = " "
Aux_Empresa_DMTRNum = " "
Aux_Empresa_DMTRPiso = " "
Aux_Empresa_DMTRDepto = " "
Aux_Empresa_DMTRCodPost = " "
Aux_Empresa_DMTRLoc = " "
Aux_Empresa_DMTRProv = " "
Aux_Empresa_TelefonoLab = " "


'Empleado
Dim Aux_Empleado_ternro As String
Dim Aux_Empleado_CUIL As String
Dim Aux_Empleado_Nombres As String
Dim Aux_Empleado_TipoNroDoc As String
Dim Aux_Empleado_Sexo As String
Dim Aux_Empleado_FecNac As String
Dim Aux_Empleado_EstadoCivil As String
Dim Aux_Empleado_DOMCalle As String
Dim Aux_Empleado_DOMNro As String
Dim Aux_Empleado_DOMPiso As String
Dim Aux_Empleado_DOMDto As String
Dim Aux_Empleado_DOMCodPost As String
Dim Aux_Empleado_DOMLocalidad As String
Dim Aux_Empleado_DOMProvincia As String
Dim Aux_Empleado_Telefono As String
Dim Aux_Empleado_OSElegida As String
Dim Aux_Empleado_FecIngreso As String
Dim Aux_Empleado_ABM As String

'Empleado
Dim Aux_Empleado_codEstCiv
Dim Aux_Empleado_codProv
Dim Aux_Empleado_nombre
Dim Aux_Empleado_apellido
Dim Aux_Empleado_nrodoc
Dim Aux_Empleado_tipodoc
Dim Aux_Empleado_nacionalidad
Dim Aux_Empleado_codigo
Dim Aux_Empleado_sitRev

'Inicializo variables de Empleado
Aux_Empleado_CUIL = " "
Aux_Empleado_Nombres = " "
Aux_Empleado_TipoNroDoc = " "
Aux_Empleado_FecNac = ""
Aux_Empleado_EstadoCivil = " "
Aux_Empleado_DOMCalle = " "
Aux_Empleado_DOMNro = " "
Aux_Empleado_DOMPiso = " "
Aux_Empleado_DOMDto = " "
Aux_Empleado_codigo = " "
Aux_Empleado_DOMLocalidad = " "
Aux_Empleado_DOMProvincia = " "
Aux_Empleado_Telefono = " "
Aux_Empleado_OSElegida = " "
Aux_Empleado_FecIngreso = ""
Aux_Empleado_ABM = " "

'Familiar
Dim Aux_Familiar_Parentesco As String
Dim Aux_Familiar_codPar As String
Dim Aux_Familiar_CUIL As String
Dim Aux_Familiar_Nombres As String
Dim Aux_Familiar_TipoNroDoc As String
Dim Aux_Familiar_FecNac As String
Dim Aux_Familiar_Sexo As String
Dim Aux_Familiar_EstadoCivil As String
Dim Aux_Familiar_Incapacidad As String
Dim Aux_Familiar_DOMCalle As String
Dim Aux_Familiar_DOMNro As String
Dim Aux_Familiar_DOMPiso As String
Dim Aux_Familiar_DOMDto As String
Dim Aux_Familiar_DOMCodPost As String
Dim Aux_Familiar_DOMLocalidad As String
Dim Aux_Familiar_DOMProvinc As String
Dim Aux_Familiar_Telefono As String
Dim Aux_Familiar_Aniocursa As String
Dim Aux_Familiar_Genderasigf As String ' Si/No
Dim Aux_Familiar_Escolaridad As String 'General basica, polimodal o no informa
Dim confval_1 As Integer
Dim confval_2 As Integer
Dim Aux_Familiar_Estudia As String


Dim Aux_Familiar_tipodoc As String
Dim Aux_Familiar_nrodoc As String
Dim Aux_Familiar_estcivcod As String
Dim Aux_Familiar_provcod As String
Dim Aux_Familiar_nacionalidad As String
Dim Aux_Familiar_apellido As String

'Familiar
Aux_Familiar_Parentesco = " "
Aux_Familiar_CUIL = " "
Aux_Familiar_Nombres = " "
Aux_Familiar_TipoNroDoc = " "
Aux_Familiar_FecNac = ""
Aux_Familiar_Sexo = " "
Aux_Familiar_EstadoCivil = " "
Aux_Familiar_Incapacidad = " "
Aux_Familiar_DOMCalle = " "
Aux_Familiar_DOMNro = " "
Aux_Familiar_DOMPiso = " "
Aux_Familiar_DOMDto = " "
Aux_Familiar_DOMCodPost = " "
Aux_Familiar_DOMLocalidad = " "
Aux_Familiar_DOMProvinc = " "
Aux_Familiar_Telefono = " "
Aux_Familiar_Aniocursa = " "
Aux_Familiar_Genderasigf = 0 ' Si/No
Aux_Familiar_Escolaridad = "N" 'General basica, polimodal o no informa


'Registros
Dim rs_Empresa As New ADODB.Recordset
Dim rs_situacionRevista As New ADODB.Recordset
Dim rs_Empleados As New ADODB.Recordset
Dim rs_Familiares As New ADODB.Recordset
Dim rs_batch_empleado As New ADODB.Recordset
Dim rs_repnro As New ADODB.Recordset
Dim rs_Confrep As New ADODB.Recordset
Dim rs_Estudio As New ADODB.Recordset
Dim rs_confval As New ADODB.Recordset
Dim Aux_Familiar_ternro As String
Dim Aux_Familiar_empleado As String
Dim Aux_Familiar_repnro As String
Dim Aux_repnro As String
Dim queryEmpleado As String
Dim sqlRev As String

Dim Tercero

On Error GoTo CE

' Comienzo la transaccion
    MyBeginTrans
    StrSql = "SELECT batch_empleado.ternro,empleado.empleg from batch_empleado inner join empleado on empleado.ternro = batch_empleado.ternro where batch_empleado.bpronro = " & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_empleado
    If rs_batch_empleado.EOF Then
        StrSql = "SELECT ternro,empleg from empleado"
        OpenRecordset StrSql, rs_batch_empleado
    End If
    'Seteo el progreso
    Progreso = 0
    CEmpleadosAProc = rs_batch_empleado.RecordCount
    If CEmpleadosAProc = 0 Then
        Flog.Writeline "------------------------------"
        Flog.Writeline "No hay empleados que procesar "
        Flog.Writeline "------------------------------"
        CEmpleadosAProc = 1
    End If
    IncPorc = (100 / CEmpleadosAProc)
    
    'Loopeo por cada empleado en la tabla batch_empleado con el ID del proceso
    Do While Not rs_batch_empleado.EOF
        Flog.Writeline "Procesando empleado id " & rs_batch_empleado!ternro
        Flog.Writeline "------------------------------"
    
        'Situacion Revista
        
        StrSql = "Select estrcodext from his_estructura "
        StrSql = StrSql & "inner join estructura on estructura.estrnro=his_estructura.estrnro "
        StrSql = StrSql & "Where estructura.tenro = 30 And ternro=" & rs_batch_empleado!ternro
        StrSql = StrSql & " order by htetdesde desc "
        OpenRecordset StrSql, rs_situacionRevista
        
        If Not IsNull(rs_situacionRevista!estrcodext) Then Aux_Empleado_sitRev = rs_situacionRevista!estrcodext


        'Traigo loa valores de la empresa
        StrSql = "select empresa.ternro EMPTERNRO, detdom.calle, detdom.nro, detdom.piso, detdom.oficdepto, detdom.codigopostal, localidad.locdesc, provincia.provdesc, telefono.telnro, empresa.empnom, ter_doc.nrodoc, provincia.provcodext "
        StrSql = StrSql & " from his_estructura"
        StrSql = StrSql & " join empresa on his_estructura.estrnro = empresa.estrnro and his_estructura.htetdesde <= " & ConvFecha(HFecha) & " and (his_estructura.htethasta >= " & ConvFecha(HFecha) & " or his_estructura.htethasta is null) and his_estructura.tenro = 10 and his_estructura.ternro = " & rs_batch_empleado!ternro
        StrSql = StrSql & " left join cabdom on empresa.ternro = cabdom.ternro and cabdom.domdefault = -1 "
        StrSql = StrSql & " left join detdom on cabdom.domnro = detdom.domnro "
        StrSql = StrSql & " left join localidad on localidad.locnro = detdom.locnro "
        StrSql = StrSql & " left join provincia on provincia.provnro = detdom.provnro "
        StrSql = StrSql & " left join telefono on telefono.domnro = cabdom.domnro and telefono.teldefault = -1 "
        StrSql = StrSql & " left join ter_doc on ter_doc.ternro = empresa.ternro and tidnro = 6"
        OpenRecordset StrSql, rs_Empresa
        
        'Asigno los valores auxiliares para la empresa
        If Not rs_Empresa.EOF Then
            If Not IsNull(rs_Empresa!EMPTERNRO) Then Aux_Empresa_ternro = rs_Empresa!EMPTERNRO
            If Not IsNull(rs_Empresa!nrodoc) Then Aux_Empresa_Cuit = rs_Empresa!nrodoc
            If Not IsNull(rs_Empresa!empnom) Then Aux_Empresa_RazonSocial = rs_Empresa!empnom
            If Not IsNull(rs_Empresa!calle) Then Aux_Empresa_DMTRCalle = rs_Empresa!calle
            If Not IsNull(rs_Empresa!nro) Then Aux_Empresa_DMTRNum = rs_Empresa!nro
            If Not IsNull(rs_Empresa!piso) Then Aux_Empresa_DMTRPiso = rs_Empresa!piso
            If Not IsNull(rs_Empresa!oficdepto) Then Aux_Empresa_DMTRDepto = rs_Empresa!oficdepto
            If Not IsNull(rs_Empresa!codigopostal) Then Aux_Empresa_DMTRCodPost = rs_Empresa!codigopostal
            If Not IsNull(rs_Empresa!locdesc) Then Aux_Empresa_DMTRLoc = rs_Empresa!locdesc
            If Not IsNull(rs_Empresa!provdesc) Then Aux_Empresa_DMTRProv = rs_Empresa!provdesc
            If Not IsNull(rs_Empresa!telnro) Then Aux_Empresa_TelefonoLab = rs_Empresa!telnro
            If Not IsNull(rs_Empresa!provcodext) Then Aux_Empresa_provcod = rs_Empresa!provcodext
        End If
        
        'Traigo loa valores del empleado
        StrSql = "select estcivil.estcivdesabr, tercero.terape, tercero.terape2, tercero.ternom, tercero.ternom2, tipodocu.tidsigla, DU.nrodoc NRODU, tercero.tersex, tercero.terfecnac, detdom.calle, detdom.nro, detdom.piso, detdom.oficdepto, detdom.codigopostal, localidad.locdesc, provincia.provdesc, telefono.telnro, CUIL.nrodoc NROCUIL, osocial.osdesc, empleado.empfaltagr,provincia.provcodext,estcivil.estcivcodext,nacionaldes from tercero "
        StrSql = StrSql & " left join cabdom on tercero.ternro = cabdom.ternro and cabdom.domdefault = -1"
        StrSql = StrSql & " left join detdom on cabdom.domnro = detdom.domnro"
        StrSql = StrSql & " left join localidad on localidad.locnro = detdom.locnro"
        StrSql = StrSql & " left join provincia on provincia.provnro = detdom.provnro"
        
        StrSql = StrSql & " left join empleado on empleado.ternro = tercero.ternro"
        StrSql = StrSql & " left join nacionalidad on tercero.nacionalnro = nacionalidad.nacionalnro"
        
        StrSql = StrSql & " left join telefono on telefono.domnro = cabdom.domnro and telefono.teldefault = -1"
        StrSql = StrSql & " left join ter_doc CUIL on CUIL.ternro = tercero.ternro and CUIL.tidnro = 10"
        StrSql = StrSql & " left join ter_doc DU on DU.ternro = tercero.ternro and (DU.tidnro = 1 or DU.tidnro = 2 or DU.tidnro = 3 or DU.tidnro = 4)"
        StrSql = StrSql & " left join tipodocu on DU.tidnro = tipodocu.tidnro"
        StrSql = StrSql & " left join estcivil on tercero.estcivnro = estcivil.estcivnro "
        StrSql = StrSql & " left join his_estructura on his_estructura.tenro = 17 and his_estructura.ternro = tercero.ternro and his_estructura.htetdesde <= " & ConvFecha(HFecha) & " and (his_estructura.htethasta >= " & ConvFecha(HFecha) & " or his_estructura.htethasta is null)"
        StrSql = StrSql & " left join replica_estr on replica_estr.estrnro = his_estructura.estrnro"
        StrSql = StrSql & " left join osocial on osocial.ternro = replica_estr.origen"
        StrSql = StrSql & " where tercero.ternro = " & rs_batch_empleado!ternro
        OpenRecordset StrSql, rs_Empleados
        'Asigno los valores auxiliares para el empleado
        
        If rs_Empleados.EOF Then
            StrSql = "select estcivil.estcivdesabr, tercero.terape, tercero.terape2, tercero.ternom, tercero.ternom2, tipodocu.tidsigla, DU.nrodoc NRODU, tercero.tersex, tercero.terfecnac, detdom.calle, detdom.nro, detdom.piso, detdom.oficdepto, detdom.codigopostal, localidad.locdesc, provincia.provdesc, telefono.telnro, CUIL.nrodoc NROCUIL, osocial.osdesc, v_empleado.empfaltagr from tercero "
            StrSql = StrSql & " left join cabdom on tercero.ternro = cabdom.ternro and cabdom.tidonro=2"
            StrSql = StrSql & " left join detdom on cabdom.domnro = detdom.domnro"
            StrSql = StrSql & " left join localidad on localidad.locnro = detdom.locnro"
            StrSql = StrSql & " left join provincia on provincia.provnro = detdom.provnro"
            StrSql = StrSql & " left join v_empleado on V_empleado.ternro = tercero.ternro"
            StrSql = StrSql & " left join telefono on telefono.domnro = cabdom.domnro and telefono.teldefault = -1"
            StrSql = StrSql & " left join ter_doc CUIL on CUIL.ternro = tercero.ternro and CUIL.tidnro = 10"
            StrSql = StrSql & " left join ter_doc DU on DU.ternro = tercero.ternro and (DU.tidnro = 1 or DU.tidnro = 2 or DU.tidnro = 3 or DU.tidnro = 4)"
            StrSql = StrSql & " left join tipodocu on DU.tidnro = tipodocu.tidnro"
            StrSql = StrSql & " left join estcivil on tercero.estcivnro = estcivil.estcivnro "
            StrSql = StrSql & " left join his_estructura on his_estructura.tenro = 17 and his_estructura.ternro = tercero.ternro and his_estructura.htetdesde <= " & ConvFecha(HFecha) & " and (his_estructura.htethasta >= " & ConvFecha(HFecha) & " or his_estructura.htethasta is null)"
            StrSql = StrSql & " left join replica_estr on replica_estr.estrnro = his_estructura.estrnro"
            StrSql = StrSql & " left join osocial on osocial.ternro = replica_estr.origen"
            OpenRecordset StrSql, rs_Empleados
        End If
        
        If Not rs_Empleados.EOF Then
            Aux_Empleado_ternro = rs_batch_empleado!ternro
            If rs_Empleados!tersex = -1 Then
                Aux_Empleado_Sexo = "M"
            Else
                Aux_Empleado_Sexo = "F"
            End If
            
            
            If Not IsNull(rs_Empleados!nacionaldes) Then Aux_Empleado_nacionalidad = rs_Empleados!nacionaldes
            If Not IsNull(rs_Empleados!estcivcodext) Then Aux_Empleado_codEstCiv = rs_Empleados!estcivcodext
            If Not IsNull(rs_Empleados!provcodext) Then Aux_Empleado_codProv = rs_Empleados!provcodext
            
            If Not IsNull(rs_Empleados!NROCuil) Then Aux_Empleado_CUIL = rs_Empleados!NROCuil
            
            If Not IsNull(rs_Empleados!NROCuil) Then Aux_Empleado_CUIL = rs_Empleados!NROCuil
            Aux_Empleado_Nombres = rs_Empleados!terape & " " & rs_Empleados!terape2 & " " & rs_Empleados!ternom & " " & rs_Empleados!ternom2
            Aux_Empleado_Nombres = Replace(Aux_Empleado_Nombres, "'", "\'")
            
            Aux_Empleado_apellido = rs_Empleados!terape & " " & rs_Empleados!terape2
            Aux_Empleado_nombre = rs_Empleados!ternom & " " & rs_Empleados!ternom2
            
            Aux_Empleado_tipodoc = rs_Empleados!tidsigla
            Aux_Empleado_nrodoc = rs_Empleados!NRODU
            
            Aux_Empleado_TipoNroDoc = rs_Empleados!tidsigla & " " & rs_Empleados!NRODU
            If Not IsNull(rs_Empleados!terfecnac) Then Aux_Empleado_FecNac = ConvFecha(rs_Empleados!terfecnac)
            If Not IsNull(rs_Empleados!estcivdesabr) Then Aux_Empleado_EstadoCivil = rs_Empleados!estcivdesabr
            If Not IsNull(rs_Empleados!calle) Then Aux_Empleado_DOMCalle = rs_Empleados!calle
            Aux_Empleado_DOMCalle = Replace(Aux_Empleado_DOMCalle, "'", "\'")
            If Not IsNull(rs_Empleados!nro) Then Aux_Empleado_DOMNro = rs_Empleados!nro
            If Not IsNull(rs_Empleados!piso) Then Aux_Empleado_DOMPiso = rs_Empleados!piso
            If Not IsNull(rs_Empleados!oficdepto) Then Aux_Empleado_DOMDto = rs_Empleados!oficdepto
            If Not IsNull(rs_Empleados!codigopostal) Then Aux_Empleado_codigo = rs_Empleados!codigopostal
            If Not IsNull(rs_Empleados!locdesc) Then Aux_Empleado_DOMLocalidad = rs_Empleados!locdesc
            Aux_Empleado_DOMLocalidad = Replace(Aux_Empleado_DOMLocalidad, "'", "\'")
            If Not IsNull(rs_Empleados!provdesc) Then Aux_Empleado_DOMProvincia = rs_Empleados!provdesc
            Aux_Empleado_DOMProvincia = Replace(Aux_Empleado_DOMProvincia, "'", "\'")
            If Not IsNull(rs_Empleados!telnro) Then Aux_Empleado_Telefono = rs_Empleados!telnro
            If Not IsNull(rs_Empleados!osdesc) Then Aux_Empleado_OSElegida = rs_Empleados!osdesc
            Aux_Empleado_OSElegida = Replace(Aux_Empleado_OSElegida, "'", "\'")
            If Not IsNull(rs_Empleados!empfaltagr) Then Aux_Empleado_FecIngreso = rs_Empleados!empfaltagr
        End If

        StrSql = "INSERT INTO rep_PS53 "
        StrSql = StrSql & "(empresa,bpronro,fecha, hora, iduser, ternro,"
        StrSql = StrSql & "abm, repfecha, cuit, razonsocial, cuil, nombre, tiponrodoc,"
        StrSql = StrSql & "sexo, fecnac, estadocivil, fecingreso, domcalle, domnro, "
        StrSql = StrSql & "dompiso, domdepto, domcodpost, domlocalidad, domprovinc, "
        StrSql = StrSql & "telefono, oselegida, dmtrcalle, dmtrnro, dmtrpiso, dmtrdepto,"
        StrSql = StrSql & "dmtrcodpost, dmtrloc, dmtrprov, telefonolab, nombres,apellido, tipodoc, nrodoc, provcod, estcivcod,nacionalidad,provcodempr,sitrev)"
        StrSql = StrSql & " VALUES("
        StrSql = StrSql & "'" & Aux_Empresa_ternro & "',"
        StrSql = StrSql & "'" & NroProcesoBatch & "',"
        StrSql = StrSql & ConvFecha(Date) & ","
        StrSql = StrSql & "'" & Time & "',"
        StrSql = StrSql & "'" & IdUser & "',"
        StrSql = StrSql & "'" & Aux_Empleado_ternro & "',"
        StrSql = StrSql & "'" & Aux_Empleado_ABM & "',"
        StrSql = StrSql & ConvFecha(HFecha) & ","
        StrSql = StrSql & "'" & Aux_Empresa_Cuit & "',"
        StrSql = StrSql & "'" & Aux_Empresa_RazonSocial & "',"
        StrSql = StrSql & "'" & Aux_Empleado_CUIL & "',"
        StrSql = StrSql & "'" & Aux_Empleado_Nombres & "',"
        StrSql = StrSql & "'" & Aux_Empleado_TipoNroDoc & "',"
        StrSql = StrSql & "'" & Aux_Empleado_Sexo & "',"
        StrSql = StrSql & Aux_Empleado_FecNac & ","
        StrSql = StrSql & "'" & Aux_Empleado_EstadoCivil & "',"
        If Aux_Empleado_FecIngreso <> "" Then
            StrSql = StrSql & ConvFecha(Aux_Empleado_FecIngreso) & ","
        Else
            StrSql = StrSql & "'" & Aux_Empleado_FecIngreso & "',"
        End If
        StrSql = StrSql & "'" & Aux_Empleado_DOMCalle & "',"
        StrSql = StrSql & "'" & Aux_Empleado_DOMNro & "',"
        StrSql = StrSql & "'" & Aux_Empleado_DOMPiso & "',"
        StrSql = StrSql & "'" & Left(Aux_Empleado_DOMDto, 8) & "',"
        StrSql = StrSql & "'" & Aux_Empleado_codigo & "',"
        StrSql = StrSql & "'" & Aux_Empleado_DOMLocalidad & "',"
        StrSql = StrSql & "'" & Aux_Empleado_DOMProvincia & "',"
        StrSql = StrSql & "'" & Aux_Empleado_Telefono & "',"
        StrSql = StrSql & "'" & Aux_Empleado_OSElegida & "',"
        StrSql = StrSql & "'" & Aux_Empresa_DMTRCalle & "',"
        StrSql = StrSql & "'" & Aux_Empresa_DMTRNum & "',"
        StrSql = StrSql & "'" & Aux_Empresa_DMTRPiso & "',"
        StrSql = StrSql & "'" & Aux_Empresa_DMTRDepto & "',"
        StrSql = StrSql & "'" & Aux_Empresa_DMTRCodPost & "',"
        StrSql = StrSql & "'" & Aux_Empresa_DMTRLoc & "',"
        StrSql = StrSql & "'" & Aux_Empresa_DMTRProv & "',"
        StrSql = StrSql & "'" & Aux_Empresa_TelefonoLab & "',"
        StrSql = StrSql & "'" & Aux_Empleado_nombre & "',"
        StrSql = StrSql & "'" & Aux_Empleado_apellido & "',"
        StrSql = StrSql & "'" & Aux_Empleado_tipodoc & "',"
        StrSql = StrSql & "'" & Aux_Empleado_nrodoc & "',"
        StrSql = StrSql & "'" & Aux_Empleado_codProv & "',"
        StrSql = StrSql & "'" & Aux_Empleado_codEstCiv & "',"
        StrSql = StrSql & "'" & Aux_Empleado_nacionalidad & "',"
        StrSql = StrSql & "'" & Aux_Empresa_provcod & "',"
        StrSql = StrSql & "'" & Aux_Empleado_sitRev & "'"
        
        
        
        
        StrSql = StrSql & ")"
        
        queryEmpleado = StrSql
               
        'Armo el query para pedir los familiares del empleado correspondiente al registro donde estoy parado, dentro del recordset rs_Empleados
        StrSql = "select familiar.ternro,parentesco.paredesc, CUIL.nrodoc NROCUIL, tercero.terape, tercero.terape2, tercero.ternom, tercero.ternom2, tipodocu.tidsigla, DU.nrodoc NRODU, tercero.terfecnac, tercero.tersex, estcivil.estcivdesabr, familiar.faminc, detdom.calle, detdom.nro, detdom.piso, detdom.oficdepto, detdom.codigopostal, localidad.locdesc, provincia.provdesc, telefono.telnro, estudio_actual.nivnro, estudio_actual.estactgra, estudio_actual.estactgra, familiar.famsalario,provincia.provcodext,estcivil.estcivcodext,nacionaldes,parcodext,famest"
        StrSql = StrSql & " from familiar join tercero on familiar.ternro = tercero.ternro and familiar.empleado = " & rs_batch_empleado!ternro
        StrSql = StrSql & " left join parentesco on familiar.parenro = parentesco.parenro"
        StrSql = StrSql & " left join cabdom on tercero.ternro = cabdom.ternro and cabdom.domdefault = -1"
        StrSql = StrSql & " left join detdom on cabdom.domnro = detdom.domnro"
        StrSql = StrSql & " left join localidad on localidad.locnro = detdom.locnro"
        StrSql = StrSql & " left join provincia on provincia.provnro = detdom.provnro"
        StrSql = StrSql & " left join telefono on telefono.domnro = cabdom.domnro and telefono.teldefault = -1"
        StrSql = StrSql & " left join ter_doc CUIL on CUIL.ternro = tercero.ternro and CUIL.tidnro = 10"
        StrSql = StrSql & " left join ter_doc DU on DU.ternro = tercero.ternro and (DU.tidnro = 1 or DU.tidnro = 2 or DU.tidnro = 3 or DU.tidnro = 4)"
        StrSql = StrSql & " left join tipodocu on DU.tidnro = tipodocu.tidnro"
        
        StrSql = StrSql & " left join nacionalidad on tercero.nacionalnro = nacionalidad.nacionalnro"
        
        StrSql = StrSql & " left join estcivil on tercero.estcivnro = estcivil.estcivnro"
        StrSql = StrSql & " left join estudio_actual on tercero.ternro = estudio_actual.ternro"
        OpenRecordset StrSql, rs_Familiares
        
        'Si NO hay familiares, entonces escapo sin agregar el reporte siquiera.
        If rs_Familiares.EOF Then
            Flog.Writeline "      El empleado no tiene familiares, no se genero el reporte."
            Flog.Writeline "      ------------------------"
        Else
            'Si hay familiares, entonces agrego el reporte, obtengo el ultimo repnro agregado y procedo a agregar empleados
            StrSql = queryEmpleado
            objConn.Execute StrSql, , adExecuteNoRecords
            'Obtengo el repnro recien agregado
            Aux_repnro = getLastIdentity(objConn, "rep_PS53")
        End If
            'Por cada familiar agrego el registro en la tabla rep_PS53_fam
        Do While Not rs_Familiares.EOF
            Flog.Writeline "      ------------------------"
            Flog.Writeline "      procesando familiar nro " & rs_Familiares!ternro

            If Not IsNull(rs_Familiares!parcodext) Then Aux_Familiar_codPar = rs_Familiares!parcodext
        
            Aux_Familiar_ternro = rs_Familiares!ternro
            Aux_Familiar_Nombres = rs_Familiares!ternom & " " & rs_Familiares!ternom2
            Aux_Familiar_apellido = rs_Familiares!terape & " " & rs_Familiares!terape2
            
            Aux_Familiar_tipodoc = rs_Familiares!tidsigla
            Aux_Familiar_nrodoc = rs_Familiares!NRODU
            If Not IsNull(rs_Familiares!nacionaldes) Then Aux_Familiar_nacionalidad = rs_Familiares!nacionaldes
            If Not IsNull(rs_Familiares!provcodext) Then Aux_Familiar_provcod = rs_Familiares!provcodext
            If Not IsNull(rs_Familiares!estcivcodext) Then Aux_Familiar_estcivcod = rs_Familiares!estcivcodext
            
            Aux_Familiar_Estudia = "N"
            
            If Not IsNull(rs_Familiares!famest) Then
                 If rs_Familiares!famest Then
                        Aux_Familiar_Estudia = "S"
                End If
            End If
            
            If Not IsNull(rs_Familiares!NROCuil) Then Aux_Familiar_CUIL = rs_Familiares!NROCuil
            If Not IsNull(rs_Familiares!paredesc) Then Aux_Familiar_Parentesco = rs_Familiares!paredesc
            If Not IsNull(rs_Familiares!terfecnac) Then Aux_Familiar_FecNac = rs_Familiares!terfecnac
            If Not IsNull(rs_Familiares!estcivdesabr) Then Aux_Familiar_EstadoCivil = rs_Familiares!estcivdesabr
            If Not IsNull(rs_Familiares!calle) Then Aux_Familiar_DOMCalle = rs_Familiares!calle
            If Not IsNull(rs_Familiares!nro) Then Aux_Familiar_DOMNro = rs_Familiares!nro
            If Not IsNull(rs_Familiares!piso) Then Aux_Familiar_DOMPiso = rs_Familiares!piso
            If Not IsNull(rs_Familiares!oficdepto) Then Aux_Familiar_DOMDto = rs_Familiares!oficdepto
            If Not IsNull(rs_Familiares!codigopostal) Then Aux_Familiar_DOMCodPost = rs_Familiares!codigopostal
            If Not IsNull(rs_Familiares!locdesc) Then Aux_Familiar_DOMLocalidad = rs_Familiares!locdesc
            If Not IsNull(rs_Familiares!provdesc) Then Aux_Familiar_DOMProvinc = rs_Familiares!provdesc
            If Not IsNull(rs_Familiares!telnro) Then Aux_Familiar_Telefono = rs_Familiares!telnro
            If Not IsNull(rs_Familiares!estactgra) Then Aux_Familiar_Aniocursa = rs_Familiares!estactgra
            If Not IsNull(rs_Familiares!famsalario) Then Aux_Familiar_Genderasigf = rs_Familiares!famsalario
            
            If CInt(Aux_Familiar_Genderasigf) = 0 Then
                Aux_Familiar_Genderasigf = "N"
            Else
                Aux_Familiar_Genderasigf = "S"
            End If
            
            If CStr(rs_Familiares!tersex) = 0 Then
                Aux_Familiar_Sexo = "F"
            Else
                Aux_Familiar_Sexo = "M"
            End If
            
            
            If CInt(rs_Familiares!faminc) = 0 Then
                    Aux_Familiar_Incapacidad = "N"
            Else
                    Aux_Familiar_Incapacidad = "S"
            End If
                    
            StrSql = "SELECT * FROM confrep WHERE repnro=" & tiprep_nro
            OpenRecordset StrSql, rs_Confrep
            If rs_Confrep.EOF Then
                HuboError = True
                Flog.Writeline "      ------------------------"
                Flog.Writeline "      ------------------------"
                Flog.Writeline "      No se encontró la configuracion para el reporte " & tiprep_nro
                Flog.Writeline "      ------------------------"
                Flog.Writeline "      ------------------------"
                MyRollbackTrans
                Exit Sub
            Else
                Flog.Writeline "      Accediendo a la configuracion del reporte para obtener la escolaridad. Tipo de reporte numero " & tiprep_nro & " , columnas 1 y 2."
            End If
            Do While Not rs_Confrep.EOF
                Select Case CInt(rs_Confrep!confnrocol)
                Case 1
                    confval_1 = rs_Confrep!confval
                    StrSql = "SELECT * FROM estudio_actual where ternro=" & Aux_Familiar_ternro
                    OpenRecordset StrSql, rs_Estudio
                    If Not rs_Estudio.EOF Then
                        If rs_Estudio!nivnro = confval_1 Then
                            Aux_Familiar_Escolaridad = "B"
                        End If
                    End If
                    If rs_Estudio.State = adStateOpen Then rs_Estudio.Close
                Case 2
                    confval_2 = rs_Confrep!confval
                    StrSql = "SELECT * FROM estudio_actual where ternro=" & Aux_Familiar_ternro
                    OpenRecordset StrSql, rs_Estudio
                    If Not rs_Estudio.EOF Then
                        If rs_Estudio!nivnro = confval_2 Then
                            Aux_Familiar_Escolaridad = "P"
                        End If
                    End If
                    If rs_Estudio.State = adStateOpen Then rs_Estudio.Close
                End Select
            rs_Confrep.MoveNext
            Loop
            
            If rs_Confrep.State = adStateOpen Then rs_Confrep.Close
            
            'Ejecuto el query
            StrSql = "INSERT INTO rep_PS53_fam "
            StrSql = StrSql & "(repnro,ternro,familiar,parentesco,cuil,nombre,tiponrodoc,fecnac,sexo,"
            StrSql = StrSql & "estadocivil,incapacidad,domcalle,domnro,dompiso,domdepto,domcodpost,"
            StrSql = StrSql & "domlocalidad,domprovinc,telefono,genderasigf,escolaridad,aniocursa,apellido,tipodoc,nrodoc,estcivcod,nacionalidad,codPar,est)"
            StrSql = StrSql & " VALUES("
            StrSql = StrSql & "'" & Aux_repnro & "',"
            StrSql = StrSql & "'" & Aux_Empleado_ternro & "',"
            StrSql = StrSql & "'" & Aux_Familiar_ternro & "',"
            StrSql = StrSql & "'" & Aux_Familiar_Parentesco & "',"
            StrSql = StrSql & "'" & Aux_Familiar_CUIL & "',"
            StrSql = StrSql & "'" & Aux_Familiar_Nombres & "',"
            StrSql = StrSql & "'" & Aux_Familiar_TipoNroDoc & "',"
            If Aux_Familiar_FecNac = " " Then
                StrSql = StrSql & "'" & Aux_Familiar_FecNac & "',"
            Else
                StrSql = StrSql & "" & ConvFecha(Aux_Familiar_FecNac) & ","
            End If
            StrSql = StrSql & "'" & Aux_Familiar_Sexo & "',"
            StrSql = StrSql & "'" & Aux_Familiar_EstadoCivil & "',"
            StrSql = StrSql & "'" & Aux_Familiar_Incapacidad & "',"
            StrSql = StrSql & "'" & Aux_Familiar_DOMCalle & "',"
            StrSql = StrSql & "'" & Aux_Familiar_DOMNro & "',"
            StrSql = StrSql & "'" & Aux_Familiar_DOMPiso & "',"
            StrSql = StrSql & "'" & Left(Aux_Familiar_DOMDto, 8) & "',"
            StrSql = StrSql & "'" & Aux_Familiar_DOMCodPost & "',"
            StrSql = StrSql & "'" & Aux_Familiar_DOMLocalidad & "',"
            StrSql = StrSql & "'" & Aux_Familiar_DOMProvinc & "',"
            StrSql = StrSql & "'" & Aux_Familiar_Telefono & "',"
            StrSql = StrSql & "'" & Aux_Familiar_Genderasigf & "',"
            StrSql = StrSql & "'" & Aux_Familiar_Escolaridad & "',"
            StrSql = StrSql & "'" & Aux_Familiar_Aniocursa & "',"
            StrSql = StrSql & "'" & Aux_Familiar_apellido & "',"
            StrSql = StrSql & "'" & Aux_Familiar_tipodoc & "',"
            StrSql = StrSql & "'" & Aux_Familiar_nrodoc & "',"
            StrSql = StrSql & "'" & Aux_Familiar_estcivcod & "',"
            StrSql = StrSql & "'" & Aux_Familiar_nacionalidad & "',"
            StrSql = StrSql & "'" & Aux_Familiar_codPar & "',"
            StrSql = StrSql & "'" & Aux_Familiar_Estudia & "'"
            StrSql = StrSql & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        rs_Familiares.MoveNext
        Loop
            If rs_Familiares.State = adStateOpen Then rs_Familiares.Close
            If rs_Empresa.State = adStateOpen Then rs_Empresa.Close
            If rs_Empleados.State = adStateOpen Then rs_Empleados.Close
        Flog.Writeline "------------------------------"
        rs_batch_empleado.MoveNext
        Progreso = Progreso + IncPorc
        TiempoAcumulado = GetTickCount
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
                 ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
                 "' WHERE bpronro = " & NroProcesoBatch
        objConn.Execute StrSql, , adExecuteNoRecords
        
        Loop
If rs_batch_empleado.State = adStateOpen Then rs_batch_empleado.Close

'Fin de la transaccion
MyCommitTrans

Set rs_Empresa = Nothing
Set rs_Empleados = Nothing
Set rs_Familiares = Nothing
Set rs_batch_empleado = Nothing

Exit Sub
CE:
    HuboError = True
    Flog.Writeline "Error: " & Err.Description
    Flog.Writeline "Ultimo sql Ejecutado: " & StrSql
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
Dim Aux As String

Dim HFecha As Date
Dim Aux_Separador As String

Aux_Separador = "@"
' Levanto cada parametro por separado, el separador de parametros es Aux_Separador

If Not IsNull(parametros) Then
    If Len(parametros) >= 1 Then
        pos1 = 1
        pos2 = InStr(pos1, parametros, Aux_Separador) - 1
        HFecha = CDate(Mid(parametros, pos1, pos2 - pos1 + 1))
    End If
End If

'Certificado Ansses de Servicios
Call Generar_Reporte(bpronro, HFecha)
End Sub




