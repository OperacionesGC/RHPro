Attribute VB_Name = "MdlRepIndividualPersonal"
Option Explicit

Const Version = 1.1 ' Gustavo Ring - Reporte Individual del Personal - Apex
Const FechaVersion = "07/11/2006"
Const tiprep_nro = 141



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
    'strCmdLine = "10516"
    
    
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
    OpenConnection strconexion, CnTraza
    
    Nombre_Arch = PathFLog & "Reporte_Individual_Personal" & "-" & NroProcesoBatch & ".log"
    
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
    StrSql = "SELECT * FROM batch_proceso WHERE (btprcnro = 141 ) AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        IdUser = rs_batch_proceso!IdUser
        Fecha = rs_batch_proceso!bprcfecha
        Hora = rs_batch_proceso!bprchora
        bprcparam = rs_batch_proceso!bprcparam
        
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        
        'Call Generar_Reporte(NroProcesoBatch, aux_sucursal, aux_mes, aux_anio, aux_empr)
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

Public Sub Generar_Reporte(ByVal bpronro As Long, ByVal sucursal As String, ByVal mes As Integer, ByVal Anio As Integer, ByVal empr As Integer)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del Reporte Individual de Personal
' Autor      : Gustavo Ring
' Fecha      : 24/10/2006
' Ult. Mod   :
'
' --------------------------------------------------------------------------------------------

'Variables auxiliares

'Parametros
Dim aux_sucursal As String
Dim aux_mes As Integer
Dim aux_anio As Integer

Dim HFecha As Date

HFecha = "01/" & mes & "/" & Anio


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
Dim Aux_repnro As String
Dim Aux_Empresa_Domicilio As String
Dim Aux_Empresa_DOMCalle As String
Dim Aux_Empresa_DOMNro As String
Dim Aux_Empresa_DOMPiso As String
Dim Aux_Empresa_DOMDto As String
Dim Aux_Empresa_DOMCodPost As String


'Empleado
Dim Aux_Empleado_ternro As String

Dim Aux_Empleado_Apellidos As String
Dim Aux_Empleado_Ficha As String
Dim Aux_Empleado_Nombres As String
Dim Aux_Empleado_Ingreso As String
Dim Aux_Empleado_Categoria As String
Dim Aux_empleado_Preaviso As String
Dim Aux_Empleado_Ocupacion As String
Dim Aux_empleado_Egreso As String
Dim Aux_Empleado_CambioOcupacion As String
Dim Aux_Empleado_FechaCambioOcupacion As String
Dim Aux_empleado_MotivoEgreso As String

Dim Aux_Empleado_VacacionesDias As String
Dim Aux_Empleado_VacacionesDesde As String
Dim Aux_Empleado_VacacionesHasta As String

Dim Aux_Empleado_SuspensionDias(3) As String
Dim Aux_Empleado_SuspensionDesde(3) As String
Dim Aux_Empleado_SuspensionHasta(3) As String

Dim Aux_Empleado_Nacimiento As String
Dim Aux_Empleado_Domicilio As String
Dim Aux_Empleado_TipoDocumento As String
Dim Aux_Empleado_Nrodocumento As String
Dim Aux_Empleado_Localidad As String
Dim Aux_Empleado_Caja As String
Dim Aux_Empleado_CajaNro As String
Dim Aux_Empleado_Provincia As String
Dim Aux_Empleado_Cuil As String
Dim Aux_Empleado_LugarNacimiento As String
Dim Aux_Empleado_Nacionalidad As String
Dim Aux_Empleado_EstadoCivil As String

Dim Aux_Empleado_DOMCalle As String
Dim Aux_Empleado_DOMNro As String
Dim Aux_Empleado_DOMPiso As String
Dim Aux_Empleado_DOMDto As String
Dim Aux_Empleado_DOMCodPost As String
Dim Aux_Empleado_Suspencion() As String

'Inicializo variables de Empleado
Aux_Empleado_ternro = ""

Aux_Empleado_Nombres = ""
Aux_Empleado_Apellidos = ""
Aux_Empleado_Categoria = ""
Aux_Empleado_Ocupacion = ""
Aux_Empleado_CambioOcupacion = ""

Aux_Empleado_FechaCambioOcupacion = ""

Aux_Empleado_Ficha = ""
Aux_Empleado_Ingreso = ""
Aux_empleado_Preaviso = ""
Aux_empleado_Egreso = ""
Aux_empleado_MotivoEgreso = ""

Aux_Empleado_VacacionesDias = ""
Aux_Empleado_VacacionesDesde = ""
Aux_Empleado_VacacionesHasta = ""

Aux_Empleado_SuspensionDias(1) = ""
Aux_Empleado_SuspensionDesde(1) = ""
Aux_Empleado_SuspensionHasta(1) = ""

Aux_Empleado_SuspensionDias(2) = ""
Aux_Empleado_SuspensionDesde(2) = ""
Aux_Empleado_SuspensionHasta(2) = ""

Aux_Empleado_SuspensionDias(3) = ""
Aux_Empleado_SuspensionDesde(3) = ""
Aux_Empleado_SuspensionHasta(3) = ""

Aux_Empleado_Nacimiento = ""
Aux_Empleado_LugarNacimiento = ""
Aux_Empleado_Nacionalidad = ""

Aux_Empleado_TipoDocumento = ""
Aux_Empleado_Nrodocumento = ""

Aux_Empleado_Caja = ""
Aux_Empleado_CajaNro = ""
Aux_Empleado_Cuil = ""

Aux_Empleado_EstadoCivil = ""
Aux_Empleado_DOMCalle = ""
Aux_Empleado_DOMNro = ""
Aux_Empleado_DOMPiso = ""
Aux_Empleado_DOMDto = ""
Aux_Empleado_DOMCodPost = ""
Aux_Empleado_Domicilio = ""
Aux_Empleado_Localidad = ""
Aux_Empleado_Provincia = ""

'Familiar

Dim Aux_Familiar_repnro As String
Dim Aux_Familiar_ternro As String
Dim Aux_Familiar_FechaNacimiento As String
Dim Aux_Familiar_Apellido As String
Dim Aux_Familiar_Nombre As String
Dim Aux_Familiar_Esposa As String
Dim Aux_Familiar_Hijo As String
Dim Aux_Familiar_HDiscapa As String
Dim Aux_Familiar_Vencimiento As String
Dim Aux_Familiar_Prenatal As String
Dim Aux_Familiar_Estudia As String

Aux_Familiar_repnro = ""
Aux_Familiar_ternro = ""
Aux_Familiar_FechaNacimiento = ""
Aux_Familiar_Apellido = ""
Aux_Familiar_Nombre = ""
Aux_Familiar_Esposa = ""
Aux_Familiar_Hijo = ""
Aux_Familiar_HDiscapa = ""
Aux_Familiar_Vencimiento = ""
Aux_Familiar_Prenatal = ""
Aux_Familiar_Estudia = ""

Dim confval_1 As String
Dim confval_2 As String
Dim confval_3 As String
Dim confval_4 As String
Dim confval_5 As String

Dim aux_fecha_desde As String
Dim aux_fecha_hasta As String
Dim aux_ultimo_dia As String
Dim aux_fecha As String
Dim aux_mes_proximo As Integer


'Registros
Dim rs_Empresa As New ADODB.Recordset
Dim rs_Consult As New ADODB.Recordset
Dim rs_Empleados As New ADODB.Recordset
Dim rs_Familiares As New ADODB.Recordset
Dim rs_vacaciones As New ADODB.Recordset
Dim rs_suspensiones As New ADODB.Recordset
Dim rs_batch_empleado As New ADODB.Recordset
Dim rs_repnro As New ADODB.Recordset
Dim rs_Confrep As New ADODB.Recordset
Dim rs_Estudio As New ADODB.Recordset
Dim rs_confval As New ADODB.Recordset
Dim rs_parientes As New ADODB.Recordset
Dim rs_causa As New ADODB.Recordset
Dim rs_puesto As New ADODB.Recordset
Dim rs_jubilacion As New ADODB.Recordset
Dim rs_categoria As New ADODB.Recordset
Dim rs_caja As New ADODB.Recordset
Dim rs_empaux As New ADODB.Recordset
Dim queryEmpleado As String
Dim Tercero
Dim I As Integer

On Error GoTo CE

' Comienzo la transaccion
    'MyBeginTrans
    
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
MyBeginTrans
    Aux_Empleado_ternro = ""

    Aux_Empleado_Nombres = ""
    Aux_Empleado_Apellidos = ""
    Aux_Empleado_Categoria = ""
    Aux_Empleado_Ocupacion = ""
    Aux_Empleado_CambioOcupacion = ""

    Aux_Empleado_FechaCambioOcupacion = ""

    Aux_Empleado_Ficha = ""
    Aux_Empleado_Ingreso = ""
    Aux_empleado_Preaviso = ""
    Aux_empleado_Egreso = ""
    Aux_empleado_MotivoEgreso = ""

    Aux_Empleado_VacacionesDias = ""
    Aux_Empleado_VacacionesDesde = ""
    Aux_Empleado_VacacionesHasta = ""

    Aux_Empleado_SuspensionDias(1) = ""
    Aux_Empleado_SuspensionDesde(1) = ""
    Aux_Empleado_SuspensionHasta(1) = ""

    Aux_Empleado_SuspensionDias(2) = ""
    Aux_Empleado_SuspensionDesde(2) = ""
    Aux_Empleado_SuspensionHasta(2) = ""

    Aux_Empleado_SuspensionDias(3) = ""
    Aux_Empleado_SuspensionDesde(3) = ""
    Aux_Empleado_SuspensionHasta(3) = ""

    Aux_Empleado_Nacimiento = ""
    Aux_Empleado_LugarNacimiento = ""
    Aux_Empleado_Nacionalidad = ""

    Aux_Empleado_TipoDocumento = ""
    Aux_Empleado_Nrodocumento = ""

    Aux_Empleado_Caja = ""
    Aux_Empleado_CajaNro = ""
    Aux_Empleado_Cuil = ""

    Aux_Empleado_EstadoCivil = ""
    Aux_Empleado_DOMCalle = ""
    Aux_Empleado_DOMNro = ""
    Aux_Empleado_DOMPiso = ""
    Aux_Empleado_DOMDto = ""
    Aux_Empleado_DOMCodPost = ""
    Aux_Empleado_Domicilio = ""
    Aux_Empleado_Localidad = ""
    Aux_Empleado_Provincia = ""

        
        
        Flog.Writeline "Procesando empleado id " & rs_batch_empleado!ternro
        Flog.Writeline "------------------------------"
        
        ' Sql Vacaciones
        StrSql = "SELECT tddesc, elfechadesde, elfechahasta, elcantdias "
        StrSql = StrSql & " FROM emp_lic INNER JOIN v_empleado ON emp_lic.empleado=v_empleado.ternro "
        StrSql = StrSql & " INNER JOIN tipdia ON emp_lic.tdnro=tipdia.tdnro "
        StrSql = StrSql & " LEFT JOIN proceso ON emp_lic.pronro = proceso.pronro "
        StrSql = StrSql & " LEFT JOIN lic_estado ON lic_estado.licestnro = emp_lic.licestnro "
        StrSql = StrSql & " Where tipdia.tdnro = 2 And ((Month(elfechadesde) = " & mes & " And Year(elfechadesde) = " & Anio & ") Or (Month(elfechahasta) =  " & mes & "  And Year(elfechahasta) =  " & Anio & " )) And ternro = " & rs_batch_empleado!ternro
    
        'Traigo las vacaciones
        
        OpenRecordset StrSql, rs_vacaciones
        If Not (rs_vacaciones.EOF) Then
            Aux_Empleado_VacacionesDias = rs_vacaciones!elcantdias
            Aux_Empleado_VacacionesDesde = rs_vacaciones!elfechadesde
            Aux_Empleado_VacacionesHasta = rs_vacaciones!elfechahasta
        End If
        
        ' Sql Suspensiones
        
        StrSql = "SELECT DISTINCT emp_licnro, tddesc, v_empleado.ternro, v_empleado.empleg, v_empleado.terape,v_empleado.ternom, elfechadesde, elfechahasta,emp_lic.pronro,prodesc, elcantdias, eltipo, elhoradesde, elhorahasta,licestdesabr, tipdia.tdest "
        StrSql = StrSql & " FROM emp_lic INNER JOIN v_empleado ON emp_lic.empleado=v_empleado.ternro "
        StrSql = StrSql & " INNER JOIN tipdia ON emp_lic.tdnro=tipdia.tdnro "
        StrSql = StrSql & " LEFT JOIN proceso ON emp_lic.pronro = proceso.pronro "
        StrSql = StrSql & " LEFT JOIN lic_estado ON lic_estado.licestnro = emp_lic.licestnro "
        StrSql = StrSql & " Where tipdia.tdnro = 16 And ((Month(elfechadesde) = " & mes & " And Year(elfechadesde) = " & Anio & ") Or (Month(elfechahasta) =  " & mes & "  And Year(elfechahasta) =  " & Anio & " )) And ternro = " & rs_batch_empleado!ternro

        'Traigo las Suspensiones
        
        OpenRecordset StrSql, rs_suspensiones
        
        I = 0
        While Not rs_suspensiones.EOF
           Aux_Empleado_SuspensionDias(I) = rs_suspensiones!elcantdias
           Aux_Empleado_SuspensionDesde(I) = rs_suspensiones!elfechadesde
           Aux_Empleado_SuspensionHasta(I) = rs_suspensiones!elfechahasta
           I = I + 1
        Wend
        
        aux_fecha_desde = "01/" & mes & "/" & Anio
        
        If CInt(mes) = 12 Then
                      aux_anio = Anio + 1
                      aux_mes_proximo = 1
                    Else
                      aux_anio = Anio
                      aux_mes_proximo = CInt(mes) + 1
        End If
        
        aux_fecha = "01/" & CStr(aux_mes_proximo) & "/" & CStr(aux_anio)
        
        aux_ultimo_dia = CStr(DateDiff("d", aux_fecha_desde, aux_fecha))
        
        aux_fecha_hasta = aux_ultimo_dia & "/" & mes & "/" & Anio
        
        'Traigo si hubo despido la causa
        StrSql = "select bajfec,caudes from fases "
        StrSql = StrSql & " inner join causa on causa.caunro=fases.caunro "
        StrSql = StrSql & " Where (month(bajfec)=" & mes & " and year(bajfec)=" & Anio & ") and empleado=" & rs_batch_empleado!ternro
        OpenRecordset StrSql, rs_causa
        
        If Not (rs_causa.EOF) Then
           Aux_empleado_Egreso = rs_causa!bajfec
           Aux_empleado_MotivoEgreso = rs_causa!caudes
        End If
        
        'Traigo la caja de jubilación
        StrSql = "select * from tipoestructura "
        StrSql = StrSql & " inner join his_estructura on his_estructura.tenro=tipoestructura.tenro "
        StrSql = StrSql & " inner join estructura on estructura.estrnro=his_estructura.estrnro "
        StrSql = StrSql & " Where tipoestructura.tenro = 15 And ternro =" & rs_batch_empleado!ternro
        StrSql = StrSql & " and (htetdesde<='" & aux_fecha_desde & "' AND (htethasta is null or htethasta>='" & aux_fecha_hasta & "')) "
        
        StrSql = StrSql & " order by htetdesde desc "
        
        OpenRecordset StrSql, rs_caja
        
        If Not (rs_caja.EOF) Then
            Aux_Empleado_Caja = rs_caja!estrdabr
        End If
        
        'Traigo el puesto
        StrSql = "select * from tipoestructura "
        StrSql = StrSql & " inner join his_estructura on his_estructura.tenro=tipoestructura.tenro"
        StrSql = StrSql & " inner join estructura on estructura.estrnro=his_estructura.estrnro"
        StrSql = StrSql & " Where tipoestructura.tenro = 4 and (htetdesde<='" & aux_fecha_desde & "' AND (htethasta is null or htethasta>='" & aux_fecha_hasta & "')) "
        StrSql = StrSql & " and ternro =" & rs_batch_empleado!ternro
        StrSql = StrSql & " order by htetdesde desc"
        
        OpenRecordset StrSql, rs_puesto

        If Not (rs_puesto.EOF) Then
            Aux_Empleado_Ocupacion = rs_puesto!estrdabr
            rs_puesto.MoveNext
            If Not (rs_puesto.EOF) Then
                   If (Month(rs_puesto!htetdesde) = mes) Then
                      Aux_Empleado_CambioOcupacion = Aux_Empleado_Ocupacion
                      rs_puesto.MoveNext
                      Aux_Empleado_Ocupacion = rs_puesto!estrdabr
                   End If
            End If
        End If
        
        'Traigo la categoria
        StrSql = "select * from tipoestructura "
        StrSql = StrSql & " inner join his_estructura on his_estructura.tenro=tipoestructura.tenro"
        StrSql = StrSql & " inner join estructura on estructura.estrnro=his_estructura.estrnro"
        StrSql = StrSql & " Where tipoestructura.tenro = 3 And (htetdesde<='" & aux_fecha_desde & "' AND (htethasta is null or htethasta>='" & aux_fecha_hasta & "')) " & " and ternro =" & rs_batch_empleado!ternro
        StrSql = StrSql & " order by htetdesde desc"

        OpenRecordset StrSql, rs_categoria

        If Not (rs_categoria.EOF) Then
            Aux_Empleado_Categoria = rs_categoria!estrdabr
        End If
        
       StrSql = "select empresa.ternro EMPTERNRO, detdom.calle, detdom.nro, detdom.piso, detdom.oficdepto, detdom.codigopostal, localidad.locdesc, provincia.provdesc, telefono.telnro, empresa.empnom, ter_doc.nrodoc, provincia.provcodext FROM "
       StrSql = StrSql & " empresa"
       StrSql = StrSql & " left join cabdom on empresa.ternro = cabdom.ternro and cabdom.domdefault = -1 "
       StrSql = StrSql & " left join detdom on cabdom.domnro = detdom.domnro "
       StrSql = StrSql & " left join localidad on localidad.locnro = detdom.locnro "
       StrSql = StrSql & " left join provincia on provincia.provnro = detdom.provnro "
       StrSql = StrSql & " left join telefono on telefono.domnro = cabdom.domnro and telefono.teldefault = -1 "
       StrSql = StrSql & " left join ter_doc on ter_doc.ternro = empresa.ternro and tidnro = 6"
       StrSql = StrSql & " where empresa.estrnro=" & empr
        
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
            If Not IsNull(rs_Empresa!provcodext) Then Aux_Empresa_provcod = rs_Empresa!provcodext
        End If
        
        Aux_Empresa_Domicilio = Aux_Empresa_DMTRCalle & " " & Aux_Empresa_DMTRNum & " " & Aux_Empresa_DMTRPiso & " " & Aux_Empresa_DMTRDepto & " " & Aux_Empresa_DOMDto & " " & Aux_Empresa_DMTRLoc & " (" & Aux_Empresa_DMTRCodPost & ")"
        
        If sucursal = "2" Then
            StrSql = " SELECT estrdabr FROM estructura WHERE estrnro = 429 "
            OpenRecordset StrSql, rs_Consult
            If Not rs_Consult.EOF Then
                 Aux_Empresa_Domicilio = rs_Consult!estrdabr
            Else
                 'Mantengo la actual
            End If
        End If

        'Traigo loa valores del empleado
        StrSql = "select estcivil.estcivdesabr, tercero.terape, tercero.terape2, tercero.ternom, tercero.ternom2, tipodocu.tidsigla, DU.nrodoc NRODU, tercero.tersex, tercero.terfecnac, detdom.calle, detdom.nro, detdom.piso, detdom.oficdepto, detdom.codigopostal, localidad.locdesc, provincia.provdesc, telefono.telnro, CUIL.nrodoc NROCUIL, osocial.osdesc, empleado.empfaltagr,provincia.provcodext,estcivil.estcivcodext,nacionaldes,empleado.empleg, pais.paisdesc from tercero "
        StrSql = StrSql & " left join cabdom on tercero.ternro = cabdom.ternro and cabdom.domdefault = -1"
        StrSql = StrSql & " left join detdom on cabdom.domnro = detdom.domnro"
        StrSql = StrSql & " left join localidad on localidad.locnro = detdom.locnro"
        StrSql = StrSql & " left join provincia on provincia.provnro = detdom.provnro"
        StrSql = StrSql & " left join pais on tercero.paisnro = pais.paisnro "
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
            StrSql = "select estcivil.estcivdesabr, tercero.terape, tercero.terape2, tercero.ternom, tercero.ternom2, tipodocu.tidsigla, DU.nrodoc NRODU, tercero.tersex, tercero.terfecnac, detdom.calle, detdom.nro, detdom.piso, detdom.oficdepto, detdom.codigopostal, localidad.locdesc, provincia.provdesc, telefono.telnro, CUIL.nrodoc NROCUIL, osocial.osdesc, v_empleado.empfaltagr,empleado.empleg,paisdesc from tercero "
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
            Aux_Empleado_Ficha = rs_Empleados!empleg
            Aux_Empleado_Apellidos = rs_Empleados!terape & " " & rs_Empleados!terape2
            Aux_Empleado_Nombres = rs_Empleados!ternom & " " & rs_Empleados!ternom2
            
            If Not IsNull(rs_Empleados!nacionaldes) Then Aux_Empleado_Nacionalidad = rs_Empleados!nacionaldes
            
            If Not IsNull(rs_Empleados!paisdesc) Then Aux_Empleado_LugarNacimiento = rs_Empleados!paisdesc
            
            If Not IsNull(rs_Empleados!estcivdesabr) Then Aux_Empleado_EstadoCivil = rs_Empleados!estcivdesabr
                        
            If Not IsNull(rs_Empleados!NROCuil) Then Aux_Empleado_Cuil = rs_Empleados!NROCuil
            
            If Not IsNull(rs_Empleados!terfecnac) Then Aux_Empleado_Nacimiento = rs_Empleados!terfecnac
            
            If Not IsNull(rs_Empleados!empfaltagr) Then Aux_Empleado_Ingreso = rs_Empleados!empfaltagr
            
            If Not IsNull(rs_Empleados!tidsigla) Then Aux_Empleado_TipoDocumento = rs_Empleados!tidsigla
            If Not IsNull(rs_Empleados!NRODU) Then Aux_Empleado_Nrodocumento = rs_Empleados!NRODU
            
            If Not IsNull(rs_Empleados!estcivdesabr) Then Aux_Empleado_EstadoCivil = rs_Empleados!estcivdesabr
            
            If Not IsNull(rs_Empleados!empfaltagr) Then Aux_Empleado_Ingreso = rs_Empleados!empfaltagr
            
            ' Domicilio
            If Not IsNull(rs_Empleados!calle) Then Aux_Empleado_DOMCalle = rs_Empleados!calle
            If Not IsNull(rs_Empleados!nro) Then Aux_Empleado_DOMNro = rs_Empleados!nro
            If Not IsNull(rs_Empleados!piso) Then Aux_Empleado_DOMPiso = rs_Empleados!piso
            If Not IsNull(rs_Empleados!oficdepto) Then Aux_Empleado_DOMDto = rs_Empleados!oficdepto
            Aux_Empleado_Domicilio = Aux_Empleado_DOMCalle & " " & Aux_Empleado_DOMNro & " " & Aux_Empleado_DOMPiso & " " & Aux_Empleado_DOMDto
            
            If Not IsNull(rs_Empleados!locdesc) Then Aux_Empleado_Localidad = rs_Empleados!locdesc
            If Not IsNull(rs_Empleados!provdesc) Then Aux_Empleado_Provincia = rs_Empleados!provdesc
            
        End If

        StrSql = "INSERT INTO rep_personal_individual "
        StrSql = StrSql & "(empresa,bpronro,fecha, hora, ternro,"
        StrSql = StrSql & "apellidos, nombres, categoria, ocupacion, cambioocupacion,fechacambio, ficha,"
        StrSql = StrSql & "ingreso,preaviso,egreso,motivoegreso,vacacionesdias,vacacionesdesde,vacacioneshasta, "
        StrSql = StrSql & "susp1dias,susp1desde,susp1hasta,susp2dias,susp2desde,susp2hasta,susp3dias,susp3desde,susp3hasta,"
        StrSql = StrSql & "nacimiento,lugarnacimiento,nrodoc,tipodoc,caja,estadocivil,cuil,"
        StrSql = StrSql & "domicilio,localidad,provincia,nacionalidad,cuit,direccionEmpresa,RazonSocial)"
        StrSql = StrSql & " VALUES("
        StrSql = StrSql & "'" & Aux_Empresa_ternro & "',"
        StrSql = StrSql & "'" & NroProcesoBatch & "',"
        StrSql = StrSql & "'" & Date & "',"
        StrSql = StrSql & "'" & Time & "',"
        StrSql = StrSql & "'" & rs_batch_empleado!ternro & "',"
        StrSql = StrSql & "'" & Aux_Empleado_Apellidos & "',"
        StrSql = StrSql & "'" & Aux_Empleado_Nombres & "',"
        StrSql = StrSql & "'" & Aux_Empleado_Categoria & "',"
        StrSql = StrSql & "'" & Aux_Empleado_Ocupacion & "',"
        StrSql = StrSql & "'" & Aux_Empleado_CambioOcupacion & "',"
        StrSql = StrSql & "'" & Aux_Empleado_FechaCambioOcupacion & "',"
        StrSql = StrSql & "'" & Aux_Empleado_Ficha & "',"
        StrSql = StrSql & "'" & Aux_Empleado_Ingreso & "',"
        StrSql = StrSql & "'" & Aux_empleado_Preaviso & "',"
        StrSql = StrSql & "'" & Aux_empleado_Egreso & "',"
        StrSql = StrSql & "'" & Aux_empleado_MotivoEgreso & "',"
        StrSql = StrSql & "'" & Aux_Empleado_VacacionesDias & "',"
        StrSql = StrSql & "'" & Aux_Empleado_VacacionesDesde & "',"
        StrSql = StrSql & "'" & Aux_Empleado_VacacionesHasta & "',"
        StrSql = StrSql & "'" & Aux_Empleado_SuspensionDias(0) & "',"
        StrSql = StrSql & "'" & Aux_Empleado_SuspensionDesde(0) & "',"
        StrSql = StrSql & "'" & Aux_Empleado_SuspensionHasta(0) & "',"
        StrSql = StrSql & "'" & Aux_Empleado_SuspensionDias(1) & "',"
        StrSql = StrSql & "'" & Aux_Empleado_SuspensionDesde(1) & "',"
        StrSql = StrSql & "'" & Aux_Empleado_SuspensionHasta(1) & "',"
        StrSql = StrSql & "'" & Aux_Empleado_SuspensionDias(2) & "',"
        StrSql = StrSql & "'" & Aux_Empleado_SuspensionDesde(2) & "',"
        StrSql = StrSql & "'" & Aux_Empleado_SuspensionHasta(2) & "',"
        StrSql = StrSql & "'" & Aux_Empleado_Nacimiento & "',"
        StrSql = StrSql & "'" & Aux_Empleado_LugarNacimiento & "',"
        StrSql = StrSql & "'" & Aux_Empleado_Nrodocumento & "',"
        StrSql = StrSql & "'" & Aux_Empleado_TipoDocumento & "',"
        StrSql = StrSql & "'" & Aux_Empleado_Caja & "',"
        StrSql = StrSql & "'" & Aux_Empleado_EstadoCivil & "',"
        StrSql = StrSql & "'" & Aux_Empleado_Cuil & "',"
        StrSql = StrSql & "'" & Aux_Empleado_Domicilio & "',"
        StrSql = StrSql & "'" & Aux_Empleado_Localidad & "',"
        StrSql = StrSql & "'" & Aux_Empleado_Provincia & "',"
        StrSql = StrSql & "'" & Aux_Empleado_Nacionalidad & "',"
        StrSql = StrSql & "'" & Aux_Empresa_Cuit & "',"
        StrSql = StrSql & "'" & Aux_Empresa_Domicilio & "',"
        StrSql = StrSql & "'" & Aux_Empresa_RazonSocial & "'"
     
        StrSql = StrSql & ")"
        
        objConn.Execute StrSql, , adExecuteNoRecords
        Aux_repnro = getLastIdentity(objConn, "rep_personal_individual")
               
        'Armo el query para pedir los familiares del empleado correspondiente al registro donde estoy parado, dentro del recordset rs_Empleados
        StrSql = "select familiar.ternro,famDGIhasta,famest, tercero.terape, tercero.terape2, tercero.ternom, tercero.ternom2,tercero.terfecnac,parcodext "
        StrSql = StrSql & " from familiar inner join tercero on familiar.ternro = tercero.ternro and familiar.empleado = " & rs_batch_empleado!ternro
        StrSql = StrSql & " inner join parentesco on familiar.parenro = parentesco.parenro"
        'StrSql = StrSql & " left join estcivil on tercero.estcivnro = estcivil.estcivnro order by familiar.ternro "
        'StrSql = StrSql & " left join estudio_actual on tercero.ternro = estudio_actual.ternro"
        OpenRecordset StrSql, rs_Familiares
        
        
        'Por cada familiar agrego el registro en la tabla rep_personal_individual_fam
       Do While Not rs_Familiares.EOF
            Flog.Writeline "      ------------------------"
            Flog.Writeline "      procesando familiar nro " & rs_Familiares!ternro

            Aux_Familiar_ternro = " "
            Aux_Familiar_FechaNacimiento = " "
            Aux_Familiar_Apellido = " "
            Aux_Familiar_Nombre = " "
            Aux_Familiar_Esposa = " "
            Aux_Familiar_Hijo = " "
            Aux_Familiar_HDiscapa = " "
            Aux_Familiar_Vencimiento = " "
            Aux_Familiar_Prenatal = " "
            Aux_Familiar_Estudia = " "
            Aux_Familiar_Esposa = " "
            Aux_Familiar_Hijo = " "
            Aux_Familiar_HDiscapa = " "
            Aux_Familiar_Estudia = " "
            
            Aux_Familiar_ternro = rs_Familiares!ternro
            Aux_Familiar_Nombre = rs_Familiares!terape & " " & rs_Familiares!terape2 & " " & rs_Familiares!ternom & " " & rs_Familiares!ternom2
            
            If Not IsNull(rs_Familiares!famDGIhasta) Then Aux_Familiar_Vencimiento = rs_Familiares!famDGIhasta
            If Not IsNull(rs_Familiares!terfecnac) Then Aux_Familiar_FechaNacimiento = rs_Familiares!terfecnac
            If Not IsNull(rs_Familiares!famest) Then
                    If rs_Familiares!famest Then
                        Aux_Familiar_Estudia = "X"
                    End If
            End If
            If Not IsNull(rs_Familiares!famDGIhasta) Then Aux_Familiar_Vencimiento = rs_Familiares!famDGIhasta
            
            StrSql = "SELECT * FROM confrep WHERE repnro=182"
            OpenRecordset StrSql, rs_Confrep
            If rs_Confrep.EOF Then
                HuboError = True
                Flog.Writeline "      -----------------------------------------------------------"
                Flog.Writeline "      -----------------------------------------------------------"
                Flog.Writeline "      No se encontró la configuracion para el reporte 182"
                Flog.Writeline "      -----------------------------------------------------------"
                Flog.Writeline "      -----------------------------------------------------------"
                MyRollbackTrans
                Exit Sub
            Else
                Flog.Writeline "      Accediendo a la configuracion del reporte para obtener los parentescos "
            End If
            
            Do While Not rs_Confrep.EOF
                Select Case CInt(rs_Confrep!confnrocol)
                  Case 1
                    confval_1 = rs_Confrep!confval2
                  Case 2
                    confval_2 = rs_Confrep!confval2
                  Case 3
                    confval_3 = rs_Confrep!confval2
                  Case 4
                    confval_4 = rs_Confrep!confval2
                End Select
                rs_Confrep.MoveNext
            Loop
            
            Select Case rs_Familiares!parcodext
                Case confval_1
                    Aux_Familiar_Esposa = "X"
                Case confval_2
                    Aux_Familiar_Hijo = "X"
                Case confval_3
                    Aux_Familiar_HDiscapa = "X"
                Case confval_4
                    Aux_Familiar_Prenatal = "X"
            End Select
            
            If rs_Confrep.State = adStateOpen Then rs_Confrep.Close
            
            'Ejecuto el query
            StrSql = "INSERT INTO rep_personal_individual_fam "
            StrSql = StrSql & "(repnro,ternro,nombre,esposa,hijo,hijoD,prenatal,estudia,vencimiento, nacimiento) "
            StrSql = StrSql & " VALUES("
            StrSql = StrSql & "'" & Aux_repnro & "',"
            StrSql = StrSql & "'" & Aux_Empleado_ternro & "',"
            StrSql = StrSql & "'" & Aux_Familiar_Nombre & "',"
            StrSql = StrSql & "'" & Aux_Familiar_Esposa & "',"
            StrSql = StrSql & "'" & Aux_Familiar_Hijo & "',"
            StrSql = StrSql & "'" & Aux_Familiar_HDiscapa & "',"
            StrSql = StrSql & "'" & Aux_Familiar_Prenatal & "',"
            StrSql = StrSql & "'" & Aux_Familiar_Estudia & "',"
            StrSql = StrSql & "'" & Aux_Familiar_Vencimiento & "',"
            StrSql = StrSql & "'" & Aux_Familiar_FechaNacimiento & "'"
            StrSql = StrSql & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
            rs_Familiares.MoveNext
 
        Loop
            If rs_Familiares.State = adStateOpen Then rs_Familiares.Close
            If rs_Empresa.State = adStateOpen Then rs_Empresa.Close
            If rs_Empleados.State = adStateOpen Then rs_Empleados.Close
        Flog.Writeline "------------------------------"
        
        'objConn.Execute StrSql, , adExecuteNoRecords
        
        rs_batch_empleado.MoveNext
        
        If Not HuboError Then
            StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
        Else
            StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
        End If
        Progreso = Progreso + IncPorc
        TiempoAcumulado = GetTickCount
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
                 ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
                 "' WHERE bpronro = " & NroProcesoBatch
        CnTraza.Execute StrSql, , adExecuteNoRecords
 MyCommitTrans
        Loop
If rs_batch_empleado.State = adStateOpen Then rs_batch_empleado.Close

'Fin de la transaccion
'MyCommitTrans

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
Dim pos3 As Integer
Dim pos4 As Integer
Dim aux_sucursal
Dim aux_mes
Dim aux_anio
Dim aux_empr
Dim ArrParametros

Dim aux As String

Dim HFecha As Date
Dim Aux_Separador As String

Aux_Separador = "@"
' Levanto cada parametro por separado, el separador de parametros es Aux_Separador

If Not IsNull(parametros) Then
    If Len(parametros) >= 1 Then
        ArrParametros = Split(parametros, Aux_Separador, -1)
        aux_sucursal = ArrParametros(0)
        aux_mes = ArrParametros(1)
        aux_anio = ArrParametros(2)
        aux_empr = ArrParametros(3)
    End If
End If

'Reporte Individual del personal
Call Generar_Reporte(bpronro, aux_sucursal, aux_mes, aux_anio, aux_empr)
End Sub




