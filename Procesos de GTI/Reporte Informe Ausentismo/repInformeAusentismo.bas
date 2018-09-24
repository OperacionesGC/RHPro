Attribute VB_Name = "repInformeAusentismo"
Option Explicit

'Global Const Version = "1.00"
'Global Const FechaModificacion = "24/08/2006"
'Global Const UltimaModificacion = "Inicial"

'Global Const Version = "1.01"
'Global Const FechaModificacion = "14/10/2006"
'Global Const UltimaModificacion = "Encriptacion de Conexion"
'Global Const UltimaModificacion1 = "Manuel Lopez"


'Global Const Version = "1.02"
'Global Const FechaModificacion = "11/07/2007"
'Global Const UltimaModificacion = "Detallado de logs" '----- esto quedó pendiente .......

'Global Const Version = "1.03"
'Global Const FechaModificacion = "22/01/2009"
'Global Const UltimaModificacion = "FGZ" ' Encriptacion de string de conexion y Schema para Oracle

'Global Const Version = "1.04"
'Global Const FechaModificacion = "17/10/2014"
'Global Const UltimaModificacion = "Sebastian Stremel - CAS-27374 - FARMOGRAFICA - CUSTOM FILTRO DE REPORTE DE AUSENTISMO - Se agrego filtrado por 3 niveles de estructura."

'Global Const Version = "1.05"
'Global Const FechaModificacion = "20/10/2014"
'Global Const UltimaModificacion = "Carmen Quintero - CAS-27570 - AGD - ERROR EN REPORTE DE AUSENTISMO - Se cambio el tipo de dato de las variables: total_dotacion_th_AR, total_dotacion_th_ANR, total_dotacion_AR y total_dotacion_ANR "

'Global Const Version = "1.06"
'Global Const FechaModificacion = "06/11/2014"
'Global Const UltimaModificacion = "Sebastian Stremel - CAS-27374 - FARMOGRAFICA - CUSTOM FILTRO DE REPORTE DE AUSENTISMO [Entrega 2] - Se agrego semilla de encriptacion."

'Global Const Version = "1.07"
'Global Const FechaModificacion = "13/11/2014"
'Global Const UltimaModificacion = "Carmen Quintero - CAS-27570 - AGD - ERROR EN REPORTE DE AUSENTISMO [Entrega 2] - Se cambio el tipo de dato de las variables: total_dotacion_th_AR, total_dotacion_th_ANR, total_dotacion_AR y total_dotacion_ANR "

'Global Const Version = "1.08"
'Global Const FechaModificacion = "02/12/2014"
'Global Const UltimaModificacion = "Fernandez, Matias-CAS-27570 - AGD - Error en filtro de reporte Gen de Rep. de ausentismo- Se agrego el filtro de empresa. "


Global Const Version = "1.09"
Global Const FechaModificacion = "12/01/2015"
Global Const UltimaModificacion = "Fernandez, Matias-CAS-27570 - AGD - ERROR EN DOTACION DE REPORTE DE AUSENTISMO - se corrigio los totales de empleados"



'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------
Dim fs, f
Dim NroLinea As Long
Dim crpNro As Long
Dim RegLeidos As Long
Dim RegError As Long
Dim RegFecha As Date
Dim NroProceso As Long

Global Path As String
Global NArchivo As String
Global HuboErrores As Boolean
Global EmpErrores As Boolean

'Global Tabulador As Long -MDF-
Global TiempoInicialProceso
Global TiempoAcumulado

Global IdUser As String
Global Fecha As Date
Global Hora As String

Private Type Horas
    thnro As Long
    thdesc As String
    toths As Double
End Type

Global Arr_THoras_AR() As Horas
Global Arr_THoras_ANR() As Horas


Private Sub Main()

Dim NombreArchivo As String
Dim Directorio As String
Dim CArchivos
Dim archivo
Dim Folder
Dim strCmdLine As String
Dim Nombre_Arch As String

Dim StrSql As String
Dim objRs As New ADODB.Recordset
Dim cantRegistros
Dim PID As String
Dim parametros As String
Dim ArrParametros

Dim fecdesde As Date
Dim fechasta As Date
Dim empresa As Long

Dim tenro1 As Integer
Dim estrnro1 As Integer
Dim tenro2 As Integer
Dim estrnro2 As Integer
Dim tenro3 As Integer
Dim estrnro3 As Integer

Dim usaParametros As Boolean
usaParametros = False

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

    TiempoInicialProceso = GetTickCount
    
    HuboErrores = False
    
    Nombre_Arch = PathFLog & "ReporteInfAusentismo" & "-" & NroProceso & ".log"
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    'Abro la conexion
    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    On Error Resume Next
    OpenConnection strconexion, objConnProgreso
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    
    On Error GoTo ME_Main:
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Version = " & Version
    Flog.writeline "Modificacion = " & UltimaModificacion
    Flog.writeline "Fecha = " & FechaModificacion
    Flog.writeline "PID = " & PID
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "Inicio :" & Now
    
    Flog.writeline Espacios(Tabulador * 0) & "Inicio Proceso Informe de Ausentismos: " & Now
    Flog.writeline Espacios(Tabulador * 0) & "Cambio el estado del proceso a Procesando"
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & ", bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objConnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline Espacios(Tabulador * 0) & "Obtengo los datos del proceso"
    
    TiempoAcumulado = GetTickCount
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, objRs
    
    If Not objRs.EOF Then
       IdUser = objRs!IdUser
       Fecha = objRs!bprcfecha
       Hora = objRs!bprchora
       
       'Obtengo los parametros del proceso
       parametros = objRs!bprcparam
       ArrParametros = Split(parametros, "@")
       
       'Empresa
       empresa = ArrParametros(0)
       
       'Fecha desde
       fecdesde = ArrParametros(1)
       
       'Fecha hasta
       fechasta = ArrParametros(2)
       
       'tenro1
       If UBound(ArrParametros) > 3 Then
            tenro1 = ArrParametros(3)
            estrnro1 = ArrParametros(4)
            tenro2 = ArrParametros(5)
            estrnro2 = ArrParametros(6)
            tenro3 = ArrParametros(7)
            estrnro3 = IIf(EsNulo(ArrParametros(8)), 0, ArrParametros(8))
            usaParametros = True
       Else
            tenro1 = 0
            estrnro1 = 0
            tenro2 = 0
            estrnro2 = 0
            tenro3 = 0
            estrnro3 = 0
       End If
       
       ' Proceso que genera los datos
       Call GenerarDatos(empresa, fecdesde, fechasta, tenro1, estrnro1, tenro2, estrnro2, tenro3, estrnro3, usaParametros)
       
    Else
       Exit Sub
    End If
    
    If objRs.State = adStateOpen Then objRs.Close
    
    'Actualizo el estado del proceso
    If Not HuboErrores Then
       StrSql = "UPDATE batch_proceso SET  bprcprogreso = 100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
       Flog.writeline "Proceso Finalizado Correctamente"
    Else
       StrSql = "UPDATE batch_proceso SET  bprcprogreso = 100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProceso
       Flog.writeline "Proceso Incompleto"
    End If
    objConnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Fin :" & Now
    Flog.Close

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
        StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Error General', bprcpid = " & PID & " WHERE bpronro = " & NroProceso
        objConnProgreso.Execute StrSql, , adExecuteNoRecords
    MyCommitTrans
End Sub
'--------------------------------------------------------------------
' Se encarga de generar los datos
'--------------------------------------------------------------------
Sub GenerarDatos(ByVal empresa As String, ByVal fecdesde As Date, ByVal fechasta As Date, ByVal tenro1 As Integer, ByVal estrnro1 As Integer, ByVal tenro2 As Integer, ByVal estrnro2 As Integer, ByVal tenro3 As Integer, ByVal estrnro3 As Integer, ByVal usaParametros As Boolean)


Dim StrSql As String
Dim rsConsult As New ADODB.Recordset
Dim rsacumdiario As New ADODB.Recordset
Dim rsconfetiq As New ADODB.Recordset

'Variables donde se guardan los datos del INSERT final
Dim EmpCuit As String
Dim EmpNombre As String
Dim EmpDire As String
Dim EmpTernro As String
Dim EmpLogo As String
Dim EmpLogoAlto
Dim EmpLogoAncho

Dim Cant_THnro As Integer

Dim lista_estrnro As String
Dim lista_AR As String
Dim lista_ANR As String
Dim Tenro As String
'Dim total_dotacion_th_AR As Integer
'Dim total_dotacion_AR As Integer
'Dim total_dotacion_th_ANR As Integer
'Dim total_dotacion_ANR As Integer
Dim estrnro_ant As Long
Dim ternro_ant As Long
Dim Progreso As Double
Dim cantidadProcesada As Integer
Dim IncPorc As Double
Dim fecdesde_empl As Date
Dim fechasta_empl As Date
Dim orden As Integer
Dim tothsestr As Double
Dim total_AR As Double
Dim total_ANR As Double
Dim confetiq_ant As String

Dim Cargar_registro As Boolean
Dim i As Integer

Dim total_dotacion_th_AR As Long
Dim total_dotacion_AR As Long
Dim total_dotacion_th_ANR As Long
Dim total_dotacion_ANR As Long

total_dotacion_AR = 0
total_dotacion_th_ANR = 0
total_dotacion_ANR = 0

On Error GoTo MError

    MyBeginTrans
    
    '-------------------------------------------------------------------------
    ' Busco los datos de la empresa
    '--------------------------------------------------------------------------
    StrSql = "SELECT empresa.empnom,empresa.ternro, detdom.calle,nro,codigopostal, localidad.locdesc "
    StrSql = StrSql & " FROM empresa "
    StrSql = StrSql & " LEFT JOIN cabdom ON empresa.ternro = cabdom.ternro "
    StrSql = StrSql & " LEFT JOIN detdom ON detdom.domnro = cabdom.domnro "
    StrSql = StrSql & " LEFT JOIN localidad ON detdom.locnro = localidad.locnro "
    StrSql = StrSql & " WHERE empresa.estrnro = " & empresa
    OpenRecordset StrSql, rsConsult
    If rsConsult.EOF Then
        Flog.writeline "***************************************************************"
        Flog.writeline "Error. No se encontro la empresa."
        Flog.writeline "***************************************************************"
        GoTo Fin_error
    Else
        EmpNombre = rsConsult!empnom
        EmpDire = rsConsult!calle & " " & rsConsult!nro & "<br>" & rsConsult!codigopostal & " " & rsConsult!locdesc
        EmpTernro = rsConsult!Ternro
    End If
    rsConsult.Close
    
    '-------------------------------------------------------------------------
    'Consulta para obtener el cuit de la empresa
    '-------------------------------------------------------------------------
    StrSql = "SELECT cuit.nrodoc FROM tercero " & _
             " INNER JOIN ter_doc cuit ON (tercero.ternro = cuit.ternro and cuit.tidnro = 6)" & _
             " Where tercero.ternro =" & EmpTernro
    OpenRecordset StrSql, rsConsult
    If rsConsult.EOF Then
        Flog.writeline "No se encontró el CUIT de la Empresa."
        EmpCuit = "&nbsp;"
    Else
        EmpCuit = rsConsult!nrodoc
    End If
    rsConsult.Close
    
    '-------------------------------------------------------------------------
    'Consulta para buscar el logo de la empresa
    '-------------------------------------------------------------------------
    StrSql = "SELECT ter_imag.terimnombre, tipoimag.tipimdire, tipoimag.tipimanchodef, tipoimag.tipimaltodef" & _
        " FROM ter_imag " & _
        " INNER JOIN tipoimag ON tipoimag.tipimnro = 1 AND tipoimag.tipimnro = ter_imag.tipimnro" & _
        " AND ter_imag.ternro =" & EmpTernro
    OpenRecordset StrSql, rsConsult
    If rsConsult.EOF Then
        Flog.writeline "No se encontró el Logo de la Empresa."
        EmpLogo = ""
        EmpLogoAlto = 0
        EmpLogoAncho = 0
    Else
        EmpLogo = rsConsult!tipimdire & rsConsult!terimnombre
        EmpLogoAlto = rsConsult!tipimaltodef
        EmpLogoAncho = rsConsult!tipimanchodef
    End If
    rsConsult.Close

    '-------------------------------------------------------------------------
    'Busco la configuración del reporte
    '-------------------------------------------------------------------------
    StrSql = "SELECT * FROM confrep WHERE repnro = 164"
    OpenRecordset StrSql, rsConsult
    
    lista_estrnro = "0"
    lista_AR = "0"
    lista_ANR = "0"
    Tenro = "0"
    Do Until rsConsult.EOF
        Select Case rsConsult!conftipo
            Case "TE":
                Tenro = Tenro & "," & rsConsult!confval
            Case "EST":
                lista_estrnro = lista_estrnro & "," & rsConsult!confval
            Case "AR":
                lista_AR = lista_AR & "," & rsConsult!confval
            Case "ANR":
                lista_ANR = lista_ANR & "," & rsConsult!confval
            Case Else:
                Flog.writeline "***************************************************************"
                Flog.writeline "Error. Tipo '" & rsConsult!conftipo & "' no reconocido en la configuración. Los tipos válidos son TE, EST, AR, ANR."
                Flog.writeline "***************************************************************"
                GoTo Fin_error
        End Select
        rsConsult.MoveNext
    Loop
    rsConsult.Close
    
    If Tenro = "0" Then
        Flog.writeline "***************************************************************"
        Flog.writeline "Error. Se deben configurar una columna de tipo TE en la configuración del reporte."
        Flog.writeline "***************************************************************"
        GoTo Fin_error
    End If
              
    Tenro = Mid(Tenro, 3, Len(Tenro) - 2)
    If InStr(Tenro, ",") > 0 Then
        Flog.writeline "***************************************************************"
        Flog.writeline "Error. Se permite una sola columna de tipo TE en la configuración del reporte."
        Flog.writeline "***************************************************************"
        GoTo Fin_error
    End If
    
    If lista_estrnro = "0" Then
        Flog.writeline "***************************************************************"
        Flog.writeline "Error. Se deben configurar columnas de tipos EST en la configuración del reporte."
        Flog.writeline "***************************************************************"
        GoTo Fin_error
    End If
              
    If lista_AR = "0" And lista_ANR = "0" Then
        Flog.writeline "***************************************************************"
        Flog.writeline "Error. Se debe configurar al menos una columna del tipo AR o ANR en la configuración del reporte."
        Flog.writeline "***************************************************************"
        GoTo Fin_error
    End If
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "    TE  --> " & Tenro
    Flog.writeline "    EST --> " & lista_estrnro
    Flog.writeline "    AR  --> " & lista_AR
    Flog.writeline "    ANR --> " & lista_ANR
    Flog.writeline "-----------------------------------------------------------------"

    Cant_THnro = UBound(Split(lista_AR, ","))
    ReDim Preserve Arr_THoras_AR(Cant_THnro) As Horas
    
    Call CargarValorInicial(lista_AR, True)
    
    
    Cant_THnro = UBound(Split(lista_ANR, ","))
    ReDim Preserve Arr_THoras_ANR(Cant_THnro) As Horas
    
    Call CargarValorInicial(lista_ANR, False)
    
    
    '-------------------------------------------------------------------------
    'Inserto los datos de la cabecera
    '-------------------------------------------------------------------------
    StrSql = "SELECT tipoestructura.tenro,tedabr FROM tipoestructura "
    StrSql = StrSql & " WHERE tipoestructura.tenro = " & Tenro
    OpenRecordset StrSql, rsConsult
        
    If Not rsConsult.EOF Then
        StrSql = "INSERT INTO rep_inf_aus (bpronro,fecdesde,fechasta,tenro,tedabr,"
        StrSql = StrSql & "empnombre,empdire,empcuit,emplogo,emplogoalto,emplogoancho,fecha,hora,IdUser) "
        StrSql = StrSql & "VALUES (" & NroProceso & "," & ConvFecha(fecdesde)
        StrSql = StrSql & "," & ConvFecha(fechasta) & "," & rsConsult!Tenro & ",'" & rsConsult!tedabr & "','"
        StrSql = StrSql & EmpNombre & "','" & EmpDire & "','" & EmpCuit & "','" & EmpLogo
        StrSql = StrSql & "'," & EmpLogoAlto & "," & EmpLogoAncho & "," & ConvFecha(Fecha) & ",'" & Hora & "','" & IdUser & "')"

        objConn.Execute StrSql, , adExecuteNoRecords
        
        Flog.writeline "    Inserto los datos de la cabecera "
    Else
        Flog.writeline "***************************************************************"
        Flog.writeline "Error. Tipo de Estructura no definida."
        Flog.writeline "***************************************************************"
    End If
        
    rsConsult.Close
    
    
    '-------------------------------------------------------------------------
    'Busco los empleados que pertenecen a las estructuras y a la empresa seleccionada en el rango de fechas
    '-------------------------------------------------------------------------
    StrSql = " SELECT h_estr.ternro, h_estr.estrnro, confrep.confetiq, estrdabr, h_estr.htetdesde estrfecdesde, "
    StrSql = StrSql & " h_estr.htethasta estrfechasta, emp.htetdesde empfecdesde, emp.htethasta empfechasta "
    If usaParametros Then
        If CStr(tenro3) <> "" And CStr(tenro3) <> "0" Then
            StrSql = StrSql & ", estact1.tenro tenro1, estact1.estrnro estrnro1, estact2.tenro tenro2, estact2.estrnro estrnro2, estact3.tenro tenro3, estact3.estrnro estrnro3  "
        Else
            If CStr(tenro2) <> "" And CStr(tenro2) <> "0" Then
                StrSql = StrSql & ", estact1.tenro tenro1, estact1.estrnro estrnro1, estact2.tenro tenro2, estact2.estrnro estrnro2 "
            Else
                If CStr(tenro1) <> "" And CStr(tenro1) <> "0" Then
                    StrSql = StrSql & ", estact1.tenro tenro1, estact1.estrnro estrnro1"
                End If
            End If
        End If
    End If 'mdf
    StrSql = StrSql & " FROM estructura "
    StrSql = StrSql & " INNER JOIN confrep ON estructura.estrnro = confrep.confval AND confrep.repnro = 164 AND conftipo = 'EST' "
    StrSql = StrSql & " INNER JOIN his_estructura h_estr ON estructura.estrnro = h_estr.estrnro "
    StrSql = StrSql & " AND h_estr.htetdesde <= " & ConvFecha(fechasta) & " AND (h_estr.htethasta IS NULL OR h_estr.htethasta >= " & ConvFecha(fecdesde) & ") AND h_estr.tenro = " & Tenro
    StrSql = StrSql & " INNER JOIN his_estructura emp ON h_estr.ternro = emp.ternro "
    StrSql = StrSql & " AND emp.htetdesde <= " & ConvFecha(fechasta) & " AND (emp.htethasta IS NULL OR emp.htethasta >= " & ConvFecha(fecdesde) & ") AND emp.tenro = 10 "
    StrSql = StrSql & " AND emp.estrnro= " & empresa 'mdf

    If usaParametros Then
        If CStr(tenro3) <> "" And CStr(tenro3) <> "0" Then
            StrSql = StrSql & " INNER JOIN his_estructura estact1 ON h_estr.ternro = estact1.ternro  AND estact1.tenro  = " & tenro1
            StrSql = StrSql & " AND (estact1.htetdesde<=" & ConvFecha(fechasta) & " AND (estact1.htethasta is null or estact1.htethasta>=" & ConvFecha(fechasta) & "))"
            If CStr(estrnro1) <> "" And CStr(estrnro1) <> "-1" Then 'cuando se le asigna un valor al nivel 1
                StrSql = StrSql & " AND estact1.estrnro =" & estrnro1
            End If
            
            StrSql = StrSql & " INNER JOIN his_estructura estact2 ON h_estr.ternro = estact2.ternro  AND estact2.tenro  = " & tenro2
            StrSql = StrSql & " AND (estact2.htetdesde<=" & ConvFecha(fechasta) & " AND (estact2.htethasta is null or estact2.htethasta>=" & ConvFecha(fechasta) & "))"
            If CStr(estrnro2) <> "" And CStr(estrnro2) <> "-1" Then 'cuando se le asigna un valor al nivel 2
                StrSql = StrSql & " AND estact2.estrnro =" & estrnro2
            End If
            
            StrSql = StrSql & " INNER JOIN his_estructura estact3 ON h_estr.ternro = estact3.ternro  AND estact3.tenro  = " & tenro3
            StrSql = StrSql & " AND (estact3.htetdesde<=" & ConvFecha(fechasta) & " AND (estact3.htethasta is null or estact3.htethasta>=" & ConvFecha(fechasta) & "))"
            If CStr(estrnro3) <> "" And CStr(estrnro3) <> "-1" Then 'cuando se le asigna un valor al nivel 3
                StrSql = StrSql & " AND estact3.estrnro =" & estrnro3
            End If
        Else
            If CStr(tenro2) <> "" And CStr(tenro2) <> "0" Then
                StrSql = StrSql & " INNER JOIN his_estructura estact1 ON h_estr.ternro = estact1.ternro  AND estact1.tenro  = " & tenro1
                StrSql = StrSql & " AND (estact1.htetdesde<=" & ConvFecha(fechasta) & " AND (estact1.htethasta is null or estact1.htethasta>=" & ConvFecha(fechasta) & "))"
                If CStr(estrnro1) <> "" And CStr(estrnro1) <> "-1" Then 'cuando se le asigna un valor al nivel 1
                    StrSql = StrSql & " AND estact1.estrnro =" & estrnro1
                End If
                
                StrSql = StrSql & " INNER JOIN his_estructura estact2 ON h_estr.ternro = estact2.ternro  AND estact2.tenro  = " & tenro2
                StrSql = StrSql & " AND (estact2.htetdesde<=" & ConvFecha(fechasta) & " AND (estact2.htethasta is null or estact2.htethasta>=" & ConvFecha(fechasta) & "))"
                If CStr(estrnro2) <> "" And CStr(estrnro2) <> "-1" Then 'cuando se le asigna un valor al nivel 2
                    StrSql = StrSql & " AND estact2.estrnro =" & estrnro2
                End If
            Else
                If CStr(tenro1) <> "" And CStr(tenro1) <> "0" Then
                    StrSql = StrSql & " INNER JOIN his_estructura estact1 ON h_estr.ternro = estact1.ternro  AND estact1.tenro  = " & tenro1
                    StrSql = StrSql & " AND (estact1.htetdesde<=" & ConvFecha(fechasta) & " AND (estact1.htethasta is null or estact1.htethasta>=" & ConvFecha(fechasta) & "))"
                    If CStr(estrnro1) <> "" And CStr(estrnro1) <> "-1" Then 'cuando se le asigna un valor al nivel 1
                        StrSql = StrSql & " AND estact1.estrnro =" & estrnro1
                    End If
                End If
            End If
        End If
            
    End If
    
    
    StrSql = StrSql & " WHERE h_estr.estrnro IN (" & lista_estrnro & ")"
    StrSql = StrSql & " ORDER BY confrep.confetiq, h_estr.ternro"
    OpenRecordset StrSql, rsConsult
    
    Flog.writeline ""
    Flog.writeline "Busco los empleados que pertenecen a las estructuras y a la empresa seleccionada en el rango de fechas"
    Flog.writeline "    SQL --> " & StrSql
    Flog.writeline ""
    
    estrnro_ant = 0
    ternro_ant = 0
    
    Progreso = 0
    cantidadProcesada = rsConsult.RecordCount
    If cantidadProcesada = 0 Then
        cantidadProcesada = 1
    End If
    IncPorc = (99 / cantidadProcesada)
    
    orden = 1
    Do Until rsConsult.EOF
    
        'Encuentro el rango de fechas en que el empleado es valido en empresa y estructura
        If CDate(rsConsult("estrfecdesde")) > CDate(rsConsult("empfecdesde")) Then
            fecdesde_empl = CDate(rsConsult!estrfecdesde)
        Else
            fecdesde_empl = CDate(rsConsult!empfecdesde)
        End If
        
        
        If EsNulo(rsConsult!estrfechasta) Or EsNulo(rsConsult("empfechasta")) Then
            If EsNulo(rsConsult!estrfechasta) And EsNulo(rsConsult!empfechasta) Then
                fechasta_empl = CDate(fechasta)
            ElseIf EsNulo(rsConsult!estrfechasta) Then
                fechasta_empl = CDate(rsConsult!empfechasta)
            Else
                fechasta_empl = CDate(rsConsult!estrfechasta)
            End If
        Else
            If CDate(rsConsult!estrfechasta) < CDate(rsConsult!empfechasta) Then
                fechasta_empl = CDate(rsConsult!estrfechasta)
            Else
                fechasta_empl = CDate(rsConsult!empfechasta)
            End If
        End If
        
        
        If fecdesde_empl <= fechasta_empl Then
           ' If (estrnro_ant <> rsConsult!estrnro) Or (ternro_ant <> rsConsult!Ternro) Then MDF
            '    total_dotacion_AR = total_dotacion_AR + 1
            '   total_dotacion_th_AR = total_dotacion_th_AR + 1
            ' End If
            
            ' Verifico que el empleado tenga gti_acumdiario AR en el rango de fechas
            StrSql = " SELECT sum(adcanthoras) canthoras, thnro "
            StrSql = StrSql & " FROM gti_acumdiario "
            StrSql = StrSql & " WHERE gti_acumdiario.ternro = " & rsConsult!Ternro
            StrSql = StrSql & " AND gti_acumdiario.thnro IN (" & lista_AR & ") "
            StrSql = StrSql & " AND adfecha >= " & ConvFecha(fecdesde) & " AND adfecha <= " & ConvFecha(fechasta)
            StrSql = StrSql & " GROUP BY thnro "
            OpenRecordset StrSql, rsacumdiario
            If Not rsacumdiario.EOF And ((ternro_ant <> rsConsult!Ternro) Or (ternro_ant <> rsConsult!Ternro)) Then '---- mdf
              total_dotacion_AR = total_dotacion_AR + 1 '---mdf
              total_dotacion_th_AR = total_dotacion_th_AR + 1 '----mdf
            End If
            Do Until rsacumdiario.EOF
                If rsacumdiario!canthoras <> 0 Then
                    Call CargarValorAR(rsacumdiario!thnro, rsacumdiario!canthoras)
                    Flog.writeline " Busco el valor de AR "
                End If
            
                rsacumdiario.MoveNext
            Loop
            
            rsacumdiario.Close
            
            
            'If (estrnro_ant <> rsConsult!estrnro) Or (ternro_ant <> rsConsult!Ternro) Then  MDF
                'total_dotacion_ANR = total_dotacion_ANR + 1
                'total_dotacion_th_ANR = total_dotacion_th_ANR + 1
            'End If
            
            ' Verifico que el empleado tenga gti_acumdiario ANR en el rango de fechas
            StrSql = " SELECT sum(adcanthoras) canthoras, thnro "
            StrSql = StrSql & " FROM gti_acumdiario "
            StrSql = StrSql & " WHERE gti_acumdiario.ternro = " & rsConsult!Ternro
            StrSql = StrSql & " AND gti_acumdiario.thnro IN (" & lista_ANR & ") "
            StrSql = StrSql & " AND adfecha >= " & ConvFecha(fecdesde) & " AND adfecha <= " & ConvFecha(fechasta)
            StrSql = StrSql & " GROUP BY thnro "
            OpenRecordset StrSql, rsacumdiario
            If Not rsacumdiario.EOF And ((ternro_ant <> rsConsult!Ternro) Or (ternro_ant <> rsConsult!Ternro)) Then '---- mdf
               total_dotacion_ANR = total_dotacion_ANR + 1
               total_dotacion_th_ANR = total_dotacion_th_ANR + 1
            End If
            
            Do Until rsacumdiario.EOF
                If rsacumdiario!canthoras <> 0 Then
                    Call CargarValorANR(rsacumdiario!thnro, rsacumdiario!canthoras)
                    Flog.writeline " Busco el valor de ANR "
                End If
            
                rsacumdiario.MoveNext
            Loop
            
            rsacumdiario.Close
            
        End If
        
        estrnro_ant = rsConsult!estrnro
        ternro_ant = rsConsult!Ternro
        confetiq_ant = rsConsult!confetiq
        
        rsConsult.MoveNext
        
        Cargar_registro = False
        If rsConsult.EOF Then
            Cargar_registro = True
        Else
            If estrnro_ant <> rsConsult!estrnro Then
                Cargar_registro = True
            End If
        End If
        
        
        If Cargar_registro Then
            ' Inserto los valores de las Ausencias Remuneradas
            For i = 0 To UBound(Arr_THoras_AR) - 1
                tothsestr = tothsestr + Arr_THoras_AR(i).toths
            Next
            
            For i = 0 To UBound(Arr_THoras_AR) - 1  'mdf insert 1
                StrSql = "INSERT INTO rep_inf_aus_det (bpronro,estrnro,estrdabr,thnro,thdesc,threm,toths,tothsestr,cantempl,orden) VALUES ("
                StrSql = StrSql & NroProceso
                StrSql = StrSql & "," & estrnro_ant
                StrSql = StrSql & ",'" & confetiq_ant & "'"
                StrSql = StrSql & ", " & Arr_THoras_AR(i).thnro
                StrSql = StrSql & ",'" & Arr_THoras_AR(i).thdesc & "'"
                StrSql = StrSql & ",-1"
                StrSql = StrSql & "," & Arr_THoras_AR(i).toths
                StrSql = StrSql & "," & tothsestr
                StrSql = StrSql & "," & total_dotacion_th_AR
                StrSql = StrSql & "," & orden & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
                
                Arr_THoras_AR(i).thdesc = ""
                Arr_THoras_AR(i).toths = 0
            Next
        
            total_dotacion_th_AR = 0 'por estructura  totaliza!!
            tothsestr = 0
            
            
            ' Inserto los valores de las Ausencias No Remuneradas
            For i = 0 To UBound(Arr_THoras_ANR) - 1
                tothsestr = tothsestr + Arr_THoras_ANR(i).toths
            Next
            
            For i = 0 To UBound(Arr_THoras_ANR) - 1 'mdf insert 2
                StrSql = "INSERT INTO rep_inf_aus_det (bpronro,estrnro,estrdabr,thnro,thdesc,threm,toths,tothsestr,cantempl,orden) VALUES ("
                StrSql = StrSql & NroProceso
                StrSql = StrSql & "," & estrnro_ant
                StrSql = StrSql & ",'" & confetiq_ant & "'"
                StrSql = StrSql & "," & Arr_THoras_ANR(i).thnro
                StrSql = StrSql & ",'" & Arr_THoras_ANR(i).thdesc & "'"
                StrSql = StrSql & ",0"
                StrSql = StrSql & "," & Arr_THoras_ANR(i).toths
                StrSql = StrSql & "," & tothsestr
                StrSql = StrSql & "," & total_dotacion_th_ANR
                StrSql = StrSql & "," & orden & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
                
                Arr_THoras_ANR(i).thdesc = ""
                Arr_THoras_ANR(i).toths = 0
            Next
        
            total_dotacion_th_ANR = 0 'por estructura  totaliza!!
            tothsestr = 0
            orden = orden + 1
            
        End If
    
        'Actualizo el progreso
        Progreso = Progreso + IncPorc
        cantidadProcesada = cantidadProcesada - 1
        
        TiempoAcumulado = GetTickCount
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
        StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'"
        StrSql = StrSql & ", bprcempleados ='" & CStr(cantidadProcesada) & "' WHERE bpronro = " & NroProceso
        objConnProgreso.Execute StrSql, , adExecuteNoRecords
        
    Loop

    rsConsult.Close
    
    ' Calculos la columna de totales para AR
    StrSql = " SELECT sum(toths) total FROM rep_inf_aus_det "
    StrSql = StrSql & " WHERE bpronro = " & NroProceso & " AND threm = -1"
    OpenRecordset StrSql, rsConsult
    If Not rsConsult.EOF Then
        total_AR = IIf(EsNulo(rsConsult!total), 0, rsConsult!total)
    End If
    rsConsult.Close
    
    
    For i = 0 To UBound(Arr_THoras_AR) - 1
        StrSql = " SELECT sum(toths) total FROM rep_inf_aus_det "
        StrSql = StrSql & " WHERE bpronro = " & NroProceso & " AND thnro = " & Arr_THoras_AR(i).thnro
        OpenRecordset StrSql, rsConsult
        
        If Not rsConsult.EOF Then 'mdf insert 3
            StrSql = "INSERT INTO rep_inf_aus_det (bpronro,estrnro,estrdabr,thnro,thdesc,threm,toths,tothsestr,cantempl,orden) VALUES ("
            StrSql = StrSql & NroProceso
            StrSql = StrSql & ",0"
            StrSql = StrSql & ",'TOTALES'"
            StrSql = StrSql & ", " & Arr_THoras_AR(i).thnro
            StrSql = StrSql & ",'" & Arr_THoras_AR(i).thdesc & "'"
            StrSql = StrSql & ",-1"
            StrSql = StrSql & "," & IIf(EsNulo(rsConsult!total), 0, rsConsult!total)
            StrSql = StrSql & "," & total_AR
            StrSql = StrSql & "," & total_dotacion_AR
            StrSql = StrSql & "," & orden & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        
        rsConsult.Close
        
    Next
    
    ' Calculos la columna de totales para ANR
    StrSql = " SELECT sum(toths) total FROM rep_inf_aus_det "
    StrSql = StrSql & " WHERE bpronro = " & NroProceso & " AND threm = 0"
    OpenRecordset StrSql, rsConsult
    If Not rsConsult.EOF Then
        total_ANR = IIf(EsNulo(rsConsult!total), 0, rsConsult!total)
    End If
    rsConsult.Close
    
    
    For i = 0 To UBound(Arr_THoras_ANR) - 1
        StrSql = " SELECT sum(toths) total FROM rep_inf_aus_det "
        StrSql = StrSql & " WHERE bpronro = " & NroProceso & " AND thnro = " & Arr_THoras_ANR(i).thnro
        OpenRecordset StrSql, rsConsult
        
        If Not rsConsult.EOF Then 'mdf insert 4
            StrSql = "INSERT INTO rep_inf_aus_det (bpronro,estrnro,estrdabr,thnro,thdesc,threm,toths,tothsestr,cantempl,orden) VALUES ("
            StrSql = StrSql & NroProceso
            StrSql = StrSql & ",0"
            StrSql = StrSql & ",'TOTALES'"
            StrSql = StrSql & ", " & Arr_THoras_ANR(i).thnro
            StrSql = StrSql & ",'" & Arr_THoras_ANR(i).thdesc & "'"
            StrSql = StrSql & ",0"
            StrSql = StrSql & "," & IIf(EsNulo(rsConsult!total), 0, rsConsult!total)
            StrSql = StrSql & "," & total_ANR
            StrSql = StrSql & "," & total_dotacion_ANR
            StrSql = StrSql & "," & orden & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        
        rsConsult.Close
        
    Next
    
    MyCommitTrans
    
Fin:
Exit Sub
            
Fin_error:
    MyRollbackTrans
    Exit Sub

MError:
    MyRollbackTrans
    Flog.writeline
    Flog.writeline "***************************************************************"
    Flog.writeline " Error: " & Err.Description
    Flog.writeline " Última Sql ejecutada: " & StrSql
    Flog.writeline "***************************************************************"
    Flog.writeline
    HuboErrores = True
    Exit Sub
End Sub
            
'----------------------------------------------------------------
' Carga los valores en los arrglos Arr_THoras_AR y Arr_THoras_ANR
'----------------------------------------------------------------
Public Function CargarValorInicial(ByVal lista_thnro As String, ByVal Aus_remuneradas As Boolean)
 Dim rs As New ADODB.Recordset
 Dim i
 
    StrSql = " SELECT tiphora.thnro, thdesc "
    StrSql = StrSql & " FROM tiphora "
    StrSql = StrSql & " WHERE tiphora.thnro IN (" & lista_thnro & ") "
    StrSql = StrSql & " ORDER BY tiphora.thnro "
    OpenRecordset StrSql, rs
    
    i = 0
    Do Until rs.EOF
        If Aus_remuneradas Then
            Arr_THoras_AR(i).thnro = rs!thnro
            Arr_THoras_AR(i).thdesc = rs!thdesc
        Else
            Arr_THoras_ANR(i).thnro = rs!thnro
            Arr_THoras_ANR(i).thdesc = rs!thdesc
        End If
            
        i = i + 1
        
        rs.MoveNext
    Loop
    
    rs.Close
            
End Function
'----------------------------------------------------------------
' Carga los valores en el arreglo Arr_THoras_AR
'----------------------------------------------------------------
Public Function CargarValorAR(ByVal thnro As Integer, ByVal canthoras As Double)
 Dim i As Integer
 Dim salir As Boolean
 
    i = 0
    salir = False
    Do Until salir Or i >= UBound(Arr_THoras_AR)
        If Arr_THoras_AR(i).thnro = thnro Then
            Arr_THoras_AR(i).toths = Arr_THoras_AR(i).toths + canthoras
            
            salir = True
        End If
        i = i + 1
    Loop
End Function

'----------------------------------------------------------------
' Carga los valores en el arreglo Arr_THoras_ANR
'----------------------------------------------------------------
Public Function CargarValorANR(ByVal thnro As Integer, ByVal canthoras As Double)
 Dim i As Integer
 Dim salir As Boolean
 
    i = 0
    salir = False
    Do Until salir Or i >= UBound(Arr_THoras_ANR)
        If Arr_THoras_ANR(i).thnro = thnro Then
            Arr_THoras_ANR(i).toths = Arr_THoras_ANR(i).toths + canthoras
            
            salir = True
        End If
        i = i + 1
    Loop
End Function



