Attribute VB_Name = "MdlF935"
Option Explicit

'Const Version = 1.01
'Const FechaVersion = "28/08/2006"

'Const Version = 1.02
'Const FechaVersion = "12/01/2007" ' Maximiliano Breglia se toco el order by para Oracle

'Const Version = 1.03
'Const FechaVersion = "18/04/2007" 'Domicilios de explotacion: Domicilio Ambulante (fijo N)
                                  'Vinculos Familiares: Default de Nro de sec, juz, escolaridad y nro acta en blanco
        
'Const Version = 1.04
'Const FechaVersion = "05/11/2009" 'Encriptacion de string de conexion - Manuel

Const Version = 1.05
Const FechaVersion = "02/12/2011" 'se agregaron 3 columnas a repf_935 convenio,categoria,condicion sijp - Sebastian Stremel
'----------------------------------------------------------------

Const tipoCod = 100

Global Nroliq As Long
Global Empresa As Long
Global fechaBaja As Date
Global ArrPar


'Autor  = Martin Ferraro

'----------------------------------------------------------


Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial del Generador de Reporte F935.
' Autor      : Martin Ferraro
' Fecha      : 29/08/2006
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
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    
    On Error Resume Next
    OpenConnection strconexion, objconnProgreso
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    On Error GoTo 0
    
    On Error GoTo ME_Main
    
    Nombre_Arch = PathFLog & "Reporte_F935" & "-" & NroProcesoBatch & ".log"
    
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
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 154 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        bprcparam = rs_batch_proceso!bprcparam
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        Call F935(NroProcesoBatch, bprcparam)
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


Public Sub F935(ByVal bpronro As Long, ByVal Parametros As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del reporte del F935
' Autor      : Martin Ferraro
' Fecha      : 28/08/2006
' --------------------------------------------------------------------------------------------


Dim pos1 As Integer
Dim pos2 As Integer
Dim Desde As Date
Dim Hasta As Date

Dim rs_Empleados As New ADODB.Recordset
Dim rs_Periodo As New ADODB.Recordset
Dim rs_Confrep As New ADODB.Recordset
Dim rs_Consult As New ADODB.Recordset
Dim rs_Estructura As New ADODB.Recordset
Dim rs_Estr_cod As New ADODB.Recordset

Dim ternro As Long
Dim Cuil As String
Dim Legajo As Long
Dim terape As String
Dim ternom As String
Dim agropecuario As String
Dim contrato As String
Dim puesto As String
Dim Osocial As String
Dim sucursal As String
Dim actEmpresa As String
Dim fechaNac As String
Dim discapacitado As String
Dim fechaFaseInicio As String
Dim fechaFasefin As String
Dim fechaRenuncia As String
Dim sitRevista As String
Dim formaLiq As String
Dim nivEst As String
Dim encontroRemu As Boolean
Dim EsConcRemu As Boolean
Dim vNumRemu As Long
Dim vStringRemu As String
Dim Remu As Double
Dim RemuStr As String
Dim mail As String
Dim tipoMail As String
Dim calle As String
Dim torre As String
Dim nroCalle As String
Dim oficdepto As String
Dim codigopostal As String
Dim piso As String
Dim provincia As String
Dim localidad As String
Dim telefono As String
Dim tipoTel As String
Dim cbu As String
Dim codMov As String

'variables agregadas 02/12/2011 sebastian stremel
Dim vConvenio As String
Dim convenio As String
Dim vCategoria As String
Dim categoria As String
Dim vtiposervicio As String
Dim tiposervicio  As String
'Inicio codigo ejecutable
On Error GoTo E_F935

' Levanto cada parametro por separado, el separador de parametros es "."
Flog.writeline "levantando parametros" & Parametros
If Not IsNull(Parametros) Then

    If Len(Parametros) >= 1 Then
    
        ArrPar = Split(Parametros, ".")
        If UBound(ArrPar) = 2 Then
            Nroliq = CLng(ArrPar(0))
            Flog.writeline Espacios(Tabulador * 0) & "Parametro Periodo = " & Nroliq
            
            Empresa = CLng(ArrPar(1))
            Flog.writeline Espacios(Tabulador * 0) & "Parametro Empresa = " & Empresa
            
            fechaBaja = CDate(ArrPar(2))
            Flog.writeline Espacios(Tabulador * 0) & "Parametro Fecha Baja = " & fechaBaja
        Else
            Flog.writeline Espacios(Tabulador * 0) & "ERROR. Numero de parametros erroneo."
            Exit Sub
        End If
        
    End If
Else
    Flog.writeline Espacios(Tabulador * 0) & "ERROR. No se encontraron los parametros."
    Exit Sub
End If


'Configuracion del Reporte
Flog.writeline Espacios(Tabulador * 0) & "Buscando configuracion del reporte."
StrSql = "SELECT * FROM confrep WHERE repnro = 170 ORDER BY confnrocol"
OpenRecordset StrSql, rs_Confrep
If rs_Confrep.EOF Then
    Flog.writeline Espacios(Tabulador * 0) & "No se encontró la configuración del Reporte."
    Exit Sub
Else
    encontroRemu = False
    Do Until rs_Confrep.EOF
        
        Select Case rs_Confrep!confnrocol
            Case 2
                
                Select Case rs_Confrep!conftipo
                    Case "CO"
                        Flog.writeline Espacios(Tabulador * 0) & "Columna 2 Tipo Concepto."
                        encontroRemu = True
                        EsConcRemu = True
                        If EsNulo(rs_Confrep!confval2) Then
                            Flog.writeline Espacios(Tabulador * 0) & "Columna 2 No se cargo el valor alfanumerico del Concepto."
                            Exit Sub
                        End If
                        vStringRemu = rs_Confrep!confval2
                    
                    Case "AC"
                        Flog.writeline Espacios(Tabulador * 0) & "Columna 2 Tipo Acumulador."
                        encontroRemu = True
                        EsConcRemu = False
                        vNumRemu = rs_Confrep!confval
                        
                                       
                    Case Else
                        Flog.writeline Espacios(Tabulador * 0) & "Error en tipo columna 2."
                        Exit Sub
                    End Select
            
            Case 3
                Select Case rs_Confrep!conftipo
                    Case "TE"
                        Flog.writeline Espacios(Tabulador * 0) & "Columna 3 Tipo Estructura."
                        If EsNulo(rs_Confrep!confval2) Then
                            Flog.writeline Espacios(Tabulador * 0) & "Columna 3 No se cargo el valor alfanumerico de la estructura."
                            Exit Sub
                        End If
                        vConvenio = rs_Confrep!confval2
                    Case Else
                        Flog.writeline Espacios(Tabulador * 0) & "Error en tipo columna 3."
                        Exit Sub
                    End Select
            
            Case 4
                Select Case rs_Confrep!conftipo
                    Case "TE"
                        Flog.writeline Espacios(Tabulador * 0) & "Columna 4 Tipo Estructura."
                        If EsNulo(rs_Confrep!confval2) Then
                            Flog.writeline Espacios(Tabulador * 0) & "Columna 4 No se cargo el valor alfanumerico de la estructura."
                            Exit Sub
                        End If
                        vCategoria = rs_Confrep!confval2
                    Case Else
                        Flog.writeline Espacios(Tabulador * 0) & "Error en tipo columna 4."
                        Exit Sub
                    End Select
            Case 5
                Select Case rs_Confrep!conftipo
                    Case "TE"
                        Flog.writeline Espacios(Tabulador * 0) & "Columna 5 Tipo Estructura."
                        If EsNulo(rs_Confrep!confval2) Then
                            Flog.writeline Espacios(Tabulador * 0) & "Columna 5 No se cargo el valor alfanumerico de la estructura."
                            Exit Sub
                        End If
                        vtiposervicio = rs_Confrep!confval2
                    Case Else
                        Flog.writeline Espacios(Tabulador * 0) & "Error en tipo columna 5."
                        Exit Sub
                    End Select
        End Select
        
        rs_Confrep.MoveNext
        
    Loop
End If
rs_Confrep.Close

If Not encontroRemu Then
    Flog.writeline Espacios(Tabulador * 0) & "No se encontró la columna 2 de la configuración del Reporte."
    Exit Sub
End If


'cargo el periodo
Flog.writeline Espacios(Tabulador * 0) & "Busco el periodo de liquidacion"
StrSql = "SELECT * FROM periodo WHERE pliqnro = " & CStr(Nroliq)
OpenRecordset StrSql, rs_Periodo
If rs_Periodo.EOF Then
    Flog.writeline Espacios(Tabulador * 0) & "ERROR. No se encontró el Periodo."
    Exit Sub
End If

Desde = rs_Periodo!pliqdesde
Hasta = rs_Periodo!pliqhasta

Flog.writeline Espacios(Tabulador * 1) & "Periodo: " & rs_Periodo!pliqdesc
Flog.writeline Espacios(Tabulador * 1) & "Fecha Desde: " & Desde
Flog.writeline Espacios(Tabulador * 1) & "Fecha Hasta: " & Hasta
Flog.writeline

'Borro todos los datos para el mismo periodo y empresa (solo puede existir 1 reporte para mes y empresa)
StrSql = " DELETE rep_f935 WHERE pliqnro = " & Nroliq & " AND empresa = " & Empresa
objConn.Execute StrSql, , adExecuteNoRecords

StrSql = " DELETE rep_f935_dom WHERE pliqnro = " & Nroliq & " AND empresa = " & Empresa
objConn.Execute StrSql, , adExecuteNoRecords

StrSql = " DELETE rep_f935_fam WHERE pliqnro = " & Nroliq & " AND empresa = " & Empresa
objConn.Execute StrSql, , adExecuteNoRecords


'Empleados dados de baja a partir la fecha hasta del filtro
StrSql = "SELECT distinct empleado.ternro, empleado.empleg, empleado.ternom, empleado.ternom2, empleado.terape, empleado.terape2 "
StrSql = StrSql & " FROM empleado "
StrSql = StrSql & " INNER JOIN his_estructura ON empleado.ternro = his_estructura.ternro AND his_estructura.estrnro= " & Empresa & " AND htetdesde <= " & ConvFecha(Hasta) & " AND (htethasta IS NULL OR htethasta >= " & ConvFecha(Hasta) & ") AND his_estructura.tenro=10 "
StrSql = StrSql & " INNER JOIN fases ON fases.empleado = empleado.ternro"
StrSql = StrSql & " AND fases.real = -1 "
StrSql = StrSql & " AND (fases.bajfec >= " & ConvFecha(fechaBaja) & ")"
StrSql = StrSql & " WHERE empleado.empest <> -1 "
StrSql = StrSql & " UNION "
'Empleados activos
StrSql = StrSql & " SELECT distinct empleado.ternro, empleado.empleg, empleado.ternom, empleado.ternom2, empleado.terape, empleado.terape2 "
StrSql = StrSql & " FROM empleado "
StrSql = StrSql & " INNER JOIN his_estructura ON empleado.ternro = his_estructura.ternro AND his_estructura.estrnro= " & Empresa & " AND htetdesde <= " & ConvFecha(Hasta) & " AND (htethasta IS NULL OR htethasta >= " & ConvFecha(Hasta) & ") AND his_estructura.tenro=10 "
StrSql = StrSql & " WHERE empleado.empest = -1 "
'StrSql = StrSql & " ORDER BY empleado.empleg "  Para que no de error en oracle
StrSql = StrSql & " ORDER BY 2 "


OpenRecordset StrSql, rs_Empleados

'seteo de las variables de progreso
Progreso = 0
CEmpleadosAProc = rs_Empleados.RecordCount
If CEmpleadosAProc = 0 Then
    Flog.writeline Espacios(Tabulador * 0) & "No hay empleados"
    CEmpleadosAProc = 1
Else
    Flog.writeline Espacios(Tabulador * 0) & "Cantidad de Empleados: " & CEmpleadosAProc
End If
IncPorc = (100 / CEmpleadosAProc)
        
        
        
Flog.writeline Espacios(Tabulador * 0) & "-------------------------------------------------------------------"
Flog.writeline Espacios(Tabulador * 0) & "Comienza el Procesamiento de Empleados"
Flog.writeline Espacios(Tabulador * 0) & "-------------------------------------------------------------------"
Do While Not rs_Empleados.EOF
    
    ternro = rs_Empleados!ternro
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "PROCESANDO EMPLEADO: " & rs_Empleados!empleg & " - " & rs_Empleados!terape & " " & rs_Empleados!ternom
    Flog.writeline Espacios(Tabulador * 0) & "-------------------------------------------------------------------"
    
    
    Legajo = rs_Empleados!empleg
    terape = IIf(EsNulo(rs_Empleados!terape), "", rs_Empleados!terape) & IIf(EsNulo(rs_Empleados!terape2), "", " " & rs_Empleados!terape2)
    ternom = IIf(EsNulo(rs_Empleados!ternom), "", rs_Empleados!ternom) & IIf(EsNulo(rs_Empleados!ternom2), "", " " & rs_Empleados!ternom2)
    
    
    codMov = "AT"
        
    '---------------------------------------------------------------------------------------------------------------
    'DATOS DEL EMPLEADO---------------------------------------------------------------------------------------------
    '----------------------------------------------------------------------------------------------------------------
    Flog.writeline Espacios(Tabulador * 1) & "Buscando Datos del empleado"
    StrSql = "SELECT * "
    StrSql = StrSql & " From tercero "
    StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = tercero.ternro "
    StrSql = StrSql & " WHERE tercero.ternro = " & ternro
    OpenRecordset StrSql, rs_Consult
    If Not rs_Consult.EOF Then
        fechaNac = rs_Consult!terfecnac
        discapacitado = IIf(rs_Consult!empdiscap = -1, "S", "N")
        mail = IIf(EsNulo(rs_Consult!empemail), "", Left(rs_Consult!empemail, 60))
        tipoMail = IIf(EsNulo(rs_Consult!empemail), "", "2")
    Else
        Flog.writeline Espacios(Tabulador * 2) & "No se encontraron datos del empleado"
        fechaNac = ""
        discapacitado = "N"
        mail = ""
        tipoMail = ""
    End If
    rs_Consult.Close
        
        
    '---------------------------------------------------------------------------------------------------------------
    'CUIL-----------------------------------------------------------------------------------------------------------
    '---------------------------------------------------------------------------------------------------------------
    Flog.writeline Espacios(Tabulador * 1) & "Buscando CUIL"
    StrSql = " SELECT nrodoc "
    StrSql = StrSql & " FROM ter_doc "
    StrSql = StrSql & " WHERE ter_doc.ternro= " & ternro
    StrSql = StrSql & " AND ter_doc.tidnro=10 "
    OpenRecordset StrSql, rs_Consult
    If Not rs_Consult.EOF Then
        Cuil = IIf(EsNulo(rs_Consult!nrodoc), "", rs_Consult!nrodoc)
        Flog.writeline Espacios(Tabulador * 2) & "CUIL = " & Cuil
    Else
        Cuil = ""
        Flog.writeline Espacios(Tabulador * 2) & "CUIL no encontrado"
    End If
    rs_Consult.Close
    
    Cuil = Replace(Cuil, "-", "")
    Cuil = Replace(Cuil, "/", "")
    Cuil = Left(Cuil, 11)
        
        
        
        
        
    'desde aca
    '---------------------------------------------------------------------------------------------------------------
    'CONVENIO-------------------------------------------------------------------------------------------------------
    '---------------------------------------------------------------------------------------------------------------
    Flog.writeline Espacios(Tabulador * 1) & "Buscando Convenio"
    StrSql = " SELECT estructura.estrnro, estructura.estrdabr "
    StrSql = StrSql & " FROM his_estructura "
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
    StrSql = StrSql & " WHERE ternro = " & ternro & " AND "
    StrSql = StrSql & " his_estructura.tenro = '" & vConvenio & "' AND "   'Convenio
    StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(Hasta) & ") AND "
    StrSql = StrSql & " ((" & ConvFecha(Hasta) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
    OpenRecordset StrSql, rs_Estructura
    If Not rs_Estructura.EOF Then
        Flog.writeline Espacios(Tabulador * 2) & "Convenio: " & rs_Estructura!Estrnro & " - " & rs_Estructura!estrdabr
        StrSql = "SELECT * FROM estr_cod WHERE estrnro =" & rs_Estructura!Estrnro
        StrSql = StrSql & " AND tcodnro = " & tipoCod
        OpenRecordset StrSql, rs_Estr_cod
        If Not rs_Estr_cod.EOF Then
            convenio = IIf(EsNulo(rs_Estr_cod!nrocod), "", Left(CStr(rs_Estr_cod!nrocod), 10))
        Else
            Flog.writeline Espacios(Tabulador * 2) & "No se encontró el codigo para el convenio."
            convenio = ""
        End If
        rs_Estr_cod.Close
    Else
        Flog.writeline Espacios(Tabulador * 2) & "No se encontró el Tipo de convenio."
        convenio = ""
    End If
    rs_Estructura.Close
    '---------------------------------------------------------------------------------------------------------------
    
    
    
    'desde aca
    '---------------------------------------------------------------------------------------------------------------
    'CATEGORIA-------------------------------------------------------------------------------------------------------
    '---------------------------------------------------------------------------------------------------------------
    Flog.writeline Espacios(Tabulador * 1) & "Buscando Convenio"
    StrSql = " SELECT estructura.estrnro, estructura.estrdabr "
    StrSql = StrSql & " FROM his_estructura "
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
    StrSql = StrSql & " WHERE ternro = " & ternro & " AND "
    StrSql = StrSql & " his_estructura.tenro = '" & vCategoria & "' AND "   'Categoria
    StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(Hasta) & ") AND "
    StrSql = StrSql & " ((" & ConvFecha(Hasta) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
    OpenRecordset StrSql, rs_Estructura
    If Not rs_Estructura.EOF Then
        Flog.writeline Espacios(Tabulador * 2) & "Convenio: " & rs_Estructura!Estrnro & " - " & rs_Estructura!estrdabr
        StrSql = "SELECT * FROM estr_cod WHERE estrnro =" & rs_Estructura!Estrnro
        StrSql = StrSql & " AND tcodnro = " & tipoCod
        OpenRecordset StrSql, rs_Estr_cod
        If Not rs_Estr_cod.EOF Then
            categoria = IIf(EsNulo(rs_Estr_cod!nrocod), "", Left(CStr(rs_Estr_cod!nrocod), 6))
        Else
            Flog.writeline Espacios(Tabulador * 2) & "No se encontró el codigo para la categoria."
            categoria = ""
        End If
        rs_Estr_cod.Close
    Else
        Flog.writeline Espacios(Tabulador * 2) & "No se encontró el Tipo de categoria."
        categoria = ""
    End If
    rs_Estructura.Close
    '---------------------------------------------------------------------------------------------------------------
    
    'desde aca
    '---------------------------------------------------------------------------------------------------------------
    'tiposervicio -------------------------------------------------------------------------------------------------------
    '---------------------------------------------------------------------------------------------------------------
    Flog.writeline Espacios(Tabulador * 1) & "Buscando Convenio"
    StrSql = " SELECT estructura.estrnro, estructura.estrdabr "
    StrSql = StrSql & " FROM his_estructura "
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
    StrSql = StrSql & " WHERE ternro = " & ternro & " AND "
    StrSql = StrSql & " his_estructura.tenro = '" & vtiposervicio & "' AND "   'tipo de servicio
    StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(Hasta) & ") AND "
    StrSql = StrSql & " ((" & ConvFecha(Hasta) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
    OpenRecordset StrSql, rs_Estructura
    If Not rs_Estructura.EOF Then
        Flog.writeline Espacios(Tabulador * 2) & "Convenio: " & rs_Estructura!Estrnro & " - " & rs_Estructura!estrdabr
        StrSql = "SELECT * FROM estr_cod WHERE estrnro =" & rs_Estructura!Estrnro
        StrSql = StrSql & " AND tcodnro = " & tipoCod
        OpenRecordset StrSql, rs_Estr_cod
        If Not rs_Estr_cod.EOF Then
            tiposervicio = IIf(EsNulo(rs_Estr_cod!nrocod), "", Left(CStr(rs_Estr_cod!nrocod), 3))
        Else
            Flog.writeline Espacios(Tabulador * 2) & "No se encontró el codigo para el tipo de servicio.(condicion sijp)"
            tiposervicio = ""
        End If
        rs_Estr_cod.Close
    Else
        Flog.writeline Espacios(Tabulador * 2) & "No se encontró el Tipo de servicio.(condicion sijp)"
        tiposervicio = ""
    End If
    rs_Estructura.Close
    '---------------------------------------------------------------------------------------------------------------
    
    
    '---------------------------------------------------------------------------------------------------------------
    'SUCURSAL-------------------------------------------------------------------------------------------------------
    '---------------------------------------------------------------------------------------------------------------
    Flog.writeline Espacios(Tabulador * 1) & "Buscando Sucursal"
    StrSql = " SELECT estructura.estrnro, estructura.estrdabr, estructura.estrcodext "
    StrSql = StrSql & " FROM his_estructura "
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
    StrSql = StrSql & " WHERE ternro = " & ternro & " AND "
    StrSql = StrSql & " his_estructura.tenro = 1 AND " 'Sucursal
    StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(Hasta) & ") AND "
    StrSql = StrSql & " ((" & ConvFecha(Hasta) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
    OpenRecordset StrSql, rs_Estructura
    If Not rs_Estructura.EOF Then
        Flog.writeline Espacios(Tabulador * 2) & "Sucursal: " & rs_Estructura!Estrnro & " - " & rs_Estructura!estrdabr
        sucursal = IIf(EsNulo(rs_Estructura!estrcodext), "", Left(rs_Estructura!estrcodext, 5))
        StrSql = "SELECT * FROM estr_cod WHERE estrnro =" & rs_Estructura!Estrnro
        StrSql = StrSql & " AND tcodnro = " & tipoCod
        OpenRecordset StrSql, rs_Estr_cod
        If Not rs_Estr_cod.EOF Then
            actEmpresa = IIf(EsNulo(rs_Estr_cod!nrocod), "", Left(CStr(rs_Estr_cod!nrocod), 6))
        Else
            Flog.writeline Espacios(Tabulador * 2) & "No se encontró el codigo para el Tipo de Sucursal."
            actEmpresa = "0"
        End If
        rs_Estr_cod.Close
    Else
        Flog.writeline Espacios(Tabulador * 2) & "No se encontró el Tipo de Sucursal."
        actEmpresa = "0"
        sucursal = ""
    End If
    rs_Estructura.Close
    
    
    fechaRenuncia = ""
    '---------------------------------------------------------------------------------------------------------------
    'FASE Y SITUCAION DE REVISTA------------------------------------------------------------------------------------
    '---------------------------------------------------------------------------------------------------------------
    Flog.writeline Espacios(Tabulador * 1) & "Buscando Fases"
    StrSql = "SELECT * "
    StrSql = StrSql & " From fases "
    StrSql = StrSql & " Where fases.Empleado = " & ternro
    StrSql = StrSql & " AND fases.real = -1 "
    StrSql = StrSql & " ORDER BY fases.altfec Desc "
    OpenRecordset StrSql, rs_Consult
    If Not rs_Consult.EOF Then
        fechaFaseInicio = rs_Consult!altfec
        fechaFasefin = IIf(EsNulo(rs_Consult!bajfec), "", rs_Consult!bajfec)
        Flog.writeline Espacios(Tabulador * 2) & "Fase encontrada"
        
        If EsNulo(rs_Consult!bajfec) Then
            Flog.writeline Espacios(Tabulador * 2) & "La fase NO es de baja."
            sitRevista = ""
        Else
            Flog.writeline Espacios(Tabulador * 2) & "La fase es de baja, Busco Situacion de revista."
            'Busco la estructura situacion de revista cuya fecha de alta coincida con la
            'fecha de baja de la fase encontrada
            codMov = "BT"
            StrSql = " SELECT estructura.estrnro, estructura.estrdabr "
            StrSql = StrSql & " FROM his_estructura "
            StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
            StrSql = StrSql & " WHERE ternro = " & ternro & " AND "
            StrSql = StrSql & " his_estructura.tenro = 30 AND " 'Situcion de revista
            StrSql = StrSql & " (his_estructura.htetdesde = " & ConvFecha(rs_Consult!bajfec) & ") "
            OpenRecordset StrSql, rs_Estructura
            If Not rs_Estructura.EOF Then
                Flog.writeline Espacios(Tabulador * 2) & "Situacion de revista: " & rs_Estructura!Estrnro & " - " & rs_Estructura!estrdabr
                StrSql = "SELECT * FROM estr_cod WHERE estrnro =" & rs_Estructura!Estrnro
                StrSql = StrSql & " AND tcodnro = " & tipoCod
                OpenRecordset StrSql, rs_Estr_cod
                If Not rs_Estr_cod.EOF Then
                    sitRevista = IIf(EsNulo(rs_Estr_cod!nrocod), "", Left(CStr(rs_Estr_cod!nrocod), 2))
                    'Si es renuncia guardo la fecha
                    If sitRevista = 21 Then
                        fechaRenuncia = CDate(DateAdd("d", -1, CDate(fechaFasefin)))
                    End If
                Else
                    Flog.writeline Espacios(Tabulador * 2) & "No se encontró el codigo para el Tipo de situacion de revista."
                    sitRevista = ""
                End If
                rs_Estr_cod.Close
            Else
                Flog.writeline Espacios(Tabulador * 2) & "No se encontró situacion de revista."
                sitRevista = ""
            End If
            rs_Estructura.Close
        End If
    Else
        fechaFaseInicio = ""
        fechaFasefin = ""
        sitRevista = ""
        Flog.writeline Espacios(Tabulador * 2) & "Fase no encontrada"
    End If
    rs_Consult.Close
    
    
    '---------------------------------------------------------------------------------------------------------------
    'TRABAJADOR AGROPECUARIO----------------------------------------------------------------------------------------
    '---------------------------------------------------------------------------------------------------------------
    Flog.writeline Espacios(Tabulador * 1) & "Buscando agropecuario"
    StrSql = " SELECT estructura.estrnro, estructura.estrdabr "
    StrSql = StrSql & " FROM his_estructura "
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
    StrSql = StrSql & " WHERE his_estructura.ternro = " & ternro & " AND "
    StrSql = StrSql & " his_estructura.tenro = 29 AND " 'Actividad
    StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(Hasta) & ") AND "
    StrSql = StrSql & " ((" & ConvFecha(Hasta) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
    StrSql = StrSql & " ORDER BY his_estructura.htetdesde"
    OpenRecordset StrSql, rs_Estructura
    If Not rs_Estructura.EOF Then
        Flog.writeline Espacios(Tabulador * 2) & "Actividad: " & rs_Estructura!Estrnro & " - " & rs_Estructura!estrdabr
        StrSql = "SELECT * FROM estr_cod WHERE estrnro =" & rs_Estructura!Estrnro
        StrSql = StrSql & " AND tcodnro = " & tipoCod
        OpenRecordset StrSql, rs_Estr_cod
        If Not rs_Estr_cod.EOF Then
            If ((rs_Estr_cod!nrocod = "31") Or (rs_Estr_cod!nrocod = "97") Or (rs_Estr_cod!nrocod = "98")) Then
                agropecuario = "S"
                Flog.writeline Espacios(Tabulador * 2) & "Actividad Agropecuaria codigo " & rs_Estr_cod!nrocod
            Else
                agropecuario = "N"
                Flog.writeline Espacios(Tabulador * 2) & "El codigo " & rs_Estr_cod!nrocod & " no es agropecuario."
            End If
        Else
            agropecuario = "N"
            Flog.writeline Espacios(Tabulador * 2) & "No se encontró el codigo para la Actividad"
        End If
        rs_Estr_cod.Close
    Else
        agropecuario = "N"
        Flog.writeline Espacios(Tabulador * 2) & "Estructura Actividad No encontrada."
    End If
    rs_Estructura.Close
    
    
    '---------------------------------------------------------------------------------------------------------------
    'CONTRATO ACTUAL------------------------------------------------------------------------------------------------
    '---------------------------------------------------------------------------------------------------------------
    Flog.writeline Espacios(Tabulador * 1) & "Buscando Contrato"
    StrSql = " SELECT estructura.estrnro, estructura.estrdabr "
    StrSql = StrSql & " FROM his_estructura "
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
    StrSql = StrSql & " WHERE ternro = " & ternro & " AND "
    StrSql = StrSql & " his_estructura.tenro = 18 AND " 'Contrato
    StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(Hasta) & ") AND "
    StrSql = StrSql & " ((" & ConvFecha(Hasta) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
    OpenRecordset StrSql, rs_Estructura
    If Not rs_Estructura.EOF Then
        Flog.writeline Espacios(Tabulador * 2) & "Contrato: " & rs_Estructura!Estrnro & " - " & rs_Estructura!estrdabr
        StrSql = "SELECT * FROM estr_cod WHERE estrnro =" & rs_Estructura!Estrnro
        StrSql = StrSql & " AND tcodnro = " & tipoCod
        OpenRecordset StrSql, rs_Estr_cod
        If Not rs_Estr_cod.EOF Then
            contrato = IIf(EsNulo(rs_Estr_cod!nrocod), "", Left(CStr(rs_Estr_cod!nrocod), 3))
        Else
            Flog.writeline Espacios(Tabulador * 2) & "No se encontró el codigo para el Tipo de Contrato."
            contrato = ""
        End If
        rs_Estr_cod.Close
    Else
        Flog.writeline Espacios(Tabulador * 2) & "No se encontró el Tipo de Contrato."
        contrato = ""
    End If
    rs_Estructura.Close
    
    
    '---------------------------------------------------------------------------------------------------------------
    'Obra Social----------------------------------------------------------------------------------------------------
    '---------------------------------------------------------------------------------------------------------------
    Flog.writeline Espacios(Tabulador * 1) & "Buscando Obra Social"
    StrSql = " SELECT estructura.estrnro, estructura.estrdabr "
    StrSql = StrSql & " FROM his_estructura "
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
    StrSql = StrSql & " WHERE ternro = " & ternro & " AND "
    StrSql = StrSql & " his_estructura.tenro = 17 AND " 'Obra Social
    StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(Hasta) & ") AND "
    StrSql = StrSql & " ((" & ConvFecha(Hasta) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
    OpenRecordset StrSql, rs_Estructura
    If Not rs_Estructura.EOF Then
        Flog.writeline Espacios(Tabulador * 2) & "Obra Social: " & rs_Estructura!Estrnro & " - " & rs_Estructura!estrdabr
        StrSql = "SELECT * FROM estr_cod WHERE estrnro =" & rs_Estructura!Estrnro
        StrSql = StrSql & " AND tcodnro = " & tipoCod
        OpenRecordset StrSql, rs_Estr_cod
        If Not rs_Estr_cod.EOF Then
            Osocial = IIf(EsNulo(rs_Estr_cod!nrocod), "000000", Left(CStr(rs_Estr_cod!nrocod), 6))
        Else
            Flog.writeline Espacios(Tabulador * 2) & "No se encontró el codigo para el Tipo de Obra Social."
            Osocial = ""
        End If
        rs_Estr_cod.Close
    Else
        Flog.writeline Espacios(Tabulador * 2) & "No se encontró el Tipo de Obra Social."
        Osocial = "000000"
    End If
    rs_Estructura.Close
    
    
    '---------------------------------------------------------------------------------------------------------------
    'FORMA DE LIQUIDACION-------------------------------------------------------------------------------------------
    '---------------------------------------------------------------------------------------------------------------
    Flog.writeline Espacios(Tabulador * 1) & "Buscando Forma de liqudacion"
    StrSql = " SELECT estructura.estrnro, estructura.estrdabr "
    StrSql = StrSql & " FROM his_estructura "
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
    StrSql = StrSql & " WHERE ternro = " & ternro & " AND "
    StrSql = StrSql & " his_estructura.tenro = 22 AND " 'Forma de liquidacion
    StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(Hasta) & ") AND "
    StrSql = StrSql & " ((" & ConvFecha(Hasta) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
    OpenRecordset StrSql, rs_Estructura
    If Not rs_Estructura.EOF Then
        Flog.writeline Espacios(Tabulador * 2) & "Forma de liqudacion: " & rs_Estructura!Estrnro & " - " & rs_Estructura!estrdabr
        StrSql = "SELECT * FROM estr_cod WHERE estrnro =" & rs_Estructura!Estrnro
        StrSql = StrSql & " AND tcodnro = " & tipoCod
        OpenRecordset StrSql, rs_Estr_cod
        If Not rs_Estr_cod.EOF Then
            formaLiq = IIf(EsNulo(rs_Estr_cod!nrocod), "", Left(CStr(rs_Estr_cod!nrocod), 1))
        Else
            Flog.writeline Espacios(Tabulador * 2) & "No se encontró el codigo para el Tipo de Forma de liqudacion."
            formaLiq = ""
        End If
        rs_Estr_cod.Close
    Else
        Flog.writeline Espacios(Tabulador * 2) & "No se encontró el Tipo de Forma de liqudacion."
        formaLiq = ""
    End If
    rs_Estructura.Close
    
    
    '---------------------------------------------------------------------------------------------------------------
    'PUESTO---------------------------------------------------------------------------------------------------------
    '---------------------------------------------------------------------------------------------------------------
    Flog.writeline Espacios(Tabulador * 1) & "Buscando Puesto"
    StrSql = " SELECT estructura.estrnro, estructura.estrdabr "
    StrSql = StrSql & " FROM his_estructura "
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
    StrSql = StrSql & " WHERE ternro = " & ternro & " AND "
    StrSql = StrSql & " his_estructura.tenro = 4 AND " 'Puesto
    StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(Hasta) & ") AND "
    StrSql = StrSql & " ((" & ConvFecha(Hasta) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
    OpenRecordset StrSql, rs_Estructura
    If Not rs_Estructura.EOF Then
        Flog.writeline Espacios(Tabulador * 2) & "Puesto: " & rs_Estructura!Estrnro & " - " & rs_Estructura!estrdabr
        StrSql = "SELECT * FROM estr_cod WHERE estrnro =" & rs_Estructura!Estrnro
        StrSql = StrSql & " AND tcodnro = " & tipoCod
        OpenRecordset StrSql, rs_Estr_cod
        If Not rs_Estr_cod.EOF Then
            puesto = IIf(EsNulo(rs_Estr_cod!nrocod), "", Left(CStr(rs_Estr_cod!nrocod), 4))
        Else
            Flog.writeline Espacios(Tabulador * 2) & "No se encontró el codigo para el Tipo de Puesto."
            puesto = ""
        End If
        rs_Estr_cod.Close
    Else
        Flog.writeline Espacios(Tabulador * 2) & "No se encontró el Tipo de Puesto."
        puesto = ""
    End If
    rs_Estructura.Close
    
    
    '---------------------------------------------------------------------------------------------------------------
    'REMUNERACION---------------------------------------------------------------------------------------------------
    '---------------------------------------------------------------------------------------------------------------
    Flog.writeline Espacios(Tabulador * 1) & "Buscando Remuneracion"
    Remu = 0
    If EsConcRemu Then
        StrSql = " SELECT detliq.dlimonto"
        StrSql = StrSql & " From Proceso"
        StrSql = StrSql & " INNER JOIN cabliq ON cabliq.pronro = proceso.pronro"
        StrSql = StrSql & " AND cabliq.empleado = " & ternro
        StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro"
        StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro"
        StrSql = StrSql & " AND concepto.conccod = '" & vStringRemu & "'"
        StrSql = StrSql & " Where Proceso.pliqnro = " & Nroliq
        OpenRecordset StrSql, rs_Consult
        Do While Not rs_Consult.EOF
           Remu = Remu + rs_Consult!dlimonto
           rs_Consult.MoveNext
        Loop
        rs_Consult.Close
    Else
        StrSql = " SELECT acu_liq.almonto"
        StrSql = StrSql & " From Proceso"
        StrSql = StrSql & " INNER JOIN cabliq ON cabliq.pronro = proceso.pronro"
        StrSql = StrSql & " AND cabliq.empleado = " & ternro
        StrSql = StrSql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro"
        StrSql = StrSql & " AND acu_liq.acunro = " & vNumRemu
        StrSql = StrSql & " Where Proceso.pliqnro = " & Nroliq
        OpenRecordset StrSql, rs_Consult
        Do While Not rs_Consult.EOF
            Remu = Remu + rs_Consult!almonto
            rs_Consult.MoveNext
        Loop
        rs_Consult.Close
    End If
    Flog.writeline Espacios(Tabulador * 2) & "Remuneracion: " & Remu
    
    RemuStr = Format(Remu, "###########0.00")
    
    
    'Si no es baja entonces busco datos de datos complementarios, CBU de las cuentas sueldos y familiares
    If codMov <> "BT" Then
        '---------------------------------------------------------------------------------------------------------------
        'NIVEL DE FORMACION---------------------------------------------------------------------------------------------
        '---------------------------------------------------------------------------------------------------------------
        Flog.writeline Espacios(Tabulador * 1) & "Buscando Nivel de Estudio"
        StrSql = " SELECT nivest.nivnro, nivcodext, nivest.nivdesc, cap_estformal.capfechas "
        StrSql = StrSql & " FROM cap_estformal "
        StrSql = StrSql & " INNER JOIN nivest ON cap_estformal.nivnro = nivest.nivnro "
        StrSql = StrSql & " WHERE cap_estformal.ternro = " & ternro
        'StrSql = StrSql & " AND cap_estformal.capcomp = -1 "
        StrSql = StrSql & " AND cap_estformal.capfechas is not null "
        StrSql = StrSql & " ORDER BY cap_estformal.capfechas DESC "
        OpenRecordset StrSql, rs_Consult
        If Not rs_Consult.EOF Then
            Flog.writeline Espacios(Tabulador * 2) & "Nivel de estudio " & rs_Consult!nivnro & " " & rs_Consult!nivdesc
            nivEst = Left(calcularMapeo(rs_Consult!nivnro, 1, "01"), 2)
        Else
            Flog.writeline Espacios(Tabulador * 2) & "No se encontro el nivel de estudio "
            nivEst = "01"
        End If
        rs_Consult.Close
        
        
        '---------------------------------------------------------------------------------------------------------------
        'DATOS DE DOMICILIO Y TELEFONO----------------------------------------------------------------------------------
        '---------------------------------------------------------------------------------------------------------------
        Flog.writeline Espacios(Tabulador * 1) & "Buscando Datos de Domicilio y Telefono"
        StrSql = " SELECT * "
        StrSql = StrSql & " From cabdom"
        StrSql = StrSql & " INNER JOIN detdom ON detdom.domnro = cabdom.domnro"
        StrSql = StrSql & " LEFT JOIN localidad ON localidad.locnro = detdom.locnro"
        StrSql = StrSql & " LEFT JOIN provincia ON provincia.provnro = detdom.provnro"
        StrSql = StrSql & " LEFT JOIN telefono ON telefono.domnro = cabdom.domnro"
        StrSql = StrSql & " AND telefono.teldefault = -1"
        StrSql = StrSql & " Where cabdom.domdefault = -1"
        StrSql = StrSql & " AND cabdom.ternro = " & ternro
        OpenRecordset StrSql, rs_Consult
        If Not rs_Consult.EOF Then
            Flog.writeline Espacios(Tabulador * 2) & "Domicilio Encontrado"
            calle = IIf(EsNulo(rs_Consult!calle), "", Left(rs_Consult!calle, 30))
            torre = IIf(EsNulo(rs_Consult!torre), "", Left(rs_Consult!torre, 5))
            nroCalle = IIf(EsNulo(rs_Consult!nro), "0", Left(rs_Consult!nro, 6))
            piso = IIf(EsNulo(rs_Consult!piso), "", Left(rs_Consult!piso, 5))
            oficdepto = IIf(EsNulo(rs_Consult!oficdepto), "", Left(rs_Consult!oficdepto, 5))
            codigopostal = IIf(EsNulo(rs_Consult!codigopostal), "", Left(rs_Consult!codigopostal, 8))
            Flog.writeline Espacios(Tabulador * 2) & "Provincia: " & rs_Consult!provnro & " - " & rs_Consult!provdesc
            provincia = Left(calcularMapeo(rs_Consult!provnro, 3, ""), 2)
            Flog.writeline Espacios(Tabulador * 2) & "Localidad: " & rs_Consult!locnro & " - " & rs_Consult!locdesc
            '18/04/2007 - Martin Ferraro - Cambio Default
            'localidad = Left(calcularMapeo(rs_Consult!locnro, 2, "0"), 10)
            localidad = Left(calcularMapeo(rs_Consult!locnro, 2, ""), 10)
            telefono = IIf(EsNulo(rs_Consult!telnro), "0", Left(rs_Consult!telnro, 15))
            tipoTel = IIf(EsNulo(rs_Consult!telnro), "", "1")
        Else
            Flog.writeline Espacios(Tabulador * 2) & "No se Encuentra el Domicilio"
            calle = ""
            torre = ""
            nroCalle = "0"
            piso = ""
            oficdepto = ""
            codigopostal = ""
            piso = ""
            provincia = ""
            '18/04/2007 - Martin Ferraro - Cambio Default
            'localidad = "0"
            localidad = ""
            telefono = "0"
            tipoTel = ""
        End If
        
        telefono = Replace(telefono, "-", "")
        
        rs_Consult.Close
        
        
        '---------------------------------------------------------------------------------------------------------------
        'DATOS DE CUENTA BANCARIA---------------------------------------------------------------------------------------
        '---------------------------------------------------------------------------------------------------------------
        Flog.writeline Espacios(Tabulador * 1) & "Buscando Datos de Cuenta Bancaria"
        StrSql = " SELECT * "
        StrSql = StrSql & " FROM ctabancaria"
        StrSql = StrSql & " INNER JOIN formapago ON ctabancaria.fpagnro = formapago.fpagnro"
        StrSql = StrSql & " AND formapago.fpagbanc = -1"
        StrSql = StrSql & " WHERE ctabancaria.ternro  = " & ternro
        StrSql = StrSql & " AND ctabancaria.ctabestado = -1"
        OpenRecordset StrSql, rs_Consult
        If Not rs_Consult.EOF Then
            Flog.writeline Espacios(Tabulador * 2) & "Cuenta Bancaria Encontrada"
            cbu = IIf(EsNulo(rs_Consult!ctabcbu), "0", Left(rs_Consult!ctabcbu, 22))
        Else
            Flog.writeline Espacios(Tabulador * 2) & "No se Encuentra la Cuenta Bancaria"
            cbu = "0"
        End If
        rs_Consult.Close
    
    Else
        'Cuando es baja estos datos no me importan
        nivEst = "01"
        calle = ""
        torre = ""
        nroCalle = "0"
        piso = ""
        oficdepto = ""
        codigopostal = ""
        piso = ""
        provincia = ""
        '18/04/2007 - Martin Ferraro - Cambio Default
        'localidad = "0"
        localidad = ""
        telefono = "0"
        tipoTel = ""
        cbu = "0"
        fechaNac = ""
        discapacitado = "N"
        mail = ""
        tipoMail = ""
    End If
    
    '---------------------------------------------------------------------------------------------------------------
    'Grabando datos en BD-------------------------------------------------------------------------------------------
    '---------------------------------------------------------------------------------------------------------------
    Flog.writeline Espacios(Tabulador * 1) & "Guardando datos en rep_f935"
    StrSql = " INSERT INTO rep_f935 "
    StrSql = StrSql & " (bpronro, pliqnro, ternro, empresa,"
    StrSql = StrSql & " apellido, nombre, legajo, cuil,"
    StrSql = StrSql & " agropec, contrato, fec_ini_rec_lab, fec_fin_rec_lab,"
    StrSql = StrSql & " osocial, sit_baja, fecrenuncia, rem, "
    StrSql = StrSql & " suc, actividad, puesto, cod_mov,"
    StrSql = StrSql & " fec_nac, formacion, incapacidad, pais,"
    StrSql = StrSql & " area, telefono, tipo_tel, mail,"
    StrSql = StrSql & " tipo_mail, desc_calle, nro_dom, torre,"
    StrSql = StrSql & " bloque , piso, departamento, cp,"
    StrSql = StrSql & " loc , provincia, cbu, mod_liq,convenio,categoria,tiposervicio)"
    StrSql = StrSql & " VALUES"
    
'StrSql = StrSql & " (bpronro, pliqnro, ternro, empresa,"
    StrSql = StrSql & "(" & NroProcesoBatch
    StrSql = StrSql & "," & Nroliq
    StrSql = StrSql & "," & ternro
    StrSql = StrSql & "," & Empresa
'StrSql = StrSql & " apellido, nombre, legajo, cuil,"
    StrSql = StrSql & ",'" & Mid(terape, 1, 100) & "'"
    StrSql = StrSql & ",'" & Mid(ternom, 1, 100) & "'"
    StrSql = StrSql & "," & Legajo
    StrSql = StrSql & ",'" & Mid(Cuil, 1, 30) & "'"
'StrSql = StrSql & " agropec, contrato, fec_ini_rec_lab, fec_fin_rec_lab,"
    StrSql = StrSql & ",'" & Mid(agropecuario, 1, 1) & "'"
    StrSql = StrSql & ",'" & Mid(contrato, 1, 50) & "'"
    If fechaFaseInicio = "" Then
        StrSql = StrSql & ",null"
    Else
       StrSql = StrSql & "," & ConvFecha(CDate(fechaFaseInicio))
    End If
    If fechaFasefin = "" Then
        StrSql = StrSql & ",null"
    Else
       StrSql = StrSql & "," & ConvFecha(CDate(fechaFasefin))
    End If
'StrSql = StrSql & " osocial, sit_baja, fecrenuncia, rem, "
    StrSql = StrSql & ",'" & Mid(Osocial, 1, 50) & "'"
    StrSql = StrSql & ",'" & Mid(sitRevista, 1, 50) & "'"
    If fechaRenuncia = "" Then
        StrSql = StrSql & ",null"
    Else
       StrSql = StrSql & "," & ConvFecha(CDate(fechaRenuncia))
    End If
    StrSql = StrSql & ",'" & RemuStr & "'"
'StrSql = StrSql & " suc, domicilio, puesto, cod_mov,"
    StrSql = StrSql & ",'" & Mid(sucursal, 1, 50) & "'"
    StrSql = StrSql & ",'" & Mid(actEmpresa, 1, 50) & "'"
    StrSql = StrSql & ",'" & Mid(puesto, 1, 50) & "'"
    StrSql = StrSql & ",'" & codMov & "'"
'StrSql = StrSql & " fec_nac, formacion, incapacidad, pais,"
    If EsNulo(fechaNac) Then
        StrSql = StrSql & ", null"
    Else
        StrSql = StrSql & "," & ConvFecha(CDate(fechaNac))
    End If
    StrSql = StrSql & ",'" & Mid(nivEst, 1, 50) & "'"
    StrSql = StrSql & ",'" & Mid(discapacitado, 1, 1) & "'"
    StrSql = StrSql & ",'0'"
'StrSql = StrSql & " area, telefono, tipo_tel, mail,"
    StrSql = StrSql & ",'0'"
    StrSql = StrSql & ",'" & Mid(telefono, 1, 50) & "'"
    StrSql = StrSql & ",'" & Mid(tipoTel, 1, 50) & "'"
    StrSql = StrSql & ",'" & Mid(mail, 1, 100) & "'"
'StrSql = StrSql & " tipo_mail, desc_calle, nro_dom, torre,"
    StrSql = StrSql & ",'" & Mid(tipoMail, 1, 50) & "'"
    StrSql = StrSql & ",'" & Mid(calle, 1, 50) & "'"
    StrSql = StrSql & ",'" & Mid(nroCalle, 1, 50) & "'"
    StrSql = StrSql & ",'" & Mid(torre, 1, 50) & "'"
'StrSql = StrSql & " bloque , piso, departamento, cp,"
    StrSql = StrSql & ",'0'"
    StrSql = StrSql & ",'" & Mid(piso, 1, 50) & "'"
    StrSql = StrSql & ",'" & Mid(oficdepto, 1, 50) & "'"
    StrSql = StrSql & ",'" & Mid(codigopostal, 1, 50) & "'"
'StrSql = StrSql & " loc , provincia, cbu, mod_liq)"
    StrSql = StrSql & ",'" & Mid(localidad, 1, 50) & "'"
    StrSql = StrSql & ",'" & Mid(provincia, 1, 50) & "'"
    StrSql = StrSql & ",'" & Mid(cbu, 1, 50) & "'"
    StrSql = StrSql & ",'" & Mid(formaLiq, 1, 50) & "'"
    StrSql = StrSql & ",'" & convenio & "'"
    StrSql = StrSql & ",'" & categoria & "'"
    StrSql = StrSql & ",'" & tiposervicio & "')"
    objConn.Execute StrSql, , adExecuteNoRecords
    
    
    '---------------------------------------------------------------------------------------------------------------
    'DATOS DE FAMILIARES--------------------------------------------------------------------------------------------
    '---------------------------------------------------------------------------------------------------------------
    If codMov <> "BT" Then
        Flog.writeline Espacios(Tabulador * 1) & "Buscando Datos de los Familiares"
        Call DatosFamiliares(ternro, Empresa, Cuil)
    End If
    
  
    '---------------------------------------------------------------------------------------------------------------
    'ACTUALIZO EL PROGRESO------------------------------------------------------------------------------------------
    '---------------------------------------------------------------------------------------------------------------
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
             "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    rs_Empleados.MoveNext
    
Loop

'---------------------------------------------------------------------------------------------------------------
'DATOS DE LAS SUCURSALES----------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------
Call DatosSucursales(Empresa)


If rs_Empleados.State = adStateOpen Then rs_Empleados.Close
If rs_Periodo.State = adStateOpen Then rs_Periodo.Close
If rs_Confrep.State = adStateOpen Then rs_Confrep.Close
If rs_Consult.State = adStateOpen Then rs_Consult.Close
If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
If rs_Estr_cod.State = adStateOpen Then rs_Estr_cod.Close

Set rs_Empleados = Nothing
Set rs_Periodo = Nothing
Set rs_Confrep = Nothing
Set rs_Consult = Nothing
Set rs_Estructura = Nothing
Set rs_Estr_cod = Nothing
Exit Sub

E_F935:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: F935"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="
    MyBeginTrans
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
             "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    MyCommitTrans
    
    HuboError = True
    Flog.writeline " Error: " & Err.Description

End Sub


Private Sub DatosFamiliares(ByVal ternro As Long, ByVal Empresa As Long, ByVal Cuil As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Busca los datos de los familiares del ternro
' Autor      : Martin Ferraro
' Fecha      : 28/08/2006
' --------------------------------------------------------------------------------------------

Dim rs_fam As New ADODB.Recordset
Dim rs_Consult As New ADODB.Recordset
Dim rs_Estructura As New ADODB.Recordset
Dim rs_Estr_cod As New ADODB.Recordset

Dim evento As String
Dim CuilFam As String
Dim famApe As String
Dim famNom As String
Dim fechaIniVinc As String
Dim fechaEmision As String
Dim tomo As String
Dim folio As String
Dim acta As String
Dim pais As String
Dim provincia As String
Dim comuna As String
Dim localidad As String
Dim tribunal As String
Dim juzgado As String
Dim secretaria As String
Dim escAnio As String
Dim escNivel As String
Dim escTipo As String
Dim escGrado As String
Dim textoDj As String
Dim codDocumento As String
Dim codDocuExt As String
Dim famTerNro As Long
Dim famPaisNro As Long
Dim famProvNro As Long
Dim famlocNro As Long

'Inicio codigo ejecutable
On Error GoTo E_DatosFamiliares

    'Inicializacion de datos no disponibles en X2
    textoDj = ""

    StrSql = "SELECT tercero.ternro,tercero.terape, tercero.ternom, famest, terfecnac, "
    StrSql = StrSql & " famsalario, famfecvto, famCargaDGI, famDGIdesde, famDGIhasta, famemergencia, "
    StrSql = StrSql & " paredesc, parcodext, famtribunal , famjuzgado, famsecretaria, "
    StrSql = StrSql & " famfec , famfecadopcion, famcertomo, famcerfolio, famceracta, parentesco.parenro, "
    StrSql = StrSql & " famcomuna, famcoddocu, famcoddocuext, familiar.paisnro, familiar.locnro, familiar.provnro "
    StrSql = StrSql & " FROM  tercero INNER JOIN familiar ON tercero.ternro=familiar.ternro "
    StrSql = StrSql & " INNER JOIN parentesco ON familiar.parenro=parentesco.parenro "
    StrSql = StrSql & " WHERE familiar.empleado = " & ternro
    StrSql = StrSql & " AND familiar.famest = -1 "
    StrSql = StrSql & " ORDER BY tercero.ternro "
    OpenRecordset StrSql, rs_fam
    If rs_fam.EOF Then
        Flog.writeline Espacios(Tabulador * 2) & "No se encontraron Familiares"
    Else
        Do While Not rs_fam.EOF
            Flog.writeline Espacios(Tabulador * 2) & "Procesando Familiar " & rs_fam!paredesc & " - " & rs_fam!terape & " " & rs_fam!ternom
            
            famApe = IIf(EsNulo(rs_fam!terape), "", rs_fam!terape)
            famNom = IIf(EsNulo(rs_fam!ternom), "", rs_fam!ternom)
            famTerNro = rs_fam!ternro
            
            Flog.writeline Espacios(Tabulador * 3) & "Buscando Mapeo Parentesco " & rs_fam!parenro & " - " & rs_fam!paredesc
            evento = Left(calcularMapeo(rs_fam!parenro, 5, ""), 2)
            
            fechaIniVinc = IIf(EsNulo(rs_fam!famfec), "", rs_fam!famfec)
            fechaEmision = IIf(EsNulo(rs_fam!famfecadopcion), "", rs_fam!famfecadopcion)
            tomo = IIf(EsNulo(rs_fam!famcertomo), "", Left(rs_fam!famcertomo, 6))
            folio = IIf(EsNulo(rs_fam!famcerfolio), "", Left(rs_fam!famcerfolio, 5))
            tribunal = IIf(EsNulo(rs_fam!famtribunal), "", Left(rs_fam!famtribunal, 50))
            '18/04/2007 - Martin Ferraro - Cambio Default
            'acta = IIf(EsNulo(rs_fam!famceracta), "0", Left(rs_fam!famceracta, 7))
            'juzgado = IIf(EsNulo(rs_fam!famjuzgado), "0", Left(rs_fam!famjuzgado, 4))
            'secretaria = IIf(EsNulo(rs_fam!famsecretaria), "0", Left(rs_fam!famsecretaria, 4))
            juzgado = IIf(EsNulo(rs_fam!famjuzgado), "", Left(rs_fam!famjuzgado, 4))
            secretaria = IIf(EsNulo(rs_fam!famsecretaria), "", Left(rs_fam!famsecretaria, 4))
            acta = IIf(EsNulo(rs_fam!famceracta), "", Left(rs_fam!famceracta, 7))
            comuna = IIf(EsNulo(rs_fam!famcomuna), "", Left(rs_fam!famcomuna, 30))
            codDocumento = IIf(EsNulo(rs_fam!famcoddocu), "0", Left(rs_fam!famcoddocu, 2))
            codDocuExt = IIf(EsNulo(rs_fam!famcoddocuext), "", Left(rs_fam!famcoddocuext, 2))
            
            
            '---------------------------------------------------------------------------------------------------------------
            'CUIL-----------------------------------------------------------------------------------------------------------
            '---------------------------------------------------------------------------------------------------------------
            Flog.writeline Espacios(Tabulador * 3) & "Buscando CUIL familiar"
            StrSql = " SELECT nrodoc "
            StrSql = StrSql & " FROM ter_doc "
            StrSql = StrSql & " WHERE ter_doc.ternro= " & famTerNro
            StrSql = StrSql & " AND ter_doc.tidnro=10 "
            OpenRecordset StrSql, rs_Consult
            If Not rs_Consult.EOF Then
                CuilFam = IIf(EsNulo(rs_Consult!nrodoc), "", rs_Consult!nrodoc)
                Flog.writeline Espacios(Tabulador * 4) & "CUIL = " & CuilFam
            Else
                CuilFam = ""
                Flog.writeline Espacios(Tabulador * 4) & "CUIL no encontrado"
            End If
            rs_Consult.Close
            
            CuilFam = Replace(CuilFam, "-", "")
            CuilFam = Replace(CuilFam, "/", "")
            CuilFam = Left(CuilFam, 11)
            
            '---------------------------------------------------------------------------------------------------------------
            'PAIS-----------------------------------------------------------------------------------------------------------
            '---------------------------------------------------------------------------------------------------------------
            If Not EsNulo(rs_fam!paisnro) Then
                Flog.writeline Espacios(Tabulador * 3) & "Buscando Pais familiar"
                famPaisNro = IIf(EsNulo(rs_fam!paisnro), 0, rs_fam!paisnro)
                
                StrSql = " SELECT * "
                StrSql = StrSql & " FROM pais "
                StrSql = StrSql & " WHERE pais.paisnro = " & famPaisNro
                OpenRecordset StrSql, rs_Consult
                
                If Not rs_Consult.EOF Then
                    Flog.writeline Espacios(Tabulador * 4) & "Pais encontrado = " & rs_Consult!paisnro & " - " & rs_Consult!paisdesc
                    pais = Left(calcularMapeo(famPaisNro, 4, ""), 4)
                Else
                    pais = ""
                    Flog.writeline Espacios(Tabulador * 4) & "Pais no encontrado"
                End If
                rs_Consult.Close
                
            Else
                pais = ""
                Flog.writeline Espacios(Tabulador * 3) & "Pais no cargado"
            End If


            '---------------------------------------------------------------------------------------------------------------
            'Provincia------------------------------------------------------------------------------------------------------
            '---------------------------------------------------------------------------------------------------------------
            If Not EsNulo(rs_fam!provnro) Then
                Flog.writeline Espacios(Tabulador * 3) & "Buscando Provincia familiar"
                famProvNro = IIf(EsNulo(rs_fam!provnro), 0, rs_fam!provnro)
                
                StrSql = " SELECT * "
                StrSql = StrSql & " FROM provincia "
                StrSql = StrSql & " WHERE provincia.provnro = " & famProvNro
                OpenRecordset StrSql, rs_Consult
                
                If Not rs_Consult.EOF Then
                    Flog.writeline Espacios(Tabulador * 4) & "Provincia encontrado = " & rs_Consult!provnro & " - " & rs_Consult!provdesc
                    provincia = Left(calcularMapeo(famProvNro, 3, "0"), 2)
                Else
                    provincia = "0"
                    Flog.writeline Espacios(Tabulador * 4) & "Provincia no encontrado"
                End If
                rs_Consult.Close
                
            Else
                provincia = "0"
                Flog.writeline Espacios(Tabulador * 3) & "Provincia no cargada"
            End If


            '---------------------------------------------------------------------------------------------------------------
            'Localidad------------------------------------------------------------------------------------------------------
            '---------------------------------------------------------------------------------------------------------------
            If Not EsNulo(rs_fam!locnro) Then
                Flog.writeline Espacios(Tabulador * 3) & "Buscando Localidad familiar"
                famlocNro = IIf(EsNulo(rs_fam!locnro), 0, rs_fam!locnro)
                
                StrSql = " SELECT * "
                StrSql = StrSql & " FROM localidad "
                StrSql = StrSql & " WHERE localidad.locnro = " & famlocNro
                OpenRecordset StrSql, rs_Consult
                
                If Not rs_Consult.EOF Then
                    Flog.writeline Espacios(Tabulador * 4) & "Localidad encontrado = " & rs_Consult!locnro & " - " & rs_Consult!locdesc
                    localidad = Left(calcularMapeo(famlocNro, 2, ""), 10)
                Else
                    localidad = ""
                    Flog.writeline Espacios(Tabulador * 4) & "Localidad no encontrado"
                End If
                rs_Consult.Close
                
            Else
                localidad = ""
                Flog.writeline Espacios(Tabulador * 3) & "Localidad no cargada"
            End If


            '---------------------------------------------------------------------------------------------------------------
            'Nivel de estudio-----------------------------------------------------------------------------------------------
            '---------------------------------------------------------------------------------------------------------------
            Flog.writeline Espacios(Tabulador * 3) & "Buscando estudio actual del familiar"
            StrSql = " SELECT estudio_actual.nivnro, estactgra, nivdesc, estactanio"
            StrSql = StrSql & " From estudio_actual"
            StrSql = StrSql & " INNER JOIN nivest ON nivest.nivnro = estudio_actual.nivnro"
            StrSql = StrSql & " Where estudio_actual.ternro = " & famTerNro
            OpenRecordset StrSql, rs_Consult
            
            If Not rs_Consult.EOF Then
                Flog.writeline Espacios(Tabulador * 4) & "Estudio Actual encontrado = " & rs_Consult!nivnro & " - " & rs_Consult!nivdesc
                escTipo = "N"
                escGrado = IIf(EsNulo(rs_Consult!estactgra), "", Left(rs_Consult!estactgra, 1))
                '18/04/2007 - Martin Ferraro - Cambio Default
                'escAnio = IIf(EsNulo(rs_Consult!estactanio), "0", Left(rs_Consult!estactanio, 4))
                escAnio = IIf(EsNulo(rs_Consult!estactanio), "", Left(rs_Consult!estactanio, 4))
                escNivel = Left(calcularMapeo(rs_Consult!nivnro, 1, "01"), 2)
            Else
                '18/04/2007 - Martin Ferraro - Cambio Default
                'escAnio = "0"
                escAnio = ""
                escNivel = "01"
                escGrado = ""
                escTipo = ""
                Flog.writeline Espacios(Tabulador * 4) & "Estudio Actual NO encontrado "
            End If


            '---------------------------------------------------------------------------------------------------------------
            'Grabando datos en BD-------------------------------------------------------------------------------------------
            '---------------------------------------------------------------------------------------------------------------
            Flog.writeline Espacios(Tabulador * 3) & "Guardando datos en rep_f935_fam"
            StrSql = " INSERT INTO rep_f935_fam "
            StrSql = StrSql & " (bpronro, pliqnro, ternro, empresa,"
            StrSql = StrSql & " famape, famnom, famternro, cuil,"
            StrSql = StrSql & " cuilfam, evento, fec_ini_vinc, fec_doc,"
            StrSql = StrSql & " tomo, folio, acta, pais, "
            StrSql = StrSql & " provincia, localidad, comuna, tribunal,"
            StrSql = StrSql & " juzgado, secretaria, escolaridad_anio, escolaridad_tipo,"
            StrSql = StrSql & " escolaridad_nivel, escolaridad_grado, dj_txt, cod_doc,"
            StrSql = StrSql & " cod_doc_ext, cod_mov)"
            StrSql = StrSql & " VALUES"
        'StrSql = StrSql & " (bpronro, pliqnro, ternro, empresa,"
            StrSql = StrSql & "(" & NroProcesoBatch
            StrSql = StrSql & "," & Nroliq
            StrSql = StrSql & "," & ternro
            StrSql = StrSql & "," & Empresa
        'StrSql = StrSql & " famape, famnom, famternro, cuil,"
            StrSql = StrSql & ",'" & Mid(famApe, 1, 50) & "'"
            StrSql = StrSql & ",'" & Mid(famNom, 1, 50) & "'"
            StrSql = StrSql & "," & famTerNro
            StrSql = StrSql & ",'" & Mid(Cuil, 1, 30) & "'"
        'StrSql = StrSql & " cuilfam, evento, fec_ini_vinc, fec_doc,"
            StrSql = StrSql & ",'" & Mid(CuilFam, 1, 30) & "'"
            StrSql = StrSql & ",'" & Mid(evento, 1, 50) & "'"
            If fechaIniVinc = "" Then
                StrSql = StrSql & ",null"
            Else
               StrSql = StrSql & "," & ConvFecha(CDate(fechaIniVinc))
            End If
            If fechaEmision = "" Then
                StrSql = StrSql & ",null"
            Else
               StrSql = StrSql & "," & ConvFecha(CDate(fechaEmision))
            End If
        'StrSql = StrSql & " tomo, folio, acta, pais, "
            StrSql = StrSql & ",'" & Mid(tomo, 1, 50) & "'"
            StrSql = StrSql & ",'" & Mid(folio, 1, 50) & "'"
            StrSql = StrSql & ",'" & Mid(acta, 1, 50) & "'"
            StrSql = StrSql & ",'" & Mid(pais, 1, 50) & "'"
        'StrSql = StrSql & " provincia, localidad, comuna, tribunal,"
            StrSql = StrSql & ",'" & Mid(provincia, 1, 50) & "'"
            StrSql = StrSql & ",'" & Mid(localidad, 1, 50) & "'"
            StrSql = StrSql & ",'" & Mid(comuna, 1, 50) & "'"
            StrSql = StrSql & ",'" & Mid(tribunal, 1, 50) & "'"
        'StrSql = StrSql & " juzgado, secretaria, escolaridad_anio, escolaridad_tipo,"
            StrSql = StrSql & ",'" & Mid(juzgado, 1, 50) & "'"
            StrSql = StrSql & ",'" & Mid(secretaria, 1, 50) & "'"
            StrSql = StrSql & ",'" & Mid(escAnio, 1, 50) & "'"
            StrSql = StrSql & ",'" & Mid(escTipo, 1, 50) & "'"
        'StrSql = StrSql & " escolaridad_nivel, escolaridad_grado, dj_txt, cod_doc,"
            StrSql = StrSql & ",'" & Mid(escNivel, 1, 50) & "'"
            StrSql = StrSql & ",'" & Mid(escGrado, 1, 50) & "'"
            StrSql = StrSql & ",'" & Mid(textoDj, 1, 70) & "'"
            StrSql = StrSql & ",'" & Mid(codDocumento, 1, 50) & "'"
        'StrSql = StrSql & " cod_doc_ext, codmov)"
            StrSql = StrSql & ",'" & Mid(codDocuExt, 1, 50) & "'"
            StrSql = StrSql & ",'AT')"
            
            objConn.Execute StrSql, , adExecuteNoRecords
            
            rs_fam.MoveNext
            
        Loop
    End If
    rs_fam.Close




If rs_fam.State = adStateOpen Then rs_fam.Close
If rs_Consult.State = adStateOpen Then rs_Consult.Close
If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
If rs_Estr_cod.State = adStateOpen Then rs_Estr_cod.Close
Set rs_fam = Nothing
Set rs_Consult = Nothing
Set rs_Estructura = Nothing
Set rs_Estr_cod = Nothing

Exit Sub
E_DatosFamiliares:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: DatosFamiliares"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="
    HuboError = True
    Flog.writeline " Error: " & Err.Description


End Sub

Private Sub DatosSucursales(ByVal Empresa As Long)
' --------------------------------------------------------------------------------------------
' Descripcion: Busca los datos de las sucursales
' Autor      : Martin Ferraro
' Fecha      : 28/08/2006
' --------------------------------------------------------------------------------------------

Dim rs_suc As New ADODB.Recordset
Dim rs_Consult As New ADODB.Recordset
Dim rs_Estr_cod As New ADODB.Recordset

Dim sucnro As Long
Dim sucTernro As Long
Dim sucDesc As String
Dim sucCod As String
Dim calle As String
Dim torre As String
Dim nroCalle As String
Dim oficdepto As String
Dim codigopostal As String
Dim piso As String
Dim provincia As String
Dim localidad As String
Dim domAmb As String
Dim actSuc As String

'Inicio codigo ejecutable
On Error GoTo E_DatosSucursales


    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "------------------------------------------------------"
    Flog.writeline Espacios(Tabulador * 0) & "------------------------------------------------------"
    Flog.writeline Espacios(Tabulador * 0) & "Buscando Datos de las Sucursales"

    'Inicializacion de datos no disponibles en X2
    '18/04/2007 - Martin Ferraro - Cambio Default
    'domAmb = ""
    domAmb = "N"
    actSuc = "0"
    sucCod = ""
    
    StrSql = " Select estructura.estrnro, estructura.estrdabr, sucursal.ternro, estructura.estrcodext "
    StrSql = StrSql & " From estructura"
    StrSql = StrSql & " INNER JOIN sucursal ON sucursal.estrnro = estructura.estrnro"
    StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro = sucursal.ternro"
    StrSql = StrSql & " Where estructura.tenro = 1"
    StrSql = StrSql & " ORDER BY estructura.estrnro"
    OpenRecordset StrSql, rs_suc
    If rs_suc.EOF Then
        Flog.writeline Espacios(Tabulador * 1) & "No se encontraron Sucursales"
    Else
        Do While Not rs_suc.EOF
            Flog.writeline
            Flog.writeline Espacios(Tabulador * 1) & "Procesando Sucursal " & rs_suc!Estrnro & " - " & rs_suc!estrdabr
            sucnro = rs_suc!Estrnro
            sucTernro = rs_suc!ternro
            sucDesc = IIf(EsNulo(rs_suc!estrdabr), "", rs_suc!estrdabr)
            sucCod = Left(IIf(EsNulo(rs_suc!estrcodext), "", rs_suc!estrcodext), 5)
            
            StrSql = "SELECT * FROM estr_cod WHERE estrnro =" & sucnro
            StrSql = StrSql & " AND tcodnro = " & tipoCod
            OpenRecordset StrSql, rs_Estr_cod
            If Not rs_Estr_cod.EOF Then
                actSuc = IIf(EsNulo(rs_Estr_cod!nrocod), "0", Left(CStr(rs_Estr_cod!nrocod), 6))
            Else
                Flog.writeline Espacios(Tabulador * 1) & "No se encontró el codigo para el Tipo de Sucursal."
                actSuc = ""
            End If
            rs_Estr_cod.Close
            
            
            '---------------------------------------------------------------------------------------------------------------
            'DATOS DE DOMICILIO Y TELEFONO----------------------------------------------------------------------------------
            '---------------------------------------------------------------------------------------------------------------
            Flog.writeline Espacios(Tabulador * 2) & "Buscando Datos de Domicilio y Telefono"
            StrSql = " SELECT * "
            StrSql = StrSql & " From cabdom"
            StrSql = StrSql & " INNER JOIN detdom ON detdom.domnro = cabdom.domnro"
            StrSql = StrSql & " LEFT JOIN localidad ON localidad.locnro = detdom.locnro"
            StrSql = StrSql & " LEFT JOIN provincia ON provincia.provnro = detdom.provnro"
            StrSql = StrSql & " Where cabdom.domdefault = -1"
            StrSql = StrSql & " AND cabdom.ternro = " & sucTernro
            OpenRecordset StrSql, rs_Consult
            If Not rs_Consult.EOF Then
                Flog.writeline Espacios(Tabulador * 3) & "Domicilio Encontrado"
                calle = IIf(EsNulo(rs_Consult!calle), "", Left(rs_Consult!calle, 30))
                torre = IIf(EsNulo(rs_Consult!torre), "", Left(rs_Consult!torre, 5))
                nroCalle = IIf(EsNulo(rs_Consult!nro), "0", Left(rs_Consult!nro, 6))
                piso = IIf(EsNulo(rs_Consult!piso), "", Left(rs_Consult!piso, 5))
                oficdepto = IIf(EsNulo(rs_Consult!oficdepto), "", Left(rs_Consult!oficdepto, 5))
                codigopostal = IIf(EsNulo(rs_Consult!codigopostal), "", Left(rs_Consult!codigopostal, 8))
                Flog.writeline Espacios(Tabulador * 3) & "Provincia: " & rs_Consult!provnro & " - " & rs_Consult!provdesc
                provincia = Left(calcularMapeo(rs_Consult!provnro, 3, ""), 2)
                Flog.writeline Espacios(Tabulador * 3) & "Localidad: " & rs_Consult!locnro & " - " & rs_Consult!locdesc
                '18/04/2007 - Martin Ferraro - Cambio Default
                'localidad = Left(calcularMapeo(rs_Consult!locnro, 2, "0"), 10)
                localidad = Left(calcularMapeo(rs_Consult!locnro, 2, ""), 10)
            Else
                Flog.writeline Espacios(Tabulador * 3) & "No se Encuentra el Domicilio"
                calle = ""
                torre = ""
                nroCalle = "0"
                piso = ""
                oficdepto = ""
                codigopostal = ""
                piso = ""
                provincia = ""
                localidad = ""
            End If
            
            
            '---------------------------------------------------------------------------------------------------------------
            'Grabando datos en BD-------------------------------------------------------------------------------------------
            '---------------------------------------------------------------------------------------------------------------
            Flog.writeline Espacios(Tabulador * 2) & "Guardando datos en rep_f935_dom"
            StrSql = " INSERT INTO rep_f935_dom "
            StrSql = StrSql & " (bpronro, pliqnro, empresa, cod_mov,"
            StrSql = StrSql & " dom_amb, desc_calle, nro_dom, torre,"
            StrSql = StrSql & " bloque , piso, departamento, cp,"
            StrSql = StrSql & " loc , provincia, actividad, suc)"
            StrSql = StrSql & " VALUES"
        'StrSql = StrSql & " (bpronro, pliqnro, ternro, cod_mov,"
            StrSql = StrSql & "(" & NroProcesoBatch
            StrSql = StrSql & "," & Nroliq
            StrSql = StrSql & "," & Empresa
            StrSql = StrSql & ",'AT'"
        'StrSql = StrSql & " dom_amb, desc_calle, nro_dom, torre,"
            StrSql = StrSql & ",'" & Mid(domAmb, 1, 50) & "'"
            StrSql = StrSql & ",'" & Mid(calle, 1, 50) & "'"
            StrSql = StrSql & ",'" & Mid(nroCalle, 1, 50) & "'"
            StrSql = StrSql & ",'" & Mid(torre, 1, 50) & "'"
        'StrSql = StrSql & " bloque , piso, departamento, cp,"
            StrSql = StrSql & ",''"
            StrSql = StrSql & ",'" & Mid(piso, 1, 50) & "'"
            StrSql = StrSql & ",'" & Mid(oficdepto, 1, 50) & "'"
            StrSql = StrSql & ",'" & Mid(codigopostal, 1, 50) & "'"
        'StrSql = StrSql & " loc , provincia, actividad, suc)"
            StrSql = StrSql & ",'" & Mid(localidad, 1, 50) & "'"
            StrSql = StrSql & ",'" & Mid(provincia, 1, 50) & "'"
            StrSql = StrSql & ",'" & Mid(actSuc, 1, 50) & "'"
            StrSql = StrSql & ",'" & Mid(sucCod, 1, 50) & "')"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            rs_suc.MoveNext
            
        Loop
    End If
    rs_suc.Close


If rs_suc.State = adStateOpen Then rs_suc.Close
If rs_Consult.State = adStateOpen Then rs_Consult.Close
If rs_Estr_cod.State = adStateOpen Then rs_Estr_cod.Close
Set rs_suc = Nothing
Set rs_Consult = Nothing
Set rs_Estr_cod = Nothing

Exit Sub
E_DatosSucursales:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: DatosSucursales"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="
    HuboError = True
    Flog.writeline " Error: " & Err.Description


End Sub


'----------------------------------------------------------------
Function calcularMapeo(ByVal Parametro, ByVal Tabla, ByVal Default)
' --------------------------------------------------------------------------------------------
' Descripcion: Resuelve el mapeo a un codigo
' Autor      : Martin Ferraro
' Fecha      : 21/12/2006
' --------------------------------------------------------------------------------------------

Dim StrSql As String
Dim rs_Mapeo As New ADODB.Recordset
Dim correcto As Boolean
Dim Salida

'Inicio codigo ejecutable
On Error GoTo E_calcularMapeo
    
    If IsNull(Parametro) Then
       correcto = False
    Else
       correcto = Parametro <> ""
    End If
           
    Salida = Default

    If correcto Then
        
        'Busco el mapeo en BD
        StrSql = " SELECT * FROM mapeo_general "
        StrSql = StrSql & " WHERE maptipnro = " & Tabla
        StrSql = StrSql & " AND mapclanro = 1 " 'Misimplificacion
        StrSql = StrSql & " AND mapgenorigen = '" & Parametro & "' "
        OpenRecordset StrSql, rs_Mapeo
        
        If Not rs_Mapeo.EOF Then
            Salida = CStr(IIf(EsNulo(rs_Mapeo!mapgendestino), Default, rs_Mapeo!mapgendestino))
        Else
            Flog.writeline Espacios(Tabulador * 3) & "No se encontro mapeo tipo " & Tabla & " para el origen " & Parametro
        End If
        
        rs_Mapeo.Close
    
    End If
    
    calcularMapeo = Salida

If rs_Mapeo.State = adStateOpen Then rs_Mapeo.Close
Set rs_Mapeo = Nothing

Exit Function
E_calcularMapeo:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: CalcularMapeo"
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="
    HuboError = True
    Flog.writeline " Error: " & Err.Description

End Function

