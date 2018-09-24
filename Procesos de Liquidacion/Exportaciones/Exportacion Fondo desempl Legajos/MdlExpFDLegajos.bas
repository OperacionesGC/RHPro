Attribute VB_Name = "MdlExpFDLegajos"
Option Explicit

Global Const Version = "1.00" 'Exportacion fondo desempleo banco nacion - legajos
Global Const FechaModificacion = "27/09/2007"
Global Const UltimaModificacion = " " 'Martin Ferraro - Version Inicial


Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial del Generador de Exportacion de Fondo desempleo Legajos.
' Autor      : Martin Ferraro
' Fecha      : 27/09/2007
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
    
    Nombre_Arch = PathFLog & "Exp. FD Legajos" & "-" & NroProcesoBatch & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaModificacion
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-------------------------------------------------"
    Flog.writeline
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0 ,bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Flog.writeline "Pone el estado en procesando"
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 201 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        bprcparam = rs_batch_proceso!bprcparam
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        Call ExportLegFondoDesempl(NroProcesoBatch, bprcparam)
    Else
        Flog.writeline "no encontró el proceso"
    End If
    
    TiempoFinalProceso = GetTickCount
    Flog.writeline
    Flog.writeline "--------------------------------------------------------------"
    Flog.writeline "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    
    If HuboError Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
    End If
    
    
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


Public Sub ExportLegFondoDesempl(ByVal bpronro As Long, ByVal Parametros As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de Exportacion de legajos de fondo de desempleo de banco nacion
' Autor      : Martin Ferraro
' Fecha      : 27/09/2007
' --------------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------------------
'Parametros
'-------------------------------------------------------------------------------------------------
Dim LegDesde As Long
Dim LegHasta As Long
Dim ParEstado As Integer
Dim Empnro As Long
Dim Tenro1 As Long
Dim Tenro2 As Long
Dim Tenro3 As Long
Dim Estrnro1 As Long
Dim Estrnro2 As Long
Dim Estrnro3 As Long
Dim FecEstr As Date

'-------------------------------------------------------------------------------------------------
'Variables
'-------------------------------------------------------------------------------------------------
Dim ArrPar
Dim ternro As Long
Dim Directorio As String
Dim Separador As String
Dim SeparadorDecimal As String
Dim DescripcionModelo As String
Dim Archivo As String
Dim fExport
Dim carpeta
Dim Aux_Linea As String
Dim casa As String
Dim IdenEmpresa As String
Dim Legajo As Long
Dim TerApeNomb As String
Dim TipoPer As String
Dim PaisDefaultNro As Long
Dim PaisNro As Long
Dim TipDoc As String
Dim DocNro As String
Dim Cuil As String
Dim Calle As String
Dim NroCalle As String
Dim Piso As String
Dim Oficdepto As String
Dim Codigopostal As String
Dim ProvinciaNro As String
Dim provincia As String
Dim LocDesc As String
Dim Telefono As String
Dim TipoCod As Long

'-------------------------------------------------------------------------------------------------
'RecordSets
'-------------------------------------------------------------------------------------------------
Dim rs_Consult As New ADODB.Recordset
Dim rs_Aux As New ADODB.Recordset
Dim rs_Empleados As New ADODB.Recordset

'Inicio codigo ejecutable
On Error GoTo E_LegFondoDesemple


'-------------------------------------------------------------------------------------------------
'Levanto cada parametro por separado, el separador de parametros es "@"
'-------------------------------------------------------------------------------------------------
Flog.writeline "levantando parametros" & Parametros
If Not IsNull(Parametros) Then

    If Len(Parametros) >= 1 Then
    
        ArrPar = Split(Parametros, "@")
        If UBound(ArrPar) = 10 Then
            LegDesde = IIf(ArrPar(0) = "", 0, CLng(ArrPar(0)))
            LegHasta = IIf(ArrPar(1) = "", 0, CLng(ArrPar(1)))
            Flog.writeline Espacios(Tabulador * 0) & "Empleados Desde " & LegDesde & " Hasta " & LegHasta
            
            ParEstado = CInt(ArrPar(2))
            Select Case ParEstado
                Case -1:
                    Flog.writeline Espacios(Tabulador * 0) & "Activos"
                Case 0:
                    Flog.writeline Espacios(Tabulador * 0) & "Inactivos"
                Case 1:
                    Flog.writeline Espacios(Tabulador * 0) & "Ambos"
            End Select
            
            Empnro = CLng(ArrPar(3))
            Flog.writeline Espacios(Tabulador * 0) & "Estructura Empresa = " & Empnro
        
            Tenro1 = CLng(ArrPar(4))
            Flog.writeline Espacios(Tabulador * 0) & "Nivel TE1 = " & Tenro1
            
            Tenro2 = CLng(ArrPar(5))
            Flog.writeline Espacios(Tabulador * 0) & "Nivel TE2 = " & Tenro2
            
            Tenro3 = CLng(ArrPar(6))
            Flog.writeline Espacios(Tabulador * 0) & "Nivel TE3 = " & Tenro3
            
            Estrnro1 = CLng(ArrPar(7))
            Flog.writeline Espacios(Tabulador * 0) & "Estruct1 = " & Estrnro1
            
            Estrnro2 = CLng(ArrPar(8))
            Flog.writeline Espacios(Tabulador * 0) & "Estruct2 = " & Estrnro2
            
            Estrnro3 = CLng(ArrPar(9))
            Flog.writeline Espacios(Tabulador * 0) & "Estruct3 = " & Estrnro3
            
            FecEstr = CDate(ArrPar(10))
            Flog.writeline Espacios(Tabulador * 0) & "Fecha Estruct = " & FecEstr
            
        Else
            Flog.writeline Espacios(Tabulador * 0) & "ERROR. Numero de parametros erroneo."
            Exit Sub
        End If
        
    End If
Else
    Flog.writeline Espacios(Tabulador * 0) & "ERROR. No se encontraron los parametros."
    Exit Sub
End If
Flog.writeline


'-------------------------------------------------------------------------------------------------
'Configuracion del Reporte
'-------------------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 0) & "Buscando configuracion del reporte."
StrSql = "SELECT * FROM confrep WHERE repnro = 215 ORDER BY confnrocol"
OpenRecordset StrSql, rs_Consult
If rs_Consult.EOF Then
    Flog.writeline Espacios(Tabulador * 1) & "No se encontró la configuración del Reporte."
Else
    casa = ""
    TipoCod = 0
    Do While Not rs_Consult.EOF
                
        Flog.writeline Espacios(Tabulador * 1) & "Columna " & rs_Consult!confnrocol & " " & rs_Consult!confetiq & " VNUM = " & rs_Consult!confval & " VALF = " & rs_Consult!confval2
        Select Case rs_Consult!confnrocol
            Case 1:
                If EsNulo(rs_Consult!confval2) Then
                    Flog.writeline Espacios(Tabulador * 1) & "Falta configurar el valor de CASA en el campo alfanumerico de la columna 1."
                    HuboError = True
                    Exit Sub
                Else
                    casa = rs_Consult!confval2
                End If
            Case 2:
                If EsNulo(rs_Consult!confval) Then
                    Flog.writeline Espacios(Tabulador * 1) & "Falta configurar el valor de Tipo de Codigo en el campo numerico de la columna 2."
                    HuboError = True
                    Exit Sub
                Else
                    TipoCod = rs_Consult!confval
                End If
        End Select
        
        rs_Consult.MoveNext
    Loop
End If
rs_Consult.Close
Flog.writeline


'-------------------------------------------------------------------------------------------------
'Validaciones campos obligatorios confrep
'-------------------------------------------------------------------------------------------------
If Len(casa) = 0 Then
    Flog.writeline Espacios(Tabulador * 1) & "Falta configurar el valor de CASA en el campo alfanumerico de la columna 1."
    HuboError = True
    Exit Sub
End If
If TipoCod = 0 Then
    Flog.writeline Espacios(Tabulador * 1) & "Falta configurar el valor de Tipo de Codigo en el campo numerico de la columna 2."
    HuboError = True
    Exit Sub
End If
    

'-------------------------------------------------------------------------------------------------
'Configuracion del Directorio de salida
'-------------------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 0) & "Buscando directorio de salida."
StrSql = "SELECT sis_direntradas FROM sistema WHERE sisnro = 1 "
OpenRecordset StrSql, rs_Consult
If Not rs_Consult.EOF Then
    Directorio = Trim(rs_Consult!sis_direntradas)
Else
    Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontró el registro de la tabla sistema nro 1"
    HuboError = True
    Exit Sub
End If
rs_Consult.Close
Flog.writeline
    
    
'-------------------------------------------------------------------------------------------------
'Configuracion del Modelo
'-------------------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 0) & "Buscando Modelo Interface."
StrSql = "SELECT * FROM modelo WHERE modnro = 915"
OpenRecordset StrSql, rs_Consult
If Not rs_Consult.EOF Then
    Directorio = Directorio & Trim(rs_Consult!modarchdefault)
    Separador = IIf(Not IsNull(rs_Consult!modseparador), rs_Consult!modseparador, ",")
    SeparadorDecimal = IIf(Not IsNull(rs_Consult!modsepdec), rs_Consult!modsepdec, ".")
    DescripcionModelo = rs_Consult!moddesc
    
    Flog.writeline Espacios(Tabulador * 1) & "Modelo 915 " & rs_Consult!moddesc
    Flog.writeline Espacios(Tabulador * 1) & "Directorio de Exportacion : " & Directorio
Else
    Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontró el modelo 915."
    HuboError = True
    Exit Sub
End If
rs_Consult.Close
Flog.writeline


'-------------------------------------------------------------------------------------------------
'Pais default
'-------------------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 0) & "Buscando Pais Default del Sistema."
PaisDefaultNro = 0
StrSql = "SELECT paisnro, paisdesc, paisdef"
StrSql = StrSql & " FROM pais"
StrSql = StrSql & " WHERE paisdef = -1"
OpenRecordset StrSql, rs_Consult
If Not rs_Consult.EOF Then
    Flog.writeline Espacios(Tabulador * 1) & "Pais Default del sistema: " & rs_Consult!PaisNro & " " & rs_Consult!paisdesc
    PaisDefaultNro = rs_Consult!PaisNro
Else
    Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontró el Pais Default."
End If
rs_Consult.Close
Flog.writeline


'-------------------------------------------------------------------------------------------------
'Datos de la empresa
'-------------------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 0) & "Buscando el tipo de codigo " & TipoCod & " de la empresa para Identificacion."
IdenEmpresa = ""
StrSql = "SELECT nrocod"
StrSql = StrSql & " FROM estructura"
StrSql = StrSql & " INNER JOIN estr_cod ON estr_cod.estrnro = estructura.estrnro"
StrSql = StrSql & " AND estr_cod.tcodnro = " & TipoCod
StrSql = StrSql & " WHERE estructura.estrnro = " & Empnro
OpenRecordset StrSql, rs_Consult
If Not rs_Consult.EOF Then
    IdenEmpresa = IIf(EsNulo(rs_Consult!nrocod), "", rs_Consult!nrocod)
Else
    Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontró el tipo de codigo de la empresa."
    HuboError = True
    Exit Sub
End If
rs_Consult.Close
Flog.writeline

'-------------------------------------------------------------------------------------------------
'Busqueda de empleados
'-------------------------------------------------------------------------------------------------
'Buscando todos los empleados
StrSql = "SELECT distinct empleado.empleg, empleado.ternro, empleado.terape, empleado.terape2, empleado.ternom, empleado.ternom2"
StrSql = StrSql & " FROM empleado"
'Empresa
StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.ternro = empleado.ternro"
StrSql = StrSql & " AND his_estructura.htetdesde <=" & ConvFecha(FecEstr) & " AND"
StrSql = StrSql & " (his_estructura.htethasta >= " & ConvFecha(FecEstr) & " OR his_estructura.htethasta IS NULL)"
StrSql = StrSql & " AND his_estructura.tenro  = 10"
StrSql = StrSql & " AND his_estructura.estrnro = " & Empnro
'Filtro por los tres niveles de estructuras
If Tenro1 <> 0 Then
    StrSql = StrSql & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro  AND estact1.tenro  = " & Tenro1
    StrSql = StrSql & " AND estact1.htetdesde<=" & ConvFecha(FecEstr) & " AND "
    StrSql = StrSql & " (estact1.htethasta >= " & ConvFecha(FecEstr) & " OR estact1.htethasta IS NULL)"
    If Estrnro1 <> -1 Then
        StrSql = StrSql & " AND estact1.estrnro =" & Estrnro1
    End If
End If
If Tenro2 <> 0 Then
    StrSql = StrSql & " INNER JOIN his_estructura estact2 ON empleado.ternro = estact2.ternro  AND estact2.tenro  = " & Tenro2
    StrSql = StrSql & " AND estact2.htetdesde<=" & ConvFecha(FecEstr) & " AND "
    StrSql = StrSql & " (estact2.htethasta >= " & ConvFecha(FecEstr) & " OR estact2.htethasta IS NULL)"
    If Estrnro2 <> -1 Then
        StrSql = StrSql & " AND estact2.estrnro =" & Estrnro2
    End If
End If
If Tenro3 <> 0 Then
    StrSql = StrSql & " INNER JOIN his_estructura estact3 ON empleado.ternro = estact3.ternro  AND estact3.tenro  = " & Tenro3
    StrSql = StrSql & " AND estact3.htetdesde<=" & ConvFecha(FecEstr) & " AND "
    StrSql = StrSql & " (estact3.htethasta >= " & ConvFecha(FecEstr) & " OR estact3.htethasta IS NULL)"
    If Estrnro3 <> -1 Then
        StrSql = StrSql & " AND estact3.estrnro =" & Estrnro3
    End If
End If

StrSql = StrSql & " WHERE"
StrSql = StrSql & " empleado.empleg >= " & LegDesde
StrSql = StrSql & " AND empleado.empleg <= " & LegHasta
Select Case ParEstado
    Case -1:
        StrSql = StrSql & " AND empleado.empest = -1"
    Case 0:
        StrSql = StrSql & " AND empleado.empest = 0"
End Select

StrSql = StrSql & " ORDER BY empleado.empleg"

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
        
        
'-------------------------------------------------------------------------------------------------
'Creacion del archivo
'-------------------------------------------------------------------------------------------------
If Not rs_Empleados.EOF Then
    'Seteo el nombre del archivo generado
    Archivo = Directorio & "\VUELFD_" & Format(Date, "dd-mm-yyyy") & "_" & bpronro & ".txt"
    Set fs = CreateObject("Scripting.FileSystemObject")
    'Activo el manejador de errores
    On Error Resume Next
    Set fExport = fs.CreateTextFile(Archivo, True)
    If Err.Number <> 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "La carpeta Destino no existe. Se creará."
        Set carpeta = fs.CreateFolder(Directorio)
        Set fExport = fs.CreateTextFile(Archivo, True)
    End If
    On Error GoTo E_LegFondoDesemple
    Flog.writeline Espacios(Tabulador * 0) & "Archivo Creado: " & "VUELFD_" & Format(Date, "dd-mm-yyyy") & "_" & bpronro & ".txt"
End If
  
        
'-------------------------------------------------------------------------------------------------
'Comienzo
'-------------------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 0) & "-------------------------------------------------------------------"
Flog.writeline Espacios(Tabulador * 0) & "Comienza el Procesamiento de Empleados"
Flog.writeline Espacios(Tabulador * 0) & "-------------------------------------------------------------------"
Do While Not rs_Empleados.EOF
    
    ternro = rs_Empleados!ternro
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "PROCESANDO EMPLEADO: " & rs_Empleados!empleg & " - " & rs_Empleados!terape & " " & rs_Empleados!ternom
    
    
    Legajo = rs_Empleados!empleg
    TerApeNomb = IIf(EsNulo(rs_Empleados!terape), "", rs_Empleados!terape)
    TerApeNomb = TerApeNomb & IIf(EsNulo(rs_Empleados!terape2), "", " " & rs_Empleados!terape2)
    TerApeNomb = TerApeNomb & ", " & IIf(EsNulo(rs_Empleados!ternom), "", rs_Empleados!ternom)
    TerApeNomb = TerApeNomb & IIf(EsNulo(rs_Empleados!ternom2), "", " " & rs_Empleados!ternom2)
    
    'Buscado Datos del tercero
    StrSql = "SELECT tercero.tersex, tercero.paisnro"
    StrSql = StrSql & " FROM tercero "
    StrSql = StrSql & " WHERE tercero.ternro = " & ternro
    OpenRecordset StrSql, rs_Consult
    TipoPer = ""
    PaisNro = 0
    If Not rs_Consult.EOF Then
        TipoPer = IIf(rs_Consult!tersex = -1, "1", "2")
        PaisNro = IIf(EsNulo(rs_Consult!PaisNro), 0, rs_Consult!PaisNro)
    End If
    rs_Consult.Close
    
    'Ver si es extranjero
    If (PaisNro <> PaisDefaultNro) Then
        TipoPer = "3"
    End If
    
    'Buscando datos del documento
    StrSql = "SELECT ter_doc.nrodoc, ter_doc.tidnro"
    StrSql = StrSql & " FROM ter_doc "
    StrSql = StrSql & " WHERE ter_doc.ternro = " & ternro
    StrSql = StrSql & " AND ter_doc.tidnro <= 5 "
    StrSql = StrSql & " ORDER BY tidnro"
    OpenRecordset StrSql, rs_Consult

    TipDoc = ""
    DocNro = ""
    If Not rs_Consult.EOF Then
        Select Case CLng(rs_Consult!TidNro)
            Case 1: 'DNI
                TipDoc = "96"
            Case 2: 'LE
                TipDoc = "90"
            Case 3: 'LC
                TipDoc = "89"
            Case 4: 'CI
                TipDoc = "01"
            Case 5: 'PAS
                TipDoc = ""
        End Select
        DocNro = rs_Consult!nrodoc
    End If
    DocNro = Replace(DocNro, "-", "")
    rs_Consult.Close
    
    
    'Buscando datos del documento CUIL
    StrSql = "SELECT ter_doc.nrodoc, ter_doc.tidnro"
    StrSql = StrSql & " FROM ter_doc "
    StrSql = StrSql & " WHERE ter_doc.ternro = " & ternro
    StrSql = StrSql & " AND ter_doc.tidnro = 10 "
    OpenRecordset StrSql, rs_Consult

    Cuil = ""
    If Not rs_Consult.EOF Then
        If Not EsNulo(rs_Consult!nrodoc) Then
            Cuil = rs_Consult!nrodoc
        End If
    End If
    Cuil = Replace(Cuil, "-", "")
    rs_Consult.Close
    
    
    'Datos del Domicilio
    Calle = ""
    NroCalle = ""
    Piso = ""
    Oficdepto = ""
    Codigopostal = ""
    ProvinciaNro = "0"
    LocDesc = ""
    Telefono = ""
    StrSql = " SELECT detdom.calle, detdom.nro, detdom.piso, detdom.oficdepto, detdom.codigopostal"
    StrSql = StrSql & " , detdom.provnro, detdom.locnro, provincia.provdesc, localidad.locdesc, telefono.telnro"
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
        If Not EsNulo(rs_Consult!Calle) Then Calle = rs_Consult!Calle
        If Not EsNulo(rs_Consult!nro) Then NroCalle = rs_Consult!nro
        If Not EsNulo(rs_Consult!Piso) Then Piso = rs_Consult!Piso
        If Not EsNulo(rs_Consult!Oficdepto) Then Oficdepto = rs_Consult!Oficdepto
        If Not EsNulo(rs_Consult!Codigopostal) Then Codigopostal = rs_Consult!Codigopostal
        If Not EsNulo(rs_Consult!provnro) Then ProvinciaNro = rs_Consult!provnro
        If Not EsNulo(rs_Consult!LocDesc) Then LocDesc = rs_Consult!LocDesc
        If Not EsNulo(rs_Consult!telnro) Then Telefono = rs_Consult!telnro
    End If
    rs_Consult.Close
    
    Telefono = Replace(Telefono, "-", "")
    provincia = calcularMapeo(ProvinciaNro, 3, "")
    Flog.writeline Espacios(Tabulador * 1) & "El mapeo de la provincia nro " & ProvinciaNro & " es " & provincia
    
    'Casa N(4)
    Aux_Linea = Format_StrNro(Left(casa, 4), 4, True, "0")
    'Identificacion Empresa C(4)
    Aux_Linea = Aux_Linea & Separador & Format_Str(Left(IdenEmpresa, 4), 4, True, " ")
    'Legajo N(9)
    Aux_Linea = Aux_Linea & Separador & Format_StrNro(Left(Legajo, 9), 9, True, "0")
    'Nombre del Titular C(30)
    Aux_Linea = Aux_Linea & Separador & Format_Str(Left(TerApeNomb, 30), 30, True, " ")
    'Tipo de Persona N(1)
    Aux_Linea = Aux_Linea & Separador & Format_StrNro(Left(TipoPer, 1), 1, True, "0")
    'Tipo de Documento N(2)
    Aux_Linea = Aux_Linea & Separador & Format_StrNro(Left(TipDoc, 2), 2, True, "0")
    'Nro de Documento N(8)
    Aux_Linea = Aux_Linea & Separador & Format_StrNro(Left(DocNro, 8), 8, True, "0")
    'CUIL N(15)
    Aux_Linea = Aux_Linea & Separador & Format_StrNro(Left(Cuil, 15), 15, True, "0")
    'Calle C(15)
    Aux_Linea = Aux_Linea & Separador & Format_Str(Left(Calle, 15), 15, True, " ")
    'Numero N(5)
    Aux_Linea = Aux_Linea & Separador & Format_StrNro(Left(NroCalle, 5), 5, True, "0")
    'Piso N(2)
    Aux_Linea = Aux_Linea & Separador & Format_StrNro(Left(Piso, 2), 2, True, "0")
    'Departamento C(2)
    Aux_Linea = Aux_Linea & Separador & Format_Str(Left(Oficdepto, 2), 2, True, " ")
    'Telefono N(8)
    Aux_Linea = Aux_Linea & Separador & Format_StrNro(Left(Telefono, 8), 8, True, "0")
    'Localidad C(15)
    Aux_Linea = Aux_Linea & Separador & Format_Str(Left(LocDesc, 15), 15, True, " ")
    'Codigo Postal N(4)
    Aux_Linea = Aux_Linea & Separador & Format_StrNro(Left(Codigopostal, 4), 4, True, "0")
    'COD/PROV/PAIS N(3)
    Aux_Linea = Aux_Linea & Separador & Format_StrNro(Left(provincia, 3), 3, True, "0")
    'Relleno C(13)
    Aux_Linea = Aux_Linea & Separador & Format_StrNro("", 13, True, " ")
    
    'Imprimo la linea
    fExport.writeline Aux_Linea
    
    
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


If rs_Empleados.State = adStateOpen Then rs_Empleados.Close
If rs_Consult.State = adStateOpen Then rs_Consult.Close
If rs_Aux.State = adStateOpen Then rs_Aux.Close


Set rs_Empleados = Nothing
Set rs_Consult = Nothing
Set rs_Aux = Nothing

Exit Sub

E_LegFondoDesemple:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: E_LegFondoDesemple"
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
        StrSql = StrSql & " AND mapclanro = 2 " 'Clase Fondo Desempleo
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

