Attribute VB_Name = "Mdlexporcitricos"
Option Explicit

'   Exportación de Citricos

Global Const Version = "1.03"
Global Const FechaVersion = "04/11/2015"
Global Const UltimaModificacion = "se amplió el tiempo de procesamiento"
Global Const UltimaModificacion1 = "Miriam Ruiz"
'CAS-33588 - NGA BASE CITRICOS - Bug en exportar interfaz progress


'Global Const Version = "1.02"
'Global Const FechaVersion = "25/08/2015"
'Global Const UltimaModificacion = "Se generan archivos de exportación para cítricos"
'Global Const UltimaModificacion1 = "Miriam Ruiz"
'CAS-30747 - NGA - Citricos - Exportacion archivos Progress -entrega 2 - se corrigió la exportación de las licencias


'Global Const Version = "1.01"
'Global Const FechaVersion = "14/08/2015"
'Global Const UltimaModificacion = "Se generan archivos de exportación para cítricos"
'Global Const UltimaModificacion1 = "Miriam Ruiz"
'CAS-30747 - NGA - Citricos - Exportacion archivos Progress -entrega 2 - se agregó parametro activo/inactivo/ambos


'Global Const Version = "1.00"
'Global Const FechaVersion = "08/07/2015"
'Global Const UltimaModificacion = "Se generan archivos de exportación para cítricos"
'Global Const UltimaModificacion1 = "Miriam Ruiz"
'CAS-30747 - NGA - Citricos - Exportacion archivos Progress
'------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------

Global nListaProc As String
Global nEmpresa As Long
Global ArchExp
Global iduser As String
Global FechaDesde As Date
Global FechaHasta As Date

Global cargo As Double
Global categorias As Double
Global conceptos As String
Global contrato As Double
Global convenio As Double
Global costos As Double
Global empresas  As Double
Global DomicilioEmp  As Double
Global tipolicencia As String
Global sectores As Double
Global chrsep As String
Global Listaempleados As String
Global Encab As String
Global licVac As String
Global Linea As String
Global nrofase As Long
Global tipliquidacion As Long
Global Fechaegreso As String
Global Separador As String
Global cuil As String


'----------------------------------------------------------

Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial de la Exportación de Citricos
' Autor      : Raul CHinestra
' Fecha      : 01/09/2006
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
    
    Nombre_Arch = PathFLog & "Exportación_progress" & "-" & NroProcesoBatch & ".log"
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaVersion
    Flog.writeline "Modificacion:              " & UltimaModificacion
    Flog.writeline "                           " & UltimaModificacion1
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-------------------------------------------------"
    Flog.writeline
    
    
    'Abro la conexion
    OpenConnection strconexion, objConn
    objConn.CommandTimeout = 600
   'objConn.ConnectionTimeout = 600
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
    
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0 ,bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Flog.writeline "Pone el estado en procesando"
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 453 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        bprcparam = rs_batch_proceso!bprcparam
        iduser = rs_batch_proceso!iduser
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        Call ExpEmp(NroProcesoBatch, bprcparam)
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

Public Function buscarseparador(ByVal chrsep As String)
Dim aux As String

    Select Case chrsep
        Case "chr(44)":   aux = ","
        Case "chr(59)":   aux = ";"
        Case "chr(46)":   aux = "."
        Case "chr(124)":   aux = "|"
        Case "chr(127)":   aux = " "
    End Select
    buscarseparador = aux
    
End Function

Public Sub Exportar(ByVal UsaEncabezado As Boolean, ByVal columna As Integer, ByVal NombreArch As String, ByVal Encab As String)

Dim rs_exportar As New ADODB.Recordset
Dim rs_exportaraux As New ADODB.Recordset
Dim Linea As String


     
    Set ArchExp = fs.CreateTextFile(NombreArch & ".csv", True)
    Flog.writeline "creando archivo" & NombreArch
    If UsaEncabezado Then
    '-------------------------------------------------------------------------------------------
    ' Exporto a un CSV
    '-------------------------------------------------------------------------------------------
    
         ArchExp.Write Encab
         ArchExp.writeline ""
     End If

    Select Case columna
        Case 2: ' cargo
                StrSql = " SELECT distinct estructura.estrnro,estructura.estrdabr "
                StrSql = StrSql & " FROM estructura "
                StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.estrnro= estructura.estrnro "
                StrSql = StrSql & " AND estructura.tenro =  " & cargo
                StrSql = StrSql & " AND (htetdesde <= " & ConvFecha(FechaHasta)
                StrSql = StrSql & " AND (htethasta >= " & ConvFecha(FechaDesde) & " OR htethasta is null))"
                If Listaempleados <> "-1" Then
                    StrSql = StrSql & " AND ternro IN (" & Listaempleados & ")"
                End If
                OpenRecordset StrSql, rs_exportar
                Do While Not rs_exportar.EOF
                   Linea = rs_exportar!Estrnro & Separador & rs_exportar!estrdabr
                    ArchExp.Write Linea
                    ArchExp.writeline ""
                   rs_exportar.MoveNext
                Loop
                ArchExp.Close
        Case 3  'categoria
                StrSql = " SELECT distinct estructura.estrnro,estructura.estrdabr "
                StrSql = StrSql & " FROM estructura "
                StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.estrnro= estructura.estrnro "
                StrSql = StrSql & " AND estructura.tenro =  " & categorias
                StrSql = StrSql & " AND (htetdesde <= " & ConvFecha(FechaHasta)
                StrSql = StrSql & " AND (htethasta >= " & ConvFecha(FechaDesde) & " OR htethasta is null))"
                If Listaempleados <> "-1" Then
                    StrSql = StrSql & " AND ternro IN (" & Listaempleados & ")"
                End If
                OpenRecordset StrSql, rs_exportar
                Do While Not rs_exportar.EOF
                   Linea = rs_exportar!Estrnro & Separador & rs_exportar!estrdabr
                    ArchExp.Write Linea
                    ArchExp.writeline ""
                   rs_exportar.MoveNext
                Loop
                ArchExp.Close
        Case 4: ' concepto
                StrSql = " SELECT  c.concnro,c.concabr,f.fordabr, cft.tpanro  FROM con_for_tpa cft"
                StrSql = StrSql & " INNER JOIN concepto C ON cft.concnro = c.concnro "
                StrSql = StrSql & " INNER JOIN formula F ON cft.fornro = f.fornro "
                StrSql = StrSql & " INNER JOIN  cft_def ON cft_def.concnro = c.concnro  AND cft_def.tpanro = cft.tpanro "
                StrSql = StrSql & " WHERE cftauto IN (0) " ' -- 0 = novedad y -1 por busqueda
                StrSql = StrSql & " AND nivelo = 2 AND cft_def.nivelc=0 "
                If conceptos <> "" And conceptos <> "0" Then
                    StrSql = StrSql & " AND c.concnro in (" & conceptos & ")"
                End If
                
                StrSql = StrSql & " ORDER BY c.concnro,f.fornro,cft.tpanro,nivel "
                Flog.writeline "conceptos" & StrSql
                 OpenRecordset StrSql, rs_exportar
                Do While Not rs_exportar.EOF
                   Linea = rs_exportar!ConcNro & Separador & rs_exportar!concabr & Separador & rs_exportar!Fordabr & Separador & rs_exportar!tpanro
                    ArchExp.Write Linea
                    ArchExp.writeline ""
                   rs_exportar.MoveNext
                Loop
                ArchExp.Close
         Case 5: ' contrato
                StrSql = " SELECT distinct estructura.estrnro,estructura.estrdabr "
                StrSql = StrSql & " FROM estructura "
                StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.estrnro= estructura.estrnro "
                StrSql = StrSql & " AND estructura.tenro =  " & contrato
                StrSql = StrSql & " AND (htetdesde <= " & ConvFecha(FechaHasta)
                StrSql = StrSql & " AND (htethasta >= " & ConvFecha(FechaDesde) & " OR htethasta is null))"
                If Listaempleados <> "-1" Then
                    StrSql = StrSql & " AND ternro IN (" & Listaempleados & ")"
                End If
                OpenRecordset StrSql, rs_exportar
                Do While Not rs_exportar.EOF
                   Linea = rs_exportar!Estrnro & Separador & rs_exportar!estrdabr
                   ArchExp.Write Linea
                    ArchExp.writeline ""
                   rs_exportar.MoveNext
                Loop
                ArchExp.Close
            
         Case 6: ' convenio
                StrSql = " SELECT distinct estructura.estrnro,estructura.estrdabr "
                StrSql = StrSql & " FROM estructura "
                StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.estrnro= estructura.estrnro "
                StrSql = StrSql & " AND estructura.tenro =  " & convenio
                StrSql = StrSql & " AND (htetdesde <= " & ConvFecha(FechaHasta)
                StrSql = StrSql & " AND (htethasta >= " & ConvFecha(FechaDesde) & " OR htethasta is null))"
                If Listaempleados <> "-1" Then
                    StrSql = StrSql & " AND ternro IN (" & Listaempleados & ")"
                End If
                OpenRecordset StrSql, rs_exportar
                Do While Not rs_exportar.EOF
                   Linea = rs_exportar!Estrnro & Separador & rs_exportar!estrdabr
                  ArchExp.Write Linea
                    ArchExp.writeline ""
                   rs_exportar.MoveNext
                Loop
                ArchExp.Close
                
         Case 7: 'centro de costos
                StrSql = " SELECT distinct estructura.estrnro,estructura.estrcodext,estructura.estrdext "
                StrSql = StrSql & " FROM estructura "
                StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.estrnro= estructura.estrnro "
                StrSql = StrSql & " AND estructura.tenro =  " & costos
                StrSql = StrSql & " AND (htetdesde <= " & ConvFecha(FechaHasta)
                StrSql = StrSql & " AND (htethasta >= " & ConvFecha(FechaDesde) & " OR htethasta is null))"
                If Listaempleados <> "-1" Then
                    StrSql = StrSql & " AND ternro IN (" & Listaempleados & ")"
                End If
                OpenRecordset StrSql, rs_exportar
                Do While Not rs_exportar.EOF
                   Linea = rs_exportar!estrcodext & Separador & rs_exportar!estrdext
                   ArchExp.Write Linea
                    ArchExp.writeline ""
                   rs_exportar.MoveNext
                Loop
                ArchExp.Close
         Case 8: 'empresas
                StrSql = " SELECT distinct estructura.estrdabr,estructura.estrcodext  "
                StrSql = StrSql & " FROM estructura "
                StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.estrnro= estructura.estrnro "
                StrSql = StrSql & " AND estructura.tenro =  " & empresas
                StrSql = StrSql & " AND (htetdesde <= " & ConvFecha(FechaHasta)
                StrSql = StrSql & " AND (htethasta >= " & ConvFecha(FechaDesde) & " OR htethasta is null))"
                If Listaempleados <> "-1" Then
                    StrSql = StrSql & " AND ternro IN (" & Listaempleados & ")"
                End If
                OpenRecordset StrSql, rs_exportar
                Do While Not rs_exportar.EOF
                   Linea = rs_exportar!estrdabr & Separador & rs_exportar!estrcodext
                   ArchExp.Write Linea
                    ArchExp.writeline ""
                   rs_exportar.MoveNext
                Loop
                ArchExp.Close
                
          Case 10: 'tipo licencias
                StrSql = " SELECT DISTINCT tipdia.tdnro,tddesc,tdsigla from tipdia "
                StrSql = StrSql & " INNER JOIN emp_lic ON emp_lic.tdnro = tipdia.tdnro"
                StrSql = StrSql & " WHERE (elfechadesde <= " & ConvFecha(FechaHasta)
                StrSql = StrSql & " AND (elfechahasta >= " & ConvFecha(FechaDesde) & " OR  elfechahasta is null))"
                If Listaempleados <> "-1" Then
                    StrSql = StrSql & " AND  empleado IN (" & Listaempleados & ")"
                End If
                OpenRecordset StrSql, rs_exportar
                Do While Not rs_exportar.EOF
                   Linea = rs_exportar!tdnro & Separador & rs_exportar!tddesc & Separador & rs_exportar!tdsigla
                   ArchExp.Write Linea
                    ArchExp.writeline ""
                   rs_exportar.MoveNext
                Loop
                ArchExp.Close
           Case 11: 'sector
                StrSql = " SELECT distinct estructura.estrnro,estructura.estrdabr "
                StrSql = StrSql & " FROM estructura "
                StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.estrnro= estructura.estrnro "
                StrSql = StrSql & " AND estructura.tenro =  " & sectores
                StrSql = StrSql & " AND (htetdesde <= " & ConvFecha(FechaHasta)
                StrSql = StrSql & " AND (htethasta >= " & ConvFecha(FechaDesde) & " OR htethasta is null))"
                If Listaempleados <> "-1" Then
                    StrSql = StrSql & " AND ternro IN (" & Listaempleados & ")"
                End If
                OpenRecordset StrSql, rs_exportar
                Do While Not rs_exportar.EOF
                   Linea = rs_exportar!Estrnro & Separador & rs_exportar!estrdabr
                   ArchExp.Write Linea
                    ArchExp.writeline ""
                   rs_exportar.MoveNext
                Loop
                ArchExp.Close
           Case 12:  ' causas de baja
                    StrSql = " SELECT DISTINCT causa.caunro,caudes FROM fases"
                    StrSql = StrSql & " INNER JOIN causa ON causa.caunro = fases.caunro "
                    StrSql = StrSql & " WHERE bajfec <= " & ConvFecha(FechaHasta)
                    StrSql = StrSql & " AND bajfec >= " & ConvFecha(FechaDesde)
                    If Listaempleados <> "-1" Then
                        StrSql = StrSql & " AND empleado IN (" & Listaempleados & ")"
                    End If
                   OpenRecordset StrSql, rs_exportar
                Do While Not rs_exportar.EOF
                   Linea = rs_exportar!caunro & Separador & rs_exportar!caudes
                   ArchExp.Write Linea
                    ArchExp.writeline ""
                   rs_exportar.MoveNext
                Loop
                ArchExp.Close
           Case 13: ' licencias
           
                 StrSql = " SELECT empresa.empnro, emp_lic.empleado, tdnro,emp_lic.elfechadesde,emp_lic.elfechahasta,emp_lic.emp_licnro, empleg,empresa.estrnro "
                 StrSql = StrSql & " FROM emp_lic "
                 StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.ternro = emp_lic.empleado and tenro = " & empresas
                 StrSql = StrSql & " AND htetdesde <= " & ConvFecha(FechaHasta)
                 StrSql = StrSql & " AND (htethasta >= " & ConvFecha(FechaDesde) & " OR htethasta is null)"
                 StrSql = StrSql & " INNER JOIN  empresa ON empresa.estrnro= his_estructura.estrnro"
                 StrSql = StrSql & " INNER JOIN empleado ON emp_lic.empleado = empleado.ternro "
                 StrSql = StrSql & " WHERE elfechadesde <= " & ConvFecha(FechaHasta)
                 StrSql = StrSql & " AND  elfechahasta >= " & ConvFecha(FechaDesde)
                 If Listaempleados <> "-1" Then
                    StrSql = StrSql & " AND emp_lic.empleado IN (" & Listaempleados & ")"
                 End If
                 StrSql = StrSql & " ORDER BY empnro,empleado,tdnro "
                  Flog.writeline "licencias:" & StrSql
                 OpenRecordset StrSql, rs_exportar
                 Do While Not rs_exportar.EOF
                    If rs_exportar!tdnro <> licVac Then
                        Linea = rs_exportar!Estrnro & Separador & rs_exportar!empleg & Separador & Year(rs_exportar!elfechadesde) & Separador & rs_exportar!tdnro & Separador & rs_exportar!elfechadesde & Separador & rs_exportar!elfechahasta
                        ArchExp.Write Linea
                        ArchExp.writeline ""
                    Else
                        
                        StrSql = " SELECT vacanio FROM emp_lic "
                        StrSql = StrSql & " INNER JOIN  lic_vacacion ON lic_vacacion.emp_licnro = emp_lic.emp_licnro "
                        StrSql = StrSql & " INNER JOIN vacacion on vacacion.vacnro = lic_vacacion.vacnro "
                        StrSql = StrSql & " Where emp_lic.emp_licnro = " & rs_exportar!emp_licnro
                        OpenRecordset StrSql, rs_exportaraux
                        If Not rs_exportaraux.EOF Then
                            Linea = rs_exportar!Estrnro & Separador & rs_exportar!empleg & Separador & rs_exportar!tdnro & Separador & Year(rs_exportaraux!vacanio) & Separador & rs_exportar!elfechadesde & Separador & rs_exportar!elfechahasta
                            ArchExp.Write Linea
                            ArchExp.writeline ""
                        End If
                        rs_exportaraux.Close
                        
                    End If
                    rs_exportar.MoveNext
                 Loop
    End Select
    rs_exportar.Close
    
End Sub

Public Sub ExpEmp(ByVal bpronro As Long, ByVal Parametros As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del reporte de Exportacion de Citricos
' Autor      : RCH
' Fecha      : 27/09/2006
' Modificado :
' --------------------------------------------------------------------------------------------

Dim ArregloParametros

Dim todos As Integer
Dim fecha_estruc As Date
Dim rs_emple As New ADODB.Recordset
Dim rs_Tercero As New ADODB.Recordset
Dim rs_aux As New ADODB.Recordset
Dim rs_Direccion As New ADODB.Recordset

Dim directorio As String
Dim Nombre_Arch(11) As String
Dim fs1
Dim carpeta
Dim totalEmpleados
Dim cantRegistros

Dim tipo As Integer
Dim empleados As String
Dim activoinactivo As Integer
Dim DirDefault As String

Dim UsaEncabezado As Boolean
Dim FechaActual As String

Listaempleados = "0"

On Error GoTo CE

TiempoAcumulado = GetTickCount

'----------------------------------------------------------------------------
' Levanto cada parametro por separado, el separador de parametros es "@"
'----------------------------------------------------------------------------

Flog.writeline "levantando parametros" & Parametros
If Not IsNull(Parametros) Then

       
    If Len(Parametros) >= 1 Then
    
        ArregloParametros = Split(Parametros, "@")
           
        FechaDesde = ArregloParametros(0)
        Flog.writeline "Fecha Desde : " & FechaDesde
        
        FechaHasta = ArregloParametros(1)
        Flog.writeline "Fecha Hasta : " & FechaHasta
              
        tipo = ArregloParametros(2)
        Flog.writeline "Tipo : " & tipo

        If tipo = 3 Then
            todos = -1
            activoinactivo = ArregloParametros(3)
            Flog.writeline "Activo/inactivo" & activoinactivo
        Else
            todos = 0
        End If
    End If

Else
    Flog.writeline "parametros nulos"
End If
Flog.writeline "terminó de levantar los parametros"

'----------------------------------------------------------------------------
' Directorio de exportacion
'----------------------------------------------------------------------------


StrSql = "SELECT sis_dirsalidas FROM sistema "
If rs.State = adStateOpen Then rs.Close
OpenRecordset StrSql, rs_aux
If Not rs_aux.EOF Then
   directorio = Trim(rs_aux!sis_dirsalidas)
    If "\" <> CStr(Right(directorio, 1)) Then
        directorio = directorio & "\"
    End If
End If
rs_aux.Close

'---------------------------------------------------------------------------
'Agrego la carpeta por usuario
directorio = directorio & "porUsr\" & iduser
If Not fs.FolderExists(directorio) Then
    Set carpeta = fs.CreateFolder(directorio)
End If
'---------------------------------------------------------------------------

'----------------------------------------------------------------------------
' Configuración
'----------------------------------------------------------------------------

cargo = 4
categorias = 3
conceptos = ""
contrato = 18
convenio = 19
costos = 5
empresas = 10
DomicilioEmp = 6
tipolicencia = 0
sectores = 2
licVac = 2
tipliquidacion = 32

StrSql = "SELECT * FROM ConfrepAdv WHERE repnro = " & 490

OpenRecordset StrSql, rs_aux

Do While Not rs_aux.EOF
    Select Case (rs_aux!confnrocol)
        Case 1:
               If Not IsNull(rs_aux!confval) Then
                    directorio = directorio & Trim(rs_aux!confval)
                Else
                     Flog.writeline Espacios(Tabulador * 1) & "El modelo no tiene configurada la carpeta destino. El archivo será generado en el directorio default"
                End If
                'Obtengo los datos del separador
                chrsep = rs_aux!confval3
                Separador = buscarseparador(chrsep)
                Flog.writeline Espacios(Tabulador * 0) & "Separador seleccionado: " & Separador
                UsaEncabezado = rs_aux!confval2
        Case 2: cargo = rs_aux!confval
                Nombre_Arch(2) = rs_aux!confetiq
        Case 3: categorias = rs_aux!confval
                Nombre_Arch(3) = rs_aux!confetiq
        Case 4:
                If rs_aux!confval2 = 0 Then
                    conceptos = rs_aux!confval
                Else
                    conceptos = "0"
                End If
                Nombre_Arch(4) = rs_aux!confetiq
        Case 5: contrato = rs_aux!confval
                Nombre_Arch(5) = rs_aux!confetiq
        Case 6: convenio = rs_aux!confval
                Nombre_Arch(6) = rs_aux!confetiq
        Case 7: costos = rs_aux!confval
                Nombre_Arch(7) = rs_aux!confetiq
        Case 8: empresas = rs_aux!confval
                Nombre_Arch(8) = rs_aux!confetiq
        Case 9: DomicilioEmp = rs_aux!confval
                Nombre_Arch(9) = rs_aux!confetiq
                cuil = rs_aux!confval2
        Case 10: tipolicencia = rs_aux!confval
                Nombre_Arch(10) = rs_aux!confetiq
        Case 11: sectores = rs_aux!confval
                Nombre_Arch(11) = rs_aux!confetiq
        Case 12: licVac = rs_aux!confval
        Case 13: tipliquidacion = rs_aux!confval
     End Select
        rs_aux.MoveNext
        
Loop


rs_aux.Close

On Error Resume Next


'Se agrega el numero de proceso y fecha.
FechaActual = Format(Now, "yyyymmdd")
If Not fs.FolderExists(directorio) Then
    Set carpeta = fs.CreateFolder(directorio)
End If
directorio = directorio & "\Exportacion-Progress" & "-" & NroProcesoBatch & "-" & FechaActual

  If Not fs.FolderExists(directorio) Then
   Set carpeta = fs.CreateFolder(directorio)
  End If
   Flog.writeline Espacios(Tabulador * 1) & "Se crea la carpeta destino." & directorio

'desactivo el manejador de errores

On Error GoTo 0

If todos = 0 Then
    StrSql = "SELECT * FROM batch_empleado "
    StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = batch_empleado.ternro "
    StrSql = StrSql & " WHERE bpronro = " & bpronro
Else
    StrSql = "SELECT * FROM empleado "
    If activoinactivo = 1 Then
        StrSql = StrSql & " WHERE empest = -1 "
    End If
    If activoinactivo = 2 Then
        StrSql = StrSql & " WHERE empest = 0 "
    End If
End If

OpenRecordset StrSql, rs_emple

cantRegistros = rs_emple.RecordCount
totalEmpleados = rs_emple.RecordCount
If todos = 0 Then
    Do While Not rs_emple.EOF
        Listaempleados = Listaempleados & "," & rs_emple!Ternro
        rs_emple.MoveNext
    Loop
Else
    Listaempleados = "-1"
End If
StrSql = "UPDATE batch_proceso SET bprcprogreso = 0  ,bprcestado = 'Procesando' " & _
            ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
            ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & bpronro

objConn.Execute StrSql, , adExecuteNoRecords

Flog.writeline "Empiezo a generar los archivos "
Encab = "Numero_Estructura;Descripcion_Estructura"
Set fs = CreateObject("Scripting.FileSystemObject")
Call Exportar(UsaEncabezado, 2, directorio & "\" & Nombre_Arch(2), Encab)
Call Exportar(UsaEncabezado, 3, directorio & "\" & Nombre_Arch(3), Encab)
Encab = "Codigo;Concepto;Formula;Cod.Parametro"
Call Exportar(UsaEncabezado, 4, directorio & "\" & Nombre_Arch(4), Encab)
Encab = "Numero_Estructura;Descripcion_Estructura"
Call Exportar(UsaEncabezado, 5, directorio & "\" & Nombre_Arch(5), Encab)
Call Exportar(UsaEncabezado, 6, directorio & "\" & Nombre_Arch(6), Encab)
Encab = "Cod_Ext_centro;Descripcion_Ext_Estructura"
Call Exportar(UsaEncabezado, 7, directorio & "\" & Nombre_Arch(7), Encab)
Encab = "Descripcion_Estructura;Codigo_Externo"
Call Exportar(UsaEncabezado, 8, directorio & "\" & Nombre_Arch(8), Encab)
Encab = "Tipo Licencia;Descripción;Sigla"
Call Exportar(UsaEncabezado, 10, directorio & "\" & Nombre_Arch(10), Encab)
Encab = "Numero_Estructura;Descripcion_Estructura"
Call Exportar(UsaEncabezado, 11, directorio & "\" & Nombre_Arch(11), Encab)
Encab = "Código;Descripción"
Call Exportar(UsaEncabezado, 12, directorio & "\" & "motibaja", Encab)
Encab = "Empresa;Legajo;Periodo;Tipo de Licencia;Fecha de Inicio;Fecha de Fin"
Call Exportar(UsaEncabezado, 13, directorio & "\" & "licencias", Encab)
Encab = "Empresa;EMPLEADO;Division_Liquidacion;Centros_de_Costo;APELLIDO;NOMBRE;DOMICILIO1;LOCALIDAD1;CODIGOPOSTAL;FECHANAC;SEXO;FECHAULTALTA;FECHAALTA;CAUSA;Convenio;MODELOESTR;CUIL;ESTADO;FECHABAJA;Categoria;Contrato;Cargo"

Set ArchExp = fs.CreateTextFile(directorio & "\legajos.csv", True)
If UsaEncabezado Then
     ArchExp.Write Encab
     ArchExp.writeline ""
End If
rs_emple.MoveFirst
Do While Not rs_emple.EOF
        Linea = ""
     'empresa
      StrSql = " SELECT estructura.estrcodext  "
                StrSql = StrSql & " FROM estructura "
                StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.estrnro= estructura.estrnro "
                StrSql = StrSql & " AND his_estructura.tenro =  " & empresas
                StrSql = StrSql & " AND (htetdesde <= " & ConvFecha(FechaHasta)
                StrSql = StrSql & " AND (htethasta >= " & ConvFecha(FechaDesde) & " OR htethasta is null))"
                StrSql = StrSql & " AND ternro = " & rs_emple!Ternro
                OpenRecordset StrSql, rs_aux
                If Not rs_aux.EOF Then
                    Linea = Linea & rs_aux!estrcodext
                End If
                Linea = Linea & Separador
                rs_aux.Close
                
       'legajo
              Linea = Linea & rs_emple("empleg") & Separador
       'sector
         StrSql = " SELECT estructura.estrnro "
                StrSql = StrSql & " FROM estructura "
                StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.estrnro= estructura.estrnro "
                StrSql = StrSql & " AND estructura.tenro =  " & sectores
                StrSql = StrSql & " AND (htetdesde <= " & ConvFecha(FechaHasta)
                StrSql = StrSql & " AND (htethasta >= " & ConvFecha(FechaDesde) & " OR htethasta is null))"
                StrSql = StrSql & " AND ternro = " & rs_emple!Ternro
                OpenRecordset StrSql, rs_aux
                If Not rs_aux.EOF Then
                    Linea = Linea & rs_aux!Estrnro
                End If
                Linea = Linea & Separador
                rs_aux.Close
        'centro de costo
        
                StrSql = " SELECT estructura.estrnro,estructura.estrcodext "
                StrSql = StrSql & " FROM estructura "
                StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.estrnro= estructura.estrnro "
                StrSql = StrSql & " AND estructura.tenro =  " & costos
                StrSql = StrSql & " AND (htetdesde <= " & ConvFecha(FechaHasta)
                StrSql = StrSql & " AND (htethasta >= " & ConvFecha(FechaDesde) & " OR htethasta is null))"
                StrSql = StrSql & " AND ternro = " & rs_emple!Ternro
                 OpenRecordset StrSql, rs_aux
                If Not rs_aux.EOF Then
                    Linea = Linea & rs_aux!estrcodext
                End If
                Linea = Linea & Separador
                rs_aux.Close
          'Apellido y nombre
                 StrSql = "SELECT * FROM tercero WHERE ternro = " & rs_emple!Ternro
                OpenRecordset StrSql, rs_Tercero
                 If Not rs_Tercero.EOF Then
                    Linea = Linea & rs_Tercero!terape & " " & rs_Tercero!terape2 & Separador & rs_Tercero!ternom & " " & rs_Tercero!ternom2 & Separador
                 End If
            'domicilio
                StrSql = " SELECT * FROM detdom " & _
                         " INNER JOIN cabdom ON detdom.domnro = cabdom.domnro " & _
                         " LEFT JOIN zona ON zona.zonanro = detdom.zonanro " & _
                         " WHERE cabdom.ternro = " & rs_emple!Ternro & " AND " & _
                         " cabdom.tidonro =  " & DomicilioEmp
                         
                          OpenRecordset StrSql, rs_Direccion
                           Flog.writeline Espacios(Tabulador * 1) & "Domicilio:" & StrSql
                  If Not rs_Direccion.EOF Then
                      If Not EsNulo(rs_Direccion!Calle) Then
                                Flog.writeline Espacios(Tabulador * 1) & "Calle:" & rs_Direccion!Calle
                                 Linea = Linea & rs_Direccion!Calle
                      End If
                      If Not EsNulo(rs_Direccion!Nro) Then
                                  Flog.writeline Espacios(Tabulador * 1) & "Nro:" & rs_Direccion!Nro
                                 Linea = Linea & " " & rs_Direccion!Nro
                       End If
                      If Not EsNulo(rs_Direccion!Piso) Then
                                  Flog.writeline Espacios(Tabulador * 1) & "Piso:" & rs_Direccion!Piso
                                 Linea = Linea & " " & rs_Direccion!Piso
                      End If
                      If Not EsNulo(rs_Direccion!Piso) Then
                                Flog.writeline Espacios(Tabulador * 1) & "oficdepto:" & rs_Direccion!oficdepto
                             Linea = Linea & " " & rs_Direccion!oficdepto
                      End If
                  End If
                   Flog.writeline Espacios(Tabulador * 1) & "linea:" & Linea
                  Linea = Linea & Separador
                    Flog.writeline Espacios(Tabulador * 1) & "localidad:"
            'localidad
                If Not rs_Direccion.EOF Then
                  If Not EsNulo(rs_Direccion!locnro) Then
                        StrSql = " SELECT locdesc FROM localidad " & _
                                " WHERE localidad.locnro= " & rs_Direccion!locnro
                        OpenRecordset StrSql, rs_aux
                        If Not rs_aux.EOF Then
                            Linea = Linea & rs_aux!locdesc
                        End If
                        rs_aux.Close
                        
                  End If
                 End If
                  Linea = Linea & Separador
                  'rs_Direccion.Close
                   Flog.writeline Espacios(Tabulador * 1) & "codigo postal:"
              'Codigo Postal
                  If Not rs_Direccion.EOF Then
                    Linea = Linea & rs_Direccion!codigopostal
                  End If
                  Linea = Linea & Separador
              ' fecha de nacimiento
               If Not EsNulo(rs_Tercero!terfecnac) Then
                   Linea = Linea & CStr(rs_Tercero!terfecnac) & Separador
               Else
                    Linea = Linea & Separador
               End If
              ' sexo
                    If rs_Tercero!tersex = -1 Then
                        Linea = Linea & "Masculino"
                    Else
                        Linea = Linea & "Femenino"
                    End If
                    Linea = Linea & Separador
              'Fecha de inicio temporada
                StrSql = " SELECT fasnro,altfec, bajfec FROM fases" & _
                        " WHERE empleado = " & rs_emple!Ternro & _
                        " ORDER BY altfec Desc "
                        OpenRecordset StrSql, rs_aux
                        nrofase = 0
                        Fechaegreso = ""
                        If Not rs_aux.EOF Then
                            nrofase = rs_aux!fasnro
                            If Not IsNull(rs_aux!bajfec) Then
                                Fechaegreso = CStr(rs_aux!bajfec)
                            Else
                                Fechaegreso = ""
                            End If
                            Linea = Linea & CStr(rs_aux!altfec)
                        End If
                        Linea = Linea & Separador
                        rs_aux.Close
               'Fecha de Ingreso
               StrSql = " SELECT altfec FROM fases" & _
                        " WHERE empleado = " & rs_emple!Ternro & _
                        " ORDER BY altfec "
                        OpenRecordset StrSql, rs_aux
                        If Not rs_aux.EOF Then
                            Linea = Linea & CStr(rs_aux!altfec)
                        End If
                        Linea = Linea & Separador
                        rs_aux.Close
                        
                
                'Motivo egreso
                
                    StrSql = " SELECT  causa.caunro FROM fases"
                    StrSql = StrSql & " INNER JOIN causa ON causa.caunro = fases.caunro "
                    StrSql = StrSql & " WHERE fases.fasnro = " & nrofase
                    StrSql = StrSql & " AND empleado = " & rs_emple!Ternro
                    OpenRecordset StrSql, rs_aux
                    If Not rs_aux.EOF Then
                          
                        Linea = Linea & CStr(rs_aux!caunro)
                    End If
                        Linea = Linea & Separador
                        rs_aux.Close
                'convenio
                StrSql = " SELECT estructura.estrnro"
                StrSql = StrSql & " FROM estructura "
                StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.estrnro= estructura.estrnro "
                StrSql = StrSql & " AND estructura.tenro =  " & convenio
                StrSql = StrSql & " AND (htetdesde <= " & ConvFecha(FechaHasta)
                StrSql = StrSql & " AND (htethasta >= " & ConvFecha(FechaDesde) & " OR htethasta is null))"
                StrSql = StrSql & " AND ternro =" & rs_emple!Ternro
                OpenRecordset StrSql, rs_aux
                    If Not rs_aux.EOF Then
                          
                        Linea = Linea & CStr(rs_aux!Estrnro)
                    End If
                        Linea = Linea & Separador
                        rs_aux.Close
                 'Tipo liquidación
                  StrSql = " SELECT estructura.estrcodext"
                StrSql = StrSql & " FROM estructura "
                StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.estrnro= estructura.estrnro "
                StrSql = StrSql & " AND estructura.tenro =  " & tipliquidacion
                StrSql = StrSql & " AND (htetdesde <= " & ConvFecha(FechaHasta)
                StrSql = StrSql & " AND (htethasta >= " & ConvFecha(FechaDesde) & " OR htethasta is null))"
                StrSql = StrSql & " AND ternro = " & rs_emple!Ternro
                OpenRecordset StrSql, rs_aux
                    If Not rs_aux.EOF Then
                          
                        Linea = Linea & CStr(rs_aux!estrcodext)
                    End If
                        Linea = Linea & Separador
                        rs_aux.Close
                 'Cuil
                 
                 StrSql = " SELECT cuil.nrodoc FROM tercero " & _
                         " INNER JOIN ter_doc cuil ON (tercero.ternro = cuil.ternro AND cuil.tidnro = " & cuil & ") " & _
                         " WHERE tercero.ternro= " & rs_emple!Ternro
                OpenRecordset StrSql, rs_aux
              
                  
                    If Not rs_aux.EOF Then
                       
                        Linea = Linea & Left(CStr(rs_aux!NroDoc), 13)
                    End If
                        Linea = Linea & Separador
                        rs_aux.Close
                       'liquida
                  If rs_emple!empest = -1 Then
                    Linea = Linea & "Activo"
                Else
                    Linea = Linea & "Inactivo"
                End If
                 Linea = Linea & Separador
                 
                'Fecha Egreso
                Linea = Linea & Fechaegreso
                Linea = Linea & Separador
                    
                   ' Categoría
                 StrSql = " SELECT estructura.estrnro "
                StrSql = StrSql & " FROM estructura "
                StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.estrnro= estructura.estrnro "
                StrSql = StrSql & " AND estructura.tenro =  " & categorias
                StrSql = StrSql & " AND (htetdesde <= " & ConvFecha(FechaHasta)
                StrSql = StrSql & " AND (htethasta >= " & ConvFecha(FechaDesde) & " OR htethasta is null))"
                StrSql = StrSql & " AND ternro = " & rs_emple!Ternro
                 OpenRecordset StrSql, rs_aux
                 If Not rs_aux.EOF Then
                       
                        Linea = Linea & rs_aux!Estrnro
                    End If
                        Linea = Linea & Separador
                        rs_aux.Close
                        
                ' Contrato
                
                StrSql = " SELECT estructura.estrnro"
                StrSql = StrSql & " FROM estructura "
                StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.estrnro= estructura.estrnro "
                StrSql = StrSql & " AND estructura.tenro =  " & contrato
                StrSql = StrSql & " AND (htetdesde <= " & ConvFecha(FechaHasta)
                StrSql = StrSql & " AND (htethasta >= " & ConvFecha(FechaDesde) & " OR htethasta is null))"
                StrSql = StrSql & " AND ternro = " & rs_emple!Ternro
                  OpenRecordset StrSql, rs_aux
                 If Not rs_aux.EOF Then
                       
                        Linea = Linea & rs_aux!Estrnro
                    End If
                Linea = Linea & Separador
                rs_aux.Close
               
                    
                 'Cargo
                StrSql = " SELECT estructura.estrnro "
                StrSql = StrSql & " FROM estructura "
                StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.estrnro= estructura.estrnro "
                StrSql = StrSql & " AND estructura.tenro =  " & cargo
                StrSql = StrSql & " AND (htetdesde <= " & ConvFecha(FechaHasta)
                StrSql = StrSql & " AND (htethasta >= " & ConvFecha(FechaDesde) & " OR htethasta is null))"
                StrSql = StrSql & " AND ternro = " & rs_emple!Ternro
                   OpenRecordset StrSql, rs_aux
                  If Not rs_aux.EOF Then
                       
                        Linea = Linea & rs_aux!Estrnro
                    End If
          
                        rs_aux.Close
                        
            
                rs_Direccion.Close
                rs_Tercero.Close
                ArchExp.Write Linea
                ArchExp.writeline ""
                
     rs_emple.MoveNext
       'Actualizo el estado del proceso
        TiempoAcumulado = GetTickCount
              
        cantRegistros = cantRegistros - 1
           
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Fix(((totalEmpleados - cantRegistros) * 100) / totalEmpleados) & _
                 ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                 ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & bpronro
        objConn.Execute StrSql, , adExecuteNoRecords
Loop
ArchExp.Close
'fs.Nothing

    
     
    



'----------------------------------------------------------------
' Borrar los Empleados de la tabla batch_proceso
'----------------------------------------------------------------

StrSql = " DELETE FROM batch_empleado "
StrSql = StrSql & " WHERE bpronro = " & bpronro
objConn.Execute StrSql, , adExecuteNoRecords


Exit Sub

CE:
    Flog.writeline "=================================================================="
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="
    MyRollbackTrans
    MyBeginTrans
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Fix(((totalEmpleados - cantRegistros) * 100) / totalEmpleados) & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
             "' WHERE bpronro = " & bpronro
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    MyCommitTrans
    
    HuboError = True
    Flog.writeline " Error: " & Err.Description

End Sub

