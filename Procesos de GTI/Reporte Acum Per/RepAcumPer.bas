Attribute VB_Name = "RepAcumPer"
Option Explicit

'Global Const Version = 1.01
'Global Const FechaVersion = "20/11/2006"
'Global Const UltimaModificacion = " " 'Martin Ferraro - Version Inicial


'Global Const Version = "1.02"
'Global Const FechaVersion = "03/11/2011"
'Global Const UltimaModificacion = " " 'FGZ - Encriptacion de string de conexion

Global Const Version = "1.03"
Global Const FechaVersion = "07/11/2011"
Global Const UltimaModificacion = " " 'Carmen Quintero - Se agrego la lista de procesos seleccionados
                                      'al momento de realizar SUM(dgticant) en la consulta principal

'---------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------
Dim fs, f

Dim NroProceso As Long

Global Path As String
Global NArchivo As String
Global Rta
Global HuboErrores As Boolean
Global tiene_turno As Boolean
Global Nro_Turno As Long
Global Tipo_Turno As Integer
Global Tiene_Justif As Boolean
Global nro_justif As Long
Global justif_turno As Boolean
Global p_turcomp As Boolean
Global Nro_Grupo As Long
Global Nro_fpgo As Integer
Global Fecha_Inicio As Date
Global P_Asignacion  As Boolean
Global Trabaja     As Boolean ' Indica si trabaja para ese dia
Global Orden_Dia As Integer
Global Nro_Dia As Integer
Global Nro_Subturno As Integer
Global Dia_Libre As Boolean
Global Dias_trabajados As Integer
Global Dias_laborables As Integer

Global EmpErrores As Boolean

Global tenro1 As Long
Global estrnro1 As Long
Global tenro2 As Long
Global estrnro2 As Long
Global tenro3 As Long
Global estrnro3 As Long
Global PerDesde As Long
Global PerHasta As Long
Global fecEstr As Date
Global tituloReporte As String
Global ordenBase As Long

'Maxima Cantidad de Columnas
Global Const maxCol = 20
'array  que almacena en cada componente que columna del confrep buscar
Global arrColConfrep(20) As Long
'array  que almacena en cada componente la etiqueta de cada columna
Global arrEtiqConfrep(20) As String
'array  que almacena en cada componente el resultado de acum en cada columna
Global arrValor(20) As Double
'Almacena cual es la maxima columna que se utiliza en los array
Global maxColUsed As Long
'record set del confrep
Global rsConfRep As New ADODB.Recordset

'FGZ - 28/10/2011 ---------------------
Global Subturno_Genera As Integer



Private Sub Main()

Dim NombreArchivo As String
Dim Directorio As String
Dim CArchivos
Dim Archivo
Dim Folder
Dim strcmdLine As String
Dim Nombre_Arch As String

Dim StrSql As String
Dim objRs As New ADODB.Recordset
Dim objRs2 As New ADODB.Recordset
Dim FechaDesde
Dim FechaHasta
Dim Fecha
Dim tipoDepuracion
Dim historico As Boolean
Dim param
Dim rsEmpl As New ADODB.Recordset
Dim PID As String
Dim arrPronro
Dim TiempoInicialProceso
Dim TiempoAcumulado
Dim totalEmpleados
Dim cantRegistros
Dim Ternro As Long
Dim ListaPer
Dim errConfrep As String

Dim ArrParametros
Dim Parametros As String

Dim listapronro
Dim Pronro
Dim I


'    strCmdLine = Command()
'    ArrParametros = Split(strCmdLine, " ", -1)
'    If UBound(ArrParametros) > 0 Then
'        If IsNumeric(ArrParametros(0)) Then
'            NroProceso = ArrParametros(0)
'            Etiqueta = ArrParametros(1)
'        Else
'            Exit Sub
'        End If
'    Else
'        If IsNumeric(strCmdLine) Then
'            NroProceso = strCmdLine
'        Else
'            Exit Sub
'        End If
'    End If


    
    strcmdLine = Command()
    ArrParametros = Split(strcmdLine, " ", -1)
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
            If IsNumeric(strcmdLine) Then
                NroProceso = strcmdLine
            Else
                Exit Sub
            End If
        End If
    End If

    Call CargarConfiguracionesBasicas
    'OpenConnection strconexion, objConn
    
    ' seteo del nombre del archivo de log
    Nombre_Arch = PathFLog & "RepAcumPer" & "-" & NroProceso & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    TiempoInicialProceso = GetTickCount
   
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
    
    TiempoInicialProceso = GetTickCount
    depurar = False
    HuboErrores = False
    
    On Error GoTo CE:

    'Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaVersion
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-------------------------------------------------"
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "Inicio :" & Now
  
    
    'Nombre_Arch = PathFLog & "RepAcumPer" & "-" & NroProceso & ".log"
    'Set fs = CreateObject("Scripting.FileSystemObject")
    'Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    'Flog.writeline "Inicio Proceso Reporte Resumen Acumulado Periodo : " & Now
    'Flog.writeline "Cambio el estado del proceso a Procesando"
    
    'Obtengo el Process ID
    'PID = GetCurrentProcessId
    'Flog.writeline "-------------------------------------------------"
    'Flog.writeline "Version                  : " & Version
    'Flog.writeline "Fecha Ultima Modificacion: " & FechaVersion
    'Flog.writeline "PID                      : " & PID
    'Flog.writeline "-------------------------------------------------"
    'Flog.writeline
    'Flog.writeline Espacios(Tabulador * 0) & "Inicio :" & Now
    
    TiempoAcumulado = GetTickCount
    
    'StrSql = "UPDATE batch_proceso SET bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "', bprcpid = " & PID & ", bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcprogreso = 0 WHERE bpronro = " & NroProceso
    StrSql = "UPDATE batch_proceso SET bprcpid = " & PID & ", bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords

    Flog.writeline "Obtengo los datos del proceso"
       
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProceso
    StrSql = StrSql & " AND btprcnro = 142"
    OpenRecordset StrSql, objRs
    
    If Not objRs.EOF Then
    
       'Obtengo los parametros del proceso
       Parametros = objRs!bprcparam
       
       Flog.writeline Espacios(Tabulador * 1) & "Parametros del proceso: " & Parametros
       
       FechaDesde = CDate(objRs!bprcfecdesde)
       FechaHasta = CDate(objRs!bprcfechasta)
       fecEstr = FechaHasta
              
       ArrParametros = Split(Parametros, "@")
       
       PerDesde = CLng(ArrParametros(0))
       PerHasta = CLng(ArrParametros(1))
       tenro1 = CLng(ArrParametros(2))
       estrnro1 = CLng(ArrParametros(3))
       tenro2 = CLng(ArrParametros(4))
       estrnro2 = CLng(ArrParametros(5))
       tenro3 = CLng(ArrParametros(6))
       estrnro3 = CLng(ArrParametros(7))
       tituloReporte = ArrParametros(8)
       
       'Obtengo la lista de procesos
       listapronro = ArrParametros(9)

       Flog.writeline Espacios(Tabulador * 1) & "Periodo Desde = " & PerDesde
       Flog.writeline Espacios(Tabulador * 1) & "Periodo Hasta = " & PerHasta
       Flog.writeline Espacios(Tabulador * 1) & "Tipo Estructura 1 = " & tenro1
       Flog.writeline Espacios(Tabulador * 1) & "Estructura 1 = " & estrnro1
       Flog.writeline Espacios(Tabulador * 1) & "Tipo Estructura 2 = " & tenro2
       Flog.writeline Espacios(Tabulador * 1) & "Estructura 2 = " & estrnro2
       Flog.writeline Espacios(Tabulador * 1) & "Tipo Estructura 3 = " & tenro3
       Flog.writeline Espacios(Tabulador * 1) & "Estructura 3 = " & estrnro3
       Flog.writeline Espacios(Tabulador * 1) & "Procesos = " & listapronro
       
       'EMPIEZA EL PROCESO
       
        Call CargarConfrep(errConfrep)
        If Len(errConfrep) <> 0 Then
            Flog.writeline Espacios(Tabulador * 1) & errConfrep
            GoTo CE
        End If

       'Armo la lista de periodos
       ListaPer = listaPeriodo(PerDesde, PerHasta)
       
       'Obtengo los empleados
       CargarEmpleados NroProceso, rsEmpl
       totalEmpleados = rsEmpl.RecordCount
       cantRegistros = totalEmpleados
       
       ordenBase = 0
       
       'Proceso cada empleado
       Do Until rsEmpl.EOF
          EmpErrores = False
          Ternro = rsEmpl!Ternro
          
          Flog.writeline
          Flog.writeline "Generando datos para el empleado " & Ternro
          
          Flog.writeline "Lista de procesos " & listapronro
                             
          Call generarDatosEmpleadoPeriodos(Ternro, ListaPer, listapronro)
          
          cantRegistros = cantRegistros - 1
          TiempoAcumulado = GetTickCount
          
          StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Fix(((totalEmpleados - cantRegistros) * 100) / totalEmpleados) & _
                      ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                      ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
                
          'objConn.Execute StrSql, , adExecuteNoRecords
          objConnProgreso.Execute StrSql, , adExecuteNoRecords
          
          'Si se proceso el empleado correctamente lo borro
          If Not EmpErrores Then
              StrSql = " DELETE FROM batch_empleado "
              StrSql = StrSql & " WHERE bpronro = " & NroProceso
              StrSql = StrSql & " AND ternro = " & Ternro
              objConn.Execute StrSql, , adExecuteNoRecords
          End If
          
          rsEmpl.MoveNext
       Loop
    
    Else
        Exit Sub
    End If
   
    'Actualizo el estado del proceso
    If Not HuboErrores Then
       StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
    Else
       StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProceso
    End If
    
    objConn.Execute StrSql, , adExecuteNoRecords
    
    
Final:
    Flog.writeline Espacios(Tabulador * 0) & "Fin :" & Now
    Flog.Close
    Exit Sub
    
CE:
    HuboErrores = True

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
' Inicializa arreglos globales de acuerdo al confrep
'--------------------------------------------------------------------
Public Sub CargarConfrep(ByRef strErr As String)
Dim I As Long
Dim rsAux As New ADODB.Recordset

    
    strErr = ""
    I = 0
    
    'Busco las distintas columnas del confrep
    StrSql = "SELECT distinct confnrocol FROM confrep WHERE repnro = 183 AND conftipo = 'TH' ORDER BY confnrocol"
    
    OpenRecordset StrSql, rsConfRep
    If rsConfRep.EOF Then
        strErr = "Error. No existe columna en confrep tipo TH."
        Exit Sub
    End If
    
    'Armo los arreglos que controlan conf rep
    Do While Not (rsConfRep.EOF Or I >= maxCol)
        
        I = I + 1
        arrColConfrep(I) = rsConfRep!confnrocol
        
        'Busco la etiqueta
        StrSql = "SELECT confetiq FROM confrep WHERE repnro = 183 AND confnrocol = " & rsConfRep!confnrocol
        
        OpenRecordset StrSql, rsAux
        If Not rsAux.EOF Then
            arrEtiqConfrep(I) = IIf(EsNulo(rsAux!confetiq), "", rsAux!confetiq)
        Else
            arrEtiqConfrep(I) = ""
        End If
        rsAux.Close

        rsConfRep.MoveNext
    Loop
    
    rsConfRep.Close
    
    'Seteo la variable global que apunta a la ultima componente usada del sistema
    maxColUsed = I
    
If rsAux.State = adStateOpen Then rsAux.Close
Set rsAux = Nothing
    
    
End Sub

'--------------------------------------------------------------------
' Busca los tipos de hora con en el confrep para el nro de col
'--------------------------------------------------------------------
Public Function listaTipoHora(ByVal nrocol As Long) As String

Dim rsAux As New ADODB.Recordset
Dim Aux As String

    Aux = ""
    
    StrSql = "SELECT confval FROM confrep WHERE repnro = 183 AND confnrocol = " & nrocol
    
    OpenRecordset StrSql, rsAux
    Do While Not rsAux.EOF
        
        If Len(Aux) = 0 Then
            Aux = rsAux!confval
        Else
            Aux = Aux & "," & rsAux!confval
        End If
        
        rsAux.MoveNext
        
    Loop
    rsAux.Close
    
    listaTipoHora = Aux

If rsAux.State = adStateOpen Then rsAux.Close
Set rsAux = Nothing

End Function


'--------------------------------------------------------------------
' Se encarga de llenar el arr de valores con 0 hasta la maxima col del confrep
'--------------------------------------------------------------------
Public Sub limpiarArrValores()
Dim I As Long

    For I = 1 To maxColUsed
        arrValor(I) = 0
    Next
    
End Sub


'--------------------------------------------------------------------
' Se encarga de generar los datos
'--------------------------------------------------------------------
Sub generarDatosEmpleadoPeriodos(ByVal Ternro As Long, ByVal listPeriodos As String, ByVal listapronro As String)


Dim StrSql As String
Dim rsConsult As New ADODB.Recordset
Dim rsPeriodo As New ADODB.Recordset
Dim rsAcum As New ADODB.Recordset
Dim listProc As String
Dim EmpEstrnro1 As Long
Dim EmpEstrnro2 As Long
Dim EmpEstrnro3 As Long
Dim Nombre As String
Dim apellido As String
Dim Legajo As Long
Dim I As Long
Dim listTHoras As String
Dim tieneDatos As Boolean

Dim rsProcesos As New ADODB.Recordset
Dim l_gpadesabr As String


EmpEstrnro1 = 0
EmpEstrnro2 = 0
EmpEstrnro3 = 0
l_gpadesabr = ""

On Error GoTo MError

'------------------------------------------------------------------
'Busco los datos del empleado
'------------------------------------------------------------------
StrSql = " SELECT empleg,terape,terape2,ternom,ternom2,empfecalta,empremu "
StrSql = StrSql & " FROM empleado "
StrSql = StrSql & " WHERE ternro= " & Ternro
       
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
   Nombre = rsConsult!ternom & " " & rsConsult!ternom2
   apellido = rsConsult!terape & " " & rsConsult!terape2
   Legajo = rsConsult!EmpLeg
   Flog.writeline Espacios(Tabulador * 1) & "Empleado: " & Legajo & " - " & apellido & " " & Nombre
Else
   Flog.writeline Espacios(Tabulador * 1) & "Error al obtener los datos del empleado"
   GoTo MError
End If
rsConsult.Close

'------------------------------------------------------------------
'Busco las estructuras seleccionadas en el filtro
'------------------------------------------------------------------
If tenro1 <> 0 Then
    If estrnro1 <> 0 Then
        EmpEstrnro1 = estrnro1
    Else
        StrSql = " SELECT * FROM his_estructura WHERE ternro = " & Ternro & " AND tenro = " & tenro1
        StrSql = StrSql & " AND (htetdesde<=" & ConvFecha(fecEstr) & " AND (htethasta is null or htethasta>=" & ConvFecha(fecEstr) & "))"
        
        OpenRecordset StrSql, rsConsult
        If Not rsConsult.EOF Then
           EmpEstrnro1 = rsConsult!estrnro
        End If
        rsConsult.Close
    End If
End If


If tenro2 <> 0 Then
    If estrnro2 <> 0 Then
        EmpEstrnro2 = estrnro2
    Else
        StrSql = " SELECT * FROM his_estructura WHERE ternro = " & Ternro & " AND tenro = " & tenro2
        StrSql = StrSql & " AND (htetdesde<=" & ConvFecha(fecEstr) & " AND (htethasta is null or htethasta>=" & ConvFecha(fecEstr) & "))"
        
        OpenRecordset StrSql, rsConsult
        If Not rsConsult.EOF Then
           EmpEstrnro2 = rsConsult!estrnro
        End If
        rsConsult.Close
    End If
End If


If tenro3 <> 0 Then
    If estrnro3 <> 0 Then
        EmpEstrnro3 = estrnro3
    Else
        StrSql = " SELECT * FROM his_estructura WHERE ternro = " & Ternro & " AND tenro = " & tenro3
        StrSql = StrSql & " AND (htetdesde<=" & ConvFecha(fecEstr) & " AND (htethasta is null or htethasta>=" & ConvFecha(fecEstr) & "))"
        
        OpenRecordset StrSql, rsConsult
        If Not rsConsult.EOF Then
           EmpEstrnro3 = rsConsult!estrnro
        End If
        rsConsult.Close
    End If
End If

'------------------------------------------------------------------
'Armo la SQL procesar todos los periodos
'------------------------------------------------------------------
StrSql = " SELECT * "
StrSql = StrSql & " FROM gti_per "
StrSql = StrSql & " WHERE pgtinro IN ( " & listPeriodos & ")"
StrSql = StrSql & " ORDER BY pgtianio, pgtimes "

OpenRecordset StrSql, rsPeriodo
If rsPeriodo.EOF Then
    Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encuentran los periodos"
    HuboErrores = True
    EmpErrores = True
    rsPeriodo.Close
    Exit Sub
Else
    
    'Busco la descripcion del proceso
    StrSql = "SELECT DISTINCT gti_procacum.gpadesabr "
    StrSql = StrSql & "FROM gti_procacum LEFT JOIN gti_cab ON gti_procacum.gpanro= gti_cab.gpanro "
    StrSql = StrSql & "WHERE gti_procacum.pgtinro IN ( " & listPeriodos & ")"
    StrSql = StrSql & "AND gti_procacum.gpanro in (" & listapronro & ")"
    
    OpenRecordset StrSql, rsProcesos
    Do While Not rsProcesos.EOF
        l_gpadesabr = l_gpadesabr & " - " & rsProcesos!gpadesabr
    rsProcesos.MoveNext
    Loop
    
    rsProcesos.Close
    
    'Proceso todos los periodos
    'Do Until rsPeriodo.EOF
        Flog.writeline Espacios(Tabulador * 1) & "Procesando Periodo " & rsPeriodo!pgtinro & " " & rsPeriodo!pgtidesabr
        
        tieneDatos = False
        Call limpiarArrValores
                
        For I = 1 To maxColUsed
            Flog.writeline Espacios(Tabulador * 2) & "Buscando los tipo de horas de la columna " & arrColConfrep(I)
            
            listTHoras = listaTipoHora(arrColConfrep(I))
            If Len(listTHoras) = 0 Then
                Flog.writeline Espacios(Tabulador * 2) & "La Columna no tiene tipo de horas."
            Else
                
                'Busco lo acumulado para el empleado por la lista de procesos
                StrSql = " SELECT SUM(dgticant) suma "
                StrSql = StrSql & " FROM gti_det "
                StrSql = StrSql & " INNER JOIN gti_cab ON gti_cab.cgtinro = gti_det.cgtinro "
                StrSql = StrSql & " AND gti_cab.ternro = " & Ternro
                StrSql = StrSql & " INNER JOIN gti_procacum ON gti_procacum.gpanro = gti_cab.gpanro "
                StrSql = StrSql & " AND gti_procacum.gpanro in (" & listapronro & ")"
                StrSql = StrSql & " INNER JOIN gti_per ON gti_per.pgtinro = gti_procacum.pgtinro "
                StrSql = StrSql & " AND gti_per.pgtinro IN ( " & listPeriodos & ")"
                StrSql = StrSql & " WHERE gti_det.thnro in (" & listTHoras & ")"
                
                OpenRecordset StrSql, rsAcum
                If Not rsAcum.EOF Then
                    If rsAcum!suma <> 0 Then
                        tieneDatos = True
                        arrValor(I) = rsAcum!suma
                    End If
                End If
                rsAcum.Close

            End If

        Next
                    
        If tieneDatos Then
            Flog.writeline Espacios(Tabulador * 3) & "Insertando Datos encontrados en base."
            Call InsertarBase(Ternro, rsPeriodo!pgtinro, rsPeriodo!pgtidesabr, Legajo, apellido, Nombre, EmpEstrnro1, EmpEstrnro2, EmpEstrnro3, listapronro, l_gpadesabr)
        Else
            Flog.writeline Espacios(Tabulador * 3) & "No se encontraron datos para el empleado y periodo."
        End If
           
        'rsPeriodo.MoveNext
    'Loop
End If

rsPeriodo.Close


If rsConsult.State = adStateOpen Then rsConsult.Close
If rsPeriodo.State = adStateOpen Then rsPeriodo.Close
If rsAcum.State = adStateOpen Then rsAcum.Close
Set rsConsult = Nothing
Set rsPeriodo = Nothing
Set rsAcum = Nothing
   
Exit Sub

MError:
    Flog.writeline "Error en empleado: " & Legajo & " Error: " & Err.Description
    HuboErrores = True
    EmpErrores = True
    Exit Sub
End Sub


'--------------------------------------------------------------------
' Inserta los datos en la base en la tabla rep_resacum
'--------------------------------------------------------------------
Public Sub InsertarBase(ByVal Ternro As Long, ByVal pgtinro As Long, ByVal pgtidesabr As String, ByVal Legajo As Long, ByVal apellido As String, ByVal Nombre As String, ByVal EmpEstrnro1 As Long, ByVal EmpEstrnro2 As Long, ByVal EmpEstrnro3 As Long, ByVal gpanro As String, ByVal gpadesabr As String)
Dim I As Long

    pgtidesabr = IIf(EsNulo(pgtidesabr), "", pgtidesabr)
    Nombre = IIf(EsNulo(Nombre), "", Nombre)
    apellido = IIf(EsNulo(apellido), "", apellido)
    
    ordenBase = ordenBase + 1
    
    StrSql = " INSERT INTO rep_resacum "
    StrSql = StrSql & " (bpronro, Ternro, pgtinro,"
    StrSql = StrSql & " apellido, Nombre, Legajo,"
    StrSql = StrSql & " tenro1, tenro2, tenro3,"
    StrSql = StrSql & " estrnro1, estrnro2, estrnro3,"
    StrSql = StrSql & " titulo, pgtidesabr, orden,"
    StrSql = StrSql & " coletiq1, colvalor1, coletiq2, colvalor2,"
    StrSql = StrSql & " coletiq3, colvalor3, coletiq4, colvalor4,"
    StrSql = StrSql & " coletiq5, colvalor5, coletiq6, colvalor6,"
    StrSql = StrSql & " coletiq7, colvalor7, coletiq8, colvalor8,"
    StrSql = StrSql & " coletiq9, colvalor9, coletiq10, colvalor10,"
    StrSql = StrSql & " coletiq11, colvalor11, coletiq12, colvalor12,"
    StrSql = StrSql & " coletiq13, colvalor13, coletiq14, colvalor14,"
    StrSql = StrSql & " coletiq15, colvalor15, coletiq16, colvalor16,"
    StrSql = StrSql & " coletiq17, colvalor17, coletiq18, colvalor18,"
    StrSql = StrSql & " coletiq19, colvalor19, coletiq20, colvalor20,"
    StrSql = StrSql & " gpanro, gpadesabr"
    
    StrSql = StrSql & " ) VALUES( "
    
    StrSql = StrSql & " " & NroProceso
    StrSql = StrSql & "," & Ternro
    StrSql = StrSql & "," & pgtinro
    
    StrSql = StrSql & ",'" & Mid(apellido, 1, 50) & "'"
    StrSql = StrSql & ",'" & Mid(Nombre, 1, 50) & "'"
    StrSql = StrSql & "," & Legajo
    
    StrSql = StrSql & "," & tenro1
    StrSql = StrSql & "," & tenro2
    StrSql = StrSql & "," & tenro3
    
    StrSql = StrSql & "," & EmpEstrnro1
    StrSql = StrSql & "," & EmpEstrnro2
    StrSql = StrSql & "," & EmpEstrnro3
    
    StrSql = StrSql & ",'" & Mid(tituloReporte, 1, 100) & "'"
    StrSql = StrSql & ",'" & Mid(pgtidesabr, 1, 30) & "'"
    StrSql = StrSql & "," & ordenBase
    
    'Columnas variables
    For I = 1 To maxCol
        
        If I <= maxColUsed Then
            StrSql = StrSql & ",'" & Mid(arrEtiqConfrep(I), 1, 50) & "'"
            StrSql = StrSql & "," & arrValor(I)
        Else
            StrSql = StrSql & ", null "
            StrSql = StrSql & ", null "
        End If
        
    Next
    
    StrSql = StrSql & ",'" & gpanro & "'"
    StrSql = StrSql & ",'" & Mid(gpadesabr, 1, 100) & "'"
    StrSql = StrSql & ")"
    
    objConn.Execute StrSql, , adExecuteNoRecords
    
End Sub


'--------------------------------------------------------------------
' Se encarga de generar un ResultSet de los empleados a cambiar
' si el RS es vacio significa que hay que aplicarlo sobre todos
'--------------------------------------------------------------------
Sub CargarEmpleados(nroproc, ByRef rsEmpl As ADODB.Recordset)

Dim StrEmpl As String

    StrEmpl = "SELECT * FROM batch_empleado WHERE bpronro = " & nroproc
    StrEmpl = StrEmpl & " ORDER BY estado "
    
    OpenRecordset StrEmpl, rsEmpl
End Sub


Function MenorIgualMesAnio(ByVal mes1 As Long, ByVal anio1 As Long, ByVal mes2 As Long, ByVal anio2 As Long) As Boolean
Dim Aux As Boolean
    
    Aux = False
    
    If anio1 <= anio2 Then
        If anio1 < anio2 Then
            Aux = True
        Else
            'los anios son iguales
            If mes1 <= mes2 Then
                Aux = True
            End If
        End If
    End If
    
    MenorIgualMesAnio = Aux
    
End Function 'MenorIgualMesAnio(mes1,anio1,mes2,anio2)


'--------------------------------------------------------------------
' Arma una lista con los periodos comprendidos entre dos periodos
'--------------------------------------------------------------------
Public Function listaPeriodo(ByVal per_desde As Long, ByVal per_hasta As Long) As String
Dim listaP As String
Dim mdesde As Long
Dim adesde As Long
Dim mhasta As Long
Dim ahasta As Long

    listaP = ""
    mdesde = 0
    adesde = 0
    mhasta = 0
    ahasta = 0
    
    'Busco Periodo desde
    StrSql = "SELECT pgtinro, pgtihasta, pgtimes, pgtianio "
    StrSql = StrSql & "FROM gti_per "
    StrSql = StrSql & "WHERE pgtinro = " & per_desde
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        mdesde = objRs!pgtimes
        adesde = objRs!pgtianio
    End If
    objRs.Close
    
    'Busco Periodo hasta
    StrSql = "SELECT pgtinro, pgtihasta, pgtimes, pgtianio "
    StrSql = StrSql & "FROM gti_per "
    StrSql = StrSql & "WHERE pgtinro = " & per_hasta
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        mhasta = objRs!pgtimes
        ahasta = objRs!pgtianio
    End If
    objRs.Close
    
   
    StrSql = " SELECT * FROM gti_per "
    OpenRecordset StrSql, objRs

    Do Until objRs.EOF
        If ((MenorIgualMesAnio(mdesde, adesde, objRs!pgtimes, objRs!pgtianio)) And (MenorIgualMesAnio(objRs!pgtimes, objRs!pgtianio, mhasta, ahasta))) Then
        
            If Len(listaP) = 0 Then
               listaP = objRs!pgtinro
            Else
               listaP = listaP & "," & objRs!pgtinro
            End If
            
        End If
        
        objRs.MoveNext
    
    Loop
    objRs.Close

    listaPeriodo = listaP

End Function 'listaPeriodo


