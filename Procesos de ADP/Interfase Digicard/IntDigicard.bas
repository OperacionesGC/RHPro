Attribute VB_Name = "IntDigicard"
Option Explicit

'Global Const Version = "1.00"
'Global Const FechaModificacion = "13/07/2009 "
'Global Const UltimaModificacion = " " 'FGZ

'Global Const Version = "1.01"
'Global Const FechaModificacion = "21/07/2009 "
'Global Const UltimaModificacion = " " 'FGZ
''                                       Se Agregó la opcion para migrar empleados


Global Const Version = "1.02"
Global Const FechaModificacion = "07/08/2009 "
Global Const UltimaModificacion = " " 'FGZ
'       Se corrigió la llamada al sp de licencias pasaba el legajo y estaba esperando el tercero
'       Ademas para la actualizacion del empleado estaba llamando a varios SP y ahora se unieron todos en uno solo

'=======================================================================
'=======================================================================

Global NroProceso As Long
Global Path As String
Global HuboErrores As Boolean

Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset

Global EmpErrores As Boolean

Global Tenro1 As Long
Global Estrnro1 As Long
Global Tenro2 As Long
Global Estrnro2 As Long
Global Tenro3 As Long
Global Estrnro3 As Long
Global Todos As Boolean
Global TodasLic As Boolean
Global FechaDesde As Date
Global FechaHasta As Date
Global Empleados As String
Global Filtro As String
Global TipoInterfase As Long
Global Const Nulo = "NULL"
Global Const Vacio = ""

Dim IdUser As String
Dim bpfecha As Date
Dim bphora As String


Private Sub Main()
Dim strCmdLine As String
Dim Nombre_Arch As String

Dim StrSql As String

Dim PID As String
Dim Parametros As String
Dim ArrParametros
Dim totalEmpleados
Dim cantRegistros

Dim Desde As Date
Dim Hasta As Date
Dim fecestrAnt As Date

Dim rs_Emp As New ADODB.Recordset
Dim rs_Lic As New ADODB.Recordset
Dim rs_Aux As New ADODB.Recordset

Dim FAlta As String
Dim FBaja As String
Dim Linea As Long
Dim Sector As Long
Dim Tdoc As Long
Dim Doc As String
Dim Legajo As String
Dim Apellido As String
Dim Nombre As String
Dim Mail As String


Dim Partes As Integer

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
    
    Nombre_Arch = PathFLog & "InterfaseDigicard" & "-" & NroProceso & ".log"
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)


    'Abro la conexion
    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If
    OpenConnection strconexion, objconnProgreso
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If
    On Error GoTo 0
    
    On Error GoTo ME_Main
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Version = " & Version
    Flog.writeline "Modificacion = " & UltimaModificacion
    Flog.writeline "Fecha = " & FechaModificacion
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline
    Flog.writeline "PID = " & PID
    Flog.writeline
    
    Flog.writeline "Inicio Proceso de Interface con Digicard : " & Now
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & ", bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Flog.writeline Espacios(Tabulador * 0) & "Obtengo los datos del proceso"
    
    TiempoAcumulado = GetTickCount
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
       'Obtengo los parametros del proceso
        IdUser = rs!IdUser
        bpfecha = rs!bprcfecha
        bphora = rs!bprchora
        Parametros = rs!bprcparam
        
        ArrParametros = Split(Parametros, "@")
        Call levantarParametros(ArrParametros)
                       
        If TipoInterfase = 3 Then
            Partes = 2
        Else
            Partes = 1
        End If
               
        'FGZ - 21/07/2009 - Se agregó la migracion de empleados
        If TipoInterfase = 2 Or TipoInterfase = 3 Then
            Flog.writeline
            Flog.writeline Espacios(Tabulador * 0) & "Actualizo Empleados..."
            
            StrSql = "SELECT empleado.ternro, empleg, terape, ternom, empemail "
            StrSql = StrSql & " FROM empleado "
            StrSql = StrSql & " INNER JOIN batch_empleado ON empleado.ternro = batch_empleado.ternro AND batch_empleado.bpronro =  " & NroProceso
            StrSql = StrSql & " ORDER BY empleado.empleg"
            OpenRecordset StrSql, rs_Emp
            
            Progreso = 0
            cantRegistros = rs_Emp.RecordCount
            If cantRegistros = 0 Then
                cantRegistros = 1
                Flog.writeline Espacios(Tabulador * 1) & "No se encontraron Empleados para el Filtro."
            End If
            IncPorc = ((100 / Partes) / cantRegistros)
            Do While Not rs_Emp.EOF
                 
'                 'Empleado
'                 'sp_emp_to_digicard( 2, :OLD.ternro, leg, ape, nom, mail);
'                 StrSql = "sp_emp_to_digicard(2," & rs_Emp!ternro & "," & rs_Emp!empleg & "," & rs_Emp!terape & "," & rs_Emp!ternom & "," & rs_Emp!empemail & ")"
'                 On Error Resume Next
'                 Err.Number = 0
'                 objConn.Execute StrSql, , adExecuteNoRecords
'                 If Err.Number <> 0 Then
'                    Flog.writeline Espacios(Tabulador * 1) & "Error ejecutando store procedure "
'                    Flog.writeline Espacios(Tabulador * 1) & StrSql
'                 End If
'                 On Error GoTo ME_Main
'
'
'                 'Fases
'                 'Busco la ultima fase
'                 StrSql = "SELECT * FROM fases "
'                 StrSql = StrSql & " WHERE empleado = " & rs_Emp!ternro
'                 StrSql = StrSql & " ORDER BY altfec DESC"
'                 OpenRecordset StrSql, rs_Aux
'                 If Not rs_Aux.EOF Then
'                    'sp_fase_to_digicard( 2, :NEW.empleado, :NEW.altfec, :NEW.bajfec);
'                    If IsNull(rs_Aux!bajfec) Then
'                        StrSql = "sp_fase_to_digicard(2," & rs_Emp!ternro & "," & ConvFecha(rs_Aux!altfec) & "," & Nulo & ")"
'                    Else
'                        StrSql = "sp_fase_to_digicard(2," & rs_Emp!ternro & "," & ConvFecha(rs_Aux!altfec) & "," & ConvFecha(rs_Aux!bajfec) & ")"
'                    End If
'                    On Error Resume Next
'                    Err.Number = 0
'                    objConn.Execute StrSql, , adExecuteNoRecords
'                    If Err.Number <> 0 Then
'                    Flog.writeline Espacios(Tabulador * 1) & "Error ejecutando store procedure "
'                    Flog.writeline Espacios(Tabulador * 1) & StrSql
'                    End If
'                    On Error GoTo ME_Main
'                 End If
'
'                 'Documento
'                 'Busco el documento
'                 StrSql = "SELECT * FROM ter_doc "
'                 StrSql = StrSql & " WHERE ternro = " & rs_Emp!ternro
'                 StrSql = StrSql & " AND tidnro <= 5 "
'                 StrSql = StrSql & " ORDER BY tidnro "
'                 OpenRecordset StrSql, rs_Aux
'                 If Not rs_Aux.EOF Then
'                    'sp_doc_to_digicard( 2, :NEW.ternro, :NEW.tidnro, :NEW.nrodoc);
'                    StrSql = "sp_doc_to_digicard(2," & rs_Emp!ternro & "," & rs_Aux!tidnro & "," & rs_Aux!nrodoc & ")"
'                    On Error Resume Next
'                    Err.Number = 0
'                    objConn.Execute StrSql, , adExecuteNoRecords
'                    If Err.Number <> 0 Then
'                    Flog.writeline Espacios(Tabulador * 1) & "Error ejecutando store procedure "
'                    Flog.writeline Espacios(Tabulador * 1) & StrSql
'                    End If
'                    On Error GoTo ME_Main
'                 End If
'
'                'Estructuras
'                StrSql = " SELECT estrnro, tenro FROM his_estructura "
'                StrSql = StrSql & " WHERE ternro = " & rs_Emp!ternro
'                StrSql = StrSql & " AND (tenro = 2 OR tenro = 12)"
'                StrSql = StrSql & " AND (htetdesde <= " & ConvFecha(Date) & ")"
'                StrSql = StrSql & " AND ((" & ConvFecha(Date) & " <= htethasta) or (htethasta is null))"
'                OpenRecordset StrSql, rs_Aux
'                Do While Not rs_Aux.EOF
'                    'sp_est_to_digicard( 2, :OLD.ternro, est, tipo_est);
'                    StrSql = "sp_est_to_digicard(2," & rs_Emp!ternro & "," & rs_Aux!estrnro & "," & rs_Aux!tenro & ")"
'                    On Error Resume Next
'                    Err.Number = 0
'                    objConn.Execute StrSql, , adExecuteNoRecords
'                    If Err.Number <> 0 Then
'                    Flog.writeline Espacios(Tabulador * 1) & "Error ejecutando store procedure "
'                    Flog.writeline Espacios(Tabulador * 1) & StrSql
'                    End If
'                    On Error GoTo ME_Main
'
'                    rs_Aux.MoveNext
'                Loop
                 
                 
                '===========================================================================================
                'Inicializo los datos
                FAlta = Nulo
                FBaja = Nulo
                Linea = 0
                Sector = 0
                Tdoc = 0
                Doc = ""
                Legajo = ""
                Apellido = ""
                Nombre = ""
                Mail = ""
                
                
                 'Empleado
                 Apellido = rs_Emp!terape
                 Nombre = rs_Emp!ternom
                 Legajo = rs_Emp!empleg
                 Mail = rs_Emp!empemail
                 
                 
                 'Fases
                 StrSql = "SELECT * FROM fases "
                 StrSql = StrSql & " WHERE empleado = " & rs_Emp!ternro
                 StrSql = StrSql & " ORDER BY altfec DESC"
                 OpenRecordset StrSql, rs_Aux
                 If Not rs_Aux.EOF Then
                    FAlta = ConvFecha2(rs_Aux!altfec)
                    
                    If IsNull(rs_Aux!bajfec) Then
                        FBaja = Nulo
                    Else
                        FBaja = ConvFecha2(rs_Aux!bajfec)
                    End If
                 End If
                 
                 
                 'Documento
                 StrSql = "SELECT * FROM ter_doc "
                 StrSql = StrSql & " WHERE ternro = " & rs_Emp!ternro
                 StrSql = StrSql & " AND tidnro <= 5 "
                 StrSql = StrSql & " ORDER BY tidnro "
                 OpenRecordset StrSql, rs_Aux
                 If Not rs_Aux.EOF Then
                    Tdoc = rs_Aux!tidnro
                    Doc = rs_Aux!nrodoc
                 End If
                 
                 
                'Estructuras
                'Sector
                StrSql = " SELECT estrnro, tenro FROM his_estructura "
                StrSql = StrSql & " WHERE ternro = " & rs_Emp!ternro
                StrSql = StrSql & " AND (tenro = 2)"
                StrSql = StrSql & " AND (htetdesde <= " & ConvFecha(Date) & ")"
                StrSql = StrSql & " AND ((" & ConvFecha(Date) & " <= htethasta) or (htethasta is null))"
                OpenRecordset StrSql, rs_Aux
                If Not rs_Aux.EOF Then
                    Sector = rs_Aux!estrnro
                End If
                'Linea
                StrSql = " SELECT estrnro, tenro FROM his_estructura "
                StrSql = StrSql & " WHERE ternro = " & rs_Emp!ternro
                StrSql = StrSql & " AND (tenro = 12)"
                StrSql = StrSql & " AND (htetdesde <= " & ConvFecha(Date) & ")"
                StrSql = StrSql & " AND ((" & ConvFecha(Date) & " <= htethasta) or (htethasta is null))"
                OpenRecordset StrSql, rs_Aux
                If Not rs_Aux.EOF Then
                    Linea = rs_Aux!estrnro
                End If
                 
                  'sp_empleado_to_digicard( 2, :OLD.ternro, leg, ape, nom, mail,tdoc, nrodoc, alta, baja, mail, sector, linea );
                 StrSql = "EXECUTE sp_empleado_to_digicard(2," & rs_Emp!ternro & ",'" & Legajo & "','" & Apellido & "','" & Nombre & "','" & Mail
                 StrSql = StrSql & "'," & Tdoc & ",'" & Doc & "'," & FAlta & "," & FBaja & "," & Sector & "," & Linea & ")"
                 On Error Resume Next
                 Err.Number = 0
                 objConn.Execute StrSql, , adExecuteNoRecords
                 If Err.Number <> 0 Then
                    Flog.writeline Espacios(Tabulador * 1) & "Error ejecutando store procedure "
                    Flog.writeline Espacios(Tabulador * 1) & StrSql
                 End If
                 On Error GoTo ME_Main
                
                '===========================================================================================
                 
                'Actualizo el progreso
                TiempoAcumulado = GetTickCount
                Progreso = Progreso + IncPorc
                StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
                StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'"
                StrSql = StrSql & " WHERE bpronro = " & NroProceso
                objConn.Execute StrSql, , adExecuteNoRecords
                 
                rs_Emp.MoveNext
            Loop
        End If
        
        'FGZ - 21/07/2009
        If TipoInterfase = 1 Or TipoInterfase = 3 Then
            Flog.writeline
            Flog.writeline Espacios(Tabulador * 0) & "Actualizo Licencias..."
             
             'Buscar las licencias
             StrSql = "SELECT empleg, empleado.ternro, emp_licnro, tdnro, licestnro, elfechadesde, elfechahasta "
             StrSql = StrSql & " FROM empleado "
             StrSql = StrSql & " INNER JOIN batch_empleado ON empleado.ternro = batch_empleado.ternro AND batch_empleado.bpronro =  " & NroProceso
             StrSql = StrSql & " INNER JOIN emp_lic ON empleado.ternro = emp_lic.empleado "
             StrSql = StrSql & " WHERE ((elfechadesde <= " & ConvFecha(FechaDesde) & " AND elfechahasta >= " & ConvFecha(FechaHasta) & ")"
             StrSql = StrSql & " OR (elfechadesde <= " & ConvFecha(FechaHasta) & " AND elfechahasta >= " & ConvFecha(FechaHasta) & ")"
             StrSql = StrSql & " OR (elfechadesde <= " & ConvFecha(FechaDesde) & " AND elfechahasta >= " & ConvFecha(FechaDesde) & ")"
             StrSql = StrSql & " OR (elfechadesde >= " & ConvFecha(FechaDesde) & " AND elfechadesde <= " & ConvFecha(FechaHasta) & " AND elfechahasta <= " & ConvFecha(FechaHasta) & "))"
             StrSql = StrSql & " ORDER BY empleado.empleg, emp_lic.elfechadesde"
             OpenRecordset StrSql, rs_Lic
            
             'seteo de las variables de progreso
             If TipoInterfase = 3 Then
                Progreso = 50
             Else
                Progreso = 0
            End If
             cantRegistros = rs_Lic.RecordCount
             If cantRegistros = 0 Then
                cantRegistros = 1
                Flog.writeline Espacios(Tabulador * 1) & "No se encontraron Licencias para el Filtro."
             End If
             IncPorc = ((100 / Partes) / cantRegistros)
               
             ' Se inicia el proceso
             Do While Not rs_Lic.EOF
                 If rs_Lic!licestnro = 2 Then    'Aprobada
                     'Inserto la Liecncia
                     'StrSql = "sp_lic_to_digicard( 2," & rs_Lic!empleg & "," & ConvFecha(rs_Lic!elfechadesde) & "," & ConvFecha(rs_Lic!elfechahasta) & "," & rs_Lic!tdnro & "," & ConvFecha(rs_Lic!elfechadesde) & "," & ConvFecha(rs_Lic!elfechahasta) & "," & rs_Lic!tdnro & "," & rs_Lic!emp_licnro & ")"
                     StrSql = "EXECUTE sp_lic_to_digicard( 2," & rs_Lic!ternro & "," & ConvFecha2(rs_Lic!elfechadesde) & "," & ConvFecha2(rs_Lic!elfechahasta) & "," & rs_Lic!tdnro & "," & ConvFecha2(rs_Lic!elfechadesde) & "," & ConvFecha2(rs_Lic!elfechahasta) & "," & rs_Lic!tdnro & "," & rs_Lic!emp_licnro & ")"
                 Else
                     'Elimino la Liecncia
                     'StrSql = "sp_lic_to_digicard( 3," & rs_Lic!empleg & "," & ConvFecha(rs_Lic!elfechadesde) & "," & ConvFecha(rs_Lic!elfechahasta) & "," & rs_Lic!tdnro & "," & ConvFecha(rs_Lic!elfechadesde) & "," & ConvFecha(rs_Lic!elfechahasta) & "," & rs_Lic!tdnro & "," & rs_Lic!emp_licnro & ")"
                     StrSql = "EXECUTE sp_lic_to_digicard( 3," & rs_Lic!ternro & "," & ConvFecha2(rs_Lic!elfechadesde) & "," & ConvFecha2(rs_Lic!elfechahasta) & "," & rs_Lic!tdnro & "," & ConvFecha2(rs_Lic!elfechadesde) & "," & ConvFecha2(rs_Lic!elfechahasta) & "," & rs_Lic!tdnro & "," & rs_Lic!emp_licnro & ")"
                 End If
                 'StrSql = "sp_xxx 3"
                On Error Resume Next
                Err.Number = 0
                objConn.Execute StrSql, , adExecuteNoRecords
                If Err.Number <> 0 Then
                    Flog.writeline Espacios(Tabulador * 1) & "Error ejecutando store procedure "
                    Flog.writeline Espacios(Tabulador * 1) & StrSql
                End If
                On Error GoTo ME_Main

                     
                 'Actualizo el progreso
                 TiempoAcumulado = GetTickCount
                 Progreso = Progreso + IncPorc
                 StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
                 StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'"
                 StrSql = StrSql & " WHERE bpronro = " & NroProceso
                 objConn.Execute StrSql, , adExecuteNoRecords
                 
                 rs_Lic.MoveNext
             Loop
        End If
    Else
        Flog.writeline Espacios(Tabulador * 0) & "No se encontraron los datos del proceso nro " & NroProceso
    End If
    
          
    'Actualizo el estado del proceso
    If Not HuboErrores Then
        StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
        Flog.writeline Espacios(Tabulador * 0) & "Proceso Finalizado Correctamente"
    Else
        StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProceso
        Flog.writeline Espacios(Tabulador * 0) & "Proceso Finalizado Incompleto"
    End If
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    TiempoFinalProceso = GetTickCount
    Flog.writeline Espacios(Tabulador * 0) & "=================================================="
    Flog.writeline Espacios(Tabulador * 0) & "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    Flog.writeline Espacios(Tabulador * 0) & "=================================================="
    Flog.Close
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    objconnProgreso.Close
    objConn.Close
    
Exit Sub
    
ME_Main:
    HuboErrores = True
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQL: " & StrSql
End Sub




 
Sub levantarParametros(ArrParametros)
'--------------------------------------------------------------
' procedimiento que levanta los parametros
'--------------------------------------------------------------
On Error GoTo ME_param
    'FGZ - 21/07/2009 - Se agregó este parametro
    TipoInterfase = CLng(ArrParametros(0))
    
    Todos = CBool(ArrParametros(1))
    Tenro1 = CLng(ArrParametros(2))
    Estrnro1 = CLng(ArrParametros(3))
    Tenro2 = CLng(ArrParametros(4))
    Estrnro2 = CLng(ArrParametros(5))
    Tenro3 = CLng(ArrParametros(6))
    Estrnro3 = CLng(ArrParametros(7))
    
    FechaDesde = CDate(ArrParametros(8))
    FechaHasta = CDate(ArrParametros(9))
    
    TodasLic = CBool(ArrParametros(10))

Exit Sub

ME_param:
    Flog.writeline "    Error: Error en la carga de Parametros "
    
End Sub

Sub Filtro_Empleados(ByVal StrSql As String, ByVal Fecha As Date)
'---------------------------------------------------------------------------------------------------
' procedimiento que busca los empleados que cumplen con lo seleccionado en el filtro
'---------------------------------------------------------------------------------------------------
Dim StrAgencia As String
Dim StrSelect As String
Dim strjoin As String
Dim StrOrder As String
Dim fecdes As String
Dim fechas As String
Dim rsfiltro As New ADODB.Recordset

On Error GoTo ME_armarsql
StrSql = ""
StrSelect = ""
strjoin = ""
StrOrder = ""
 
If Tenro1 <> 0 Then  ' Cuando solo selecionamos el primer nivel
    strjoin = strjoin & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro  AND estact1.tenro  = " & Tenro1
    strjoin = strjoin & "  AND (estact1.htetdesde<=" & ConvFecha(Fecha) & " AND (estact1.htethasta IS NULL OR estact1.htethasta>=" & ConvFecha(Fecha) & "))"
    If Estrnro1 <> 0 Then
        strjoin = strjoin & " AND estact1.estrnro =" & Estrnro1
    End If
    strjoin = strjoin & " INNER JOIN estructura estructura1 ON estructura1.estrnro=estact1.estrnro "
    
    If StrOrder <> "" Then
        StrOrder = StrOrder & ", "
    End If
    StrOrder = StrOrder & " tenro1, estrnro1 "
End If

If Tenro2 <> 0 Then  ' ocurre cuando se selecciono hasta el segundo nivel
    strjoin = strjoin & " INNER JOIN his_estructura estact2 ON empleado.ternro = estact2.ternro  AND estact2.tenro  = " & Tenro2
    strjoin = strjoin & " AND (estact2.htetdesde<=" & ConvFecha(Fecha) & " AND (estact2.htethasta IS NULL OR estact2.htethasta>=" & ConvFecha(Fecha) & "))"
    If Estrnro2 <> 0 Then
        strjoin = strjoin & " AND estact2.estrnro =" & Estrnro2
    End If
    strjoin = strjoin & " INNER JOIN estructura estructura2 ON estructura2.estrnro=estact2.estrnro "
    
    If StrOrder <> "" Then
        StrOrder = StrOrder & ", "
    End If
    StrOrder = StrOrder & " tenro2, estrnro2 "
End If

If Tenro3 <> 0 Then  ' esto ocurre solo cuando se seleccionan los tres niveles
    strjoin = strjoin & " INNER JOIN his_estructura estact3 ON empleado.ternro = estact3.ternro  AND estact3.tenro =" & Tenro3
    strjoin = strjoin & "   AND (estact3.htetdesde<=" & ConvFecha(Fecha) & " AND (estact3.htethasta IS NULL OR  estact3.htethasta>=" & ConvFecha(Fecha) & "))"
    If Estrnro3 <> 0 Then 'cuando se le asigna un valor al nivel 3
        strjoin = strjoin & " AND estact3.estrnro =" & Estrnro3
    End If
    strjoin = strjoin & " INNER JOIN estructura estructura3 ON estructura3.estrnro=estact3.estrnro "
    If StrOrder <> "" Then
        StrOrder = StrOrder & ", "
    End If
    StrOrder = StrOrder & " tenro3, estrnro3 "
End If

StrSql = " SELECT DISTINCT empleado.ternro  "   '  empleado.empest, tercero.tersex,
StrSql = StrSql & " FROM empleado "
StrSql = StrSql & " INNER JOIN tercero ON empleado.ternro = tercero.ternro "
StrSql = StrSql & strjoin
StrSql = StrSql & " WHERE " & Filtro
OpenRecordset StrSql, rsfiltro
Empleados = "(0"
While Not rsfiltro.EOF
    Empleados = Empleados & "," & rsfiltro!ternro
    rsfiltro.MoveNext
Wend
Empleados = Empleados & ")"

Exit Sub

ME_armarsql:
    Flog.writeline " Error: Armar consulta del Filtro.- " & Err.Description
    Flog.writeline " Búsqueda de empleados filtrados: " & StrSql
End Sub


Private Sub Actualizar_Progreso(Progreso As Integer)

    TiempoAcumulado = GetTickCount

    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
    StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'"
    StrSql = StrSql & " WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
End Sub

Public Function ConvFecha2(ByVal dteFecha As Date) As String
    
    ConvFecha2 = "'" & Format(dteFecha, strformatoFservidor) & "'"
    
End Function

