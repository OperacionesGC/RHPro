Attribute VB_Name = "CambioCantEmpxSem"
Option Explicit

Global Const Version = "1.01"
Global Const FechaModificacion = "19/12/2006"
Global Const UltimaModificacion = " " 'FAF - Version Inicial

'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------

Global fs, f
Global Flog
Global NroProceso As Long

Global Path As String
Global HuboErrores As Boolean

Global IdUser As String
Global Fecha As Date
Global Hora As String

Private Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion : Procedimiento inicial
' Autor       : Fernando Favre
' Fecha       : 19/12/2006
' Ultima Mod  :
' Descripcion :
' ---------------------------------------------------------------------------------------------
Dim strCmdLine As String
Dim Nombre_Arch As String

Dim rs As New ADODB.Recordset
Dim PID As String
Dim Parametros As String
Dim ArrParametros

Dim empresa As Integer
Dim mes As Integer
Dim Anio As Integer
Dim todasSuc As Integer
Dim estrnroSuc As Long


    strCmdLine = Command()
    ArrParametros = Split(strCmdLine, " ", -1)
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

    MyBeginTrans
     
    On Error GoTo ME_Main
    
    TiempoInicialProceso = GetTickCount
    HuboErrores = False
    
    Nombre_Arch = PathFLog & "Cantidad_Empleados_x_Semana" & "-" & NroProceso & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Version = " & Version
    Flog.writeline "Modificacion = " & UltimaModificacion
    Flog.writeline "Fecha = " & FechaModificacion
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline
    Flog.writeline "PID = " & PID
   
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & ", bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline Espacios(Tabulador * 0) & "Obtengo los datos del proceso"
    
    TiempoAcumulado = GetTickCount
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProceso
    StrSql = StrSql & " AND btprcnro = 153"
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        IdUser = rs!IdUser
        Fecha = rs!bprcfecha
        Hora = rs!bprchora
        
        Parametros = rs!bprcparam
        ArrParametros = Split(Parametros, "@")
        
        empresa = CInt(ArrParametros(0))
        mes = CInt(ArrParametros(1))
        Anio = CInt(ArrParametros(2))
        todasSuc = CInt(ArrParametros(3))
        estrnroSuc = CLng(ArrParametros(4))
        
        Call Generar_Datos(empresa, mes, Anio, todasSuc, estrnroSuc)
    Else
        Flog.writeline Espacios(Tabulador * 0) & "No se encontraron los datos del proceso nro " & NroProceso & " de tipo 153."
    End If

Fin:
    'Actualizo el estado del proceso
    If Not HuboErrores Then
       StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
       Flog.writeline Espacios(Tabulador * 0) & "Proceso Finalizado Correctamente"
       MyCommitTrans
    Else
       StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProceso
       Flog.writeline Espacios(Tabulador * 0) & "Proceso Finalizado Incompleto"
       MyRollbackTrans
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
    Flog.writeline "Ultima SQL ejecutada: " & StrSql
    GoTo Fin
End Sub

Private Sub Generar_Datos(ByVal empresa As Integer, ByVal mes As Integer, ByVal Anio As Integer, ByVal todasSuc As Integer, ByVal estrnroSuc As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion : Procedimiento que genera los datos del reporte
' Autor       : Fernando Favre
' Fecha       : 19/12/2006
' Ultima Mod  :
' Descripcion :
' ---------------------------------------------------------------------------------------------
Dim cantRegistros As Long
Dim semanas_fec(5) As Date
Dim ternro_empl_ant As Long
Dim estrnro_suc As Long
Dim estrdabr_suc As String
Dim estrcodext_suc As String
Dim estrnro_pue As Long
Dim estrdabr_pue As String
Dim estrcodext_pue As String
Dim estrnro_reghor As Long
Dim estrdabr_reghor As String
Dim estrcodext_reghor As String
Dim estrdabrSuc As String
Dim EmpNombre As String
Dim EmpDire As String
Dim ternro_emp As Integer
Dim EmpCuit As String
Dim EmpLogo As String
Dim EmpLogoAlto
Dim EmpLogoAncho

Dim rs_empleados As New ADODB.Recordset
Dim rs_fases As New ADODB.Recordset
Dim rs_Domicilio As New ADODB.Recordset
Dim rs_cuit As New ADODB.Recordset
Dim rs_logo As New ADODB.Recordset
Dim rs_estructura As New ADODB.Recordset

    On Error GoTo ME_Local
      
    Call cargar_semanas(semanas_fec, mes, Anio)
    
    StrSql = "SELECT empleado.ternro, empleado.empleg "
    StrSql = StrSql & "FROM empleado "
    StrSql = StrSql & " INNER JOIN his_estructura ON empleado.ternro = his_estructura.ternro "
    StrSql = StrSql & " WHERE his_estructura.tenro = 1 AND his_estructura.htethasta IS NULL AND empleado.empest=-1 "
    If todasSuc <> -1 And Not EsNulo(estrnroSuc) Then
        StrSql = StrSql & " AND his_estructura.estrnro = " & estrnroSuc
    End If
    StrSql = StrSql & " ORDER BY empleado.ternro"
    OpenRecordset StrSql, rs_empleados
    
    'seteo de las variables de progreso
    Progreso = 0
    If Not rs_empleados.EOF Then
        cantRegistros = rs_empleados.RecordCount
        If cantRegistros = 0 Then
           cantRegistros = 1
           Flog.writeline Espacios(Tabulador * 1) & "No se encontraron empleados con la estructura Sucursal activa. SQL: " & StrSql
        End If
    Else
        cantRegistros = 1
        Flog.writeline Espacios(Tabulador * 1) & "No se encontraron empleados con la estructura Sucursal activa. SQL: " & StrSql
    End If
    IncPorc = (99 / cantRegistros)
    
    ternro_empl_ant = 0
    
    If Not rs_empleados.EOF Then
        estrdabrSuc = ""
        If todasSuc <> -1 Then
            StrSql = " SELECT estrdabr FROM estructura WHERE estrnro = " & estrnroSuc
            OpenRecordset StrSql, rs_estructura
            If Not rs_estructura.EOF Then
                estrdabrSuc = rs_estructura!estrdabr
            End If
            rs_estructura.Close
        End If
        
        'Consulta para obtener la direccion de la empresa
        StrSql = "SELECT empresa.ternro,empresa.empnom,detdom.calle,detdom.nro,detdom.piso,detdom.oficdepto,localidad.locdesc FROM empresa "
        StrSql = StrSql & " INNER JOIN cabdom ON cabdom.ternro = empresa.ternro"
        StrSql = StrSql & " INNER JOIN detdom ON detdom.domnro = cabdom.domnro"
        StrSql = StrSql & " INNER JOIN localidad ON detdom.locnro = localidad.locnro "
        StrSql = StrSql & " WHERE empresa.empnro = " & empresa
        OpenRecordset StrSql, rs_Domicilio
        If rs_Domicilio.EOF Then
            Flog.writeline "No se encontró el Domicilio de la Empresa."
            'Exit Sub
            EmpNombre = ""
            EmpDire = ""
            ternro_emp = 0
        Else
            ternro_emp = rs_Domicilio!ternro
            EmpNombre = rs_Domicilio!empnom
            EmpDire = rs_Domicilio!calle & " " & rs_Domicilio!Nro
            If Not EsNulo(rs_Domicilio!piso) Then
                EmpDire = EmpDire & " " & rs_Domicilio!piso
            End If
            If Not EsNulo(rs_Domicilio!oficdepto) Then
                EmpDire = EmpDire & " Dpto. " & rs_Domicilio!oficdepto
            End If
        End If
        
        'Consulta para obtener el cuit de la empresa
        StrSql = "SELECT cuit.nrodoc FROM tercero "
        StrSql = StrSql & " INNER JOIN ter_doc cuit ON (tercero.ternro = cuit.ternro and cuit.tidnro = 6)"
        StrSql = StrSql & " WHERE tercero.ternro = " & ternro_emp
        OpenRecordset StrSql, rs_cuit
        If rs_cuit.EOF Then
            Flog.writeline "No se encontró el CUIT de la Empresa."
            'Exit Sub
            EmpCuit = ""
        Else
            EmpCuit = rs_cuit!nrodoc
        End If
        
        'Consulta para buscar el logo de la empresa
        StrSql = "SELECT ter_imag.terimnombre, tipoimag.tipimdire, tipoimag.tipimanchodef, tipoimag.tipimaltodef"
        StrSql = StrSql & " FROM ter_imag "
        StrSql = StrSql & " INNER JOIN tipoimag ON tipoimag.tipimnro = 1 AND tipoimag.tipimnro = ter_imag.tipimnro"
        StrSql = StrSql & " AND ter_imag.ternro =" & ternro_emp
        OpenRecordset StrSql, rs_logo
        If rs_logo.EOF Then
            Flog.writeline "No se encontró el Logo de la Empresa."
            'Exit Sub
            EmpLogo = ""
            EmpLogoAlto = 0
            EmpLogoAncho = 0
        Else
            EmpLogo = rs_logo!tipimdire & rs_logo!terimnombre
            EmpLogoAlto = rs_logo!tipimaltodef
            EmpLogoAncho = rs_logo!tipimanchodef
        End If
        
        
        StrSql = " INSERT INTO rep_empl_sem (bpronro,mes,anio,todas_suc,suc_nro,suc_dabr,empnombre,empdire,"
        StrSql = StrSql & "empcuit,emplogo,emplogoalto,emplogoancho,fecha,hora,iduser) VALUES ("
        StrSql = StrSql & NroProceso & ","
        StrSql = StrSql & mes & ","
        StrSql = StrSql & Anio & ","
        StrSql = StrSql & CInt(todasSuc) & ","
        StrSql = StrSql & estrnroSuc & ","
        StrSql = StrSql & "'" & estrdabrSuc & "',"
        StrSql = StrSql & "'" & EmpNombre & "',"
        StrSql = StrSql & "'" & EmpDire & "',"
        StrSql = StrSql & "'" & EmpCuit & "',"
        StrSql = StrSql & "'" & EmpLogo & "',"
        StrSql = StrSql & EmpLogoAlto & ","
        StrSql = StrSql & EmpLogoAncho & ","
        StrSql = StrSql & ConvFecha(Fecha) & ","
        StrSql = StrSql & "'" & Hora & "',"
        StrSql = StrSql & "'" & IdUser & "')"
        objConn.Execute StrSql, , adExecuteNoRecords
        
    End If
    
    Do While Not rs_empleados.EOF
        If rs_empleados!ternro <> ternro_empl_ant Then
            ' Primer semana
            StrSql = "SELECT * FROM fases WHERE fases.empleado = " & rs_empleados!ternro
            StrSql = StrSql & " AND altfec <= " & ConvFecha(semanas_fec(0)) & " AND (bajfec IS NULL OR "
            StrSql = StrSql & " bajfec >= " & ConvFecha(DateAdd("d", 7, semanas_fec(0))) & ")"
            OpenRecordset StrSql, rs_fases
            If Not rs_fases.EOF Then
                ' Busco la sucursal
                Call buscar_estr(CLng(rs_empleados!ternro), 1, CLng(estrnroSuc), semanas_fec(0), estrnro_suc, estrdabr_suc, estrcodext_suc)
            
                ' Busco el puesto
                Call buscar_estr(CLng(rs_empleados!ternro), 4, CLng(0), semanas_fec(0), estrnro_pue, estrdabr_pue, estrcodext_pue)
            
                ' Busco el Regimen Horario
                Call buscar_estr(CLng(rs_empleados!ternro), 21, CLng(0), semanas_fec(0), estrnro_reghor, estrdabr_reghor, estrcodext_reghor)
                
                If estrnro_suc <> 0 And estrnro_pue <> 0 And estrnro_reghor <> 0 Then
                    Call guardar_valores(1, rs_empleados!empleg, estrnro_suc, estrdabr_suc, estrcodext_suc, estrnro_pue, estrdabr_pue, estrnro_reghor, estrdabr_reghor)
                End If
                
            End If
            rs_fases.Close
        
            ' Segunda semana
            StrSql = "SELECT * FROM fases WHERE fases.empleado = " & rs_empleados!ternro
            StrSql = StrSql & " AND altfec <= " & ConvFecha(semanas_fec(1)) & " AND (bajfec IS NULL OR "
            StrSql = StrSql & " bajfec >= " & ConvFecha(DateAdd("d", 7, semanas_fec(1))) & ")"
            OpenRecordset StrSql, rs_fases
            If Not rs_fases.EOF Then
                ' Busco la sucursal
                Call buscar_estr(CLng(rs_empleados!ternro), 1, CLng(estrnroSuc), semanas_fec(1), estrnro_suc, estrdabr_suc, estrcodext_suc)
            
                ' Busco el puesto
                Call buscar_estr(CLng(rs_empleados!ternro), 4, CLng(0), semanas_fec(1), estrnro_pue, estrdabr_pue, estrcodext_pue)
            
                ' Busco el Regimen Horario
                Call buscar_estr(CLng(rs_empleados!ternro), 21, CLng(0), semanas_fec(1), estrnro_reghor, estrdabr_reghor, estrcodext_reghor)
                
                If estrnro_suc <> 0 And estrnro_pue <> 0 And estrnro_reghor <> 0 Then
                    Call guardar_valores(2, rs_empleados!empleg, estrnro_suc, estrdabr_suc, estrcodext_suc, estrnro_pue, estrdabr_pue, estrnro_reghor, estrdabr_reghor)
                End If
                
            End If
            rs_fases.Close
        
            ' Tercer semana
            StrSql = "SELECT * FROM fases WHERE fases.empleado = " & rs_empleados!ternro
            StrSql = StrSql & " AND altfec <= " & ConvFecha(semanas_fec(2)) & " AND (bajfec IS NULL OR "
            StrSql = StrSql & " bajfec >= " & ConvFecha(DateAdd("d", 7, semanas_fec(2))) & ")"
            OpenRecordset StrSql, rs_fases
            If Not rs_fases.EOF Then
                ' Busco la sucursal
                Call buscar_estr(CLng(rs_empleados!ternro), 1, CLng(estrnroSuc), semanas_fec(2), estrnro_suc, estrdabr_suc, estrcodext_suc)
            
                ' Busco el puesto
                Call buscar_estr(CLng(rs_empleados!ternro), 4, CLng(0), semanas_fec(2), estrnro_pue, estrdabr_pue, estrcodext_pue)
            
                ' Busco el Regimen Horario
                Call buscar_estr(CLng(rs_empleados!ternro), 21, CLng(0), semanas_fec(2), estrnro_reghor, estrdabr_reghor, estrcodext_reghor)
                
                If estrnro_suc <> 0 And estrnro_pue <> 0 And estrnro_reghor <> 0 Then
                    Call guardar_valores(3, rs_empleados!empleg, estrnro_suc, estrdabr_suc, estrcodext_suc, estrnro_pue, estrdabr_pue, estrnro_reghor, estrdabr_reghor)
                End If
                
            End If
            rs_fases.Close
        
            ' Cuarta semana
            StrSql = "SELECT * FROM fases WHERE fases.empleado = " & rs_empleados!ternro
            StrSql = StrSql & " AND altfec <= " & ConvFecha(semanas_fec(3)) & " AND (bajfec IS NULL OR "
            StrSql = StrSql & " bajfec >= " & ConvFecha(DateAdd("d", 7, semanas_fec(3))) & ")"
            OpenRecordset StrSql, rs_fases
            If Not rs_fases.EOF Then
                ' Busco la sucursal
                Call buscar_estr(CLng(rs_empleados!ternro), 1, CLng(estrnroSuc), semanas_fec(3), estrnro_suc, estrdabr_suc, estrcodext_suc)
            
                ' Busco el puesto
                Call buscar_estr(CLng(rs_empleados!ternro), 4, CLng(0), semanas_fec(3), estrnro_pue, estrdabr_pue, estrcodext_pue)
            
                ' Busco el Regimen Horario
                Call buscar_estr(CLng(rs_empleados!ternro), 21, CLng(0), semanas_fec(3), estrnro_reghor, estrdabr_reghor, estrcodext_reghor)
                
                If estrnro_suc <> 0 And estrnro_pue <> 0 And estrnro_reghor <> 0 Then
                    Call guardar_valores(4, rs_empleados!empleg, estrnro_suc, estrdabr_suc, estrcodext_suc, estrnro_pue, estrdabr_pue, estrnro_reghor, estrdabr_reghor)
                End If
                
            End If
            rs_fases.Close
        
            ' Quinta semana
            StrSql = "SELECT * FROM fases WHERE fases.empleado = " & rs_empleados!ternro
            StrSql = StrSql & " AND altfec <= " & ConvFecha(semanas_fec(4)) & " AND (bajfec IS NULL OR "
            StrSql = StrSql & " bajfec >= " & ConvFecha(DateAdd("d", 7, semanas_fec(4))) & ")"
            OpenRecordset StrSql, rs_fases
            If Not rs_fases.EOF Then
                ' Busco la sucursal
                Call buscar_estr(CLng(rs_empleados!ternro), 1, CLng(estrnroSuc), semanas_fec(4), estrnro_suc, estrdabr_suc, estrcodext_suc)
            
                ' Busco el puesto
                Call buscar_estr(CLng(rs_empleados!ternro), 4, CLng(0), semanas_fec(4), estrnro_pue, estrdabr_pue, estrcodext_pue)
            
                ' Busco el Regimen Horario
                Call buscar_estr(CLng(rs_empleados!ternro), 21, CLng(0), semanas_fec(4), estrnro_reghor, estrdabr_reghor, estrcodext_reghor)
                
                If estrnro_suc <> 0 And estrnro_pue <> 0 And estrnro_reghor <> 0 Then
                    Call guardar_valores(5, rs_empleados!empleg, estrnro_suc, estrdabr_suc, estrcodext_suc, estrnro_pue, estrdabr_pue, estrnro_reghor, estrdabr_reghor)
                End If
                
            End If
            rs_fases.Close
            
        End If
        rs_empleados.MoveNext
    
        'Actualizo el estado del proceso
        TiempoAcumulado = GetTickCount
        cantRegistros = cantRegistros - 1
        Progreso = Progreso + IncPorc
        
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
        StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'"
        StrSql = StrSql & ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
        
    Loop
    rs_empleados.Close
    
Fin:
    'Cierro y libero todo
    If rs_empleados.State = adStateOpen Then rs_empleados.Close
    Set rs_empleados = Nothing
    If rs_fases.State = adStateOpen Then rs_fases.Close
    Set rs_fases = Nothing
    If rs_Domicilio.State = adStateOpen Then rs_Domicilio.Close
    Set rs_Domicilio = Nothing
    If rs_cuit.State = adStateOpen Then rs_cuit.Close
    Set rs_cuit = Nothing
    If rs_logo.State = adStateOpen Then rs_logo.Close
    Set rs_logo = Nothing
    If rs_estructura.State = adStateOpen Then rs_estructura.Close
    Set rs_estructura = Nothing

Exit Sub

ME_Local:
    Flog.writeline
'    Resume Next
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Description
    Flog.writeline Espacios(Tabulador * 1) & "Ultimo SQL: " & StrSql
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.writeline
    HuboErrores = True
End Sub

Sub guardar_valores(ByVal semana As Integer, ByVal legajo As String, ByVal estrnro_suc As Long, ByVal estrdabr_suc As String, ByVal estrcodext_suc As String, ByVal estrnro_pue As Long, ByVal estrdabr_pue As String, ByVal estrnro_reghor As Long, ByVal estrdabr_reghor As String)
 Dim rs As New ADODB.Recordset
 Dim coeficiente As Double
 Dim area As String
 Dim cantidad1 As Integer
 Dim cantidad2 As Integer
 Dim cantidad3 As Integer
 Dim cantidad4 As Integer
 Dim cantidad5 As Integer
 Dim total1 As Double
 Dim total2 As Double
 Dim total3 As Double
 Dim total4 As Double
 Dim total5 As Double
 Dim legajos1 As String
 Dim legajos2 As String
 Dim legajos3 As String
 Dim legajos4 As String
 Dim legajos5 As String
 
    StrSql = "SELECT * FROM rep_empl_sem_det WHERE estrnro_suc = " & estrnro_suc
    StrSql = StrSql & " AND estrnro_pue = " & estrnro_pue
    StrSql = StrSql & " AND estrnro_regh = " & estrnro_reghor
    StrSql = StrSql & " AND bpronro = " & NroProceso
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        'update
        cantidad1 = IIf(semana = 1, rs!cantidad1 + 1, rs!cantidad1)
        cantidad2 = IIf(semana = 2, rs!cantidad2 + 1, rs!cantidad2)
        cantidad3 = IIf(semana = 3, rs!cantidad3 + 1, rs!cantidad3)
        cantidad4 = IIf(semana = 4, rs!cantidad4 + 1, rs!cantidad4)
        cantidad5 = IIf(semana = 5, rs!cantidad5 + 1, rs!cantidad5)
        
        total1 = IIf(semana = 1, cantidad1 * rs!coeficiente, rs!total1)
        total2 = IIf(semana = 2, cantidad2 * rs!coeficiente, rs!total2)
        total3 = IIf(semana = 3, cantidad3 * rs!coeficiente, rs!total3)
        total4 = IIf(semana = 4, cantidad4 * rs!coeficiente, rs!total4)
        total5 = IIf(semana = 5, cantidad5 * rs!coeficiente, rs!total5)
        
        legajos1 = IIf(semana = 1, IIf(rs!legajos1 <> "", IIf(Len(rs!legajos1 & " " & legajo) > 1000, Mid(rs!legajos1 & " " & legajo, 1, 1000), rs!legajos1 & " " & legajo), legajo), rs!legajos1)
        legajos2 = IIf(semana = 2, IIf(rs!legajos2 <> "", IIf(Len(rs!legajos2 & " " & legajo) > 1000, Mid(rs!legajos2 & " " & legajo, 1, 1000), rs!legajos2 & " " & legajo), legajo), rs!legajos2)
        legajos3 = IIf(semana = 3, IIf(rs!legajos3 <> "", IIf(Len(rs!legajos3 & " " & legajo) > 1000, Mid(rs!legajos3 & " " & legajo, 1, 1000), rs!legajos3 & " " & legajo), legajo), rs!legajos3)
        legajos4 = IIf(semana = 4, IIf(rs!legajos4 <> "", IIf(Len(rs!legajos4 & " " & legajo) > 1000, Mid(rs!legajos4 & " " & legajo, 1, 1000), rs!legajos4 & " " & legajo), legajo), rs!legajos4)
        legajos5 = IIf(semana = 5, IIf(rs!legajos5 <> "", IIf(Len(rs!legajos5 & " " & legajo) > 1000, Mid(rs!legajos5 & " " & legajo, 1, 1000), rs!legajos5 & " " & legajo), legajo), rs!legajos5)
        
        StrSql = " UPDATE rep_empl_sem_det SET cantidad1 = " & cantidad1
        StrSql = StrSql & " ,total1 = " & total1
        StrSql = StrSql & " ,legajos1 = '" & legajos1 & "'"
        StrSql = StrSql & " ,cantidad2 = " & cantidad2
        StrSql = StrSql & " ,total2 = " & total2
        StrSql = StrSql & " ,legajos2 = '" & legajos2 & "'"
        StrSql = StrSql & " ,cantidad3 = " & cantidad3
        StrSql = StrSql & " ,total3 = " & total3
        StrSql = StrSql & " ,legajos3 = '" & legajos3 & "'"
        StrSql = StrSql & " ,cantidad4 = " & cantidad4
        StrSql = StrSql & " ,total4 = " & total4
        StrSql = StrSql & " ,legajos4 = '" & legajos4 & "'"
        StrSql = StrSql & " ,cantidad5 = " & cantidad5
        StrSql = StrSql & " ,total5 = " & total5
        StrSql = StrSql & " ,legajos5 = '" & legajos5 & "'"
        StrSql = StrSql & " WHERE bpronro = " & NroProceso
        StrSql = StrSql & " AND estrnro_suc = " & estrnro_suc
        StrSql = StrSql & " AND estrnro_pue = " & estrnro_pue
        StrSql = StrSql & " AND estrnro_regh = " & estrnro_reghor
        objConn.Execute StrSql, , adExecuteNoRecords
    
    Else
        'insert
        
        'Busco el area de la sucursal, cargado como un codigo de la estructura de tipo Sucursal
        StrSql = "SELECT nrocod FROM estr_cod WHERE tcodnro = 38 AND estrnro = " & estrnro_suc
        OpenRecordset StrSql, rs
        area = ""
        If Not rs.EOF Then
            area = IIf(Not EsNulo(rs!nrocod!), CStr(rs!nrocod), " ")
        End If
        rs.Close
        
        'Busco el coeficiente del reg. horario, cargado como un codigo de la estructura de tipo Reg. Horario
        StrSql = "SELECT nrocod FROM estr_cod WHERE tcodnro = 39 AND estrnro = " & estrnro_reghor
        OpenRecordset StrSql, rs
        coeficiente = 0
        If Not rs.EOF Then
            coeficiente = IIf(Not EsNulo(rs!nrocod!), CDbl(rs!nrocod), 0)
        End If
        rs.Close
        
        cantidad1 = IIf(semana = 1, 1, 0)
        cantidad2 = IIf(semana = 2, 1, 0)
        cantidad3 = IIf(semana = 3, 1, 0)
        cantidad4 = IIf(semana = 4, 1, 0)
        cantidad5 = IIf(semana = 5, 1, 0)
        
        total1 = IIf(semana = 1, coeficiente, 0)
        total2 = IIf(semana = 2, coeficiente, 0)
        total3 = IIf(semana = 3, coeficiente, 0)
        total4 = IIf(semana = 4, coeficiente, 0)
        total5 = IIf(semana = 5, coeficiente, 0)
        
        legajos1 = IIf(semana = 1, legajo, "")
        legajos2 = IIf(semana = 2, legajo, "")
        legajos3 = IIf(semana = 3, legajo, "")
        legajos4 = IIf(semana = 4, legajo, "")
        legajos5 = IIf(semana = 5, legajo, "")
        
        'Busco el coeficiente de la sucursal, cargado como un codigo de la estructura de tipo Sucursal
        StrSql = " INSERT INTO rep_empl_sem_det (bpronro,estrnro_suc,estrdabr_suc,desc_suc,estrcodext_suc,estrnro_pue,"
        StrSql = StrSql & "estrdabr_pue,estrnro_regh,estrdabr_regh,coeficiente,cantidad1,total1,legajos1,cantidad2,total2,"
        StrSql = StrSql & "legajos2,cantidad3,total3,legajos3,cantidad4,total4,legajos4,cantidad5,total5,legajos5) "
        StrSql = StrSql & " VALUES ("
        StrSql = StrSql & NroProceso & ","
        StrSql = StrSql & estrnro_suc & ","
        StrSql = StrSql & "'" & estrdabr_suc & "',"
        StrSql = StrSql & "'" & area & "',"
        StrSql = StrSql & "'" & estrcodext_suc & "',"
        StrSql = StrSql & estrnro_pue & ","
        StrSql = StrSql & "'" & estrdabr_pue & "',"
        StrSql = StrSql & estrnro_reghor & ","
        StrSql = StrSql & "'" & estrdabr_reghor & "',"
        StrSql = StrSql & coeficiente & ","
        StrSql = StrSql & cantidad1 & ","
        StrSql = StrSql & total1 & ","
        StrSql = StrSql & "'" & legajos1 & "',"
        StrSql = StrSql & cantidad2 & ","
        StrSql = StrSql & total2 & ","
        StrSql = StrSql & "'" & legajos2 & "',"
        StrSql = StrSql & cantidad3 & ","
        StrSql = StrSql & total3 & ","
        StrSql = StrSql & "'" & legajos3 & "',"
        StrSql = StrSql & cantidad4 & ","
        StrSql = StrSql & total4 & ","
        StrSql = StrSql & "'" & legajos4 & "',"
        StrSql = StrSql & cantidad5 & ","
        StrSql = StrSql & total5 & ","
        StrSql = StrSql & "'" & legajos5 & "')"
        objConn.Execute StrSql, , adExecuteNoRecords
    
    End If
    
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    
End Sub

Sub buscar_estr(ByVal pe_ternro As Long, ByVal pe_tenro As Integer, ByVal pe_estrnro As Long, ByVal pe_fecha As Date, ByRef ps_estrnro As Long, ByRef ps_estrdabr As String, ByRef ps_estrcodext As String)
 Dim rs_his_estr As New ADODB.Recordset
    
    StrSql = "SELECT estructura.estrnro, estrdabr, estrcodext FROM his_estructura "
    StrSql = StrSql & " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro "
    StrSql = StrSql & " WHERE his_estructura.ternro = " & pe_ternro
    StrSql = StrSql & " AND his_estructura.tenro = " & pe_tenro & " AND htetdesde <" & ConvFecha(DateAdd("d", 7, pe_fecha))
    StrSql = StrSql & " AND (htethasta IS NULL OR htethasta >= " & ConvFecha(pe_fecha) & ")"
    If Not EsNulo(pe_estrnro) And pe_estrnro <> 0 Then
        StrSql = StrSql & " AND his_estructura.estrnro = " & pe_estrnro
    End If
    StrSql = StrSql & " ORDER BY htetdesde DESC"
    OpenRecordset StrSql, rs_his_estr
    ps_estrnro = 0
    ps_estrdabr = ""
    If Not rs_his_estr.EOF Then
        ps_estrnro = rs_his_estr!estrnro
        ps_estrdabr = rs_his_estr!estrdabr
        ps_estrcodext = rs_his_estr!estrcodext
    End If
    
    If rs_his_estr.State = adStateOpen Then rs_his_estr.Close
    Set rs_his_estr = Nothing

End Sub

Sub cargar_semanas(ByRef semanas_fec() As Date, ByVal p_mes As Integer, ByVal p_anio As Integer)
Dim Fecha As Date
Dim nro_semana As Integer
    
    nro_semana = 0
    Fecha = CDate("01/" & CStr(p_mes) & "/" & CStr(p_anio))
    Do While CInt(Month(Fecha)) = CInt(p_mes)
        If Weekday(Fecha) = vbMonday Then
            semanas_fec(nro_semana) = Fecha
            nro_semana = nro_semana + 1
        End If
        Fecha = DateAdd("d", 1, Fecha)
    Loop
    
End Sub
