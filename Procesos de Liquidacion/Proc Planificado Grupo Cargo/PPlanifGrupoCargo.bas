Attribute VB_Name = "ProcPlanificadoGrupoCargo"
Option Explicit

'Global Const Version = "1.00"
'Global Const FechaModificacion = "20/10/2011"
'Global Const UltimaModificacion = "Inicial"
'AUTOR = Manterola Maria Magdalena

Global Const Version = "1.01"
Global Const FechaModificacion = "04/06/2012" 'Manterola Maria Magdalena
Global Const UltimaModificacion = "Se eliminaron ciertos mensajes que se habian creado para debuggear y se modificó los nombres de columnas en la tabla CentroCosto"

Dim fs, f

Dim NroProceso As Long

Global Path As String
Global Rta
Global HuboErrores As Boolean

Global IdUser As String
Global Fecha As Date
Global hora As String

Global objconnBaseGrupoCargo As New ADODB.Connection


Private Sub Main()

Dim Directorio As String
Dim CArchivos
Dim archivo
Dim Folder
Dim strCmdLine As String
Dim Nombre_Arch As String

Dim StrSql As String
Dim objRs As New ADODB.Recordset

Dim PID As String
Dim parametros As String
Dim ArrParametros

Dim fechadesde As Date
Dim fechahasta As Date
Dim horadesde As String
Dim horahasta As String
Dim sinprocesar As Integer
Dim hayfiltro As Boolean
Dim fechainiej As String
Dim horainiej As String

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

    Nombre_Arch = PathFLog & "Interface_Grupo_Cargo" & "-" & NroProceso & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    TiempoInicialProceso = GetTickCount
    On Error Resume Next
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
    
    On Error GoTo ME_Main:
    
    HuboErrores = False
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Version = " & Version
    Flog.writeline "Modificacion = " & UltimaModificacion
    Flog.writeline "Fecha = " & FechaModificacion
    Flog.writeline "-----------------------------------------------------------------"
    
    Flog.writeline
    Flog.writeline "PID = " & PID
    
    Flog.writeline "Inicio Proceso de Planificacion Grupo Cargo : " & Now
    Flog.writeline "Cambio el estado del proceso a Procesando"
    
    'Cambio el estado del proceso a Procesando
    fechainiej = ConvFecha(Date)
    horainiej = Format(Now, "hh:mm:ss ")
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & horainiej & "', bprcfecinicioej = " & fechainiej & ", bprcestado = 'Procesando', bprcpid = " & PID & ", bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Obtengo los parametros del proceso"
    
    TiempoAcumulado = GetTickCount
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, objRs
    
    If Not objRs.EOF Then
       IdUser = objRs!IdUser
       Fecha = objRs!bprcfecha
       hora = objRs!bprchora
       
       'Obtengo los parametros del proceso
              
       If IsNull(objRs!bprcparam) Then
            parametros = ""
       Else
            parametros = objRs!bprcparam
       End If
       
       hayfiltro = False
       sinprocesar = 0
       
       ' Armar una alerta que avice que una auditoria no se modifico.
       If parametros <> "" Then
            ArrParametros = Split(parametros, "@")
            
            fechadesde = ArrParametros(0)
            horadesde = ArrParametros(1)
            fechahasta = ArrParametros(2)
            horahasta = ArrParametros(3)
            sinprocesar = ArrParametros(4)
            
            hayfiltro = True
       Else
            ' Si no hay parametros, esta programado automaticamente. Sacar la fecha desde de ultimo
            ' proceso del mismo tipo
            
            StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 305 AND bpronro = " & NroProceso & " ORDER BY bpronro DESC"
            OpenRecordset StrSql, objRs
            If Not objRs.EOF Then
                fechadesde = objRs!bprcfecInicioEj
                horadesde = objRs!bprcHoraInicioEj
                fechahasta = Format(Mid(fechainiej, 2, Len(fechainiej) - 2), FormatoInternoFecha)
                horahasta = horainiej
                sinprocesar = -1
            Else
                fechadesde = Format(Mid(fechainiej, 2, Len(fechainiej) - 2), FormatoInternoFecha)
                horadesde = Format("00:00:00", "hh:mm:ss ")
                fechahasta = Format(Mid(fechainiej, 2, Len(fechainiej) - 2), FormatoInternoFecha)
                horahasta = horainiej
                sinprocesar = -1
            End If
       End If
       
       Flog.writeline Espacios(Tabulador * 1) & "FECHA DESDE    :" & fechadesde
       Flog.writeline Espacios(Tabulador * 1) & "HORA DESDE     :" & horadesde
       Flog.writeline Espacios(Tabulador * 1) & "FECHA HASTA    :" & fechahasta
       Flog.writeline Espacios(Tabulador * 1) & "HORA HASTA     :" & horahasta
       Flog.writeline Espacios(Tabulador * 1) & "AUD. SIN PROC. :" & sinprocesar
       Flog.writeline
       
       ' Proceso que migra (exporta) los datos
       Call ComenzarTransferencia(hayfiltro, fechadesde, fechahasta, horadesde, horahasta, sinprocesar)
       
       ' Hacer el pasaje a la base del cliente
       'Call Comenzarmigracion
       
    Else
        Exit Sub
    End If
   
    'Actualizo el estado del proceso
    If Not HuboErrores Then
       StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
       Flog.writeline
       Flog.writeline "Proceso Finalizado Correctamente"
    Else
       StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProceso
       Flog.writeline
       Flog.writeline "Proceso Incompleto"
    End If
    
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
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
        StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Error General', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
    MyCommitTrans

End Sub


Public Sub ComenzarTransferencia(ByVal hayfiltro As Boolean, ByVal fechadesde As Date, ByVal fechahasta As Date, ByVal horadesde As String, ByVal horahasta As String, ByVal sinprocesar As Integer)

Dim objAuditoria As New ADODB.Recordset
Dim objrS1 As New ADODB.Recordset
Dim objRs2 As New ADODB.Recordset
Dim v_legajo As String
Dim v_desc As String
Dim v_apellido_nombre As String
Dim sin_error As Boolean

Dim v_fecha As Date
Dim v_estado As String
Dim v_ccosto As String
Dim v_fecha_ingreso As Date
Dim v_est_Nivel As String
Dim v_calledom As String
Dim v_nrodom As String
Dim v_piso As String
Dim v_depto As String
Dim v_tel As String
Dim v_ternro As Integer
Dim v_codigo As String
    
    If HuboErrores Then
        GoTo Fin:
    End If
        
    StrSql = "SELECT * FROM  auditoria "
    StrSql = StrSql & " WHERE (auditoria.aud_fec >= " & ConvFecha(fechadesde) & " ) AND "
    StrSql = StrSql & " (auditoria.aud_fec <= " & ConvFecha(fechadesde) & " ) "
    StrSql = StrSql & " ORDER BY aud_fec ASC,aud_hor ASC"
    Flog.writeline "CONSULTA GENERAL --> " & StrSql
    Flog.writeline
    OpenRecordset StrSql, objAuditoria
    
    'Determino la proporcion de progreso
    Progreso = 0
    CEmpleadosAProc = objAuditoria.RecordCount
    If CEmpleadosAProc = 0 Then
        CEmpleadosAProc = 1
    End If
    IncPorc = (95 / CEmpleadosAProc)
    
    'Actualizo el progreso
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & IncPorc & " WHERE bpronro = " & NroProceso
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    v_ternro = 0
    
    If objAuditoria.EOF Then
        Flog.writeline "----------------------------------------------------------------"
        Flog.writeline "No se registraron cambios en la auditoria."
        Flog.writeline "Revisar que los cambios se hayan realizando sobre los datos del Empleado y/o sobre los datos del Centro de Costo del Empleado (Historico)."
        Flog.writeline "Cualquier duda ver Documento de Configuracion."
        Flog.writeline "----------------------------------------------------------------"
        
    End If
    
    Do Until objAuditoria.EOF
               
        Flog.writeline "----------------------------------------------------------------"
       
        sin_error = True
        
        If Not IsNull(objAuditoria!aud_ternro) Then
            
            If v_ternro <> objAuditoria!aud_ternro And objAuditoria!aud_ternro <> 0 Then

                v_ternro = objAuditoria!aud_ternro
                
                StrSql = "SELECT tercero.terape, tercero.ternom, empleado.empest,empleado.empleg,empleado.empfaltagr, "
                StrSql = StrSql & " detdom.calle, detdom.nro, detdom.piso, detdom.oficdepto, telefono.telnro  "
                StrSql = StrSql & " FROM tercero INNER JOIN empleado ON tercero.ternro=empleado.ternro "
                StrSql = StrSql & " LEFT JOIN cabdom ON cabdom.ternro = empleado.ternro AND cabdom.tidonro = 2 "
                StrSql = StrSql & " LEFT JOIN detdom ON cabdom.domnro = detdom.domnro "
                StrSql = StrSql & " LEFT JOIN telefono ON telefono.domnro = cabdom.domnro "
                StrSql = StrSql & " WHERE tercero.ternro = " & v_ternro
                StrSql = StrSql & " ORDER BY telefono.tipotel ASC "
                            
                OpenRecordset StrSql, objrS1
                
                               
                If Not objrS1.EOF Then
                    v_legajo = objrS1!empleg
                    v_apellido_nombre = objrS1!terape & " , " & objrS1!ternom
                    
                    Flog.writeline Espacios(Tabulador * 1) & "INICIO EMPLEADO:" & v_legajo & " " & v_apellido_nombre
                    Flog.writeline
                    
                    If Not IsNull(objrS1!calle) Then
                        v_calledom = objrS1!calle
                    Else
                        v_calledom = ""
                    End If
                    If Not IsNull(objrS1!nro) Then
                        v_nrodom = objrS1!nro
                    Else
                        v_nrodom = ""
                    End If
                    If Not IsNull(objrS1!piso) Then
                        v_piso = objrS1!piso
                    Else
                        v_piso = ""
                    End If
                    If Not IsNull(objrS1!oficdepto) Then
                        v_depto = objrS1!oficdepto
                    Else
                        v_depto = ""
                    End If
                    If Not IsNull(objrS1!telnro) Then
                        v_tel = objrS1!telnro
                    Else
                        v_tel = ""
                    End If
                                                    
                    If objrS1!empest = -1 Then
                        v_estado = "A"
                    Else
                        v_estado = "I"
                    End If
                                                    
                    If Not IsNull(objrS1!empfaltagr) And objrS1!empfaltagr <> "" Then
                        v_fecha_ingreso = objrS1!empfaltagr
                    Else
                        v_fecha_ingreso = "01/01/1900"
                    End If
                    
                    Call InsertarItemEmple1y2(fechadesde, fechahasta, v_ternro, v_legajo, v_apellido_nombre, v_estado, v_calledom, v_nrodom, v_piso, v_depto, v_tel, v_fecha_ingreso, sin_error)
                Else
                    Flog.writeline "La siguiente Consulta es Vacia --> " & StrSql
                    Flog.writeline "Por lo que no se insertaron items en Emple1 y Emple2"
                    Flog.writeline
                End If
                                                                                                    
                StrSql = " SELECT empleado.empleg, hist.tenro, hist.estrnro, "
                StrSql = StrSql & " hist.htetdesde, hist.htethasta, est.estrdabr "
                StrSql = StrSql & " FROM empleado "
                StrSql = StrSql & " INNER JOIN his_estructura hist ON hist.ternro = empleado.ternro "
                StrSql = StrSql & " AND hist.htetdesde <= " & ConvFecha(fechadesde)
                StrSql = StrSql & " AND (hist.htethasta >= " & ConvFecha(fechadesde)
                StrSql = StrSql & " OR hist.htethasta IS NULL )"
                StrSql = StrSql & " INNER JOIN estructura est ON hist.estrnro = est.estrnro "
                StrSql = StrSql & " WHERE empleado.ternro = " & v_ternro
                StrSql = StrSql & " AND hist.tenro = 5 "
                StrSql = StrSql & " ORDER BY hist.htetdesde DESC "
                    
                OpenRecordset StrSql, objrS1
                                                       
                If Not objrS1.EOF Then
                    
                    v_desc = objrS1!estrdabr
                    v_codigo = objrS1!Estrnro
                    v_legajo = objrS1!empleg
                    If IsNull(objrS1!htethasta) Or objrS1!htethasta = "" Then
                        v_fecha = CDate(objrS1!htetdesde)
                        v_estado = "Activo"
                    Else
                        v_fecha = CDate(objrS1!htethasta)
                        v_estado = "Inactivo"
                    End If
                    If sin_error Then
                        Call InsertarItemCentroCosto(v_codigo, v_desc, v_estado, sin_error)
                    End If
                Else
                    Flog.writeline "La siguiente Consulta es Vacia --> " & StrSql
                    Flog.writeline "Por lo que no se insertaron items en CentroCosto"
                    Flog.writeline
                End If
                
            End If
        Else
            Flog.writeline "El valor de objAuditoria!aud_ternro es Nulo"
            Flog.writeline
        End If
                                 
    
        'Actualizo el progreso
        Progreso = Progreso + IncPorc
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & " WHERE bpronro = " & NroProceso
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
        
        objAuditoria.MoveNext
        
    Loop
    Flog.writeline "----------------------------------------------------------------"
        
Fin:
    Exit Sub
    

End Sub

Private Sub EliminarItemCentroCosto(ByVal codigo As String)

'CentroCosto
'@cto_codigo
'cto_nombre
'estado (activo / inactivo)

MyBeginTrans

    Flog.writeline Espacios(Tabulador * 1) & "ELIMINANDO ITEM CentroCosto"
    
    StrSql = "DELETE FROM CentroCosto "
    StrSql = StrSql & " WHERE cto_codigo = '" & codigo & "'"
    
    objConn.Execute StrSql, , adExecuteNoRecords
       
MyCommitTrans
    

End Sub

Private Sub InsertarItemCentroCosto(ByVal codigo As String, ByVal descripcion As String, ByVal estado As String, ByVal no_error As Boolean)

Dim objInsert As New ADODB.Recordset

'CentroCosto
'@cto_codigo
'cto_nombre
'estado (activo / inactivo)


MyBeginTrans
    Flog.writeline Espacios(Tabulador * 1) & "ACTUALIZANDO CentroCosto"
    
    StrSql = "SELECT * FROM CentroCosto WHERE cto_codigo = '" & codigo & "'"
    
    OpenRecordset StrSql, objInsert
    
    If objInsert.EOF Then
        StrSql = "INSERT INTO CentroCosto(cto_codigo,cto_nombre,estado)"
        StrSql = StrSql & " VALUES ('" & Mid(codigo, 1, 6) & "','"
        StrSql = StrSql & Mid(descripcion, 1, 60) & "','"
        StrSql = StrSql & Mid(estado, 1, 8) & "')"
    Else
        StrSql = "UPDATE CentroCosto SET "
        StrSql = StrSql & "cto_nombre = '" & Mid(descripcion, 1, 60) & "',"
        StrSql = StrSql & "estado= '" & Mid(estado, 1, 8) & "'"
        StrSql = StrSql & " WHERE cto_codigo = " & objInsert!cto_codigo
    End If
    
    objConn.Execute StrSql, , adExecuteNoRecords
        
MyCommitTrans

Fin:
    If objInsert.State = adStateOpen Then objInsert.Close
    Set objInsert = Nothing
    Exit Sub
    

End Sub
Private Sub EliminarItemEple1y2(ByVal Legajo As Integer, ByVal Ternro As Integer)

Dim objInsert As New ADODB.Recordset
Dim CodSuc As Integer

MyBeginTrans

    CodSuc = 0
    
    Flog.writeline Espacios(Tabulador * 1) & "ELIMINANDO ITEM EMPLE1 Y EMPLE2"
    
    'Primero busco la sucursal del empleado
    StrSql = " SELECT * FROM his_estructura hist "
    StrSql = StrSql & " INNER JOIN estructura est ON est.estrnro = hist.estrnro "
    StrSql = StrSql & " WHERE hist.ternro = " & Ternro
    StrSql = StrSql & " AND hist.tenro = 1 "
    OpenRecordset StrSql, objInsert
    
    
    If Not objInsert.EOF Then
        CodSuc = objInsert!Estrnro
    End If
        
    'Elimino el registro de la tabla Emple1
    StrSql = " DELETE FROM Emple1 "
    StrSql = StrSql & " WHERE Emple1.Emp_legajo = " & Legajo
    If CodSuc <> "" Then
        StrSql = StrSql & " AND Emple1.suc_codigo = " & CodSuc
    End If
    objConn.Execute StrSql, , adExecuteNoRecords
    
    
    'Elimino el registro de la tabla Emple2
    StrSql = " DELETE FROM Emple2 "
    StrSql = StrSql & " WHERE suc_codigo = " & CodSuc
    If CodSuc <> "" Then
        StrSql = StrSql & " AND Emp_legajo = " & Legajo
    End If
    objConn.Execute StrSql, , adExecuteNoRecords
    

    
MyCommitTrans

Fin:
    If objInsert.State = adStateOpen Then objInsert.Close
    Set objInsert = Nothing
    Exit Sub
    

End Sub

Private Sub InsertarItemEmple1y2(ByVal fechadesde As Date, ByVal fechahasta As Date, ByVal Ternro As Integer, ByVal Legajo As String, ByVal ape_nom As String, ByVal estado As String, ByVal calle As String, ByVal nrodomi As String, ByVal piso As String, ByVal depto As String, ByVal tel As String, ByVal Fecha As Date, ByVal no_error As Boolean)

Dim objInsert As New ADODB.Recordset
Dim CodSuc As Integer
Dim CodCC As String
Dim CodDep As Integer
Dim CodDir As String
Dim CodCateg As Integer
Dim CodSind As Integer
Dim FechaSuc As Date
Dim CodFormLiq As String
Dim HorasExtras As String

Dim TerCuil As String

MyBeginTrans

    CodSuc = 0
    CodCC = ""
    CodDep = 0
    CodDir = ""
    CodCateg = 0
    CodSind = 0
    
    Flog.writeline Espacios(Tabulador * 1) & "ACTUALIZANDO EMPLE1 Y EMPLE2"
    
    'Primero busco los códigos y descripciones de las estructuras para el empleado
    Flog.writeline Espacios(Tabulador * 1) & "Busqueda de estructuas del empleado"

    StrSql = " SELECT * FROM his_estructura hist "
    StrSql = StrSql & " INNER JOIN estructura est ON est.estrnro = hist.estrnro "
    StrSql = StrSql & " WHERE hist.ternro = " & Ternro
    StrSql = StrSql & " AND hist.htetdesde <=" & ConvFecha(fechadesde)
    StrSql = StrSql & " AND (hist.htethasta >= " & ConvFecha(fechadesde)
    StrSql = StrSql & " OR hist.htethasta IS NULL )"
    StrSql = StrSql & " AND hist.tenro IN (1,3,5,9,16,22,35)"
    OpenRecordset StrSql, objInsert
       
    Do Until objInsert.EOF
        Select Case objInsert!Tenro
            Case 1: 'Sucursal
                    CodSuc = objInsert!Estrnro
                    FechaSuc = objInsert!htetdesde
            Case 5: 'Centro de Costo
                    CodCC = objInsert!estrdabr
            Case 9: 'Departamento
                    CodDep = objInsert!Estrnro
            Case 35: 'Direccion/Gerencia
                    CodDir = objInsert!estrdabr
            Case 3: 'Categoria
                    CodCateg = objInsert!Estrnro
            Case 16: 'Sindicato
                    CodSind = objInsert!Estrnro
            Case 22: 'Forma De Liquidación
                If objInsert!estrcodext = "M" Then
                    CodFormLiq = "M"
                Else
                    If objInsert!estrcodext = "J" Then
                        CodFormLiq = "J"
                    End If
                End If
                
        End Select
        objInsert.MoveNext
    Loop
        
      
    'Ahora busco el cuil del empleado
    Flog.writeline Espacios(Tabulador * 1) & "Busqueda del Cuil del Empleado"

    StrSql = " SELECT * FROM ter_doc "
    StrSql = StrSql & " WHERE ter_doc.ternro = " & Ternro
    StrSql = StrSql & " AND ter_doc.tidnro = 10 "
    OpenRecordset StrSql, objInsert
    If Not objInsert.EOF Then
        TerCuil = Mid(Replace(objInsert!nrodoc, "-", ""), 1, 11)
    Else
        TerCuil = ""
    End If
    
    
    'Busco si el empleado hace horas extras o no
    Flog.writeline Espacios(Tabulador * 1) & "Busqueda de Horas Extras del Empleado"

    StrSql = " SELECT * FROM gti_cabparte "
    StrSql = StrSql & " INNER JOIN gti_autdet ON gti_autdet.gcpnro = gti_cabparte.gcpnro "
    StrSql = StrSql & " WHERE gti_autdet.ternro = " & Ternro
    StrSql = StrSql & " AND gti_cabparte.gcpdesde <= " & ConvFecha(fechadesde)
    StrSql = StrSql & " AND gti_cabparte.gcphasta >= " & ConvFecha(fechadesde)
    OpenRecordset StrSql, objInsert
    If objInsert.EOF Then
        HorasExtras = "N"
    Else
        HorasExtras = "S"
    End If
   
    'Ahora busco si ese registro ya esta en la tabla Emple1
    StrSql = "SELECT * FROM Emple1 WHERE Emp_legajo = " & Legajo
    OpenRecordset StrSql, objInsert
    
    If objInsert.EOF Then
        StrSql = "INSERT INTO Emple1(suc_codigo,Emp_legajo,Emp_nombre,Emp_cuil)"
        StrSql = StrSql & " VALUES (" & CodSuc & ","
        StrSql = StrSql & Legajo & ",'"
        StrSql = StrSql & Mid(ape_nom, 1, 30) & "','"
        StrSql = StrSql & TerCuil & "')"
    Else
        StrSql = "UPDATE Emple1 SET "
        StrSql = StrSql & " Emp_nombre= '" & Mid(ape_nom, 1, 30) & "',"
        StrSql = StrSql & " Emp_cuil= '" & TerCuil & "',"
        StrSql = StrSql & " suc_codigo= " & CodSuc
        StrSql = StrSql & " WHERE Emp_legajo = " & Legajo
    End If
        
    objConn.Execute StrSql, , adExecuteNoRecords

    'Ahora busco si ese registro ya esta en la tabla Emple2
    StrSql = "SELECT * FROM Emple2 WHERE Emp_legajo = " & Legajo
    OpenRecordset StrSql, objInsert
    
    If objInsert.EOF Then
        StrSql = "INSERT INTO Emple2(suc_codigo,Emp_legajo,Emp_vigenc,Emp_cuil,"
        StrSql = StrSql & " Cto_codigo,Emp_situac,Emp_hsext,Emp_tipliq,"
        StrSql = StrSql & " Emp_fecing,Emp_domcal,Emp_domnum,Emp_dompis,Emp_domdep,"
        StrSql = StrSql & " Emp_telefo,Empdpto,Empdpdes,Kat_codigo,Emp_sinnro)"
        StrSql = StrSql & " VALUES (" & CodSuc & ","
        StrSql = StrSql & Legajo & ","
        StrSql = StrSql & ConvFecha(FechaSuc) & ",'"
        StrSql = StrSql & TerCuil & "','"
        StrSql = StrSql & Mid(CodCC, 1, 6) & "','"
        StrSql = StrSql & estado & "','"
        StrSql = StrSql & HorasExtras & "','"
        StrSql = StrSql & CodFormLiq & "',"
        StrSql = StrSql & ConvFecha(Fecha) & ",'"
        StrSql = StrSql & Mid(calle, 1, 30) & "','"
        StrSql = StrSql & Mid(nrodomi, 1, 5) & "','"
        StrSql = StrSql & Mid(piso, 1, 3) & "','"
        StrSql = StrSql & Mid(depto, 1, 3) & "','"
        StrSql = StrSql & Mid(Replace(tel, " ", ""), 1, 12) & "',"
        StrSql = StrSql & CodDep & ",'"
        StrSql = StrSql & Mid(CodDir, 1, 35) & "',"
        StrSql = StrSql & CodCateg & ","
        StrSql = StrSql & CodSind & ")"
    Else
        StrSql = "UPDATE Emple2 SET "
        StrSql = StrSql & " suc_codigo= " & CodSuc & ","
        StrSql = StrSql & " Emp_vigenc = " & ConvFecha(FechaSuc) & ","
        StrSql = StrSql & " Emp_cuil = '" & TerCuil & "',"
        StrSql = StrSql & " Cto_codigo = '" & Mid(CodCC, 1, 6) & "',"
        StrSql = StrSql & " Emp_situac = '" & estado & "',"
        StrSql = StrSql & " Emp_hsext = '" & HorasExtras & "',"
        StrSql = StrSql & " Emp_tipliq = '" & CodFormLiq & "',"
        StrSql = StrSql & " Emp_fecing = " & ConvFecha(Fecha) & ","
        StrSql = StrSql & " Emp_domcal = '" & Mid(calle, 1, 30) & "',"
        StrSql = StrSql & " Emp_domnum = '" & Mid(nrodomi, 1, 5) & "',"
        StrSql = StrSql & " Emp_dompis = '" & Mid(piso, 1, 3) & "',"
        StrSql = StrSql & " Emp_domdep = '" & Mid(depto, 1, 3) & "',"
        StrSql = StrSql & " Emp_telefo = '" & Mid(Replace(tel, " ", ""), 1, 12) & "',"
        StrSql = StrSql & " Empdpto = " & CodDep & ","
        StrSql = StrSql & " Empdpdes = '" & Mid(CodDir, 1, 35) & "',"
        StrSql = StrSql & " Kat_codigo = " & CodCateg & ","
        StrSql = StrSql & " Emp_sinnro = " & CodSind
        StrSql = StrSql & " WHERE Emp_legajo = " & Legajo
    End If
    
    objConn.Execute StrSql, , adExecuteNoRecords
    
MyCommitTrans

Fin:
    If objInsert.State = adStateOpen Then objInsert.Close
    Set objInsert = Nothing
    Exit Sub
    

End Sub

Private Sub Comenzarmigracion()
Dim objRs As New ADODB.Recordset
Dim objRsCargo As New ADODB.Recordset

    On Error Resume Next

    OpenConnection strconexion, objconnBaseGrupoCargo
    
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    If Err.Number <> 0 Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion Grupo Cargo"
        Exit Sub
    End If
    

MyBeginTrans
    
'-- CENTRO DE COSTO -----------------
    StrSql = "SELECT * FROM CentroCosto "
    
    OpenRecordset StrSql, objRs
    
    CEmpleadosAProc = objRs.RecordCount
    If CEmpleadosAProc = 0 Then
        CEmpleadosAProc = 1
    End If
    IncPorc = (1 / CEmpleadosAProc)
    
    Do Until objRs.EOF
        StrSql = "SELECT * FROM CentroCosto WHERE cto_codigo= " & objRs!cto_codigo
        
        If objRsCargo.State <> adStateClosed Then
            If objRsCargo.lockType <> adLockReadOnly Then objRsCargo.UpdateBatch
            objRsCargo.Close
        End If
        objRsCargo.CacheSize = 500
        objRsCargo.Open StrSql, objconnBaseGrupoCargo, adOpenDynamic, adLockReadOnly, adCmdText
    
        If objRsCargo.EOF Then
            'Insert
            StrSql = "INSERT INTO CentroCosto(cto_codigo,cto_nombre,estado)"
            StrSql = StrSql & " VALUES ('" & Mid(objRs!cto_codigo, 1, 6) & "','"
            StrSql = StrSql & Mid(objRs!cto_nombre, 1, 60) & "','"
            StrSql = StrSql & Mid(objRs!estado, 1, 8) & "')"
        Else
            'Update
            StrSql = "UPDATE CentroCosto SET "
            If Not EsNulo(objRs!cto_nombre) Then
                StrSql = StrSql & "cto_nombre = '" & Mid(objRs!cto_nombre, 1, 60) & "',"
            End If
            StrSql = StrSql & "estado= '" & Mid(objRs!estado, 1, 8) & "'"
            StrSql = StrSql & " WHERE cto_codigo = '" & Mid(objRs!cto_codigo, 1, 6) & "'"
            
        End If
        objconnBaseGrupoCargo.Execute StrSql, , adExecuteNoRecords
        
        'Actualizo el progreso
        Progreso = Progreso + IncPorc
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & " WHERE bpronro = " & NroProceso
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
        
        objRs.MoveNext
    Loop
    
    ' Borro los datos en la tabla temporal
    'StrSql = "DELETE FROM CentroCosto"
    'objConn.Execute StrSql, , adExecuteNoRecords
    
'-- EMPLE1 -----------------
    StrSql = "SELECT * FROM Emple1 "
    
    OpenRecordset StrSql, objRs
    
    CEmpleadosAProc = objRs.RecordCount
    If CEmpleadosAProc = 0 Then
        CEmpleadosAProc = 1
    End If
    IncPorc = (1 / CEmpleadosAProc)
    
    Do Until objRs.EOF
    
        StrSql = "SELECT * FROM Emple1 WHERE suc_codigo = '" & objRs!suc_codigo & "'"
        StrSql = StrSql & " AND Emp_legajo = " & objRs!Emp_legajo
        
        If objRsCargo.State <> adStateClosed Then
            If objRsCargo.lockType <> adLockReadOnly Then objRsCargo.UpdateBatch
            objRsCargo.Close
        End If
        objRsCargo.CacheSize = 500
        objRsCargo.Open StrSql, objconnBaseGrupoCargo, adOpenDynamic, adLockReadOnly, adCmdText

        If objRsCargo.EOF Then
            'INSERT
            StrSql = "INSERT INTO Emple1(suc_codigo,Emp_legajo,Emp_nombre,Emp_cuil)"
            StrSql = StrSql & " VALUES (" & objRs!suc_codigo & ","
            StrSql = StrSql & objRs!Emp_legajo & ",'"
            StrSql = StrSql & Mid(objRs!Emp_nombre, 1, 30) & "','"
            StrSql = StrSql & Mid(Replace(objRs!Emp_cuil, "-", ""), 1, 11) & "')"
        Else
            'UPDATE
            StrSql = "UPDATE Emple1 SET "
            StrSql = StrSql & " Emp_nombre= '" & Mid(objRs!Emp_nombre, 1, 30) & "',"
            StrSql = StrSql & " Emp_cuil= '" & Mid(Replace(objRs!Emp_cuil, "-", ""), 1, 11) & "'"
            StrSql = StrSql & " WHERE suc_codigo = " & objRs!suc_codigo
            StrSql = StrSql & " AND Emp_legajo = " & objRs!Emp_legajo
        End If
        objconnBaseGrupoCargo.Execute StrSql, , adExecuteNoRecords
        
        'Actualizo el progreso
        Progreso = Progreso + IncPorc
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & " WHERE bpronro = " & NroProceso
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
        
        objRs.MoveNext
    Loop
    
    ' Borro los datos en la tabla temporal
    'StrSql = "DELETE FROM Emple1"
    'objConn.Execute StrSql, , adExecuteNoRecords
        
'-- EMPLE2 -----------------
    StrSql = "SELECT * FROM Emple2 "
    
    OpenRecordset StrSql, objRs
    
    CEmpleadosAProc = objRs.RecordCount
    If CEmpleadosAProc = 0 Then
        CEmpleadosAProc = 1
    End If
    IncPorc = (1 / CEmpleadosAProc)
    
    Do Until objRs.EOF
        
        StrSql = "SELECT * FROM Emple2 WHERE suc_codigo = " & objRs!suc_codigo
        StrSql = StrSql & " AND Emp_legajo = " & objRs!Emp_legajo
        
        If objRsCargo.State <> adStateClosed Then
            If objRsCargo.lockType <> adLockReadOnly Then objRsCargo.UpdateBatch
            objRsCargo.Close
        End If
        objRsCargo.CacheSize = 500
        objRsCargo.Open StrSql, objconnBaseGrupoCargo, adOpenDynamic, adLockReadOnly, adCmdText
    
        If objRsCargo.EOF Then
            'INSERT
            StrSql = "INSERT INTO Emple2(suc_codigo,Emp_legajo,Emp_vigenc,Emp_cuil,"
            StrSql = StrSql & " Cto_codigo,Emp_situac,Emp_hsext,Emp_tipliq,"
            StrSql = StrSql & " Emp_fecing,Emp_domcal,Emp_domnum,Emp_dompis,Emp_domdep,"
            StrSql = StrSql & " Emp_telefo,Empdpto,Empdpdes,Kat_codigo,Emp_sinnro)"
            StrSql = StrSql & " VALUES (" & objRs!suc_codigo & ","
            StrSql = StrSql & objRs!Emp_legajo & ",'"
            StrSql = StrSql & objRs!Emp_vigenc & "','"
            StrSql = StrSql & Mid(Replace(objRs!Emp_cuil, "-", ""), 1, 11) & "','"
            StrSql = StrSql & Mid(objRs!cto_codigo, 1, 6) & "','"
            StrSql = StrSql & objRs!Emp_situac & "','"
            StrSql = StrSql & objRs!Emp_hsext & "','"
            StrSql = StrSql & objRs!Emp_tipliq & "','"
            StrSql = StrSql & objRs!Emp_fecing & "','"
            StrSql = StrSql & Mid(objRs!Emp_domcal, 1, 30) & "','"
            StrSql = StrSql & Mid(objRs!Emp_domnum, 1, 5) & "','"
            StrSql = StrSql & Mid(objRs!Emp_dompis, 1, 3) & "','"
            StrSql = StrSql & Mid(objRs!Emp_domdep, 1, 3) & "','"
            StrSql = StrSql & Mid(Replace(objRs!Emp_telefo, " ", ""), 1, 12) & "',"
            StrSql = StrSql & objRs!Empdpto & ",'"
            StrSql = StrSql & Mid(objRs!Empdpdes, 1, 35) & "',"
            StrSql = StrSql & objRs!Kat_codigo & ","
            StrSql = StrSql & objRs!Emp_sinnro & ")"
        Else
            'UPDATE
            StrSql = "UPDATE Emple2 SET "
            StrSql = StrSql & " Emp_vigenc = '" & objRs!Emp_vigenc & "','"
            StrSql = StrSql & " Emp_cuil = " & Mid(Replace(objRs!Emp_cuil, "-", ""), 1, 11) & "','"
            StrSql = StrSql & " Cto_codigo = " & Mid(objRs!cto_codigo, 1, 6) & "','"
            StrSql = StrSql & " Emp_situac = " & objRs!Emp_situac & "','"
            StrSql = StrSql & " Emp_hsext = " & objRs!Emp_hsext & "','"
            StrSql = StrSql & " Emp_tipliq = " & objRs!Emp_tipliq & "','"
            StrSql = StrSql & " Emp_fecing = " & objRs!Emp_fecing & "','"
            If Not EsNulo(objRs!Emp_domcal) Then
                StrSql = StrSql & " Emp_domcal = " & Mid(objRs!Emp_domcal, 1, 30) & "','"
            End If
            If Not EsNulo(objRs!Emp_domnum) Then
                StrSql = StrSql & " Emp_domnum = " & Mid(objRs!Emp_domnum, 1, 5) & "','"
            End If
            If Not EsNulo(objRs!Emp_dompis) Then
                StrSql = StrSql & " Emp_dompis = " & Mid(objRs!Emp_dompis, 1, 3) & "','"
            End If
            If Not EsNulo(objRs!Emp_domdep) Then
                StrSql = StrSql & " Emp_domdep = " & Mid(objRs!Emp_domdep, 1, 3) & "','"
            End If
            If Not EsNulo(objRs!Emp_telefo) Then
                StrSql = StrSql & " Emp_telefo = " & Mid(Replace(objRs!Emp_telefo, " ", ""), 1, 12) & "',"
            End If
            If Not EsNulo(objRs!Empdpto) Then
                StrSql = StrSql & " Empdpto = " & objRs!Empdpto & ",'"
            End If
            If Not EsNulo(objRs!Empdpdes) Then
                StrSql = StrSql & " Empdpdes = " & Mid(objRs!Empdpdes, 1, 35) & "',"
            End If
            If Not EsNulo(objRs!Kat_codigo) Then
                StrSql = StrSql & " Kat_codigo = " & objRs!Kat_codigo & ","
            End If
            If Not EsNulo(objRs!Emp_sinnro) Then
                StrSql = StrSql & " Emp_sinnro = " & objRs!Emp_sinnro
            End If
            
            StrSql = StrSql & " WHERE suc_codigo = " & objRs!suc_codigo
            StrSql = StrSql & " AND Emp_legajo = " & objRs!Emp_legajo
        End If

        objconnBaseGrupoCargo.Execute StrSql, , adExecuteNoRecords
        
        'Actualizo el progreso
        Progreso = Progreso + IncPorc
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & " WHERE bpronro = " & NroProceso
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
        
        objRs.MoveNext
    Loop
    
    ' Borro los datos en la tabla temporal
    'StrSql = "DELETE FROM Emple2"
    'objConn.Execute StrSql, , adExecuteNoRecords
     

MyCommitTrans
    
Fin:
    Exit Sub


End Sub
