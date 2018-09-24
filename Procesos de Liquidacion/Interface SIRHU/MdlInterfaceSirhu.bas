Attribute VB_Name = "MdlInterfaceSirhu"
Option Explicit

'Const Version = "1.0"
'Const FechaVersion = "29/10/2015" ' Sebastian Stremel - Version Inicial - CAS-32254 - Custom Interface SIRHU - INTERCARGO
'Const UltimaModificacion = ""

Const Version = "1.1"
Const FechaVersion = "15/02/2016" ' Sebastian Stremel - Se guarda el conccod en lugar del concnro - CAS-32254 - Custom Interface SIRHU - INTERCARGO [Entrega 2]
Const UltimaModificacion = ""


Global NroProceso As Long
Global HuboErrores As Boolean
Type confrep
    conftipo As String
    confval As Integer
    confval2 As String
End Type

Type concepto
    Ternro As Integer
    descabr As String
    ConcNro As Integer
    ConcCod As String
    tipoconc As Integer
End Type
Global objInsert As New ADODB.Connection
Global FechaDesde As String
Global FechaHasta As String
Global listaProcesos As String
Global arrConcepto() As concepto
Global porc As Integer

Private Sub Main()

Dim NombreArchivo As String
Dim directorio As String
Dim CArchivos
Dim archivo
Dim Folder
Dim strCmdLine As String
Dim Nombre_Arch As String

Dim StrSql As String
Dim objRs As New ADODB.Recordset
Dim objRs2 As New ADODB.Recordset

Dim ArrParametros
Dim PID As String
Dim parametros

Dim legdesde As Long
Dim leghasta As Long
Dim estado As Integer
Dim pliqdesde As Integer
Dim pliqhasta As Integer
Dim aprobado As Integer
'Dim listaProcesos As String
Dim tenro1 As Integer
Dim estrnro1 As Integer
Dim tenro2 As Integer
Dim estrnro2 As Integer
Dim tenro3 As Integer
Dim estrnro3 As Integer
Dim fecestr As String
Dim orden As String
Dim lote As Integer

Dim listaEmpleados As String
listaEmpleados = "0"
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

    NroProceso = NroProcesoBatch
    
    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas
    

    TiempoInicialProceso = GetTickCount

    Nombre_Arch = PathFLog & "InterfaceSirhu" & "-" & NroProceso & ".log"
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    Flog.writeline "Inicio Proceso de Interface SIRHU: " & Now
    Flog.writeline "Cambio el estado del proceso a Procesando"
   
    'Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Version = " & Version
    Flog.writeline "Modificacion = " & UltimaModificacion
    Flog.writeline "Fecha = " & FechaVersion
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline

    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If

    OpenConnection strconexion, objconnProgreso
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If

    HuboErrores = False
    On Error GoTo CE
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcprogreso = 0, bprcpid = " & PID & " WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Obtengo los datos del proceso"
    
    TiempoAcumulado = GetTickCount
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        'Obtengo los parametros del proceso
        parametros = objRs!bprcparam
        
        Flog.writeline "Parametros del proceso: " & parametros
        
        ArrParametros = Split(parametros, "@")
       
        legdesde = ArrParametros(0)
        Flog.writeline "legajo desde: " & legdesde
        
        leghasta = ArrParametros(1)
        Flog.writeline "legajo hasta: " & leghasta
        
        estado = ArrParametros(2)
        Flog.writeline "Estado de empleados: " & estado
        
        pliqdesde = ArrParametros(3)
        Flog.writeline "periodo desde: " & pliqdesde
        
        pliqhasta = ArrParametros(4)
        Flog.writeline "periodo hasta: " & pliqhasta
        
        aprobado = ArrParametros(5)
        Flog.writeline "Estado de los procesos: " & aprobado
        
        listaProcesos = ArrParametros(6)
        Flog.writeline "Lista de procesos: " & listaProcesos
        
        tenro1 = ArrParametros(7)
        Flog.writeline "Tipo Estructura 1: " & tenro1
        
        estrnro1 = ArrParametros(8)
        Flog.writeline "Estructura 1: " & estrnro1
        
        tenro2 = ArrParametros(9)
        Flog.writeline "Tipo Estructura 2: " & tenro2
        
        estrnro2 = ArrParametros(10)
        Flog.writeline "Estructura 2: " & estrnro2
        
        tenro3 = ArrParametros(11)
        Flog.writeline "Tipo Estructura 3: " & tenro3
        
        estrnro3 = ArrParametros(12)
        Flog.writeline "Estructura 3: " & estrnro3
        
        fecestr = ArrParametros(13)
        Flog.writeline "Fecha Estructuras: " & fecestr
        
        orden = ArrParametros(14)
        Flog.writeline "Orden: " & orden
       
        
        '------------------------------------------
        'Busco los empleados a procesar
        '------------------------------------------
        StrSql = "SELECT distinct(empleado) FROM cabliq "
        StrSql = StrSql & " INNER JOIN empleado on empleado.ternro = cabliq.empleado "
        If tenro1 <> 0 Then
            If estrnro1 <> 0 Then
                StrSql = StrSql & " INNER JOIN his_estructura h1 ON h1.ternro = empleado.ternro And h1.Tenro = " & tenro1 & " AND h1.estrnro=" & estrnro1
            Else
                StrSql = StrSql & " INNER JOIN his_estructura h1 ON h1.ternro = empleado.ternro And h1.Tenro = " & tenro1
            End If
            StrSql = StrSql & " AND (h1.htetdesde<=" & ConvFecha(fecestr) & " AND (h1.htethasta is null or h1.htethasta>=" & ConvFecha(fecestr) & "))"
        End If
        
        If tenro2 <> 0 Then
            If estrnro2 <> 0 Then
                StrSql = StrSql & " INNER JOIN his_estructura h2 ON h2.ternro = empleado.ternro And h2.Tenro = " & tenro2 & " AND h2.estrnro=" & estrnro2
            Else
                StrSql = StrSql & " INNER JOIN his_estructura h2 ON h2.ternro = empleado.ternro And h2.Tenro = " & tenro2
            End If
            StrSql = StrSql & " AND (h2.htetdesde<=" & ConvFecha(fecestr) & " AND (h2.htethasta is null or h2.htethasta>=" & ConvFecha(fecestr) & "))"
        End If
        
        If tenro3 <> 0 Then
            If estrnro3 <> 0 Then
                StrSql = StrSql & " INNER JOIN his_estructura h3 ON h3.ternro = empleado.ternro And h3.Tenro = " & tenro3 & " AND h3.estrnro=" & estrnro3
            Else
                StrSql = StrSql & " INNER JOIN his_estructura h3 ON h3.ternro = empleado.ternro And h3.Tenro = " & tenro3
            End If
            StrSql = StrSql & " AND (h3.htetdesde<=" & ConvFecha(fecestr) & " AND (h3.htethasta is null or h3.htethasta>=" & ConvFecha(fecestr) & "))"
        End If
        StrSql = StrSql & " Where (Empleado.empleg >= " & legdesde & " And Empleado.empleg <= " & leghasta & ")"
        If estado <> 1 Then
            StrSql = StrSql & " AND empest= " & estado
        End If
        StrSql = StrSql & " AND cabliq.pronro in (" & listaProcesos & ")"
        Flog.writeline "****" & StrSql & "****"
        OpenRecordset StrSql, objRs2
        If Not objRs2.EOF Then
            Do While Not objRs2.EOF
                listaEmpleados = listaEmpleados & "," & objRs2!Empleado
            objRs2.MoveNext
            Loop
        Else
            Flog.writeline "No hay empleados para procesar"
            StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
            objConn.Execute StrSql, , adExecuteNoRecords
            Flog.writeline "Fin :" & Now
            Flog.Close
            Exit Sub
        End If
        objRs2.Close
        '------------------------------------------
        'FIN BUSQUEDA EMPLEADOS
        '------------------------------------------
        
        '------------------------------------------
        'BUSCO EL LOTE ANTERIOR
        StrSql = "SELECT max(lote) lote FROM rep_sirhu_cab "
        OpenRecordset StrSql, objRs2
        If Not objRs2.EOF And Not EsNulo(objRs2!lote) Then
            lote = CLng(objRs2!lote) + 1
        Else
            lote = 1
        End If
        objRs2.Close
        '------------------------------------------
        
        '------------------------------------------
        'INSERTO LA CABECERA
        '------------------------------------------
        OpenConnection strconexion, objInsert
        StrSql = " INSERT INTO rep_sirhu_cab "
        StrSql = StrSql & "("
        StrSql = StrSql & " rep_bpronro, rep_legDesde, rep_legHasta,"
        StrSql = StrSql & " rep_empEstado, rep_pliqdesde, rep_pliqhasta,"
        StrSql = StrSql & " rep_procesos, rep_tenro1, rep_estrnro1,"
        StrSql = StrSql & " rep_tenro2, rep_estrnro2, rep_tenro3,"
        StrSql = StrSql & " rep_estrnro3, rep_orden, rep_fechaGeneracion, lote"
        StrSql = StrSql & ")"
        StrSql = StrSql & " VALUES "
        StrSql = StrSql & "("
        StrSql = StrSql & NroProceso & "," & legdesde & "," & leghasta & ","
        StrSql = StrSql & estado & "," & pliqdesde & "," & pliqhasta & ","
        StrSql = StrSql & "'" & listaProcesos & "'," & tenro1 & "," & estrnro1 & ","
        StrSql = StrSql & tenro2 & "," & estrnro2 & "," & tenro3 & ","
        StrSql = StrSql & estrnro3 & ",'" & orden & "'," & ConvFecha(Date) & "," & lote
        StrSql = StrSql & ")"
        objInsert.Execute StrSql, , adExecuteNoRecords
        '------------------------------------------
        'HASTA ACA
        '------------------------------------------
        
        'busco la fecha desde y la fecha hasta de los periodos
        StrSql = " SELECT min(pliqdesde) fechadesde, max(pliqhasta) fechahasta "
        StrSql = StrSql & " FROM periodo "
        StrSql = StrSql & " WHERE pliqnro IN (" & pliqdesde & "," & pliqhasta & ") "
        OpenRecordset StrSql, rs
        If Not rs.EOF Then
            FechaDesde = rs!FechaDesde
            FechaHasta = rs!FechaHasta
        Else
            FechaDesde = ""
            FechaHasta = ""
        End If
        rs.Close

        Call datosPersonas(listaEmpleados, pliqdesde, pliqhasta)
        
        Call cabezalHaberes(listaEmpleados, pliqdesde, pliqhasta)
        
        Call detalleHaberes(listaEmpleados, pliqdesde, pliqhasta)
        
        Call conceptoHaberes(listaEmpleados, pliqdesde, pliqhasta)
    Else
        Exit Sub
    End If
   
    If Not HuboErrores Then
        'Actualizo el estado del proceso
        StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
    Else
        StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProceso
    End If
    
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Fin :" & Now
    Flog.Close

    Exit Sub
    
CE:
    HuboErrores = True
    Flog.writeline "************************************************************"
    Flog.writeline " Error: " & Err.Description & Now
    Flog.writeline " Ultimo sql ejecutado: " & StrSql
    Flog.writeline "************************************************************"
End Sub

Public Sub datosPersonas(ByVal empleados As String, ByVal pliqdesde As Integer, ByVal pliqhasta As Integer)


Dim arrEmpleados

Dim objEmpleado As New datosPersonales
Dim rs As New ADODB.Recordset

Dim j As Integer
Dim Ternro As Long

'variables de los empleados
Dim TipoDoc As Integer

Dim empTipoDoc As String

Dim empNroDoc As String
Dim empNombre2 As String
Dim empNombre As String
Dim empApellido As String
Dim empApellido2 As String
Dim empApeNom As String

Dim empFechaNac As String
Dim empSexo As String
Dim empEstCivil As String
Dim empCodInstitucion As String
Dim empFechaIngreso As String
Dim empCodNacionalidad As Integer
Dim empTitulo As String
Dim empCuil As String
Dim empSistPrevisional As String
Dim empCodSistPrevisional As String
Dim empCodObraSocial As String
Dim empNumeroAfiliacion As String
Dim empTipoHorario As Integer
Dim empNivel As String
Dim docpais As Integer


Dim arrConfrep() As confrep

'levanto los datos del confrep
StrSql = "SELECT * FROM confrep "
StrSql = StrSql & " WHERE repnro=498"
StrSql = StrSql & " ORDER BY confnrocol ASC "
OpenRecordset StrSql, rs
If Not rs.EOF Then
    Do While Not rs.EOF
        Select Case rs!confnrocol
            Case 1: 'codigo de institucion
                If UCase(rs!conftipo) = "TE" Then
                    ReDim Preserve arrConfrep(rs!confnrocol)
                    arrConfrep(rs!confnrocol).conftipo = rs!conftipo
                    arrConfrep(rs!confnrocol).confval = rs!confval
                    arrConfrep(rs!confnrocol).confval2 = IIf(EsNulo(rs!confval2), "estrcodext", rs!confval2)
                Else
                    Flog.writeline "La columna 1 debe ser de tipo TE."
                End If
            
            Case 2: 'Obra social
                If UCase(rs!conftipo) = "TE" Then
                    ReDim Preserve arrConfrep(rs!confnrocol)
                    arrConfrep(rs!confnrocol).conftipo = rs!conftipo
                    arrConfrep(rs!confnrocol).confval = rs!confval
                    arrConfrep(rs!confnrocol).confval2 = IIf(EsNulo(rs!confval2), "estrcodext", rs!confval2)
                Else
                    Flog.writeline "La columna 2 debe ser de tipo TE."
                End If
        End Select
    rs.MoveNext
    Loop
Else
    Flog.writeline "No esta configurado el confrep del reporte 498. Se aborta el procesamiento"
    Exit Sub
End If
rs.Close
'hasta aca


arrEmpleados = Split(empleados, ",")

IncPorc = 25 / UBound(arrEmpleados)
Progreso = 0

For j = 1 To UBound(arrEmpleados)
    Ternro = arrEmpleados(j) 'se guarda el ternro de c/u
    Progreso = Progreso + IncPorc
    'Busco los datos personales como nombre, apellido, etc..
    objEmpleado.buscarDatosPersonales Ternro
    Flog.writeline "PROCESANDO LEGAJO :" & objEmpleado.obtenerLegajo
        
    'BUSCO LOS DATOS DEL DOCUMENTO DEL EMPLEADO
    
    docpais = objEmpleado.obtenerDocPais
    Flog.writeline "Doc Pais: " & docpais
    
    objEmpleado.buscarNroDoc Ternro, 0, docpais
    
    TipoDoc = objEmpleado.obtenerTipoDoc
    Flog.writeline "Tipo de documento :" & TipoDoc
    
    'busco el mapeo del dato
    empTipoDoc = mapearDato("sirhu_tipdoc", TipoDoc)
    Flog.writeline "Tipo de documento mapeado:" & empTipoDoc
    
    empNroDoc = objEmpleado.obtenerNroDoc
    If EsNulo(empNroDoc) Then
        empNroDoc = 0
    End If
    Flog.writeline "Nro de documento :" & empNroDoc
   
    empNombre2 = objEmpleado.obtenerNombreApellido("nombre2")
    Flog.writeline "Segundo Nombre: " & empNombre2
    
    empNombre = objEmpleado.obtenerNombreApellido("nombre")
    Flog.writeline "Nombre: " & empNombre
    
    empApellido2 = objEmpleado.obtenerNombreApellido("apellido2")
    Flog.writeline "Segundo Apellido: " & empApellido2

    empApellido = objEmpleado.obtenerNombreApellido("apellido")
    Flog.writeline "Apellido: " & empApellido
    
    'armo el nombre y apellido
    empApeNom = empApellido & " " & empApellido2 & "," & empNombre & " " & empNombre2
    
    empFechaNac = objEmpleado.obtenerFNacimiento
    Flog.writeline "Fecha de nacimiento: " & empFechaNac
    
    empSexo = objEmpleado.obtenerSexo
    Flog.writeline "Sexo: " & empSexo
    empSexo = mapearDato("sirhu_tersex", empSexo)
    Flog.writeline "Sexo con mapeo: " & empSexo
    
    empEstCivil = objEmpleado.obtenerEstadoCivil
    Flog.writeline "Estado Civil: " & empEstCivil
    empEstCivil = mapearDato("sirhu_estcivil", empEstCivil)
    Flog.writeline "Estado Civil Mapeado: " & empEstCivil
    
    objEmpleado.buscarEstructuras2Fechas Ternro, arrConfrep(1).confval, FechaDesde, FechaHasta, arrConfrep(1).confval2, False
    empCodInstitucion = objEmpleado.obtenerEstructura2Fechas
    Flog.writeline "Codigo Institucion: " & empCodInstitucion
    
    objEmpleado.buscarEstructuras2Fechas Ternro, arrConfrep(1).confval, FechaDesde, FechaHasta, "estrdabr", False
    empCodInstitucion = empCodInstitucion & "@" & objEmpleado.obtenerEstructura2Fechas
    Flog.writeline "Codigo Institucion: " & empCodInstitucion
    
    
    empFechaIngreso = objEmpleado.obtenerFechaIngreso
    Flog.writeline "Fecha de ingreso: " & empFechaIngreso
    
    empCodNacionalidad = objEmpleado.obtenerNacionalidad
    Flog.writeline "Cod de nacionalidad: " & empCodNacionalidad
    empCodNacionalidad = mapearDato("sirhu_nac", empCodNacionalidad)
    Flog.writeline "Cod de nacionalidad mapeado: " & empCodNacionalidad
    
    
    objEmpleado.buscarTitulo Ternro
    empTitulo = objEmpleado.obtenerTitulo
    Flog.writeline "Titulo: " & empTitulo
    
    empNivel = objEmpleado.obtenerCodNivel
    Flog.writeline "Nivel: " & empNivel
    empNivel = mapearDato("sirhu_nivest", empNivel)
    Flog.writeline "Nivel Mapeado: " & empNivel
    
    empCuil = objEmpleado.obtenerCUIL
    Flog.writeline "CUIL: " & empCuil
    
    empSistPrevisional = "C"
    Flog.writeline "Sistema Previsional: " & empSistPrevisional
    
    empCodSistPrevisional = "90"
    Flog.writeline "Codigo Sistema Previsional: " & empCodSistPrevisional
    
    objEmpleado.buscarEstructuras2Fechas Ternro, arrConfrep(2).confval, FechaDesde, FechaHasta, arrConfrep(2).confval2, False
    empCodObraSocial = objEmpleado.obtenerEstructura2Fechas
    Flog.writeline "Codigo Obra social: " & empCodObraSocial
    
    empNumeroAfiliacion = ""
    Flog.writeline "Numero de afiliacion: " & empNumeroAfiliacion
    
    empTipoHorario = 1
    Flog.writeline "Tipo Horario: " & empTipoHorario
    
    
    'INSERTO LOS DATOS DEL EMPLEADO
    StrSql = "INSERT INTO rep_sirhu_datosPersonales "
    StrSql = StrSql & "("
    StrSql = StrSql & " rep_bpronro, rep_tipoDoc, rep_nroDoc, "
    StrSql = StrSql & " rep_apellido_nombre, rep_fecha_nac, rep_sexo, "
    StrSql = StrSql & " rep_estCivil, rep_codInst, rep_fechaIngreso, "
    StrSql = StrSql & " rep_codNac, rep_codEduc, rep_descTitulo, "
    StrSql = StrSql & " rep_cuil, rep_sistPrev, rep_codSistPrev, "
    StrSql = StrSql & " rep_codOSocial, rep_nroAfil, rep_tipoHorario "
    StrSql = StrSql & ")"
    StrSql = StrSql & " VALUES "
    StrSql = StrSql & "("
    StrSql = StrSql & NroProceso & ",'" & Left(empTipoDoc, 3) & "','" & empNroDoc & "',"
    StrSql = StrSql & "'" & Left(empApeNom, 60) & "',"
    If EsNulo(empFechaNac) Then
        StrSql = StrSql & " NULL,'" & Left(empSexo, 4) & "',"
    Else
        StrSql = StrSql & ConvFecha(empFechaNac) & ",'" & Left(empSexo, 4) & "',"
    End If
    StrSql = StrSql & "'" & Left(empEstCivil, 4) & "','" & Left(empCodInstitucion, 200) & "',"
    If EsNulo(empFechaIngreso) Then
        StrSql = StrSql & "NULL,"
    Else
        StrSql = StrSql & ConvFecha(empFechaIngreso) & ","
    End If
    StrSql = StrSql & empCodNacionalidad & ",'" & Left(empNivel, 2) & "','" & Left(empTitulo, 30) & "','" & Left(empCuil, 11) & "',"
    StrSql = StrSql & "'" & empSistPrevisional & "','" & Left(empCodSistPrevisional, 4) & "','" & Left(empCodObraSocial, 10) & "', "
    StrSql = StrSql & "'" & empNumeroAfiliacion & "'," & Left(empTipoHorario, 2)
    StrSql = StrSql & ")"
    Flog.writeline "Insert: " & StrSql
    objInsert.Execute StrSql, , adExecuteNoRecords
    Flog.writeline " SE INSERTO EL EMPLEADO: " & objEmpleado.obtenerLegajo
    'HASTA ACA
    
    'ACTUALIZO EL PROGRESO
    actualizarProgreso (Progreso)
    Flog.writeline "Porcentaje: " & Progreso
    'HASTA ACA
Next j

End Sub
Public Function actualizarProgreso(porc)
    
On Error GoTo error
    StrSql = "UPDATE batch_proceso SET  bprcprogreso =" & CInt(porc) & ", bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesando' WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords

Exit Function

error:
    Flog.writeline "Se produjo un error al actualizar el progreso"
    Flog.writeline Err.Description
End Function
Public Function mapearDato(ByVal infotipo As String, ByVal codInterno As String) As String

Dim rs As New ADODB.Recordset
Dim aux
StrSql = "SELECT codexterno FROM mapeo_sap"
StrSql = StrSql & " WHERE UPPER(infotipo)='" & UCase(infotipo) & "'"
StrSql = StrSql & " AND codinterno='" & codInterno & "'"
OpenRecordset StrSql, rs
If Not rs.EOF Then
    aux = IIf(EsNulo(rs!codexterno), codInterno, rs!codexterno)
Else
    aux = codInterno
End If
rs.Close

mapearDato = aux

End Function

Public Sub cabezalHaberes(ByVal empleados As String, ByVal pliqdesde As Integer, ByVal pliqhasta As Integer)

'On Error GoTo siguiente
Dim arrEmpleados

Dim objEmpleado As New datosPersonales
Dim rs As New ADODB.Recordset

Dim j As Integer
Dim Ternro As Long

Dim arrConfrep() As confrep

'variables del empleado
Dim empInstitucion As String
Dim TipoDoc As Integer
Dim empTipoDoc As String
Dim empNroDoc As String
Dim empEscalafon As String
Dim empCodAgrupamiento As String
Dim empCodNivel As String
Dim empCodGrado As String
Dim empCodUnidad  As String
Dim empCodNudo As String
Dim empCodJur As String
Dim empCodSubJur As String
Dim empCodEntidad As String
Dim empCodProg As String
Dim empCodSubProg As String
Dim empCodProy As String
Dim empCodAct As String
Dim empCodUbigeo As String
Dim empPeriodo As String
Dim empTipoPlanta As String
Dim empFechaIngreso As String
Dim empCodFinanciamiento As Integer
Dim empMarcaEstado As String

Dim docpais As Integer

'levanto los datos del confrep
StrSql = "SELECT * FROM confrep "
StrSql = StrSql & " WHERE repnro=498"
StrSql = StrSql & " ORDER BY confnrocol ASC "
OpenRecordset StrSql, rs
If Not rs.EOF Then
    Do While Not rs.EOF
        Select Case rs!confnrocol
            Case 1: 'codigo de institucion
                If UCase(rs!conftipo) = "TE" Then
                    ReDim Preserve arrConfrep(rs!confnrocol)
                    arrConfrep(rs!confnrocol).conftipo = rs!conftipo
                    arrConfrep(rs!confnrocol).confval = rs!confval
                    arrConfrep(rs!confnrocol).confval2 = IIf(EsNulo(rs!confval2), "estrcodext", rs!confval2)
                Else
                    Flog.writeline "La columna 1 debe ser de tipo TE. Codigo de Institucion"
                End If
            Case 3: 'Cod. de escalafon
                If UCase(rs!conftipo) = "TE" Then
                    ReDim Preserve arrConfrep(rs!confnrocol)
                    arrConfrep(rs!confnrocol).conftipo = rs!conftipo
                    arrConfrep(rs!confnrocol).confval = rs!confval
                    arrConfrep(rs!confnrocol).confval2 = IIf(EsNulo(rs!confval2), "estrcodext", rs!confval2)
                Else
                    Flog.writeline "La columna 3 debe ser de tipo TE. Codigo de escalafon"
                End If
            Case 4: 'Cod. nivel
                If UCase(rs!conftipo) = "TE" Then
                    ReDim Preserve arrConfrep(rs!confnrocol)
                    arrConfrep(rs!confnrocol).conftipo = rs!conftipo
                    arrConfrep(rs!confnrocol).confval = rs!confval
                    arrConfrep(rs!confnrocol).confval2 = IIf(EsNulo(rs!confval2), "estrcodext", rs!confval2)
                Else
                    Flog.writeline "La columna 4 debe ser de tipo TE. Codigo nivel"
                End If
            Case 5: 'Cod. Grado o categoria
                If UCase(rs!conftipo) = "TE" Then
                    ReDim Preserve arrConfrep(rs!confnrocol)
                    arrConfrep(rs!confnrocol).conftipo = rs!conftipo
                    arrConfrep(rs!confnrocol).confval = rs!confval
                    arrConfrep(rs!confnrocol).confval2 = IIf(EsNulo(rs!confval2), "estrcodext", rs!confval2)
                Else
                    Flog.writeline "La columna 5 debe ser de tipo TE. Codigo Grado o Categoria"
                End If
            Case 6: 'NUDO
                If UCase(rs!conftipo) = "TE" Then
                    ReDim Preserve arrConfrep(rs!confnrocol)
                    arrConfrep(rs!confnrocol).conftipo = rs!conftipo
                    arrConfrep(rs!confnrocol).confval = rs!confval
                    arrConfrep(rs!confnrocol).confval2 = IIf(EsNulo(rs!confval2), "estrcodext", rs!confval2)
                Else
                    Flog.writeline "La columna 6 debe ser de tipo TE. NUDO"
                End If
            Case 7: 'Cod. de ubicacion geografica
                If UCase(rs!conftipo) = "TE" Then
                    ReDim Preserve arrConfrep(rs!confnrocol)
                    arrConfrep(rs!confnrocol).conftipo = rs!conftipo
                    arrConfrep(rs!confnrocol).confval = rs!confval
                    arrConfrep(rs!confnrocol).confval2 = IIf(EsNulo(rs!confval2), "estrcodext", rs!confval2)
                Else
                    Flog.writeline "La columna 6 debe ser de tipo TE. Codigo de ubicacion geografica"
                End If
        
        End Select
    rs.MoveNext
    Loop
Else
    Flog.writeline "No esta configurado el confrep del reporte 498. Se aborta el procesamiento"
    Exit Sub
End If
rs.Close
'hasta aca

arrEmpleados = Split(empleados, ",")
IncPorc = 25 / UBound(arrEmpleados)
'Progreso = 0

For j = 1 To UBound(arrEmpleados)
    Progreso = Progreso + IncPorc
    Ternro = arrEmpleados(j) 'se guarda el ternro de c/u
    
    'Busco los datos personales como nombre, apellido, etc..
    objEmpleado.buscarDatosPersonales Ternro
    Flog.writeline "PROCESANDO LEGAJO :" & objEmpleado.obtenerLegajo
    
    docpais = objEmpleado.obtenerDocPais
    Flog.writeline "Doc Pais: " & docpais
    
    'busco el cod de la institucion
    objEmpleado.buscarEstructuras2Fechas Ternro, arrConfrep(1).confval, FechaDesde, FechaHasta, arrConfrep(1).confval2, False
    empInstitucion = objEmpleado.obtenerEstructura2Fechas
    
    objEmpleado.buscarEstructuras2Fechas Ternro, arrConfrep(1).confval, FechaDesde, FechaHasta, "estrdabr", False
    empInstitucion = empInstitucion & "@" & objEmpleado.obtenerEstructura2Fechas
    Flog.writeline "Codigo Institucion: " & empInstitucion
    
    'BUSCO LOS DATOS DEL DOCUMENTO DEL EMPLEADO
    objEmpleado.buscarNroDoc Ternro, 0, docpais
    
    TipoDoc = objEmpleado.obtenerTipoDoc
    Flog.writeline "Tipo de documento :" & TipoDoc
    'busco el mapeo del dato
    empTipoDoc = mapearDato("sirhu_tipdoc", TipoDoc)
    Flog.writeline "Tipo de documento mapeado:" & empTipoDoc
    
    empNroDoc = objEmpleado.obtenerNroDoc
    If EsNulo(empNroDoc) Then
        empNroDoc = 0
    End If
    Flog.writeline "Nro de documento :" & empNroDoc
    
    'codigo de escalafon
    objEmpleado.buscarEstructuras2Fechas Ternro, arrConfrep(3).confval, FechaDesde, FechaHasta, arrConfrep(3).confval2, False
    empEscalafon = objEmpleado.obtenerEstructura2Fechas
    Flog.writeline "Cod. Escalafon :" & empEscalafon
    
    'Codigo de agrupamiento
    empCodAgrupamiento = "G"
    Flog.writeline "Cod. de agrupamiento fijo: " & empCodAgrupamiento
    
    'Codigo de nivel
    objEmpleado.buscarEstructuras2Fechas Ternro, arrConfrep(4).confval, FechaDesde, FechaHasta, arrConfrep(4).confval2, False
    empCodNivel = objEmpleado.obtenerEstructura2Fechas
    Flog.writeline "Cod. Nivel :" & empCodNivel
    
    'Codigo de grado o categoria
    objEmpleado.buscarEstructuras2Fechas Ternro, arrConfrep(5).confval, FechaDesde, FechaHasta, arrConfrep(5).confval2, False
    empCodGrado = objEmpleado.obtenerEstructura2Fechas
    Flog.writeline "Cod. grado o categoria :" & empCodGrado
        
    'Codigo de unidad
    empCodUnidad = "56000079201"
    Flog.writeline "Cod. de unidad fijo: " & empCodUnidad
    
    'Codigo de NUDO
    objEmpleado.buscarEstructuras2Fechas Ternro, arrConfrep(6).confval, FechaDesde, FechaHasta, arrConfrep(6).confval2, False
    empCodNudo = objEmpleado.obtenerEstructura2Fechas
    Flog.writeline "Cod. Nudo :" & empCodNudo
    
    'Codigo Jurisdiccion
    empCodJur = "056"
    Flog.writeline "Cod. jurisdiccion fijo: " & empCodJur
    
    'Codigo SubJurisdiccion
    empCodSubJur = "00"
    Flog.writeline "Cod. Subjurisdiccion fijo: " & empCodSubJur
    
    'Codigo Entidad
    empCodEntidad = "792"
    Flog.writeline "Cod. Entidad fijo: " & empCodEntidad
    
    empCodProg = ""
    Flog.writeline "Codigo de programa fijo vacio: " & empCodProg
    
    empCodSubProg = ""
    Flog.writeline "Codigo de subprograma fijo vacio: " & empCodSubProg
    
    empCodProy = ""
    Flog.writeline "Codigo de proyecto fijo vacio: " & empCodProy
    
    empCodAct = ""
    Flog.writeline "Codigo de actividad fijo vacio: " & empCodAct
    
    'Codigo de Ubicacion Geografica
    empCodUbigeo = ""
    
    'Periodo correspondiente a la informacion  aaaamm año y mes de liq exportado
    empPeriodo = FechaDesde
    Flog.writeline "Periodo de liquidacion exportado: " & empPeriodo
    
    'Tipo de Planta
    empTipoPlanta = "P"
    Flog.writeline "Tipo de Planta: " & empTipoPlanta
    
    'Fecha de ingreso
    empFechaIngreso = objEmpleado.obtenerFechaIngreso
    Flog.writeline "Fecha de ingreso: " & empFechaIngreso
        
    'Codigo de financiamiento
    empCodFinanciamiento = 0
    Flog.writeline "Codigo de financiamiento fijo vacio"
    
    'Marca de estado
    empMarcaEstado = ""
    Flog.writeline "Marca de estado fijo vacio"
    
    StrSql = "INSERT INTO rep_sirhu_cabezalHaberes "
    StrSql = StrSql & "("
    StrSql = StrSql & " rep_bpronro, rep_tipoDoc, rep_nroDoc, "
    StrSql = StrSql & " rep_codInst, rep_codEscalafon, rep_codAgrupamiento, "
    StrSql = StrSql & " rep_codNivel, rep_codGrado, rep_codUnidad, "
    StrSql = StrSql & " rep_codNudo, rep_codJur, rep_codSubJur, "
    StrSql = StrSql & " rep_codEntidad, rep_codProg, rep_codSubProg, "
    StrSql = StrSql & " rep_codProy, rep_codAct, rep_codUbigeo, "
    StrSql = StrSql & " rep_periodo, rep_tipoPlanta, rep_fechaIngr, "
    StrSql = StrSql & " rep_codFinan, rep_estado"
    StrSql = StrSql & ")"
    StrSql = StrSql & " VALUES "
    StrSql = StrSql & "("
    StrSql = StrSql & NroProceso & ",'" & empTipoDoc & "','" & empNroDoc & "',"
    StrSql = StrSql & "'" & Left(empInstitucion, 200) & "','" & Left(empEscalafon, 4) & "','" & Left(empCodAgrupamiento, 4) & "',"
    StrSql = StrSql & "'" & Left(empCodNivel, 4) & "','" & Left(empCodGrado, 4) & "','" & Left(empCodUnidad, 11) & "',"
    StrSql = StrSql & "'" & Left(empCodNudo, 3) & "','" & Left(empCodJur, 2) & "','" & Left(empCodSubJur, 2) & "',"
    StrSql = StrSql & "'" & Left(empCodEntidad, 3) & "','" & Left(empCodProg, 2) & "','" & Left(empCodSubProg, 2) & "',"
    StrSql = StrSql & "'" & Left(empCodProy, 3) & "','" & Left(empCodAct, 3) & "','" & Left(empCodUbigeo, 3) & "',"
    If EsNulo(empPeriodo) Then
        StrSql = StrSql & "NULL,"
    Else
        StrSql = StrSql & ConvFecha(empPeriodo) & ","
    End If
    
    StrSql = StrSql & "'" & Left(empTipoPlanta, 1) & "',"
    
    If EsNulo(empFechaIngreso) Then
        StrSql = StrSql & "NULL,"
    Else
        StrSql = StrSql & ConvFecha(empFechaIngreso) & ","
    End If
    StrSql = StrSql & empCodFinanciamiento & ",'" & Left(empMarcaEstado, 1) & "'"
    StrSql = StrSql & ")"
    
    If guardarDatos(StrSql) Then
        Flog.writeline " SE INSERTO EL EMPLEADO: " & objEmpleado.obtenerLegajo
    Else
        Flog.writeline " SE PRODUJO UN ERROR AL INSERTAR EL EMPLEADO: " & objEmpleado.obtenerLegajo
    End If
    
    'ACTUALIZO EL PROGRESO
    actualizarProgreso (Progreso)
    Flog.writeline "Porcentaje: " & Progreso
    'HASTA ACA
    
Next j
        
        
End Sub

Public Sub detalleHaberes(ByVal empleados As String, ByVal pliqdesde As Integer, ByVal pliqhasta As Integer)
Dim arrEmpleados

Dim objEmpleado As New datosPersonales
Dim rs As New ADODB.Recordset

Dim j As Integer
Dim Ternro As Long
Dim I As Integer

Dim arrConfrep() As confrep
arrEmpleados = Split(empleados, ",")

Dim empTipoDoc As String
Dim empNroDoc As String
Dim empInstitucion As String
Dim TipoDoc As Integer
Dim empEscalafon As String
Dim entro As Boolean

Dim docpais As Integer

'levanto los datos del confrep
StrSql = "SELECT * FROM confrep "
StrSql = StrSql & " WHERE repnro=498"
StrSql = StrSql & " ORDER BY confnrocol ASC "
OpenRecordset StrSql, rs
If Not rs.EOF Then
    Do While Not rs.EOF
        Select Case rs!confnrocol
            Case 1: 'codigo de institucion
                If UCase(rs!conftipo) = "TE" Then
                    ReDim Preserve arrConfrep(rs!confnrocol)
                    arrConfrep(rs!confnrocol).conftipo = rs!conftipo
                    arrConfrep(rs!confnrocol).confval = rs!confval
                    arrConfrep(rs!confnrocol).confval2 = IIf(EsNulo(rs!confval2), "estrcodext", rs!confval2)
                Else
                    Flog.writeline "La columna 1 debe ser de tipo TE."
                End If
                
            Case 3: 'Cod. de escalafon
                If UCase(rs!conftipo) = "TE" Then
                    ReDim Preserve arrConfrep(rs!confnrocol)
                    arrConfrep(rs!confnrocol).conftipo = rs!conftipo
                    arrConfrep(rs!confnrocol).confval = rs!confval
                    arrConfrep(rs!confnrocol).confval2 = IIf(EsNulo(rs!confval2), "estrcodext", rs!confval2)
                Else
                    Flog.writeline "La columna 3 debe ser de tipo TE. Codigo de escalafon"
                End If
                
            Case 8: 'Tipo de Concepto Imprimible Puente
                    ReDim Preserve arrConfrep(rs!confnrocol)
                    arrConfrep(rs!confnrocol).conftipo = rs!conftipo
                    arrConfrep(rs!confnrocol).confval = rs!confval
                    arrConfrep(rs!confnrocol).confval2 = IIf(EsNulo(rs!confval2), "estrcodext", rs!confval2)
            
            Case 9: 'Tipo de Concepto Imprimible No Puente
                    ReDim Preserve arrConfrep(rs!confnrocol)
                    arrConfrep(rs!confnrocol).conftipo = rs!conftipo
                    arrConfrep(rs!confnrocol).confval = rs!confval
                    arrConfrep(rs!confnrocol).confval2 = IIf(EsNulo(rs!confval2), "estrcodext", rs!confval2)
            
            Case 10: 'Tipo de Concepto No Impr. Puente
                    ReDim Preserve arrConfrep(rs!confnrocol)
                    arrConfrep(rs!confnrocol).conftipo = rs!conftipo
                    arrConfrep(rs!confnrocol).confval = rs!confval
                    arrConfrep(rs!confnrocol).confval2 = IIf(EsNulo(rs!confval2), "estrcodext", rs!confval2)
                        
            Case 11: 'Tipo de Concepto No Impr. No Puente
                    ReDim Preserve arrConfrep(rs!confnrocol)
                    arrConfrep(rs!confnrocol).conftipo = rs!conftipo
                    arrConfrep(rs!confnrocol).confval = rs!confval
                    arrConfrep(rs!confnrocol).confval2 = IIf(EsNulo(rs!confval2), "estrcodext", rs!confval2)
            
        End Select
    rs.MoveNext
    Loop
Else
    Flog.writeline "No esta configurado el confrep del reporte 498. Se aborta el procesamiento"
    Exit Sub
End If
rs.Close
'hasta aca

IncPorc = 25 / UBound(arrEmpleados)
'Progreso = 0

For j = 1 To UBound(arrEmpleados)
    Progreso = Progreso + IncPorc
    Ternro = arrEmpleados(j) 'se guarda el ternro de c/u
    
    'Busco los datos personales como nombre, apellido, etc..
    objEmpleado.buscarDatosPersonales Ternro
    Flog.writeline "PROCESANDO LEGAJO :" & objEmpleado.obtenerLegajo
    
    docpais = objEmpleado.obtenerDocPais
    Flog.writeline "Doc Pais: " & docpais

    'busco el cod de la institucion
    objEmpleado.buscarEstructuras2Fechas Ternro, arrConfrep(1).confval, FechaDesde, FechaHasta, arrConfrep(1).confval2, False
    empInstitucion = objEmpleado.obtenerEstructura2Fechas
    Flog.writeline "INSTITUCION : " & empInstitucion
    
    objEmpleado.buscarEstructuras2Fechas Ternro, arrConfrep(1).confval, FechaDesde, FechaHasta, "estrdabr", False
    empInstitucion = empInstitucion & "@" & objEmpleado.obtenerEstructura2Fechas
    Flog.writeline "Codigo Institucion: " & empInstitucion
    
    'BUSCO LOS DATOS DEL DOCUMENTO DEL EMPLEADO
    objEmpleado.buscarNroDoc Ternro, 0, docpais
    
    TipoDoc = objEmpleado.obtenerTipoDoc
    Flog.writeline "Tipo de documento :" & TipoDoc
    'busco el mapeo del dato
    empTipoDoc = mapearDato("sirhu_tipdoc", TipoDoc)
    Flog.writeline "Tipo de documento mapeado:" & empTipoDoc
    
    empNroDoc = objEmpleado.obtenerNroDoc
    If EsNulo(empNroDoc) Then
        empNroDoc = 0
    End If
    Flog.writeline "Nro de documento :" & empNroDoc
    
    'codigo de escalafon
    objEmpleado.buscarEstructuras2Fechas Ternro, arrConfrep(3).confval, FechaDesde, FechaHasta, arrConfrep(3).confval2, False
    empEscalafon = objEmpleado.obtenerEstructura2Fechas
    Flog.writeline "Cod. Escalafon :" & empEscalafon
    
    'VOY A BUSCAR LOS CONCEPTOS DE LA LIQUIDACION
    StrSql = "SELECT concepto.conccod, concepto.concnro, concepto.concabr, SUM(dlimonto) monto, concepto.concpuente, concepto.concimp, concepto.tconnro "
    StrSql = StrSql & " FROM cabliq c "
    StrSql = StrSql & " INNER JOIN detliq d ON d.cliqnro = c.cliqnro"
    StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = d.concnro "
    StrSql = StrSql & " WHERE c.pronro IN (" & listaProcesos & ")"
    StrSql = StrSql & " AND c.empleado = " & Ternro
    StrSql = StrSql & " GROUP BY concepto.conccod, concepto.concnro, concepto.concabr, concepto.concpuente, concepto.concimp, concepto.tconnro "
    StrSql = StrSql & " ORDER BY concepto.conccod asc "
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        Do While Not rs.EOF
            entro = False
            If (rs!concimp = -1) And (rs!concpuente = -1) Then 'miro si es imprimible y puente
                If (InStr(arrConfrep(8).confval2, "," & rs!tconnro & ",") > 0) Then
                    entro = True
                    'StrSqlAux = StrSqlAux & rs!ConcNro & "," & FormatNumber(rs!Monto, 2)
                End If
            Else
                If (rs!concimp = -1) And (rs!concpuente = 0) Then 'Imprimible y No Puente
                    If (InStr(arrConfrep(9).confval2, "," & rs!tconnro & ",") > 0) Then
                        'StrSqlAux = StrSqlAux & rs!ConcNro & "," & FormatNumber(rs!Monto, 2)
                        entro = True
                    End If
                Else
                    If (rs!concimp = 0) And (rs!concpuente = -1) Then 'No Impr. Puente
                        If (InStr(arrConfrep(10).confval2, "," & rs!tconnro & ",") > 0) Then
                            'StrSqlAux = StrSqlAux & rs!ConcNro & "," & FormatNumber(rs!Monto, 2)
                            entro = True
                        End If
                    Else
                        If (rs!concimp = 0) And (rs!concpuente = 0) Then 'No Impr. No Puente
                            If (InStr(arrConfrep(11).confval2, "," & rs!tconnro & ",") > 0) Then
                                entro = True
                            '    StrSqlAux = StrSqlAux & rs!ConcNro & "," & FormatNumber(rs!Monto, 2)
                            End If
                        End If
                    End If
                End If
            End If
            
            If entro Then
                I = I + 1
                ReDim Preserve arrConcepto(I)
                arrConcepto(I).Ternro = Ternro
                arrConcepto(I).descabr = rs!concabr
                arrConcepto(I).ConcCod = rs!ConcCod
                arrConcepto(I).ConcNro = rs!ConcNro
                arrConcepto(I).tipoconc = rs!tconnro
                
                StrSql = " INSERT INTO rep_sirhu_detalleHaberes (rep_bpronro, rep_tipoDoc, rep_nroDoc, rep_codInst,"
                StrSql = StrSql & " rep_codEscalafon, rep_codConcepto, rep_importeConcepto, rep_unidadFisica,"
                StrSql = StrSql & " rep_cantUniFisica, rep_periodoExportacion ) "
                StrSql = StrSql & " VALUES "
                StrSql = StrSql & " ( "
                StrSql = StrSql & NroProceso & ",'" & empTipoDoc & "','" & empNroDoc & "','" & Left(empInstitucion, 200) & "',"
                StrSql = StrSql & "'" & Left(empEscalafon, 4) & "','" & rs!ConcCod & "'," & Replace(FormatNumber(rs!Monto, 2), ",", "") & ",99,1," & ConvFecha(FechaDesde)
                StrSql = StrSql & " ) "
                Flog.writeline "MONTO: " & rs!Monto
                'Flog.writeline "MONTO CON FORMAT NUMBER: " & FormatNumber(rs!Monto, 2)
                guardarDatos (StrSql)
            End If
        rs.MoveNext
        Loop
    End If
    rs.Close
    'HASTA ACA
    
    'ACTUALIZO EL PROGRESO
    actualizarProgreso (Progreso)
    Flog.writeline "Porcentaje: " & Progreso
    'HASTA ACA
Next j

End Sub

Public Sub conceptoHaberes(ByVal empleados As String, ByVal pliqdesde As Integer, ByVal pliqhasta As Integer)

Dim empInstitucion As String
Dim empEscalafon As String
Dim codConcepto As String
Dim descConcepto As String
Dim caracRem As Integer
Dim tipoconc As Integer

Dim j As Integer

Dim arrConfrep() As confrep

Dim objEmpleado As New datosPersonales

'levanto los datos del confrep
StrSql = "SELECT * FROM confrep "
StrSql = StrSql & " WHERE repnro=498"
StrSql = StrSql & " ORDER BY confnrocol ASC "
OpenRecordset StrSql, rs
If Not rs.EOF Then
    Do While Not rs.EOF
        Select Case rs!confnrocol
            Case 1: 'codigo de institucion
                If UCase(rs!conftipo) = "TE" Then
                    ReDim Preserve arrConfrep(rs!confnrocol)
                    arrConfrep(rs!confnrocol).conftipo = rs!conftipo
                    arrConfrep(rs!confnrocol).confval = rs!confval
                    arrConfrep(rs!confnrocol).confval2 = IIf(EsNulo(rs!confval2), "estrcodext", rs!confval2)
                Else
                    Flog.writeline "La columna 1 debe ser de tipo TE."
                End If
                
            Case 3: 'Cod. de escalafon
                If UCase(rs!conftipo) = "TE" Then
                    ReDim Preserve arrConfrep(rs!confnrocol)
                    arrConfrep(rs!confnrocol).conftipo = rs!conftipo
                    arrConfrep(rs!confnrocol).confval = rs!confval
                    arrConfrep(rs!confnrocol).confval2 = IIf(EsNulo(rs!confval2), "estrcodext", rs!confval2)
                Else
                    Flog.writeline "La columna 3 debe ser de tipo TE. Codigo de escalafon"
                End If
        End Select
    rs.MoveNext
    Loop
End If

IncPorc = 25 / UBound(arrConcepto)
'Progreso = 0
'recorro el arreglo de los conceptos
For j = 1 To UBound(arrConcepto)
    Progreso = Progreso + IncPorc
    'busco el cod de la institucion
    objEmpleado.buscarEstructuras2Fechas arrConcepto(j).Ternro, arrConfrep(1).confval, FechaDesde, FechaHasta, arrConfrep(1).confval2, False
    empInstitucion = objEmpleado.obtenerEstructura2Fechas
    
    objEmpleado.buscarEstructuras2Fechas arrConcepto(j).Ternro, arrConfrep(1).confval, FechaDesde, FechaHasta, "estrdabr", False
    empInstitucion = empInstitucion & "@" & objEmpleado.obtenerEstructura2Fechas
    Flog.writeline "Codigo Institucion: " & empInstitucion
    
    'codigo de escalafon
    objEmpleado.buscarEstructuras2Fechas arrConcepto(j).Ternro, arrConfrep(3).confval, FechaDesde, FechaHasta, arrConfrep(3).confval2, False
    empEscalafon = objEmpleado.obtenerEstructura2Fechas
    Flog.writeline "Cod. Escalafon :" & empEscalafon
    
    'codConcepto
    codConcepto = arrConcepto(j).ConcCod
    
    'descripcion del concepto
    descConcepto = arrConcepto(j).descabr
    
    'caracter remunerativo
    caracRem = mapearDato("sirhu_caracRem", arrConcepto(j).tipoconc)
    
    'Tipo de concepto
    tipoconc = mapearDato("sirhu_tipoConc", arrConcepto(j).tipoconc)
    
    StrSql = " INSERT INTO rep_sirhu_conceptoHaberes "
    StrSql = StrSql & "("
    StrSql = StrSql & "rep_bpronro, rep_codInst, rep_codEscalafon,"
    StrSql = StrSql & "rep_codConcepto, rep_descConcepto, rep_caracRem, rep_tipoConc"
    StrSql = StrSql & ")"
    StrSql = StrSql & " VALUES "
    StrSql = StrSql & "("
    StrSql = StrSql & NroProceso & ",'" & Left(empInstitucion, 200) & "','" & Left(empEscalafon, 4) & "',"
    StrSql = StrSql & "'" & codConcepto & "','" & Left(descConcepto, 40) & "'," & caracRem & "," & tipoconc
    StrSql = StrSql & ")"
    guardarDatos (StrSql)
    
    'ACTUALIZO EL PROGRESO
    actualizarProgreso (Progreso)
    Flog.writeline "Porcentaje: " & Progreso
    'HASTA ACA
Next

End Sub
Public Function guardarDatos(ByVal query As String) As Boolean

On Error GoTo error:

'ejecuto el insert
objInsert.Execute query, , adExecuteNoRecords
guardarDatos = True

Exit Function

error:
    Flog.writeline "Se produjo un error al guardar los datos, query: " & query
    guardarDatos = False

End Function
