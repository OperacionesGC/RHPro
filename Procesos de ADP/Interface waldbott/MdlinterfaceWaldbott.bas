Attribute VB_Name = "MdlinterfaceWaldbott"
Option Explicit

'Const Version = "1.00"
'Const FechaVersion = "04/09/2012" ' cas 16698 - Deluchi Ezequiel - Version Inicial

'Const Version = "1.01"
'Const FechaVersion = "03/12/2012" ' cas 16698 agregados - Deluchi Ezequiel - se cambio la forma de obtener e informar los siguiente campos
' codigo de situacion de revista, contrato, jurisdiccion, grupo, generasijp, centro de costo-imputacion

'Const Version = "1.02"
'Const FechaVersion = "10/12/2012" ' CAS-16698 - Samconsult - tablas valores-horas - Deluchi Ezequiel
' Se inserta en caso de alta de empleado, el nuevo legajo en las tablas horas y valores, de samconsult

'Const Version = "1.03"
'Const FechaVersion = "13/12/2012" ' CAS-16698 - Samconsult
' Se agrego numero de legajo en algunos logs

'Const Version = "1.04"
'Const FechaVersion = "26/12/2012" ' CAS-16698 - Samconsult
' Se informa el codigo externo de la estructura contrato

'Const Version = "1.05"
'Const FechaVersion = "27/03/2013" ' CAS-18942 - Samconsult - Error interfaz Waldbott
' Correcion sobre descripcion de la tabla banco, update de familiares se agrego en el where el nombre del familiar, correccion en la de la fecha de alta
' Correcion en la variable que guardaba tipo de doc de familiar, Verificacion de instancias anteriores ejecutando.
' Los Empleados inactivos, en rhpro no se vuelven a poner inactivo.
' No se envían emails si no se encontraron errores en los empleados.
' Si la estructura del Tipo "Contrato" NO esta configurada como "tiempo determinado", se informa FECVTOCONT en blanco.

'Const Version = "1.06"
'Const FechaVersion = "22/04/2013" ' CAS-18942 - Samconsult - Error interfaz Waldbott
' Correcion cuando se buscan empleados inactivos, del lado de rhpro.

'Const Version = "1.07"
'Const FechaVersion = "02/05/2013" ' CAS-18942 - Samconsult - Error interfaz Waldbott
' Se agrego filtro de empleados que tengan la estructura configurada en el confrep columna 22.
' Correcion en la sincronizacion.

'Const Version = "1.08"
'Const FechaVersion = "22/05/2013" ' CAS-18942 - Samconsult - Error Interfaz Waldbott (Emails Vacíos)
' Se quitaron del mail los empleados que no dieron error.

'Const Version = "1.09"
'Const FechaVersion = "03/06/2013" ' CAS-18942 - Samconsult - Error Interfaz Waldbott (Emails Vacíos) [Entrega 2]
' Se agrego descripcion de error de categoria.

Const Version = "1.10"
Const FechaVersion = "18/06/2013" ' CAS-18942 - Samconsult - Error Interfaz Waldbott (Emails Vacíos) [Entrega 3]
' Solo se analiza los empleados del proceso.

'---------------------------------------------------------------------------------------------------------------------------------------------
Dim dirsalidas As String
Dim usuario As String
Dim Incompleto As Boolean
'-------------------------------------------------------------------------------------------------
'Conexion Externa
'-------------------------------------------------------------------------------------------------
Global ExtConn As New ADODB.Connection
Global ExtConnOra As New ADODB.Connection
Global ExtConnAccess As New ADODB.Connection
Global ExtConnAccess2 As New ADODB.Connection
Global ConnLE As New ADODB.Connection
Global Usa_LE As Boolean
Global Misma_BD As Boolean
Private Type ConexionEmpresa
    ConexionOracle As String    'Guarda la conexion de oracle para la empresa
    ConexionAcces As String     'Guarda la conexion de Acces para la empresa
    estrnroEmpresa As String    'Codigo de la estrucutra empresa configurada
End Type




Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial de la interface Waldbott.
' Autor      : Deluchi Ezequiel
' Fecha      : 04/09/2012
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

    Nombre_Arch = PathFLog & "InterfaceWaldbott-" & NroProcesoBatch & ".log"
    
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
    
    
    On Error Resume Next
    'Abro la conexion
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
    
   
    If App.PrevInstance Then
        StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0 ,bprcestado = 'Pendiente', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
        Flog.writeline "Hay una instancia previa del proceso ejecutando, se pone el proceso en estado pendiente."
        End
    End If

    On Error GoTo 0
    
    On Error GoTo ME_Main
    
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0 ,bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Flog.writeline "Pone el estado en procesando"
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 379 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        bprcparam = IIf(EsNulo(rs_batch_proceso!bprcparam), "", rs_batch_proceso!bprcparam)
        usuario = rs_batch_proceso!iduser
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        Call interfaceWaldbott(NroProcesoBatch, bprcparam)
    Else
        Flog.writeline "no se encontró el proceso"
    End If
    
    TiempoFinalProceso = GetTickCount
    Flog.writeline
    Flog.writeline "--------------------------------------------------------------"
    Flog.writeline "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    
    If HuboError Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Else
        If Incompleto Then
            StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProcesoBatch
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
        Else
            StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
        End If
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


Public Sub interfaceWaldbott(ByVal bpronro As Long, ByVal Parametros As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Interface Waldbott
' Autor      : Deluchi Ezequiel
' Fecha      : 04/09/2012
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------


'-------------------------------------------------------------------------------------------------
'Parametros
'-------------------------------------------------------------------------------------------------
Dim Manual As Boolean
Dim FecDesde As Date
Dim FecHasta As Date


'-------------------------------------------------------------------------------------------------
'Variables
'-------------------------------------------------------------------------------------------------

Dim listaEmpresas   'Lista de las empresas a actualizar los datos
Dim arrayEmpresa    'Arreglo para recorrer por las diferentes empresas
Dim tipDocWDT       'Tipo de Documento para obtener el mapeo de legajo waldbott
Dim tipCodWDT       'Tipo de Codigo que se asignaran a las estructuras
ReDim arr_Errores(71) As String 'Arreglo para guardar los errores, el indice indica el nro de columna segun la especificacion del modelo waldbott
Dim hubo_error As Boolean
Dim hubo_Error_fam As Boolean
Dim convenio As String
Dim tconvenio As Integer
Dim categoria As String
Dim tcategoria As Integer
Dim sindicato As String
Dim tsindicato As Integer
Dim puesto As String
Dim tpuesto As Integer
Dim centroDeCosto As String
Dim tcentroDeCosto As Integer
Dim ttipoEmpleado As Long
Dim ttipoEmpleadoEstrnro As Long
Dim cajaJubilacion As String
Dim tcajaJubilacion As Integer
Dim obraSocialLey As String
Dim tobraSocialLey As Integer
Dim planOSLey As String
Dim tplanOSLey As Integer
Dim contrato As String
Dim tcontrato As Integer
Dim regimenHorario As String
Dim tregimenHorario As Integer
Dim actividad As String
Dim tactividad As Integer
Dim situacionRevista As String
Dim tsituacionRevista As Integer
Dim estado As String
Dim testado As Integer
Dim banco As String
Dim nroCuenta As String
Dim nroCBU As String
Dim emp_Remuneracion
Dim estrnro As String
Dim legajoWaldbott As Long
Dim apellido_Nombre As String
Dim fecha_Nacimiento As String
Dim nacionalidad As String
Dim estado_civil As String
Dim sexo As String
Dim fecha_Alta As Date
Dim fecha_ingreso As Date
Dim estudios As String
Dim tipo_Documento As String
Dim nro_doc As String
Dim nro_cuil As String
Dim calle As String
Dim nro_Direccion As String
Dim piso As String
Dim ofic_depto As String
Dim codigo_Postal As String
Dim localidad As String
Dim provincia As String
Dim telefono As String
Dim convenio_CodExt As String
Dim puesto_CodExt As String
Dim sitRev_CodAFIP As String
Dim estado_Empleado As String
Dim fecha_VtoContrato As String
Dim fecha_Baja As String
Dim causa_Baja As String
Dim empleados_Actualizados As String
Dim EmpPorc As Long         'para llevar el progreso
Dim str_error As String     ' arma la tabla con errores que se enviara por mail
Dim empresa As String
Dim indice As Long
Dim dias
Dim imputacion As String
Dim timputacion As Integer
Dim jurisdiccion As String
Dim tjurisdiccion As Integer
Dim grupo As String
Dim tgrupo As Integer
Dim Aux_Cod_sitr As String
Dim Aux_Cod_sitr1
Dim Aux_Cod_sitr2
Dim Aux_Cod_sitr3
Dim Aux_diainisr1
Dim Aux_diainisr2
Dim Aux_diainisr3
Dim estrnroEmpresa As Long
'-------------------------------------------------------------------------------------------------
'RecordSets
'-------------------------------------------------------------------------------------------------
Dim rs_Consult As New ADODB.Recordset
Dim rs_Empleado As New ADODB.Recordset
Dim rs_familiar As New ADODB.Recordset
Dim rs_aux As New ADODB.Recordset
Dim rs_aux2 As New ADODB.Recordset
Dim rs_access As New ADODB.Recordset
Dim strAccess As String

'Inicio codigo ejecutable
On Error GoTo E_interfaceWaldbott

'-------------------------------------------------------------------------------------------------
'Configuracion del Reporte
'-------------------------------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 0) & "Buscando configuracion del reporte 383."

StrSql = "SELECT * FROM confrep WHERE repnro = 383 ORDER BY confnrocol"
OpenRecordset StrSql, rs_Consult
If rs_Consult.EOF Then
    Flog.writeline Espacios(Tabulador * 1) & "No se encontró la configuración del Reporte 383."
    HuboError = True
    Exit Sub
Else
    ReDim arrayEmpresaConexion(rs_Consult.RecordCount) As ConexionEmpresa
    
    Dim I
    I = 1
    listaEmpresas = "0"
    Do While Not rs_Consult.EOF
        
        Select Case UCase(rs_Consult!conftipo)
             Case "BDO":
                arrayEmpresaConexion(I).ConexionOracle = IIf(EsNulo(rs_Consult!confval), 0, rs_Consult!confval)
                arrayEmpresaConexion(I).estrnroEmpresa = IIf(EsNulo(rs_Consult!confval2), 0, rs_Consult!confval2)
            Case "BDA":
                arrayEmpresaConexion(I).ConexionAcces = IIf(EsNulo(rs_Consult!confval), 0, rs_Consult!confval)
                I = I + 1
            Case "EMP"
                listaEmpresas = listaEmpresas & "," & CStr(rs_Consult!confval)
            Case "DOC"
                tipDocWDT = IIf(EsNulo(rs_Consult!confval), 0, rs_Consult!confval)
            Case "EW"
                tipCodWDT = IIf(EsNulo(rs_Consult!confval), 0, rs_Consult!confval)

        End Select

        
        Select Case UCase(rs_Consult!confnrocol)
            Case 15: 'Tipo estructura Imputacion
                timputacion = CInt(rs_Consult!confval)
            Case 16: 'Tipo estructura Jurisdiccion
                tjurisdiccion = CInt(rs_Consult!confval)
            Case 17:
                tgrupo = CInt(rs_Consult!confval)
            Case 20: 'Estructura Grupo
                testado = CInt(rs_Consult!confval)
            Case 22: 'Estructura Tipo Empleado
                ttipoEmpleado = CLng(rs_Consult!confval)
                ttipoEmpleadoEstrnro = CLng(rs_Consult!confval2)

        End Select
        
        rs_Consult.MoveNext
    Loop
End If
rs_Consult.Close

'cargo los tipos de estructura standard
tconvenio = 19
tcategoria = 3
tsindicato = 16
tpuesto = 4
tcentroDeCosto = 5
tcajaJubilacion = 15
tobraSocialLey = 24
tplanOSLey = 25
tcontrato = 18
tregimenHorario = 21
tactividad = 29
tsituacionRevista = 30

If UBound(arrayEmpresaConexion) < 2 Then
    Flog.writeline Espacios(Tabulador * 1) & "No se configuro el valor de alguna de las 2 conexiones externas (Oracle o Acces)."
    HuboError = True
    Exit Sub
End If

indice = 1

'recupero las empresas configuradas
arrayEmpresa = Split(listaEmpresas, ",")
EmpPorc = 100 / UBound(arrayEmpresa)
Progreso = 0

'Por cada empresa busco los empleados
Do While indice <= UBound(arrayEmpresa)
    Incompleto = False
    estrnroEmpresa = arrayEmpresa(indice)
    Flog.writeline Espacios(Tabulador * 0) & "Buscando Empleados para la empresa " & arrayEmpresa(indice)
    'busco los empleados activos de la empresa con el documento cargado waldbott y que no esten sincronizados
    StrSql = " SELECT distinct empleado.empleg, empleado.ternro, empleado.terape, empleado.terape2, empleado.ternom, empleado.ternom2, tercero.terfecnac, nacionaldes, estructura.estrdabr, "
    StrSql = StrSql & " Tercero.estcivnro, Tercero.paisnro, Tercero.tersex, empleado.empremu, Empleado.empfecalta, pais.paisdesc, estcivil.estcivdesext "
    StrSql = StrSql & " FROM empsinc "
    StrSql = StrSql & " INNER JOIN empleado ON empsinc.esternro = empleado.ternro"
    StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.ternro = empleado.ternro AND his_estructura.tenro = 10 "
    StrSql = StrSql & " AND his_estructura.estrnro = " & arrayEmpresa(indice)
    StrSql = StrSql & " AND (htetdesde <= " & ConvFecha(Now) & " AND (htethasta IS NULL OR htethasta >= " & ConvFecha(Now) & ")) "
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro"
    'temp es la estructura de tipo empleado configurada en la columna 22 del confrep
    StrSql = StrSql & " INNER JOIN his_estructura temp ON temp.ternro = empleado.ternro AND temp.tenro = " & ttipoEmpleado
    StrSql = StrSql & " AND temp.estrnro = " & ttipoEmpleadoEstrnro
    StrSql = StrSql & " AND (temp.htetdesde <= " & ConvFecha(Now) & " AND (temp.htethasta IS NULL OR temp.htethasta >= " & ConvFecha(Now) & ")) "
    StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro = empleado.ternro "
    StrSql = StrSql & " INNER JOIN pais ON pais.paisnro = tercero.paisnro"
    StrSql = StrSql & " INNER JOIN batch_proceso ON batch_proceso.bprcempleados = empleado.ternro AND bpronro = " & bpronro
    StrSql = StrSql & " INNER JOIN nacionalidad ON nacionalidad.nacionalnro = tercero.nacionalnro "
    StrSql = StrSql & " INNER JOIN estcivil ON estcivil.estcivnro = tercero.estcivnro"
    StrSql = StrSql & " WHERE essinc = 0 "
    OpenRecordset StrSql, rs_Empleado
    If rs_Empleado.RecordCount = 0 Then
        IncPorc = EmpPorc / 1
    Else
        IncPorc = EmpPorc / rs_Empleado.RecordCount
    End If
    str_error = ""
    '-----------------------------------------------------------------------------------------------
    'Busco la conexion Access de la empresa
    '-----------------------------------------------------------------------------------------------
    StrSql = " SELECT cnstring FROM conexion "
    StrSql = StrSql & " WHERE cnnro = " & arrayEmpresaConexion(indice).ConexionAcces
    OpenRecordset StrSql, rs_aux
    
    If Not rs_aux.EOF Then 'Si no existe la fecha de alta reconocida me quedo con la fecha de alta del empleado
        strconexion = rs_aux!cnstring
    Else
        Flog.writeline Espacios(Tabulador * 0) & "Error en la configuracion de la conexion Access."
    End If
    If rs_aux.State = adStateOpen Then rs_aux.Close
    
    'Abro la conexion con las tablas de Access
    OpenConnExt strconexion, ExtConnAccess 'ExtConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline "Problemas en la conexion Access"
        Exit Sub
    End If
    Flog.writeline Espacios(Tabulador * 0) & "Conexion con Access establecida."

    '-----------------------------------------------------------------------------------------------
    'Busco la conexion Oracle de la empresa
    '-----------------------------------------------------------------------------------------------
    StrSql = " SELECT cnstring FROM conexion "
    StrSql = StrSql & " WHERE cnnro = " & arrayEmpresaConexion(indice).ConexionOracle
    OpenRecordset StrSql, rs_aux
    If Not rs_aux.EOF Then
        strconexion = rs_aux!cnstring
    Else
        Flog.writeline Espacios(Tabulador * 0) & "Error en la configuracion de la conexion Oracle."
    End If
    If rs_aux.State = adStateOpen Then rs_aux.Close
    
    
    'Abro la conexion con las tablas de Oracle
    OpenConnExt strconexion, ExtConnOra
    'OpenConnection strconexion, ExtConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline "Problemas en la conexion Oracle"
        Exit Sub
    End If

    Flog.writeline Espacios(Tabulador * 0) & "Conexion con Oracle establecida."
    
    Dim arr_i
    empleados_Actualizados = "0"
    If Not rs_Empleado.EOF Then
        empresa = rs_Empleado!estrdabr
    Else
        Flog.writeline Espacios(Tabulador * 0) & "No se encontraron empleados desincronizados para la empresa " & arrayEmpresa(indice) & "."
    End If
    Do While Not rs_Empleado.EOF
        For arr_i = 0 To UBound(arr_Errores)
            arr_Errores(arr_i) = ""
        Next
        hubo_error = False
        Progreso = Progreso + IncPorc
        Flog.writeline Espacios(Tabulador * 0) & "*************************************************************************************************************"
        Flog.writeline Espacios(Tabulador * 0) & "Busco los datos de el ternro: " & rs_Empleado!ternro & ". Legajo: " & rs_Empleado!empleg
        arr_Errores(0) = rs_Empleado!ternro
        
        
        StrSql = " SELECT nrodoc FROM ter_doc WHERE tidnro = " & tipDocWDT & " AND ternro = " & rs_Empleado!ternro
        OpenRecordset StrSql, rs_aux
        If Not rs_aux.EOF Then
            legajoWaldbott = rs_aux!nrodoc 'char 7
        Else
            legajoWaldbott = "0" 'char 7
        End If
        If rs_aux.State = adStateOpen Then rs_aux.Close
        
        
        Call checkError(legajoWaldbott, True, 7, bpronro, "legajo", arr_Errores)
       'Error Critico
        If arr_Errores(1) <> "" Then
            hubo_error = True
        End If
        
        apellido_Nombre = rs_Empleado!terape & IIf(EsNulo(rs_Empleado!terape2), " ", " " & rs_Empleado!terape2 & " ") & rs_Empleado!ternom & IIf(EsNulo(rs_Empleado!ternom2), "", " " & rs_Empleado!ternom2) 'char 30
        Call checkError(apellido_Nombre, True, 29, bpronro, "Nombre", arr_Errores)
        'Error Critico
        If arr_Errores(2) <> "" Then
            hubo_error = True
        End If
        
        fecha_Nacimiento = rs_Empleado!terfecnac 'date
        Call checkError(fecha_Nacimiento, True, 10, bpronro, "FechaNacimiento", arr_Errores)
        'Error Critico
        If arr_Errores(4) <> "" Then
            hubo_error = True
        End If
        
        'Nacionalidad
        nacionalidad = rs_Empleado!nacionaldes
        strAccess = " SELECT codigo FROM vw_nacionalidades"
        strAccess = strAccess & " WHERE funcion = '" & nacionalidad & "'"
        OpenRecordsetExt strAccess, rs_access, ExtConnAccess
    
        If Not rs_access.EOF Then
            nacionalidad = rs_access!codigo
            Flog.writeline Espacios(Tabulador * 2) & "Nacionalidad existente en la tabla Access."
        Else
            Flog.writeline Espacios(Tabulador * 2) & "Error: Nacionalidad no existente en la tabla Access."
            arr_Errores(5) = "Error"
            'Error Critico
            hubo_error = True
        End If
        If rs_access.State = adStateOpen Then rs_access.Close
        Call checkError(nacionalidad, False, 1, bpronro, "nacionalidad", arr_Errores)
        
        estado_civil = rs_Empleado!estcivdesext 'char 1
        Call checkError(estado_civil, True, 1, bpronro, "estadoCivil", arr_Errores)
        If arr_Errores(6) <> "" Then
            hubo_error = True
        End If

        sexo = IIf(rs_Empleado!tersex = -1, "M", "F") 'char 1
        Call checkError(sexo, True, 1, bpronro, "sexo", arr_Errores)
    
        'Busco la fecha de alta (fecha desde fase activa)
        '**********************************************************************************
        StrSql = " SELECT altfec FROM fases where empleado = " & rs_Empleado!ternro & " ORDER BY altfec DESC "
        OpenRecordset StrSql, rs_aux
        If Not rs_aux.EOF Then
            fecha_Alta = rs_aux!altfec
            Call checkError(fecha_Alta, True, 10, bpronro, "FechaAlta", arr_Errores)
            Flog.writeline Espacios(Tabulador * 1) & "Fecha de alta para ternro: " & rs_Empleado!ternro & ".  Legajo: " & rs_aux!altfec
        End If
        If arr_Errores(8) <> "" Then
            hubo_error = True
            Flog.writeline Espacios(Tabulador * 1) & "Error Fecha de Alta para el ternro: " & rs_Empleado!ternro & ".  Legajo: " & rs_Empleado!empleg
        End If
        'Busco y la fecha de alta reconocida y chequeo si existe
        '**********************************************************************************
        Dim Fecha_limite As String
        Dim Fecha_baja_nueva As String
        Dim entra As Boolean
        
        StrSql = " SELECT altfec, bajfec FROM fases where real = -1 and fasrecofec = -1 AND empleado = " & rs_Empleado!ternro
        OpenRecordset StrSql, rs_aux
        dias = 0
        If Not rs_aux.EOF Then
            If Not EsNulo(rs_aux!bajfec) Then
                Fecha_limite = " AND altfec > " & ConvFecha(rs_aux!bajfec)
                Fecha_baja_nueva = rs_aux!bajfec
                entra = True
            Else
                Fecha_limite = " AND 1 = 2" 'La fase esta como fase reconocida y activa
                Fecha_baja_nueva = ""
                entra = False
                
            End If
            
            If Not EsNulo(rs_aux!bajfec) And Not EsNulo(rs_aux!altfec) Then
                dias = dias + DateDiff("d", rs_aux!altfec, rs_aux!bajfec)
            Else
                If EsNulo(rs_aux!bajfec) And Not EsNulo(rs_aux!altfec) Then
                    dias = dias + DateDiff("d", rs_aux!altfec, Now()) - 1
                End If
            End If
        Else
            Fecha_limite = ""
        End If
        'Calculo la diferencia de dias sobre las otras fases para restarle a la fecha de ingreso
        StrSql = " SELECT altfec, bajfec FROM fases where real = -1 and empleado = " & rs_Empleado!ternro & Fecha_limite
        OpenRecordset StrSql, rs_aux2
        
        Do While Not rs_aux2.EOF
            'Le sumo por el dia actual
            If Not EsNulo(rs_aux2!bajfec) And Not EsNulo(rs_aux2!altfec) Then
                dias = dias + DateDiff("d", rs_aux2!altfec, rs_aux2!bajfec)
                If Fecha_baja_nueva <> "" Then
                    If CDate(Fecha_baja_nueva) < CDate(rs_aux2!bajfec) Then
                        Fecha_baja_nueva = rs_aux2!bajfec
                    End If
                End If
                
            Else
                If EsNulo(rs_aux2!bajfec) And Not EsNulo(rs_aux2!altfec) Then
                    dias = dias + DateDiff("d", rs_aux2!altfec, Now())
                End If
            End If
            rs_aux2.MoveNext
        Loop
        

        dias = (dias + 1) * -1
        fecha_ingreso = DateAdd("d", dias, Now())
        
        Call checkError(fecha_ingreso, True, 10, bpronro, "FechaAlta", arr_Errores)
        'Error Critico
        If arr_Errores(8) <> "" Then
            hubo_error = True
        End If
        
        '********************************* Estudios *************************************************
        StrSql = " SELECT titdesabr FROM cap_estformal "
        StrSql = StrSql & " INNER JOIN titulo ON titulo.titnro = cap_estformal.titnro "
        StrSql = StrSql & " Where ternro = " & rs_Empleado!ternro
        OpenRecordset StrSql, rs_aux
        If Not rs_aux.EOF Then
            estudios = IIf(EsNulo(rs_aux!titdesabr), "", rs_aux!titdesabr) 'char 30
        End If
        If rs_aux.State = adStateOpen Then rs_aux.Close
        Call checkError(estudios, False, 30, bpronro, "Estudios", arr_Errores)
        
        '******************************* Tipo de Documento *******************************************
        StrSql = " SELECT tidnro, nrodoc FROM  ter_doc"
        StrSql = StrSql & " WHERE ternro = " & rs_Empleado!ternro & " AND tidnro <= 11 ORDER BY tidnro ASC"
        OpenRecordset StrSql, rs_aux
        If Not rs_aux.EOF Then
            Select Case CInt(rs_aux!tidnro)
                Case 1:
                    tipo_Documento = "DN"
                Case 2:
                    tipo_Documento = "LE"
                Case 3:
                    tipo_Documento = "LC"
                Case 4:
                    tipo_Documento = "CI"
                Case 5:
                    tipo_Documento = "PA"
                Case 11:
                    tipo_Documento = "CM"
                    
                Case Else:
                    Flog.writeline Espacios(Tabulador * 1) & "Error no existe documento para el ternro: " & rs_Empleado!ternro & ".  Legajo: " & rs_Empleado!empleg
                    tipo_Documento = ""
                    arr_Errores(11) = "Error"
                    hubo_error = True
            End Select
            nro_doc = rs_aux!nrodoc
            Call checkError(nro_doc, True, 9, bpronro, "nrodoc", arr_Errores)
            'Error Critico
            If arr_Errores(12) <> "" Then
                hubo_error = True
            End If

        End If
        If rs_aux.State = adStateOpen Then rs_aux.Close
        Call checkError(tipo_Documento, True, 2, bpronro, "tipoDocumento", arr_Errores)
      
        
        '********************************* CUIL ******************************************************
        StrSql = " SELECT nrodoc FROM  ter_doc"
        StrSql = StrSql & " WHERE ternro = " & rs_Empleado!ternro & " AND tidnro = 10"
        OpenRecordset StrSql, rs_aux
        If Not rs_aux.EOF Then
            nro_cuil = IIf(EsNulo(rs_aux!nrodoc), "", rs_aux!nrodoc) 'char 15
            Call checkError(nro_cuil, True, 15, bpronro, "cuil", arr_Errores)
            'Error Critico
            If arr_Errores(13) <> "" Then
                hubo_error = True
            End If
        Else
            Flog.writeline Espacios(Tabulador * 1) & "Error no existe cuil para el ternro: " & rs_Empleado!ternro & ". Legajo: " & rs_Empleado!empleg
        End If
        If rs_aux.State = adStateOpen Then rs_aux.Close
        
        
        '********************************** Domicilio ************************************************
        '*********************************************************************************************
        StrSql = " SELECT calle, nro, piso, oficdepto,codigopostal, locdesc, provdesc, telnro FROM cabdom "
        StrSql = StrSql & " INNER JOIN detdom ON detdom.domnro = cabdom.domnro AND domdefault = -1 "
        StrSql = StrSql & " INNER JOIN localidad ON localidad.locnro = detdom.locnro "
        StrSql = StrSql & " INNER JOIN provincia ON provincia.provnro = detdom.provnro "
        StrSql = StrSql & " LEFT JOIN telefono ON telefono.domnro = cabdom.domnro "
        StrSql = StrSql & " WHERE ternro = " & rs_Empleado!ternro
        OpenRecordset StrSql, rs_aux
        If Not rs_aux.EOF Then
            calle = IIf(EsNulo(rs_aux!calle), "", rs_aux!calle) 'char 25
            nro_Direccion = IIf(EsNulo(rs_aux!nro), "", rs_aux!nro) 'char 5
            piso = IIf(EsNulo(rs_aux!piso), "", rs_aux!piso) 'char 2
            ofic_depto = IIf(EsNulo(rs_aux!oficdepto), "", rs_aux!oficdepto) 'char 4
            codigo_Postal = IIf(EsNulo(rs_aux!codigopostal), "", rs_aux!codigopostal) 'char 8
            localidad = IIf(EsNulo(rs_aux!locdesc), "", rs_aux!locdesc) 'char 18
            provincia = IIf(EsNulo(rs_aux!provdesc), "", rs_aux!provdesc) 'char 2 access
            telefono = IIf(EsNulo(rs_aux!telnro), "", rs_aux!telnro) 'char 14
        End If
        If rs_aux.State = adStateOpen Then rs_aux.Close
        
        Call checkError(calle, False, 25, bpronro, "calle", arr_Errores)
        Call checkError(nro_Direccion, False, 5, bpronro, "nroDireccion", arr_Errores)
        Call checkError(piso, False, 2, bpronro, "piso", arr_Errores)
        Call checkError(ofic_depto, False, 4, bpronro, "oficdepto", arr_Errores)
        Call checkError(codigo_Postal, False, 8, bpronro, "cp", arr_Errores)
        Call checkError(localidad, False, 18, bpronro, "localidad", arr_Errores)
        
        
        strAccess = " SELECT id_provincia FROM vw_provincias"
        strAccess = strAccess & " WHERE descripcion = '" & provincia & "'"
        OpenRecordsetExt strAccess, rs_access, ExtConnAccess

        If Not rs_access.EOF Then
            provincia = rs_access!id_provincia
            Call checkError(provincia, True, 2, bpronro, "provincia", arr_Errores)
            Flog.writeline Espacios(Tabulador * 2) & "provincia existente en la tabla Access."
        Else
            Flog.writeline Espacios(Tabulador * 2) & "Error: provincia no existente en la tabla Access."
            arr_Errores(20) = "Error"
            provincia = ""
            hubo_error = True
        End If
        If rs_access.State = adStateOpen Then rs_access.Close
        
        Call checkError(telefono, False, 14, bpronro, "telefono", arr_Errores)
        
        '*******************************  Busco los datos de la estructura convenio ***********************************
        '**************************************************************************************************************
        buscarEstructura CLng(rs_Empleado!ternro), tconvenio, convenio, estrnro, "estrcodext" 'varchar7
        convenio_CodExt = estrnro
        Call checkError(convenio_CodExt, False, 7, bpronro, "Convenio", arr_Errores)
        
        '*******************************  Busco los datos de la estructura categoria ***********************************
        '***************************************************************************************************************
        buscarEstructura CLng(rs_Empleado!ternro), tcategoria, categoria, estrnro, "" 'varchar 4 access
        If categoria = "" Then
            Flog.writeline Espacios(Tabulador * 1) & "Error: no existe categoria para el ternro: " & rs_Empleado!ternro & ". Legajo: " & rs_Empleado!empleg
        End If
        strAccess = " SELECT codigo FROM vw_categorias"
        strAccess = strAccess & " WHERE funcion = '" & categoria & "'"
        OpenRecordsetExt strAccess, rs_access, ExtConnAccess
        If Not rs_access.EOF Then
            categoria = rs_access!codigo
            Call checkError(categoria, True, 4, bpronro, "categoria", arr_Errores)
            Flog.writeline Espacios(Tabulador * 2) & "categoria existente en la tabla Access."
            If arr_Errores(23) <> "" Then
                hubo_error = True
            End If
        Else
            Flog.writeline Espacios(Tabulador * 2) & "Error: categoria no existente en la tabla Access."
            arr_Errores(23) = "Error"
            hubo_error = True
            categoria = ""
        End If
        If rs_access.State = adStateOpen Then rs_access.Close
        
        
        'Si no se informa categoria es obligatorio
        emp_Remuneracion = IIf(EsNulo(rs_Empleado!empremu), 0, rs_Empleado!empremu) ' 9(15).99
        If categoria <> "" Then
            Call checkError(emp_Remuneracion, False, 1, bpronro, "Remuneracion", arr_Errores)
        Else
            Call checkError(emp_Remuneracion, False, 1, bpronro, "Remuneracion", arr_Errores)
            'Error Critico
            If arr_Errores(42) <> "" Then
                hubo_error = True
            End If
        End If

        
        '*******************************  Busco los datos de la estructura sindicato ***********************************
        '***************************************************************************************************************
        buscarEstructura CLng(rs_Empleado!ternro), tsindicato, sindicato, estrnro, "" 'varchar 2 access
        If sindicato = "" Then
            Flog.writeline Espacios(Tabulador * 1) & "Error: no existe sindicato para el ternro: " & rs_Empleado!ternro & ". Legajo: " & rs_Empleado!empleg
        End If
        strAccess = " SELECT codigo FROM vw_sindicatos"
        strAccess = strAccess & " WHERE descripcion = '" & sindicato & "'"
        OpenRecordsetExt strAccess, rs_access, ExtConnAccess
        If Not rs_access.EOF Then
            sindicato = rs_access!codigo
            Flog.writeline Espacios(Tabulador * 2) & "sindicato existente en la tabla Access."
            Call checkError(sindicato, True, 2, bpronro, "sindicato", arr_Errores)
            If arr_Errores(24) <> "" Then
                hubo_error = True
            End If
        Else
            Flog.writeline Espacios(Tabulador * 2) & "Error: sindicato no existente en la tabla Access."
            arr_Errores(24) = "Error"
            hubo_error = True
            sindicato = ""
        End If
        If rs_access.State = adStateOpen Then rs_access.Close
        

        
        '*******************************  Busco los datos de la estructura jurisdiccion ********************************
        '***************************************************************************************************************
        buscarEstructura CLng(rs_Empleado!ternro), tjurisdiccion, jurisdiccion, estrnro, "" 'varchar 2 access
        If jurisdiccion = "" Then
            Flog.writeline Espacios(Tabulador * 1) & "No existe jurisdiccion para el ternro: " & rs_Empleado!ternro & ". Legajo: " & rs_Empleado!empleg
        End If
        strAccess = " SELECT codigo FROM vw_jurisdicciones "
        strAccess = strAccess & " WHERE descripcion = '" & jurisdiccion & "'"
        OpenRecordsetExt strAccess, rs_access, ExtConnAccess
        If Not rs_access.EOF Then
            jurisdiccion = rs_access!codigo
            Flog.writeline Espacios(Tabulador * 2) & "Jurisdiccion existente en la tabla Access."
            Call checkError(jurisdiccion, True, 2, bpronro, "jurisdiccion", arr_Errores)
            If arr_Errores(71) <> "" Then
                hubo_error = True
            End If
        Else
            Flog.writeline Espacios(Tabulador * 2) & "jurisdiccion no existente en la tabla Access."
            arr_Errores(71) = "Error"
            hubo_error = True
            jurisdiccion = ""
        End If
        If rs_access.State = adStateOpen Then rs_access.Close
        
        '********************************  Busco los datos de la estructura grupo **************************************
        '***************************************************************************************************************
        buscarEstructura CLng(rs_Empleado!ternro), tgrupo, grupo, estrnro, "" 'varchar 2 access
        If grupo = "" Then
            Flog.writeline Espacios(Tabulador * 2) & "No existe grupo para el ternro: " & rs_Empleado!ternro & ". Legajo: " & rs_Empleado!empleg
        Else
            Flog.writeline Espacios(Tabulador * 2) & "Grupo existente."
            Call checkError(grupo, True, 4, bpronro, "grupo", arr_Errores)
        End If
        
        '*******************************  Busco los datos de la estructura Puesto ***********************************
        '***************************************************************************************************************
        buscarEstructura CLng(rs_Empleado!ternro), tpuesto, puesto, estrnro, "estrnro" 'varchar 20
        Call checkError(puesto, False, 20, bpronro, "Puesto", arr_Errores)
    
    
        '*******************************  Busco los datos de la estructura Imputacion **********************************
        '***************************************************************************************************************
        buscarEstructura CLng(rs_Empleado!ternro), timputacion, imputacion, estrnro, "" 'varchar 20
        Call checkError(imputacion, False, 10, bpronro, "Imputacion", arr_Errores)
    
        '*****************************  Busco los datos de la estructura Centro de Costo *******************************
        '***************************************************************************************************************
        buscarEstructura CLng(rs_Empleado!ternro), tcentroDeCosto, centroDeCosto, estrnro, "" 'varchar 10 access
        If centroDeCosto = "" Then
            Flog.writeline Espacios(Tabulador * 1) & "Error: no existe Centro de Costo para el ternro: " & rs_Empleado!ternro & ". Legajo: " & rs_Empleado!empleg
        End If
        strAccess = " SELECT codigo FROM vw_ccostos"
        strAccess = strAccess & " WHERE nombre = '" & centroDeCosto & "'"
        OpenRecordsetExt strAccess, rs_access, ExtConnAccess
        If Not rs_access.EOF Then
            centroDeCosto = rs_access!codigo
            Flog.writeline Espacios(Tabulador * 2) & "Centro De Costo existente en la tabla Access."
            Call checkError(centroDeCosto, True, 10, bpronro, "ccosto", arr_Errores)
        Else
            Flog.writeline Espacios(Tabulador * 2) & "Error: Centro De Costo no existente en la tabla Access (imposible armar legajo)."
            arr_Errores(26) = "Error"
            centroDeCosto = "" 'sin el centro de costo no se puede armar legajo aborto empleado
            hubo_error = True
        End If
        If rs_access.State = adStateOpen Then rs_access.Close
        
    
        '**************************  Busco los datos de la estructura Caja de Jubilacion *******************************
        '***************************************************************************************************************
        buscarEstructura CLng(rs_Empleado!ternro), tcajaJubilacion, cajaJubilacion, estrnro, "" 'varchar 2 access
        If cajaJubilacion = "" Then
            Flog.writeline Espacios(Tabulador * 1) & "Error: no existe caja de jubilacion para el ternro: " & rs_Empleado!ternro & ". Legajo: " & rs_Empleado!empleg
        End If
        strAccess = " SELECT codigo FROM vw_jubilacion"
        strAccess = strAccess & " WHERE descripcion = '" & cajaJubilacion & "'"
        OpenRecordsetExt strAccess, rs_access, ExtConnAccess
        If Not rs_access.EOF Then
            cajaJubilacion = rs_access!codigo
            Call checkError(cajaJubilacion, True, 2, bpronro, "cjubilacion", arr_Errores)
            Flog.writeline Espacios(Tabulador * 2) & "Caja de Jubilacion existente en la tabla Access."
            If arr_Errores(27) <> "" Then
                hubo_error = True
            End If
        Else
            Flog.writeline Espacios(Tabulador * 2) & "Error: Caja de Jubilacion no existente en la tabla Access."
            arr_Errores(27) = "Error"
            cajaJubilacion = ""
            hubo_error = True
        End If
        If rs_access.State = adStateOpen Then rs_access.Close
        
        
        '**************************  Busco los datos de la estructura Obra social **************************************
        '***************************************************************************************************************
        buscarEstructura CLng(rs_Empleado!ternro), tobraSocialLey, obraSocialLey, estrnro, "" 'varchar 2 access
        If obraSocialLey = "" Then
            Flog.writeline Espacios(Tabulador * 1) & "Error: no existe obra social para el ternro: " & rs_Empleado!ternro & ". Legajo: " & rs_Empleado!empleg
        End If
        strAccess = " SELECT codigo FROM vw_obras_sociales"
        strAccess = strAccess & " WHERE descripcion = '" & obraSocialLey & "'"
        OpenRecordsetExt strAccess, rs_access, ExtConnAccess
        If Not rs_access.EOF Then
            obraSocialLey = rs_access!codigo
            Flog.writeline Espacios(Tabulador * 2) & "obraSocialLey existente en la tabla Access."
            Call checkError(obraSocialLey, True, 2, bpronro, "obraSocial", arr_Errores)
            If arr_Errores(28) <> "" Then
                hubo_error = True
            End If
        Else
            Flog.writeline Espacios(Tabulador * 2) & "Error: obraSocialLey no existente en la tabla Access."
            arr_Errores(28) = obraSocialLey
            obraSocialLey = ""
            hubo_error = True
        End If
        If rs_access.State = adStateOpen Then rs_access.Close
        
        
        '***********************  Busco los datos de la estructura Plan de Obra social *********************************
        '***************************************************************************************************************
        buscarEstructura CLng(rs_Empleado!ternro), tplanOSLey, planOSLey, estrnro, ""   'varchar 1 access
        If planOSLey = "" Then
            Flog.writeline Espacios(Tabulador * 2) & "No existe Plan de Obra social para el ternro: " & rs_Empleado!ternro & "(No obligatorio). Legajo: " & rs_Empleado!empleg
        End If
        strAccess = " SELECT codigo FROM vw_planes_obra_social"
        strAccess = strAccess & " WHERE descripcion = '" & planOSLey & "'"
        OpenRecordsetExt strAccess, rs_access, ExtConnAccess
        If Not rs_access.EOF Then
            planOSLey = rs_access!codigo
            Flog.writeline Espacios(Tabulador * 2) & "Plan OS Ley existente en la tabla Access."
            Call checkError(planOSLey, False, 1, bpronro, "planObraSocial", arr_Errores)
        Else
            Flog.writeline Espacios(Tabulador * 2) & "Plan OS Ley no existente en la tabla Access (No obligatorio)."
            planOSLey = ""
        End If
        If rs_access.State = adStateOpen Then rs_access.Close
        
    
        '******************************  Busco los datos de la estructura Contrato *************************************
        '***************************************************************************************************************
        buscarEstructura CLng(rs_Empleado!ternro), tcontrato, contrato, estrnro, "estrnro,htetdesde,estrcodext"
        If estrnro = "" Then
            Flog.writeline Espacios(Tabulador * 0) & "El ternro: " & rs_Empleado!ternro & " no tiene contrato. Legajo: " & rs_Empleado!empleg
            contrato = ""
        Else
            'busco en la tabla access
            '----------------------------------------------------------------------------------------------------------------------
            'strAccess = " SELECT codigo FROM vw_contrato"
            'strAccess = strAccess & " WHERE descripcion = '" & contrato & "'"
            'OpenRecordsetExt strAccess, rs_access, ExtConnAccess
            'If Not rs_access.EOF Then
                'contrato = rs_access!codigo
                'Flog.writeline Espacios(Tabulador * 2) & "Contrato existente en la tabla Access."
                Flog.writeline Espacios(Tabulador * 2) & "Contrato existente."
                contrato = Split(estrnro, ",")(2)
                Call checkError(contrato, False, 3, bpronro, "contrato", arr_Errores)
                
                StrSql = " SELECT tcind, tcanios, tcmeses FROM tipocont WHERE estrnro = " & Split(estrnro, ",")(0)
                OpenRecordset StrSql, rs_aux
                If Not rs_aux.EOF Then
                    If CLng(rs_aux!tcind) = 0 Then
                        fecha_VtoContrato = ""
                    Else
                        Call fechaCalculada(CDate(Split(estrnro, ",")(1)), IIf(EsNulo(rs_aux!tcanios), 0, rs_aux!tcanios), IIf(EsNulo(rs_aux!tcmeses), 0, rs_aux!tcmeses), 0, fecha_VtoContrato)
                    End If
                Else
                    Flog.writeline Espacios(Tabulador * 0) & "El contrato no tiene fecha de vencimiento (ternro: " & rs_Empleado!ternro & ") Legajo: " & rs_Empleado!empleg
                End If
                If rs_aux.State = adStateOpen Then rs_aux.Close
                Call checkError(fecha_VtoContrato, False, 10, bpronro, "fecha", arr_Errores)
            'Else
            '    Flog.writeline Espacios(Tabulador * 2) & "Contrato no existente en la tabla Access (No obligatorio)."
            '    contrato = ""
            'End If
            If rs_access.State = adStateOpen Then rs_access.Close
            '----------------------------------------------------------------------------------------------------------------------
        End If
        
        '***********************  Busco los datos de la estructura Regimen Horario**************************************
        '***************************************************************************************************************
        buscarEstructura CLng(rs_Empleado!ternro), tregimenHorario, regimenHorario, estrnro, ""
        If regimenHorario = "" Then
            Flog.writeline Espacios(Tabulador * 2) & "No existe Regimen Horario para el ternro: " & rs_Empleado!ternro & ". Legajo: " & rs_Empleado!empleg
        End If
        Call checkError(regimenHorario, False, 20, bpronro, "regHorario", arr_Errores)
    
        '***********************  Busco los datos de la estructura Banco ***********************************************
        '***************************************************************************************************************
    
        StrSql = " SELECT bandesc,ctabnro, ctabcbu, estrdabr FROM ctabancaria " & _
                 " LEFT JOIN banco ON ctabancaria.banco = banco.ternro " & _
                 " LEFT JOIN estructura ON estructura.estrnro = banco.estrnro " & _
                 " WHERE ctabancaria.ternro = " & rs_Empleado!ternro & " AND ctabestado = -1 "
        OpenRecordset StrSql, rs_aux
        If Not rs_aux.EOF Then
            banco = IIf(EsNulo(rs_aux!estrdabr), "", rs_aux!estrdabr)
            nroCuenta = IIf(EsNulo(rs_aux!ctabnro), "", rs_aux!ctabnro)
            nroCBU = IIf(EsNulo(rs_aux!ctabcbu), "", rs_aux!ctabcbu)
            'Busco el banco en la tabla access
            '---------------------------------------------------------------------------------------
            strAccess = " SELECT codigo FROM vw_bancos"
            strAccess = strAccess & " WHERE descripcion = '" & banco & "'"
            OpenRecordsetExt strAccess, rs_access, ExtConnAccess
            If Not rs_access.EOF Then
                banco = rs_access!codigo
                Flog.writeline Espacios(Tabulador * 2) & "Banco existente en la tabla Access."
                Call checkError(banco, False, 2, bpronro, "banco", arr_Errores)
                Call checkError(nroCuenta, False, 16, bpronro, "nroCuenta", arr_Errores)
                Call checkError(nroCBU, False, 30, bpronro, "nroCBU", arr_Errores)
            Else
                Flog.writeline Espacios(Tabulador * 2) & "Banco no existente en la tabla Access (No obligatorio)."
                banco = ""
                nroCuenta = ""
                nroCBU = ""
            End If
            '---------------------------------------------------------------------------------------
        Else
            Flog.writeline Espacios(Tabulador * 1) & "No existe Cuenta bancaria o banco para el ternro: " & rs_Empleado!ternro & ". Legajo: " & rs_Empleado!empleg
        End If
        If rs_aux.State = adStateOpen Then rs_aux.Close
    
        '***********************  Busco los datos de la estructura Actividad SIJP **************************************
        '***************************************************************************************************************
        buscarEstructura CLng(rs_Empleado!ternro), tactividad, actividad, estrnro, ""
        If actividad = "" Then
            Flog.writeline Espacios(Tabulador * 1) & "Error: no existe Actividad SIJP para el ternro: " & rs_Empleado!ternro & ". Legajo: " & rs_Empleado!empleg
        End If
        'Busco el banco en la tabla access
        '---------------------------------------------------------------------------------------
        strAccess = " SELECT codigo FROM vw_actividades"
        strAccess = strAccess & " WHERE descripcion = '" & actividad & "'"
        OpenRecordsetExt strAccess, rs_access, ExtConnAccess
        If Not rs_access.EOF Then
            actividad = rs_access!codigo
            Flog.writeline Espacios(Tabulador * 2) & "Actividad existente en la tabla Access."
            Call checkError(actividad, True, 2, bpronro, "actividad", arr_Errores)
            If arr_Errores(37) <> "" Then
                hubo_error = True
            End If
        Else
            Flog.writeline Espacios(Tabulador * 2) & "Actividad no existente en la tabla Access."
            arr_Errores(37) = "Error"
            actividad = ""
            hubo_error = True
        End If
    
        Call sit_rev(CLng(rs_Empleado!ternro), estrnroEmpresa, Aux_Cod_sitr, Aux_Cod_sitr1, Aux_Cod_sitr2, Aux_Cod_sitr3, Aux_diainisr1, Aux_diainisr2, Aux_diainisr3)
        'Fin cambio version 1.1
        '**************************  Busco los datos de la estructura Estado *******************************************
        '***************************************************************************************************************
        buscarEstructura CLng(rs_Empleado!ternro), testado, estado, estado_Empleado, "estrcodext"
        If estado_Empleado = "" Then
            Flog.writeline Espacios(Tabulador * 2) & "Error el estado del empleado(ternro: " & rs_Empleado!ternro & ") no existe. Legajo: " & rs_Empleado!empleg
        End If
        Call checkError(estado_Empleado, True, 1, bpronro, "estadoEmpleado", arr_Errores)
        'Error Critico
        If arr_Errores(39) <> "" Then
            hubo_error = True
        End If

        '****************************  Busco los datos de la Causa de baja *********************************************
        '***************************************************************************************************************
        StrSql = " SELECT bajfec, caudes FROM fases "
        StrSql = StrSql & " INNER JOIN causa on causa.caunro = fases.caunro "
        StrSql = StrSql & " WHERE bajfec is not null AND not EXISTS(select fasnro from fases where bajfec is null and empleado = " & rs_Empleado!ternro & " ) and empleado = " & rs_Empleado!ternro
        StrSql = StrSql & " ORDER BY fasnro DESC "
        OpenRecordset StrSql, rs_aux
        If Not rs_aux.EOF Then
            fecha_Baja = rs_aux!bajfec
            causa_Baja = rs_aux!caudes
            Select Case UCase(causa_Baja)
                Case UCase("Abandono trabajo"):
                    causa_Baja = "A"
                Case UCase("fin contrato"):
                    causa_Baja = "C"
                Case UCase("Despido"):
                    causa_Baja = "D"
                Case UCase("Relevo"):
                    causa_Baja = "E"
                Case UCase("Fallecimiento"):
                    causa_Baja = "F"
                Case UCase("Despido con Causa"):
                    causa_Baja = "J"
                Case UCase("Clientela"):
                    causa_Baja = "L"
                Case UCase("Despido Maternidad"):
                    causa_Baja = "M"
                Case UCase("Otros"):
                    causa_Baja = "O"
                Case UCase("despidos Sin preaviso"):
                    causa_Baja = "P"
                Case UCase("Causa Fuerza Mayor"):
                    causa_Baja = "Q"
                Case UCase("renuncia"):
                    causa_Baja = "R"
                Case UCase("Despido Matrimonio"):
                    causa_Baja = "T"
                Case Else:
                    Flog.writeline Espacios(Tabulador * 2) & "Error: El ternro: " & rs_Empleado!ternro & ", esta activo.  Legajo: " & rs_Empleado!empleg
            End Select
            Call checkError(causa_Baja, False, 1, bpronro, "causaBaja", arr_Errores)
        Else
            Flog.writeline Espacios(Tabulador * 2) & "El ternro: " & rs_Empleado!ternro & ", esta activo. Legajo: " & rs_Empleado!empleg
            causa_Baja = ""
            fecha_Baja = ""
        End If
        If rs_aux.State = adStateOpen Then rs_aux.Close
    
        
        
        'Si hubo errores salto el empleado e informo
        If hubo_error Then
            Flog.writeline Espacios(Tabulador * 0) & "Se encontraron errores en el empleado: " & rs_Empleado!ternro & ", no se actualizara. Legajo: " & rs_Empleado!empleg
            GoTo prox_emp
        End If
        Flog.writeline Espacios(Tabulador * 0) & "Busco los datos de el ternro: " & rs_Empleado!ternro & ". Legajo: " & rs_Empleado!empleg
        'Inserto o actualizo la tabla de empleado waldbott
        '---------------------------------------------------------------------------------------------------------
        StrSql = " SELECT empleado FROM empleado " & _
                 " WHERE empleado = " & legajoWaldbott
        OpenRecordsetExt StrSql, rs_access, ExtConnOra
        If Not rs_access.EOF Then
            Flog.writeline Espacios(Tabulador * 0) & "Se encontro el empleado en Waldbott, legajo: " & legajoWaldbott & "."
            'Actualizar el empleado
            StrSql = " UPDATE empleado SET "
            StrSql = StrSql & " nombre = '" & apellido_Nombre & "', "
            StrSql = StrSql & " fecnacto = TO_DATE (" & cambiaFecha(fecha_Nacimiento) & ", 'DD-MM-YY') , "
            StrSql = StrSql & " nacionalid = '" & nacionalidad & "', "
            StrSql = StrSql & " estcivil = '" & estado_civil & "', "
            StrSql = StrSql & " sexo = '" & sexo & "', "
            StrSql = StrSql & " fecingreso = TO_DATE (" & cambiaFecha(fecha_Alta) & ", 'DD-MM-YY') , "
            StrSql = StrSql & " fecantig1 = TO_DATE (" & cambiaFecha(fecha_ingreso) & ", 'DD-MM-YY') , "
            StrSql = StrSql & " estudios = '" & estudios & "', "
            StrSql = StrSql & " tipodocto = '" & tipo_Documento & "', "
            StrSql = StrSql & " nrodocto = " & nro_doc & ", "
            StrSql = StrSql & " nrocuil = '" & nro_cuil & "', "
            StrSql = StrSql & " calle = '" & calle & "', "
            StrSql = StrSql & " numero = '" & nro_Direccion & "', "
            StrSql = StrSql & " piso = '" & piso & "', "
            StrSql = StrSql & " depto = '" & ofic_depto & "', "
            StrSql = StrSql & " codpostal = '" & codigo_Postal & "', "
            StrSql = StrSql & " localidad = '" & localidad & "', "
            StrSql = StrSql & " provincia = '" & provincia & "', "
            StrSql = StrSql & " telefono = '" & telefono & "', "
            StrSql = StrSql & " convcol = '" & convenio_CodExt & "', "
            StrSql = StrSql & " categoria = '" & categoria & "', "
            StrSql = StrSql & " codsindic = '" & sindicato & "', "
            StrSql = StrSql & " tarea = '" & puesto & "', "
            'StrSql = StrSql & " imputac = '" & centroDeCosto & "', "
            StrSql = StrSql & " codjubilac = '" & cajaJubilacion & "', "
            StrSql = StrSql & " codosocial = '" & obraSocialLey & "', "
            StrSql = StrSql & " planosoc = '" & planOSLey & "', "
            'StrSql = StrSql & " contrato = '" & contrato & "', "
            StrSql = StrSql & " fecvtocont = TO_DATE (" & cambiaFecha(fecha_VtoContrato) & ", 'DD-MM-YY') , "
            StrSql = StrSql & " horario = '" & regimenHorario & "', "
            StrSql = StrSql & " codbanco = '" & banco & "',"
            StrSql = StrSql & " ctabanco = '" & nroCuenta & "',"
            StrSql = StrSql & " empleoant = '" & nroCBU & "',"
            StrSql = StrSql & " actividad = '" & actividad & "', "
            StrSql = StrSql & " rebcontpat = '" & contrato & "',"
            StrSql = StrSql & " CSITREV1 = '" & Aux_Cod_sitr1 & "',"
            StrSql = StrSql & " DIASR1 = '" & Aux_diainisr1 & "',"
            StrSql = StrSql & " CSITREV2 = '" & Aux_Cod_sitr2 & "',"
            StrSql = StrSql & " DIASR2 = '" & Aux_diainisr2 & "',"
            StrSql = StrSql & " CSITREV3 = '" & Aux_Cod_sitr3 & "',"
            StrSql = StrSql & " DIASR3 = '" & Aux_diainisr3 & "',"
            StrSql = StrSql & " Condicion = '" & estado_Empleado & "', "
            StrSql = StrSql & " motegreso = '" & causa_Baja & "', "
            StrSql = StrSql & " fecbaja = TO_DATE (" & cambiaFecha(fecha_Baja) & ", 'DD-MM-YY') , "
            StrSql = StrSql & " sdojornal = " & emp_Remuneracion & ","
            StrSql = StrSql & " imputac = '" & imputacion & "', "
            StrSql = StrSql & " jurisdic = '" & jurisdiccion & "', "
            StrSql = StrSql & " grupo = '" & grupo & "',"
            StrSql = StrSql & " generasijp = 'S' "
            StrSql = StrSql & " WHERE empleado = " & legajoWaldbott
    
            ExtConnOra.Execute StrSql, , adExecuteNoRecords
            
            Flog.writeline Espacios(Tabulador * 0) & "Actualizado el empleado en Waldbott, legajo: " & legajoWaldbott & "."
            Flog.writeline Espacios(Tabulador * 0) & "*************************************************************************************************************"
        Else
            Flog.writeline Espacios(Tabulador * 0) & "No se encontro el empleado en Waldbott, legajo: " & legajoWaldbott & "."
            Dim nuevo_legajo As Long
            rs_access.Close
            StrSql = " SELECT empleado FROM empleado "
            StrSql = StrSql & " WHERE Empleado >= " & centroDeCosto & "0000 And Empleado <= " & centroDeCosto & "9999"
            StrSql = StrSql & " ORDER BY empleado DESC "
            OpenRecordsetExt StrSql, rs_access, ExtConnOra
            If Not rs_access.EOF Then
                nuevo_legajo = CLng(centroDeCosto & Right(rs_access!Empleado, 4)) + 1
            Else
                nuevo_legajo = CLng(centroDeCosto & "0001")
            End If
            
            Flog.writeline Espacios(Tabulador * 0) & "No existe el empleado."
            'El empleado no existe hay que insertarlo
            StrSql = " INSERT INTO empleado (EMPLEADO,NOMBRE,FECNACTO,NACIONALID,ESTCIVIL,SEXO,FECINGRESO,FECANTIG1,ESTUDIOS "
            StrSql = StrSql & " ,TIPODOCTO,NRODOCTO,NROCUIL,CALLE,NUMERO,PISO,DEPTO,CODPOSTAL,LOCALIDAD,PROVINCIA ,TELEFONO, "
            'StrSql = StrSql & " CONVCOL,CATEGORIA,CODSINDIC,TAREA,IMPUTAC,CODJUBILAC,CODOSOCIAL,PLANOSOC ,CONTRATO,FECVTOCONT,"
            StrSql = StrSql & " CONVCOL,CATEGORIA,CODSINDIC,TAREA,IMPUTAC,CODJUBILAC,CODOSOCIAL,PLANOSOC ,FECVTOCONT,"
            StrSql = StrSql & " HORARIO,CODBANCO,CTABANCO,EMPLEOANT,ACTIVIDAD,REBCONTPAT,Condicion,MOTEGRESO,FECBAJA,SDOJORNAL, "
            'StrSql = StrSql & " HORARIO,CODBANCO,CTABANCO,EMPLEOANT,ACTIVIDAD,Condicion,MOTEGRESO,FECBAJA,SDOJORNAL, "
            'version 1.1
            StrSql = StrSql & " CSITREV1,DIASR1,CSITREV2,DIASR2,CSITREV3,DIASR3, "
            StrSql = StrSql & " JURISDIC,GRUPO,generasijp) "
            StrSql = StrSql & " VALUES (" & nuevo_legajo & ",'" & apellido_Nombre & "', TO_DATE (" & cambiaFecha(fecha_Nacimiento) & ", 'DD-MM-YY') , "
            StrSql = StrSql & " '" & nacionalidad & "', '" & estado_civil & "', '" & sexo & "', TO_DATE (" & cambiaFecha(fecha_Alta) & ", 'DD-MM-YY') , "
            StrSql = StrSql & " TO_DATE (" & cambiaFecha(fecha_ingreso) & ", 'DD-MM-YY') , '" & estudios & "', '" & tipo_Documento & "', "
            StrSql = StrSql & "" & nro_doc & ",'" & nro_cuil & "','" & calle & "','" & nro_Direccion & "','" & piso & "','" & ofic_depto & "', "
            StrSql = StrSql & "'" & codigo_Postal & "','" & localidad & "','" & provincia & "','" & telefono & "','" & convenio_CodExt & "', "
            StrSql = StrSql & "'" & categoria & "','" & sindicato & "','" & puesto & "','" & imputacion & "','" & cajaJubilacion & "', "
            'StrSql = StrSql & "'" & obraSocialLey & "','" & planOSLey & "','" & contrato & "',TO_DATE (" & cambiaFecha(fecha_VtoContrato) & ", 'DD-MM-YY') , "
            StrSql = StrSql & "'" & obraSocialLey & "','" & planOSLey & "',TO_DATE (" & cambiaFecha(fecha_VtoContrato) & ", 'DD-MM-YY') , "
            'StrSql = StrSql & "'" & regimenHorario & "','" & banco & "','" & nroCuenta & "','" & nroCBU & "','" & actividad & "','" & sitRev_CodAFIP & "',"
            'version 1.1
            StrSql = StrSql & "'" & regimenHorario & "','" & banco & "','" & nroCuenta & "','" & nroCBU & "','" & actividad & "','" & contrato & "',"

            StrSql = StrSql & "'" & estado_Empleado & "','" & causa_Baja & "',TO_DATE (" & cambiaFecha(fecha_Baja) & ", 'DD-MM-YY') , " & emp_Remuneracion & ", "
            StrSql = StrSql & "'" & Aux_Cod_sitr1 & "','" & Aux_diainisr1 & "','" & Aux_Cod_sitr2 & "','" & Aux_diainisr2 & "','" & Aux_Cod_sitr3 & "','" & Aux_diainisr3 & "',"
            StrSql = StrSql & "'" & jurisdiccion & "','" & grupo & "','S')"
                
            ExtConnOra.Execute StrSql, , adExecuteNoRecords
            Flog.writeline Espacios(Tabulador * 0) & "Se inserto un nuevo empleado en Waldbott, legajo: " & nuevo_legajo & "."
            
            
            '10/12/2012 - 1.02 insercion de legajo en la tablas Horas y Valores
            StrSql = " INSERT INTO HORAS (empleado) VALUES ('" & nuevo_legajo & "')"
            ExtConnOra.Execute StrSql, , adExecuteNoRecords
            Flog.writeline Espacios(Tabulador * 0) & "Se inserto el nuevo: " & nuevo_legajo & ", en la tabla horas."
            
            StrSql = " INSERT INTO VALORES (empleado) VALUES ('" & nuevo_legajo & "')"
            ExtConnOra.Execute StrSql, , adExecuteNoRecords
            Flog.writeline Espacios(Tabulador * 0) & "Se inserto el nuevo: " & nuevo_legajo & ", en la tabla valores."
            'FIN 10/12/2012
            
            
            'Chequeo si tiene algun documento, si lo tiene lo actualizao
            StrSql = " SELECT ternro FROM ter_doc WHERE tidnro = " & tipDocWDT & "AND ternro = " & rs_Empleado!ternro
            OpenRecordset StrSql, rs_aux
            If rs_aux.EOF Then
                'Inserto el documento waldbott para el empleado del lado de RHPRO
                StrSql = " INSERT INTO ter_doc (tidnro,ternro,nrodoc) VALUES "
                StrSql = StrSql & " (" & tipDocWDT & "," & rs_Empleado!ternro & ",'" & nuevo_legajo & "') "
                
            Else
                StrSql = " UPDATE ter_doc SET nrodoc = '" & nuevo_legajo & "'" & _
                         " WHERE tidnro = " & tipDocWDT & " AND ternro = " & rs_Empleado!ternro
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
            Flog.writeline Espacios(Tabulador * 0) & "Se actualizo el legajo waldbott en rhpro, legajo: " & nuevo_legajo & "."
            Flog.writeline Espacios(Tabulador * 0) & "*************************************************************************************************************"
        End If
        If rs_access.State = adStateOpen Then rs_access.Close
        
        '-------------------------------------------------------------------------------------------------
        'Busqueda de familiares
        '-------------------------------------------------------------------------------------------------
        Dim apellido_Nombre_Fam As String
        Dim nacionalidad_Fam As String
        Dim fecha_Nacimiento_Fam As String
        Dim estado_civil_fam As String
        Dim sexo_Fam As String
        Dim parentesco As String
        Dim incapacitado As String
        Dim estudia_Fam As String
        Dim nro_doc_fam As String
        Dim tipo_Documento_Fam As String
        Dim obra_social_Fam As String
        Dim salario_Fam As String
        Dim ganancias_Fam As String
        Dim inicio_Fam As String
        Dim vencimiento_Fam As String
        
        StrSql = " SELECT tercero.ternro, tercero.terfecnac, tercero.terape,tercero.terape2, tercero.ternom, tercero.ternom2,parcodext, nacionaldes,tercero.estcivnro "
        StrSql = StrSql & " , tercero.tersex, estcivil.estcivdesext, familiar.faminc, familiar.famestudia, osocial, familiar.famsalario, familiar.famcargaDGI, familiar.famfec, familiar.famfecvto "
        StrSql = StrSql & " From Tercero "
        StrSql = StrSql & " INNER JOIN familiar ON tercero.ternro=familiar.ternro "
        StrSql = StrSql & " INNER JOIN nacionalidad ON nacionalidad.nacionalnro = tercero.nacionalnro "
        StrSql = StrSql & " INNER JOIN estcivil ON estcivil.estcivnro = tercero.estcivnro "
        StrSql = StrSql & " LEFT JOIN parentesco ON familiar.parenro=parentesco.parenro "
        StrSql = StrSql & " WHERE familiar.Empleado = " & rs_Empleado!ternro
        OpenRecordset StrSql, rs_familiar
        
         Do While Not rs_familiar.EOF
            
            hubo_Error_fam = False
            apellido_Nombre_Fam = rs_familiar!terape & IIf(EsNulo(rs_familiar!terape2), " ", " " & rs_familiar!terape2 & " ") & rs_familiar!ternom & IIf(EsNulo(rs_familiar!ternom2), "", " " & rs_familiar!ternom2) 'char 30
            Flog.writeline Espacios(Tabulador * 0) & "Comienza analisis de familiar: " & rs_familiar!ternro & ". Nombre: " & apellido_Nombre_Fam
            Call checkError(apellido_Nombre_Fam, True, 29, bpronro, "nombreFam", arr_Errores)
            'Error Critico
            If arr_Errores(52) <> "" Then
                hubo_Error_fam = True
            End If

            fecha_Nacimiento_Fam = rs_familiar!terfecnac 'date
            Call checkError(fecha_Nacimiento_Fam, False, 10, bpronro, "fecha", arr_Errores)
            
            nacionalidad_Fam = rs_familiar!nacionaldes 'char 1 tabla access
            'Busco la nacionalidad en la tabla access
            '-----------------------------------------------------------------------------------------
            strAccess = " SELECT codigo FROM vw_nacionalidades "
            strAccess = strAccess & " WHERE funcion = '" & nacionalidad_Fam & "'"
            OpenRecordsetExt strAccess, rs_access, ExtConnAccess
            If Not rs_access.EOF Then
                nacionalidad_Fam = rs_access!codigo
                Flog.writeline Espacios(Tabulador * 2) & "Nacionalidad: " & nacionalidad_Fam & " del familiar existente en la tabla Access."
                Call checkError(nacionalidad_Fam, False, 1, bpronro, "nacionalidad", arr_Errores)
            Else
                Flog.writeline Espacios(Tabulador * 2) & "Nacionalidad: " & nacionalidad_Fam & " del familiar no existente en la tabla Access."
                nacionalidad_Fam = ""
            End If
            If rs_access.State = adStateOpen Then rs_access.Close
            
            'Estado civil familiar
            '-----------------------------------------------------------------------------------------
            If EsNulo(rs_familiar!estcivdesext) Then
                'Call errorFam("Estado Civil")
                arr_Errores(56) = "Error"
                hubo_Error_fam = True
            Else
                estado_civil_fam = rs_familiar!estcivdesext
            End If
            Call checkError(estado_civil_fam, True, 1, bpronro, "estadoCivilFam", arr_Errores)
            If arr_Errores(70) <> "" Then
                hubo_Error_fam = True
            End If
            
            sexo_Fam = IIf(rs_familiar!tersex = -1, "M", "F") 'char 1
            If sexo_Fam = "" Then
                'Error Critico
                arr_Errores(57) = "Error"
                hubo_Error_fam = True
            End If
            Call checkError(sexo_Fam, True, 1, bpronro, "sexo", arr_Errores)


            parentesco = rs_familiar!parcodext
            Call checkError(parentesco, True, 1, bpronro, "parentesco", arr_Errores)
            'Error Critico
            If arr_Errores(58) <> "" Then
                hubo_Error_fam = True
            End If

    
            incapacitado = IIf(rs_familiar!faminc = -1, "S", "N")
            Call checkError(incapacitado, True, 1, bpronro, "incapacitado", arr_Errores)
            
            estudia_Fam = IIf(rs_familiar!famestudia = -1, "S", "N")
            Call checkError(estudia_Fam, True, 1, bpronro, "estudia", arr_Errores)
                
            '******************************* Tipo de Documento *******************************************
            StrSql = " SELECT tidnro, nrodoc FROM  ter_doc"
            StrSql = StrSql & " WHERE ternro = " & rs_familiar!ternro & " AND tidnro <= 11 "
            OpenRecordset StrSql, rs_aux
            If Not rs_aux.EOF Then
                Select Case CInt(rs_aux!tidnro)
                    Case 1:
                        tipo_Documento_Fam = "DN"
                    Case 2:
                        tipo_Documento_Fam = "LE"
                    Case 3:
                        tipo_Documento_Fam = "LC"
                    Case 4:
                        tipo_Documento_Fam = "CI"
                    Case 5:
                        tipo_Documento_Fam = "PA"
                    Case 11:
                        tipo_Documento_Fam = "CM"
                        
                    Case Else:
                        Flog.writeline Espacios(Tabulador * 1) & "Error no existe documento para el ternro: " & rs_familiar!ternro & "."
                End Select
                nro_doc_fam = rs_aux!nrodoc
                Call checkError(nro_doc_fam, True, 9, bpronro, "nrodoc", arr_Errores)
            Else
                Flog.writeline Espacios(Tabulador * 1) & "Error no hay documento cargado para el ternro: " & rs_familiar!ternro & "."
                GoTo prox_fam
            End If
            If rs_aux.State = adStateOpen Then rs_aux.Close
            Call checkError(tipo_Documento, True, 2, bpronro, "tipoDocumento", arr_Errores)
            
            obra_social_Fam = IIf(EsNulo(rs_familiar!osocial), "N", "S")
            Call checkError(obra_social_Fam, False, 1, bpronro, "osocial", arr_Errores)
            
            salario_Fam = IIf(rs_familiar!famsalario = -1, "S", "N")
            Call checkError(salario_Fam, False, 1, bpronro, "salario", arr_Errores)
            
            ganancias_Fam = IIf(rs_familiar!famcargaDGI = -1, "S", "N")
            Call checkError(ganancias_Fam, False, 1, bpronro, "ganancias", arr_Errores)
            
            inicio_Fam = IIf(EsNulo(rs_familiar!famfec), "", rs_familiar!famfec)
            Call checkError(inicio_Fam, False, 8, bpronro, "fecha", arr_Errores) 'Si la columna 33 es verdadero obligatorio = true
            
            vencimiento_Fam = IIf(EsNulo(rs_familiar!famfecvto), "", rs_familiar!famfecvto)
            Call checkError(vencimiento_Fam, False, 10, bpronro, "venciminetofam", arr_Errores) 'Si la columna 33 es verdadero obligatorio = true
    
            If hubo_Error_fam Then
                hubo_error = True
                Flog.writeline Espacios(Tabulador * 1) & "Hay errores en el Familiar: " & apellido_Nombre_Fam & ", no se actualizara."
                GoTo prox_fam
            End If
            
            'Chequeo existe el familiar
            StrSql = " SELECT empleado, nrohijo FROM fliares "
            StrSql = StrSql & " WHERE nombre = '" & apellido_Nombre_Fam & "'"
            StrSql = StrSql & " AND empleado = " & legajoWaldbott
            OpenRecordsetExt StrSql, rs_access, ExtConnOra
            'actualizo los familiares
            '-----------------------------------------------------------------------------------------------
           If Not rs_access.EOF Then
                 'Call MyBeginTransExt(ExtConnOra)

                Flog.writeline Espacios(Tabulador * 1) & "Actualizando Familiar: " & apellido_Nombre_Fam & "."
                StrSql = "UPDATE fliares SET "
                StrSql = StrSql & " nombre =  '" & apellido_Nombre_Fam & "',"
                StrSql = StrSql & " fecnacto =  TO_DATE (" & cambiaFecha(fecha_Nacimiento_Fam) & ", 'DD-MM-YY') ,"
                StrSql = StrSql & " nacional =  '" & nacionalidad_Fam & "',"
                StrSql = StrSql & " estcivil =  '" & estado_civil_fam & "',"
                StrSql = StrSql & " sexo =  '" & sexo_Fam & "',"
                StrSql = StrSql & " hijoincap = '" & incapacitado & "',"
                StrSql = StrSql & " codfliar = '" & parentesco & "',"
                StrSql = StrSql & " escolarid =  '" & estudia_Fam & "',"
                StrSql = StrSql & " tipodocto =  '" & tipo_Documento_Fam & "',"
                StrSql = StrSql & " nrodocto =  " & nro_doc_fam & ","
                StrSql = StrSql & " osocial =  '" & obra_social_Fam & "',"
                StrSql = StrSql & " asigfliar =  '" & salario_Fam & "',"
                StrSql = StrSql & " ganancias =  '" & ganancias_Fam & "',"
                StrSql = StrSql & " fecalta =  TO_DATE (" & cambiaFecha(inicio_Fam) & ", 'DD-MM-YY') ,"
                StrSql = StrSql & " fecbaja =  TO_DATE (" & cambiaFecha(vencimiento_Fam) & ", 'DD-MM-YY') "
                StrSql = StrSql & " WHERE Empleado = " & legajoWaldbott & " AND nombre =  '" & apellido_Nombre_Fam & "'"
                ExtConnOra.Execute StrSql, , adExecuteNoRecords
                'Call MyCommitTransExt(ExtConnOra)
          Else
          
            Dim nuevo_hijo As Long
            rs_access.Close
            StrSql = " SELECT nrohijo FROM fliares WHERE empleado = " & legajoWaldbott & " ORDER BY to_number(nrohijo) DESC "
            OpenRecordsetExt StrSql, rs_access, ExtConnOra
            If Not rs_access.EOF Then
                nuevo_hijo = CLng(rs_access!nrohijo) + 1
            Else
                nuevo_hijo = 1
            End If
            rs_access.Close
            Flog.writeline Espacios(Tabulador * 0) & "El Familiar: " & apellido_Nombre_Fam & ", no existe se insertara."
            'Call MyBeginTransExt(ExtConnOra)
            StrSql = "INSERT INTO fliares (empleado,nombre,fecnacto,nrohijo,nacional , estcivil, sexo, codfliar, hijoincap, escolarid, "
            StrSql = StrSql & " tipodocto,nrodocto,osocial,asigfliar,ganancias,fecalta,fecbaja,idfamiliar) VALUES "
            StrSql = StrSql & " ('" & legajoWaldbott & "','" & apellido_Nombre_Fam & "', TO_DATE (" & cambiaFecha(fecha_Nacimiento_Fam) & ", 'DD-MM-YY') ," & nuevo_hijo & ","
            StrSql = StrSql & "'" & nacionalidad_Fam & "','" & estado_civil_fam & "','" & sexo_Fam & "','" & parentesco & "','" & incapacitado & "',"
            StrSql = StrSql & "'" & estudia_Fam & "','" & tipo_Documento_Fam & "'," & nro_doc_fam & ",'" & obra_social_Fam & "',"
            StrSql = StrSql & "'" & salario_Fam & "','" & ganancias_Fam & "', TO_DATE (" & cambiaFecha(inicio_Fam) & ", 'DD-MM-YY') ,"
            StrSql = StrSql & " TO_DATE (" & cambiaFecha(vencimiento_Fam) & ", 'DD-MM-YY'),0)"

            ExtConnOra.Execute StrSql, , adExecuteNoRecords
            'Call MyCommitTransExt(ExtConnOra)
          End If
          'guardo los empleados que tengo que marcar como sincronizados
          
prox_fam:
            rs_familiar.MoveNext
        Loop
        If rs_familiar.State = adStateOpen Then rs_familiar.Close
        
prox_emp:
        If Not hubo_error Then
            empleados_Actualizados = empleados_Actualizados & "," & rs_Empleado!ternro
        Else
            Incompleto = True
            Call insertarError(bpronro, arr_Errores, str_error, empresa, hubo_error)
        End If
                
        rs_Empleado.MoveNext
        
        'Actualizo el progreso
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & " WHERE bpronro = " & bpronro
        objconnProgreso.Execute StrSql, , adExecuteNoRecords

    Loop
    
    If rs_Empleado.State = adStateOpen Then rs_Empleado.Close
    
    indice = indice + 1
    If empresa <> "" And Incompleto Then
        Call crearProcesosMensajeria(str_error, empresa)
    End If
    
    Call importacion(testado)
    Call sincronizar(empleados_Actualizados)

Loop

If rs_Empleado.State = adStateOpen Then rs_Empleado.Close
If rs_familiar.State = adStateOpen Then rs_familiar.Close
If rs_aux.State = adStateOpen Then rs_aux.Close
If rs_access.State = adStateOpen Then rs_access.Close
If rs_Consult.State = adStateOpen Then rs_Consult.Close
Set rs_Consult = Nothing
Set rs_Empleado = Nothing
Set rs_familiar = Nothing
Set rs_aux = Nothing
Set rs_access = Nothing

'Libero las conexiones externo
ExtConnAccess.Close
ExtConnOra.Close
Set ExtConnOra = Nothing
Set ExtConnAccess = Nothing

Exit Sub

E_interfaceWaldbott:
    Flog.writeline "=================================================================="
    Flog.writeline "Procedimiento: interfaceWaldbott"
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


Public Function cambiaFecha(ByVal Fecha As String) As String
' ---------------------------------------------------------------------------------------------
' Descripcion: Setea la fecha al formato de insercion de la base de datos.
' Autor      : Martin Ferraro
' Fecha      : 18/04/2009
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------


    If EsNulo(Fecha) Then
        cambiaFecha = "NULL"
    Else
        cambiaFecha = ConvFecha(Fecha)
    End If

End Function

Public Sub OpenConnExt(strConnectionString As String, ByRef objConn As ADODB.Connection)
' ---------------------------------------------------------------------------------------------
' Descripcion: Abre conexion externa
' Autor      : Martin Ferraro
' Fecha      : 18/04/2009
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

    If objConn.State <> adStateClosed Then objConn.Close
    objConn.CursorLocation = adUseClient
    
    'Indica que desde una transacción se pueden ver cambios que no se han producido en otras transacciones.
    objConn.IsolationLevel = adXactReadUncommitted
    
    objConn.CommandTimeout = 3600 'segundos
    objConn.ConnectionTimeout = 60 'segundos
    objConn.Open strConnectionString
End Sub


Public Sub OpenRecordsetExt(strSQLQuery As String, ByRef objRs As ADODB.Recordset, ByVal objConnE As ADODB.Connection, Optional lockType As LockTypeEnum = adLockReadOnly)
' ---------------------------------------------------------------------------------------------
' Descripcion: Abre recordset de conexion externa
' Autor      : Martin Ferraro
' Fecha      : 18/04/2009
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim pos1 As Long
Dim pos2 As Long
Dim aux As String

    'Abre un recordset con la consulta strSQLQuery
    If objRs.State <> adStateClosed Then
        If objRs.lockType <> adLockReadOnly Then objRs.UpdateBatch
        objRs.Close
    End If
    
    objRs.CacheSize = 500

    objRs.Open strSQLQuery, objConnE, adOpenDynamic, lockType, adCmdText
    
    Cantidad_de_OpenRecordset = Cantidad_de_OpenRecordset + 1
    
End Sub


Public Function BuscarTerceroXMail(ByVal Mail As String) As Long
' ---------------------------------------------------------------------------------------------
' Descripcion: Funcion encargada de dado un mail obtener un tercero.
' Autor      : Martin Ferraro
' Fecha      : 18/04/2009
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim rs_tercero As New ADODB.Recordset

    StrSql = "SELECT tercero.ternro, teremail"
    StrSql = StrSql & " FROM tercero"
    StrSql = StrSql & " INNER JOIN ter_tip ON ter_tip.ternro = tercero.ternro"
    StrSql = StrSql & " AND ter_tip.tipnro = 14"
    StrSql = StrSql & " WHERE teremail = '" & Mail & "'"
    OpenRecordset StrSql, rs_tercero
    
    If rs_tercero.EOF Then
        BuscarTerceroXMail = 0
    Else
        BuscarTerceroXMail = rs_tercero!ternro
    End If
    
    rs_tercero.Close

Set rs_tercero = Nothing

End Function


Public Function BuscarTerceroTempXMail(ByVal Mail As String) As Long
' ---------------------------------------------------------------------------------------------
' Descripcion: Dado un mail obtiene el campo tercerotemp.
' Autor      : Margiotta, Emanuel
' Fecha      : 04/01/2010
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_terceroTemp As New ADODB.Recordset

    'Si LE esta instalado y esta en otra BD copia los datos
    If (Usa_LE) Then
        StrSql = "SELECT pos_postulante.tercerotemp"
        StrSql = StrSql & " FROM tercero"
        StrSql = StrSql & " INNER JOIN ter_tip ON ter_tip.ternro = tercero.ternro"
        StrSql = StrSql & " AND ter_tip.tipnro = 14"
        StrSql = StrSql & " INNER JOIN pos_postulante ON tercero.ternro = pos_postulante.ternro"
        StrSql = StrSql & " WHERE teremail = '" & Mail & "'"
        OpenRecordset StrSql, rs_terceroTemp
    
        If rs_terceroTemp.EOF Then
            BuscarTerceroTempXMail = 0
        Else
            BuscarTerceroTempXMail = rs_terceroTemp!terceroTemp
        End If
    
        rs_terceroTemp.Close

        Set rs_terceroTemp = Nothing
    Else
    
        BuscarTerceroTempXMail = 0
    End If

End Function


Public Sub OpenRecordsetWithConn(strSQLQuery As String, ByRef objRs As ADODB.Recordset, ByRef Conn As ADODB.Connection, Optional lockType As LockTypeEnum = adLockReadOnly)
    'Abre un recordset con la consulta strSQLQuery, usando la conexion Conn
    If objRs.State <> adStateClosed Then
        If objRs.lockType <> adLockReadOnly Then objRs.UpdateBatch
        objRs.Close
    End If
    objRs.CacheSize = 500
    objRs.Open strSQLQuery, Conn, adOpenDynamic, lockType, adCmdText
End Sub

Public Sub buscarEstructura(ternro As Long, tenro As Integer, ByRef estrDescr As String, ByRef retorno As String, tipo_Retorno As String)
'----------------------------------------------------------------------------------------------------------------
'   ternro:         numero de tercero
'   tenro:          tipo de estructura
'   estrDescr:      valor de retorno de la descripcion de la estructura
'   retorno:        valor de retorno dependiendo del tipo de retorno
'   tipo_retorno:   string indicando segun el case que valor devolver
'----------------------------------------------------------------------------------------------------------------
Dim rs_Estructura As New ADODB.Recordset

    'Consulta de la estructura activa del empleado
    StrSql = " SELECT estructura.estrnro, estructura.estrdabr, his_estructura.htetdesde "
    If InStr(tipo_Retorno, "estrcodext") <> 0 Then
        StrSql = StrSql & " , estrcodext "
    End If
    StrSql = StrSql & " From his_estructura "
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
    StrSql = StrSql & " WHERE his_estructura.tenro = " & tenro & " AND (htetdesde <= " & ConvFecha(Now) & " AND (htethasta IS NULL OR htethasta >= " & ConvFecha(Now) & ")) and ternro = " & ternro
    OpenRecordset StrSql, rs_Estructura
    If Not rs_Estructura.EOF Then
        estrDescr = rs_Estructura!estrdabr
        Select Case tipo_Retorno
            Case "estrcodext":
                retorno = rs_Estructura!estrcodext
            Case "estrnro":
                retorno = CStr(rs_Estructura!estrnro)
            Case "estrnro,htetdesde":
                retorno = CStr(rs_Estructura!estrnro) & "," & CStr(rs_Estructura!htetdesde)
            Case "estrnro,htetdesde,estrcodext":
                retorno = CStr(rs_Estructura!estrnro) & "," & CStr(rs_Estructura!htetdesde) & "," & CStr(rs_Estructura!estrcodext)
            
            Case Else
                retorno = ""
        End Select
            
    Else
        estrDescr = ""
        retorno = ""
    End If
    If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
End Sub

Public Sub checkError(ByRef Dato As Variant, obligatorio As Boolean, Longitud As Integer, bpronro As Long, columna As String, ByRef arr_Error() As String)
'----------------------------------------------------------------------------------------------------------------
'   Checkea longitudes, campos obligatorios, en caso de cortar la corta a la longitud deseada, si el dato no se puede cortar se informa error
'   Dato:           Dato a chequear
'   obligatorio:    Verdadero o falso para checkear si el dato es obligatorio o no
'   longitud:       longitud de la cadena del dato de entrada
'   bpronro:        nro de proceso
'----------------------------------------------------------------------------------------------------------------
    If obligatorio Then
            If Len(Dato) < 0 Then
                Flog.writeline Espacios(Tabulador * 2) & "Error: " & columna & " es obligatorio."
            Else
                
                Select Case UCase(columna)
                    Case UCase("legajo"):
                        If Len(Dato) > Longitud Then
                            Flog.writeline Espacios(Tabulador * 2) & "Error: " & columna & " tiene una longitud mayor a: " & Longitud
                            arr_Error(1) = "Error"
                        Else
                            Dato = Left(Dato, Longitud)
                        End If
                    Case UCase("Nombre"):
                        If Len(Dato) < 1 Then
                            arr_Error(2) = "Error"
                        Else
                            Dato = Left(Dato, Longitud)
                        End If
                    
                    Case UCase("NombreFam"):
                        If Len(Dato) < 1 Then
                            arr_Error(52) = "Error"
                        Else
                            Dato = Left(Dato, Longitud)
                        End If
                    Case UCase("FechaNacimiento"):
                        If Len(Dato) = 0 Then
                            Flog.writeline Espacios(Tabulador * 2) & "Error: Formato de " & columna & " no valido."
                            arr_Error(4) = "Error"
                        End If
                                        
                    Case UCase("FechaAlta"):
                        If Len(Dato) = 0 Then
                            Flog.writeline Espacios(Tabulador * 2) & "Error: Formato de " & columna & " no valido."
                            arr_Error(8) = "Error"
                        End If
                                        
                    Case UCase("Fecha"):
                        If Len(Dato) = 0 Then
                            Flog.writeline Espacios(Tabulador * 2) & "Error: Formato de " & columna & " no valido."
                        End If
                    
                    Case UCase("cuil"):
                        If Len(Dato) > Longitud Then
                            Flog.writeline Espacios(Tabulador * 2) & "Error: El " & columna & " tiene una longitud mayor a: " & Longitud
                            arr_Error(13) = "Error"
                        Else
                            Dato = Left(Dato, Longitud)
                        End If
                        
                    Case UCase("estadoCivil"):
                        If Len(Dato) <> Longitud Then
                            Flog.writeline Espacios(Tabulador * 2) & "Error: El " & columna & " tiene una longitud mayor o menor a: " & Longitud
                            arr_Error(6) = "Error"
                        End If
                    
                    
                    Case UCase("estadoCivilFam"):
                        If Len(Dato) <> Longitud Then
                            Flog.writeline Espacios(Tabulador * 2) & "Error: El estado civil del familiar tiene una longitud mayor o menor a: " & Longitud
                            arr_Error(70) = "Error"
                        End If
                    
                    Case UCase("categoria"):
                        If Len(Dato) > Longitud Then
                            Flog.writeline Espacios(Tabulador * 2) & "Error: La " & columna & " tiene una longitud mayor o menor a: " & Longitud
                            arr_Error(23) = "Error"
                        End If
                    
                    Case UCase("Remuneracion"):
                        Dato = Replace(FormatNumber(Dato, 2, 0, 0, False), ",", ".")
                        If Len(Split(Dato, ".")(1)) > 2 Then
                            Flog.writeline Espacios(Tabulador * 2) & "Error: Formato de " & columna & ", se permiten hasta 2 decimales."
                            arr_Error(42) = "Error"
                        Else
                            If Len(Split(Dato, ".")(0)) > 15 Then
                                Flog.writeline Espacios(Tabulador * 2) & "Error: Formato de " & columna & ", se permiten hasta 15 enteros."
                                arr_Error(42) = "Error"
                            Else
                                Dato = Left(Split(Dato, ".")(0), 15) & "." & Left(Split(Dato, ".")(1), 2)
                            End If
                        End If
                                                    
                    Case UCase("estadoEmpleado"):
                        If Len(Dato) <> 1 Then
                            Flog.writeline Espacios(Tabulador * 2) & "Error: El " & columna & " tiene una longitud distinta a: " & Longitud
                            arr_Error(39) = "Error"
                        End If
                    
                    Case UCase("Parentesco"):
                        If Len(Dato) < 1 Then
                            arr_Error(58) = "Error"
                        Else
                            Dato = Left(Dato, Longitud)
                        End If
                   
                   Case UCase("jurisdiccion"):
                        If Len(Dato) < 2 Then
                            arr_Error(71) = "Error"
                        Else
                            Dato = Left(Dato, Longitud)
                        End If
                    
                    Case UCase("grupo"):
                        Dato = Left(Dato, Longitud)
                        
                    Case Else:
                        Dato = Left(Dato, Longitud)
                End Select
                    
            End If
    Else
    
        Select Case UCase(columna)
            Case UCase("Nombre"):
                Dato = Left(Dato, Longitud)
                
            Case UCase("Fecha"):
                Dato = Left(Dato, 6) & Right(Dato, 2)
                
            Case UCase("Remuneracion"):
                Dato = Replace(FormatNumber(Dato, 2, 0, 0, False), ",", ".")
                If Len(Split(Dato, ".")(1)) > 4 Then
                    Flog.writeline Espacios(Tabulador * 2) & "Error: Formato de " & columna & ", se permiten hasta 4 decimales."
                Else
                    If Len(Split(Dato, ".")(0)) > 15 Then
                        Flog.writeline Espacios(Tabulador * 2) & "Error: Formato de " & columna & ", se permiten hasta 15 enteros."
                    Else
                        Dato = Left(Split(Dato, ".")(0), 15) & "." & Left(Split(Dato, ".")(1), 4)
                    End If
                End If
                
            Case Else:
                Dato = Left(Dato, Longitud)
        End Select
    
    End If
End Sub

Public Sub importacion(ByVal testado As Long)
Dim rs_Empleado As New ADODB.Recordset
Dim rs_egresado As New ADODB.Recordset
Dim rs_aux As New ADODB.Recordset
Dim rs_access As New ADODB.Recordset
Dim estrnro_egresado
    StrSql = " SELECT confval FROM confrep WHERE confnrocol = 14 AND repnro = 383 "
    OpenRecordset StrSql, rs_egresado
    If Not rs_egresado.EOF Then
        Flog.writeline Espacios(Tabulador * 0) & "Configuracion de la estructura egresado OK."
        estrnro_egresado = rs_egresado!confval
    Else
        estrnro_egresado = 0
        Flog.writeline Espacios(Tabulador * 0) & "No esta configurada la estructura egresado."
    End If
        
    If rs_egresado.State = adStateOpen Then rs_egresado.Close
    'Busco bajas del lado de waldbott
    '-------------------------------------------------------------------------------------
    StrSql = " SELECT empleado FROM empleado WHERE condicion = '0' "
    
    OpenRecordsetExt StrSql, rs_access, ExtConnOra
    If Not rs_access.EOF Then
        
        Do While Not rs_access.EOF
            'Si el empleado ya tiene el estado inactivo, es que ya esta sincronizada la baja
            StrSql = " SELECT ter_doc.ternro FROM ter_doc " & _
                     " INNER JOIN empleado on ter_doc.ternro = empleado.ternro " & _
                     " WHERE nrodoc = '" & rs_access!Empleado & "' AND empest = -1 "
            OpenRecordset StrSql, rs_Empleado
            Do While Not rs_Empleado.EOF
                Flog.writeline Espacios(Tabulador * 0) & "Actualizando baja para empleado con legajo Waldbott: " & rs_access!Empleado & "."
                MyBeginTrans
                'Cierro la fase del empleado
                StrSql = " UPDATE fases SET "
                StrSql = StrSql & " bajfec = " & cambiaFecha(Now())
                StrSql = StrSql & " WHERE empleado = " & rs_Empleado!ternro & " AND bajfec Is Null "
                objConn.Execute StrSql, , adExecuteNoRecords
                Flog.writeline Espacios(Tabulador * 0) & "Cierro la fase en RHPro del empleado, ternro: " & rs_Empleado!ternro & "."
                'Cierro la estructura estado que tenia
                StrSql = " UPDATE his_estructura SET "
                StrSql = StrSql & " htethasta = " & cambiaFecha(DateAdd("d", -1, Now()))
                StrSql = StrSql & " WHERE ternro = " & rs_Empleado!ternro & " AND tenro = " & testado & " AND htetdesde Is Null "
                objConn.Execute StrSql, , adExecuteNoRecords
                Flog.writeline Espacios(Tabulador * 0) & "Cierro la estructura de tipo estado en RHPro del empleado, ternro: " & rs_Empleado!ternro & "."
                'Chequeo que la estructura no este cargada
                StrSql = " SELECT ternro FROM his_estructura WHERE estrnro = " & estrnro_egresado & " AND htetdesde <= " & cambiaFecha(Now())
                OpenRecordset StrSql, rs_aux
                If rs_aux.EOF Then
                    'Cargo la estructura egresado al empleado
                    StrSql = " INSERT INTO his_estructura (tenro, ternro, estrnro, htetdesde) VALUES "
                    StrSql = StrSql & "( " & testado & "," & rs_Empleado!ternro & "," & estrnro_egresado & "," & cambiaFecha(Now()) & ")"
                    objConn.Execute StrSql, , adExecuteNoRecords
                    Flog.writeline Espacios(Tabulador * 0) & "Inserto la estructura de tipo estado (egresado) en RHPro del empleado, ternro: " & rs_Empleado!ternro & "."
                Else
                    Flog.writeline Espacios(Tabulador * 0) & "El empleado ya tiene la estructura egresado."
                End If
                'Cierro la fase del empleado
                'StrSql = " UPDATE fases SET "
                'StrSql = StrSql & " bajfec = " & cambiaFecha(Now())
                'StrSql = StrSql & " WHERE empleado = " & rs_Empleado!ternro & " AND bajfec Is Null "
                'objConn.Execute StrSql, , adExecuteNoRecords
                
                'Pongo el empleado en estado inactivo
                StrSql = " UPDATE empleado SET "
                StrSql = StrSql & " empest = 0 "
                StrSql = StrSql & " WHERE ternro = " & rs_Empleado!ternro
                objConn.Execute StrSql, , adExecuteNoRecords
                Flog.writeline Espacios(Tabulador * 0) & "Actualizo el estado del empleado en 0 (inactivo) en RHPro, ternro: " & rs_Empleado!ternro & "."
                MyCommitTrans
                
                rs_Empleado.MoveNext
            Loop
            If rs_Empleado.State = adStateOpen Then rs_Empleado.Close
            rs_access.MoveNext
        Loop
    Else
        Flog.writeline Espacios(Tabulador * 0) & "No Hay empleados dado de baja."
    End If
    If rs_access.State = adStateOpen Then rs_access.Close
End Sub

Public Sub insertarError(ByVal bpronro As Long, ByRef arr_Error() As String, ByRef str_error As String, ByVal empresa As String, ByVal hubo_error As Boolean)
Dim rs_error As New ADODB.Recordset
Dim indice As Integer
Dim encabezado As String
    
        
    str_error = str_error & "<TABLE width='50%' border='1' bordercolor='#333333' cellpadding='0' cellspacing='0'>" & vbCrLf
    If UBound(arr_Error) > 0 Then
        StrSql = " SELECT empleg FROM empleado WHERE ternro = " & arr_Error(0)
        OpenRecordset StrSql, rs_error
        If Not rs_error.EOF Then
            encabezado = "Legajo Empleado: " & rs_error!empleg
        End If
        If rs_error.State = adStateOpen Then rs_error.Close
        str_error = str_error & "<tr><TH style='background-color:#d13528;color:#FFFFFF;' colspan='2' align='center'><b>" & empresa & "</b></TH></tr>" & vbCrLf
        str_error = str_error & "<tr><TH style='background-color:#d13528;color:#FFFFFF;' align='center'><b>Tipo Error</b></TH>" & vbCrLf
        str_error = str_error & "<TH style='background-color:#d13528;color:#FFFFFF;' align='center'><b>" & encabezado & "</b></TH></tr>" & vbCrLf
    End If
    If hubo_error Then
        For indice = 0 To UBound(arr_Error)
            If arr_Error(indice) <> "" Then
    
                Select Case indice
                    
                    Case 1: 'legajo waldbott
                        str_error = str_error & "<tr><td>Sistema RHPro</td>" & vbCrLf
                        str_error = str_error & "<td>Numero de Documento Waldbott incorrecto</td></tr>" & vbCrLf
                    Case 2: 'Nombre
                        str_error = str_error & "<tr><td>Sistema RHPro</td>" & vbCrLf
                        str_error = str_error & "<td>Nombre no existe</td></tr>" & vbCrLf
                    Case 4: 'Fecha nacimiento
                        str_error = str_error & "<tr><td>Sistema RHPro</td>" & vbCrLf
                        str_error = str_error & "<td>No hay fecha de nacimiento</td></tr>" & vbCrLf
                    Case 5: 'Nacionalidad
                        str_error = str_error & "<tr><td>Sistema RHPro</td>" & vbCrLf
                        str_error = str_error & "<td>No hay nacionalidad cargada</td></tr>" & vbCrLf
                    Case 6: 'Estado Civil
                        str_error = str_error & "<tr><td>Sistema RHPro</td>" & vbCrLf
                        str_error = str_error & "<td>Codigo Estado Civil invalido</td></tr>" & vbCrLf
                    Case 8: 'Fecha de Alta
                        str_error = str_error & "<tr><td>Sistema RHPro</td>" & vbCrLf
                        str_error = str_error & "<td>No hay fecha de alta cargada</td></tr>" & vbCrLf
                    Case 11: 'Tipo de documento
                        str_error = str_error & "<tr><td>Sistema RHPro</td>" & vbCrLf
                        str_error = str_error & "<td>Tipo de Documento no valido</td></tr>" & vbCrLf
                    Case 12: 'Nro de documento
                        str_error = str_error & "<tr><td>Sistema RHPro</td>" & vbCrLf
                        str_error = str_error & "<td>Numero de Documento no valido</td></tr>" & vbCrLf
                    Case 13: 'Nro CUIL
                        str_error = str_error & "<tr><td>Sistema RHPro</td>" & vbCrLf
                        str_error = str_error & "<td>Numero de Cuil no valido</td></tr>" & vbCrLf
                    Case 20: 'Provincia
                        str_error = str_error & "<tr><td>Base access</td>" & vbCrLf
                        str_error = str_error & "<td>Provincia no existente</td></tr>" & vbCrLf
                    Case 23: 'Categoria
                        str_error = str_error & "<tr><td>Base access</td>" & vbCrLf
                        str_error = str_error & "<td>Categoria o Codigo invalido</td></tr>" & vbCrLf
                    Case 24: 'Sindicato
                        str_error = str_error & "<tr><td>Base access</td>" & vbCrLf
                        str_error = str_error & "<td>Sindicato no existente</td></tr>" & vbCrLf
                    Case 26: 'Centro de Costo
                        str_error = str_error & "<tr><td>Base access</td>" & vbCrLf
                        str_error = str_error & "<td>Centro de Costo no existente</td></tr>" & vbCrLf
                    Case 27: 'Caja de Jubilacion
                        str_error = str_error & "<tr><td>Base Access</td>" & vbCrLf
                        str_error = str_error & "<td>Caja de jubilacion no existente</td></tr>" & vbCrLf
                    Case 28: 'Obra social
                        str_error = str_error & "<tr><td>Base Access</td>" & vbCrLf
                        str_error = str_error & "<td>Obra social no existente</td></tr>" & vbCrLf
                    Case 37: 'Actividad
                        str_error = str_error & "<tr><td>Base Access</td>" & vbCrLf
                        str_error = str_error & "<td>Actividad no existente</td></tr>" & vbCrLf
                    'Case 38: 'Situacion de Revista - Comentado version 1.1
                    '    str_error = str_error & "<tr><td>Sistema RHPro</td>" & vbCrLf
                    '    str_error = str_error & "<td>Codigo de Situacion inexistente</td></tr>" & vbCrLf
                    Case 39: 'Estado el Empleado
                        str_error = str_error & "<tr><td>Sistema RHPro</td>" & vbCrLf
                        str_error = str_error & "<td>Estado no valido o inexistente</td></tr>" & vbCrLf
                    Case 42: 'Remuneracion
                        str_error = str_error & "<tr><td>Sistema RHPro</td>" & vbCrLf
                        str_error = str_error & "<td>Remuneracion no valido</td></tr>" & vbCrLf
                    Case 52: 'Nombre Familiar
                        str_error = str_error & "<tr><td>Sistema RHPro</td>" & vbCrLf
                        str_error = str_error & "<td>Nombre no existente</td></tr>" & vbCrLf
                    Case 56: 'Estado Civil Familiar
                        str_error = str_error & "<tr><td>Sistema RHPro</td>" & vbCrLf
                        str_error = str_error & "<td>Estado Civil no valido</td></tr>" & vbCrLf
                    Case 57: 'Sexo Familiar
                        str_error = str_error & "<tr><td>Sistema RHPro</td>" & vbCrLf
                        str_error = str_error & "<td>Sexo familiar no valido</td></tr>" & vbCrLf
                    Case 58: 'Parentesco Familiar
                        str_error = str_error & "<tr><td>Sistema RHPro</td>" & vbCrLf
                        str_error = str_error & "<td>Parentesco familiar no valido</td></tr>" & vbCrLf
                    Case 70: 'Estado Civil Familiar
                        str_error = str_error & "<tr><td>Sistema RHPro</td>" & vbCrLf
                        str_error = str_error & "<td>Estado Civil del familiar no valido o inexistente</td></tr>" & vbCrLf
                    Case 71: 'Juridisdiccion
                        str_error = str_error & "<tr><td>Sistema RHPro</td>" & vbCrLf
                        str_error = str_error & "<td>Jurisdiccion no valido o inexistente</td></tr>" & vbCrLf
                End Select
            End If
        Next
    Else
        str_error = str_error & "<tr><td colspan='2' align='center'><b>No hay Errores</b></td></tr>" & vbCrLf
    End If
    str_error = str_error & "</table>" & vbCrLf
    str_error = str_error & " <table><tr><td>&nbsp;</td></tr><tr><td>&nbsp;</td></tr></table> "
    
    If rs_error.State = adStateOpen Then rs_error.Close
End Sub

Sub crearProcesosMensajeria(ByRef str_error As String, ByVal empresa As String)

Dim objRs As New ADODB.Recordset
Dim fs2, MsgFile
Dim titulo As String
Dim bpronroMail As Long
Dim mails As String
Dim notiNro As Long
Dim mailFileName As String
Dim mailFile


    ' Directorio Salidas
    StrSql = "SELECT sis_dirsalidas FROM sistema"
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        dirsalidas = objRs!sis_dirsalidas & "\attach"
        Flog.writeline "Directorio de Salidas: " & dirsalidas
    Else
        Flog.writeline "No se encuentra configurado sis_dirsalidas"
        Exit Sub
    End If
    If objRs.State = adStateOpen Then objRs.Close

    'Busco el codigo de la notificacion
    StrSql = "SELECT confval FROM confrep "
    StrSql = StrSql & " WHERE conftipo = 'TN' AND repnro = 383"
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        notiNro = objRs!confval
    Else
        Flog.writeline Espacios(Tabulador * 0) & "No esta configurado el tipo de alerta, para el envio de mail."
        notiNro = 0
    End If
    If objRs.State = adStateOpen Then objRs.Close


    'FGZ - 04/09/2006 - Saco esto y lo pongo afuera
    StrSql = "insert into batch_proceso "
    StrSql = StrSql & "(btprcnro, bprcfecha, iduser, bprchora, bprcfecdesde, bprcfechasta, bprcparam, "
    StrSql = StrSql & "bprcestado, bprcprogreso, bprcfecfin, bprchorafin, bprctiempo, empnro, bprcempleados) "
    StrSql = StrSql & "values (25," & ConvFecha(Date) & ",'" & usuario & "','" & FormatDateTime(Time, 4) & ":00'"
    StrSql = StrSql & ",null,null,'1','Pendiente',null,null,null,null,0,null)"
    
    objConn.Execute StrSql, , adExecuteNoRecords
    
    bpronroMail = getLastIdentity(objConn, "batch_proceso")


    '--------------------------------------------------
    'Busco todos los usuarios a los cuales les tengo que enviar los mails
    StrSql = "SELECT usremail FROM user_per "
    StrSql = StrSql & "inner join noti_usuario on user_per.iduser = noti_usuario.iduser "
    StrSql = StrSql & "where notinro = " & notiNro
    OpenRecordset StrSql, objRs
    mails = ""
    Do Until objRs.EOF
        If Not IsNull(objRs!usremail) Then
            If Len(objRs!usremail) > 0 Then
                mails = mails & objRs!usremail & ";"
            End If
        End If
        objRs.MoveNext
    Loop
    
    mailFileName = dirsalidas & "\msg_" & bpronroMail & "_interface_waldbott_" & Replace(FormatDateTime(Date, 2), "/", "-") & "_" & Replace(FormatDateTime(Time, 4), ":", "-") & "-" & String(2 - Len(Second(Now)), "0") & Second(Now)
    Set fs2 = CreateObject("Scripting.FileSystemObject")
    Set mailFile = fs2.CreateTextFile(mailFileName & ".html", True)
    
    mailFile.writeline "<html><head>"
    mailFile.writeline "<title> Interface Waldbott - RHPro &reg; </title></head><body>"
    'mailFile.writeline "<h4>Errores Detectados</h4>"
    mailFile.writeline str_error
    mailFile.writeline "</body></html>"
    mailFile.Close
    '--------------------------------------------------


    Set fs2 = CreateObject("Scripting.FileSystemObject")
    
    'FGZ - 05/09/2006
    'Los nombres de los archivos para los mails de esta alerta empiezan con el bpronro de este proceso
    Set MsgFile = fs2.CreateTextFile(mailFileName & ".msg", True)
    
    MsgFile.writeline "[MailMessage]"
    MsgFile.writeline "FromName=RHPro - Errores waldbott"
    MsgFile.writeline "Subject=Informe Errores waldbott " & empresa
    MsgFile.writeline "Body1="
    If Len(mailFileName) > 0 Then
       MsgFile.writeline "Attachment=" & mailFileName & ".html"
    Else
       MsgFile.writeline "Attachment="
    End If
    
    MsgFile.writeline "Recipients=" & mails
    
    If objRs.State = adStateOpen Then objRs.Close
    
    StrSql = "select cfgemailfrom,cfgemailhost,cfgemailport,cfgemailuser,cfgemailpassword from conf_email where cfgemailest = -1"
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        MsgFile.writeline "FromAddress=" & objRs!cfgemailfrom
        MsgFile.writeline "Host=" & objRs!cfgemailhost
        MsgFile.writeline "Port=" & objRs!cfgemailport
        MsgFile.writeline "User=" & objRs!cfgemailuser
        MsgFile.writeline "Password=" & objRs!cfgemailpassword
    Else
        Flog.writeline "No existen datos configurados para el envio de emails, o no existe configuracion activa"
        Exit Sub
    End If
    MsgFile.writeline "CCO="
    MsgFile.writeline "CC="
    MsgFile.writeline "HTMLBody="
    MsgFile.writeline "HTMLMailHeader="

    
    If objRs.State = adStateOpen Then objRs.Close

End Sub

Public Sub sincronizar(ByVal empleados As String)
    
    Flog.writeline "Entrando a sincronizar."
    MyBeginTrans
    
    'Marco a los empleados como sincronizados
    StrSql = " UPDATE empsinc SET "
    StrSql = StrSql & " essinc = -1 "
    StrSql = StrSql & " where esternro in (" & empleados & ") "
    
    objConn.Execute StrSql, , adExecuteNoRecords
    MyCommitTrans
    Flog.writeline "Empleados sincronizados: (" & empleados & ") (Ternros)"
End Sub


Public Sub sit_rev(ByVal ternro As Long, ByVal empresa As Long, ByRef Aux_Cod_sitr, ByRef Aux_Cod_sitr1, ByRef Aux_Cod_sitr2, ByRef Aux_Cod_sitr3, ByRef Aux_diainisr1, ByRef Aux_diainisr2, ByRef Aux_diainisr3)
Dim rs_fases As New ADODB.Recordset
Dim rs_aux As New ADODB.Recordset
Dim rs_HisEstructuras As New ADODB.Recordset
Dim rs_Estructura As New ADODB.Recordset
Dim rs_Estr_cod As New ADODB.Recordset
Dim Fecha_Inicio_Fase As String
Dim Fecha_Fin_Fase As String
Dim sr_cod(3)
Dim sr_dia(3)
Dim es_ultimo As Boolean
Dim p As Integer

    StrSql = " SELECT * FROM fases WHERE empleado = " & ternro & _
             " ORDER BY altfec"
    
   OpenRecordset StrSql, rs_fases
   '---------inicio ver 1.34
   ' Creo el Select para verificar si el empleado tiene un Contrato de tipo 11 (Afip)
   StrSql = " SELECT htetdesde,htethasta FROM his_estructura " & _
            " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro " & _
            " INNER JOIN estr_cod ON estr_cod.estrnro = estructura.estrnro " & _
            " INNER JOIN tipocont ON tipocont.estrnro = his_estructura.estrnro " & _
            " WHERE estructura.tenro = 18 AND nrocod = '11' and his_estructura.ternro = " & ternro & _
            " ORDER BY htetdesde DESC "
   
   OpenRecordset StrSql, rs_aux
   'ver 1.34
   'Si dicho empleado tiene dicha estructura, asigno como fecha de inicio de Fase, el inicio del periodo, más allá de que el empleado tenga una fase abierta en dicho mes
   'Según resolución AFIP para el SICORE
   If Not rs_aux.EOF Then
        Fecha_Inicio_Fase = rs_aux!htetdesde
        Fecha_Fin_Fase = IIf(EsNulo(rs_aux!htethasta), Date, rs_aux!htethasta)
   Else
        If rs_fases.RecordCount > 1 Then rs_fases.MoveFirst
        If rs_fases.RecordCount > 0 Then
                Flog.writeline Espacios(Tabulador * 2) & "Comienza proceso de comparación de fechas de fases con las del período del SIJP"
                Do While Not rs_fases.EOF
                    'Asigno la fecha de alta de la fase si es mayor a la del periodo
                    Fecha_Inicio_Fase = IIf(rs_fases!altfec < Date, rs_fases!altfec, Date)
                    Flog.writeline Espacios(Tabulador * 2) & "Asigno a fecha de inicio de fase el valor " & Fecha_Inicio_Fase
                    If Not EsNulo(rs_fases!bajfec) Then
                        Flog.writeline Espacios(Tabulador * 2) & "El valor de fecha de baja no es nulo"
                        'Asigno la fecha de baja de la fase si es menor a la del periodo
                        Fecha_Fin_Fase = IIf(rs_fases!bajfec < Date, rs_fases!bajfec, Date)
                        Flog.writeline Espacios(Tabulador * 2) & "Asigno a fecha de fin de fase el valor " & Fecha_Fin_Fase
                    Else
                        Fecha_Fin_Fase = Date
                        Flog.writeline Espacios(Tabulador * 2) & "El valor de fecha de baja es nulo"
                        Flog.writeline Espacios(Tabulador * 2) & "El valor asignado a la Fecha de Fin de Fase es " & Fecha_Fin_Fase
                    End If
 
                    rs_fases.MoveNext
            Loop
        End If
    End If '------ fin ver 1.34

    Flog.writeline Espacios(Tabulador * 2) & "Fecha de alta de fase definitiva para calculo de sijp: " & Fecha_Inicio_Fase
    Flog.writeline Espacios(Tabulador * 2) & "Fecha de baja de fase definitiva para calculo de sijp: " & Fecha_Fin_Fase
    
    
    Flog.writeline Espacios(Tabulador * 2) & "Buscar Situacion de Revista Actual"
    ' ----------------------------------------------------------------
    ' FGZ - 28/04/2004 - Codigos
    'Buscar Situacion de Revista Actual
    StrSql = " SELECT * FROM his_estructura " & _
             " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro " & _
             " INNER JOIN estr_cod ON estr_cod.estrnro = his_estructura.estrnro AND tcodnro = 1 " & _
             " WHERE his_estructura.ternro = " & ternro & " AND " & _
             " his_estructura.tenro = 30 AND " & _
             " (his_estructura.htetdesde <= " & ConvFecha(Fecha_Fin_Fase) & ") AND " & _
             " ((" & ConvFecha(Fecha_Inicio_Fase) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null)) " & _
             " ORDER BY his_estructura.htetdesde "
             
    OpenRecordset StrSql, rs_Estructura
    'Inicializo las variables
    Aux_Cod_sitr1 = ""
    Aux_diainisr1 = ""
    Aux_Cod_sitr2 = ""
    Aux_diainisr2 = ""
    Aux_Cod_sitr3 = ""
    Aux_diainisr3 = ""
    
    Select Case rs_Estructura.RecordCount
                Case 0:
                    Flog.writeline Espacios(Tabulador * 2) & "No hay situaciones de revista."
                
                Case 1:
                    'Aux_Cod_sitr1 = rs_Estructura!nrocod
                    StrSql = "SELECT * FROM estr_cod WHERE estrnro =" & rs_Estructura!estrnro
                    StrSql = StrSql & " AND tcodnro = 1"
                    OpenRecordset StrSql, rs_Estr_cod
                    If Not rs_Estr_cod.EOF Then
                        Aux_Cod_sitr1 = Left(CStr(rs_Estr_cod!nrocod), 2)
                    Else
                        Flog.writeline Espacios(Tabulador * 2) & "No se encontró el codigo interno para la Situacion de Revista"
                        Aux_Cod_sitr1 = 1
                    End If
                    
                    If CDate(rs_Estructura!htetdesde) < CDate(Fecha_Inicio_Fase) Then
                        Aux_diainisr1 = 1
                    Else
                        Aux_diainisr1 = CStr(Day(rs_Estructura!htetdesde))
                    End If
                    'FGZ - 08/07/2005
                    If CInt(Aux_diainisr1) > CInt(Day(Fecha_Fin_Fase)) Then
                        Aux_diainisr1 = CStr(Day(Fecha_Fin_Fase))
                    End If
                    
                    'Aux_Cod_sitr = Aux_Cod_sitr1
                    Flog.writeline Espacios(Tabulador * 2) & "hay 1 situaciones de revista"
                Case 2:
                    'Primer situacion
                    'Aux_Cod_sitr1 = rs_Estructura!nrocod
                    StrSql = "SELECT * FROM estr_cod WHERE estrnro =" & rs_Estructura!estrnro
                    StrSql = StrSql & " AND tcodnro = 1"
                    OpenRecordset StrSql, rs_Estr_cod
                    If Not rs_Estr_cod.EOF Then
                        Aux_Cod_sitr1 = Left(CStr(rs_Estr_cod!nrocod), 2)
                    Else
                        Flog.writeline Espacios(Tabulador * 2) & "No se encontró el codigo interno para la Situacion de Revista"
                        Aux_Cod_sitr1 = 1
                    End If
                    
                    If CDate(rs_Estructura!htetdesde) < CDate(Fecha_Inicio_Fase) Then
                        Aux_diainisr1 = 1
                    Else
                        Aux_diainisr1 = CStr(Day(rs_Estructura!htetdesde))
                    End If
                    'FGZ - 08/07/2005
                    If CInt(Aux_diainisr1) > CInt(Day(Fecha_Fin_Fase)) Then
                        Aux_diainisr1 = CStr(Day(Fecha_Fin_Fase))
                    End If
                    
                    'siguiente situacion
                    rs_Estructura.MoveNext
                    StrSql = "SELECT * FROM estr_cod WHERE estrnro =" & rs_Estructura!estrnro
                    StrSql = StrSql & " AND tcodnro = 1"
                    OpenRecordset StrSql, rs_Estr_cod
                    If Aux_Cod_sitr1 <> Left(CStr(rs_Estr_cod!nrocod), 2) Then ' Agregado ver. 1.31
                            If Not rs_Estr_cod.EOF Then
                                Aux_Cod_sitr2 = Left(CStr(rs_Estr_cod!nrocod), 2)
                            Else
                                Flog.writeline Espacios(Tabulador * 2) & "No se encontró el codigo interno para la Situacion de Revista"
                                Aux_Cod_sitr2 = 1
                            End If
                            'Aux_Cod_sitr2 = rs_Estructura!nrocod
                            Aux_diainisr2 = CStr(Day(rs_Estructura!htetdesde))
                            'FGZ - 08/07/2005
                            If CInt(Aux_diainisr2) > CInt(Day(Fecha_Fin_Fase)) Then
                                Aux_diainisr2 = CStr(Day(Fecha_Fin_Fase))
                            End If
                            'Aux_Cod_sitr = Aux_Cod_sitr2
                            Flog.writeline Espacios(Tabulador * 2) & "hay 2 situaciones de revista"
                    Else 'FGZ - 11/01/2012 --------------------------------------------
                        'Si es la misma sit de revista ==> le asigno la anterior
                        'Aux_Cod_sitr = Aux_Cod_sitr1
                    End If
                    
                Case 3:
                    'Primer situacion (1)
                    'Aux_Cod_sitr1 = rs_Estructura!nrocod
                    StrSql = "SELECT * FROM estr_cod WHERE estrnro =" & rs_Estructura!estrnro
                    StrSql = StrSql & " AND tcodnro = 1"
                    OpenRecordset StrSql, rs_Estr_cod
                    If Not rs_Estr_cod.EOF Then
                        Aux_Cod_sitr1 = Left(CStr(rs_Estr_cod!nrocod), 2)
                    Else
                        Flog.writeline Espacios(Tabulador * 2) & "No se encontró el codigo interno para la Situacion de Revista"
                        Aux_Cod_sitr1 = 1
                    End If
                    
                    If CDate(rs_Estructura!htetdesde) < CDate(Fecha_Inicio_Fase) Then
                        Aux_diainisr1 = 1
                    Else
                        Aux_diainisr1 = CStr(Day(rs_Estructura!htetdesde))
                    End If
                    'FGZ - 08/07/2005
                    If CInt(Aux_diainisr1) > CInt(Day(Fecha_Fin_Fase)) Then
                        Aux_diainisr1 = CStr(Day(Fecha_Fin_Fase))
                    End If
                    
                    'siguiente situacion (2)
                    rs_Estructura.MoveNext
                    'Aux_Cod_sitr2 = rs_Estructura!nrocod
                    StrSql = "SELECT * FROM estr_cod WHERE estrnro =" & rs_Estructura!estrnro
                    StrSql = StrSql & " AND tcodnro = 1"
                    OpenRecordset StrSql, rs_Estr_cod
                    If Aux_Cod_sitr1 <> Left(CStr(rs_Estr_cod!nrocod), 2) Then 'Agregado ver 1.31
                        If Not rs_Estr_cod.EOF Then
                            Aux_Cod_sitr2 = Left(CStr(rs_Estr_cod!nrocod), 2)
                        Else
                            Flog.writeline Espacios(Tabulador * 2) & "No se encontró el codigo interno para la Situacion de Revista"
                            Aux_Cod_sitr2 = 1
                        End If
                        Aux_diainisr2 = CStr(Day(rs_Estructura!htetdesde))
                        'FGZ - 08/07/2005
                        If CInt(Aux_diainisr2) > CInt(Day(Fecha_Fin_Fase)) Then
                            Aux_diainisr2 = CStr(Day(Fecha_Fin_Fase))
                        End If
                    Else 'FGZ - 11/01/2012 --------------------------------------------
                        'Si es la misma sit de revista ==> le asigno la anterior
                        'Aux_Cod_sitr = Aux_Cod_sitr1
                    End If
                    
                    'siguiente situacion (3)
                    rs_Estructura.MoveNext
                    'Aux_Cod_sitr3 = rs_Estructura!nrocod
                    StrSql = "SELECT * FROM estr_cod WHERE estrnro =" & rs_Estructura!estrnro
                    StrSql = StrSql & " AND tcodnro = 1"
                    OpenRecordset StrSql, rs_Estr_cod
                    If Aux_Cod_sitr2 <> Left(CStr(rs_Estr_cod!nrocod), 2) Then 'Agregado ver 1.31
                        If Not rs_Estr_cod.EOF Then
                            If Aux_Cod_sitr2 <> "" Then 'Agregado ver 1.37 - JAZ
                                Aux_Cod_sitr3 = Left(CStr(rs_Estr_cod!nrocod), 2)
                            Else
                                Aux_Cod_sitr2 = Left(CStr(rs_Estr_cod!nrocod), 2)
                            End If
                        Else
                            Flog.writeline Espacios(Tabulador * 2) & "No se encontró el codigo interno para la Situacion de Revista"
                            Aux_Cod_sitr3 = 1
                        End If
                        If Aux_Cod_sitr3 <> "" Then 'Agregado ver 1.37 - JAZ
                            Aux_diainisr3 = CStr(Day(rs_Estructura!htetdesde))
                        Else
                            Aux_diainisr2 = CStr(Day(rs_Estructura!htetdesde))
                        End If
                        'FGZ - 08/07/2005
                        If Aux_diainisr3 <> "" Then 'Agregado ver 1.37 - JAZ
                            If CInt(Aux_diainisr3) > CInt(Day(Fecha_Fin_Fase)) Then
                                Aux_diainisr3 = CStr(Day(Fecha_Fin_Fase))
                            End If
                            'Aux_Cod_sitr = Aux_Cod_sitr3
                        Else
                            If CInt(Aux_diainisr2) > CInt(Day(Fecha_Fin_Fase)) Then
                                Aux_diainisr2 = CStr(Day(Fecha_Fin_Fase))
                            End If
                            'Aux_Cod_sitr = Aux_Cod_sitr2
                        End If
                        Flog.writeline Espacios(Tabulador * 2) & "hay 3 situaciones de revista"
                    Else 'FGZ - 11/01/2012 --------------------------------------------
                        'Si es la misma sit de revista ==> le asigno la anterior
'                        If Aux_Cod_sitr2 <> "" Then 'Modificado ver 1.36 - JAZ
'                            Aux_Cod_sitr = Aux_Cod_sitr2
'                        Else
'                            Aux_Cod_sitr = Aux_Cod_sitr1
'                        End If
                    End If
                    
                Case Else 'mas de tres situaciones ==> toma las ulwtimas tres pero verifica que no haya situaciones iguales en dif periodos
                     rs_Estructura.MoveLast
                     es_ultimo = False
                     p = 3
                     Do While Not (p = 0 Or rs_Estructura.EOF)
                        If Not rs_Estructura!nrocod Then
                           sr_cod(p) = rs_Estructura!nrocod
                           Flog.writeline Espacios(Tabulador * 2) & "codigo = " & sr_cod(p)
                        Else
                           sr_cod(p) = 1
                        End If
                        rs_Estructura.MovePrevious
                        If Not rs_Estructura.EOF Then
                            If sr_cod(p) = rs_Estructura!nrocod Then
                               Do While es_ultimo = False And Not rs_Estructura.EOF
                                  rs_Estructura.MovePrevious
                                  If Not rs_Estructura.EOF Then 'Agregado ver 1.36 - JAZ
                                    If sr_cod(p) <> rs_Estructura!nrocod Then
                                       es_ultimo = True
                                       rs_Estructura.MoveNext
                                       If CDate(rs_Estructura!htetdesde) < CDate(Fecha_Inicio_Fase) Then
                                          sr_dia(p) = ""
                                       Else
                                          sr_dia(p) = CStr(Day(rs_Estructura!htetdesde))
                                          Flog.writeline Espacios(Tabulador * 2) & "dia por dentro = " & sr_dia(p)
                                       End If
                                       rs_Estructura.MovePrevious
                                       p = p - 1
                                    End If
                                 Else 'Agregado ver 1.36 - JAZ - desde acá
                                    es_ultimo = True
                                    rs_Estructura.MoveNext
                                    If CDate(rs_Estructura!htetdesde) < CDate(Fecha_Inicio_Fase) Then
                                       sr_dia(p) = ""
                                    Else
                                       sr_dia(p) = CStr(Day(rs_Estructura!htetdesde))
                                       Flog.writeline Espacios(Tabulador * 2) & "dia por dentro = " & sr_dia(p)
                                    End If
                                    rs_Estructura.MovePrevious
                                    p = p - 1
                                 End If 'Agregado ver 1.36 - JAZ - hasta acá
                               Loop
                               es_ultimo = False
                            Else
                                rs_Estructura.MoveNext
                                If CDate(rs_Estructura!htetdesde) < CDate(Fecha_Inicio_Fase) Then
                                    sr_dia(p) = ""
                                Else
                                    sr_dia(p) = CStr(Day(rs_Estructura!htetdesde))
                                    Flog.writeline Espacios(Tabulador * 2) & "dia por fuera = " & sr_dia(p)
                                End If
                                p = p - 1
                                rs_Estructura.MovePrevious
                            End If
                        Else
                                rs_Estructura.MoveNext
                                If CDate(rs_Estructura!htetdesde) < CDate(Fecha_Inicio_Fase) Then
                                    sr_dia(p) = ""
                                Else
                                    sr_dia(p) = CStr(Day(rs_Estructura!htetdesde))
                                    Flog.writeline Espacios(Tabulador * 2) & "dia por fuera = " & sr_dia(p)
                                End If
                                p = p - 1
                                rs_Estructura.MovePrevious
                        End If
                     Loop
                     Aux_Cod_sitr3 = sr_cod(3)
                     Aux_Cod_sitr2 = sr_cod(2)
                     Aux_Cod_sitr1 = sr_cod(1)
                     Aux_diainisr3 = sr_dia(3)
                     Aux_diainisr2 = sr_dia(2)
                     Aux_diainisr1 = sr_dia(1)
                     'Aux_Cod_sitr = Aux_Cod_sitr3
                     
                     Flog.writeline Espacios(Tabulador * 2) & "hay + de 3 situaciones de revista"

                End Select
                
                'FGZ - 28/12/2004
                'No puede haber situaciones de revista iguales consecutivas.
                'Antes ese caso, me quedo con la primera de las iguales y consecutivas
                If Aux_Cod_sitr3 = Aux_Cod_sitr2 Then
                    'Elimino la situacion de revista 3
                    Aux_Cod_sitr3 = ""
                    Aux_diainisr3 = ""
                End If
                If Aux_Cod_sitr2 = Aux_Cod_sitr1 Then
                    'Elimino la situacion de revista 2 y la 3 la pongo en la 2
                    Aux_Cod_sitr2 = Aux_Cod_sitr3
                    Aux_diainisr2 = Aux_diainisr3
                    
                    Aux_Cod_sitr3 = ""
                    Aux_diainisr3 = ""
                End If
    
End Sub
