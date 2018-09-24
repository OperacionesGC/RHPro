Attribute VB_Name = "MdlinterfaceModelos"
Option Explicit

'Const Version = "1.00"
'Const FechaVersion = "15/01/2013" ' cas 17468 - Deluchi Ezequiel - Version Inicial

Const Version = "1.01"
Const FechaVersion = "24/02/2016" ' LED - CAS-34811 - Monresa - Adec Recibo Digital y ESS - Se agregaron los siguientes modelos:
                                  ' 405 Sincronizacion Recibos ESS, 406 Sincronizacion GTI, 407 Sincronizacion IRPF por periodo


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

    Nombre_Arch = PathFLog & "InterfaceModelos-" & NroProcesoBatch & ".log"
    
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
    On Error GoTo 0
    
    On Error GoTo ME_Main
    
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0 ,bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Flog.writeline "Acutaliza el estado en procesando"
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 384 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        bprcparam = IIf(EsNulo(rs_batch_proceso!bprcparam), "", rs_batch_proceso!bprcparam)
        usuario = rs_batch_proceso!iduser
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        Call interfaceModelo(NroProcesoBatch, bprcparam)
    Else
        Flog.writeline "no se encontró el proceso"
    End If
    
    TiempoFinalProceso = GetTickCount
    Flog.writeline
    Flog.writeline "--------------------------------------------------------------"
    Flog.writeline "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    
    If huboError Then
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

Public Sub interfaceModelo(ByVal bpronro As Long, ByVal Parametros As String)
' Parametros =  1º Manual (-1) | Planificado (0) (todas)
'               2º Si es Exportacion (0) o Importacion (1) o Exportacion WEB (2)
'               3º Empresa si es 0 son todas (solo se usa en exportacion manual, Exportacion o importacion)
'               4º versionOrigen!!versionDestino ej: R3 (3) a R2 (2) (solo en Exportacion o importacion)
'               5º Modelos separados por "!!" el primer valor -1 no se usa (para todos por igual)
'               6º Legajo desde y hasta (solo se usa en exportacion manual, Exportacion o importacion)
'               7º Fecha desde (solo se usa en exportacion web)
'               8º Fecha hasta (solo se usa en exportacion web)
'
Dim arrayParametros
Dim empresa
Dim Origen
Dim destino
Dim modelos
Dim ProcManual As Integer
Dim legDesde As Long
Dim legHasta As Long
Dim fechaDesde As String
Dim fechaHasta As String
    
    arrayParametros = Split(Parametros, "@")
    
    ProcManual = arrayParametros(0)
    empresa = arrayParametros(2)
    Origen = Split(arrayParametros(3), "!!")(0)
    destino = Split(arrayParametros(3), "!!")(1)
    modelos = arrayParametros(4)
    legDesde = Split(arrayParametros(5), "!!")(0)
    legHasta = Split(arrayParametros(5), "!!")(1)
    If UBound(arrayParametros) > 5 Then
        fechaDesde = arrayParametros(6)
        fechaHasta = arrayParametros(7)
    End If
    
    
    'Exportacion
    If CInt(arrayParametros(1)) = 0 Then
        Call exportacion(ProcManual, bpronro, empresa, Origen, destino, modelos, legDesde, legHasta)
    End If
    
    'Importacion
    If CInt(arrayParametros(1)) = 1 Then
        Call importacion(bpronro, empresa, Origen, destino)
    End If
    
    'Exportacion web
    If CInt(arrayParametros(1)) = 2 Then
        Call exportacionWeb(ProcManual, bpronro, modelos, fechaDesde, fechaHasta)
    End If
    
End Sub

Public Sub exportacion(ByVal ProcManual As Long, ByVal bpronro As Long, ByVal empresa As Long, ByVal Origen As Long, ByVal destino As Long, ByVal modelos As String, ByVal legDesde As Long, ByVal legHasta As Long)
 Dim rsEmpresas  As New ADODB.Recordset
 Dim detalle As String
 Dim listModelos
 Dim I As Integer
 Dim Nombre_Arch As String
 Dim rsModelos As New ADODB.Recordset
 Dim rsEmpleados As New ADODB.Recordset
 Dim separador As String
 Dim strLineaModelo As String
 Dim archModelo
 Dim porc As Double
 Dim cantEmpleados As Integer
 Dim cantEmpresa As Integer
 Dim progreso As Double
 Dim estrnro As Long
 
    'hay q levantar los empleados de batch_empleado ya los filtro el asp los desincronizados
    'preguntar si es todas o no y hacer la cosnutla
    On Error GoTo CE
    
    If empresa = 0 Then
        StrSql = "SELECT estrnro,estrdabr FROM estructura WHERE tenro=10"
        OpenRecordset StrSql, rsEmpresas
    Else
        StrSql = "SELECT estrnro,estrdabr FROM estructura WHERE tenro=10 and estrnro= " & empresa
        OpenRecordset StrSql, rsEmpresas
    End If


    'EAM- Obtiene los modelos segun si el importación es Manual o Planificada
    If (CLng(ProcManual) = -1) Then
        If EsNulo(modelos) Then
            Flog.writeline Espacios(Tabulador * 0) & "No se encontraron modelos para importar."
            GoTo CE
        Else
            listModelos = Split(modelos, "!!")
        End If
    Else
        StrSql = "SELECT distinct modelo FROM empsinc " & _
                " INNER JOIN empsinc_det ON empsinc_det.ternro = empsinc.esternro " & _
                " WHERE essinc=0 ORDER BY modelo ASC "
        OpenRecordset StrSql, rsModelos
        modelos = 0
        Do While Not rsModelos.EOF
            modelos = modelos & "!!" & rsModelos!modelo
            rsModelos.MoveNext
        Loop
        listModelos = Split(modelos, "!!")
    End If
    
    cantEmpresa = rsEmpresas.RecordCount
    progreso = 0
    
    estrnro = 0
    Do While Not rsEmpresas.EOF
        
        For I = 1 To UBound(listModelos)
            'EAM- Obtiene los empleados que hay que sincronizar
            If (CLng(ProcManual) = 0) Then
                StrSql = "SELECT * FROM empleado " & _
                        " INNER JOIN his_estructura ON his_estructura.ternro = empleado.ternro " & _
                        " INNER JOIN empsinc ON his_estructura.ternro = empsinc.esternro " & _
                        " INNER JOIN empsinc_det ON empsinc_det.ternro = empsinc.esternro " & _
                        " WHERE his_estructura.tenro = 10 And his_estructura.estrnro = " & rsEmpresas!estrnro & " AND modelo= " & listModelos(I) & _
                        " AND (his_estructura.htetdesde <= " & ConvFecha(Date) & ") " & " AND ((" & ConvFecha(Date) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
                OpenRecordset StrSql, rsEmpleados
            Else
                StrSql = "SELECT empleado.ternro FROM empleado " & _
                        " INNER JOIN his_estructura ON his_estructura.ternro = empleado.ternro " & _
                        " INNER JOIN empsinc ON his_estructura.ternro = empsinc.esternro " & _
                        " INNER JOIN empsinc_det ON empsinc_det.ternro = empsinc.esternro " & _
                        " WHERE his_estructura.tenro = 10 And his_estructura.estrnro = " & rsEmpresas!estrnro & " AND modelo= " & listModelos(I) & _
                        " AND (his_estructura.htetdesde <= " & ConvFecha(Date) & ") " & " AND ((" & ConvFecha(Date) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))" & _
                        " and ((empleado.empleg>= " & legDesde & " ) AND (empleado.empleg<= " & legHasta & "))"
                OpenRecordset StrSql, rsEmpleados
            End If

            porc = CLng((100 / cantEmpresa) / UBound(listModelos))
            cantEmpleados = rsEmpleados.RecordCount
            If cantEmpleados = 0 Then
                progreso = progreso + porc
                cantEmpleados = 1
            Else
                If estrnro <> rsEmpresas!estrnro Then
                    estrnro = rsEmpresas!estrnro
                    'EAM- Crea el archivo de exportación de la empresa
                    Nombre_Arch = PathModelo(1000) & "\Movimiento-" & rsEmpresas!estrdabr & "-" & rsEmpresas!estrnro & "-" & Day(Date) & Month(Date) & Year(Date) & "-" & Hour(Now) & Minute(Now) & ".txt"
                    separador = SeparadorModelo(1000)
                     Set fs = CreateObject("Scripting.FileSystemObject")
                     Set archModelo = fs.CreateTextFile(Nombre_Arch, True)
                End If
            End If
            
            porc = CLng(porc) / CLng(cantEmpleados)
            
            Do While Not rsEmpleados.EOF
                progreso = progreso + porc
                Flog.writeline Espacios(Tabulador * 0) & " Ternro: " & rsEmpleados!ternro
                Select Case listModelos(I)
                    Case 1000:
                        strLineaModelo = expModelo1000(rsEmpleados!ternro, separador)
                        Flog.writeline Espacios(Tabulador * 0) & "entro 1000"
                    Case 1001:
                        strLineaModelo = expModelo1001(rsEmpleados!ternro, separador)
                        Flog.writeline Espacios(Tabulador * 0) & "entro 1001"
                    Case 1002:
                        strLineaModelo = expModelo1002(rsEmpleados!ternro, separador)
                        Flog.writeline Espacios(Tabulador * 0) & "entro 1002"
                    Case 1003:
                        strLineaModelo = expModelo1003(rsEmpleados!ternro, separador)
                        Flog.writeline Espacios(Tabulador * 0) & "entro 1003"
                    Case 1004:
                        strLineaModelo = expModelo1004(rsEmpleados!ternro, separador)
                        Flog.writeline Espacios(Tabulador * 0) & "entro 1004"
                    Case 1005:
                        strLineaModelo = expModelo1005(rsEmpleados!ternro, separador)
                        Flog.writeline Espacios(Tabulador * 0) & "entro 1005"
                End Select
                                
                If Not EsNulo(strLineaModelo) Then
                    archModelo.writeline strLineaModelo
                    strLineaModelo = ""
                End If
                
                StrSql = "UPDATE batch_proceso SET bprcprogreso = " & progreso & " WHERE bpronro = " & bpronro
                objconnProgreso.Execute StrSql, , adExecuteNoRecords
                
                Call sincronizar(rsEmpleados!ternro)
                rsEmpleados.MoveNext
            Loop
            StrSql = "UPDATE batch_proceso SET bprcprogreso = " & progreso & " WHERE bpronro = " & bpronro
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
        Next
        
        progreso = 100
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & progreso & " WHERE bpronro = " & bpronro
                objconnProgreso.Execute StrSql, , adExecuteNoRecords
        rsEmpresas.MoveNext
    Loop
    
    GoTo Procesado
CE:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 0) & "Error: " & Err.Description
    Flog.writeline Espacios(Tabulador * 0) & "SQL: " & StrSql
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline
Procesado:
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 0) & "Los datos fueron Exportados Exitosamente."
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline
End Sub

Public Sub exportacionWeb(ByVal ProcManual As Long, ByVal bpronro As Long, ByVal modelos As String, ByVal fechaDesde As String, ByVal fechaHasta As String)
 Dim rsEmpresas  As New ADODB.Recordset
 Dim detalle As String
 Dim listModelos
 Dim I As Integer
 Dim Nombre_Arch As String
 Dim rsModelos As New ADODB.Recordset
 Dim rsEmpleados As New ADODB.Recordset
 Dim rs_Datos As New ADODB.Recordset
 Dim strConexionExt As String
 Dim separador As String
 Dim strLineaModelo As String
 Dim archModelo
 Dim porc As Double

 Dim cantEmpresa As Integer
 Dim progreso As Double
 Dim estrnro As Long
 Dim cnnro As Long
 
    'hay q levantar los empleados de batch_empleado ya los filtro el asp los desincronizados
    'preguntar si es todas o no y hacer la cosnutla
    On Error GoTo CE
    
    'EAM- Obtiene los modelos segun si el importación es Manual o Planificada
    If EsNulo(modelos) Then
        Flog.writeline Espacios(Tabulador * 0) & "No se encontraron modelos para importar."
        GoTo CE
    Else
        listModelos = Split(modelos, "!!")
    End If
    
    'Levanto el codigo de la conexion para buscar en la tabla
    StrSql = " SELECT confval, confval2 FROM confrep WHERE repnro = 502 AND upper(conftipo) = 'CON'"
    OpenRecordset StrSql, rs_Datos
    If Not rs_Datos.EOF Then
        If Not IsNull(rs_Datos!confval) Then
                cnnro = rs_Datos!confval
        Else
            Flog.writeline Espacios(Tabulador * 0) & "Conexion externa no configurada"
        End If
    Else
        Flog.writeline Espacios(Tabulador * 0) & "Conexion externa no configurada"
    End If
    
    'Abro la conexion externa una sola vez para los modelos
    StrSql = " SELECT cndesc, cnstring FROM conexion WHERE cnnro = " & cnnro
    OpenRecordset StrSql, rs_Datos
    If Not rs_Datos.EOF Then
        Flog.writeline Espacios(Tabulador * 0) & "Conexion externa encontrada"
        strConexionExt = rs_Datos!cnstring
    
        OpenConnection strConexionExt, ExtConn
        If Err.Number <> 0 Or Error_Encrypt Then
            Flog.writeline "Problemas en la conexion Externa: " & rs_Datos!cndesc
            Exit Sub
        End If
        Flog.writeline Espacios(Tabulador * 0) & "Conexion externa satisfactoria: " & rs_Datos!cndesc
    Else
        Flog.writeline Espacios(Tabulador * 0) & "No se configuro correctamente la conexion externa"
        huboError = True
        Exit Sub
    End If
    
    progreso = 0
    porc = CLng(100 / UBound(listModelos))
    
    For I = 1 To UBound(listModelos)
        
        separador = SeparadorModelo(listModelos(I))
        progreso = progreso + porc
        
        Select Case listModelos(I)
            Case 405: 'modelo de recibos de sueldo con aprobacion desde el sistema y conformidad del empleado
                Call expModelo405(progreso, bpronro)
            
            Case 406: 'modelo de vista reducida del tablero de GTI
                Call expModelo406(fechaDesde, fechaHasta, progreso, bpronro)
        
            Case 407: 'modelo de vista reducida de los reportes IRPF por periodo
                Call expModelo407(progreso, bpronro)
        End Select

            
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & progreso & " WHERE bpronro = " & bpronro
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
        
        'Call sincronizar(rsEmpleados!ternro)
        
        'StrSql = "UPDATE batch_proceso SET bprcprogreso = " & progreso & " WHERE bpronro = " & bpronro
        'objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Next
    
    progreso = 100
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & progreso & " WHERE bpronro = " & bpronro
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    
    GoTo Procesado
CE:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 0) & "Error: " & Err.Description
    Flog.writeline Espacios(Tabulador * 0) & "SQL: " & StrSql
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline
Procesado:
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 0) & "Los datos fueron Exportados Exitosamente."
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline
End Sub

Public Sub importacion(ByVal bpronro As Long, ByVal empresa As Long, ByVal Origen As Long, ByVal destino As Long)
Dim directorio As String
Dim CArchivos
Dim archivo
Dim Folder
Dim separador
Dim EncontroAlguno
Dim Path
Dim cantArchivos As Long
Dim porc As Long


    StrSql = "SELECT sis_direntradas FROM sistema WHERE sisnro = 1 "
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        directorio = Trim(objRs!sis_direntradas)
    Else
        Flog.writeline "No esta configurado el directorio de sistema."
        Exit Sub
    End If

    If objRs.State = adStateOpen Then objRs.Close
    
    StrSql = "SELECT * FROM modelo WHERE modnro = 1000 " '& NroModelo
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        directorio = directorio & Trim(objRs!modarchdefault)

        separador = IIf(Not IsNull(objRs!modseparador), objRs!modseparador, ",")
        
        Flog.writeline "Directorio de importación: " & directorio
     Else
        Flog.writeline "No se encontró el modelo para obtener el directorio de lectura."
        Exit Sub
    End If
    If objRs.State = adStateOpen Then objRs.Close
    
    Set fs = CreateObject("Scripting.FileSystemObject")
        
        Path = directorio
        
        Dim fc, F1, s2
        Set Folder = fs.GetFolder(directorio)
        Set CArchivos = Folder.Files
        
        EncontroAlguno = False
        
        cantArchivos = CArchivos.Count
        If cantArchivos = 0 Then
            cantArchivos = 1
        End If
        porc = 100 / cantArchivos
        progreso = 0
        For Each archivo In CArchivos
            progreso = progreso + porc
            EncontroAlguno = True
                Flog.writeline "Procesando archivo " & archivo.Name
                Call LeeArchivo(directorio & "\" & archivo.Name, destino)

            StrSql = "UPDATE batch_proceso SET bprcprogreso = " & progreso & " WHERE bpronro = " & bpronro
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
        Next
        If Not EncontroAlguno Then
            Flog.writeline "No se encontró ningun archivo."
            progreso = 100
            StrSql = "UPDATE batch_proceso SET bprcprogreso = " & progreso & " WHERE bpronro = " & bpronro
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
            
        End If
End Sub

Private Sub LeeArchivo(ByVal nombreArchivo As String, ByVal destino As Long)
' Descripcion: Lee todos los archivos del directorio y linea por linea
' Autor      : Deluchi Ezequiel
' Fecha      :
' Modificado :

Const ForReading = 1
Const TristateFalse = 0
Dim strLinea As String
Dim Archivo_Aux As String
Dim rs_Lineas As New ADODB.Recordset
Dim rs_Modelo As New ADODB.Recordset
Dim Ciclos As Long
Dim str_error As String
Dim separador As String


    If App.PrevInstance Then
        Flog.writeline Espacios(Tabulador * 0) & "Hay una instancia previa del proceso corriendo - El proceso actual queda en estado Pendiente."
        ProcPendiente = True ' para dejar el proceso pendiente
        Exit Sub
    End If
    
    'Espero hasta que se crea el archivo
    
    On Error Resume Next
    Err.Number = 1
    Ciclos = 0
    Do Until Err.Number = 0
        Err.Number = 0
        Set f = fs.GetFile(nombreArchivo)
        If f.Size = 0 Then
            If Ciclos > 100 Then
                Flog.writeline "No anda el getfile."
            Else
                Err.Number = 1
                Ciclos = Ciclos + 1
            End If
        End If
    Loop
    On Error GoTo 0
    Flog.writeline "Archivo encontrado: " & nombreArchivo
   
   'Abro el archivo
    On Error GoTo CE
    Set f = fs.OpenTextFile(nombreArchivo, ForReading, TristateFalse)
    
    str_error = ""
    separador = SeparadorModelo(1000)
    
    Do While Not f.AtEndOfStream
        strLinea = f.ReadLine
        
        If Trim(strLinea) <> "" Then
            
            'el primer valor de la linea es el modelo
            Select Case Split(strLinea, separador)(0)
                Case 1000:
                    
                    Call import_modelo1000(strLinea, destino, str_error)
                    
                Case 1001:
                    Call import_modelo1001(strLinea, destino, str_error)
                    
                Case 1002:
                    Call import_modelo1002(strLinea, destino, str_error)
                    
                Case 1003:
                    Call import_modelo1003(strLinea, destino, str_error)
                    
                Case 1004:
                    Call import_modelo1004(strLinea, destino, str_error)
                    
                Case 1005:
                    Call import_modelo1005(strLinea, destino, str_error)
                    
            End Select
            
        End If
        
    Loop
        Call crearProcesosMensajeria(str_error, nombreArchivo)

    f.Close
    Flog.writeline
    Flog.writeline "Archivo procesado: " & nombreArchivo & " " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    
    'Borrar el archivo
    fs.DeleteFile nombreArchivo, True
    
Fin:
    If rs_Lineas.State = adStateOpen Then rs_Lineas.Close
    Set rs_Lineas = Nothing
    Exit Sub
    
CE:
    huboError = True
    
    MyRollbackTrans
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 0) & "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline Espacios(Tabulador * 0) & "Error: " & Err.Number
    Flog.writeline Espacios(Tabulador * 0) & "Descripcion: " & Err.Description
    Flog.writeline
    
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "SQL Ejecutado: " & StrSql
        Flog.writeline
    End If
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline
    GoTo Fin
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

Sub crearProcesosMensajeria(ByRef str_error As String, ByVal nombreArchivo As String)

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
    StrSql = StrSql & " WHERE conftipo = 'TN' AND repnro = 390"
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
    
    mailFileName = dirsalidas & "\msg_" & bpronroMail & "_interface_modelos_" & Replace(FormatDateTime(Date, 2), "/", "-") & "_" & Replace(FormatDateTime(Time, 4), ":", "-") & "-" & String(2 - Len(Second(Now)), "0") & Second(Now)
    Set fs2 = CreateObject("Scripting.FileSystemObject")
    Set mailFile = fs2.CreateTextFile(mailFileName & ".html", True)
    
    mailFile.writeline "<html><head>"
    mailFile.writeline "<title> Interface Modelos - RHPro &reg; </title></head><body>"
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
    MsgFile.writeline "FromName=RHPro - detalle importacion"
    MsgFile.writeline "Subject=Informe archivo " & Split(nombreArchivo, "\")(UBound(Split(nombreArchivo, "\")))
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


Public Sub sincronizar(ByVal ternro As Long)
 Dim rsAux As New ADODB.Recordset
 
    MyBeginTrans
    
    StrSql = "SELECT * FROM empsinc_det WHERE ternro= " & ternro
    OpenRecordset StrSql, rsAux
    
    If rsAux.EOF Then
        'Marco a los empleados como sincronizados
        StrSql = " UPDATE empsinc SET essinc = -1 WHERE esternro in (" & ternro & ") "
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
    
    
    MyCommitTrans

End Sub

Public Sub sincronizar_Det(ByVal empleados As String, ByVal nroModelo As Long)

    MyBeginTrans
    
    If nroModelo > 0 Then
        'Borro el detalle del modelo para el empleado
        StrSql = " DELETE FROM empsinc_det WHERE ternro= " & empleados & " AND modelo= " & nroModelo
    Else
        'Marco a los empleados como sincronizados
        Flog.writeline "No Se pudo borrar el detalle. Nro de Modelo " & nroModelo & " incorrectoesta. tercero: " & empleados
    End If
    
    objConn.Execute StrSql, , adExecuteNoRecords
    MyCommitTrans

End Sub

'Obtiene el directorio configurado para el modelo
Public Function PathModelo(nroModelo)
 Dim directorio As String
 Dim rsAux As New ADODB.Recordset
 
    StrSql = "SELECT sis_direntradas FROM sistema WHERE sisnro = 1 "
    OpenRecordset StrSql, rsAux
    
    If Not rsAux.EOF Then
        directorio = Trim(rsAux!sis_direntradas)
    Else
        Flog.writeline "No esta configurado el directorio de sistema."
        Exit Function
    End If
    
    StrSql = "SELECT * FROM modelo WHERE modnro= " & nroModelo
    OpenRecordset StrSql, rsAux
    
    If Not rsAux.EOF Then
        directorio = directorio & Trim(rsAux!modarchdefault)
        Flog.writeline "Directorio del modelo: " & directorio
     Else
        Flog.writeline "No se encontró el modelo para obtener el directorio de lectura."
        Exit Function
    End If
        
    PathModelo = directorio
End Function


Public Function SeparadorModelo(nroModelo)
 Dim separador As String
 Dim rsAux As New ADODB.Recordset

    StrSql = "SELECT modseparador FROM modelo WHERE modnro= " & nroModelo
    OpenRecordset StrSql, rsAux
    
    If Not rsAux.EOF Then
        separador = Trim(rsAux!modseparador)
        Flog.writeline "Separador del modelo: " & separador
     Else
        Flog.writeline "No se encontró el modelo para obtener el directorio de lectura."
        Exit Function
    End If
        
    SeparadorModelo = separador
End Function

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
