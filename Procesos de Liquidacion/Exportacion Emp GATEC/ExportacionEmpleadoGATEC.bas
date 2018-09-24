Attribute VB_Name = "ExpEmpSIGA"
Option Explicit

'Inicio de versión desde la 1.03 del anterior proceso de exportación empleados SIGA
'Global Const Version = "1.01" 'Zamarbide Juan Alberto - CAS-15555 - TABACAL - Creación nueva vista SQL - GATEC Funcionario
'Global Const FechaModificacion = "26/04/2012" ' Version Inicial adaptando al Nuevo sistema GATEC para Tabacal
'Global Const UltimaModificacion = " "

'Global Const Version = "1.02" 'Gonzalez Nicolás - CAS-15555 - TABACAL - Creación nueva vista SQL - GATEC Funcionario
'Global Const FechaModificacion = "29/05/2012" ' Se corrigió error en SQL
'Global Const UltimaModificacion = " " ' Se agregó control de errores en base a la conexion y sale del proceso en caso de error.

'Global Const Version = "1.03" 'Manterola Maria Magdalena - CAS-15555 - TABACAL - Creación nueva vista SQL - GATEC Funcionario
'Global Const FechaModificacion = "26/06/2012" ' Se corrigieron errores en SQL
'Global Const UltimaModificacion = "Se agregaron chequeos por nulo en ciertos campos de las consultas sql"

'Global Const Version = "1.04" 'Gonzalez Nicolás - CAS-19165 - TABACAL - Proceso Gatec Exportacion empleados en planificador
'Global Const FechaModificacion = "04/05/2013"
'Global Const UltimaModificacion = "Se agregaron parámetros adicionales para el procesamiento planificado"


'Global Const Version = "1.05" 'Gonzalez Nicolás - CAS-19165 - TABACAL - Proceso Gatec Exportacion empleados en planificador
'Global Const FechaModificacion = "13/06/2013"
'Global Const UltimaModificacion = "Se corrigieron errores en query"

'Global Const Version = "1.06" 'Gonzalez Nicolás - CAS-20063 - TABACAL - Exportacion Empledos GATEC - Bug
'Global Const FechaModificacion = "18/06/2013"
'Global Const UltimaModificacion = "Se corrigió validación del campo seg_vence"

Global Const Version = "1.07" 'Ruiz Miriam - CAS-27665 - TABACAL - Proceso planificado - parámetro a pasar
Global Const FechaModificacion = "22/10/2014"
Global Const UltimaModificacion = "Se controla que si la fecha viene vacía desde el palnificador de procesos, coloque la fecha actual"


'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
'Global fs, f
'Global Flog
Global NroProceso As Long

'Global Path As String
Global HuboErrores As Boolean

'NUEVAS
Global EmpErrores As Boolean

Global tenro1 As Integer
Global estrnro1 As Integer
Global tenro2 As Integer
Global estrnro2 As Integer
Global tenro3 As Integer
Global estrnro3 As Integer


Global legdesde As String
Global leghasta As String
Global orden As String
Global ordenado As String
Global empest As String

Global errorConfrep As Boolean

Global TipoCols4(200)
Global CodCols4(200)
Global TipoCols5(200)
Global CodCols5(200)

Global mes1 As String
Global mesPorc1 As String
Global mes2 As String
Global mesPorc2 As String
Global mes3 As String
Global mes4 As String
Global mesPorc3 As String
Global mes5 As String
Global mesPorc4 As String
Global mes6 As String


Global mesPeriodo As Integer
Global anioPeriodo As Integer
Global mesAnterior1 As Integer
Global mesAnterior2 As Integer
Global anioAnterior1 As Integer
Global anioAnterior2 As Integer

Global cantColumna4
Global cantColumna5

Global estrnomb1
Global estrnomb2
Global estrnomb3
Global testrnomb1
Global testrnomb2
Global testrnomb3

Global tprocNro As Integer
Global tprocDesc As String
Global proDesc As String
Global ConcNro As Integer
Global ConcCod As String
Global concabr As String
Global tconnro As Integer
Global tconDesc As String
Global concimp As Integer
Global concpuente As Integer
Global fecEstr As String
Global Formato As Integer
Global Modelo As Long
Global TituloRep As String
Global descDesde
Global descHasta
Global FechaHasta
Global FechaDesde
Global ArchExp
Global UsaEncabezado As Integer
Global Encabezado As Boolean


Private Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial exportacion
' Autor      : FAF
' Fecha      : 27/04/2007
' Ultima Mod :
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim strCmdLine As String
Dim Nombre_Arch As String

Dim rs As New ADODB.Recordset
Dim PID As String
Dim Parametros As String
Dim ArrParametros

Dim empresa As Long
Dim Tenro As Long
Dim Estrnro As Long
Dim Fecha As String
Dim informa_fecha As Date
Dim TipoCall As Integer



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
    
    TiempoInicialProceso = GetTickCount
    HuboErrores = False
    
    Nombre_Arch = PathFLog & "ExpEmpGATEC" & "-" & NroProceso & ".log"
    
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
    StrSql = StrSql & " AND btprcnro = 167"
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
       'Obtengo los parametros del proceso
       Parametros = rs!bprcparam
       ArrParametros = Split(Parametros, "@")
       
       
       tenro1 = CInt(ArrParametros(0))
       estrnro1 = CInt(ArrParametros(1))
       tenro2 = CInt(ArrParametros(2))
       estrnro2 = CInt(ArrParametros(3))
       tenro3 = CInt(ArrParametros(4))
       estrnro3 = CInt(ArrParametros(5))
       If ArrParametros(6) <> "" Then
            informa_fecha = CDate(ArrParametros(6))
       Else
            informa_fecha = Date
       End If
       TipoCall = CInt(ArrParametros(7))
       empresa = CLng(ArrParametros(8))
       If UBound(ArrParametros) > 8 Then
           legdesde = ArrParametros(9)
           leghasta = ArrParametros(10)
           orden = ArrParametros(11)
           ordenado = ArrParametros(12)
           empest = ArrParametros(13)
       End If
       
       '52@1474@50@1470@6@1467@01/07/2013@2@1240@1@9999999999@empleg@Asc@1
       
       Call Escribir_Tabla(informa_fecha, TipoCall, empresa)
    Else
        Flog.writeline Espacios(Tabulador * 0) & "No se encontraron los datos del proceso nro " & NroProceso
    End If
    
    
    StrSql = "DELETE FROM batch_empleado "
    StrSql = StrSql & " WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
    
    
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


Private Sub Escribir_Tabla(ByVal informa_fecha_alta As Date, ByVal tipo_llamada As Integer, ByVal empresa As Integer)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que genera la carga en la base de datos Externa
' Autor      : JAZ
' Fecha      : 24/04/2012
' Ultima Mod : 29/05/2012 - Gonzalez Nicolás - Se agregó estructura. en CARGO PUESTO.
'                                            - Se inicializa var conexion en 0
'              18/06/2013 - Gonzalez Nicolás - Se cambio la forma de validar el campo seg_vence
' ---------------------------------------------------------------------------------------------
'Dim objRs As New ADODB.Recordset
'Dim Nombre_Arch As String
'Dim Directorio As String
'Dim carpeta
'Dim fs1

Dim cantRegistros As Long

Dim NroReporte As Integer
Dim lista_estructuras As String
Dim lista_ccosto As String
Dim tipo_estructura As Integer
Dim fecha_fase As String
Dim Seguir As Boolean
Dim v_fecalta
Dim v_fecbaja
Dim v_fecestruc
Dim v_grupo
Dim v_dni
Dim v_cargo
Dim v_tipoemp
Dim v_contrato
Dim v_codpos
Dim cod_empresa
Dim seg_valor
Dim seg_vence
Dim depto
Dim cdepto
Dim turma
Dim salario
Dim v_domicilio
Dim v_domcomp
Dim v_provincia
Dim v_localidad
Dim v_barrio
Dim v_telefono
Dim Conexion
Dim Texto
Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset
Dim rsConsult As New ADODB.Recordset
Dim conex2 As New ADODB.Connection
Dim c2_Exists As New ADODB.Recordset



    On Error GoTo ME_Local
      
    '--------------------------------Configuracion del Reporte-------------------------------------------------------
    If CStr(informa_fecha_alta) = "" Or IsNull(informa_fecha_alta) Then
        informa_fecha_alta = Date
    End If
    NroReporte = 192
    StrSql = "SELECT * FROM confrep "
    StrSql = StrSql & " WHERE repnro = " & NroReporte
    If rs.State = adStateOpen Then rs.Close
    OpenRecordset StrSql, rs
    lista_estructuras = "0"
    lista_ccosto = "0"
    tipo_estructura = 32
    fecha_fase = "01/01/2000"
    Conexion = 0
    If rs.EOF Then
        Flog.writeline "No se encontró la configuración del Reporte"
        Flog.writeline "   Se deben configurar 3 tipos de columnas:"
        Flog.writeline "     FF : Indica la fecha de baja a partir de la cual se consideran las fases. Default 01/01/2000. Unico"
        Flog.writeline "     TE : Tipo de estructura   No se encontró la configuración del Reporte. Default 32 (Grupo de Liquidacion). Unico"
        Flog.writeline "     EST: Lista de estructuras del tipo anterior. Una o mas"
        Flog.writeline "     CC: Lista de Centros Costos restrictivos del Grupo 13."
        Flog.writeline "     CE: Código de la Empresa (Campo obligatorio de la Tabla FUN_FUNCIONARIO) "
        Flog.writeline "     VAS: Valor del Seguro (Campo NO obligatorio de la Tabla FUN_FUNCIONARIO) "
        Flog.writeline "     VES: Fecha de Vencimiento del Seguro DD/MM/AAAA (Campo NO obligatorio de la Tabla FUN_FUNCIONARIO) "
        Flog.writeline "     DEP: Departamento (Campo NO obligatorio de la Tabla FUN_FUNCIONARIO)"
        Flog.writeline "     CDP: Código de Departamento (Campo NO obligatorio de la Tabla FUN_FUNCIONARIO)"
        Flog.writeline "     TUR: Turma - Equipo de Empleados (Campo NO obligatorio de la Tabla FUN_FUNCIONARIO) "
        Flog.writeline "     SAL: Salario (Campo obligatorio de la Tabla FUN_FUNCIONARIO)"
        Flog.writeline "     CON: Conexión - Hace referencia al cnnro de la tabla conexion que tiene el String de Conexion de la misma  "
        Exit Sub
    Else
        Do Until rs.EOF
            Select Case rs!conftipo
                Case "FF":
                    fecha_fase = rs!confval2
                Case "TE":
                    tipo_estructura = rs!confval
                Case "EST":
                    lista_estructuras = lista_estructuras & "," & rs!confval
                Case "CC":
                    lista_ccosto = lista_ccosto & "," & rs!confval
                Case "CE": 'Código Empresa
                    cod_empresa = rs!confval
                Case "VAS": 'Valor del Seguro
                    seg_valor = rs!confval
                Case "VES": 'Vencimiento del Seguro
                    seg_vence = rs!confval2
                Case "DEP": 'Departamento
                    depto = rs!confval2
                Case "CDP": 'Código de Departamento
                    cdepto = rs!confval2
                Case "TUR": 'Turma (Equipo de Empleados)
                    turma = rs!confval
                Case "SAL": 'Salario
                    salario = rs!confval
                Case "CON": 'Conexion
                    Conexion = rs!confval
            End Select
            
            rs.MoveNext
        Loop
    End If
    rs.Close
    lista_estructuras = lista_estructuras & ","
    lista_ccosto = lista_ccosto & ","
    Flog.writeline "     "
    
    If tipo_llamada = 1 Then
        StrSql = "SELECT empleado.ternro, empleg, empest, empleado.ternom, empleado.terape,empleado.ternom2, empleado.terape2, tercero.terfecnac, empemail "
        StrSql = StrSql & " FROM batch_empleado "
        StrSql = StrSql & " INNER JOIN empleado ON batch_empleado.ternro = empleado.ternro "
        StrSql = StrSql & " INNER JOIN tercero ON empleado.ternro = tercero.ternro "
        StrSql = StrSql & " WHERE bpronro = " & NroProceso
        StrSql = StrSql & " ORDER BY estado "
    Else
        StrSql = "SELECT empleado.ternro, empleg, empest, empleado.ternom, empleado.terape, empleado.ternom2, empleado.terape2,tercero.terfecnac, empemail "
        StrSql = StrSql & " FROM empleado "
        StrSql = StrSql & " INNER JOIN tercero ON empleado.ternro = tercero.ternro "
        StrSql = StrSql & " INNER JOIN his_estructura he ON empleado.ternro = he.ternro  AND he.tenro  = 10 "
        StrSql = StrSql & " AND (he.htetdesde<=" & ConvFecha(informa_fecha_alta) & " AND (he.htethasta is null or he.htethasta>=" & ConvFecha(informa_fecha_alta) & "))"
        StrSql = StrSql & " AND he.estrnro =" & empresa
        
        '---F1
        If tenro1 <> 0 Then
            StrSql = StrSql & " INNER JOIN his_estructura he1 ON empleado.ternro = he1.ternro  AND he1.tenro  = " & tenro1
            StrSql = StrSql & " AND (he1.htetdesde<=" & ConvFecha(informa_fecha_alta) & " AND (he1.htethasta is null or he1.htethasta>=" & ConvFecha(informa_fecha_alta) & "))"
            If estrnro1 <> 0 Then
                StrSql = StrSql & " AND he1.estrnro =" & estrnro1
            End If
        End If
        
        
        If tenro2 <> 0 Then
            StrSql = StrSql & " INNER JOIN his_estructura he2 ON empleado.ternro = he2.ternro  AND he2.tenro  = " & tenro2
            StrSql = StrSql & " AND (he2.htetdesde<=" & ConvFecha(informa_fecha_alta) & " AND (he2.htethasta is null or he2.htethasta>=" & ConvFecha(informa_fecha_alta) & "))"
            If estrnro2 <> 0 Then
                StrSql = StrSql & " AND he2.estrnro =" & estrnro2
            End If
        End If

        If tenro3 <> 0 Then
            StrSql = StrSql & " INNER JOIN his_estructura he3 ON empleado.ternro = he3.ternro AND he3.tenro  = " & tenro3
            StrSql = StrSql & " AND (he3.htetdesde<=" & ConvFecha(informa_fecha_alta) & " AND (he3.htethasta is null or he3.htethasta>=" & ConvFecha(informa_fecha_alta) & "))"
            If estrnro3 <> 0 Then
                StrSql = StrSql & " AND he3.estrnro =" & estrnro3
            End If
        End If

        
        If tipo_llamada = 2 Then
            StrSql = StrSql & " AND empleado.empleg >= " & legdesde
            StrSql = StrSql & " AND empleado.empleg <= " & leghasta
            If empest <> 1 Then
                StrSql = StrSql & " AND empleado.empest = " & empest
            End If
            StrSql = StrSql & " ORDER BY " & orden & " " & ordenado
        End If
    End If
    OpenRecordset StrSql, rs
    Flog.writeline "SQL-> " & StrSql
    'seteo de las variables de progreso
    Progreso = 0
    If Not rs.EOF Then
        cantRegistros = rs.RecordCount
        If cantRegistros = 0 Then
           cantRegistros = 1
           Flog.writeline Espacios(Tabulador * 1) & "No se encontraron datos a Exportar."
        End If
    Else
        cantRegistros = 1
        Flog.writeline Espacios(Tabulador * 1) & "No se encontraron datos a Exportar."
    End If
    IncPorc = (99 / cantRegistros)
    
    '-------------------------------- CONEXION A BD EXTERNA --------------------------------
    StrSql = " SELECT cnnro,cndesc, cnstring FROM conexion WHERE cnnro = " & Conexion
    OpenRecordset StrSql, rs2
    If rs2.EOF Then
        Flog.writeline Espacios(Tabulador * 0) & "No se encuentra la conexion a la BD Externa"
        HuboErrores = True
        Exit Sub
        'Flog.writeline Espacios(Tabulador * 0) & "Se replicaran los Datos en la Tabla"
    Else
        On Error Resume Next
        'Abro la conexion a la BD Externa
        OpenConnection rs2!cnstring, conex2
        
        Flog.writeline Espacios(Tabulador * 0) & "Conexión utilizada para inserción de datos: " & rs2!cndesc
        Flog.writeline ""
        
        'If Err.Number <> 0 Then
        If Error_Encrypt = True Then 'NG - 29/05/2012
            Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion. Debe Configurar bien la conexion a la BD temporal."
            HuboErrores = True
            Exit Sub
        End If
        
        '========================================================================================
        'NG - 29/05/2012 - Valido que exista la tabla GATEC_FUNCIONARIO Para la BD seleccionada
        '========================================================================================
        StrSql = "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES where TABLE_NAME = 'GATEC_FUNCIONARIO'"
        c2_Exists.CacheSize = 500
        c2_Exists.Open StrSql, conex2, adOpenDynamic, adLockReadOnly, adCmdText
        
        
        If c2_Exists.EOF Then
            Flog.writeline "No existe la tabla GATEC_FUNCIONARIO para la BD configurada."
            HuboErrores = True
            Exit Sub
        End If
        c2_Exists.Close
        '========================================================================================
        
    End If
    rs2.Close
    
    
    'Flog.writeline "HuboErrores; " & HuboErrores
    '--------------------------------------- COMIENZO A PROCESAR --------------------------------------
    Do Until rs.EOF
        Seguir = True
        Flog.writeline "Procesando el Empleado: " & rs!empleg & " - " & rs!terape & ", " & rs!ternom
        
        '-------------------------------- FASES -------------------------------------------------------
        ' Fases: SE EXPORTAN SOLO EMPLEADOS ACTIVOS O INACTIVOS CON FECHA DE BAJA MAYOR AL 1/1/2000
        StrSql = "SELECT * FROM fases WHERE empleado =" & rs!Ternro & " ORDER BY altfec DESC"
        OpenRecordset StrSql, rs2
        If rs2.EOF Then
            Flog.writeline "*** Empleado sin fechas de alta/baja (fases), no se informa: " & rs!empleg
            Seguir = False
        End If
        
   
        If Seguir Then
            If Not IsNull(rs2!bajfec) And rs2!bajfec <> "" Then
                If CDate(rs2!bajfec) < CDate(fecha_fase) Then
                    Seguir = False
                End If
            End If
        End If
        
        If Seguir Then
            v_fecalta = rs2!altfec
            v_fecbaja = rs2!bajfec
            v_fecestruc = rs2!bajfec
            If IsNull(rs2!bajfec) Or rs2!bajfec = "" Then
                v_fecestruc = Date
            End If
        End If
        rs2.Close
        '-------------------------------- VER ESTOOOOOOOOOOOOOOOOOOOOOOOO
        'v_grupo = 0
        'v_estrnro = 0
        'If Seguir Then
         '   StrSql = " SELECT his_estructura.estrnro, estructura.estrcodext FROM his_estructura "
         '   StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
'            StrSql = StrSql & " AND htetdesde <= " & ConvFecha(v_fecestruc)
'            StrSql = StrSql & " AND (htethasta Is Null Or htethasta >= " & ConvFecha(v_fecestruc) & ") "
         '   StrSql = StrSql & " AND his_estructura.tenro = " & tipo_estructura & " AND his_estructura.ternro = " & rs!Ternro
'            StrSql = StrSql & " AND his_estructura.estrnro IN (" & lista_estructuras & ")"
         '   StrSql = StrSql & " ORDER BY his_estructura.htetdesde DESC, his_estructura.htethasta "
         '   OpenRecordset StrSql, rs2
         '   If rs2.EOF Then
         '       Seguir = True
         '   Else
         '       v_grupo = rs2!estrcodext
         '       v_estrnro = rs2!Estrnro
         '   End If
        'End If
        '--------------------------------------------------
        If Seguir Then
                '--------------------------- Documento --------------------------------------------
                StrSql = " SELECT ter_doc.nrodoc "
                StrSql = StrSql & " FROM ter_doc "
                StrSql = StrSql & " WHERE ter_doc.tidnro = 10 and ter_doc.ternro= " & rs!Ternro
                OpenRecordset StrSql, rsConsult
                v_dni = ""
                If Not rsConsult.EOF Then
                   v_dni = rsConsult!NroDoc
                Else
                   Flog.writeline "El empleado no posee Número de Documento"
                End If
                rsConsult.Close
                
                '------------------------ Cargo - Puesto -------------------------------------------
                StrSql = " SELECT estrdabr,estructura.estrnro "
                StrSql = StrSql & " From his_estructura"
                StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
                StrSql = StrSql & " AND htetdesde <= " & ConvFecha(informa_fecha_alta)
                StrSql = StrSql & " And (htethasta Is Null Or htethasta >= " & ConvFecha(informa_fecha_alta) & ") "
                StrSql = StrSql & " And his_estructura.tenro = 4 And his_estructura.ternro = " & rs!Ternro
                StrSql = StrSql & " ORDER BY his_estructura.htetdesde DESC"
                OpenRecordset StrSql, rsConsult
                v_cargo = ""
                If Not rsConsult.EOF Then
                   v_cargo = rsConsult!Estrnro
                Else
                    
                End If
                rsConsult.Close
                
                '------------------- DOMICILIOS ---------------------------------------
                StrSql = " SELECT calle, nro, sector, torre, piso, oficdepto, manzana, codigopostal, telnro, provcodext, locdesc, barrio "
                StrSql = StrSql & " From cabdom"
                StrSql = StrSql & " inner join detdom on cabdom.domnro = detdom.domnro "
                StrSql = StrSql & " inner join localidad on localidad.locnro = detdom.locnro "
                StrSql = StrSql & " inner join provincia on provincia.provnro = detdom.provnro "
                StrSql = StrSql & " inner join telefono on telefono.domnro = detdom.domnro "
                StrSql = StrSql & " WHERE cabdom.ternro = " & rs!Ternro
                StrSql = StrSql & " AND domdefault = -1 "
                
                OpenRecordset StrSql, rsConsult
                
                If Not rsConsult.EOF Then
                    v_domicilio = rsConsult!calle & " " & rsConsult!nro
                    If rsConsult!piso <> "" Or Not IsNull(rsConsult!piso) Then
                        v_domcomp = rsConsult!piso & "° " & rsConsult!oficdepto 'sector, torre y manzana?
                    End If
                    v_codpos = rsConsult!codigopostal
                    v_provincia = rsConsult!provcodext
                    v_localidad = rsConsult!locdesc
                    v_barrio = rsConsult!barrio
                    v_telefono = rsConsult!telnro
                End If
                '-------------------------------------------------------------------------
                '------------------------- TIPO DE EMPLEADO ------------------------------
                StrSql = " SELECT estructura.estrnro, estrcodext "
                StrSql = StrSql & " From his_estructura"
                StrSql = StrSql & " inner join estructura on his_estructura.estrnro = estructura.estrnro "
                StrSql = StrSql & " WHERE htetdesde <= " & ConvFecha(informa_fecha_alta)
                StrSql = StrSql & " And (htethasta Is Null Or htethasta >= " & ConvFecha(informa_fecha_alta) & ") "
                StrSql = StrSql & " And his_estructura.tenro = 29 And his_estructura.ternro = " & rs!Ternro
                StrSql = StrSql & " ORDER BY his_estructura.htetdesde DESC"
                OpenRecordset StrSql, rsConsult
                v_tipoemp = ""
                If Not rsConsult.EOF Then
                   v_tipoemp = rsConsult!estrcodext
                End If
                rsConsult.Close
                
                              
                'v_tipemple = ""
                'If v_grupo = "1" Or v_grupo = "2" Or v_grupo = "8" Or v_grupo = "9" Then
                '    v_tipemple = "J"
                'End If
                
                'If v_grupo = "11" Or v_grupo = "12" Then
                '    v_tipemple = "H"
                'End If
                
                If Seguir Then
                    StrSql = "SELECT * FROM GATEC_FUNCIONARIO WHERE FUN_CODIGO = " & rs!empleg
                    c2_Exists.CacheSize = 500
                    c2_Exists.Open StrSql, conex2, adOpenDynamic, adLockReadOnly, adCmdText
                    
                    If Not c2_Exists.EOF Then
                        StrSql = " UPDATE GATEC_FUNCIONARIO "
                        StrSql = StrSql & " SET "
                        StrSql = StrSql & " FUN_CODIGO = '" & rs!empleg & "',"
                        StrSql = StrSql & " COD_EMPR = '" & cod_empresa & "',"
                        StrSql = StrSql & " FUN_NOME = '" & Mid(rs!ternom & " " & rs!ternom2, 1, 40) & "',"
                        StrSql = StrSql & " FUN_ABV = '" & Mid(rs!terape & " " & rs!terape2, 1, 15) & "',"
                        StrSql = StrSql & " FUN_DT_ADM = " & ConvFecha(v_fecalta) & ","
                        If EsNulo(v_fecbaja) Then
                            StrSql = StrSql & " FUN_DT_DESL = null" & ","
                        Else
                            StrSql = StrSql & " FUN_DT_DESL = " & ConvFecha(v_fecbaja) & ","
                        End If
                        StrSql = StrSql & IIf(EsNulo(v_cargo), " FUN_CARGO = " & 0, " FUN_CARGO = '" & v_cargo & "'") & ","
                        StrSql = StrSql & " FUN_ATIVO = '" & Abs(rs!empest) & "',"
                        '*****************************************************************************
                        'MMM - 26/06/2012
                        If EsNulo(seg_valor) Then
                            StrSql = StrSql & " FUN_VLR_SEGURO = null" & ","
                        Else
                            StrSql = StrSql & " FUN_VLR_SEGURO = " & seg_valor & ","
                        End If
                        'StrSql = StrSql & IIf(seg_valor = "" Or IsNull(seg_valor), " FUN_VLR_SEGURO = null", " FUN_VLR_SEGURO = " & seg_valor) & ","
                        
                        'If seg_vence = "" Or ConvFecha(seg_vence) <> "30/12/1899" Then
                        
                        'NG - 18/06/2013 -
                        If EsNulo(seg_vence) = True Then
                            StrSql = StrSql & " FUN_VENC_SEGURO = null" & ","
                        Else
                            StrSql = StrSql & " FUN_VENC_SEGURO = " & ConvFecha(seg_vence) & ","
                        End If
                        'StrSql = StrSql & IIf(ConvFecha(seg_vence) <> "30/12/1899", " FUN_VENC_SEGURO = null", " FUN_VENC_SEGURO = " & ConvFecha(seg_vence)) & ","
                        '*****************************************************************************
                        StrSql = StrSql & IIf(EsNulo(depto), " FUN_COD_ESTAB = null", " FUN_COD_ESTAB = '" & depto & "'") & ","
                        StrSql = StrSql & " FUN_ENDE = '" & v_domicilio & "',"
                        StrSql = StrSql & IIf(EsNulo(v_domcomp), " FUN_COMP = null", " FUN_COMP = '" & v_domcomp & "'") & ","
                        StrSql = StrSql & IIf(EsNulo(v_barrio), " FUN_BAIR = null", " FUN_BAIR = '" & v_barrio & "'") & ","
                        StrSql = StrSql & " FUN_CIDE = '" & v_localidad & "',"
                        StrSql = StrSql & " FUN_UF = '" & v_provincia & "',"
                        StrSql = StrSql & IIf(EsNulo(v_codpos), " FUN_CEP = null", " FUN_CEP = '" & v_codpos & "'") & ","
                        StrSql = StrSql & IIf(EsNulo(v_dni), " FUN_CPF = null", " FUN_CPF = '" & Mid(v_dni, 1, 14) & "'") & ","
                        StrSql = StrSql & IIf(EsNulo(v_telefono), " FUN_FONE = null", " FUN_FONE = '" & v_telefono & "'") & ","
                        StrSql = StrSql & IIf(EsNulo(rs!empemail), " FUN_EMAI = null", " FUN_EMAI = '" & rs!empemail & "'") & ","
                        StrSql = StrSql & IIf(EsNulo(cdepto), " FUN_COD_DEPTO = null", " FUN_COD_DEPTO = '" & cdepto & "'") & ","
                        StrSql = StrSql & IIf(EsNulo(turma), " FUN_COD_EMP_REIT = null", " FUN_COD_EMP_REIT = " & turma) & ","
                        StrSql = StrSql & IIf(EsNulo(v_tipoemp), " FUN_TIPO = null", " FUN_TIPO = '" & Mid(v_tipoemp, 1, 1) & "'") & ","
                        StrSql = StrSql & " FUN_SALARIO = " & Replace(salario, ",", ".")
                        StrSql = StrSql & " WHERE FUN_CODIGO = " & rs!empleg
                        
                        Texto = "Actualizo los datos del Empleado."
                    Else
                        StrSql = " INSERT INTO GATEC_FUNCIONARIO"
                        StrSql = StrSql & "(FUN_CODIGO,COD_EMPR,FUN_NOME,FUN_ABV,FUN_DT_ADM,FUN_DT_DESL,"
                        StrSql = StrSql & "FUN_CARGO,FUN_ATIVO,FUN_REGISTRADE,FUN_VLR_SEGURO,FUN_VENC_SEGURO,"
                        StrSql = StrSql & "FUN_PROPRIO,FUN_COD_ESTAB,FUN_ENDE,FUN_COMP,FUN_BAIR,FUN_CIDE,"
                        StrSql = StrSql & "FUN_UF,FUN_CEP,FUN_CAIX_POST,FUN_CPF,FUN_DDD,FUN_FONE,FUN_EMAI,"
                        StrSql = StrSql & "FUN_EX_MEDICO,FUN_COD_DEPTO,FUN_COD_EMP_REIT,FUN_TIPO,FUN_SALARIO)"
                        StrSql = StrSql & " VALUES ("
                        StrSql = StrSql & "'" & rs!empleg & "',"
                        StrSql = StrSql & "'" & cod_empresa & "',"
                        StrSql = StrSql & "'" & Mid(rs!ternom & " " & rs!ternom2, 1, 40) & "',"
                        StrSql = StrSql & "'" & Mid(rs!terape & " " & rs!terape2, 1, 15) & "',"
                        StrSql = StrSql & ConvFecha(v_fecalta) & ","
                        If EsNulo(v_fecbaja) Then
                            StrSql = StrSql & "null" & ","
                        Else
                            StrSql = StrSql & ConvFecha(v_fecbaja) & ","
                        End If
                        StrSql = StrSql & IIf(EsNulo(v_cargo), 0, "'" & v_cargo & "'") & ","
                        StrSql = StrSql & "'" & Abs(rs!empest) & "',"
                        StrSql = StrSql & "'S',"
                        '*****************************************************************************
                        'MMM - 26/06/2012
                        If EsNulo(seg_valor) Then
                            StrSql = StrSql & " null" & ","
                        Else
                            StrSql = StrSql & seg_valor & ","
                        End If
                        'StrSql = StrSql & IIf(seg_valor = "" Or IsNull(seg_valor), "null", seg_valor) & ","
                        If seg_vence = "" Or ConvFecha(seg_vence) <> "30/12/1899" Then
                            StrSql = StrSql & "null" & ","
                        Else
                            StrSql = StrSql & ConvFecha(seg_vence) & ","
                        End If
                        'StrSql = StrSql & IIf(ConvFecha(seg_vence) <> "30/12/1899", "null", ConvFecha(seg_vence)) & ","
                        '*****************************************************************************
                        StrSql = StrSql & "'S',"
                        StrSql = StrSql & IIf(EsNulo(depto), "null", "'" & depto & "'") & ","
                        StrSql = StrSql & "'" & v_domicilio & "',"
                        StrSql = StrSql & IIf(EsNulo(v_domcomp), "null", "'" & v_domcomp & "'") & ","
                        StrSql = StrSql & IIf(EsNulo(v_barrio), "null", "'" & v_barrio & "'") & ","
                        StrSql = StrSql & "'" & v_localidad & "',"
                        StrSql = StrSql & "'" & v_provincia & "',"
                        StrSql = StrSql & IIf(EsNulo(v_codpos), "null", "'" & v_codpos & "'") & ","
                        StrSql = StrSql & "null" & ","
                        StrSql = StrSql & IIf(EsNulo(v_dni), "null", "'" & Mid(v_dni, 1, 14) & "'") & ","
                        StrSql = StrSql & "null" & ","
                        StrSql = StrSql & IIf(EsNulo(v_telefono), "null", "'" & v_telefono & "'") & ","
                        StrSql = StrSql & IIf(EsNulo(rs!empemail), "null", "'" & rs!empemail & "'") & ","
                        StrSql = StrSql & "'S',"
                        StrSql = StrSql & IIf(EsNulo(cdepto), "null", "'" & cdepto & "'") & ","
                        StrSql = StrSql & IIf(EsNulo(turma), "null", turma) & ","
                        StrSql = StrSql & IIf(EsNulo(v_tipoemp), "null", "'" & Mid(v_tipoemp, 1, 1) & "'") & ","
                        StrSql = StrSql & Replace(salario, ",", ".") & ")"
                        
                        Texto = "Inserto los datos del Empleado."
                    End If
                    
                    '*****************************************************************************
                    'MMM - 26/06/2012
                    Flog.writeline
                    Flog.writeline "CONSULTA A INSERTAR O ACTUALIZAR: " & StrSql
                    Flog.writeline
                    '*****************************************************************************
                    
                    c2_Exists.Close
                    conex2.Execute StrSql, , adExecuteNoRecords
                    
                    On Error GoTo ME_Local
                    
                    'Flog.writeline "Se han registrado los datos del Empleado."
                    Flog.writeline Texto
                  
                End If
                
            End If
            Flog.writeline ""
        
        
        'Actualizo el estado del proceso
        TiempoAcumulado = GetTickCount
        cantRegistros = cantRegistros - 1
        Progreso = Progreso + IncPorc
        
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
        StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'"
        StrSql = StrSql & ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
        
        rs.MoveNext
        
    Loop
    
    
   
Fin:
    'Cierro y libero todo
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    If rs2.State = adStateOpen Then rs2.Close
    Set rs2 = Nothing
    If rsConsult.State = adStateOpen Then rsConsult.Close
    Set rsConsult = Nothing
    conex2.Close
    Set conex2 = Nothing
Exit Sub

ME_Local:
    Flog.writeline
 '   Resume Next
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Description
    Flog.writeline Espacios(Tabulador * 1) & "Ultimo SQL: " & StrSql
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.writeline
End Sub


