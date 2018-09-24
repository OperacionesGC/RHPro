Attribute VB_Name = "ExpRenatre"
Option Explicit

'Global Const Version = "1.01"
'Global Const FechaModificacion = "11/04/2007"
'Global Const UltimaModificacion = " " 'Version Inicial

'Global Const Version = "1.02"
'Global Const FechaModificacion = "04/10/2007"
'Global Const UltimaModificacion = "Se agrergo el mapeo a los estudios" 'Version Inicial
                                'se corrigio la sql de las obras sociales.

Global Const Version = "1.03"
Global Const FechaModificacion = "31/07/2009"
Global Const UltimaModificacion = "MB - Encriptacion de string connection"

'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
Global fs, f
'Global Flog
Global NroProceso As Long

Global Path As String
Global HuboErrores As Boolean

'NUEVAS
Global EmpErrores As Boolean

Global tenro1 As Integer
Global estrnro1 As Integer
Global tenro2 As Integer
Global estrnro2 As Integer
Global tenro3 As Integer
Global estrnro3 As Integer

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
Global concnro As Integer
Global Conccod As String
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
' Fecha      : 11/04/2007
' Ultima Mod :
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim strCmdLine As String
Dim Nombre_Arch As String

Dim rs As New ADODB.Recordset
Dim PID As String
Dim Parametros As String
Dim ArrParametros

Dim Empresa As Long
Dim Tenro As Long
Dim Estrnro As Long
Dim Fecha As String
Dim informa_fecha As Integer


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

    On Error Resume Next
    'Abro la conexion
'    OpenConnection strconexion, objConn
'    If Err.Number <> 0 Then
'        Flog.writeline "Problemas en la conexion"
'        Exit Sub
'    End If
'    OpenConnection strconexion, objconnProgreso
'    If Err.Number <> 0 Then
'        Flog.writeline "Problemas en la conexion"
'        Exit Sub
'    End If
    
    
    On Error GoTo ME_Main
    
    TiempoInicialProceso = GetTickCount
    HuboErrores = False
    
    Nombre_Arch = PathFLog & "ExportacionRenatre" & "-" & NroProceso & ".log"
    
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

    On Error GoTo 0

    On Error GoTo ME_Main
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & ", bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline Espacios(Tabulador * 0) & "Obtengo los datos del proceso"
    
    TiempoAcumulado = GetTickCount
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProceso
    StrSql = StrSql & " AND btprcnro = 164"
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
       'Obtengo los parametros del proceso
       Parametros = rs!bprcparam
       ArrParametros = Split(Parametros, "@")
       
       Empresa = CLng(ArrParametros(0))
       Fecha = ArrParametros(1)
       Tenro = CLng(ArrParametros(2))
       Estrnro = CLng(ArrParametros(3))
       informa_fecha = CInt(ArrParametros(4))
       
       Call Generar_Archivo(Empresa, Fecha, Tenro, Estrnro, informa_fecha)
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


Private Sub Generar_Archivo(ByVal Empnro As Long, ByVal fecha_alta As String, ByVal tipoestr As Long, ByVal estructura As Long, ByVal informa_fecha_alta As Integer)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que genera la exportacion
' Autor      : FAF
' Fecha      : 11/04/2007
' Ultima Mod :
' Descripcion:
' ---------------------------------------------------------------------------------------------
'Dim objRs As New ADODB.Recordset
Dim Nombre_Arch As String
Dim Directorio As String
Dim Carpeta
Dim fs1

Dim Sep As String
Dim SepDec As String

Dim cantRegistros As Long
Dim Empresa As String

Dim Legajo As String
Dim ternro_emp As Long

Dim linea As String
Dim tipo As Integer
Dim f_paren As String
Dim f_apellido As String
Dim f_nombre As String
Dim f_fec_nac As String
Dim f_tersex As String
Dim f_tipodoc As String
Dim f_Nrodoc As String
Dim f_empleo As String
Dim f_asig As String
Dim f_incap As String
Dim f_esco As String
Dim apellido As String
Dim nombre As String
Dim Cuil As String
Dim Calle As String
Dim depto As String
Dim cuartel As String
Dim Nro As String
Dim Piso As String
Dim nrodep As String
Dim Loc As String
Dim prov As String
Dim Telefono As String
Dim Sexo As String
Dim estciv As String
Dim tipodoc As String
Dim Nrodoc As String
Dim fec_nac As String
Dim nacional As String
Dim condi As String
Dim nroUATRE As String
Dim tareas As String
Dim estudio As String
Dim otrosest As String
Dim osocial As String
Dim cajjub As String
Dim afjp As String
Dim sindica As String
Dim nroafil As String
Dim domtrab As String
Dim provtrab As String
Dim loctrab As String
Dim cpostrab As String
Dim tipCodRenatre As Integer

Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
    
    On Error GoTo ME_Local
      
    'Directorio de exportacion
    StrSql = "SELECT sis_dirsalidas FROM sistema WHERE sisnro = 1 "
    If rs.State = adStateOpen Then rs.Close
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        Directorio = Trim(rs!sis_dirsalidas) & "\ExpRenatre"
    End If
    
    Nombre_Arch = Directorio & "\renatre.txt"
    Flog.writeline Espacios(Tabulador * 1) & "Se crea el archivo: " & Nombre_Arch
    Set fs = CreateObject("Scripting.FileSystemObject")
    'desactivo el manejador de errores
    On Error Resume Next
    
    Set Carpeta = fs.getFolder(Directorio)
    If Err.Number <> 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "El directorio " & Directorio & " no existe. Se creará."
        Err.Number = 0
        Set Carpeta = fs.CreateFolder(Directorio)
        
        If Err.Number <> 0 Then
            Flog.writeline Espacios(Tabulador * 1) & "No se puede crear el directorio " & Directorio & ". Verifique los derechos de acceso o puede crearlo."
            HuboErrores = True
            GoTo Fin
        End If
    End If
    
    Set ArchExp = fs.CreateTextFile(Nombre_Arch, True)
    
    If Err.Number <> 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "No se puede crear el archivo " & Nombre_Arch & " en el directorio " & Directorio
        HuboErrores = True
        GoTo Fin
    End If
    
    On Error GoTo ME_Local
    
    'Configuracion del Reporte
        '    NroReporte = 155
        '    StrSql = "SELECT * FROM confrep "
        '    StrSql = StrSql & " WHERE repnro = " & NroReporte
        '    StrSql = StrSql & " AND confnrocol = 1"
        '    If rs.State = adStateOpen Then rs.Close
        '    OpenRecordset StrSql, rs
        '    If rs.EOF Then
        '        Flog.writeline "No se encontró la configuración del Reporte"
        '        Exit Sub
        '    Else
        '        Acumulador = rs!confval
        '    End If
    
    
    '------------------------------------------------------------------
    'Codigo RENATRE
    '------------------------------------------------------------------
    tipCodRenatre = 0
    StrSql = "SELECT tcodnro FROM tipocod "
    StrSql = StrSql & " WHERE tcodnom = 'RENATRE'"
    If rs.State = adStateOpen Then rs.Close
    OpenRecordset StrSql, rs
    If rs.EOF Then
        Flog.writeline "No se encontró la configuración del Codigo para renatre"
        Exit Sub
    Else
        tipCodRenatre = rs!tcodnro
    End If
    rs.Close

    '------------------------------------------------------------------
    'Busco los datos
    '------------------------------------------------------------------
    StrSql = " SELECT distinct empleado.ternro, empleado.terape, empleado.ternom, tersex, estcivdesabr, terfecnac, "
    StrSql = StrSql & " pais.paisdesc, ter_doc.nrodoc, tipodocu.tidsigla   "
    StrSql = StrSql & " FROM empleado "
    StrSql = StrSql & " INNER JOIN tercero ON empleado.ternro = tercero.ternro "
    StrSql = StrSql & " INNER JOIN his_estructura empresa ON empleado.ternro = empresa.ternro "
    StrSql = StrSql & " AND empresa.tenro = 10 AND empresa.htetdesde <= " & ConvFecha(fecha_alta)
    StrSql = StrSql & " AND (empresa.htethasta >= " & ConvFecha(fecha_alta)
    StrSql = StrSql & " OR empresa.htethasta IS NULL) AND empresa.estrnro = " & Empnro
    If tipoestr <> 0 Or estructura <> 0 Then
        StrSql = StrSql & " INNER JOIN his_estructura ON empleado.ternro = his_estructura.ternro "
        StrSql = StrSql & " AND his_estructura.tenro = " & tipoestr & " AND his_estructura.htetdesde <= " & ConvFecha(fecha_alta)
        StrSql = StrSql & " AND (his_estructura.htethasta >= " & ConvFecha(fecha_alta)
        StrSql = StrSql & " OR his_estructura.htethasta IS NULL) "
        If estructura <> 0 Then
            StrSql = StrSql & " AND his_estructura.estrnro = " & estructura
        End If
    End If
    StrSql = StrSql & " LEFT JOIN estcivil ON tercero.estcivnro = estcivil.estcivnro "
    StrSql = StrSql & " LEFT JOIN pais ON tercero.paisnro = pais.paisnro "
    StrSql = StrSql & " LEFT JOIN ter_doc ON tercero.ternro = ter_doc.ternro AND tidnro <= 4 "
    StrSql = StrSql & " LEFT JOIN tipodocu ON tipodocu.tidnro = ter_doc.tidnro "
    StrSql = StrSql & " WHERE empleado.empest = -1"
    If CInt(informa_fecha_alta) = -1 Then
        StrSql = StrSql & " AND empleado.empfaltagr = " & ConvFecha(fecha_alta)
    End If
        
    OpenRecordset StrSql, rs
    
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
    
    Do Until rs.EOF
        tipo = "0"
        apellido = ""
        nombre = ""
        ternro_emp = IIf(IsNull(rs!ternro), "", rs!ternro)
        apellido = IIf(IsNull(rs!terape), "", rs!terape)
        nombre = IIf(IsNull(rs!ternom), "", rs!ternom)
        If CInt(rs!tersex) = -1 Then
            Sexo = "Masc."
        Else
            Sexo = "Fem."
        End If
        estciv = ""
        If rs!estcivdesabr <> "" And Not IsNull(rs!estcivdesabr) Then
            estciv = rs!estcivdesabr
        End If
        fec_nac = ""
        If rs!terfecnac <> "" And Not IsNull(rs!terfecnac) Then
            fec_nac = Day(rs!terfecnac) & "-" & Month(rs!terfecnac) & "-" & Year(rs!terfecnac)
        End If
        nacional = ""
        If rs!paisdesc <> "" And Not IsNull(rs!paisdesc) Then
            nacional = rs!paisdesc
        End If
        tipodoc = ""
        If rs!tidsigla <> "" And Not IsNull(rs!tidsigla) Then
            tipodoc = rs!tidsigla
        End If
        Nrodoc = ""
        If rs!Nrodoc <> "" And Not IsNull(rs!Nrodoc) Then
            Nrodoc = rs!Nrodoc
        End If
        
        ' CUIL
        StrSql = " SELECT nrodoc "
        StrSql = StrSql & " FROM ter_doc WHERE ter_doc.ternro = " & rs!ternro & " AND ter_doc.tidnro = 10"
        OpenRecordset StrSql, rs2
        Cuil = ""
        If Not rs2.EOF And Not IsNull(rs2!Nrodoc) Then
            Cuil = rs2!Nrodoc
        End If
        rs2.Close
        
        ' Domicilio
        StrSql = " SELECT calle, nro, partnom, piso, oficdepto, locdesc, provdesc, telnro "
        StrSql = StrSql & " FROM cabdom "
        StrSql = StrSql & " INNER JOIN detdom ON cabdom.domnro = detdom.domnro "
        StrSql = StrSql & " LEFT JOIN partido ON partido.partnro = detdom.partnro"
        StrSql = StrSql & " LEFT JOIN localidad ON localidad.locnro = detdom.locnro"
        StrSql = StrSql & " LEFT JOIN provincia ON provincia.provnro = detdom.provnro"
        StrSql = StrSql & " LEFT JOIN telefono ON telefono.domnro = detdom.domnro AND telefono.teldefault = -1 "
        StrSql = StrSql & " WHERE cabdom.domdefault = -1 AND cabdom.ternro = " & rs!ternro
        OpenRecordset StrSql, rs2
        depto = ""
        Loc = ""
        prov = ""
        Telefono = ""
        Calle = ""
        cuartel = ""
        Piso = ""
        nrodep = ""
        Nro = "1"
        If Not rs2.EOF Then
            Calle = IIf(IsNull(rs2!Calle), "", rs2!Calle)
            If rs2!Nro <> "" And Not IsNull(rs2!Nro) Then
                Nro = rs2!Nro
            End If
            If rs2!partnom <> "" And Not IsNull(rs2!partnom) Then
                depto = rs2!partnom
            End If
            Piso = IIf(IsNull(rs2!Piso), "", rs2!Piso)
            nrodep = IIf(IsNull(rs2!oficdepto), "", rs2!oficdepto)
            If rs2!locdesc <> "" And Not IsNull(rs2!locdesc) Then
                Loc = rs2!locdesc
            End If
            If rs2!provdesc <> "" And Not IsNull(rs2!provdesc) Then
                prov = rs2!provdesc
            End If
            If rs2!telnro <> "" And Not IsNull(rs2!telnro) Then
                Telefono = rs2!telnro
            End If
        End If
        rs2.Close
    
        ' Tipo Contrato
        StrSql = " SELECT tcnro, tcdabr, nrocod "
        StrSql = StrSql & " FROM his_estructura "
        StrSql = StrSql & " INNER JOIN tipocont ON tipocont.estrnro = his_estructura.estrnro"
        StrSql = StrSql & " LEFT JOIN estr_cod ON his_estructura.estrnro = estr_cod.estrnro AND estr_cod.tcodnro = " & tipCodRenatre
        StrSql = StrSql & " WHERE his_estructura.ternro = " & rs!ternro & " AND his_estructura.tenro = 18"
        StrSql = StrSql & " AND his_estructura.htetdesde <= " & ConvFecha(fecha_alta)
        StrSql = StrSql & " AND (his_estructura.htethasta >= " & ConvFecha(fecha_alta) & " OR his_estructura.htethasta IS NULL)"
        OpenRecordset StrSql, rs2
        condi = ""
        If Not rs2.EOF Then
            If IsNull(rs2!Nrocod) Then
                condi = ""
            Else
                condi = CStr(rs2!Nrocod)
            End If
            'condi = IIf(IsNull(rs2!tcdabr), "", rs2!tcdabr)
            'Select Case rs2!tcnro
            '    Case 1: condi = "Permanente"
            '    Case 2: condi = "De Temporada"
            '    Case 3: condi = "No Permanente"
            'End Select
        End If
        rs2.Close
        
        nroUATRE = "541"
        
        ' Categoria
        StrSql = " SELECT catnro, estrdabr, nrocod "
        StrSql = StrSql & " FROM his_estructura "
        StrSql = StrSql & " INNER JOIN categoria ON categoria.estrnro = his_estructura.estrnro"
        StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro"
        StrSql = StrSql & " LEFT JOIN estr_cod ON his_estructura.estrnro = estr_cod.estrnro AND estr_cod.tcodnro = " & tipCodRenatre
        StrSql = StrSql & " WHERE his_estructura.ternro = " & rs!ternro & " AND his_estructura.tenro = 3"
        StrSql = StrSql & " AND his_estructura.htetdesde <= " & ConvFecha(fecha_alta)
        StrSql = StrSql & " AND (his_estructura.htethasta >= " & ConvFecha(fecha_alta) & " OR his_estructura.htethasta IS NULL)"
        OpenRecordset StrSql, rs2
        tareas = ""
        If Not rs2.EOF Then
            If IsNull(rs2!Nrocod) Then
                tareas = ""
            Else
                tareas = CStr(rs2!Nrocod)
            End If
            'tareas = IIf(IsNull(rs2!estrdabr), "", rs2!estrdabr)
            'Select Case rs2!catnro
            '    Case 15: tareas = "Pe¢n General"
            '    Case 30: tareas = "Pe¢n General"
            '    Case 31: tareas = "Conductor Tractorista"
            'End Select
        End If
        rs2.Close
    
        ' Nivel de Estudio
        StrSql = " SELECT nivest.nivnro, nivest.nivdesc "
        StrSql = StrSql & " FROM cap_estformal "
        StrSql = StrSql & " INNER JOIN nivest ON cap_estformal.nivnro = nivest.nivnro"
        StrSql = StrSql & " WHERE cap_estformal.ternro = " & rs!ternro & " AND cap_estformal.capcomp = -1"
        StrSql = StrSql & " ORDER BY nivest.nivdesc"
        OpenRecordset StrSql, rs2
        estudio = "Sin Estudio"
        If Not rs2.EOF Then
            'Prescolar, EGB1, EGB2, EGB3, Polimodal, Terciario, Univestario, No
            
            estudio = IIf(IsNull(rs2!nivdesc), "", rs2!nivdesc)
            Flog.writeline Espacios(Tabulador * 2) & "Nivel de estudio " & rs2!nivnro & " " & rs2!nivdesc
            estudio = Left(calcularMapeo(estudio, 1, "No"), 20)
        Else
            Flog.writeline Espacios(Tabulador * 2) & "No se encontro el nivel de estudio "
            estudio = "No"
'            Select Case rs2!nivnro
'                Case 4: estudio = "Secundario Completo"
'                Case 1: estudio = "Primario Incompleto"
'                Case 2: estudio = "Primario Incompleto"
'                Case 3: estudio = "Primario Completo"
'                Case 5: estudio = "Terciario"
'                Case 6: estudio = "Universitario"
'                Case Else: estudio = "Sin Estudio"
'            End Select
        End If
        rs2.Close
        
        '   FIND FIRST nivest
        '   WHERE nivest.nivnro = capacitacion.nivest NO-LOCK NO-ERROR.
        '   ASSIGN renglon.otrosest = IF AVAIL nivest AND nivest.nivnro <> ? THEN REPLACE(nivest.nivdesc,'"','\"') ELSE "Ningun otro estudio". */
        '   IF AVAIL capacitacion
        '   THEN DO:
        '     Case capacitacion.nivest:
        '      WHEN 4 THEN ASSIGN renglon.otrosest = "Secundario Completo".
        '      WHEN 1 THEN ASSIGN renglon.otrosest = "Primario Incompleto".
        '      WHEN 2 THEN ASSIGN renglon.otrosest = "Primario Incompleto".
        '      WHEN 3 THEN ASSIGN renglon.otrosest = "Primario Completo".
        '      WHEN 5 THEN ASSIGN renglon.otrosest = "Terciario".
        '      WHEN 6 THEN ASSIGN renglon.otrosest = "Universitario".
        '      OTHERWISE ASSIGN renglon.otrosest = "Sin Estudio".
        '     END CASE.
        '   END.
        '   ELSE ASSIGN renglon.otrosest = "Sin Estudio".
        otrosest = "Sin Estudio"
        
        ' Obra Social
        StrSql = " SELECT terrazsoc "
        StrSql = StrSql & " FROM osocial, his_estructura , replica_estr, tercero "
        StrSql = StrSql & " WHERE his_estructura.ternro = " & rs!ternro & " AND his_estructura.tenro = 17"
        StrSql = StrSql & " AND tercero.ternro = osocial.ternro"
        StrSql = StrSql & " AND replica_estr.estrnro = his_estructura.estrnro AND replica_estr.origen = osocial.ternro"
        StrSql = StrSql & " AND his_estructura.htetdesde <= " & ConvFecha(fecha_alta)
        StrSql = StrSql & " AND (his_estructura.htethasta >= " & ConvFecha(fecha_alta) & " OR his_estructura.htethasta IS NULL)"
        OpenRecordset StrSql, rs2
        osocial = ""
        If Not rs2.EOF Then
            osocial = IIf(IsNull(rs2!terrazsoc), "", rs2!terrazsoc)
        Else
            rs2.Close
            StrSql = " SELECT terrazsoc "
            StrSql = StrSql & " FROM osocial, his_estructura , replica_estr, tercero "
            StrSql = StrSql & " WHERE his_estructura.ternro = " & rs!ternro & " AND his_estructura.tenro = 24"
            StrSql = StrSql & " AND replica_estr.estrnro = his_estructura.estrnro AND replica_estr.origen = osocial.ternro"
            StrSql = StrSql & " AND his_estructura.htetdesde <= " & ConvFecha(fecha_alta)
            StrSql = StrSql & " AND (his_estructura.htethasta >= " & ConvFecha(fecha_alta) & " OR his_estructura.htethasta IS NULL)"
            OpenRecordset StrSql, rs2
            If Not rs2.EOF Then
                osocial = IIf(IsNull(rs2!terrazsoc), "", rs2!terrazsoc)
            End If
        End If
        rs2.Close
        
        ' Caja de jubilacion
        StrSql = " SELECT cajjub.ticnro, estructura.estrdabr "
        StrSql = StrSql & " FROM his_estructura "
        StrSql = StrSql & " INNER JOIN cajjub ON cajjub.estrnro = his_estructura.estrnro"
        StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro"
        StrSql = StrSql & " WHERE his_estructura.ternro = " & rs!ternro & " AND his_estructura.tenro = 15"
        StrSql = StrSql & " AND his_estructura.htetdesde <= " & ConvFecha(fecha_alta)
        StrSql = StrSql & " AND (his_estructura.htethasta >= " & ConvFecha(fecha_alta) & " OR his_estructura.htethasta IS NULL)"
        OpenRecordset StrSql, rs2
        cajjub = ""
        afjp = ""
        If Not rs2.EOF Then
            afjp = IIf(IsNull(rs2!estrdabr), "", rs2!estrdabr)
            If CInt(rs2!ticnro) = 1 Then
                cajjub = "1"
            Else
                cajjub = "0"
            End If
        End If
        rs2.Close
    
        ' Sindicato
        StrSql = " SELECT estructura.estrdabr "
        StrSql = StrSql & " FROM his_estructura "
        StrSql = StrSql & " INNER JOIN gremio ON gremio.estrnro = his_estructura.estrnro"
        StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro"
        StrSql = StrSql & " WHERE his_estructura.ternro = " & rs!ternro & " AND his_estructura.tenro = 16"
        StrSql = StrSql & " AND his_estructura.htetdesde <= " & ConvFecha(fecha_alta)
        StrSql = StrSql & " AND (his_estructura.htethasta >= " & ConvFecha(fecha_alta) & " OR his_estructura.htethasta IS NULL)"
        OpenRecordset StrSql, rs2
        sindica = ""
        nroafil = "0"
        If Not rs2.EOF Then
            sindica = IIf(IsNull(rs2!estrdabr), "", rs2!estrdabr)
        End If
        rs2.Close
    
        ' Domicilio Trabajo
        StrSql = " SELECT detdom.calle, detdom.nro, provincia.provdesc, localidad.locdesc, detdom.codigopostal "
        StrSql = StrSql & " FROM his_estructura "
        StrSql = StrSql & " INNER JOIN empresa ON empresa.estrnro = his_estructura.estrnro"
        StrSql = StrSql & " INNER JOIN cabdom ON cabdom.ternro = empresa.ternro "
        StrSql = StrSql & " INNER JOIN detdom ON detdom.domnro = cabdom.domnro "
        StrSql = StrSql & " INNER JOIN localidad ON detdom.locnro = localidad.locnro "
        StrSql = StrSql & " INNER JOIN provincia ON detdom.provnro = provincia.provnro "
        StrSql = StrSql & " WHERE his_estructura.ternro = " & rs!ternro & " AND his_estructura.tenro = 10"
        StrSql = StrSql & " AND his_estructura.htetdesde <= " & ConvFecha(fecha_alta)
        StrSql = StrSql & " AND (his_estructura.htethasta >= " & ConvFecha(fecha_alta) & " OR his_estructura.htethasta IS NULL)"
        OpenRecordset StrSql, rs2
        domtrab = ""
        provtrab = ""
        loctrab = ""
        cpostrab = ""
        If Not rs2.EOF Then
            domtrab = IIf(IsNull(rs2!Calle), "", rs2!Calle) & " " & IIf(IsNull(rs2!Nro), "", rs2!Nro)
            provtrab = IIf(IsNull(rs2!provdesc), "", rs2!provdesc)
            loctrab = IIf(IsNull(rs2!locdesc), "", rs2!locdesc)
            cpostrab = IIf(IsNull(rs2!codigopostal), "", rs2!codigopostal)
        End If
        rs2.Close
            
        'RENGLON DE TRABAJADOR
        'tipo , "Apellidos", "Nombres", "CUIL", "Calle/Ruta", "Depto.", "Cuartel", "Num/Km", "Piso", "Dto.", "Localidad", "Provincia", "Telefono", "Sexo", "Estado Civil", "Documento Tipo", "Documento Número", "Fecha de Nacimiento", "Nacionalidad", "Cond. de Trabajo", "Número Sec. UATRE", "Tarea que realiza", "Estudios", "Otros Estudios", "Obra Social", "SISTEMA DE JUBILACIONES", "Nombre AFJP", "Sindicato", "Número de Afiliado", "Domicilio donde presta servicios", "Provincia donde presta servicios", "Localidad donde presta servicios", "Codigo Postal del lugar donde presta servicios"
        
        linea = """" & tipo & """,""" & apellido & """,""" & nombre & """,""" & _
                Cuil & """,""" & Calle & """,""" & depto & """,""" & _
                cuartel & """,""" & Nro & """,""" & Piso & """,""" & _
                nrodep & """,""" & Loc & """,""" & prov & """,""" & _
                Telefono & """,""" & Sexo & """,""" & estciv & """,""" & _
                tipodoc & """,""" & Nrodoc & """,""" & fec_nac & """,""" & _
                nacional & """,""" & condi & """,""" & nroUATRE & """,""" & _
                tareas & """,""" & estudio & """,""" & otrosest & """,""" & _
                osocial & """,""" & cajjub & """,""" & afjp & """,""" & _
                sindica & """,""" & nroafil & """,""" & domtrab & """,""" & _
                provtrab & """,""" & loctrab & """,""" & cpostrab & """"
                
        ArchExp.Write linea
        ArchExp.writeline
        
        
        ' Conyuge
        tipo = "1"

        StrSql = "SELECT familiar.famtrab,familiar.faminc,familiar.famsalario,tipodocu.tidsigla,ter_doc.nrodoc,"
        StrSql = StrSql & " parentesco.paredesc,tercero.terfecnac,tercero.terape,tercero.ternom,tercero.tersex"
        StrSql = StrSql & " FROM familiar INNER JOIN tercero ON familiar.ternro = tercero.ternro"
        StrSql = StrSql & " INNER JOIN parentesco ON parentesco.parenro = familiar.parenro"
        StrSql = StrSql & " LEFT JOIN ter_doc ON ter_doc.ternro = familiar.ternro AND ter_doc.tidnro <= 4"
        StrSql = StrSql & " LEFT JOIN tipodocu ON tipodocu.tidnro = ter_doc.tidnro"
        StrSql = StrSql & " WHERE familiar.famest = -1 AND familiar.parenro = 3 AND familiar.empleado = " & ternro_emp
        OpenRecordset StrSql, rs2
        f_tipodoc = ""
        f_Nrodoc = ""
        f_paren = ""
        f_fec_nac = ""
        f_empleo = ""
        f_asig = ""
        f_incap = ""
        f_apellido = ""
        f_nombre = ""
        f_tersex = ""
        If Not rs2.EOF Then
            f_tipodoc = IIf(IsNull(rs2!tidsigla), "", rs2!tidsigla)
            f_Nrodoc = IIf(IsNull(rs2!Nrodoc), "", rs2!Nrodoc)
            f_paren = IIf(IsNull(rs2!paredesc), "", rs2!paredesc)
            If rs2!terfecnac <> "" And Not IsNull(rs2!terfecnac) Then
                f_fec_nac = Format(Day(rs2!terfecnac), "00") & "-" & Format(Month(rs2!terfecnac), "00") & "-" & Format(Year(rs2!terfecnac), "0000")
            End If
            If CInt(rs2!famtrab) = -1 Then
                f_empleo = 1
            Else
                f_empleo = 0
            End If
            If CInt(rs2!faminc) = -1 Then
                f_asig = 1
            Else
                f_asig = 0
            End If
            If CInt(rs2!famsalario) = -1 Then
                f_incap = 1
            Else
                f_incap = 0
            End If
            f_apellido = IIf(IsNull(rs2!terape), "", rs2!terape)
            f_nombre = IIf(IsNull(rs2!ternom), "", rs2!ternom)
            If CInt(rs2!tersex) = -1 Then
                f_tersex = "Masc."
            Else
                f_tersex = "Fem."
            End If
            
            
            'RENGLON DE CONYUGE
            'tipo , "Tipo de Conyuge" (Conyuge/Conviviente)","Apellidos",
            '"Nombres","Fecha de Nacimiento",
            '"Documento Tipo","Documento Número",EMPLEO,
            'ASIG,INCAP
            
            linea = """" & tipo & """,""" & f_paren & """,""" & f_apellido & """,""" & _
                    f_nombre & """,""" & f_fec_nac & """,""" & _
                    f_tipodoc & """,""" & f_Nrodoc & """,""" & f_empleo & """,""" & _
                    f_asig & """,""" & f_incap & """"
            
            ArchExp.Write linea
            ArchExp.writeline
            
        End If
        rs2.Close
        
        
        ' Otros Familiares
        tipo = "2"

        StrSql = "SELECT familiar.famtrab,familiar.famsalario,familiar.famestudia,tipodocu.tidsigla,ter_doc.nrodoc,"
        StrSql = StrSql & " parentesco.paredesc,tercero.terfecnac,tercero.terape,tercero.ternom"
        StrSql = StrSql & " FROM familiar INNER JOIN tercero ON familiar.ternro = tercero.ternro"
        StrSql = StrSql & " INNER JOIN parentesco ON parentesco.parenro = familiar.parenro"
        StrSql = StrSql & " LEFT JOIN ter_doc ON ter_doc.ternro = familiar.ternro AND ter_doc.tidnro <= 4"
        StrSql = StrSql & " LEFT JOIN tipodocu ON tipodocu.tidnro = ter_doc.tidnro"
        StrSql = StrSql & " WHERE familiar.famest = -1 AND familiar.parenro <> 3 AND familiar.empleado = " & ternro_emp
        StrSql = StrSql & " ORDER BY familiar.parenro"
        OpenRecordset StrSql, rs2
        Do Until rs2.EOF
            f_tipodoc = ""
            f_Nrodoc = ""
            f_paren = ""
            f_fec_nac = ""
            f_empleo = ""
            f_esco = ""
            f_incap = ""
            f_apellido = ""
            f_nombre = ""
                
            f_tipodoc = IIf(IsNull(rs2!tidsigla), "", rs2!tidsigla)
            f_Nrodoc = IIf(IsNull(rs2!Nrodoc), "", rs2!Nrodoc)
            f_paren = IIf(IsNull(rs2!paredesc), "", rs2!paredesc)
            If rs2!terfecnac <> "" And Not IsNull(rs2!terfecnac) Then
                f_fec_nac = Format(Day(rs2!terfecnac), "00") & "-" & Format(Month(rs2!terfecnac), "00") & "-" & Format(Year(rs2!terfecnac), "0000")
            End If
            If CInt(rs2!famestudia) = -1 Then
                f_esco = 1
            Else
                f_esco = 0
            End If
            If CInt(rs2!famsalario) = -1 Then
                f_incap = 1
            Else
                f_incap = 0
            End If
            f_apellido = IIf(IsNull(rs2!terape), "", rs2!terape)
            f_nombre = IIf(IsNull(rs2!ternom), "", rs2!ternom)
        
            'FAMILIAR
            'TIPO,"Apellidos",
            '"Nombres"","Parentesco","Fecha de Nacimiento","Sexo",
            '"Documento Tipo","Documento Número","Escolaridad",INCAP
            
            linea = """" & tipo & """,""" & f_apellido & """,""" & _
                    f_nombre & """,""" & f_paren & """,""" & f_fec_nac & """,""" & f_tersex & """,""" & _
                    f_tipodoc & """,""" & f_Nrodoc & """,""" & f_esco & """,""" & f_incap & """"
            
            ArchExp.Write linea
            ArchExp.writeline
            
            rs2.MoveNext
            
        Loop
        rs2.Close
        
        rs.MoveNext
        
        'Actualizo el estado del proceso
        TiempoAcumulado = GetTickCount
        cantRegistros = cantRegistros - 1
        Progreso = Progreso + IncPorc
        
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
        StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'"
        StrSql = StrSql & ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
        
    Loop
    
            
    ArchExp.Close
    
Fin:
    'Cierro y libero todo
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    If rs2.State = adStateOpen Then rs2.Close
    Set rs2 = Nothing
  

Exit Sub

ME_Local:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Description
    Flog.writeline Espacios(Tabulador * 1) & "Ultimo SQL: " & StrSql
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.writeline
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
        StrSql = StrSql & " AND mapclanro = 3 " 'RENATRE
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

