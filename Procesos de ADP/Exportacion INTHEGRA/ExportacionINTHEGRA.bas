Attribute VB_Name = "ExpINTHEGRA"
Option Explicit

'Global Const Version = "1.00"
'Global Const FechaModificacion = "21/04/2014"
'Global Const UltimaModificacion = "Version Inicial - CAS 21681 - TABACAL - Interfaz INTHEGRA - Automatización"

'Global Const Version = "1.01"
'Global Const FechaModificacion = "24/04/2014"
'Global Const UltimaModificacion = "CAS 21681 - se modifico nombre de archivo generado a INTHEGRA_AAAAMMDD.txt, donde AAAAMMDD, sea año, mes y día de la generación"
                                    'se corrigio bug en query al obtener los valores de los conceptos de limite de compra y limite de credito
                                    
'Global Const Version = "1.02"
'Global Const FechaModificacion = "25/04/2014"
'Global Const UltimaModificacion = "CAS 21681 - se estaba filtrando conceptos por campo concnro en lugar de conccod"

'Global Const Version = "1.03"
'Global Const FechaModificacion = "03/12/2014"
'Global Const UltimaModificacion = "CAS-26790 - se agregaron a exportacion CBU del empleado y codigo externo de estructura empresa"

Global Const Version = "1.04"
Global Const FechaModificacion = "08/01/2015"
Global Const UltimaModificacion = "CAS-26790 - se corrige bug en funcion CBUCuentaBancaria "



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

Dim rs_Confrep As New ADODB.Recordset


Private Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial exportacion
' Autor      : MDZ
' Fecha      : 21/04/2014
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
    
    Nombre_Arch = PathFLog & "ReporteDotacionPersonal" & "-" & NroProceso & ".log"
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
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

    On Error GoTo ME_Main
    
    TiempoInicialProceso = GetTickCount
    HuboErrores = False
    
    Nombre_Arch = PathFLog & "ExpINTHEGRA" & "-" & NroProceso & ".log"
    
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
    'StrSql = StrSql & " AND btprcnro = 164"
    OpenRecordset StrSql, rs
    
    If Not rs.EOF Then
       'Obtengo los parametros del proceso
        If IsNull(rs!bprcparam) Then
            Parametros = ""
        Else
            Parametros = rs!bprcparam
        End If
    
       ArrParametros = Split(Parametros, "@")
       
       'Empresa = CLng(ArrParametros(0))
       'Fecha = ArrParametros(1)
       'Tenro = CLng(ArrParametros(2))
       'Estrnro = CLng(ArrParametros(3))
       'informa_fecha = CInt(ArrParametros(4))
       
       'Call Generar_Archivo(Empresa, Fecha, Tenro, Estrnro, informa_fecha)
       Call Generar_Archivo
       
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


'Private Sub Generar_Archivo(ByVal Empnro As Long, ByVal fecha_alta As String, ByVal tipoestr As Long, ByVal estructura As Long, ByVal informa_fecha_alta As Integer)
Private Sub Generar_Archivo()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que genera la exportacion
' Autor      : MDZ
' Fecha      : 21/04/2014
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
'Dim Sexo As String
Dim estciv As String
'Dim TipoDoc As String
'Dim Nrodoc As String
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

Dim ConcLimiteCompra As String
Dim ConcLimiteCredito As String
Dim estrnroQuincenal As Integer
Dim estrnroMensual As Integer
Dim NroReporte As Integer


ConcLimiteCompra = ""
ConcLimiteCredito = ""
estrnroMensual = 0
estrnroQuincenal = 0

Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
    
    On Error GoTo ME_Local
      
    'Directorio de exportacion
    StrSql = "SELECT sis_dirsalidas FROM sistema WHERE sisnro = 1 "
    If rs.State = adStateOpen Then rs.Close
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        Directorio = Trim(rs!sis_dirsalidas) & "\INTHEGRA"
    End If
    
    Nombre_Arch = Directorio & "\INTHEGRA_" & Format(Date, "yyyyMMdd") & ".txt"

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
    
    '------------------------------------------------------------------
    'Configuracion del Reporte
    '------------------------------------------------------------------
    NroReporte = 418
    StrSql = "SELECT * FROM confrep "
    StrSql = StrSql & " WHERE repnro = " & NroReporte
    OpenRecordset StrSql, rs_Confrep
    Do While Not rs_Confrep.EOF
        If rs_Confrep("conftipo") = "LCO" Then
            ConcLimiteCompra = rs_Confrep("confval2")
        ElseIf rs_Confrep("conftipo") = "LCR" Then
            ConcLimiteCredito = rs_Confrep("confval2")
        End If
        rs_Confrep.MoveNext
    Loop
    
    
    If ConcLimiteCompra = "" Or ConcLimiteCredito = "" Then
       Flog.writeline Espacios(Tabulador * 1) & "Reporte " & NroReporte & " mal configurado."
       Flog.writeline Espacios(Tabulador * 1) & "Faltan Configurar los Conceptos para el Limite de Compra o el Limite de Credito "
       HuboErrores = True
       GoTo Fin
    End If
    
    
    '------------------------------------------------------------------
    'Busco los datos
    '------------------------------------------------------------------
    StrSql = "select e.ternro, empleg, e.terape, e.terape2, e.ternom, e.ternom2, tersex, empest, d.tidnro, d.nrodoc, " & _
            " (SELECT TOP 1 dlimonto FROM proceso INNER JOIN cabliq ON (proceso.pronro=cabliq.pronro) INNER JOIN detliq ON (cabliq.cliqnro=detliq.cliqnro) INNER JOIN concepto ON (detliq.concnro=concepto.concnro) WHERE empleado=e.ternro AND conccod=" & ConcLimiteCompra & " ORDER BY profecfin DESC ) lim_com," & _
            " (SELECT TOP 1 dlimonto FROM proceso INNER JOIN cabliq ON (proceso.pronro=cabliq.pronro) INNER JOIN detliq ON (cabliq.cliqnro=detliq.cliqnro) INNER JOIN concepto ON (detliq.concnro=concepto.concnro) WHERE empleado=e.ternro AND conccod=" & ConcLimiteCredito & " ORDER BY profecfin DESC  ) lim_cre," & _
            " (select top 1 estrnro from his_estructura h where h.ternro=e.ternro and tenro=22 and htetdesde<=" & ConvFecha(Now) & " and (htethasta>=" & ConvFecha(Now) & " or htethasta is null)) condicion " & _
            " FROM empleado e inner join tercero t on e.ternro=t.ternro left join ter_doc d on d.ternro=e.ternro order by e.ternro, d.tidnro"

    Flog.writeline
    Flog.writeline StrSql
    Flog.writeline
    
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
    
    
    
    Dim l_tercero
    Dim l_condicion

    l_tercero = 0

    Do Until rs.EOF
        If l_tercero <> rs!ternro Then
        
            If Not IsNull(rs!Condicion) Then  ' si no tiene la estructura tipo 32
                
                
                l_condicion = Condicion(rs("condicion"))
                    
                
                If l_condicion <> "" Then ' si la estructura no esta configurada en confrep
                    
                    linea = "" & l_condicion                                                    ' condicion 1:Mensual, 2:Quincenal
                    linea = linea & imprimirNumero(rs!empleg, 20)                          ' legajo
                    linea = linea & "1"                                                         ' Tipo Fijo "1"
                    linea = linea & imprimirTexto(rs!terape & " " & rs!terape2, 40)   ' apellido
                    linea = linea & imprimirTexto(rs!ternom & " " & rs!ternom2, 40)   ' nombre
    
                    If Not IsNull(rs("nrodoc")) Then
                        linea = linea & imprimirNumero(TipoDoc(rs!tidnro), 1)              ' tipo de documento
                        linea = linea & imprimirNumero(rs!Nrodoc, 15)                      ' numero de documento
                    Else
                        linea = linea & imprimirTexto("-", 1)                               ' tipo de documento
                        linea = linea & imprimirTexto("---------------", 15)                    ' numero de documento
                    End If
                    
                    linea = linea & imprimirNumero(Sexo(rs!tersex), 5)                     ' sexo
                    
                    If Not IsNull(rs("lim_com")) Then
                        linea = linea & imprimirNumero(Replace(Replace(FormatNumber(Abs(rs!lim_com), 2), ".", ""), ",", ""), 10)  ' Limite de Compra
                    Else
                        linea = linea & imprimirNumero(0, 10)                                   ' Limite de Compra
                    End If
                    
                    If Not IsNull(rs!lim_cre) Then
                        linea = linea & imprimirNumero(Replace(Replace(FormatNumber(Abs(rs!lim_cre), 2), ".", ""), ",", ""), 10)  ' Limite de Credito
                    Else
                        linea = linea & imprimirNumero(0, 10)                                   ' Limite de Credito
                    End If
                    
                    linea = linea & imprimirNumero(Estado(rs!empest), 1)                   ' status
                    
                    '03/12/2014 - MDZ
                    linea = linea & imprimirNumero(CodigoExterno(10, rs!ternro), 1)       ' codigo externo de estructura empresa
                    linea = linea & imprimirTexto(CBUCuentaBancaria(rs!ternro), 22)       ' cbu del empleado
                    
                    
                    'ArchExp.Write linea
                    ArchExp.writeline linea
                Else
                    Flog.writeline "la estructura " & rs!Condicion & " del tipo 22 no corresponde a ninguna condicion configurada." & _
                                    " El Empleado con legajo " & rs!empleg & " no se exportará!"
                End If
            End If
        End If
        l_tercero = rs!ternro
       
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
    Flog.writeline "Cierro el archivo. FIN!!!"
    
Fin:
    'Cierro y libero todo
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    'If rs2.State = adStateOpen Then rs2.Close
    'Set rs2 = Nothing
  

Exit Sub

ME_Local:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Description
    Flog.writeline Espacios(Tabulador * 1) & "Ultimo SQL: " & StrSql
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.writeline
    
End Sub

Function imprimirTexto(Texto, Longitud)
    
    imprimirTexto = Left(Texto & String(Longitud, " "), Longitud)
    
End Function


Function imprimirNumero(Texto, Longitud)
    
    'imprimirNumero = Format(texto, String("0", longitud))
    imprimirNumero = Right(String(Longitud, "0") & Texto, Longitud)

End Function


Function Condicion(estructura)
    
    'mapeo de estructura configurado en confrep
    'TIPO       = CON
    'confval    = estructura
    'columna    = 1:Mensual, 2:Quincenal
    
    Condicion = ""
    
    rs_Confrep.MoveFirst
    Do While Not rs_Confrep.EOF
        
        If UCase(rs_Confrep!conftipo) = "CON" Then
            If CInt(rs_Confrep!confval) = CInt(estructura) Then
                Condicion = rs_Confrep!confval2
                Exit Do
            End If
        End If
        rs_Confrep.MoveNext
    Loop
    

End Function

Function TipoDoc(tipo)

    'mapeo el tipo con el codigo configurado en confrep
    'TIPO       = DOC
    'confval    = tipo
    'columna    = 1 DNI
    '             2 Libreta Enrolamiento
    '             3 Libreta Cívica
    '             4 Pasaporte
    '             5 Cédula de Identidad
    '             6 Otros
    
    TipoDoc = 6
    
    rs_Confrep.MoveFirst
    Do While Not rs_Confrep.EOF
        
        If UCase(rs_Confrep!conftipo) = "DOC" Then
            If CInt(rs_Confrep!confval) = CInt(tipo) Then
                TipoDoc = rs_Confrep!confval2
                Exit Do
            End If
        End If
        rs_Confrep.MoveNext
    Loop
           
End Function

Function Sexo(tipo)

    '1  Masculino
    '2  Femenino
    If tipo Then
        Sexo = 1
    Else
        Sexo = 2
    End If

End Function

Function Estado(est)

    '1  activo
    '2  inactivo
    
    If est Then
        Estado = 1
    Else
        Estado = 2
    End If

End Function

'devuelve el codigo externo de la estructura correspondiente al tipo y tercero pasados para la fecha actual
Function CodigoExterno(ByVal te, ByVal ternro)
    
    CodigoExterno = " "
    Dim rs As New ADODB.Recordset
    
    StrSql = "select estrcodext from his_estructura h inner join estructura e ON (h.estrnro=e.estrnro) " & _
            "where ternro=" & ternro & " and h.tenro=" & te & " and " & _
            "h.htetdesde<=GETDATE() and (h.htethasta is null or h.htethasta>=GETDATE())"

    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        If Not IsNull(rs("estrcodext")) Then
            CodigoExterno = Left(rs("estrcodext"), 1)
        End If
    End If
End Function

'devuelve el CBU del tercero
Function CBUCuentaBancaria(ByVal ternro)
    CBUCuentaBancaria = ""
    Dim rs As New ADODB.Recordset
    StrSql = "select ctabcbu from ctabancaria where ctabestado=-1 and ternro=" & ternro
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        If Not IsNull(rs("ctabcbu")) Then
            CBUCuentaBancaria = Trim(rs("ctabcbu"))
        End If
    End If
End Function


