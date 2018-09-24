Attribute VB_Name = "FeedbackSIGA"
Option Explicit

'Global Const Version = "1.01"
'Global Const FechaModificacion = "02/05/2007"
'Global Const UltimaModificacion = " " 'Version Inicial

'Global Const Version = "1.02"
'Global Const FechaModificacion = "23/07/2007"
'Global Const UltimaModificacion = " " 'Version Inicial FAF - Se restringio a mostrar los conceptos imprimibles y/o de tipo Contribuciones Empleador

Global Const Version = "1.03"
Global Const FechaModificacion = "31/07/2009"
Global Const UltimaModificacion = " " 'MB - Encriptacion de string connection

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
' Fecha      : 02/05/2007
' Ultima Mod :
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim strCmdLine As String
Dim Nombre_Arch As String

Dim rs As New ADODB.Recordset
Dim PID As String
Dim Parametros As String
Dim ArrParametros

Dim Periodo As Long
Dim TipogruLiq As Integer


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
'    On Error GoTo 0

    On Error GoTo ME_Main
    
    TiempoInicialProceso = GetTickCount
    HuboErrores = False
    
    Nombre_Arch = PathFLog & "ExpFeedbackSIGA" & "-" & NroProceso & ".log"
    
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
    
    On Error GoTo ME_Main
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & ", bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline Espacios(Tabulador * 0) & "Obtengo los datos del proceso"
    
    TiempoAcumulado = GetTickCount
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProceso
    StrSql = StrSql & " AND btprcnro = 169"
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
       'Obtengo los parametros del proceso
       Parametros = rs!bprcparam
       ArrParametros = Split(Parametros, "@")
       
       Periodo = CLng(ArrParametros(0))
       TipogruLiq = CInt(ArrParametros(2))
       
       Flog.writeline "  Periodo     : " & Periodo
       Flog.writeline "  Opcion      : " & TipogruLiq & " (1.- AGR 2.- PIN)"
       Flog.writeline
        
       Call Generar_Archivo(Periodo, TipogruLiq)
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


Private Sub Generar_Archivo(ByVal pliqnro As Long, ByVal TipogruLiq As Integer)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que genera la exportacion - expliq01.p - San Martin de Tabacal
' Autor      : FAF
' Fecha      : 02/05/2007
' Ultima Mod :
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Nombre_Arch As String
Dim Directorio As String
Dim Carpeta
Dim fs1

Dim cantRegistros As Long

Dim NroReporte As Integer
Dim c_fecdesde As String
Dim c_fechasta As String
Dim lista_estruc As String
Dim tipo_estructura
Dim v_categ As String
Dim v_cc As String
Dim v_thnro As Long
Dim fechasta As Date

Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
    
    On Error GoTo ME_Local
    
    StrSql = " SELECT pliqdesde, pliqhasta FROM periodo WHERE pliqnro=" & pliqnro
    OpenRecordset StrSql, rs
    
    c_fecdesde = Date
    c_fechasta = Date
    fechasta = Date
    If Not rs.EOF Then
        c_fecdesde = CStr(Format(Year(rs!pliqdesde), "0000") & Format(Month(rs!pliqdesde), "00") & Format(Day(rs!pliqdesde), "00"))
        c_fechasta = CStr(Format(Year(rs!pliqhasta), "0000") & Format(Month(rs!pliqhasta), "00") & Format(Day(rs!pliqhasta), "00"))
        fechasta = rs!pliqhasta
    End If
    rs.Close
    
    'Directorio de exportacion
    StrSql = "SELECT sis_dirsalidas FROM sistema WHERE sisnro = 1 "
    If rs.State = adStateOpen Then rs.Close
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        Directorio = Trim(rs!sis_dirsalidas) & "\tmp"
    End If
    
    If TipogruLiq = 1 Then
        Nombre_Arch = Directorio & "\AGR_" & c_fecdesde & "a" & c_fechasta & ".txt"
    Else
        Nombre_Arch = Directorio & "\PIN_" & c_fecdesde & "a" & c_fechasta & ".txt"
    End If
    
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
    NroReporte = 194
    StrSql = "SELECT * FROM confrep "
    StrSql = StrSql & " WHERE repnro = " & NroReporte
    If rs.State = adStateOpen Then rs.Close
    OpenRecordset StrSql, rs
    lista_estruc = "0"
    tipo_estructura = 32
    If rs.EOF Then
        Flog.writeline "No se encontró la configuración del Reporte"
        Flog.writeline "   Se deben configurar 2 tipos de columnas:"
        Flog.writeline "     TE : Tipo de estructura   No se encontró la configuración del Reporte. Default 32 (Grupo de Liquidacion). Unico"
        Flog.writeline "     EST: Lista de estructuras del tipo anterior. Una o mas"
        Exit Sub
    Else
        Do Until rs.EOF
            Select Case rs!conftipo
                Case "TE":
                    tipo_estructura = rs!confval
                Case "EST":
                    If UCase(Mid(rs!confetiq, 1, 3)) = "AGR" And TipogruLiq = 1 Then
                        lista_estruc = lista_estruc & "," & rs!confval
                    ElseIf UCase(Mid(rs!confetiq, 1, 3)) = "PIN" And TipogruLiq = 2 Then
                        lista_estruc = lista_estruc & "," & rs!confval
                    End If
            End Select
            rs.MoveNext
        Loop
    End If
    rs.Close
    
    Flog.writeline "     "
    
    
    StrSql = "SELECT estrcodext, empleado.ternro, empleg, tprocdesc, cliqnro "
    StrSql = StrSql & " FROM proceso "
    StrSql = StrSql & " INNER JOIN tipoproc ON tipoproc.tprocnro = proceso.tprocnro "
    StrSql = StrSql & " INNER JOIN cabliq ON cabliq.pronro = proceso.pronro "
    StrSql = StrSql & " INNER JOIN empleado ON cabliq.empleado = empleado.ternro "
    StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.ternro = empleado.ternro "
    StrSql = StrSql & " AND htetdesde <= " & ConvFecha(fechasta)
    StrSql = StrSql & " AND (htethasta Is Null Or htethasta >= " & ConvFecha(fechasta) & ") "
    StrSql = StrSql & " AND his_estructura.tenro = " & tipo_estructura
    StrSql = StrSql & " AND his_estructura.estrnro IN (" & lista_estruc & ")"
    StrSql = StrSql & " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro "
    StrSql = StrSql & " WHERE proceso.pliqnro = " & pliqnro
    StrSql = StrSql & " ORDER BY proceso.profecini, proceso.pronro, his_estructura.estrnro, empleado.empleg "
    
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
    
        StrSql = " SELECT estructura.estrcodext FROM his_estructura "
        StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
        StrSql = StrSql & " AND htetdesde <= " & ConvFecha(fechasta)
        StrSql = StrSql & " AND (htethasta Is Null Or htethasta >= " & ConvFecha(fechasta) & ") "
        StrSql = StrSql & " AND his_estructura.tenro = 5 AND his_estructura.ternro = " & rs!ternro
        OpenRecordset StrSql, rs2
        If Not rs2.EOF Then
            v_cc = CStr(rs2!estrcodext)
        Else
            v_cc = ""
        End If
        rs2.Close
        
        StrSql = " SELECT estructura.estrcodext FROM his_estructura "
        StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
        StrSql = StrSql & " AND htetdesde <= " & ConvFecha(fechasta)
        StrSql = StrSql & " AND (htethasta Is Null Or htethasta >= " & ConvFecha(fechasta) & ") "
        StrSql = StrSql & " AND his_estructura.tenro = 3 AND his_estructura.ternro = " & rs!ternro
        OpenRecordset StrSql, rs2
        If Not rs2.EOF Then
            v_categ = CStr(rs2!estrcodext)
        Else
            v_categ = ""
        End If
        rs2.Close
        
        StrSql = " SELECT concepto.concnro, concepto.conccod, detliq.dlimonto, detliq.dlicant, tiph_con.thnro FROM detliq "
        StrSql = StrSql & " INNER JOIN concepto ON detliq.concnro = concepto.concnro "
        StrSql = StrSql & " LEFT JOIN tiph_con ON tiph_con.concnro = detliq.concnro "
        StrSql = StrSql & " WHERE (concepto.concimp = -1 OR concepto.tconnro=11) AND cliqnro = " & rs!cliqnro
        OpenRecordset StrSql, rs2
        
        Do Until rs2.EOF
            If rs2!thnro = "" Or IsNull(rs2!thnro) Then
                v_thnro = 0
            Else
                v_thnro = CLng(rs2!thnro)
            End If
            
            Call imprimirTexto(c_fecdesde, ArchExp, 8, 1)
            Call imprimirTexto("|", ArchExp, 1, 1)
            Call imprimirTexto(rs!tprocDesc, ArchExp, Len(CStr(rs!tprocDesc)), 1)
            Call imprimirTexto("|", ArchExp, 1, 1)
            Call imprimirTexto(rs!estrcodext, ArchExp, Len(CStr(rs!estrcodext)), 1)
            Call imprimirTexto("|", ArchExp, 1, 1)
            Call imprimirTexto(v_categ, ArchExp, Len(v_categ), 1)
            Call imprimirTexto("|", ArchExp, 1, 1)
            Call imprimirTexto(v_cc, ArchExp, Len(v_cc), 1)
            Call imprimirTexto("|", ArchExp, 1, 1)
            Call imprimirTexto(rs!empleg, ArchExp, Len(CStr(rs!empleg)), 1)
            Call imprimirTexto("|", ArchExp, 1, 1)
            Call imprimirTexto(rs2!Conccod, ArchExp, Len(CStr(rs2!Conccod)), 1)
            Call imprimirTexto("|", ArchExp, 1, 1)
            Call imprimirTexto(v_thnro, ArchExp, Len(CStr(v_thnro)), 1)
            Call imprimirTexto("|", ArchExp, 1, 1)
            Call imprimirTexto(Format(rs2!dlicant, "0.00"), ArchExp, Len(CStr(Format(rs2!dlicant, "0.00"))), 1)
            Call imprimirTexto("|", ArchExp, 1, 1)
            Call imprimirTexto(Format(rs2!dlimonto, "0.00"), ArchExp, Len(CStr(Format(rs2!dlimonto, "0.00"))), 1)
            
            ArchExp.writeline
                        
            rs2.MoveNext
            
        Loop
        rs2.Close
        
        
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
    
    ArchExp.Close
   
Fin:
    'Cierro y libero todo
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
  
Exit Sub

ME_Local:
    Flog.writeline
'    Resume Next
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Description
    Flog.writeline Espacios(Tabulador * 1) & "Ultimo SQL: " & StrSql
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.writeline
End Sub


Sub imprimirTexto(Texto, archivo, Longitud, derecha)
'Rutina genérica para imprimir un TEXTO, de una LONGITUD determinada.
'Los sobrantes se rellenan con CARACTER

Dim cadena
Dim txt
Dim u
Dim longTexto
    
    If IsNull(Texto) Then
        longTexto = 1
        cadena = " "
    Else
        longTexto = Len(Texto)
        cadena = Mid(CStr(Texto), 1, Longitud) & String(Longitud - longTexto, " ")
    End If
    
    archivo.Write cadena
    
End Sub


