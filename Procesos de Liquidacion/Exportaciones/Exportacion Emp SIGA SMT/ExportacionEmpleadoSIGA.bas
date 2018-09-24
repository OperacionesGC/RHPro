Attribute VB_Name = "ExpEmpSIGA"
Option Explicit

'Global Const Version = "1.01"
'Global Const FechaModificacion = "27/04/2007"
'Global Const UltimaModificacion = " " 'Version Inicial

'Global Const Version = "1.02"
'Global Const FechaModificacion = "27/05/2007"
'Global Const UltimaModificacion = " " ' FAF - Error en la funcion imprimirtexto

'Global Const Version = "1.03"
'Global Const FechaModificacion = "04/06/2007"
'Global Const UltimaModificacion = " " ' FAF - No salian los empleados activos. Se busca el ultimo grupo de liquidacion al que pertenecio

Global Const Version = "1.04"
Global Const FechaModificacion = "31/07/2009"
Global Const UltimaModificacion = " " 'Encriptacion de string connection

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

    On Error GoTo 0

    On Error GoTo ME_Main
    
    TiempoInicialProceso = GetTickCount
    HuboErrores = False
    
    Nombre_Arch = PathFLog & "ExpEmpSIGA" & "-" & NroProceso & ".log"
    
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
       
       Empresa = CLng(ArrParametros(0))
       Fecha = ArrParametros(1)
       Tenro = CLng(ArrParametros(2))
       Estrnro = CLng(ArrParametros(3))
       informa_fecha = CInt(ArrParametros(4))
       
       Call Generar_Archivo(Empresa, Fecha, Tenro, Estrnro, informa_fecha)
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


Private Sub Generar_Archivo(ByVal Empnro As Long, ByVal fecha_alta As String, ByVal tipoestr As Long, ByVal estructura As Long, ByVal informa_fecha_alta As Integer)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que genera la exportacion
' Autor      : FAF
' Fecha      : 27/04/2007
' Ultima Mod :
' Descripcion:
' ---------------------------------------------------------------------------------------------
'Dim objRs As New ADODB.Recordset
Dim Nombre_Arch As String
Dim Directorio As String
Dim Carpeta
Dim fs1

Dim cantRegistros As Long

Dim NroReporte As Integer
Dim lista_estructuras As String
Dim tipo_estructura As Integer
Dim fecha_fase As String
Dim Seguir As Boolean
Dim v_fecalta
Dim v_fecbaja
Dim v_fecestruc
Dim v_grupo
Dim v_dni
Dim v_categ
Dim v_ccosto
Dim v_contrato
Dim v_tipemple
Dim v_estrnro

Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rsConsult As New ADODB.Recordset
    
    On Error GoTo ME_Local
      
    'Directorio de exportacion
    StrSql = "SELECT sis_dirsalidas FROM sistema WHERE sisnro = 1 "
    If rs.State = adStateOpen Then rs.Close
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        Directorio = Trim(rs!sis_dirsalidas) & "\datexportados"
    End If
    
    Nombre_Arch = Directorio & "\expemple.txt"
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
    NroReporte = 192
    StrSql = "SELECT * FROM confrep "
    StrSql = StrSql & " WHERE repnro = " & NroReporte
    If rs.State = adStateOpen Then rs.Close
    OpenRecordset StrSql, rs
    lista_estructuras = "0"
    tipo_estructura = 32
    fecha_fase = "01/01/2000"
    If rs.EOF Then
        Flog.writeline "No se encontró la configuración del Reporte"
        Flog.writeline "   Se deben configurar 3 tipos de columnas:"
        Flog.writeline "     FF : Indica la fecha de baja a partir de la cual se consideran las fases. Default 01/01/2000. Unico"
        Flog.writeline "     TE : Tipo de estructura   No se encontró la configuración del Reporte. Default 32 (Grupo de Liquidacion). Unico"
        Flog.writeline "     EST: Lista de estructuras del tipo anterior. Una o mas"
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
            End Select
            rs.MoveNext
        Loop
    End If
    rs.Close
    lista_estructuras = lista_estructuras & ","
    Flog.writeline "     "
    
    StrSql = "SELECT empleado.ternro, empleg, empest, empleado.ternom, empleado.terape, tercero.terfecnac "
    StrSql = StrSql & " FROM batch_empleado "
    StrSql = StrSql & " INNER JOIN empleado ON batch_empleado.ternro = empleado.ternro "
    StrSql = StrSql & " INNER JOIN tercero ON empleado.ternro = tercero.ternro "
    StrSql = StrSql & " WHERE bpronro = " & NroProceso
    StrSql = StrSql & " ORDER BY estado "
    
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
        Seguir = True
        
        ' Fases: SE EXPORTAN SOLO EMPLEADOS ACTIVOS O INACTIVOS CON FECHA DE BAJA MAYOR AL 1/1/2000
        StrSql = "SELECT * FROM fases WHERE empleado =" & rs!ternro & " ORDER BY altfec DESC"
        OpenRecordset StrSql, rs2
        If rs2.EOF Then
            Flog.writeline "*** Empleado sin fechas de alta/baja (fases), no se informa: " & rs!empleg
            Seguir = False
        End If
        
'        If seguir And Not rs!empest Then
'            seguir = False
'        End If
        
        If Seguir Then
            If Not IsNull(rs2!bajfec) And rs2!bajfec <> "" Then
                If CDate(rs2!bajfec) < CDate(fecha_fase) Then
                    Seguir = False
                End If
'            Else
'                seguir = False
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
        
        v_grupo = 0
        v_estrnro = 0
        If Seguir Then
            StrSql = " SELECT his_estructura.estrnro, estructura.estrcodext FROM his_estructura "
            StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
'            StrSql = StrSql & " AND htetdesde <= " & ConvFecha(v_fecestruc)
'            StrSql = StrSql & " AND (htethasta Is Null Or htethasta >= " & ConvFecha(v_fecestruc) & ") "
            StrSql = StrSql & " AND his_estructura.tenro = " & tipo_estructura & " AND his_estructura.ternro = " & rs!ternro
'            StrSql = StrSql & " AND his_estructura.estrnro IN (" & lista_estructuras & ")"
            StrSql = StrSql & " ORDER BY his_estructura.htetdesde DESC, his_estructura.htethasta "
            OpenRecordset StrSql, rs2
            If rs2.EOF Then
                Seguir = True
            Else
                v_grupo = rs2!estrcodext
                v_estrnro = rs2!Estrnro
            End If
        End If
            
        If Seguir Then
            If InStr(1, lista_estructuras, CStr("," & v_estrnro & ",")) > 0 Then
                ' Documento
                StrSql = " SELECT ter_doc.nrodoc "
                StrSql = StrSql & " FROM ter_doc "
                StrSql = StrSql & " WHERE ter_doc.tidnro <= 4 and ter_doc.ternro= " & rs!ternro
                OpenRecordset StrSql, rsConsult
                v_dni = ""
                If Not rsConsult.EOF Then
                   v_dni = rsConsult!Nrodoc
                End If
                rsConsult.Close
                
                ' Categoria
                StrSql = " SELECT catnro "
                StrSql = StrSql & " From his_estructura"
                StrSql = StrSql & " INNER JOIN categoria ON categoria.estrnro=his_estructura.estrnro "
                'StrSql = StrSql & " AND htetdesde <= " & ConvFecha(v_fecestruc)
                'StrSql = StrSql & " And (htethasta Is Null Or htethasta >= " & ConvFecha(v_fecestruc) & ") "
                StrSql = StrSql & " And his_estructura.tenro = 3 And his_estructura.ternro = " & rs!ternro
                StrSql = StrSql & " ORDER BY his_estructura.htetdesde DESC"
                OpenRecordset StrSql, rsConsult
                v_categ = ""
                If Not rsConsult.EOF Then
                   v_categ = rsConsult!catnro
                End If
                rsConsult.Close
                
                ' Centro Costo
                StrSql = " SELECT estrcodext "
                StrSql = StrSql & " From his_estructura"
                StrSql = StrSql & " inner join estructura on his_estructura.estrnro = estructura.estrnro "
                'StrSql = StrSql & " WHERE htetdesde <= " & ConvFecha(v_fecestruc)
                'StrSql = StrSql & " And (htethasta Is Null Or htethasta >= " & ConvFecha(v_fecestruc) & ") "
                StrSql = StrSql & " And his_estructura.tenro = 5 And his_estructura.ternro = " & rs!ternro
                StrSql = StrSql & " ORDER BY his_estructura.htetdesde DESC"
                OpenRecordset StrSql, rsConsult
                v_ccosto = ""
                If Not rsConsult.EOF Then
                   v_ccosto = rsConsult!estrcodext
                End If
                rsConsult.Close
                
                ' Categoria
                StrSql = " SELECT tcnro "
                StrSql = StrSql & " From his_estructura"
                StrSql = StrSql & " INNER JOIN tipocont ON tipocont.estrnro=his_estructura.estrnro "
                'StrSql = StrSql & " AND htetdesde <= " & ConvFecha(v_fecestruc)
                'StrSql = StrSql & " And (htethasta Is Null Or htethasta >= " & ConvFecha(v_fecestruc) & ") "
                StrSql = StrSql & " And his_estructura.tenro = 18 And his_estructura.ternro = " & rs!ternro
                StrSql = StrSql & " ORDER BY his_estructura.htetdesde DESC"
                OpenRecordset StrSql, rsConsult
                v_contrato = ""
                If Not rsConsult.EOF Then
                   v_contrato = rsConsult!tcnro
                End If
                rsConsult.Close
                
                v_tipemple = ""
                If v_grupo = "1" Or v_grupo = "2" Or v_grupo = "8" Or v_grupo = "9" Then
                    v_tipemple = "J"
                End If
                
                If v_grupo = "11" Or v_grupo = "12" Then
                    v_tipemple = "H"
                End If
                
                        
                Call imprimirTexto(Format(v_grupo, "0000"), ArchExp, 4, 1)
                Call imprimirTexto(Format(rs!empleg, "00000000"), ArchExp, 8, 1)
                Call imprimirTexto(rs!terape, ArchExp, 20, 1)
                Call imprimirTexto(rs!ternom, ArchExp, 20, 1)
                Call imprimirTexto(Format(v_categ, "0000"), ArchExp, 4, 1)
                Call imprimirTexto(Format(v_fecalta, "dd/mm/yy"), ArchExp, 8, 1)
                Call imprimirTexto(Format(v_fecbaja, "dd/mm/yy"), ArchExp, 8, 1)
                Call imprimirTexto(v_dni, ArchExp, 13, 1)
                Call imprimirTexto(Format(rs!terfecnac, "dd/mm/yyyy"), ArchExp, 10, 1)
                Call imprimirTexto(v_ccosto, ArchExp, 20, 1)
                Call imprimirTexto(Format(v_contrato, "00"), ArchExp, 2, 1)
                Call imprimirTexto(v_tipemple, ArchExp, 6, 1)
                
                ArchExp.writeline
                
            End If
            
        End If
        
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
    If rs2.State = adStateOpen Then rs2.Close
    Set rs2 = Nothing
  

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
        If longTexto >= Longitud Then
            cadena = Mid(CStr(Texto), 1, Longitud)
        Else
            cadena = Mid(CStr(Texto), 1, Longitud) & String(Longitud - longTexto, " ")
        End If
    End If
    
    archivo.Write cadena
    
End Sub


