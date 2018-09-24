Attribute VB_Name = "GenArchivoHSBC"
Option Explicit

'Global Const Version = "1.01"
'Global Const FechaModificacion = "13/06/2006"
'Global Const UltimaModificacion = " " 'Version Inicial

'Global Const Version = "1.02"
'Global Const FechaModificacion = "13/07/2006"
'Global Const UltimaModificacion = " " 'Version Inicial

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

Private Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion : Procedimiento inicial exportacion
' Autor       : Fernando Favre
' Fecha       : 31/03/2006
' Ultima Mod  :
' Descripcion :
' ---------------------------------------------------------------------------------------------
Dim strCmdLine As String
Dim Nombre_Arch As String

Dim rs As New ADODB.Recordset
Dim PID As String
Dim Parametros As String
Dim ArrParametros

Dim intnro As Long
Dim empresa As Long
Dim Fecha_Calculo As String
Dim mostrar_log As Boolean

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
    
    On Error GoTo 0

    On Error GoTo ME_Main
    
    TiempoInicialProceso = GetTickCount
    HuboErrores = False
    
    Nombre_Arch = PathFLog & "Exp-HSBC" & "-" & NroProceso & ".log"
    
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
    On Error GoTo 0

    On Error GoTo ME_Main
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & ", bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline Espacios(Tabulador * 0) & "Obtengo los datos del proceso"
    
    TiempoAcumulado = GetTickCount
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProceso
    StrSql = StrSql & " AND btprcnro = 133"
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
       'Obtengo los parametros del proceso
       Parametros = rs!bprcparam
       ArrParametros = Split(Parametros, "@")
       
       empresa = CLng(ArrParametros(0))
       
       Fecha_Calculo = ArrParametros(1)
       
       mostrar_log = False
       If UBound(ArrParametros) > 1 Then
            mostrar_log = True
       End If
       
       Call Generar_Archivo(intnro, empresa, Fecha_Calculo, mostrar_log)
    Else
        Flog.writeline Espacios(Tabulador * 0) & "No se encontraron los datos del proceso nro " & NroProceso & " de tipo ."
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
    Flog.writeline "Ultima SQL ejecutada: " & StrSql
End Sub

Private Sub Generar_Archivo(ByVal intnro As Long, ByVal empresa As Long, ByVal Fecha As String, ByVal mostrarLog As Boolean)
' ---------------------------------------------------------------------------------------------
' Descripcion : Procedimiento que genera la exportacion
' Autor       : Fernando Favre
' Fecha       : 13/06/2006
' Ultima Mod  : 23/06/2006
' Descripcion : Se migro el cÛdigo a VB
' ---------------------------------------------------------------------------------------------
'Dim objRs As New ADODB.Recordset
Dim Nombre_Arch As String
Dim NroModelo As Long
Dim Directorio As String
Dim ArchExpExcel
Dim ArchExpTXT
Dim Carpeta
Dim fs1

Dim Sep As String
Dim SepDec As String

Dim cantRegistros As Long

Dim lista_emp As String
Dim cont As Integer
Dim Valor(6) As String
Dim separador As String
Dim intnomarch As String
Dim intfhasta As String
Dim v_linea As Integer
Dim topicorder_ant As Integer

Dim sysid As String
Dim linea As String
Dim v As String

Dim EmpNombre
Dim EmpTer
Dim EmpEstrnro
Dim empcuit

Dim rs As New ADODB.Recordset


Dim rs_topic As New ADODB.Recordset
Dim rs_topic_field As New ADODB.Recordset
Dim rs_field_value2 As New ADODB.Recordset
Dim rs_intconfgen As New ADODB.Recordset
Dim rs_interfaz As New ADODB.Recordset
Dim rs_tt_topicline As New ADODB.Recordset
Dim rs_tt_line As New ADODB.Recordset

' Registro para cada empleado en particular, tambiÈn se usa para conyugue de empleado

Dim rs_emp As New ADODB.Recordset

Dim ternro As Long

' Declaro las variables que van a aparecer en el reporte

Dim l_tidsigla As String
Dim l_nrodoc As String
Dim l_nrodoc1 As String
Dim l_empleg As String
Dim l_tersex As String
Dim l_paisnro As String
Dim l_estcivnro As Integer
Dim l_ternom As String
Dim l_terape As String
Dim l_terfecnac As Date
Dim l_paiscodex As String
Dim l_calle As String
Dim l_nro As String
Dim l_piso As String
Dim l_oficdepto As String
Dim l_provcod_bco As String
Dim l_codigopostal As String
Dim l_telnro As String
Dim l_telnro1 As String
Dim l_locdesc As String
Dim l_monto1
Dim l_monto2
Dim l_monto3
Dim l_monto4
Dim l_monto5
Dim l_monto6
Dim l_bono
Dim l_ticket
Dim l_domnro As Integer
Dim l_telefono As String
Dim l_empfaltagr As Date
Dim l_extciv_bco
Dim l_paiscodext
Dim Longitud As Integer

' Datos del conyuge

Dim l_tidsigla1 As String
Dim l_nrodoc2 As String
Dim l_ternom1 As String
Dim l_terape1 As String
Dim l_terfecnac1 As String
Dim l_paiscodext1 As String
Dim conyuge As Boolean

  On Error GoTo ME_Local
      
    'Directorio de exportacion
    StrSql = "SELECT sis_dirsalidas FROM sistema WHERE sisnro = 1 "
    If rs.State = adStateOpen Then rs.Close
    OpenRecordset StrSql, rs
    If rs.EOF Then
        Flog.writeline Espacios(Tabulador * 0) & "No se encontro el registro 1 de la tabla sistema."
    Else
        Directorio = Trim(rs!sis_dirsalidas)
    End If
    rs.Close
    
    'Creo el archivo HSBC.txt
    Nombre_Arch = Directorio & "\ExpHSBC\HSBC.txt"
    Flog.writeline Espacios(Tabulador * 1) & "Se crea el archivo: " & Nombre_Arch
    Set fs = CreateObject("Scripting.FileSystemObject")
    'desactivo el manejador de errores
    On Error Resume Next
    
    Set Carpeta = fs.getFolder(Directorio)
    If Err.Number <> 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "La carpeta Destino no existe. Se crear·."
        Err.Number = 0
        Set Carpeta = fs.CreateFolder(Directorio)
        
        If Err.Number <> 0 Then
            Flog.writeline Espacios(Tabulador * 1) & "No se puede crear el directorio. Verifique los derechos de acceso o puede crearlo."
            GoTo Fin
        End If
    End If
    
    Set Carpeta = fs.getFolder(Directorio & "\ExpHSBC")
    If Err.Number <> 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "El directorio " & Directorio & "\ExpHSBC no existe. Se crear·."
        Err.Number = 0
        Set Carpeta = fs.CreateFolder(Directorio & "\ExpHSBC")
        
        If Err.Number <> 0 Then
            Flog.writeline Espacios(Tabulador * 1) & "No se puede crear el directorio. Verifique los derechos de acceso o puede crearlo."
            GoTo Fin
        End If
    End If

    Set ArchExpTXT = fs.CreateTextFile(Nombre_Arch, True)
    
    If Err.Number <> 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "No se puede crear el archivo " & Nombre_Arch & " en el directorio " & Directorio & "\HSBC "
        GoTo Fin
    End If

    On Error GoTo ME_Local
    
    
    'Creo el archivo HSBC.csv
    Nombre_Arch = Directorio & "\ExpHSBC\HSBC.csv"
    Flog.writeline Espacios(Tabulador * 1) & "Se crea el archivo: " & Nombre_Arch
'    Set fs = CreateObject("Scripting.FileSystemObject")
    'desactivo el manejador de errores
    On Error Resume Next
    
    Set Carpeta = fs.getFolder(Directorio)
    If Err.Number <> 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "La carpeta Destino no existe. Se crear·."
        Err.Number = 0
        Set Carpeta = fs.CreateFolder(Directorio)
        
        If Err.Number <> 0 Then
            Flog.writeline Espacios(Tabulador * 1) & "No se puede crear el directorio. Verifique los derechos de acceso o puede crearlo."
            GoTo Fin
        End If
    End If
    
    Set Carpeta = fs.getFolder(Directorio & "\ExpHSBC")
    If Err.Number <> 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "El directorio " & Directorio & "\ExpHSBC no existe. Se crear·."
        Err.Number = 0
        Set Carpeta = fs.CreateFolder(Directorio & "\ExpHSBC")
        
        If Err.Number <> 0 Then
            Flog.writeline Espacios(Tabulador * 1) & "No se puede crear el directorio. Verifique los derechos de acceso o puede crearlo."
            GoTo Fin
        End If
    End If

    Set ArchExpExcel = fs.CreateTextFile(Nombre_Arch, True)
    
    If Err.Number <> 0 Then
        Flog.writeline Espacios(Tabulador * 1) & "No se puede crear el archivo " & Nombre_Arch & " en el directorio " & Directorio & "\ExpHSBC "
        GoTo Fin
    End If

    On Error GoTo ME_Local
    
    
    StrSql = "SELECT * FROM empresa "
    StrSql = StrSql & " WHERE empresa.empnro = " & empresa
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        EmpNombre = Mid(rs!empnom, 1, 40)
        EmpTer = rs!ternro
        EmpEstrnro = rs!estrnro
    Else
        Flog.writeline Espacios(Tabulador * 1) & "No se encontro la empresa."
        GoTo Fin
    End If
    rs.Close
    
    ' CUIT de la empresa
    StrSql = "SELECT nrodoc FROM ter_doc WHERE ter_doc.ternro = " & EmpTer & " AND ter_doc.tidnro = 6"
    OpenRecordset StrSql, rs
    If rs.EOF Then
        Flog.writeline Espacios(Tabulador * 1) & "No se encontrÛ el CUIT de la Empresa."
        empcuit = String(11, " ")
    Else
        empcuit = Replace(rs!nrodoc, "-", "")
    End If
    rs.Close
    
'    Fecha = Date
    
    StrSql = "SELECT * FROM empleado "
    StrSql = StrSql & " INNER JOIN tercero ON empleado.ternro = tercero.ternro "
    StrSql = StrSql & " INNER JOIN his_estructura ON empleado.ternro = his_estructura.ternro "
    StrSql = StrSql & " WHERE his_estructura.htetdesde <=" & ConvFecha(Fecha) & " AND "
    StrSql = StrSql & " (his_estructura.htethasta >= " & ConvFecha(Fecha)
    StrSql = StrSql & " OR his_estructura.htethasta IS NULL)"
    StrSql = StrSql & " AND his_estructura.tenro = 10"
    StrSql = StrSql & " AND his_estructura.estrnro = " & EmpEstrnro
    StrSql = StrSql & " AND empleado.empest = -1 "
    StrSql = StrSql & " ORDER BY empleg"
    OpenRecordset StrSql, rs
    
    If rs.EOF Then
        Flog.writeline Espacios(Tabulador * 1) & "No se encontraron empleados."
    Else
        cantRegistros = rs.RecordCount
        Progreso = 0
        If cantRegistros = 0 Then
            cantRegistros = 1
            Flog.writeline Espacios(Tabulador * 1) & " No hay Empleados para procesar."
        End If
        IncPorc = (99 / cantRegistros)
        
        Flog.writeline Espacios(Tabulador * 1) & "==> SQL de los empleados a mostrar"
        Flog.writeline Espacios(Tabulador * 1) & StrSql
        Flog.writeline " "
        Flog.writeline Espacios(Tabulador * 1) & "Creando registro de cabecera."
        
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Graba la de cabecera
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        ArchExpTXT.Write "   "
        ArchExpTXT.Write Format(Fecha, "yyyymmdd")
        ArchExpTXT.Write empcuit
        ArchExpTXT.Write " "
        ArchExpTXT.Write String(6 - Len(CStr(cantRegistros)), "0")
        ArchExpTXT.Write cantRegistros
        ArchExpTXT.Write EmpNombre
        ArchExpTXT.Write String(40 - Len(CStr(EmpNombre)), " ")
        ArchExpTXT.Write String(375, " ")
        ArchExpTXT.writeline
    
        rs.MoveFirst
        
        While Not rs.EOF
          
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' Busco a cada empleado particular
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          
            ternro = rs!ternro
          
          
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' Calculo los datos de los empleados que saco de tercero
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            l_empleg = rs!empleg
            Flog.writeline Espacios(Tabulador * 1) & "Empleado " & l_empleg
            
            If Len(Str(l_empleg)) > 6 Then
                l_empleg = Mid(l_empleg, 1, 6)
                Flog.writeline Espacios(Tabulador * 2) & "La longitud del legajo es mayor a 6 caracteres. Se truncara a " & l_empleg
            End If
            
            l_ternom = "" & rs!ternom & " " & rs!ternom2
            If Len(l_ternom) > 40 Then
                l_ternom = Mid(l_ternom, 1, 40)
            End If
          
            l_terape = "" & rs!terape & " " & rs!terape2
            If Len(l_terape) > 40 Then
                l_terape = Mid(l_terape, 1, 40)
            End If
          
            l_tersex = rs!tersex
            l_estcivnro = "" & rs!estcivnro
            l_terfecnac = "" & rs!terfecnac
            l_empfaltagr = "" & rs!empfaltagr
            l_paisnro = CInt("0" & rs!paisnro)
          
          
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' Obtengo el numero y tipo de documento
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            StrSql = "SELECT ter_doc.nrodoc,tipodocu.tidsigla FROM ter_doc inner join tipodocu ON ter_doc.tidnro"
            StrSql = StrSql & "=tipodocu.tidnro WHERE ter_doc.ternro = " & ternro & " and " & "tipodocu.tidnro<=3"
          
            If rs_emp.State = adStateOpen Then rs_emp.Close
            OpenRecordset StrSql, rs_emp
            If rs_emp.EOF Then
                Flog.writeline Espacios(Tabulador * 2) & "No se encontro el nro o el tipo de documento. SQL ==> " & StrSql
                l_tidsigla = ""
                l_nrodoc = ""
            Else
                l_tidsigla = "" & rs_emp!tidsigla
                l_nrodoc = rs_emp!nrodoc
            End If
            l_nrodoc = Replace(l_nrodoc, ".", "")
            rs_emp.Close
            
            If l_tidsigla = "" Then
                l_tidsigla = String(2, " ")
            End If
            If Len(l_tidsigla) > 3 Then
                l_tidsigla = Mid(l_tidsigla, 1, 3)
            End If
            
            If l_nrodoc = "" Then
                l_nrodoc = String(8, " ")
            End If
            If Len(l_nrodoc) > 8 Then
                l_nrodoc = Mid(l_nrodoc, 1, 8)
            End If
            
            If mostrarLog Then
                Flog.writeline Espacios(Tabulador * 3) & " Ya calculo el Tipo y numero de documento del empleado " & l_empleg
            End If
            
            
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' Obtengo el cuil
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            StrSql = "select * from ter_doc where tidnro=10 and ternro=" & ternro
          
            If rs_emp.State = adStateOpen Then rs_emp.Close
            OpenRecordset StrSql, rs_emp
            If rs_emp.EOF Then
                Flog.writeline Espacios(Tabulador * 2) & "No se encontro el cuil. SQL ==> " & StrSql
                l_nrodoc1 = ""
            Else
                l_nrodoc1 = Replace(rs_emp!nrodoc, "-", "")
            End If
            If l_nrodoc1 = "" Then
                l_nrodoc1 = String(11, " ")
            End If
         
            If Len(l_nrodoc1) > 11 Then
                Flog.Write Espacios(Tabulador * 2) & "La longitud del Nro documento es mayor a 11 caracteres. Se truncara de " & l_nrodoc1
                l_nrodoc1 = Mid(l_nrodoc1, 1, 11)
                Flog.writeline Espacios(Tabulador * 2) & " a " & l_nrodoc1
            End If
         
            rs_emp.Close
         
            If mostrarLog Then
                Flog.writeline Espacios(Tabulador * 3) & " Ya calculo el cuil del empleado " & l_empleg
            End If
         
         
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' Obtengo la nacionalidad
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            StrSql = "select paiscodext from pais where paisnro=" & l_paisnro
          
            If rs_emp.State = adStateOpen Then rs_emp.Close
            OpenRecordset StrSql, rs_emp
            
            If rs_emp.EOF Then
                Flog.writeline Espacios(Tabulador * 2) & "No se encontro cod. externro del pais. SQL ==> " & StrSql
                l_paiscodext = ""
            Else
                l_paiscodext = "" & rs_emp!paiscodext
            End If
            rs_emp.Close
            If Len(l_paiscodext) > 3 Then
                Flog.Write Espacios(Tabulador * 2) & "La longitud de la nacionalidad es mayor a 3 caracteres. Se truncara de " & l_paiscodext
                l_paiscodext = Mid(l_paiscodext, 1, 3)
                Flog.writeline Espacios(Tabulador * 2) & " a " & l_paiscodext
            End If
         
            If mostrarLog Then
                Flog.writeline Espacios(Tabulador * 3) & " Ya calculo la nacionalidad del empleado " & l_empleg
            End If
         
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' Obtengo el domicilio del empleado
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            StrSql = "select * from detdom inner join cabdom on detdom.domnro="
            StrSql = StrSql & "cabdom.domnro and cabdom.ternro=" & ternro
            StrSql = StrSql & " inner join provincia on detdom.provnro=provincia.provnro "
            StrSql = StrSql & " inner join localidad on detdom.locnro=localidad.locnro"
            If rs_emp.State = adStateOpen Then rs_emp.Close
            OpenRecordset StrSql, rs_emp
            If rs_emp.EOF Then
                Flog.writeline Espacios(Tabulador * 2) & "No se encontraron el domicilio. SQL ==> " & StrSql
                l_domnro = 0
                l_calle = ""
                l_nro = ""
                l_piso = ""
                l_oficdepto = ""
                l_provcod_bco = ""
                l_codigopostal = ""
                l_locdesc = ""
            Else
                l_calle = "" & rs_emp!calle
                l_calle = Replace(l_calle, ",", "")
                l_nro = "" & rs_emp!nro
                l_piso = "" & rs_emp!piso
                l_oficdepto = "" & rs_emp!oficdepto
                l_provcod_bco = "" & rs_emp!provcod_bco
                l_codigopostal = "" & rs_emp!codigoPostal
                l_locdesc = "" & rs_emp!locdesc
                l_domnro = rs_emp!domnro
            End If
            rs_emp.Close
                    
            If Len(l_calle) > 60 Then l_calle = Mid(l_calle, 1, 60)
            If Len(l_nro) > 5 Then l_nro = Mid(l_nro, 1, 5)
            If Len(l_piso) > 2 Then l_piso = Mid(l_piso, 1, 2)
            If Len(l_oficdepto) > 2 Then l_oficdepto = Mid(l_oficdepto, 1, 2)
            If Len(l_provcod_bco) > 1 Then l_provcod_bco = Mid(l_provcod_bco, 1, 1)
            If l_piso = "" Then l_piso = "00"
            If l_oficdepto = "" Then l_oficdepto = "00"
            If Len(l_codigopostal) > 8 Then
                l_codigopostal = Mid(l_codigopostal, 1, 8)
            Else
                If l_codigopostal = "" Then l_codigopostal = "1005"
            End If
            If l_provcod_bco = "" Then l_provcod_bco = " "
            If Len(l_locdesc) > 70 Then l_locdesc = Mid(l_locdesc, 1, 70)
            l_locdesc = Replace(l_locdesc, "(", " ")
            l_locdesc = Replace(l_locdesc, ")", " ")
            
            Call DejarSoloNumeros(l_nro)
            Call DejarSoloNumeros(l_piso)
            Call DejarSoloNumeros(l_codigopostal)
            
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' Obtengo el telÈfono del empleado
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            l_telnro = ""
            l_telnro1 = ""
            If l_domnro <> 0 Then
                StrSql = "select * from telefono where domnro=" & l_domnro
                 
                If rs_emp.State = adStateOpen Then rs_emp.Close
                OpenRecordset StrSql, rs_emp
                
                If Not rs_emp.EOF Then
                    l_telnro1 = Replace(rs_emp!telnro, "-", "")
                    l_telnro1 = Replace(l_telnro1, "(", "")
                    l_telnro1 = Replace(l_telnro1, ")", "")
                    l_telnro1 = Replace(l_telnro1, " ", "")
                    l_telnro1 = Replace(l_telnro1, "/", "")
                End If
                rs_emp.Close
                
            End If
            
            
            If l_telnro1 = "" Then l_telnro1 = String(8, "0")
            Longitud = Len(l_telnro1)
            If Longitud >= 8 Then
                If Longitud - 8 > 0 Then
                    l_telnro = Mid(l_telnro1, 1, Longitud - 8)
                    If Len(l_telnro) > 5 Then
                        l_telnro = Mid(l_telnro, 1, 5)
                    End If
                End If
                l_telnro1 = Mid(l_telnro1, Longitud - 7, 8)
            End If
            
            Call DejarSoloNumeros(l_telnro1)
            Call DejarSoloNumeros(l_telnro)
            
            If mostrarLog Then
                Flog.writeline Espacios(Tabulador * 3) & " Ya calculo el domicilio y telefono del empleado " & l_empleg
            End If
            
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' Obtengo el  extciv del empleado
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            StrSql = "select extciv_bco from estcivil where estcivnro=" & l_estcivnro
             
            If rs_emp.State = adStateOpen Then rs_emp.Close
            OpenRecordset StrSql, rs_emp
         
            If rs_emp.EOF Then
                Flog.writeline Espacios(Tabulador * 2) & "No se encontro datos del estado civil. SQL ==> " & StrSql
                l_extciv_bco = "0"
            Else
                l_extciv_bco = "" & rs_emp!extciv_bco
                If l_extciv_bco = "" Then l_extciv_bco = "0"
            End If
            rs_emp.Close
         
            
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' Obtengo los ˙ltimos 6 ingresos mensuales
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            StrSql = "select * from acu_mes where acunro=6 and ammes=11 and amanio=2004 and"
            StrSql = StrSql & " ternro =" & ternro
             
            If rs_emp.State = adStateOpen Then rs_emp.Close
            OpenRecordset StrSql, rs_emp
            l_monto1 = 0
            l_monto2 = 0
            l_monto3 = 0
            l_monto4 = 0
            l_monto5 = 0
            l_monto6 = 0
            If rs_emp.EOF Then
                Flog.writeline Espacios(Tabulador * 2) & "No se encontro datos de los ingresos. SQL ==> " & StrSql
            Else
                l_monto1 = Format(CDbl("0" & rs_emp!ammonto), "000000")
                If Len(l_monto1) > 6 Then
                    Flog.writeline Espacios(Tabulador * 2) & "Error. El monto (sueldo) " & l_monto1 & " es mayor a 6 caracteres. En el archivo se imprime en 6 posiciones."
                    GoTo ME_Local
                End If
                l_monto2 = l_monto1
                l_monto3 = l_monto1
                l_monto4 = l_monto1
                l_monto5 = l_monto1
                l_monto6 = l_monto1
            End If
            rs_emp.Close
                        
             
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' Ultima BonificaciÛn
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            StrSql = "select * from acu_liq inner join cabliq on cabliq.empleado=" & ternro
            StrSql = StrSql & " and cabliq.pronro=484 and acu_liq.acunro = 6"
            StrSql = StrSql & " and cabliq.cliqnro = acu_liq.cliqnro"
            
            If rs_emp.State = adStateOpen Then rs_emp.Close
            OpenRecordset StrSql, rs_emp
            l_bono = 0
            If rs_emp.EOF Then
                Flog.writeline Espacios(Tabulador * 2) & "No se encontro datos de las bonificacuones. SQL ==> " & StrSql
            Else
                l_bono = Format(CDbl("0" & rs_emp!almonto), "000000")
                If Len(l_bono) > 6 Then
                    Flog.writeline Espacios(Tabulador * 2) & "Error. El monto (bonificacion) " & l_bono & " es mayor a 6 caracteres. En el archivo se imprime en 6 posiciones."
                    GoTo ME_Local
                End If
            End If
            rs_emp.Close
                    
          
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' Tickets
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            StrSql = "select * from proceso inner join cabliq on "
            StrSql = StrSql & "proceso.pronro=cabliq.pronro and proceso.pliqnro=75 "
            StrSql = StrSql & "inner join acu_liq on cabliq.empleado=" & ternro
            StrSql = StrSql & " and acu_liq.acunro = 124"
            StrSql = StrSql & " and acu_liq.cliqnro = cabliq.cliqnro"
            
            If rs_emp.State = adStateOpen Then rs_emp.Close
            OpenRecordset StrSql, rs_emp
            l_ticket = 0
            If rs_emp.EOF Then
                Flog.writeline Espacios(Tabulador * 2) & "No se encontro datos de los ticket. SQL ==> " & StrSql
            Else
                l_ticket = Format(CDbl("0" & rs_emp!almonto), "000000")
                If Len(l_ticket) > 6 Then
                    Flog.writeline Espacios(Tabulador * 2) & "Error. El monto (ticket) " & l_ticket & " es mayor a 6 caracteres. En el archivo se imprime en 6 posiciones."
                    GoTo ME_Local
                End If
            End If
            rs_emp.Close
                    
         
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'InformaciÛn del conyuge
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' Busco al conyuge
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            l_ternom1 = ""
            l_terape1 = ""
            l_tidsigla1 = ""
            l_nrodoc2 = ""
            l_terfecnac1 = ""
            l_paiscodext1 = ""
            
            StrSql = "SELECT * FROM familiar WHERE empleado = " & ternro
            StrSql = StrSql & " and parenro=3"
            
            If rs_emp.State = adStateOpen Then rs_emp.Close
            OpenRecordset StrSql, rs_emp
            If rs_emp.EOF Then
                Flog.writeline Espacios(Tabulador * 2) & "No esta casado. SQL ==> " & StrSql
                conyuge = False
            Else
                ternro = rs_emp!ternro
                conyuge = True
            
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ' Calculo los datos de los conyugues
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                StrSql = "select * from tercero inner join pais on tercero.paisnro=pais.paisnro and tercero.ternro=" & ternro
                If rs_emp.State = adStateOpen Then rs_emp.Close
                OpenRecordset StrSql, rs_emp
                If rs_emp.EOF Or Not (conyuge) Then
                    Flog.writeline Espacios(Tabulador * 2) & "No se encontro los datos del conyuge. SQL ==> " & StrSql
                Else
                    l_ternom1 = "" & rs_emp!ternom & " " & rs_emp!ternom2
                    l_terape1 = "" & rs_emp!terape & " " & rs_emp!terape2
                    l_paiscodext1 = "" & rs_emp!paiscodext
                    l_terfecnac1 = "" & rs_emp!terfecnac
                End If
            
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ' Obtengo el numero y tipo de documento
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                StrSql = "SELECT ter_doc.nrodoc,tipodocu.tidsigla FROM ter_doc inner join tipodocu ON ter_doc.tidnro"
                StrSql = StrSql & "=tipodocu.tidnro WHERE ter_doc.ternro = " & ternro & " and " & "tipodocu.tidnro<=3"
                               
                If rs_emp.State = adStateOpen Then rs_emp.Close
                OpenRecordset StrSql, rs_emp
                
                If rs_emp.EOF Or Not (conyuge) Then
                    Flog.writeline Espacios(Tabulador * 2) & "No se encontro el nro o el tipo de documento del conyugue. SQL ==> " & StrSql
                Else
                    l_tidsigla1 = "" & rs_emp!tidsigla
                    l_nrodoc2 = CStr(rs_emp!nrodoc)
                End If
            End If
            
            l_ternom1 = Mid(l_ternom1, 1, 40)
            l_terape1 = Mid(l_terape1, 1, 40)
            l_tidsigla1 = Mid(l_tidsigla1, 1, 3)
            l_nrodoc2 = Mid(l_nrodoc2, 1, 8)
            l_paiscodext1 = Mid(l_paiscodext1, 1, 3)
            
            If l_ternom1 = "" Then l_ternom1 = String(40, "0")
            If l_terape1 = "" Then l_terape1 = String(40, "0")
            If l_paiscodext1 = "" Then l_paiscodext1 = "ARG"
            
            If l_tidsigla1 = "" Then l_tidsigla1 = "DNI"
            If l_nrodoc2 = "" Then l_nrodoc2 = String(8, "0")
             
            rs_emp.Close
            
            If mostrarLog Then
                Flog.writeline Espacios(Tabulador * 3) & " Calculo los valores del conjuge para el empleado " & l_empleg
            End If
          
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'Genero el txt
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ArchExpTXT.Write l_tidsigla
            ArchExpTXT.Write String(3 - Len(l_tidsigla), " ")
            ArchExpTXT.Write l_nrodoc
            ArchExpTXT.Write String(8 - Len(l_nrodoc), " ")
            ArchExpTXT.Write l_nrodoc1
            ArchExpTXT.Write String(11 - Len(l_nrodoc1), " ")
            
            If mostrarLog Then
                Flog.writeline Espacios(Tabulador * 3) & " Ya imprimio en el archivo de texto el Tipo, nro Doc y cuil"
            End If
            
            If l_tersex = -1 Then
                ArchExpTXT.Write "1"
            Else
                ArchExpTXT.Write "2"
            End If
            ArchExpTXT.Write String(6 - Len(l_empleg), "0")
            ArchExpTXT.Write l_empleg
            ArchExpTXT.Write l_ternom
            ArchExpTXT.Write String(40 - Len(l_ternom), " ")
            ArchExpTXT.Write l_terape
            ArchExpTXT.Write String(40 - Len(l_terape), " ")
            
            If mostrarLog Then
                Flog.writeline Espacios(Tabulador * 3) & " Ya imprimio sexo, legajo, nombre y apellido"
            End If
            
            If Not EsNulo(l_terfecnac) Then
                ArchExpTXT.Write String(4 - Len(Year(CDate(l_terfecnac))), "0")
                ArchExpTXT.Write Year(CDate(l_terfecnac))
                
                ArchExpTXT.Write String(2 - Len(Month(CDate(l_terfecnac))), "0")
                ArchExpTXT.Write Month(CDate(l_terfecnac))
                
                ArchExpTXT.Write String(2 - Len(Day(CDate(l_terfecnac))), "0")
                ArchExpTXT.Write Day(CDate(l_terfecnac))
            Else
                ArchExpTXT.Write "00000000"
            End If
            
            If mostrarLog Then
                Flog.writeline Espacios(Tabulador * 3) & " Ya imprimio la fecha de nacimiento"
            End If
            
            ArchExpTXT.Write l_paiscodext
            ArchExpTXT.Write String(3 - Len(l_paiscodext), " ")
            
            If l_extciv_bco = "" Then
                ArchExpTXT.Write " "
            Else
                ArchExpTXT.Write Mid(l_extciv_bco, 1, 1)
            End If
'            ArchExpTXT.Write l_extciv_bco
'            ArchExpTXT.Write String(1 - Len(l_extciv_bco), " ")
            
            ArchExpTXT.Write l_calle
            ArchExpTXT.Write String(60 - Len(l_calle), " ")
            
            ArchExpTXT.Write l_nro
            ArchExpTXT.Write String(5 - Len(l_nro), " ")
            
            ArchExpTXT.Write l_piso
            ArchExpTXT.Write String(2 - Len(l_piso), " ")
            
            ArchExpTXT.Write l_oficdepto
            ArchExpTXT.Write String(2 - Len(l_oficdepto), " ")
            
            If mostrarLog Then
                Flog.writeline Espacios(Tabulador * 3) & " Ya imprimio Pais (paiscodext), el estado civil (extciv_bco), calle, nro, piso y oficina"
            End If
            
            If l_telnro <> "" Then
                ArchExpTXT.Write l_telnro
                ArchExpTXT.Write String(5 - Len(l_telnro), " ")
            Else
                ArchExpTXT.Write String(5, "0")
            End If
            
            ArchExpTXT.Write String(8 - Len(l_telnro1), "0")
            ArchExpTXT.Write l_telnro1
            
            ArchExpTXT.Write l_locdesc
            ArchExpTXT.Write String(70 - Len(l_locdesc), " ")
            
            ArchExpTXT.Write l_provcod_bco
            ArchExpTXT.Write String(1 - Len(l_provcod_bco), "0")
            
            ArchExpTXT.Write l_codigopostal
            ArchExpTXT.Write String(8 - Len(l_codigopostal), " ")
            
            If mostrarLog Then
                Flog.writeline Espacios(Tabulador * 3) & " Ya imprimio el TE, localidad, provincia (provcod_bco) y codigo postal"
            End If
            
            ArchExpTXT.Write "001"
                    
            ArchExpTXT.Write String(6 - Len(CStr(l_monto1)), "0")
            ArchExpTXT.Write l_monto1
            
            ArchExpTXT.Write String(6 - Len(CStr(l_monto2)), "0")
            ArchExpTXT.Write l_monto2
            
            ArchExpTXT.Write String(6 - Len(CStr(l_monto3)), "0")
            ArchExpTXT.Write l_monto3
            
            ArchExpTXT.Write String(6 - Len(CStr(l_monto4)), "0")
            ArchExpTXT.Write l_monto4
            
            ArchExpTXT.Write String(6 - Len(CStr(l_monto5)), "0")
            ArchExpTXT.Write l_monto5
                    
            ArchExpTXT.Write String(6 - Len(CStr(l_monto6)), "0")
            ArchExpTXT.Write l_monto6
            
            ArchExpTXT.Write String(6 - Len(CStr(l_bono)), "0")
            ArchExpTXT.Write l_bono
            
            ArchExpTXT.Write String(6 - Len(CStr(l_ticket)), "0")
            ArchExpTXT.Write l_ticket
            
            If mostrarLog Then
                Flog.writeline Espacios(Tabulador * 3) & " Ya imprimio el 001, monto1, monto2, monto3, monto4, monto5, monto6, bono y ticket"
            End If
            
            If Not EsNulo(l_empfaltagr) Then
                ArchExpTXT.Write String(4 - Len(Year(CDate(l_empfaltagr))), "0")
                ArchExpTXT.Write Year(CDate(l_empfaltagr))
                
                ArchExpTXT.Write String(2 - Len(Month(CDate(l_empfaltagr))), "0")
                ArchExpTXT.Write Month(CDate(l_empfaltagr))
                
                ArchExpTXT.Write String(2 - Len(Day(CDate(l_empfaltagr))), "0")
                ArchExpTXT.Write Day(CDate(l_empfaltagr))
            Else
                ArchExpTXT.Write "00000000"
            End If
            
            ' Datos del conyuge
            
            If mostrarLog Then
                Flog.writeline Espacios(Tabulador * 3) & " Ya imprimio la fecha de alta"
            End If
            
            
            ArchExpTXT.Write "N"
            
            ArchExpTXT.Write l_tidsigla1
            ArchExpTXT.Write String(3 - Len(l_tidsigla1), " ")
            ArchExpTXT.Write l_nrodoc2
            ArchExpTXT.Write String(8 - Len(l_nrodoc2), " ")
            
            If mostrarLog Then
                Flog.writeline Espacios(Tabulador * 3) & " Ya imprimio el Tipo y Doc "
            End If
            
            If conyuge And Not EsNulo(l_terfecnac1) Then
                ArchExpTXT.Write String(4 - Len(Year(CDate(l_terfecnac1))), "0")
                ArchExpTXT.Write Year(CDate(l_terfecnac1))
            
                ArchExpTXT.Write String(2 - Len(Month(CDate(l_terfecnac1))), "0")
                ArchExpTXT.Write Month(CDate(l_terfecnac1))
            
                ArchExpTXT.Write String(2 - Len(Day(CDate(l_terfecnac1))), "0")
                ArchExpTXT.Write Day(CDate(l_terfecnac1))
            Else
                ArchExpTXT.Write "19500101"
            End If
            
            If mostrarLog Then
                Flog.writeline Espacios(Tabulador * 3) & " Ya imprimio la fecha de nacimiento del conjuge"
            End If
            
            ArchExpTXT.Write String(3 - Len(l_paiscodext1), " ")
            ArchExpTXT.Write l_paiscodext1
            
            ArchExpTXT.Write l_ternom1
            ArchExpTXT.Write String(40 - Len(l_ternom1), " ")
            ArchExpTXT.Write l_terape1
            ArchExpTXT.Write String(40 - Len(l_terape1), " ")
            ArchExpTXT.writeline
             
    
            If mostrarLog Then
                Flog.writeline Espacios(Tabulador * 3) & " Ya imprimio el pais, nombre y apellido del conjuge"
            End If
            
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'Genero el archivo excel
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            If mostrarLog Then
                Flog.writeline Espacios(Tabulador * 3) & "Fin archivo txt. Comienzo excel"
            End If
            
            ArchExpExcel.Write l_tidsigla
            ArchExpExcel.Write String(3 - Len(l_tidsigla), " ")
            ArchExpExcel.Write ","
            
            ArchExpExcel.Write l_nrodoc
            ArchExpExcel.Write String(8 - Len(l_nrodoc), " ")
            ArchExpExcel.Write ","
            
            ArchExpExcel.Write l_nrodoc1
            ArchExpExcel.Write String(11 - Len(l_nrodoc1), " ")
            ArchExpExcel.Write ","
            
            If mostrarLog Then
                Flog.writeline Espacios(Tabulador * 3) & " Ya imprimio el Tipo, nro Doc y cuil en excel"
            End If
            
            If l_tersex = -1 Then
                ArchExpExcel.Write "1"
            Else
                ArchExpExcel.Write "2"
            End If
            ArchExpExcel.Write ","
            
            ArchExpExcel.Write String(6 - Len(l_empleg), "0")
            ArchExpExcel.Write l_empleg
            ArchExpExcel.Write ","
            
            ArchExpExcel.Write l_ternom
            ArchExpExcel.Write String(40 - Len(l_ternom), " ")
            ArchExpExcel.Write ","
            
            ArchExpExcel.Write l_terape
            ArchExpExcel.Write String(40 - Len(l_terape), " ")
            ArchExpExcel.Write ","
            
            If mostrarLog Then
                Flog.writeline Espacios(Tabulador * 3) & " Ya imprimio sexo, legajo, nombre y apellido en excel"
            End If
            
            If Not EsNulo(l_terfecnac) Then
                ArchExpExcel.Write String(4 - Len(Year(CDate(l_terfecnac))), "0")
                ArchExpExcel.Write Year(CDate(l_terfecnac))
                ArchExpExcel.Write ","
                
                ArchExpExcel.Write String(2 - Len(Month(CDate(l_terfecnac))), "0")
                ArchExpExcel.Write Month(CDate(l_terfecnac))
                ArchExpExcel.Write ","
                
                ArchExpExcel.Write String(2 - Len(Day(CDate(l_terfecnac))), "0")
                ArchExpExcel.Write Day(CDate(l_terfecnac))
                ArchExpExcel.Write ","
            Else
                ArchExpExcel.Write "0000,00,00"
            End If
            
            If mostrarLog Then
                Flog.writeline Espacios(Tabulador * 3) & " Ya imprimio la fecha de nacimiento en excel"
            End If
            
            ArchExpExcel.Write l_paiscodext
            ArchExpExcel.Write String(3 - Len(l_paiscodext), " ")
            ArchExpExcel.Write ","
            
            If l_extciv_bco = "" Then
                ArchExpExcel.Write " "
            Else
                ArchExpExcel.Write Mid(l_extciv_bco, 1, 1)
            End If
            'ArchExpExcel.Write l_extciv_bco
            'ArchExpExcel.Write String(1 - Len(l_extciv_bco), " ")
            ArchExpExcel.Write ","
            
            ArchExpExcel.Write l_calle
            ArchExpExcel.Write String(60 - Len(l_calle), " ")
            ArchExpExcel.Write ","
            
            ArchExpExcel.Write l_nro
            ArchExpExcel.Write String(5 - Len(l_nro), " ")
            ArchExpExcel.Write ","
            
            ArchExpExcel.Write l_piso
            ArchExpExcel.Write String(2 - Len(l_piso), " ")
            ArchExpExcel.Write ","
            
            ArchExpExcel.Write l_oficdepto
            ArchExpExcel.Write String(2 - Len(l_oficdepto), " ")
            ArchExpExcel.Write ","
            
            If l_telnro <> "" Then
                ArchExpExcel.Write l_telnro
                ArchExpExcel.Write String(5 - Len(l_telnro), " ")
            Else
                ArchExpExcel.Write String(5, "0")
            End If
            
            ArchExpExcel.Write ","
            
            If mostrarLog Then
                Flog.writeline Espacios(Tabulador * 3) & " Ya imprimio Pais (paiscodext), el estado civil (extciv_bco), calle, nro, piso y oficina. En excel"
            End If
            
            ArchExpExcel.Write String(8 - Len(l_telnro1), "0")
            ArchExpExcel.Write l_telnro1
            ArchExpExcel.Write ","
            
            ArchExpExcel.Write l_locdesc
            ArchExpExcel.Write String(70 - Len(l_locdesc), " ")
            ArchExpExcel.Write ","
            
            ArchExpExcel.Write l_provcod_bco
            ArchExpExcel.Write String(1 - Len(l_provcod_bco), "0")
            ArchExpExcel.Write ","
            
            ArchExpExcel.Write l_codigopostal
            ArchExpExcel.Write String(8 - Len(l_codigopostal), " ")
            
            
            If mostrarLog Then
                Flog.writeline Espacios(Tabulador * 3) & " Ya imprimio el TE, localidad, provincia (provcod_bco) y codigo postal. En excel"
            End If
            
            
            ArchExpExcel.Write ",001,"
            
                    
            ArchExpExcel.Write String(6 - Len(CStr(l_monto1)), "0")
            ArchExpExcel.Write l_monto1
            ArchExpExcel.Write ","
            
            ArchExpExcel.Write String(6 - Len(CStr(l_monto2)), "0")
            ArchExpExcel.Write l_monto2
            ArchExpExcel.Write ","
            
            ArchExpExcel.Write String(6 - Len(CStr(l_monto3)), "0")
            ArchExpExcel.Write l_monto3
            ArchExpExcel.Write ","
            
            ArchExpExcel.Write String(6 - Len(CStr(l_monto4)), "0")
            ArchExpExcel.Write l_monto4
            ArchExpExcel.Write ","
            
            ArchExpExcel.Write String(6 - Len(CStr(l_monto5)), "0")
            ArchExpExcel.Write l_monto5
            ArchExpExcel.Write ","
                    
            ArchExpExcel.Write String(6 - Len(CStr(l_monto6)), "0")
            ArchExpExcel.Write l_monto6
            ArchExpExcel.Write ","
            
            ArchExpExcel.Write String(6 - Len(CStr(l_bono)), "0")
            ArchExpExcel.Write l_bono
            ArchExpExcel.Write ","
            
            ArchExpExcel.Write String(6 - Len(CStr(l_ticket)), "0")
            ArchExpExcel.Write l_ticket
            ArchExpExcel.Write ","
            
            If mostrarLog Then
                Flog.writeline Espacios(Tabulador * 3) & " Ya imprimio el 001, monto1, monto2, monto3, monto4, monto5, monto6, bono y ticket. En excel"
            End If
            
            If Not EsNulo(l_empfaltagr) Then
                ArchExpExcel.Write String(4 - Len(Year(CDate(l_empfaltagr))), "0")
                ArchExpExcel.Write Year(CDate(l_empfaltagr))
                ArchExpExcel.Write ","
                
                ArchExpExcel.Write String(2 - Len(Month(CDate(l_empfaltagr))), "0")
                ArchExpExcel.Write Month(CDate(l_empfaltagr))
                ArchExpExcel.Write ","
                
                ArchExpExcel.Write String(2 - Len(Day(CDate(l_empfaltagr))), "0")
                ArchExpExcel.Write Day(CDate(l_empfaltagr))
            Else
                ArchExpExcel.Write "0000,00,00"
            End If
            
            
            If mostrarLog Then
                Flog.writeline Espacios(Tabulador * 3) & " Ya imprimio la fecha de alta en excel"
            End If
            
            ' Datos del conyuge
            
            
            ArchExpExcel.Write ",N,"
            
            ArchExpExcel.Write l_tidsigla1
            ArchExpExcel.Write String(3 - Len(l_tidsigla1), " ")
            ArchExpExcel.Write ","
            
            ArchExpExcel.Write l_nrodoc2
            ArchExpExcel.Write String(8 - Len(l_nrodoc2), " ")
            ArchExpExcel.Write ","
            
            
            If conyuge And Not EsNulo(l_terfecnac1) Then
                ArchExpExcel.Write String(4 - Len(Year(CDate(l_terfecnac1))), "0")
                ArchExpExcel.Write Year(CDate(l_terfecnac1))
                ArchExpExcel.Write ","
                       
                ArchExpExcel.Write String(2 - Len(Month(CDate(l_terfecnac1))), "0")
                ArchExpExcel.Write Month(CDate(l_terfecnac1))
                ArchExpExcel.Write ","
                        
                ArchExpExcel.Write String(2 - Len(Day(CDate(l_terfecnac1))), "0")
                ArchExpExcel.Write Day(CDate(l_terfecnac1))
            Else
                ArchExpExcel.Write "1950,01,01"
            End If
            ArchExpExcel.Write ","
            
            If mostrarLog Then
                Flog.writeline Espacios(Tabulador * 3) & " Ya imprimio el Tipo y Doc y fecha de nacimiento del conjuge en excel"
            End If
            
            ArchExpExcel.Write String(3 - Len(l_paiscodext1), " ")
            ArchExpExcel.Write l_paiscodext1
            ArchExpExcel.Write ","
                
            ArchExpExcel.Write l_ternom1
            ArchExpExcel.Write String(40 - Len(l_ternom1), " ")
            ArchExpExcel.Write ","
            
            ArchExpExcel.Write l_terape1
            ArchExpExcel.Write String(40 - Len(l_terape1), " ")
                    
            If mostrarLog Then
                Flog.writeline Espacios(Tabulador * 3) & " Ya imprimio el pais, nombre y apellido del conjuge en excel"
            End If
            
            ArchExpExcel.writeline
            
            rs.MoveNext
            
            'Actualizo el estado del proceso
            TiempoAcumulado = GetTickCount
            cantRegistros = cantRegistros - 1
            Progreso = Progreso + IncPorc
     
            StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
            StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'"
            StrSql = StrSql & ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
        Wend
    End If
    
    ArchExpExcel.Close
    ArchExpTXT.Close

'progress
'       FIRST tercero OF empleado NO-LOCK
'        BREAK BY empleado.empleg:
'
'    /***** Registro de cabecera *****/
'    ASSIGN cabecera = FILL(" ",3) +                 /* Blancos */
'                      STRING(YEAR(TODAY),"9999") +  /* Fecha de envio */
'                      STRING(MONTH(TODAY),"99") +
'                      STRING(DAY(TODAY),"99") +
'                      cuit +                        /* CUIT */
'                      FILL(" ",1) +                 /* Blancos */
'                      FILL(" ",6) +                 /* Cant. de registros */
'                      STRING(nom-empresa,"x(40)") + /* Nombre de la empresa */
'                      FILL(" ",375).                /* Blancos */
'
'    PUT STREAM salida UNFORMATTED cabecera SKIP.'
'
'    /* Guardo la posicion de escritura para completar */
'    ASSIGN puntero = 23#
'
'    /***** Registro de detalle *****/
'    FOR EACH empleado WHERE empleado.empest AND empleado.empleg < 1000 NO-LOCK,
'        FIRST tercero OF empleado NO-LOCK
'        BREAK BY empleado.empleg:
'
'        /* Cantidad de Registros */
'        ASSIGN cant - Reg = cant - Reg + 1#
'
'        /* Busco el documento del empleado (DNI,LE o LC) */
'        FIND FIRST ter_doc WHERE ter_doc.ternro = empleado.ternro AND
'                                 ter_doc.tidnro <= 3 NO-LOCK NO-ERROR.
'        IF AVAIL ter_doc THEN DO:
'            FIND FIRST tipodocu OF ter_doc NO-LOCK NO-ERROR.
'            IF AVAIL tipodocu THEN
'                ASSIGN sigla = STRING(tipodocu.tidsigla,"x(3)").
'            Else
'                ASSIGN sigla = FILL(" ",3).
'            ASSIGN nrodocu = STRING(REPLACE(ter_doc.nrodoc,".",""),"99999999").
'        END.
'        Else
'            ASSIGN sigla = FILL(" ", 3)
'                   nrodocu = FILL(" ",8).
'
'        /* Busco el CUIL del empleado */
'        FIND FIRST ter_doc WHERE ter_doc.ternro = empleado.ternro AND
'                                 ter_doc.tidnro = 10 NO-LOCK NO-ERROR.
'        IF AVAIL ter_doc THEN
'            ASSIGN cuil = REPLACE(ter_doc.nrodoc,"-","").
'        Else
'            ASSIGN cuil = FILL(" ",11).
'
'        /* Sexo */
'        ASSIGN sexo = IF tercero.tersex THEN "1" ELSE "2".
'
'        /* Pais */
'        FIND FIRST pais OF tercero NO-LOCK NO-ERROR.
'
'        /* Estado Civil */
'        FIND FIRST estcivil OF tercero WHERE estcivil.estcivnro < 6 NO-LOCK NO-ERROR.
'
'        /* Domicilio */
'        FIND FIRST cabdom WHERE cabdom.ternro = empleado.ternro NO-LOCK NO-ERROR.
'        IF AVAIL cabdom THEN DO:
'            FIND FIRST detdom OF cabdom NO-LOCK NO-ERROR.
'            IF AVAIL detdom THEN DO:
'                FIND FIRST telefono OF detdom NO-LOCK NO-ERROR.
'                IF AVAIL telefono THEN DO:
'                    ASSIGN numerotel = Replace(telefono.telnro, "-", "")
'                           numerotel = Replace(numerotel, "(", "")
'                           numerotel = Replace(numerotel, ")", "")
'                           numerotel = Replace(numerotel, " ", "")
'                           numerotel = REPLACE(numerotel,"/","").
'                    If LENGTH(numerotel) > 8 Then
''''''''''''''                  ASSIGN prefijo = SUBSTRING(numerotel, 1, LENGTH(numerotel) - 8)
''''''''''''''                  numtel  = SUBSTRING(numerotel,LENGTH(numerotel) - 7,8).
'                    Else
'                        ASSIGN prefijo = ""
'                               numtel  = numerotel.
'                END.
'                FIND FIRST localidad OF detdom NO-LOCK NO-ERROR.
'                FIND FIRST provincia OF detdom NO-LOCK NO-ERROR.
'
'                ASSIGN calle = String(Replace(detdom.calle, ",", ""), "x(60)")
'                       calle = Replace(calle, "•", "N")
'                       calle = Replace(calle, "§", "n")
'                       calle = Replace(calle, "†", "a")
'                       calle = Replace(calle, "Ç", "e")
'                       calle = Replace(calle, "°", "i")
'                       calle = Replace(calle, "¢", "o")
'                       calle = Replace(calle, "£", "u")
'                       calle = Replace(calle, "µ", "A")
'                       calle = Replace(calle, "ê", "E")
'                       calle = Replace(calle, "÷", "I")
'                       calle = Replace(calle, "‡", "O")
'                       calle = Replace(calle, "È", "u")
'                       calle = Replace(calle, "(", " ")
'                       calle = Replace(calle, ")", " ")
'                       piso = ""
'                       depto = String(detdom.oficdepto, "x(2)")
'                       tel1 = ""
'                       tel2 = ""
'                       local = STRING(IF AVAIL localidad THEN localidad.locdes ELSE FILL(" ",70),"x(70)")
'                       local = REPLACE(local,"("," ")
'                       local = REPLACE(local,")"," ")
'                       prov  = STRING(IF AVAIL provincia THEN provcod_bco ELSE FILL(" ",1))
'                       cpost = ""
'                       nro   = "".
'                DO loop = 1 TO LENGTH(detdom.nro):
'                    ASSIGN puente = INTEGER(SUBSTRING(detdom.nro,loop,1)) NO-ERROR.
'                    IF NOT ERROR-STATUS:ERROR THEN
'                        ASSIGN nro = nro + SUBSTRING(detdom.nro,loop,1).
'                END.
'                DO loop = 1 TO LENGTH(detdom.piso):
'                    ASSIGN puente = INTEGER(SUBSTRING(detdom.piso,loop,1)) NO-ERROR.
'                    IF NOT ERROR-STATUS:ERROR THEN
'                        ASSIGN piso = piso + SUBSTRING(detdom.piso,loop,1).
'                END.
'                DO loop = 1 TO LENGTH(prefijo):
'                    ASSIGN puente = INTEGER(SUBSTRING(prefijo,loop,1)) NO-ERROR.
'                    IF NOT ERROR-STATUS:ERROR THEN
'                        ASSIGN tel1 = tel1 + SUBSTRING(prefijo,loop,1).
'                END.
'                DO loop = 1 TO LENGTH(numtel):
'                    ASSIGN puente = INTEGER(SUBSTRING(numtel,loop,1)) NO-ERROR.
'                    IF NOT ERROR-STATUS:ERROR THEN
'                        ASSIGN tel2 = tel2 + SUBSTRING(numtel,loop,1).
'               END.
'                DO loop = 1 TO LENGTH(detdom.codigopostal):
'                    ASSIGN puente = INTEGER(SUBSTRING(detdom.codigopostal,loop,1)) NO-ERROR.
'                    IF NOT ERROR-STATUS:ERROR THEN
'                        ASSIGN cpost = cpost + SUBSTRING(detdom.codigopostal,loop,1).
'                END.
'            END.
'        END.
'        Else
'            ASSIGN calle = FILL("0", 60)
'                   nro = FILL("0", 5)
'                   piso = FILL("0", 2)
'                   depto = FILL("0", 2)
'                   tel1 = FILL("0", 5)
'                   tel2 = FILL("0", 8)
'                   local = FILL("0",70)
'                   prov = "B"
'                   cpost = "1005    ".

'        /* Ultimos 6 ingresos mensuales */
'        FIND FIRST acu_mes OF empleado WHERE acu_mes.acunro = 6 AND
'                                             acu_mes.amanio = 2004 NO-LOCK NO-ERROR.
'        IF AVAIL acu_mes THEN
'            ASSIGN monto1 = acu_mes.ammonto[11]
'                   monto2 = acu_mes.ammonto[11]
'                   monto3 = acu_mes.ammonto[11]
'                   monto4 = acu_mes.ammonto[11]
'                   monto5 = acu_mes.ammonto[11]
'                   monto6 = acu_mes.ammonto[11].
'
'
'        /* Ultimos Tickets */
'        ASSIGN ticket = 0#
'        FOR EACH proceso WHERE proceso.pliqnro = 75 NO-LOCK,
'            EACH cabliq OF proceso WHERE cabliq.empleado = empleado.ternro NO-LOCK,
'            FIRST acu_liq OF cabliq WHERE acu_liq.acunro = 124 NO-LOCK:
'            ASSIGN ticket = acu_liq.almonto.
'        END.
'
'        ASSIGN apellido = Replace(Empleado.terape, "•", "N")
'               apellido = Replace(apellido, "§", "n")
'               apellido = Replace(apellido, "†", "a")
'               apellido = Replace(apellido, "Ç", "e")
'               apellido = Replace(apellido, "°", "i")
'               apellido = Replace(apellido, "¢", "o")
'               apellido = Replace(apellido, "£", "u")
'               apellido = Replace(apellido, "µ", "A")
'               apellido = Replace(apellido, "ê", "E")
'               apellido = Replace(apellido, "÷", "I")
'               apellido = Replace(apellido, "‡", "O")
'               apellido = Replace(apellido, "È", "u")
'               nombre = Replace(Empleado.ternom, "•", "N")
'               nombre = Replace(nombre, "§", "n")
'               nombre = Replace(nombre, "†", "a")
'               nombre = Replace(nombre, "Ç", "e")
'               nombre = Replace(nombre, "°", "i")
'               nombre = Replace(nombre, "¢", "o")
'               nombre = Replace(nombre, "£", "u")
'               nombre = Replace(nombre, "µ", "A")
'               nombre = Replace(nombre, "ê", "E")
'               nombre = Replace(nombre, "÷", "I")
'               nombre = Replace(nombre, "‡", "O")
'               nombre = REPLACE(nombre,"È","u").
'
'        ASSIGN detalle = sigla +
'                         nrodocu +
'                         cuil +
'                         sexo +
'                         STRING(empleado.empleg,"999999") +
'                         STRING(nombre,"x(40)") +
'                         STRING(apellido,"x(40)") +
'                         STRING(YEAR(tercero.terfecnac),"9999") +
'                         STRING(MONTH(tercero.terfecnac),"99") +
'                         STRING(DAY(tercero.terfecnac),"99") +
'                         STRING(IF AVAIL pais THEN SUBSTRING(pais.paiscodext,1,3) ELSE "   ") +
'                         STRING(IF AVAIL estcivil THEN extciv_bco ELSE "0") +
'                         STRING(calle,"x(60)") +
'                         STRING(IF nro = "     " THEN "00000" ELSE STRING(nro,"99999")) +
'                         STRING(IF piso = "  " THEN "00" ELSE STRING(piso,"99")) +
'                        STRING(IF depto = "  " THEN "00" ELSE depto) +
'                         STRING(IF tel1 = "     " THEN "00000" ELSE STRING(tel1,"99999")) +
'                         STRING(IF tel2 = "        " THEN "00000000" ELSE STRING(tel2,"99999999")) +
'                         local +
'                         prov +
'                         STRING(IF cpost = "        " THEN "1005    " ELSE STRING(cpost,"x(8)")) +
'                         "001" +
'                         STRING(ROUND(monto1,0),"999999") +
'                         STRING(ROUND(monto2,0),"999999") +
'                         STRING(ROUND(monto3,0),"999999") +
'                         STRING(ROUND(monto4,0),"999999") +
'                         STRING(ROUND(monto5,0),"999999") +
'                         STRING(ROUND(monto6,0),"999999") +
'                         STRING(ROUND(bono,0),"999999") +
'                         STRING(ROUND(ticket,0),"999999") +
'                         STRING(YEAR(empleado.empfaltagr),"9999") +
'                         STRING(MONTH(empleado.empfaltagr),"99") +
'                         STRING(DAY(empleado.empfaltagr),"99") +
'                         "N".

'       ASSIGN detalle-excel = sigla + "," +
'                         nrodocu + "," +
'                         cuil +  "," +
'                         sexo + "," +
'                         STRING(empleado.empleg,"999999") + "," +
'                         STRING(nombre,"x(40)") + "," +
'                         STRING(apellido,"x(40)") + "," +
'                         STRING(YEAR(tercero.terfecnac),"9999") + "," +
'                         STRING(MONTH(tercero.terfecnac),"99") + "," +
'                         STRING(DAY(tercero.terfecnac),"99") + "," +
'                         STRING(IF AVAIL pais THEN SUBSTRING(pais.paiscodext,1,3) ELSE "   ") + "," +
'                         STRING(IF AVAIL estcivil THEN extciv_bco ELSE "0") +  "," +
'                         STRING(calle,"x(60)") +  "," +
'                         STRING(IF nro = "     " THEN "00000" ELSE STRING(nro,"99999")) +  "," +
'                         STRING(IF piso = "  " THEN "00" ELSE STRING(piso,"99")) +  "," +
'                         STRING(IF depto = "  " THEN "00" ELSE depto) +  "," +
'                         STRING(IF tel1 = "     " THEN "00000" ELSE STRING(tel1,"99999")) +  "," +
'                         STRING(IF tel2 = "        " THEN "00000000" ELSE STRING(tel2,"99999999")) +  "," +
'                         local + "," +
'                         prov + "," +
'''''''''''''''''''       STRING(IF cpost = "        " THEN "1005    " ELSE STRING(cpost,"x(8)")) +  "," +
'                         "001" +  "," +
'                         STRING(ROUND(monto1,0),"999999") +  "," +
'                         STRING(ROUND(monto2,0),"999999") +  "," +
'                         STRING(ROUND(monto3,0),"999999") +  "," +
'                         STRING(ROUND(monto4,0),"999999") +  "," +
'                         STRING(ROUND(monto5,0),"999999") +  "," +
'                         STRING(ROUND(monto6,0),"999999") + "," +
'                         STRING(ROUND(bono,0),"999999") +  "," +
'                         STRING(ROUND(ticket,0),"999999") + "," +
'                         STRING(YEAR(empleado.empfaltagr),"9999") + "," +
'                         STRING(MONTH(empleado.empfaltagr),"99") + "," +
'                         STRING(DAY(empleado.empfaltagr),"99") + "," +
'                         "N".
                         
'        /* Busco los datos del conyuge */
'        FIND FIRST familiar WHERE familiar.empleado = empleado.ternro AND
'                                  familiar.parenro = 3 NO-LOCK NO-ERROR.
'        IF AVAIL familiar THEN DO:
'            FIND FIRST fam-tercero OF familiar NO-LOCK NO-ERROR.
'            IF AVAIL fam-tercero THEN DO:
'                /* Busco el documento del familiar (DNI,LE o LC) */
'                FIND FIRST ter_doc WHERE ter_doc.ternro = familiar.ternro AND
'                                         ter_doc.tidnro <= 3 NO-LOCK NO-ERROR.
'                IF AVAIL ter_doc THEN DO:
'                    FIND FIRST tipodocu OF ter_doc NO-LOCK NO-ERROR.
'                    IF AVAIL tipodocu THEN
'                        ASSIGN sigla1 = STRING(tipodocu.tidsigla,"x(3)").
'                    Else
'                        ASSIGN sigla1 = "DNI".
'                    ASSIGN nrodocu1 = STRING(REPLACE(ter_doc.nrodoc,".",""),"99999999").
'                END.
'                Else
'                    ASSIGN sigla1 = "DNI"
'                           nrodocu1 = "00000000".
                           
'                /* Pais del conyuge */
'                FIND FIRST pais OF fam-tercero NO-LOCK NO-ERROR.
'
'                ASSIGN apellido = Replace(fam - Tercero.terape, "•", "N")
'                       apellido = Replace(apellido, "§", "n")
'                       apellido = Replace(apellido, "†", "a")
'                       apellido = Replace(apellido, "Ç", "e")
'                       apellido = Replace(apellido, "°", "i")
'                       apellido = Replace(apellido, "¢", "o")
'                       apellido = Replace(apellido, "£", "u")
'                       apellido = Replace(apellido, "µ", "A")
'                       apellido = Replace(apellido, "ê", "E")
'                       apellido = Replace(apellido, "÷", "I")
'                       apellido = Replace(apellido, "‡", "O")
'                       apellido = Replace(apellido, "È", "u")
'                       nombre = Replace(fam - Tercero.ternom, "•", "N")
'                       nombre = Replace(nombre, "§", "n")
'                       nombre = Replace(nombre, "§", "n")
'                       nombre = Replace(nombre, "†", "a")
'                       nombre = Replace(nombre, "Ç", "e")
'                       nombre = Replace(nombre, "°", "i")
'                       nombre = Replace(nombre, "¢", "o")
'                       nombre = Replace(nombre, "£", "u")
'                       nombre = Replace(nombre, "µ", "A")
'                       nombre = Replace(nombre, "ê", "E")
'                       nombre = Replace(nombre, "÷", "I")
'                       nombre = Replace(nombre, "‡", "O")
'                       nombre = REPLACE(nombre,"È","u").

'                ASSIGN detalle = detalle +
'                                 sigla1 +
'                                 STRING(nrodocu1,"99999999") +
'                                 STRING(IF fam-tercero.terfecnac = ? THEN
'                                            "19500101"
'                                         Else
'                                            STRING(YEAR(fam-tercero.terfecnac),"9999") +
'                                            STRING(MONTH(fam-tercero.terfecnac),"99") +
'                                            STRING(DAY(fam-tercero.terfecnac),"99"),"99999999") +
'                                 STRING(IF AVAIL pais THEN SUBSTRING(pais.paiscodext,1,3) ELSE "ARG") +
'                                 STRING(nombre,"x(40)") +
'                                 STRING(apellido,"x(40)").
'                ASSIGN detalle-excel = detalle-excel + "," +
'                                 sigla1 +  "," +
'                                 STRING(nrodocu1,"99999999") + "," +
'                                 IF fam-tercero.terfecnac = ? THEN
'                                    "1950,01,01,"
'                                 Else
'                                    STRING(STRING(YEAR(fam-tercero.terfecnac),"9999") + "," +
'                                           STRING(MONTH(fam-tercero.terfecnac),"99") + "," +
'                                           STRING(DAY(fam-tercero.terfecnac),"99") + ",") +
'                                 STRING(IF AVAIL pais THEN SUBSTRING(pais.paiscodext,1,3) ELSE "ARG") + "," +
'                                 STRING(nombre,"x(40)") + "," +
'                                 STRING(apellido,"x(40)").
'            END.
'            Else
'                ASSIGN detalle = detalle +
'                                 "DNI0000000019500101ARG" +
'                                 FILL("0",80)
'                       detalle-excel = detalle-excel + "," +
'                                 "DNI,00000000,1950,01,01,ARG," +
'                                 FILL("0",40) + "," +
'                                 FILL("0",40).
'        END.
'        Else
'            ASSIGN detalle = detalle +
'                             "DNI0000000019500101ARG" +
'                             FILL("0",80)
'                    detalle-excel = detalle-excel + "," +
'                              "DNI,00000000,1950,01,01,ARG," +
'                              FILL("0",40) + "," +
'                              FILL("0",40).
        
'        PUT STREAM salida UNFORMATTED detalle SKIP.
'        PUT STREAM salida-excel UNFORMATTED detalle-excel SKIP.

'    END.
    
'    /* Imprimo en el encabezado la cantidad de registros */
'    ASSIGN cabecera = STRING(cant-reg,"999999").
'    SEEK STREAM salida TO puntero.
'    PUT STREAM salida UNFORMATTED cabecera.
    
'    OUTPUT STREAM salida CLOSE.
'    SESSION:SET-WAIT-STATE("").
'END.
    
    
    
 
Exit Sub
    
Fin:

ME_Local:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Description
    Flog.writeline Espacios(Tabulador * 1) & "Ultimo SQL: " & StrSql
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.writeline
    HuboErrores = True
End Sub

Sub DejarSoloNumeros(ByRef Texto As String)
Dim i
Dim aux As String
Dim txt_numero As String
Dim caracter As Integer

    If Not EsNulo(Texto) Then
        aux = Texto
        txt_numero = ""
        For i = 1 To Len(aux)
            caracter = Asc(Mid(aux, i, 1))
            If caracter >= 48 And caracter <= 57 Then
                txt_numero = txt_numero & Mid(aux, i, 1)
            End If
        Next
        Texto = txt_numero
    End If

End Sub
