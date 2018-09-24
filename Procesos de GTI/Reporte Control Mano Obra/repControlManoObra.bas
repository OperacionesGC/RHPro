Attribute VB_Name = "repControlManoObra"
Option Explicit

'Global Const Version = "1.00"
'Global Const FechaModificacion = "28/08/2006"
'Global Const UltimaModificacion = "Inicial"

Global Const Version = "1.01"
Global Const FechaModificacion = "19/09/2006"
Global Const UltimaModificacion = "" ' Se elimino la inicializacion de los dias habiles luego de actualizar la BD

Dim fs, f

Dim NroProceso As Long

Global Path As String
Global NArchivo As String
Global HuboErrores As Boolean
Global EmpErrores As Boolean

Global Tabulador As Long
Global TiempoInicialProceso
Global TiempoAcumulado

Global IdUser As String
Global Fecha As Date
Global Hora As String

Private Sub Main()

Dim NombreArchivo As String
Dim Directorio As String
Dim CArchivos
Dim archivo
Dim Folder
Dim strCmdLine As String
Dim Nombre_Arch As String

Dim StrSql As String
Dim objRs As New ADODB.Recordset
Dim cantRegistros
Dim PID As String
Dim Parametros As String
Dim ArrParametros

Dim empl_desde As Long
Dim empl_hasta As Long
Dim empl_estado As Integer
Dim empresa As Long
Dim tenro1 As Long
Dim estrnro1 As Long
Dim tenro2 As Long
Dim estrnro2 As Long
Dim tenro3 As Long
Dim estrnro3 As Long
Dim fecdesde As Date
Dim fechasta As Date

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

    TiempoInicialProceso = GetTickCount
    OpenConnection strconexion, objConn
    OpenConnection strconexion, objConnProgreso
    
    HuboErrores = False
    
    Nombre_Arch = PathFLog & "ReporteControlManoObra" & "-" & NroProceso & ".log"
    
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
    
    Flog.writeline "Inicio Proceso Control de Mano de Obra: " & Now
    Flog.writeline "Cambio el estado del proceso a Procesando"
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & ", bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objConnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Obtengo los datos del proceso"
    
    TiempoAcumulado = GetTickCount
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, objRs
    
    If Not objRs.EOF Then
       IdUser = objRs!IdUser
       Fecha = objRs!bprcfecha
       Hora = objRs!bprchora
       
       'Obtengo los parametros del proceso
       Parametros = objRs!bprcparam
       ArrParametros = Split(Parametros, "@")
       
       'Empleado - Desde Legajo
       empl_desde = ArrParametros(0)
       
       'Empleado - Hasta Legajo
       empl_hasta = ArrParametros(1)
       
       'Empleado - Estado
       empl_estado = ArrParametros(2)
       
       'Empresa
       empresa = ArrParametros(3)
       
       'Primer nivel organizacional
       tenro1 = ArrParametros(4)
       estrnro1 = ArrParametros(5)
       
       'Segundo nivel organizacional
       tenro2 = ArrParametros(6)
       estrnro2 = ArrParametros(7)
       
       'Tercero nivel organizacional
       tenro3 = ArrParametros(8)
       estrnro3 = ArrParametros(9)
       
       'Fecha desde
       fecdesde = ArrParametros(10)
       
       'Fecha hasta
       fechasta = ArrParametros(11)
       
       ' Proceso que genera los datos
       Call GenerarDatos(empl_desde, empl_hasta, empl_estado, empresa, tenro1, estrnro1, tenro2, estrnro2, tenro3, estrnro3, fecdesde, fechasta)
       
    Else
       Exit Sub
    End If
    
    If objRs.State = adStateOpen Then objRs.Close
    
    'Actualizo el estado del proceso
    If Not HuboErrores Then
       StrSql = "UPDATE batch_proceso SET  bprcprogreso = 100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
       Flog.writeline "Proceso Finalizado Correctamente"
    Else
       StrSql = "UPDATE batch_proceso SET  bprcprogreso = 100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProceso
       Flog.writeline "Proceso Incompleto"
    End If
    
    objConnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Fin :" & Now
    Flog.Close

    Exit Sub
    
CE:
    HuboErrores = True
    Flog.writeline " Error: " & Err.Description & Now
End Sub
'--------------------------------------------------------------------
' Se encarga de generar los datos
'--------------------------------------------------------------------
Sub GenerarDatos(ByVal desde As Long, ByVal hasta As Long, ByVal estado As Integer, ByVal empresa As Long, ByVal tenro1 As Long, ByVal estrnro1 As Long, ByVal tenro2 As Long, ByVal estrnro2 As Long, ByVal tenro3 As Long, ByVal estrnro3 As Long, ByVal fecdesde As Date, ByVal fechasta As Date)

Dim StrSql As String
Dim rsConsult As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim objFeriado As New Feriado

'Variables donde se guardan los datos del INSERT final
Dim EmpCuit As String
Dim EmpNombre As String
Dim EmpDire As String
Dim EmpTernro As String
Dim EmpLogo As String
Dim EmpLogoAlto
Dim EmpLogoAncho

Dim CoefResta As Double
Dim CoefHsNormal As Double
Dim CoefHsPeAct As Double
Dim lista_HsExtras As String
Dim lista_ResPuesto As String
Dim lista_PagoHoras As String
Dim lista_NoPagoHoras_S As String
Dim lista_NoPagoHoras_R As String

Dim fecCalculo As Date

Dim tedabr1 As String
Dim estrdabr1 As String
Dim tedabr2 As String
Dim estrdabr2 As String
Dim tedabr3 As String
Dim estrdabr3 As String

Dim estrnro1_ant As Long
Dim estrnro2_ant As Long
Dim estrnro3_ant As Long
Dim estrdabr1_ant As String
Dim estrdabr2_ant As String
Dim estrdabr3_ant As String
Dim guardar As Boolean
Dim hsextras As Double
Dim resvpuesto As Double
Dim pgohs As Double
Dim nopgohs As Double

Dim hsNormal
Dim hspreact
Dim tothsaus
Dim pgohsporc
Dim nopgohsporc
Dim hspres
Dim hspresporc
Dim hsextrasporc
Dim toths
Dim tothsporc
                        
Dim cantemp As Integer
Dim dias_Hab As Integer
Dim contarDiasHab As Boolean
Dim esFeriado As Boolean
Dim insertarRegistro As Boolean

Dim Progreso As Double
Dim cantidadProcesada As Integer
Dim IncPorc As Double

Dim fecdesde_empl As Date
Dim fechasta_empl As Date
Dim orden As Integer
Dim tothsestr As Double
Dim total_AR As Double
Dim total_ANR As Double
Dim confetiq_ant As String

Dim Cargar_registro As Boolean
Dim encontre_rango As Boolean
Dim i As Integer

On Error GoTo MError

    MyBeginTrans
    
    '-------------------------------------------------------------------------
    ' Busco los datos de la empresa
    '--------------------------------------------------------------------------
    StrSql = "SELECT empresa.empnom,empresa.ternro, detdom.calle,nro,codigopostal, localidad.locdesc "
    StrSql = StrSql & " FROM empresa "
    StrSql = StrSql & " LEFT JOIN cabdom ON empresa.ternro = cabdom.ternro "
    StrSql = StrSql & " LEFT JOIN detdom ON detdom.domnro = cabdom.domnro "
    StrSql = StrSql & " LEFT JOIN localidad ON detdom.locnro = localidad.locnro "
    StrSql = StrSql & " WHERE empresa.estrnro = " & empresa
    OpenRecordset StrSql, rsConsult
    If rsConsult.EOF Then
        Flog.writeline Espacios(Tabulador * 1) & "***************************************************************"
        Flog.writeline Espacios(Tabulador * 1) & "Error. No se encontro la empresa."
        Flog.writeline Espacios(Tabulador * 1) & "***************************************************************"
        GoTo Fin_error
    Else
        EmpNombre = rsConsult!empnom
        EmpDire = rsConsult!calle & " " & rsConsult!nro & "<br>" & rsConsult!codigopostal & " " & rsConsult!locdesc
        EmpTernro = rsConsult!Ternro
    End If
    rsConsult.Close
    
    '-------------------------------------------------------------------------
    'Consulta para obtener el cuit de la empresa
    '-------------------------------------------------------------------------
    StrSql = "SELECT cuit.nrodoc FROM tercero " & _
             " INNER JOIN ter_doc cuit ON (tercero.ternro = cuit.ternro and cuit.tidnro = 6)" & _
             " Where tercero.ternro =" & EmpTernro
    OpenRecordset StrSql, rsConsult
    If rsConsult.EOF Then
        Flog.writeline Espacios(Tabulador * 1) & "No se encontró el CUIT de la Empresa."
        EmpCuit = "&nbsp;"
    Else
        EmpCuit = rsConsult!nrodoc
    End If
    rsConsult.Close
    
    '-------------------------------------------------------------------------
    'Consulta para buscar el logo de la empresa
    '-------------------------------------------------------------------------
    StrSql = "SELECT ter_imag.terimnombre, tipoimag.tipimdire, tipoimag.tipimanchodef, tipoimag.tipimaltodef" & _
        " FROM ter_imag " & _
        " INNER JOIN tipoimag ON tipoimag.tipimnro = 1 AND tipoimag.tipimnro = ter_imag.tipimnro" & _
        " AND ter_imag.ternro =" & EmpTernro
    OpenRecordset StrSql, rsConsult
    If rsConsult.EOF Then
        Flog.writeline Espacios(Tabulador * 1) & "No se encontró el Logo de la Empresa."
        EmpLogo = ""
        EmpLogoAlto = 0
        EmpLogoAncho = 0
    Else
        EmpLogo = rsConsult!tipimdire & rsConsult!terimnombre
        EmpLogoAlto = rsConsult!tipimaltodef
        EmpLogoAncho = rsConsult!tipimanchodef
    End If
    rsConsult.Close
    
    
    '-------------------------------------------------------------------------
    'Busco la descripcion del Primer Nivel Organizacional
    '-------------------------------------------------------------------------
    If tenro1 <> 0 Then
        StrSql = "SELECT tenro,tedabr FROM tipoestructura "
        StrSql = StrSql & " WHERE tipoestructura.tenro = " & tenro1
        OpenRecordset StrSql, rsConsult
        If Not rsConsult.EOF Then
            tedabr1 = rsConsult!tedabr
        End If
        rsConsult.Close
        
        If estrnro1 <> -1 Then
            StrSql = "SELECT estrnro,estrdabr FROM estructura "
            StrSql = StrSql & " WHERE estrnro = " & estrnro1
            OpenRecordset StrSql, rsConsult
            If Not rsConsult.EOF Then
                estrdabr1 = rsConsult!estrdabr
            End If
            rsConsult.Close
        End If
    End If
    
    '-------------------------------------------------------------------------
    'Busco la descripcion del Segundo Nivel Organizacional
    '-------------------------------------------------------------------------
    If tenro2 <> 0 Then
        StrSql = "SELECT tenro,tedabr FROM tipoestructura "
        StrSql = StrSql & " WHERE tipoestructura.tenro = " & tenro2
        OpenRecordset StrSql, rsConsult
        If Not rsConsult.EOF Then
            tedabr2 = rsConsult!tedabr
        End If
        rsConsult.Close
        
        If estrnro2 <> -1 Then
            StrSql = "SELECT estrnro,estrdabr FROM estructura "
            StrSql = StrSql & " WHERE estrnro = " & estrnro2
            OpenRecordset StrSql, rsConsult
            If Not rsConsult.EOF Then
                estrdabr2 = rsConsult!estrdabr
            End If
            rsConsult.Close
        End If
    End If
    
    '-------------------------------------------------------------------------
    'Busco la descripcion del tercer Nivel Organizacional
    '-------------------------------------------------------------------------
    If tenro3 <> 0 Then
        StrSql = "SELECT tenro,tedabr FROM tipoestructura "
        StrSql = StrSql & " WHERE tipoestructura.tenro = " & tenro3
        OpenRecordset StrSql, rsConsult
        If Not rsConsult.EOF Then
            tedabr3 = rsConsult!tedabr
        End If
        rsConsult.Close
        
        If estrnro3 <> -1 Then
            StrSql = "SELECT estrnro,estrdabr FROM estructura "
            StrSql = StrSql & " WHERE estrnro = " & estrnro3
            OpenRecordset StrSql, rsConsult
            If Not rsConsult.EOF Then
                estrdabr3 = rsConsult!estrdabr
            End If
            rsConsult.Close
        End If
    End If
    
    '-------------------------------------------------------------------------
    'Inserto la cabecera del reporte
    '-------------------------------------------------------------------------
    StrSql = "INSERT INTO rep_mano_obra  (bpronro,empldesde,emplhasta,emplest,empnombre,empdire,empcuit,emplogo,"
    StrSql = StrSql & "emplogoalto,emplogoancho,tenro1,tedabr1,estrnro1,estrdabr1,tenro2,tedabr2,estrnro2,estrdabr2,"
    StrSql = StrSql & "tenro3,tedabr3,estrnro3,estrdabr3,fecdesde,fechasta,fecha,hora,IdUser) VALUES ("
    StrSql = StrSql & NroProceso & ","
    StrSql = StrSql & desde & ","
    StrSql = StrSql & hasta & ","
    StrSql = StrSql & estado & ","
    StrSql = StrSql & "'" & EmpNombre & "',"
    StrSql = StrSql & "'" & EmpDire & "',"
    StrSql = StrSql & "'" & EmpCuit & "',"
    StrSql = StrSql & "'" & EmpLogo & "',"
    StrSql = StrSql & EmpLogoAlto & ","
    StrSql = StrSql & EmpLogoAncho & ","
    StrSql = StrSql & tenro1 & ","
    StrSql = StrSql & "'" & tedabr1 & "',"
    StrSql = StrSql & estrnro1 & ","
    StrSql = StrSql & "'" & estrdabr1 & "',"
    StrSql = StrSql & tenro2 & ","
    StrSql = StrSql & "'" & tedabr2 & "',"
    StrSql = StrSql & estrnro2 & ","
    StrSql = StrSql & "'" & estrdabr2 & "',"
    StrSql = StrSql & tenro3 & ","
    StrSql = StrSql & "'" & tedabr3 & "',"
    StrSql = StrSql & estrnro3 & ","
    StrSql = StrSql & "'" & estrdabr3 & "',"
    StrSql = StrSql & ConvFecha(fecdesde) & ","
    StrSql = StrSql & ConvFecha(fechasta) & ","
    StrSql = StrSql & ConvFecha(Fecha) & ","
    StrSql = StrSql & "'" & Hora & "',"
    StrSql = StrSql & "'" & IdUser & "')"
    
    objConn.Execute StrSql, , adExecuteNoRecords
    
    '-------------------------------------------------------------------------
    'Busco la configuración del reporte
    '-------------------------------------------------------------------------
    StrSql = "SELECT * FROM confrep WHERE repnro = 171"
    OpenRecordset StrSql, rsConsult
    
    Flog.writeline " "
    Flog.writeline "***************************************************************"
    Flog.writeline "Buscando valores en la configuración del reporte (confrep). Los tipos válidos son:"
    Flog.writeline "        CRE - Coeficiente Resta (V.AlfaNum)"
    Flog.writeline "        CHN - Coeficiente Horas Normal (V.AlfaNum)"
    Flog.writeline "        CHP - Coeficiente Horas Pe.Act (V.AlfaNum)"
    Flog.writeline "        TH - Tipos de Horas"
    Flog.writeline "             Columna 5  - Horas Reserva Puesto"
    Flog.writeline "             Columna 7  - Horas Pago"
    Flog.writeline "             columna 8  - Horas No Pago - Accion SUMA"
    Flog.writeline "             Columna 8  - Horas No Pago - Accion RESTA. Se le resta el CRE"
    Flog.writeline "             Columna 11 - Horas Extras"
    
    CoefResta = 0
    CoefHsNormal = 1
    CoefHsPeAct = 1
    lista_HsExtras = "0"
    lista_ResPuesto = "0"
    lista_PagoHoras = "0"
    lista_NoPagoHoras_S = "0"
    lista_NoPagoHoras_R = "0"
    
    Do Until rsConsult.EOF
        Select Case rsConsult!conftipo
            Case "CRE":
                CoefResta = Replace(rsConsult!confval2, ",", ".")
            Case "CHN":
                CoefHsNormal = Replace(rsConsult!confval2, ",", ".")
            Case "CHP":
                CoefHsPeAct = Replace(rsConsult!confval2, ",", ".")
            Case "TH":
                Select Case rsConsult!confnrocol
                    Case 11:
                        lista_HsExtras = lista_HsExtras & "," & rsConsult!confval
                    Case 5:
                        lista_ResPuesto = lista_ResPuesto & "," & rsConsult!confval
                    Case 7:
                        lista_PagoHoras = lista_PagoHoras & "," & rsConsult!confval
                    Case 8:
                        If rsConsult!confaccion = "sumar" Then
                            lista_NoPagoHoras_S = lista_NoPagoHoras_S & "," & rsConsult!confval
                        Else
                            lista_NoPagoHoras_R = lista_NoPagoHoras_R & "," & rsConsult!confval
                        End If
                    Case Else:
                        Flog.writeline Espacios(Tabulador * 1) & "***************************************************************"
                        Flog.writeline Espacios(Tabulador * 1) & "Error. Tipo TH. El nro columna '" & rsConsult!confnrocol & "' no es valido."
                        Flog.writeline Espacios(Tabulador * 1) & "***************************************************************"
                        GoTo Fin_error
                End Select
            Case Else:
                Flog.writeline Espacios(Tabulador * 1) & "***************************************************************"
                Flog.writeline Espacios(Tabulador * 1) & "Error. Tipo '" & rsConsult!conftipo & "' no reconocido en la configuración."
                Flog.writeline Espacios(Tabulador * 1) & "***************************************************************"
                GoTo Fin_error
        End Select
        rsConsult.MoveNext
    Loop
    rsConsult.Close
    
    
    '-------------------------------------------------------------------------
    'Comiensa a procesar
    '-------------------------------------------------------------------------
    Progreso = 0
    cantidadProcesada = DateDiff("d", fecdesde, fechasta) + 1
    If cantidadProcesada = 0 Then
        cantidadProcesada = 1
    End If
    IncPorc = (99 / cantidadProcesada)
    
    fecCalculo = fecdesde
    Do While fecCalculo <= fechasta
        
        Flog.writeline "** Día --> " & fecCalculo
        
        ' Busco los empleados que respetan el filtro inicial y posean datos en gti_acumdiario a la fecCalculo
        ' ordenados por los niveles organizacionales
        StrSql = "SELECT DISTINCT empleado.ternro "
        If tenro1 <> 0 Then
            StrSql = StrSql & ",his1.estrnro estrnro1,est1.estrdabr estrdabr1"
        End If
        If tenro2 <> 0 Then
            StrSql = StrSql & ",his2.estrnro estrnro2,est2.estrdabr estrdabr2"
        End If
        If tenro3 <> 0 Then
            StrSql = StrSql & ",his3.estrnro estrnro3,est3.estrdabr estrdabr3"
        End If
        StrSql = StrSql & " FROM empleado "
        StrSql = StrSql & " INNER JOIN his_estructura emp ON empleado.ternro = emp.ternro AND emp.tenro = 10"
        StrSql = StrSql & " AND (emp.htetdesde<=" & ConvFecha(fecCalculo) & " AND (emp.htethasta IS NULL OR emp.htethasta>=" & ConvFecha(fecCalculo) & "))"
        StrSql = StrSql & " AND emp.estrnro = " & empresa
        StrSql = StrSql & " INNER JOIN gti_acumdiario ON empleado.ternro = gti_acumdiario.ternro AND gti_acumdiario.adfecha = " & ConvFecha(fecCalculo)
        If tenro1 <> 0 Then
            StrSql = StrSql & " INNER JOIN his_estructura his1 ON empleado.ternro = his1.ternro AND his1.tenro = " & tenro1
            StrSql = StrSql & " AND (his1.htetdesde<=" & ConvFecha(fecCalculo) & " AND (his1.htethasta IS NULL OR his1.htethasta>=" & ConvFecha(fecCalculo) & "))"
            If estrnro1 <> -1 Then
                StrSql = StrSql & " AND his1.estrnro = " & estrnro1
            End If
            StrSql = StrSql & " INNER JOIN estructura est1 ON his1.estrnro = est1.estrnro "
        End If
        If tenro2 <> 0 Then
            StrSql = StrSql & " INNER JOIN his_estructura his2 ON empleado.ternro = his2.ternro AND his2.tenro = " & tenro2
            StrSql = StrSql & " AND (his2.htetdesde<=" & ConvFecha(fecCalculo) & " AND (his2.htethasta IS NULL OR his2.htethasta>=" & ConvFecha(fecCalculo) & "))"
            If estrnro2 <> -1 Then
                StrSql = StrSql & " AND his2.estrnro = " & estrnro2
            End If
            StrSql = StrSql & " INNER JOIN estructura est2 ON his2.estrnro = est2.estrnro "
        End If
        If tenro3 <> 0 Then
            StrSql = StrSql & " INNER JOIN his_estructura his3 ON empleado.ternro = his3.ternro AND his3.tenro = " & tenro3
            StrSql = StrSql & " AND (his3.htetdesde<=" & ConvFecha(fecCalculo) & " AND (his3.htethasta IS NULL OR his3.htethasta>=" & ConvFecha(fecCalculo) & "))"
            If estrnro3 <> -1 Then
                StrSql = StrSql & " AND his3.estrnro = " & estrnro3
            End If
            StrSql = StrSql & " INNER JOIN estructura est3 ON his3.estrnro = est3.estrnro "
        End If
        StrSql = StrSql & " WHERE empleg >= " & desde & " AND empleg <= " & hasta
        If estado <> 1 Then
            StrSql = StrSql & " AND empest = " & estado
        End If
        StrSql = StrSql & " ORDER BY "
        If tenro1 <> 0 Then
            StrSql = StrSql & "estrnro1, estrdabr1,"
        End If
        If tenro2 <> 0 Then
            StrSql = StrSql & "estrnro2, estrdabr2,"
        End If
        If tenro3 <> 0 Then
            StrSql = StrSql & "estrnro3, estrdabr3,"
        End If
        StrSql = StrSql & "empleado.ternro"
        
        OpenRecordset StrSql, rsConsult
        
        dias_Hab = 0
        hsNormal = 0
        resvpuesto = 0
        hspreact = 0
        pgohs = 0
        nopgohs = 0
        tothsaus = 0
        pgohsporc = 0
        nopgohsporc = 0
        hspres = 0
        hspresporc = 0
        hsextras = 0
        hsextrasporc = 0
        toths = 0
        tothsporc = 0
        cantemp = 0
        
        If rsConsult.EOF Then
        
            Flog.writeline "     No se encontraron Acumulados Diarios."
            Flog.writeline "      SQL --> " & StrSql
        
        Else
            
            contarDiasHab = True
            
            Do Until rsConsult.EOF
            
                'Determino si es un dia habil
                esFeriado = False
                If (Weekday(fecCalculo) = 7 Or Weekday(fecCalculo) = 1 Or objFeriado.Feriado(fecCalculo, rsConsult!Ternro, False)) Then
                    esFeriado = True
                Else
                
                    cantemp = cantemp + 1
                
                    If contarDiasHab Then
                        dias_Hab = dias_Hab + 1
                        contarDiasHab = False
                    End If
                End If
                
                'Horas Extras
                StrSql = "SELECT adcanthoras FROM gti_acumdiario "
                StrSql = StrSql & " WHERE thnro IN (" & lista_HsExtras & ") AND adfecha = " & ConvFecha(fecCalculo)
                StrSql = StrSql & " AND ternro = " & rsConsult!Ternro
                OpenRecordset StrSql, rs
                Do Until rs.EOF
                    hsextras = hsextras + rs!adcanthoras
                    rs.MoveNext
                Loop
                rs.Close
                
                If Not esFeriado Then
                    'Reserva Puesto
                    StrSql = "SELECT adcanthoras FROM gti_acumdiario "
                    StrSql = StrSql & " WHERE thnro IN (" & lista_ResPuesto & ") AND adfecha = " & ConvFecha(fecCalculo)
                    StrSql = StrSql & " AND ternro = " & rsConsult!Ternro
                    OpenRecordset StrSql, rs
                    Do Until rs.EOF
                        resvpuesto = resvpuesto + (rs!adcanthoras - CoefResta)
                        rs.MoveNext
                    Loop
                    rs.Close
                
                    'Pago de Horas
                    StrSql = "SELECT adcanthoras FROM gti_acumdiario "
                    StrSql = StrSql & " WHERE thnro IN (" & lista_PagoHoras & ") AND adfecha = " & ConvFecha(fecCalculo)
                    StrSql = StrSql & " AND ternro = " & rsConsult!Ternro
                    OpenRecordset StrSql, rs
                    Do Until rs.EOF
                        pgohs = pgohs + (rs!adcanthoras - CoefResta)
                        rs.MoveNext
                    Loop
                    rs.Close
                
                    'No Pago de Horas - Sumar
                    StrSql = "SELECT adcanthoras FROM gti_acumdiario "
                    StrSql = StrSql & " WHERE thnro IN (" & lista_NoPagoHoras_S & ") AND adfecha = " & ConvFecha(fecCalculo)
                    StrSql = StrSql & " AND ternro = " & rsConsult!Ternro
                    OpenRecordset StrSql, rs
                    Do Until rs.EOF
                        nopgohs = nopgohs + rs!adcanthoras
                        rs.MoveNext
                    Loop
                    rs.Close
                
                    'No Pago de Horas - Resta
                    StrSql = "SELECT adcanthoras FROM gti_acumdiario "
                    StrSql = StrSql & " WHERE thnro IN (" & lista_NoPagoHoras_R & ") AND adfecha = " & ConvFecha(fecCalculo)
                    StrSql = StrSql & " AND ternro = " & rsConsult!Ternro
                    OpenRecordset StrSql, rs
                    Do Until rs.EOF
                        nopgohs = nopgohs + Abs(rs!adcanthoras - CoefResta)
                        rs.MoveNext
                    Loop
                    rs.Close
                End If
                
                If tenro1 <> 0 Then
                    estrnro1_ant = rsConsult!estrnro1
                    estrdabr1_ant = rsConsult!estrdabr1
                End If
                If tenro2 <> 0 Then
                    estrnro2_ant = rsConsult!estrnro2
                    estrdabr2_ant = rsConsult!estrdabr2
                End If
                If tenro3 <> 0 Then
                    estrnro3_ant = rsConsult!estrnro3
                    estrdabr3_ant = rsConsult!estrdabr3
                End If
                
                rsConsult.MoveNext
                
                guardar = False
                If rsConsult.EOF Then
                    guardar = True
                Else
                    If tenro1 <> 0 Then
                        If estrdabr1_ant <> rsConsult!estrdabr1 Then
                            guardar = True
                        End If
                    End If
                    
                    If tenro2 <> 0 Then
                        If estrdabr2_ant <> rsConsult!estrdabr2 Then
                            guardar = True
                        End If
                    End If
                
                    If tenro3 <> 0 Then
                        If estrdabr3_ant <> rsConsult!estrdabr3 Then
                            guardar = True
                        End If
                    End If
                End If
                
                If guardar Then
                    ' Busco si se encuentran valores para los niveles organizacionales
                    ' Si no encuentro --> inserto
                    ' Si encuentro valores
                    '       Verifico si el ultimo dia (fhasta) es imnediatamente anterior al fecCalculo y la dotacion de personal coincide
                    '           Recalculo los valores
                    '           fhasta = fecCalculo
                    '       Sino
                    '           inserto un nuevo valor con nuevo intervalo de fecha y dotacion de personal
                    
                    StrSql = "SELECT * FROM rep_mano_obra_det WHERE bpronro = " & NroProceso
                    If tenro1 <> 0 Then
                        StrSql = StrSql & " AND estrnro1 = " & estrnro1_ant
                    End If
                    If tenro2 <> 0 Then
                        StrSql = StrSql & " AND estrnro2 = " & estrnro2_ant
                    End If
                    If tenro3 <> 0 Then
                        StrSql = StrSql & " AND estrnro3 = " & estrnro3_ant
                    End If
                    OpenRecordset StrSql, rs
                    
                    If rs.EOF Then
                        'Insertar
                        insertarRegistro = True
                    Else
                        encontre_rango = True
                        Do While (encontre_rango)
                            If rs.EOF Then
                                encontre_rango = False
                            Else
                                If (rs!fhasta = DateAdd("d", -1, fecCalculo) And (rs!dotacion = cantemp Or esFeriado)) Then
                                    encontre_rango = False
                                Else
                                    rs.MoveNext
                                End If
                            End If
                        Loop
                        
                        If rs.EOF Then
                            ' Insertar
                            insertarRegistro = True
                        Else
                            ' Recalcular
                            insertarRegistro = False
                        End If
                    End If
                    
                    
                    If insertarRegistro Then
                        'Insertar registro
                        hsNormal = (cantemp * CoefHsNormal * dias_Hab)
                        hspreact = ((cantemp * CoefHsPeAct * dias_Hab) - resvpuesto)
                        tothsaus = pgohs + nopgohs
                        If tothsaus <> 0 Then
                            pgohsporc = (pgohs * 100) / tothsaus
                            nopgohsporc = (nopgohs * 100) / tothsaus
                        End If
                        hspres = hspreact - tothsaus
                        If hsNormal <> 0 Then
                            hspresporc = (hspres * 100) / hsNormal
                        End If
                        If hspres <> 0 Then
                            hsextrasporc = (hsextras * 100) / hspres
                        End If
                        toths = hspres + hsextras
                        If hspres <> 0 Then
                            tothsporc = (toths * 100) / hspres
                        End If
                        'Inserto
                        StrSql = "INSERT INTO rep_mano_obra_det (bpronro,estrnro1,estrdabr1,estrnro2,estrdabr2,estrnro3," & _
                            "estrdabr3,fdesde,fhasta,diasHab,dotacion,hsnormal,resvpuesto,hspreact,pgohs,pgohsporc,nopgohs," & _
                            "nopgohsporc,tothsaus,hspres,hspresporc,hsextras,hsextrasporc,toths,tothsporc) VALUES (" & _
                            NroProceso & "," & estrnro1_ant & ",'" & estrdabr1_ant & "'," & estrnro2_ant & ",'" & _
                            estrdabr2_ant & "'," & estrnro3_ant & ",'" & estrdabr3_ant & "'," & _
                            ConvFecha(fecCalculo) & "," & ConvFecha(fecCalculo) & "," & dias_Hab & "," & cantemp & "," & _
                            hsNormal & "," & resvpuesto & "," & hspreact & "," & pgohs & "," & pgohsporc & "," & nopgohs & "," & _
                            nopgohsporc & "," & tothsaus & "," & hspres & "," & hspresporc & "," & _
                            hsextras & "," & hsextrasporc & "," & toths & "," & tothsporc & ")"
                        objConn.Execute StrSql, , adExecuteNoRecords
                        
                        Flog.writeline "     Inserto Registro."
                        Flog.writeline "      SQL --> " & StrSql
                        
                    Else
                        'Update registro
                        dias_Hab = dias_Hab + rs!diasHab
                        hsNormal = (rs!dotacion * CoefHsNormal * dias_Hab)
                        resvpuesto = resvpuesto + rs!resvpuesto
                        hspreact = ((rs!dotacion * CoefHsPeAct * dias_Hab) - resvpuesto)
                        pgohs = pgohs + rs!pgohs
                        nopgohs = nopgohs + rs!nopgohs
                        tothsaus = pgohs + nopgohs
                        If tothsaus <> 0 Then
                            pgohsporc = (pgohs * 100) / tothsaus
                            nopgohsporc = (nopgohs * 100) / tothsaus
                        End If
                        hspres = hspreact - tothsaus
                        If hsNormal <> 0 Then
                            hspresporc = (hspres * 100) / hsNormal
                        End If
                        hsextras = hsextras + rs!hsextras
                        If hspres <> 0 Then
                            hsextrasporc = (hsextras * 100) / hspres
                        End If
                        toths = hspres + hsextras
                        If hspres <> 0 Then
                            tothsporc = (toths * 100) / hspres
                        End If
                        'Update
                        StrSql = "UPDATE rep_mano_obra_det SET " & _
                            "hsnormal = " & hsNormal & ",fhasta = " & ConvFecha(fecCalculo) & ",diasHab=" & dias_Hab & _
                            ",resvpuesto = " & resvpuesto & ",hspreact = " & hspreact & ",pgohs = " & pgohs & _
                            ",pgohsporc = " & pgohsporc & ",nopgohs = " & nopgohs & _
                            ",nopgohsporc = " & nopgohsporc & ",tothsaus = " & tothsaus & ",hspres = " & hspres & _
                            ",hspresporc = " & hspresporc & ",hsextras = " & hsextras & _
                            ",hsextrasporc = " & hsextrasporc & ",toths = " & toths & _
                            ",tothsporc = " & tothsporc & _
                            " WHERE bpronro = " & NroProceso & " AND estrnro1 = " & rs!estrnro1 & _
                            " AND estrnro2 = " & rs!estrnro2 & " AND estrnro3 = " & rs!estrnro3 & _
                            " AND fdesde = " & ConvFecha(rs!fdesde) & " AND fhasta = " & ConvFecha(rs!fhasta)
                        objConn.Execute StrSql, , adExecuteNoRecords
                    
                        Flog.writeline "     Update Registro."
                        Flog.writeline "      SQL --> " & StrSql
                        
                    End If
                    rs.Close
                                            
                    'Inicialiso los valores
                    'dias_Hab = 0
                    hsNormal = 0
                    resvpuesto = 0
                    hspreact = 0
                    pgohs = 0
                    nopgohs = 0
                    tothsaus = 0
                    pgohsporc = 0
                    nopgohsporc = 0
                    hspres = 0
                    hspresporc = 0
                    hsextras = 0
                    hsextrasporc = 0
                    toths = 0
                    tothsporc = 0
                    cantemp = 0
                End If
            Loop
                
                
        End If
        
        rsConsult.Close
        
        'Actualizo el progreso
        Progreso = Progreso + IncPorc
        cantidadProcesada = cantidadProcesada - 1
        
        TiempoAcumulado = GetTickCount
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
        StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'"
        StrSql = StrSql & ", bprcempleados ='" & CStr(cantidadProcesada) & "' WHERE bpronro = " & NroProceso
        objConnProgreso.Execute StrSql, , adExecuteNoRecords
        
        fecCalculo = DateAdd("d", 1, fecCalculo)
        
    Loop
    
    Flog.writeline " "
    
    MyCommitTrans
    
Fin:
Exit Sub
            
Fin_error:
    MyRollbackTrans
    Exit Sub

MError:
    MyRollbackTrans
'    Resume Next
    Flog.writeline
    Flog.writeline "***************************************************************"
    Flog.writeline " Error: " & Err.Description
    Flog.writeline " Última Sql ejecutada: " & StrSql
    Flog.writeline "***************************************************************"
    Flog.writeline
    HuboErrores = True
    Exit Sub
End Sub
            
