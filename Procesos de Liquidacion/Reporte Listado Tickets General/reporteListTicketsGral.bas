Attribute VB_Name = "repListTicketsGral"
Option Explicit

Dim fs, f
Global Flog

Dim NroLinea As Long
Dim crpNro As Long
Dim RegLeidos As Long
Dim RegError As Long
Dim RegFecha As Date
Dim NroProceso As Long

Global Path As String
Global NArchivo As String
Global Rta
Global HuboErrores As Boolean
Global EmpErrores As Boolean

Global pliqnro As Integer
Global pliqdesc As String
Global empresa As Integer
Global procaprob As Boolean
Global todospro As Boolean
Global pronro As Integer
Global columna(8) As String
Global Tabulador As Long

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
Dim rsPeriodos As New ADODB.Recordset
Dim tipoDepuracion
Dim historico As Boolean
Dim param
Dim rsConceptos As New ADODB.Recordset
Dim i
Dim totalAcum
Dim cantRegistros
Dim PID As String
Dim parametros As String
Dim ArrParametros
Dim concnro As Integer

Dim pliqdesc As String
Dim prodesc As String
Dim EmpTernro As Long
Dim Empnro As Long
Dim EmpEstrnro As Integer
Dim EmpNombre As String
Dim EmpDire As String
Dim EmpCuit As String
Dim EmpLogo As String
Dim EmpFirma As String
Dim EmpLogoAlto As Integer
Dim EmpLogoAncho As Integer
Dim EmpFirmaAlto As Integer
Dim EmpFirmaAncho As Integer


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
    
    HuboErrores = False
    
    Nombre_Arch = PathFLog & "ReporteListTicketGral" & "-" & NroProceso & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    Flog.Writeline "Inicio Proceso de Listado de tickets General : " & Now
    Flog.Writeline "Cambio el estado del proceso a Procesando"
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.Writeline "PID = " & PID
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & ", bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Flog.Writeline "Obtengo los datos del proceso"
    
    TiempoAcumulado = GetTickCount
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, objRs
    
    If Not objRs.EOF Then
       IdUser = objRs!IdUser
       Fecha = objRs!bprcfecha
       Hora = objRs!bprchora
               
       'Obtengo los parametros del proceso
       parametros = objRs!bprcparam
       ArrParametros = Split(parametros, "@")
       
       'Obtengo la empresa
       empresa = ArrParametros(0)
       
       'Obtengo el Período
       pliqnro = ArrParametros(1)
       
       'Obtengo la opcion de solo procesos aprobados
       procaprob = ArrParametros(2)
       
       'Obtengo la opcion de todos los procesos
       todospro = ArrParametros(3)
       
       'Obtengo el proceso, si se eligio alguno en especial
       pronro = ArrParametros(4)
       
               
       ' Obtengo los valores del confrep
       StrSql = "SELECT * FROM confrep WHERE confrep.repnro= 132 ORDER BY confrep.confnrocol "
       OpenRecordset StrSql, objRs
       
       For i = 1 To 7
           columna(i) = 0
       Next
       
       Do Until objRs.EOF
            If objRs!conftipo = "AC" Then
               columna(objRs!confnrocol) = objRs!confval
            ElseIf objRs!conftipo = "CO" Then
               columna(objRs!confnrocol) = objRs!confval2
            End If
            objRs.MoveNext
       Loop
       If objRs.State = adStateOpen Then objRs.Close

       
       'EMPIEZA EL PROCESO
       
       StrSql = "SELECT * FROM empresa WHERE estrnro = " & empresa
       OpenRecordset StrSql, objRs
    
       EmpEstrnro = 0
       EmpTernro = 0
       If objRs.EOF Then
            Flog.Writeline "No se encontró la empresa"
            Exit Sub
       Else
            Empnro = objRs!Empnro
            EmpNombre = objRs!empnom
            EmpTernro = objRs!ternro
       End If
       If objRs.State = adStateOpen Then objRs.Close
       
       'Consulta para obtener la direccion de la empresa
       StrSql = "SELECT detdom.calle,detdom.nro,localidad.locdesc,codigopostal From cabdom " & _
            " INNER JOIN detdom ON detdom.domnro = cabdom.domnro AND cabdom.ternro =" & EmpTernro & _
            " INNER JOIN localidad ON detdom.locnro = localidad.locnro "
       OpenRecordset StrSql, objRs
       If objRs.EOF Then
            Flog.Writeline "No se encontró el domicilio de la empresa"
            EmpDire = " "
       Else
            EmpDire = objRs!calle & " " & objRs!Nro & " (" & objRs!codigopostal & " ) " & objRs!locdesc
       End If
       If objRs.State = adStateOpen Then objRs.Close
       
       'Consulta para obtener el cuit de la empresa
       StrSql = "SELECT cuit.nrodoc FROM tercero " & _
                 " INNER JOIN ter_doc cuit ON (tercero.ternro = cuit.ternro and cuit.tidnro = 6)" & _
                 " Where tercero.ternro =" & EmpTernro
       OpenRecordset StrSql, objRs
       If objRs.EOF Then
            Flog.Writeline "No se encontró el CUIT de la Empresa"
            EmpCuit = "  "
       Else
            EmpCuit = objRs!nrodoc
       End If
       If objRs.State = adStateOpen Then objRs.Close
       
       'Consulta para buscar el logo de la empresa
       StrSql = "SELECT ter_imag.terimnombre, tipoimag.tipimdire, tipoimag.tipimanchodef, tipoimag.tipimaltodef" & _
            " From ter_imag " & _
            " INNER JOIN tipoimag ON tipoimag.tipimnro = 1 AND tipoimag.tipimnro = ter_imag.tipimnro" & _
            " AND ter_imag.ternro =" & EmpTernro
       OpenRecordset StrSql, objRs
       If objRs.EOF Then
            Flog.Writeline "No se encontró el Logo de la Empresa"
            EmpLogo = ""
            EmpLogoAlto = 0
            EmpLogoAncho = 0
       Else
            EmpLogo = objRs!tipimdire & objRs!terimnombre
            EmpLogoAlto = objRs!tipimaltodef
            EmpLogoAncho = objRs!tipimanchodef
       End If
       If objRs.State = adStateOpen Then objRs.Close
       
       'Busco el periodo desde
       StrSql = "SELECT * FROM periodo WHERE pliqnro = " & pliqnro
       OpenRecordset StrSql, objRs
        
       If Not objRs.EOF Then
          pliqdesc = objRs!pliqdesc
       Else
          Flog.Writeline "No se encontro el Período."
          Exit Sub
       End If
        
       If objRs.State = adStateOpen Then objRs.Close
       
       If pronro <> 0 Then
            'Busco el proceso
            StrSql = "SELECT * FROM proceso WHERE pronro = " & pronro
            OpenRecordset StrSql, objRs
             
            If Not objRs.EOF Then
               prodesc = objRs!prodesc
            Else
               Flog.Writeline "No se encontro el Proceso."
               Exit Sub
            End If
             
            If objRs.State = adStateOpen Then objRs.Close
       End If
       
       ' Inserto el encabezado del reporte
       Flog.Writeline "Inserto la cabecera del reporte."
       
       StrSql = "INSERT INTO rep_list_tick (bpronro,pliqnro,pliqdesc,todospro,procaprob," & _
                "pronro,prodesc,empnombre,empdire,empcuit,emplogo,emplogoalto," & _
                "emplogoancho,fecha,hora,iduser) " & _
                "VALUES (" & _
                 NroProceso & "," & pliqnro & ",'" & pliqdesc & "'," & CInt(todospro) & "," & _
                 CInt(procaprob) & "," & pronro & ",'" & prodesc & "','" & EmpNombre & "','" & _
                 EmpDire & "','" & EmpCuit & "','" & EmpLogo & "'," & EmpLogoAlto & "," & _
                 EmpLogoAncho & "," & ConvFecha(Fecha) & ",'" & Hora & "','" & IdUser & "')"
       objConn.Execute StrSql, , adExecuteNoRecords
       
       ' Proceso que genera los datos
       Call GenerarDatosProceso(Empnro, pliqnro, procaprob, todospro, pronro)
       
    Else
        Exit Sub
    End If
   
    'Actualizo el estado del proceso
    If Not HuboErrores Then
       StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
       Flog.Writeline "Proceso Finalizado Correctamente"
    Else
       StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProceso
       Flog.Writeline "Proceso Incompleto"
    End If
    
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Flog.Writeline "Fin :" & Now
    Flog.Close

    Exit Sub
    
CE:
    HuboErrores = True
    Flog.Writeline " Error: " & Err.Description & Now

End Sub
'--------------------------------------------------------------------
' Se encarga de generar los datos
'--------------------------------------------------------------------
Sub GenerarDatosProceso(ByVal empresa As Integer, ByVal pliqnro As Integer, ByVal procaprob As Boolean, ByVal todospro As Boolean, ByVal pronro As Integer)

Dim StrSql As String
Dim rsConsult As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset
Dim rs4 As New ADODB.Recordset


'Variables donde se guardan los datos del INSERT final
Dim sueldo As Double
Dim suetot As Double
Dim imptic As Double
Dim imptr As Double
Dim imptc As Double
Dim imptp As Double
Dim porctr As Double
Dim porctc As Double
Dim centrocosto As String
Dim categoria As String
Dim ternro As Long
Dim empleg As Long
Dim apellido As String
Dim nombre As String
Dim ccostocodext As String
Dim catcodext As String

Dim Cantidad As Integer
Dim cantidadProcesada As Integer

On Error GoTo MError

'------------------------------------------------------------------
' Busco los empleados
'------------------------------------------------------------------
If todospro Then
   StrSql = " SELECT e1.estrnro, e1.estrdabr, e2.estrnro, e2.estrdabr, ve.empleg, ve.terape, ve.terape2, ve.ternom, ve.ternom2, ve.ternro, e1.estrcodext, e2.estrcodext  "
   StrSql = StrSql & " FROM periodo "
   StrSql = StrSql & " INNER JOIN proceso ON proceso.pliqnro = periodo.pliqnro AND proceso.empnro= " & empresa
   StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
   StrSql = StrSql & " INNER JOIN v_empleado ve ON ve.ternro = cabliq.empleado "
   StrSql = StrSql & " INNER JOIN his_estructura he1 ON he1.ternro = cabliq.empleado AND he1.htethasta IS NULL AND he1.tenro = 5 " 'Centro Costo
   StrSql = StrSql & " INNER JOIN estructura e1 ON he1.estrnro = e1.estrnro "
   StrSql = StrSql & " INNER JOIN his_estructura he2 ON he2.ternro = cabliq.empleado AND he2.htethasta IS NULL AND he2.tenro = 3 " 'Categoria
   StrSql = StrSql & " INNER JOIN estructura e2 ON he2.estrnro = e2.estrnro "
   StrSql = StrSql & " WHERE periodo.pliqnro = " & pliqnro
   StrSql = StrSql & " GROUP BY e1.estrnro, e1.estrdabr, e2.estrnro, e2.estrdabr, ve.empleg, ve.terape, ve.terape2, ve.ternom, ve.ternom2, ve.ternro, e1.estrcodext, e2.estrcodext  "
   StrSql = StrSql & " ORDER BY e1.estrcodext, e2.estrcodext, ve.empleg "
ElseIf procaprob Then
   StrSql = " SELECT e1.estrnro, e1.estrdabr, e2.estrnro, e2.estrdabr, ve.empleg, ve.terape, ve.terape2, ve.ternom, ve.ternom2, ve.ternro, e1.estrcodext, e2.estrcodext  "
   StrSql = StrSql & " FROM periodo "
   StrSql = StrSql & " INNER JOIN proceso ON proceso.pliqnro = periodo.pliqnro AND (proceso.proaprob = 3 OR proceso.proaprob = 2) "
   StrSql = StrSql & " AND proceso.empnro= " & empresa
   StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
   StrSql = StrSql & " INNER JOIN v_empleado ve ON ve.ternro = cabliq.empleado "
   StrSql = StrSql & " INNER JOIN his_estructura he1 ON he1.ternro = cabliq.empleado AND he1.htethasta IS NULL AND he1.tenro = 5 " 'Centro Costo
   StrSql = StrSql & " INNER JOIN estructura e1 ON he1.estrnro = e1.estrnro "
   StrSql = StrSql & " INNER JOIN his_estructura he2 ON he2.ternro = cabliq.empleado AND he2.htethasta IS NULL AND he2.tenro = 3 " 'Categoria
   StrSql = StrSql & " INNER JOIN estructura e2 ON he2.estrnro = e2.estrnro "
   StrSql = StrSql & " WHERE periodo.pliqnro = " & pliqnro
   StrSql = StrSql & " GROUP BY e1.estrnro, e1.estrdabr, e2.estrnro, e2.estrdabr, ve.empleg, ve.terape, ve.terape2, ve.ternom, ve.ternom2, ve.ternro, e1.estrcodext, e2.estrcodext  "
   StrSql = StrSql & " ORDER BY e1.estrcodext, e2.estrcodext, ve.empleg "
Else 'Se seleciono un proceso en particular
   StrSql = " SELECT e1.estrnro, e1.estrdabr, e2.estrnro, e2.estrdabr, ve.empleg, ve.terape, ve.terape2, ve.ternom, ve.ternom2, ve.ternro, e1.estrcodext, e2.estrcodext "
   StrSql = StrSql & " FROM periodo "
   StrSql = StrSql & " INNER JOIN proceso ON proceso.pliqnro = periodo.pliqnro AND proceso.pronro = " & pronro
   StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
   StrSql = StrSql & " INNER JOIN v_empleado ve ON ve.ternro = cabliq.empleado "
   StrSql = StrSql & " INNER JOIN his_estructura he1 ON he1.ternro = cabliq.empleado AND he1.htethasta IS NULL AND he1.tenro = 5 " 'Centro Costo
   StrSql = StrSql & " INNER JOIN estructura e1 ON he1.estrnro = e1.estrnro "
   StrSql = StrSql & " INNER JOIN his_estructura he2 ON he2.ternro = cabliq.empleado AND he2.htethasta IS NULL AND he2.tenro = 3 " 'Categoria
   StrSql = StrSql & " INNER JOIN estructura e2 ON he2.estrnro = e2.estrnro "
   StrSql = StrSql & " WHERE periodo.pliqnro = " & pliqnro
   StrSql = StrSql & " GROUP BY e1.estrnro, e1.estrdabr, e2.estrnro, e2.estrdabr, ve.empleg, ve.terape, ve.terape2, ve.ternom, ve.ternom2, ve.ternro, e1.estrcodext, e2.estrcodext  "
   StrSql = StrSql & " ORDER BY e1.estrcodext, e2.estrcodext, ve.empleg "
End If

OpenRecordset StrSql, rsConsult

'Seteo el progreso
If rsConsult.RecordCount <> 0 Then
    Cantidad = rsConsult.RecordCount
Else
    Cantidad = 1
End If
IncPorc = 99 / Cantidad
cantidadProcesada = Cantidad

Do Until rsConsult.EOF
    
    sueldo = 0
    suetot = 0
    imptic = 0
    imptr = 0
    imptc = 0
    imptp = 0
    porctr = 0
    porctc = 0
    
    centrocosto = rsConsult(1)
    categoria = rsConsult(3)
    ccostocodext = rsConsult(10)
    catcodext = rsConsult(11)
    empleg = rsConsult!empleg
    ternro = rsConsult!ternro
    apellido = rsConsult!terape
    If rsConsult!terape2 <> "" Then
        apellido = apellido & " " & rsConsult!terape2
    End If
    nombre = rsConsult!ternom
    If rsConsult!ternom2 <> "" Then
        nombre = nombre & " " & rsConsult!ternom2
    End If
    
    ' Busco el Sueldo
    StrSql = " SELECT cabliq.empleado, sum(almonto) as almonto "
    StrSql = StrSql & " FROM cabliq "
    StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro AND proceso.pliqnro = " & pliqnro
    If procaprob Then
        StrSql = StrSql & " AND (proceso.proaprob = 3 OR proceso.proaprob = 2) "
    ElseIf Not todospro Then
        StrSql = StrSql & " AND proceso.pronro = " & pronro
    End If
    StrSql = StrSql & " INNER JOIN acu_liq ON acu_liq.acunro = " & columna(1) & " AND acu_liq.cliqnro = cabliq.cliqnro "
    StrSql = StrSql & " WHERE cabliq.empleado = " & rsConsult!ternro
    StrSql = StrSql & " GROUP BY cabliq.empleado "
    OpenRecordset StrSql, rs2
     
    If Not rs2.EOF Then
        sueldo = rs2!almonto
    End If
     
     
    'Para ver si tiene TICKETS
    StrSql = " SELECT concnro "
    StrSql = StrSql & " FROM concepto "
    StrSql = StrSql & " WHERE concepto.conccod = '" & columna(2) & "'"
    
    OpenRecordset StrSql, rs3
     
    If Not rs3.EOF Then
        StrSql = " SELECT cabliq.empleado, sum(dlimonto) as dlimonto "
        StrSql = StrSql & " FROM cabliq "
        StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro AND proceso.pliqnro = " & pliqnro
        If procaprob Then
           StrSql = StrSql & " AND (proceso.proaprob = 3 OR proceso.proaprob = 2) "
        ElseIf Not todospro Then
           StrSql = StrSql & " AND proceso.pronro = " & pronro
        End If
        StrSql = StrSql & " INNER JOIN detliq ON detliq.concnro = " & rs3!concnro & " AND detliq.cliqnro = cabliq.cliqnro "
        StrSql = StrSql & " WHERE cabliq.empleado = " & rsConsult!ternro
        StrSql = StrSql & " GROUP BY cabliq.empleado "
    
        OpenRecordset StrSql, rs4
     
        If Not rs4.EOF Then 'Si tiene ticket
            If Not rs2.EOF Then
                suetot = Round(CDbl((CDbl(rs2!almonto) / 0.9091)), 0)
            End If
        Else 'No tiene tickets
            If Not rs2.EOF Then
                suetot = rs2!almonto
            End If
        End If
        If rs4.State = adStateOpen Then rs4.Close
     End If
     If rs3.State = adStateOpen Then rs3.Close
     
     'Total Tickets Restaurant %
     StrSql = " SELECT concnro "
     StrSql = StrSql & " FROM concepto "
     StrSql = StrSql & " WHERE concepto.conccod = '" & columna(3) & "'"
     OpenRecordset StrSql, rs3
     
     If Not rs3.EOF Then
        StrSql = " SELECT cabliq.empleado, sum(dlimonto) as dlimonto "
        StrSql = StrSql & " FROM cabliq "
        StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro AND proceso.pliqnro = " & pliqnro
        If procaprob Then
           StrSql = StrSql & " AND (proceso.proaprob = 3 OR proceso.proaprob = 2) "
        ElseIf Not todospro Then
           StrSql = StrSql & " AND proceso.pronro = " & pronro
        End If
        StrSql = StrSql & " INNER JOIN detliq ON detliq.concnro = " & rs3!concnro & " AND detliq.cliqnro = cabliq.cliqnro "
        StrSql = StrSql & " WHERE cabliq.empleado = " & rsConsult!ternro
        StrSql = StrSql & " GROUP BY cabliq.empleado "
        OpenRecordset StrSql, rs4
     
        If Not rs4.EOF Then 'Si tiene ticket
           imptr = rs4!dlimonto
        End If
        If rs4.State = adStateOpen Then rs4.Close
     End If
     If rs3.State = adStateOpen Then rs3.Close
     
     'Total Tickets Restaurant suma fija
     StrSql = " SELECT concnro "
     StrSql = StrSql & " FROM concepto "
     StrSql = StrSql & " WHERE concepto.conccod = '" & columna(4) & "'"
     OpenRecordset StrSql, rs3
     
     If Not rs3.EOF Then
        StrSql = " SELECT cabliq.empleado, sum(dlimonto) as dlimonto "
        StrSql = StrSql & " FROM cabliq "
        StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro AND proceso.pliqnro = " & pliqnro
        If procaprob Then
           StrSql = StrSql & " AND (proceso.proaprob = 3 OR proceso.proaprob = 2) "
        ElseIf Not todospro Then
           StrSql = StrSql & " AND proceso.pronro = " & pronro
        End If
        StrSql = StrSql & " INNER JOIN detliq ON detliq.concnro = " & rs3!concnro & " AND detliq.cliqnro = cabliq.cliqnro "
        StrSql = StrSql & " WHERE cabliq.empleado = " & rsConsult!ternro
        StrSql = StrSql & " GROUP BY cabliq.empleado "
        OpenRecordset StrSql, rs4
     
        If Not rs4.EOF Then 'Si tiene ticket
           imptr = CDbl(imptr) + CDbl(rs4!dlimonto)
        End If
        If rs4.State = adStateOpen Then rs4.Close
     End If
     If rs3.State = adStateOpen Then rs3.Close
     
     'Total Tickets Canasta %
     StrSql = " SELECT concnro "
     StrSql = StrSql & " FROM concepto "
     StrSql = StrSql & " WHERE concepto.conccod = '" & columna(5) & "'"
     OpenRecordset StrSql, rs3
     
     If Not rs3.EOF Then
        StrSql = " SELECT cabliq.empleado, sum(dlimonto) as dlimonto "
        StrSql = StrSql & " FROM cabliq "
        StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro AND proceso.pliqnro = " & pliqnro
        If procaprob Then
           StrSql = StrSql & " AND (proceso.proaprob = 3 OR proceso.proaprob = 2) "
        ElseIf Not todospro Then
           StrSql = StrSql & " AND proceso.pronro = " & pronro
        End If
        StrSql = StrSql & " INNER JOIN detliq ON detliq.concnro = " & rs3!concnro & " AND detliq.cliqnro = cabliq.cliqnro "
        StrSql = StrSql & " WHERE cabliq.empleado = " & rsConsult!ternro
        StrSql = StrSql & " GROUP BY cabliq.empleado "
        OpenRecordset StrSql, rs4
     
        If Not rs4.EOF Then 'Si tiene ticket
           imptc = CDbl(rs4!dlimonto)
        End If
        If rs4.State = adStateOpen Then rs4.Close
     End If
     If rs3.State = adStateOpen Then rs3.Close
     
     'Total Tickets Canasta suma fija
     StrSql = " SELECT concnro "
     StrSql = StrSql & " FROM concepto "
     StrSql = StrSql & " WHERE concepto.conccod = '" & columna(6) & "'"
     OpenRecordset StrSql, rs3
     
     If Not rs3.EOF Then
        StrSql = " SELECT cabliq.empleado, sum(dlimonto) as dlimonto "
        StrSql = StrSql & " FROM cabliq "
        StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro AND proceso.pliqnro = " & pliqnro
        If procaprob Then
           StrSql = StrSql & " AND (proceso.proaprob = 3 OR proceso.proaprob = 2) "
        ElseIf Not todospro Then
           StrSql = StrSql & " AND proceso.pronro = " & pronro
        End If
        StrSql = StrSql & " INNER JOIN detliq ON detliq.concnro = " & rs3!concnro & " AND detliq.cliqnro = cabliq.cliqnro "
        StrSql = StrSql & " WHERE cabliq.empleado = " & rsConsult!ternro
        StrSql = StrSql & " GROUP BY cabliq.empleado "
        OpenRecordset StrSql, rs4
     
        If Not rs4.EOF Then 'Si tiene ticket
           imptc = CDbl(imptc) + CDbl(rs4!dlimonto)
        End If
        If rs4.State = adStateOpen Then rs4.Close
     End If
     If rs3.State = adStateOpen Then rs3.Close
     
     'Total Tickets Plus
     StrSql = " SELECT concnro "
     StrSql = StrSql & " FROM concepto "
     StrSql = StrSql & " WHERE concepto.conccod = '" & columna(7) & "'"
     OpenRecordset StrSql, rs3
     
     If Not rs3.EOF Then
        StrSql = " SELECT cabliq.empleado, sum(dlimonto) as dlimonto "
        StrSql = StrSql & " FROM cabliq "
        StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro AND proceso.pliqnro = " & pliqnro
        If procaprob Then
           StrSql = StrSql & " AND (proceso.proaprob = 3 OR proceso.proaprob = 2) "
        ElseIf Not todospro Then
           StrSql = StrSql & " AND proceso.pronro = " & pronro
        End If
        StrSql = StrSql & " INNER JOIN detliq ON detliq.concnro = " & rs3!concnro & " AND detliq.cliqnro = cabliq.cliqnro "
        StrSql = StrSql & " WHERE cabliq.empleado = " & rsConsult!ternro
        StrSql = StrSql & " GROUP BY cabliq.empleado "
        OpenRecordset StrSql, rs4
     
        If Not rs4.EOF Then 'Si tiene ticket
           imptp = CDbl(rs4!dlimonto)
           suetot = CDbl(suetot) + CDbl(imptp)
        End If
        If rs4.State = adStateOpen Then rs4.Close
     End If
     If rs3.State = adStateOpen Then rs3.Close
    
     imptic = CDbl(imptr) + CDbl(imptc)
     
     If imptic <> 0 Then
        porctr = Round((CDbl(imptr) * 100) / CDbl(imptic), 0)
        porctc = 100 - porctr
     End If
     '-------------------------------------------------------------------------------
     'Inserto los datos en la BD

     '-------------------------------------------------------------------------------
     StrSql = "INSERT INTO rep_list_tick_det (bpronro,ternro,empleg,apellido,nombre,categoria," & _
              "centrocosto,sueldo,suetot,imptic,porctr,porctc,imptr,imptc,imptp, ccostocodext, catcodext) VALUES (" & _
              NroProceso & "," & ternro & "," & empleg & ",'" & apellido & "','" & nombre & "','" & _
              categoria & "','" & centrocosto & "'," & sueldo & "," & suetot & "," & imptic & "," & _
              porctr & "," & porctc & "," & imptr & "," & imptc & "," & imptp & ",'" & ccostocodext & "','" & catcodext & "')"
     objConn.Execute StrSql, , adExecuteNoRecords
     
     rsConsult.MoveNext
    
     'Actualizo el progreso
     TiempoAcumulado = GetTickCount
     Progreso = Progreso + IncPorc
     cantidadProcesada = cantidadProcesada - 1
     StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
     StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'"
     StrSql = StrSql & ", bprcempleados ='" & CStr(cantidadProcesada) & "' WHERE bpronro = " & NroProceso
     objConn.Execute StrSql, , adExecuteNoRecords
Loop

If rsConsult.State = adStateOpen Then rsConsult.Close

Exit Sub
            
MError:
    Flog.Writeline " Error: " & Err.Description
    HuboErrores = True
    EmpErrores = True
    Exit Sub
End Sub
