Attribute VB_Name = "repEstadAccid"
Option Explicit

'Const Version = 1.01
'Const FechaVersion = "19/10/2006"
''Modificaciones: Mariano Capriz
''               Se adapto el inicio del main para que corra con el nuevo appserver
''               Se agrego la version y log inicial

Global Const Version = "1.02" ' Cesar Stankunas
Global Const FechaVersion = "05/08/2009"
Global Const UltimaModificacion = ""    'Encriptacion de string connection

'---------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------

Dim fs, f

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

Global periDescDesde
Global periDescHasta
Global teDAbr1
Global tenro1
Global teDAbr2
Global tenro2
Global teDAbr3
Global tenro3

Global IdUser
Global Fecha
Global Hora

Global incapMuerte
Global incapPerman
Global thTrab1
Global thTrab2
Global thExtras1
Global thExtras2
Global thAccid1

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
Dim objRs2 As New ADODB.Recordset
Dim pgtidesde
Dim pgtihasta
Dim fdesde
Dim fhasta
Dim tipoDepuracion
Dim historico As Boolean
Dim param
Dim rsEstr As New ADODB.Recordset
Dim rsPeri As New ADODB.Recordset
Dim PID As String
Dim TiempoInicialProceso
Dim TiempoAcumulado
Dim cantRegistros
Dim totalRegistros
Dim arr
Dim cantPeriodos
Dim ArrParametros

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
    
    Call CargarConfiguracionesBasicas
    
    Nombre_Arch = PathFLog & "RepEstadAccid" & "-" & NroProceso & ".log"
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaVersion
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "Inicio Proceso Reporte Estadisticas Accidentes: " & Now
    
    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    depurar = False
    HuboErrores = False
    
    TiempoInicialProceso = GetTickCount
    
    Flog.writeline Espacios(Tabulador * 0) & "Inicio :" & Now
    
    'Cambio el estado del proceso a Procesando
    TiempoAcumulado = GetTickCount
    
    Flog.writeline "Cambio el estado del proceso a Procesando"
    
    StrSql = "UPDATE batch_proceso SET bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "', bprcpid = " & PID & ", bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Obtengo los datos del proceso"
       
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, objRs
    
    If Not objRs.EOF Then
    
       'Obtengo los parametros del proceso
       arr = Split(objRs!bprcparam, "@")
       tenro1 = arr(0)
       tenro2 = arr(1)
       tenro3 = arr(2)
       pgtidesde = arr(3)
       pgtihasta = arr(4)
       
       IdUser = objRs!IdUser
       Fecha = objRs!bprcfecha
       Hora = objRs!bprchora
       
       'Busco la fecha desde
       StrSql = "SELECT * FROM gti_per WHERE pgtinro = " & pgtidesde
       OpenRecordset StrSql, objRs
       
       If Not objRs.EOF Then
          fdesde = objRs!pgtidesde
          periDescDesde = objRs!pgtidesabr
       Else
          Flog.writeline "No se encontro el periodo desde"
          Exit Sub
       End If
       
       objRs.Close
       
       'Busco la fecha hasta
       StrSql = "SELECT * FROM gti_per WHERE pgtinro = " & pgtihasta
       OpenRecordset StrSql, objRs
       
       If Not objRs.EOF Then
          fhasta = objRs!pgtihasta
          periDescHasta = objRs!pgtidesabr
       Else
          Flog.writeline "No se encontro el periodo hasta"
          Exit Sub
       End If
       
       objRs.Close
       
       'Busco el tipo de estructura 1
       StrSql = "SELECT * FROM tipoestructura WHERE tenro = " & tenro1
       OpenRecordset StrSql, objRs
       
       If Not objRs.EOF Then
          teDAbr1 = objRs!teDAbr
       Else
          Flog.writeline "No se encontro el tipo de estructura 1"
          'Exit Sub
       End If
       
       'Busco el tipo de estructura 2
       StrSql = "SELECT * FROM tipoestructura WHERE tenro = " & tenro2
       OpenRecordset StrSql, objRs
       
       If Not objRs.EOF Then
          teDAbr2 = objRs!teDAbr
       Else
          Flog.writeline "No se encontro el tipo de estructura 2"
          'Exit Sub
       End If
       
       'Busco el tipo de estructura 3
       StrSql = "SELECT * FROM tipoestructura WHERE tenro = " & tenro3
       OpenRecordset StrSql, objRs
       
       If Not objRs.EOF Then
          teDAbr3 = objRs!teDAbr
       Else
          Flog.writeline "No se encontro el tipo de estructura 3"
          'Exit Sub
       End If
       
       objRs.Close
       
       'Obtengo la cantidad de periodos
       StrSql = "SELECT * FROM gti_per WHERE pgtidesde >= " & ConvFecha(fdesde)
       StrSql = StrSql & " AND pgtihasta <= " & ConvFecha(fhasta) & " ORDER BY pgtihasta "
       
       OpenRecordset StrSql, objRs
       
       cantPeriodos = 0
       
       If Not objRs.EOF Then
          cantPeriodos = objRs.RecordCount
       Else
          Flog.writeline "No se encontraron periodos"
          Exit Sub
       End If
       
       objRs.Close
       
       Flog.writeline "Generando datos para el tipo de estructura: " & teDAbr1
       Flog.writeline "En el rango de Fechas: " & fdesde & "-" & fhasta
       
       'Busco en el confrep los datos del tipo de hora extra y de las incapacidades
       StrSql = " SELECT * FROM confrep "
       StrSql = StrSql & " WHERE repnro = 87 "
      
       OpenRecordset StrSql, objRs2
       
       incapMuerte = 0
       incapPerman = 0
       thExtras1 = 0
       thExtras2 = 0
       thTrab1 = 0
       thTrab2 = 0
       thAccid1 = 0
       
       If objRs2.EOF Then
          Flog.writeline "No esta configurado el ConfRep"
          Exit Sub
       End If
       
       Flog.writeline "Obtengo los datos del confrep"
       
       Do Until objRs2.EOF
       
          Select Case objRs2!confnrocol
             Case 1
                  incapMuerte = objRs2!confval
             Case 2
                  incapPerman = objRs2!confval
             Case 3
                  thTrab1 = objRs2!confval
             Case 4
                  thTrab2 = objRs2!confval
             Case 5
                  thExtras1 = objRs2!confval
             Case 6
                  thExtras2 = objRs2!confval
             Case 7
                  thAccid1 = objRs2!confval
                  
          End Select
       
          objRs2.MoveNext
       Loop
       
       'EMPIEZA EL PROCESO
       
       'Obtengo las estructuras
       StrSql = "SELECT * FROM estructura WHERE tenro = " & tenro1 & " ORDER BY estrdabr "
       
       OpenRecordset StrSql, rsEstr
       
       totalRegistros = rsEstr.RecordCount
       totalRegistros = totalRegistros * cantPeriodos
       cantRegistros = totalRegistros
       
       'Genero por cada empleado/fecha los horarios
       Do Until rsEstr.EOF
       
            'Obtengo los periodos
            StrSql = "SELECT * FROM gti_per WHERE pgtidesde >= " & ConvFecha(fdesde)
            StrSql = StrSql & " AND pgtihasta <= " & ConvFecha(fhasta) & " ORDER BY pgtihasta "
            
            OpenRecordset StrSql, rsPeri
       
            Do Until rsPeri.EOF
          
                Flog.writeline "Generando datos para la estructura " & rsEstr!estrdabr & " en el periodo " & rsPeri!pgtidesabr
          
                Call generarDatos(rsEstr!estrnro, rsEstr!estrdabr, rsPeri!pgtinro, rsPeri!pgtidesde, rsPeri!pgtihasta, rsPeri!pgtianio, rsPeri!pgtimes)
                
                cantRegistros = cantRegistros - 1
                
                rsPeri.MoveNext
            Loop
          
            TiempoAcumulado = GetTickCount
          
            StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Fix(((totalRegistros - cantRegistros) * 100) / totalRegistros) & _
                      ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                      ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
                 
            objConn.Execute StrSql, , adExecuteNoRecords
          
            rsEstr.MoveNext
       Loop
    
    Else
        Exit Sub
    End If
   
    'Actualizo el estado del proceso
    If Not HuboErrores Then
       StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
    Else
       StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProceso
    End If
    
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Fin :" & Now
    Flog.Close

    Exit Sub
    
CE:
    HuboErrores = True
    'Flog.writeline " Error: " & Err.Description & Now
    Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
        Flog.writeline Espacios(Tabulador * 0) & "Reporte abortado por Error:" & " " & Now
        Flog.writeline Espacios(Tabulador * 0) & " Error General " & Err.Description
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "SQL Ejecutado: " & StrSql
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
        Flog.writeline
End Sub

'--------------------------------------------------------------------
' Se encarga de generar los datos
'--------------------------------------------------------------------
Sub generarDatos(estrnro, estrdabr, pgtinro, pgtidesde As Date, pgtihasta As Date, Anio As Integer, Mes As Integer)

Dim StrSql As String
Dim SqlEmpleados As String
Dim SqlAccid As String
Dim rsConsult As New ADODB.Recordset
Dim rsEmpl As New ADODB.Recordset

Dim dt_m
Dim em_m
Dim acc_m
Dim dp_m
Dim mu_m
Dim ilp_m
Dim hextras

On Error GoTo MError

dt_m = 0
em_m = 0
acc_m = 0
dp_m = 0
mu_m = 0
ilp_m = 0
hextras = 0

'------------------------------------------------------------------
'Armo la SQL que se encarga de indicar los empleados de la estructura para el periodo
'------------------------------------------------------------------
SqlEmpleados = " SELECT DISTINCT gti_horcumplido.ternro "
SqlEmpleados = SqlEmpleados & " From gti_horcumplido "
SqlEmpleados = SqlEmpleados & " INNER JOIN his_estructura ON his_estructura.ternro = gti_horcumplido.ternro AND estrnro = " & estrnro
SqlEmpleados = SqlEmpleados & " WHERE thnro IN (" & thTrab1 & "," & thTrab2 & ")"
SqlEmpleados = SqlEmpleados & " AND gti_horcumplido.horfecrep >= " & ConvFecha(pgtidesde) & " AND gti_horcumplido.horfecrep <= " & ConvFecha(pgtihasta)

'------------------------------------------------------------------
'Controlo si ya existe un historico con datos
'------------------------------------------------------------------

StrSql = " SELECT * "
StrSql = StrSql & " From gti_rep_estacc_his "
StrSql = StrSql & " Where estrnro = " & estrnro
StrSql = StrSql & " AND pgtinro = " & pgtinro

OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
    'Si existe uso los datos almacenados
    
    em_m = IIf(IsNull(rsConsult!em), 0, rsConsult!em)
    dt_m = IIf(IsNull(rsConsult!dt), 0, rsConsult!dt)
    hextras = IIf(IsNull(rsConsult!hextras), 0, rsConsult!hextras)
    
    rsConsult.Close
    
Else
    'Si no existe busco los datos

    rsConsult.Close

    '------------------------------------------------------------------
    'Busco los empleados de la estructura
    '------------------------------------------------------------------
   
    OpenRecordset SqlEmpleados, rsEmpl
    
    em_m = rsEmpl.RecordCount
    
    'Si no hay empleados en la estructura no busco datos
    If em_m = 0 Then
       rsEmpl.Close
       Exit Sub
    End If
    
    rsEmpl.Close
    
    'Busco la cantidad de dias habiles
    StrSql = " SELECT sum(horcant) AS total "
    StrSql = StrSql & " From gti_horcumplido "
    StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.ternro = gti_horcumplido.ternro AND his_estructura.estrnro = " & estrnro
    StrSql = StrSql & " Where gti_horcumplido.horfecrep >= " & ConvFecha(pgtidesde) & " AND gti_horcumplido.horfecrep <= " & ConvFecha(pgtihasta)
    StrSql = StrSql & " AND thnro IN (" & thTrab1 & "," & thTrab2 & ")"
    
    OpenRecordset StrSql, rsConsult
    
    dt_m = 0
    
    If Not rsConsult.EOF Then
       If IsNull(rsConsult!total) Then
          dt_m = 0
       Else
          dt_m = Round(Round(rsConsult!total, 0) / 8, 0)
       End If
    End If
    
    rsConsult.Close

    '------------------------------------------------------------------
    'Busco la cantidad de horas extras usando el horario cumplido
    '------------------------------------------------------------------
    
    StrSql = " SELECT sum(horcant) AS total "
    StrSql = StrSql & " From gti_horcumplido "
    StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.ternro = gti_horcumplido.ternro AND his_estructura.estrnro = " & estrnro
    StrSql = StrSql & " Where gti_horcumplido.horfecrep >= " & ConvFecha(pgtidesde) & " AND gti_horcumplido.horfecrep <= " & ConvFecha(pgtihasta)
    StrSql = StrSql & " AND thnro IN (" & thExtras1 & "," & thExtras2 & ")"
    
    OpenRecordset StrSql, rsConsult
    
    If Not rsConsult.EOF Then
       If IsNull(rsConsult!total) Then
          hextras = 0
       Else
          hextras = Round(rsConsult!total, 0)
       End If
    End If
    
    rsConsult.Close

End If

'------------------------------------------------------------------
'Busco la cantidad de accidentes en el mes
'------------------------------------------------------------------

StrSql = " SELECT count(accnro) AS suma "
StrSql = StrSql & " From soaccidente "
StrSql = StrSql & " Where accfecha >= " & ConvFecha(pgtidesde) & " AND accfecha <= " & ConvFecha(pgtihasta)
StrSql = StrSql & " AND empleado IN (" & SqlEmpleados & ")"
StrSql = StrSql & " AND accestadistica = -1 "
StrSql = StrSql & " AND accreapertura = 0 "

OpenRecordset StrSql, rsConsult

acc_m = 0

If Not rsConsult.EOF Then
   acc_m = Round(rsConsult!suma, 0)
End If

rsConsult.Close

'------------------------------------------------------------------
'Busco la cantidad de licencias por accidente, que no tengan accidentes relacionadas
'------------------------------------------------------------------
    
StrSql = " SELECT count(emp_lic.emp_licnro) AS total "
StrSql = StrSql & " From emp_lic "
StrSql = StrSql & " INNER JOIN lic_accid ON lic_accid.emp_licnro = emp_lic.emp_licnro AND lic_accid.accnro = 0 "
StrSql = StrSql & " Where "
StrSql = StrSql & " (  (elfechadesde >= " & ConvFecha(pgtidesde) & " AND elfechadesde <= " & ConvFecha(pgtihasta) & " ) "
StrSql = StrSql & " OR (elfechahasta >= " & ConvFecha(pgtidesde) & " AND elfechahasta <= " & ConvFecha(pgtihasta) & " ) "
StrSql = StrSql & " OR (elfechadesde <= " & ConvFecha(pgtidesde) & " AND elfechahasta >= " & ConvFecha(pgtihasta) & " ) "
StrSql = StrSql & " ) "
StrSql = StrSql & " AND emp_lic.empleado IN (" & SqlEmpleados & ")"

OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   acc_m = acc_m + Round(rsConsult!total, 0)
End If

rsConsult.Close

'------------------------------------------------------------------
'Busco la cantidad de dias perdidos por accidentes
'------------------------------------------------------------------

StrSql = " SELECT emp_lic.* "
StrSql = StrSql & " From emp_lic "
StrSql = StrSql & " INNER JOIN lic_accid ON lic_accid.emp_licnro = emp_lic.emp_licnro "
StrSql = StrSql & " INNER JOIN soaccidente ON soaccidente.accnro = lic_accid.accnro  AND accestadistica = -1 "
StrSql = StrSql & " Where "
StrSql = StrSql & " (  (elfechadesde >= " & ConvFecha(pgtidesde) & " AND elfechadesde <= " & ConvFecha(pgtihasta) & " ) "
StrSql = StrSql & " OR (elfechahasta >= " & ConvFecha(pgtidesde) & " AND elfechahasta <= " & ConvFecha(pgtihasta) & " ) "
StrSql = StrSql & " OR (elfechadesde <= " & ConvFecha(pgtidesde) & " AND elfechahasta >= " & ConvFecha(pgtihasta) & " ) "
StrSql = StrSql & " ) "
StrSql = StrSql & " AND emp_lic.empleado IN (" & SqlEmpleados & ")"

OpenRecordset StrSql, rsConsult

Do Until rsConsult.EOF
   dp_m = dp_m + CantidadDeDias(pgtidesde, pgtihasta, rsConsult!elfechadesde, rsConsult!elfechahasta)
   
   rsConsult.MoveNext
Loop

rsConsult.Close

'------------------------------------------------------------------
'Busco la cantidad de muertes en el mes
'------------------------------------------------------------------

StrSql = " SELECT count(accnro) AS suma "
StrSql = StrSql & " From soaccidente "
StrSql = StrSql & " Where accfecha >= " & ConvFecha(pgtidesde) & " AND accfecha <= " & ConvFecha(pgtihasta)
StrSql = StrSql & " AND empleado IN (" & SqlEmpleados & ")"
StrSql = StrSql & " AND accestadistica = -1 "
StrSql = StrSql & " AND incapnro = " & incapMuerte

OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   mu_m = rsConsult!suma
End If

rsConsult.Close

'------------------------------------------------------------------
'Busco la cantidad de incapacidades permanentes en el mes
'------------------------------------------------------------------

StrSql = " SELECT count(accnro) AS suma "
StrSql = StrSql & " From soaccidente "
StrSql = StrSql & " Where accfecha >= " & ConvFecha(pgtidesde) & " AND accfecha <= " & ConvFecha(pgtihasta)
StrSql = StrSql & " AND empleado IN (" & SqlEmpleados & ")"
StrSql = StrSql & " AND accestadistica = -1 "
StrSql = StrSql & " AND incapnro = " & incapPerman

OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   ilp_m = rsConsult!suma
End If

rsConsult.Close

'------------------------------------------------------------------
'Armo la SQL para guardar los datos
'------------------------------------------------------------------
    
StrSql = " INSERT INTO gti_rep_estad_acci "
StrSql = StrSql & " (bpronro, Fecha, Hora, iduser,"
StrSql = StrSql & "  tenro1, estrnro1, tedabr1, estrdabr1,"
StrSql = StrSql & "  periodo_desde, periodo_hasta, pgtinro, dt_m,"
StrSql = StrSql & "  em_m, acc_m, dp_m, mu_m,"
StrSql = StrSql & "  ilp_m , hextras ) "
StrSql = StrSql & " VALUES"
StrSql = StrSql & "(" & NroProceso
StrSql = StrSql & "," & ConvFecha(Fecha)
StrSql = StrSql & ",'" & Hora & "'"
StrSql = StrSql & ",'" & IdUser & "'"
StrSql = StrSql & "," & tenro1
StrSql = StrSql & "," & estrnro
StrSql = StrSql & ",'" & teDAbr1 & "'"
StrSql = StrSql & ",'" & estrdabr & "'"
StrSql = StrSql & ",'" & periDescDesde & "'"
StrSql = StrSql & ",'" & periDescHasta & "'"
StrSql = StrSql & "," & pgtinro
StrSql = StrSql & "," & dt_m
StrSql = StrSql & "," & em_m
StrSql = StrSql & "," & acc_m
StrSql = StrSql & "," & dp_m
StrSql = StrSql & "," & mu_m
StrSql = StrSql & "," & ilp_m
StrSql = StrSql & "," & Int(hextras)
StrSql = StrSql & ")"

'------------------------------------------------------------------
'Guardo los datos en la BD
'------------------------------------------------------------------
objConn.Execute StrSql, , adExecuteNoRecords
    
Exit Sub

MError:
    Flog.writeline "Error en periodo: " & pgtinro & " Error: " & Err.Description
    HuboErrores = True
    Exit Sub

End Sub


Public Function bus_DiasHabiles(Ternro As Long, FechaDesde As Date, FechaHasta As Date)

Dim DiasHabiles As Single
Dim Dia As Date
Dim esFeriado As Boolean
Dim HACE_TRAZA As Boolean
    HACE_TRAZA = False

Dim objFeriado As New Feriado
  
' inicializacion de variables
Set objFeriado.Conexion = objConn
    'Set objFeriado.ConexionTraza = objConn

Bien = False
    
Dia = FechaDesde
DiasHabiles = 0
Do While Dia <= FechaHasta
    
    esFeriado = objFeriado.Feriado(Dia, Ternro, HACE_TRAZA)
    
    If Not esFeriado And Not Weekday(Dia) = 1 Then
        ' No es feriado no Domingo
        If Weekday(Dia) = 7 Then 'Sabado
           'DiasHabiles = DiasHabiles + 0.5
        Else
            DiasHabiles = DiasHabiles + 1
        End If
    End If
    Dia = Dia + 1
Loop

Bien = True
bus_DiasHabiles = DiasHabiles

End Function

