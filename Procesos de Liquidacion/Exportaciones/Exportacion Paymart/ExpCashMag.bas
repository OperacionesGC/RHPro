Attribute VB_Name = "ExpCashManagement"
Option Explicit

'Global Const Version = "1.00"
'Global Const FechaModificacion = "25/10/2010" 'Verónica Bogado
'Global Const UltimaModificacion = "Version Inicial"

'Global Const Version = "1.10"
'Global Const FechaModificacion = "16/11/2010" ' MB
'global Const UltimaModificacion = "Correcciones varias3"

'Global Const Version = "1.11"
'Global Const FechaModificacion = "09/12/2010" ' MB
'Global Const UltimaModificacion = "Correcciones varias"

'Global Const Version = "1.12"
'Global Const FechaModificacion = "07/01/2011" ' MB
'Global Const UltimaModificacion = "Error en desbordamiento por var entera de legajo"

'Global Const Version = "1.13"
'Global Const FechaModificacion = "07/02/2011" ' Verónica Bogado
'Global Const UltimaModificacion = "Inconsitencia entre montos Employee y Header"

Global Const Version = "1.14"
Global Const FechaModificacion = "09/03/2011" ' Verónica Bogado
Global Const UltimaModificacion = "Informar cantidad en Pay_Detail solo para los conceptos incluidos en acumulador de horas (18)"


Public Sub Main()
Dim objconnMain As New ADODB.Connection
Dim strCmdLine
Dim ArrParametros
Dim Nombre_Arch As String
Dim rs_batch_proceso As New ADODB.Recordset
Dim PID As String
Dim bprcparam As String
    
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
        
    'Abro el log del proceso
    Nombre_Arch = PathFLog & "RHProExpPayMart-" & NroProcesoBatch & ".log"
    
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
  
On Error GoTo ME_Main
            
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
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0 ,bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Flog.writeline "Pone el estado en procesando"
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 277 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        bprcparam = rs_batch_proceso!bprcparam
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        Call generaExp(NroProcesoBatch, bprcparam)
    Else
        Flog.writeline "no encontró el proceso"
    End If
    
    TiempoFinalProceso = GetTickCount
    Flog.writeline "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    
    If Not HuboError Then
        StrSql = "UPDATE batch_proceso SET bprcprogreso = 100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
    Else
        StrSql = "UPDATE batch_proceso SET bprcprogreso = 100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
    End If
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
Fin:
    Flog.Close
    'If objConn.State = adStateOpen Then objConn.Close
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


Public Function FechaSbarra(Fecha) As String
Dim auxi As String
    auxi = Year(Fecha)
    If Month(Fecha) < 10 Then
      auxi = auxi & "0" & Month(Fecha)
    Else
      auxi = auxi & Month(Fecha)
    End If
    If Day(Fecha) < 10 Then
      auxi = auxi & "0" & Day(Fecha)
    Else
      auxi = auxi & Day(Fecha)
    End If
    FechaSbarra = auxi
End Function


Public Sub generaExp(ByVal BproNro As Long, ByVal Parametros As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que genera los archivos correspondientes al grupo de exportación.
'Arreglo que contiene los parametros
Dim arrParam

'Parametros desde ASP
Dim ListaProceso As String


'RecordSet
Dim rs_Empleados As New ADODB.Recordset
Dim rs_Consult As New ADODB.Recordset
Dim rs_confrep As New ADODB.Recordset
Dim rs_Headerm As New ADODB.Recordset
Dim rs_Headerh As New ADODB.Recordset
Dim rs_Countemp As New ADODB.Recordset
Dim rs_paismon As New ADODB.Recordset


'Variables
Dim fs, f
Dim Path As String
Dim NroReporte As Integer
Dim NroModelo As Long
Dim directorio As String
Dim Carpeta
Dim ArchDef As String
Dim Nombre_Arch As String
Dim Dir_Arch1 As String
Dim Dir_Arch2 As String
Dim Dir_Arch3 As String
Dim Dir_Arch4 As String
Dim LineaEmploy As String
Dim LineaPayelem As String
Dim LineaHead As String
Dim Sep As String
Dim hizhead As Boolean
Dim ArrParametros As String
Dim ArchExpHeader
Dim ArchExpEmployee
Dim ArchExpPayelements
Dim ArchExpDetail
Dim ArchHeader As String
Dim ArchEmployee As String
Dim ArchPayelements As String
Dim ArchDetail As String
Dim Empant As Double
Dim Secuencia As String
Dim FecCorrida As String
'Acumuladores del confrep
Dim Lista_Moneyamount As String
Dim Lista_Hsamount As String
Dim Lista_Totalproc As String
Dim Lista_Tothsproc As String

'Empleado anterior, por defecto 0 para que tome el primero
Empant = 0
' Inicio codigo ejecutable
On Error GoTo CE

' Levanto cada parametro por separado, el separador de parametros es "@"
Flog.writeline Espacios(Tabulador * 0) & "levantando parametros" & Parametros
If Not IsNull(Parametros) Then
  Flog.writeline Espacios(Tabulador * 1) & "Parametros " & Parametros
  arrParam = Split(Parametros, "@")
  If UBound(arrParam) = 1 Then
    ListaProceso = arrParam(0)
    Flog.writeline Espacios(Tabulador * 1) & "Procesos = " & ListaProceso
    Nombre_Arch = arrParam(1)
    Flog.writeline Espacios(Tabulador * 1) & "Archivo a generar = " & Nombre_Arch
  Else
    Flog.writeline Espacios(Tabulador * 0) & "ERROR. La cantidad de parametros no es la esperada."
    HuboError = True
    Exit Sub
  End If
Else
  Flog.writeline Espacios(Tabulador * 0) & "ERROR. No se encuentran los paramentros."
  HuboError = True
  Exit Sub
End If
Flog.writeline


'---------------------------------------------------------------------------------------------------
'Datos del modelo y reporte
'---------------------------------------------------------------------------------------------------
NroModelo = 331
NroReporte = 294
Flog.writeline Espacios(Tabulador * 0) & "Buscando datos del modelo " & NroModelo
StrSql = "SELECT sis_dirsalidas FROM sistema"
OpenRecordset StrSql, rs_Consult
If Not rs_Consult.EOF Then
   directorio = Trim(rs_Consult!sis_dirsalidas)
End If
rs_Consult.Close

StrSql = "SELECT * FROM modelo WHERE modnro = " & NroModelo
OpenRecordset StrSql, rs_Consult
If Not rs_Consult.EOF Then
   Sep = IIf(Not IsNull(rs_Consult!modseparador), rs_Consult!modseparador, "")
   If Not IsNull(rs_Consult!modarchdefault) Then
      ArchDef = Trim(rs_Consult!modarchdefault)
   Else
      Flog.writeline Espacios(Tabulador * 1) & "ERROR. El modelo no tiene configurada la carpeta destino."
      HuboError = True
      Exit Sub
   End If
Else
    Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontró el modelo " & NroModelo
    HuboError = True
    Exit Sub
End If
         
directorio = directorio & Trim(rs_Consult!modarchdefault)
'directorio = "C:\log\ExpPaymart\" para pruebas locales

         
'---------------------------------------------------------------------------------------------------
'Creacion de los archivos
'---------------------------------------------------------------------------------------------------
'Asigna los sectores fijos y variables del nombre de cada archivo según contenido
'---------------------------------------------------------------------------------------------------


FecCorrida = FechaSbarra(Date)


'verifica el nro de secuencia a asignar buscando en la carpeta
Secuencia = Secuen(FecCorrida, directorio)

ArchHeader = "PAY_HEADER.RHPRO." & FecCorrida & "." & Secuencia & ".txt"
ArchEmployee = "EMPLOYEES.RHPRO." & FecCorrida & "." & Secuencia & ".txt"
ArchPayelements = "PAY_ELEMENTS.RHPRO." & FecCorrida & "." & Secuencia & ".txt"
ArchDetail = "PAY_DETAIL.RHPRO." & FecCorrida & "." & Secuencia & ".txt"

Dir_Arch1 = directorio & ArchHeader
Dir_Arch2 = directorio & ArchEmployee
Dir_Arch3 = directorio & ArchPayelements
Dir_Arch4 = directorio & ArchDetail

Flog.writeline Espacios(Tabulador * 1) & "Generando Archivo " & Dir_Arch1
Flog.writeline Espacios(Tabulador * 1) & "Generando Archivo " & Dir_Arch2
Flog.writeline Espacios(Tabulador * 1) & "Generando Archivo " & Dir_Arch3
Flog.writeline Espacios(Tabulador * 1) & "Generando Archivo " & Dir_Arch4

On Error Resume Next
Set fs = CreateObject("Scripting.FileSystemObject")
Set ArchExpHeader = fs.CreateTextFile(Dir_Arch1, True)
Set ArchExpEmployee = fs.CreateTextFile(Dir_Arch2, True)
Set ArchExpPayelements = fs.CreateTextFile(Dir_Arch3, True)
Set ArchExpDetail = fs.CreateTextFile(Dir_Arch4, True)

If Err.Number <> 0 Then
     Flog.writeline Espacios(Tabulador * 1) & "La carpeta Destino no existe. Se creará."
     Set Carpeta = fs.CreateFolder(directorio)
     Set ArchExpHeader = fs.CreateTextFile(Dir_Arch1, True)
     Set ArchExpEmployee = fs.CreateTextFile(Dir_Arch2, True)
     Set ArchExpPayelements = fs.CreateTextFile(Dir_Arch3, True)
     Set ArchExpDetail = fs.CreateTextFile(Dir_Arch4, True)
End If
'desactivo el manejador de errores
On Error GoTo CE
'--------------------------------------------------------------------------------
'Carga los parámetros de confrep
'--------------------------------------------------------------------------------
StrSql = "select * from confrep where repnro= " & NroReporte
OpenRecordset StrSql, rs_confrep
If rs_confrep.EOF Then
  Flog.writeline "Falta la configuración del reporte en confrep"
Else
  'Asigna cero inicialmente para concatenar derecho la coma y que quede bien.
  Lista_Moneyamount = "0"
  Lista_Hsamount = "0"
  Lista_Totalproc = "0"
  Lista_Tothsproc = "0"
  
  While Not rs_confrep.EOF
    Select Case UCase(rs_confrep("conftipo"))
      Case "ACM":
          Lista_Moneyamount = Lista_Moneyamount & "," & rs_confrep("confval")
      Case "ACH":
          Lista_Hsamount = Lista_Hsamount & "," & rs_confrep("confval")
      Case Else
        Flog.writeline "columna de configuración de reporte no reconocida " & rs_confrep("conftipo")
    End Select
  rs_confrep.MoveNext
  Wend
End If

'---------------------------------------------------------------------------------
'Consulta Principal
'---------------------------------------------------------------------------------
StrSql = "SELECT proceso.pronro, prodesc, pliqnro, tprocnro, profecpago,"
StrSql = StrSql & "profecini, profecfin, cabliq.cliqnro, empleado, empleg "
StrSql = StrSql & "FROM proceso INNER JOIN cabliq "
StrSql = StrSql & "ON cabliq.pronro = proceso.pronro "
StrSql = StrSql & "INNER JOIN empleado ON cabliq.empleado=empleado.ternro "
StrSql = StrSql & "INNER JOIN tercero ON tercero.ternro=empleado.ternro "
'StrSql = StrSql & "INNER JOIN pais ON tercero.paisnro=pais.paisnro "
'StrSql = StrSql & "INNER JOIN moneda ON moneda.paisnro = pais.paisnro "
'Se quita join a pais y moneda, los datos de moneda de pago se informan
'según la moneda defecto de sistema, que es la del país de ubicación
StrSql = StrSql & " WHERE proceso.pronro IN (" & ListaProceso & ") "
StrSql = StrSql & " ORDER BY empleado.empleg"
OpenRecordset StrSql, rs_Empleados

'Control Line
Flog.writeline "Consulta principal con parámetros cargados: " & StrSql

'--------------------------------------------------------------------------------
'             Consulta a País y moneda defecto (Para todos)
'--------------------------------------------------------------------------------
StrSql = "SELECT paiscodext codpais, codext codmoneda, paiscod_bco FROM pais"
StrSql = StrSql & " INNER JOIN moneda ON pais.paisnro=moneda.paisnro"
StrSql = StrSql & " WHERE pais.paisdef=-1"
OpenRecordset StrSql, rs_paismon

Flog.writeline "Consulta país y moneda defecto: " & StrSql

'---------------------------------------------------------------------------------
'seteo de las variables de progreso
'---------------------------------------------------------------------------------
Progreso = 0
CEmpleadosAProc = rs_Empleados.RecordCount
If CEmpleadosAProc = 0 Then
   Flog.writeline "no hay empleados"
   CEmpleadosAProc = 1
End If
IncPorc = (99 / CEmpleadosAProc)
        
        
Flog.writeline
Flog.writeline
Flog.writeline
Flog.writeline Espacios(Tabulador * 0) & "------------------------------------------------------------------------"
Flog.writeline Espacios(Tabulador * 0) & "Comienza el procesamiento de Registros"
Flog.writeline Espacios(Tabulador * 0) & "------------------------------------------------------------------------"
        
If Not (rs_Empleados.EOF) Then
  ArchExpHeader.writeline "EFFECTIVE_DATE" & Sep & "LEGIS_CD" & Sep & "SRC_BATCH_NAME" & Sep & "SRC_BATCH_SUSPENSE_D" & Sep & "PROVIDER_NAME" & Sep & "PAY_PERIOD_START" & Sep & "PAY_PERIOD_END" & Sep & "RECORD_COUNT" & Sep & "PERSON_COUNT" & Sep & "CURRENCY_CODE" & Sep & "TOTAL_MONEY_AMOUNT" & Sep & "TOTAL_HOURS_AMOUNT"
  Call LineaHeader(Sep, LineaHead, ListaProceso, ArchExpHeader, Lista_Moneyamount, Lista_Hsamount, ArchDetail, rs_Empleados!profecpago, rs_Empleados!profecini, rs_Empleados!profecfin, rs_paismon!codmoneda, rs_paismon!codpais)
  ArchExpHeader.writeline LineaHead
  Call LineaPayelements(ArchExpPayelements, rs_Empleados!profecpago, rs_paismon!codpais, rs_Empleados!Empleado, ListaProceso, Sep, ArchDetail)
  ArchExpEmployee.writeline "PERSON_ID" & Sep & "EMP_NUM" & Sep & "LEGIS_CD" & Sep & "CURRENCY_CODE" & Sep & "MONEY_AMOUNT" & Sep & "HOURS_AMOUNT" & Sep & "SRC_BATCH_NAME" & Sep & "SRC_BATCH_SUSPENSE_D"
  ArchExpDetail.writeline "EFFECTIVE_DATE" & Sep & "ELEMENT_NAME" & Sep & "ELEMENT_ID" & Sep & "PERSON_ID" & Sep & "LEGIS_CD" & Sep & "CURRENCY_CODE" & Sep & "MONEY_AMOUNT" & Sep & "HOURS_AMOUNT" & Sep & "PAY_PERIOD_START" & Sep & "PAY_PERIOD_END" & Sep & "SRC_BATCH_NAME" & Sep & "SRC_BATCH_SUSPENSE_D"
  Do While Not rs_Empleados.EOF
    If Empant <> rs_Empleados!Empleado Then
      Call LineaPaydetail(rs_Empleados!Empleado, ArchDetail, ArchExpDetail, ListaProceso, rs_Empleados!profecpago, Sep, ArchDetail, rs_Empleados!empleg, rs_paismon!codpais, rs_paismon!codmoneda, rs_Empleados!profecini, rs_Empleados!profecfin)
      Call LineaEmployees(rs_Empleados!Empleado, ListaProceso, LineaEmploy, rs_Empleados!profecpago, ArchDetail, rs_Empleados!empleg, rs_paismon!codpais, rs_paismon!codmoneda, rs_Empleados!profecini, rs_Empleados!profecfin, Sep, Lista_Moneyamount, Lista_Hsamount)
      ArchExpEmployee.writeline LineaEmploy
      Empant = rs_Empleados!Empleado
    End If
    '---------------------------------------------------------------------------------
    'Actualizo el progreso
    '---------------------------------------------------------------------------------
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
             "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    rs_Empleados.MoveNext
  Loop
Else
  Flog.writeline Espacios(Tabulador * 0) & "La consulta principal no produjo registros"
End If


If rs_Empleados.State = adStateOpen Then rs_Empleados.Close
If rs_Consult.State = adStateOpen Then rs_Consult.Close


Set rs_Empleados = Nothing
Set rs_Consult = Nothing

Exit Sub

CE:
    Flog.writeline "=================================================================="
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="
    
    
    MyBeginTrans
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
             "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    MyCommitTrans
    
    HuboError = True

End Sub

Sub LineaPaydetail(Empleado, Archivo, ArchName, Proceso, Fecheff, Sep, ArchStr, Legajo, Cpais, Cmone, Fini, Ffin)
      
Dim LineaDetail As String
Dim rsLinpay As New ADODB.Recordset
Dim rs_elemhs As New ADODB.Recordset
Dim StrSql As String
Dim Rowhs As Variant
Dim ix As Integer
Dim banconc As Boolean 'Controla que la cantidad se haya imprimido para colocar ceros
Dim Archadd

'Arma la consulta con los conceptos imprimibles a detallar
StrSql = "SELECT concepto.concabr concep , concepto.conccod, cabliq.empleado, detliq.dlimonto Monto, "
StrSql = StrSql & " detliq.dlicant Cant, concepto.concnro"
StrSql = StrSql & " FROM cabliq INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro and concepto.concimp=-1"
StrSql = StrSql & " WHERE cabliq.pronro in (" & Proceso & ") AND cabliq.empleado=" & Empleado
'StrSql = StrSql & " and (concepto.concimp = -1 or concepto.concnro in (select concnro from con_acum where acunro=18))"
'Para que no muestre "No imprimibles" que podrían estar dentro del acumulador 18
StrSql = StrSql & " order by conccod"
OpenRecordset StrSql, rsLinpay

Flog.writeline "Consulta para Detail: " & StrSql

'Trae todos los conceptos para en el acumulador de Horas, fijo 18
'Para discriminar al momento de imprimir.
StrSql = "select concnro from con_acum where acunro=18"
OpenRecordset StrSql, rs_elemhs
Flog.writeline "Trae los conceptos en el acumulador de Horas: " & StrSql

Rowhs = rs_elemhs.GetRows
rs_elemhs.Close

LineaDetail = ""
While Not rsLinpay.EOF
  'EFFECTIVE_DATE
  LineaDetail = LineaDetail & FechaSbarra(Fecheff)
  'ELEMENT NAME
  LineaDetail = LineaDetail & Sep & UCase(rsLinpay!concep) 'Nombre del concepto
  'ELEMENT ID
  LineaDetail = LineaDetail & Sep & rsLinpay!Conccod 'Código de concepto x cada concepto.
  'PERSON_ID
  LineaDetail = LineaDetail & Sep & Legajo 'Nro de Legajo del empleado
  'LEGIS CD
  LineaDetail = LineaDetail & Sep & Cpais 'Codigo país
  'CURRENCY CODE
  LineaDetail = LineaDetail & Sep & Cmone 'Código de Moneda (Externo)
  'MONEY AMOUNT
  LineaDetail = LineaDetail & Sep & Format(rsLinpay!Monto, "########0.00") 'valor del concepto
  'HOURS AMOUNT (Solo las cantidades de conceptos que implican horas y están incluidos en el acumulador 18
  'Enarbolado de if para que no moleste el select
  For ix = 0 To UBound(Rowhs, 2) '- 1
    If rsLinpay!concnro = Rowhs(0, ix) Then
      LineaDetail = LineaDetail & Sep & Format(rsLinpay!cant, "########0.00") 'Informa cant de horas formateada a nros
      banconc = True
    End If
  Next
  If banconc = False Then
    LineaDetail = LineaDetail & Sep & Format(0, "########0.00") 'No es de horas
  Else
    banconc = False 'Para que funcione en la siguiente vuelta
  End If
  'PAY PERIOD START
  LineaDetail = LineaDetail & Sep & FechaSbarra(Fini) 'Fecha del proceso
  'PAY PERIOD END
  LineaDetail = LineaDetail & Sep & FechaSbarra(Ffin) 'Fecha de fin del proceso
  'SRC_BATCH_NAME
  LineaDetail = LineaDetail & Sep & ArchStr
  'SRC BATCH SUSPENSE D
  LineaDetail = LineaDetail & Sep & FechaSbarra(Fecheff) 'Fecha planeada del proceso
  ArchName.writeline LineaDetail
  'Flog.writeline "Línea Pay Details: " & LineaDetail
rsLinpay.MoveNext
LineaDetail = ""
Wend
End Sub

Sub LineaEmployees(Empleado, Procesos, LineaEmploy, Fecproc, ArchName, Legajo, Cpais, Cmone, Fecini, Fecfin, Sep, Lst_Totalproc, Lst_Tothsproc)
  
Dim Fecstr As String
Dim rs_Headerm As New ADODB.Recordset
Dim rs_Headerh As New ADODB.Recordset

'Busco el monto total del proceso para el header.
StrSql = "SELECT sum(almonto) Monto FROM cabliq c INNER JOIN acu_liq al ON c.cliqnro=al.cliqnro"
StrSql = StrSql & " where c.empleado = " & Empleado & " AND al.acunro in (" & Lst_Totalproc & ") and c.pronro IN (" & Procesos & ")"
OpenRecordset StrSql, rs_Headerm

'Busco el acumulador de horas totales para el header
StrSql = "SELECT sum(alcant) Monto FROM cabliq c INNER JOIN acu_liq al ON c.cliqnro=al.cliqnro"
StrSql = StrSql & " where c.empleado = " & Empleado & " AND al.acunro in (" & Lst_Tothsproc & ") and c.pronro IN (" & Procesos & ")"
OpenRecordset StrSql, rs_Headerh
  
  
  
  Fecstr = FechaSbarra(Fecproc)
  'PERSON_ID - Legajo del empleado
  'LineaEmploy = Fecstr
  LineaEmploy = "99999"
  
  'EMP_NUM - Legajo del empleado
  LineaEmploy = LineaEmploy & Sep & Legajo
  
  'LEGIS_CD - Código de país
  LineaEmploy = LineaEmploy & Sep & Cpais
  
  'CURRENCY_CODE - Codigo de moneda
  LineaEmploy = LineaEmploy & Sep & Cmone
  
  'MONEY_AMOUNT - Monto ac. del empleado
  If EsNulo(rs_Headerm!Monto) Then
    LineaEmploy = LineaEmploy & Sep & "0.00"
  Else
    LineaEmploy = LineaEmploy & Sep & Format(rs_Headerm!Monto, "########0.00")
  End If
  
  'HOURS_AMOUNT - Total hs. acum del empleado
  If EsNulo(rs_Headerh!Monto) Then
    LineaEmploy = LineaEmploy & Sep & "0.00"
  Else
    LineaEmploy = LineaEmploy & Sep & Format(rs_Headerh!Monto, "########0.00")
  End If
  'SRC_BATCH_NAME - Nombre del archivo Employees
  LineaEmploy = LineaEmploy & Sep & ArchName
  
  'SRC_BATCH_SUSPENSE_D - fecha del proceso
  LineaEmploy = LineaEmploy & Sep & FechaSbarra(Fecproc)
  
  'Flog.writeline "Línea de Empleado: " & Legajo
End Sub

Sub LineaPayelements(ArchName, Fpago, Pais, Empleado, Procesos, Sep, ArchStr) '1 lin por cada concepto
  Dim RsPayelem As New Recordset
  Dim LineaPayelem As String
  Dim Fepago As String
  Fepago = FechaSbarra(Fpago)
  StrSql = "SELECT DISTINCT(concepto.concabr) concep , concepto.conccod "
  StrSql = StrSql & " FROM cabliq INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
  StrSql = StrSql & "INNER JOIN concepto ON concepto.concnro = detliq.concnro AND concepto.concimp = -1 "
  StrSql = StrSql & "WHERE cabliq.pronro IN (" & Procesos & ") "
  StrSql = StrSql & "ORDER BY concepto.conccod "
  OpenRecordset StrSql, RsPayelem
  
  ArchName.writeline "EFFECTIVE_DATE" & Sep & "LEGIS_CD" & Sep & "SRC_BATCH_NAME" & Sep & "SRC_BATCH_SUSPENSE_D" & Sep & "PROVIDER_NAME" & Sep & "ELEMENT_NAME" & Sep & "ELEMENT_ID"
  LineaPayelem = ""
  While Not RsPayelem.EOF
  'EFFECTIVE DATE
  LineaPayelem = LineaPayelem & Fepago
  'LEGIS_CD
  LineaPayelem = LineaPayelem & Sep & Pais
  'SRC_BATCH_NAME
  LineaPayelem = LineaPayelem & Sep & ArchStr
  'SRC_BATCH_SUSPENSE_D
  LineaPayelem = LineaPayelem & Sep & Fepago
  'PROVIDER_NAME
  LineaPayelem = LineaPayelem & Sep & "RHPRO"
  'ELEMENT_NAME
  LineaPayelem = LineaPayelem & Sep & UCase(RsPayelem!concep)
  'ELEMENT_ID
  LineaPayelem = LineaPayelem & Sep & RsPayelem!Conccod
  ArchName.writeline LineaPayelem
  'Flog.writeline "Linea en payelements " & LineaPayelem
  RsPayelem.MoveNext
  LineaPayelem = ""
  Wend

End Sub

Sub LineaHeader(Sep, LineaHead, Procesos, Archivo, Lst_Totalproc, Lst_Tothsproc, ArchNomina, FechaEffec, Inipago, Finpago, Cmone, Cpais)
Dim TotNomi As Integer

Dim rs_Headerm As New ADODB.Recordset
Dim rs_Headerh As New ADODB.Recordset
Dim rs_Countemp As New ADODB.Recordset
Dim rs_Totconce As New ADODB.Recordset

Flog.writeline "linea Header"

LineaHead = ""

'Busco el monto total del proceso para el header.
StrSql = "SELECT sum(round(almonto,2)) total FROM cabliq c INNER JOIN acu_liq al ON c.cliqnro=al.cliqnro"
StrSql = StrSql & " where al.acunro in (" & Lst_Totalproc & ") and c.pronro IN (" & Procesos & ")"
OpenRecordset StrSql, rs_Headerm

Flog.writeline "SQL Monto Total: " & StrSql


'Busco el acumulador de horas totales para el header
StrSql = "SELECT sum(round(alcant,2)) total FROM cabliq c INNER JOIN acu_liq al ON c.cliqnro=al.cliqnro"
StrSql = StrSql & " where al.acunro in (" & Lst_Tothsproc & ") and c.pronro IN (" & Procesos & ")"
OpenRecordset StrSql, rs_Headerh

Flog.writeline "SQL Horas Total: " & StrSql


'Busco el total de empleados de la corrida
StrSql = "SELECT distinct(empleado) from cabliq where pronro IN( " & Procesos & ")"
OpenRecordset StrSql, rs_Countemp
TotNomi = rs_Countemp.RecordCount

Flog.writeline "Cantidad de Empl: " & TotNomi

'Busco el total de conceptos imprimibles
StrSql = "SELECT count(*) Totimp FROM cabliq INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
StrSql = StrSql & "INNER JOIN concepto ON concepto.concnro = detliq.concnro"
StrSql = StrSql & " Where concepto.concimp = -1 "
StrSql = StrSql & "AND cabliq.pronro IN (" & Procesos & ")"
OpenRecordset StrSql, rs_Totconce

Flog.writeline "Cantidad de Conceptos Imprimibles: " & rs_Totconce!Totimp

  'EFFECTIVE_DATE (Fecha del proceso)
  LineaHead = LineaHead & FechaSbarra(FechaEffec)
  
  'LEGIS_CD (Codigo de país, tabla iso, cod externo)
  LineaHead = LineaHead & Sep & Cpais
  
  'SRC_BATCH_NAME (Nombre del archivo de la nómina de empleados)
  LineaHead = LineaHead & Sep & ArchNomina
  
  'SRC_BATCH_SUSPENSE_D (Fecha del proceso)
  LineaHead = LineaHead & Sep & FechaSbarra(FechaEffec)
  
  'PROVIDER NAME (Nombre del proveedor. RHPro fijo por pedido)
  LineaHead = LineaHead & Sep & "RHPRO"
  
  'PAY_PERIOD_START (Fecha de inicio del proceso)
  LineaHead = LineaHead & Sep & FechaSbarra(Inipago)
  
  'PAY_PERIOD_END (Fecha de fin del proceso - planeada -)
  LineaHead = LineaHead & Sep & FechaSbarra(Finpago)
  
  'RECORD_COUNT (Nro de registros en el fichero de detalle)
  LineaHead = LineaHead & Sep & rs_Totconce!Totimp
  
  'PERSON_COUNT Empleados liquidados en la corrida
  LineaHead = LineaHead & Sep & TotNomi
  
  'CURRENCY_CODE (Codigo de moneda en cod externo iso)
  LineaHead = LineaHead & Sep & Cmone
  
  'TOTAL_MONEY_AMOUNT (total a pagar en la corrida, configurado por confrep)
  If EsNulo(rs_Headerm!total) Then
    LineaHead = LineaHead & Sep & "0.00"
  Else
    LineaHead = LineaHead & Sep & Format(rs_Headerm!total, "########0.00")
  End If
  
  'TOTAL_HOURS_AMOUNT (total de horas usadas, configurable de confrep)
  If EsNulo(rs_Headerh!total) Then
    LineaHead = LineaHead & Sep & "0.00"
  Else
    LineaHead = LineaHead & Sep & Format(rs_Headerh!total, "########0.00")
  End If
  
  'Flog.writeline "Linea del Header: " & LineaHead
End Sub

Function Secuen(ByVal Fec As String, directorio As String) As String
Dim i As Integer
Dim hay As Integer
Dim Elarch As String
Dim fil
Dim salir As Integer
i = 0
hay = 0
salir = 0
Set fil = CreateObject("Scripting.FileSystemObject")
While (salir = 0 And i <= 99)
  If i < 10 Then
    Elarch = directorio & "EMPLOYEES.RHPRO." & Fec & "." & "0" & i & ".txt"
  Else
    Elarch = directorio & "EMPLOYEES.RHPRO." & Fec & "." & i & ".txt"
  End If
  If (fil.fileExists(Elarch)) Then
        hay = i + 1
    Else
        salir = 1
  End If
i = i + 1
Wend
If hay < 10 Then
  Secuen = "0" & hay
Else
  Secuen = hay
End If
End Function
