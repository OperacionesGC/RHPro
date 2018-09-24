Attribute VB_Name = "mdlSAP"
Option Explicit

' 27/02/2003 O.D.A. Cambio en la sentencia SQL que simula la vista de los empleados
'                   a los que puede acceder el usuario.

'Const Version = 1.01    'Inicial con nro de version
'Const FechaVersion = "11/10/2005"

Const Version = 1.01    'exportacion con Desdglose
Const FechaVersion = "18/11/2005"

'---------------------------------------------------------------------------
'---------------------------------------------------------------------------
Const ForReading = 1

Type Templeado
    Legajo As Long
    Ternro As Long
    Grupo As Long
End Type

Global empleado As Templeado
Global objBTurno As New BuscarTurno

Global Fecha As Date

Global fs
Global fSAP, f
Global FLOG
Global Nombre_Arch As String
Global StrLegajo As String
Global StrGrupo As String
Global strFecha As String


Global CEmpleadosAProc As Integer
Global CDiasAProc As Integer
Global IncPorc As Single
Global IncPorcEmpleado As Single
Global Progreso As Single
Global Estructura_Producto As Long
Global Lista_Thnro As String
Global Etiqueta
Global Repnro As Long

Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long

    
Public Sub Main()

Dim fechaDesde As Date
Dim fechaHasta As Date
Dim Legajo As Long
Dim nroProceso As Long
Dim myrs As New ADODB.Recordset

Dim strcmdLine As String
Dim tinicio
Dim i As Long


Dim Ternro As Long
Dim objRs As New ADODB.Recordset
Dim objrs_vemple As New ADODB.Recordset
Dim objrs_ea36 As New ADODB.Recordset
Dim TipoHora As Long
Dim Archivo As String
Dim strline As String
Dim pos1 As Integer
Dim pos2 As Integer

'Dim path As String
Dim NombreE36 As String

Dim LegajoAnterior As Long

Dim rs_Batch_Proceso As New ADODB.Recordset
Dim rs_His_Batch_Proceso As New ADODB.Recordset
Dim rs_Confrep As New ADODB.Recordset
Dim PID As String
Dim ArrParametros

    strcmdLine = Command()
    ArrParametros = Split(strcmdLine, " ", -1)
    If UBound(ArrParametros) > 0 Then
        If IsNumeric(ArrParametros(0)) Then
            nroProceso = ArrParametros(0)
            Etiqueta = ArrParametros(1)
        Else
            Exit Sub
        End If
    Else
        If IsNumeric(strcmdLine) Then
            nroProceso = strcmdLine
        Else
            Exit Sub
        End If
    End If

    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas

        
    Nombre_Arch = PathFLog & "Exportacion_SAP" & "-" & nroProceso & ".log"
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set FLOG = fs.CreateTextFile(Nombre_Arch, True)
    tinicio = Now
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    FLOG.writeline "-------------------------------------------------"
    FLOG.writeline "Version                  : " & Version
    FLOG.writeline "Fecha Ultima Modificacion: " & FechaVersion
    FLOG.writeline "PID                      : " & PID
    FLOG.writeline "-------------------------------------------------"
    FLOG.writeline
    
    'Abro la conexion
    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Then
        FLOG.writeline "Problemas en la conexion"
        Exit Sub
    End If
'    OpenConnection strconexion, objConnProgreso
'    If Err.Number <> 0 Then
'        FLOG.writeline "Problemas en la conexion"
'        Exit Sub
'    End If
    On Error GoTo 0
    
    On Error GoTo ME_Main
       
    'Abre el archivo INI
    FLOG.writeline "Abre el archivo INI"
    ''--------------------------------------------------------------------------------
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.OpenTextFile(App.Path & "\rhproSAP.INI", ForReading, 0)
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        PathSAP = Mid(strline, pos1, pos2 - pos1)
        If Right(PathSAP, 1) <> "\" Then PathSAP = PathSAP & "\"
    End If
    f.Close
    ''--------------------------------------------------------------------------------
    
    'Cambio el estado del proceso a Procesando
    FLOG.writeline "Cambio el estado del proceso a Procesando"
    StrSql = "UPDATE batch_proceso SET bprcpid = " & PID & ", bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcprogreso = 0 WHERE bpronro = " & nroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
       
    FLOG.writeline "Busco los datos del Proceso"
    StrSql = "SELECT bprcfecdesde,bprcfechasta,bprcparam,iduser FROM batch_proceso WHERE bpronro = " & nroProceso
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        fechaDesde = objRs!bprcfecdesde
        fechaHasta = objRs!bprcfechasta
    Else
        FLOG.writeline "No se encontró el proceso " & nroProceso
        Exit Sub
    End If
 
' Pseudo Vista de Empleado
'---------------------------------------------------------------------------
'StrSql = "SELECT * FROM empleado emp, estruc_actual ea1, estruc_actual ea2"
'StrSql = StrSql & " WHERE emp.Ternro = ea1.Ternro "
'StrSql = StrSql & " AND ea1.estrnro IN (SELECT upv.estrnro "
'StrSql = StrSql & " FROM   usupuedever upv "
'StrSql = StrSql & " Where upv.iduser = '" & objrs!iduser & "'"
'StrSql = StrSql & " AND upv.tenro   = 36 AND upv.estrnro = ea1.estrnro)"
'StrSql = StrSql & " AND emp.Ternro = ea2.Ternro "
'StrSql = StrSql & " AND ea2.estrnro IN (SELECT upv.estrnro "
'StrSql = StrSql & " FROM   usupuedever upv "
'StrSql = StrSql & " Where upv.iduser = '" & objrs!iduser & "'"
'StrSql = StrSql & " AND upv.tenro = 7 AND upv.estrnro = ea2.estrnro)"
' Se deja sin efecto esta version, por la que sigue. O.D.A. 27/02/2003
'---------------------------------------------------------------------------

'StrSql = "SELECT *"
'StrSql = StrSql & " FROM  empleado      emp,"
'StrSql = StrSql & "       estruc_actual ea07,"
'StrSql = StrSql & "       usupuedever   upv07,"
'StrSql = StrSql & "       estruc_actual ea36,"
'StrSql = StrSql & "       usupuedever   upv36"


' FGZ - 04/082003
' Cambio estruc_actual por his_estructura
StrSql = "SELECT ea36.tenro as ea36tenro, ea36.estrnro as ea36estrnro, emp.empleg, emp.ternro "
StrSql = StrSql & " FROM  empleado       emp,"
StrSql = StrSql & "       his_estructura ea07,"
StrSql = StrSql & "       usupuedever    upv07,"
StrSql = StrSql & "       his_estructura ea36,"
StrSql = StrSql & "       usupuedever    upv36"
' --------------------------------------------------

If (objRs!bprcparam <> "") Then
  StrSql = StrSql & ",    his_estructura he"
End If

StrSql = StrSql & " WHERE ea07.ternro    = emp.ternro"
StrSql = StrSql & " AND   ea07.tenro     = 7"
StrSql = StrSql & " AND   upv07.tenro    = ea07.tenro"
StrSql = StrSql & " AND   upv07.estrnro  = ea07.estrnro"
StrSql = StrSql & " AND   ea07.htethasta IS NULL"
StrSql = StrSql & " AND   upv07.iduser   = '" & objRs!IdUser & "'"

StrSql = StrSql & " AND   ea36.ternro    = emp.ternro"
StrSql = StrSql & " AND   ea36.tenro     = 36"
StrSql = StrSql & " AND   upv36.tenro    = ea36.tenro"
StrSql = StrSql & " AND   upv36.estrnro  = ea36.estrnro"
StrSql = StrSql & " AND   ea36.htethasta IS NULL"
StrSql = StrSql & " AND   upv36.iduser   = '" & objRs!IdUser & "'"
' Se simula la vista sobre los empleados, armando la lista de legajos
' a los que puede acceder el usuario actual, según las estructuras elegidas
' en USUPUEDEVER para los "Grupos de Seguridad" (7) y los "Centros de Empaque" (36).
' O.D.A. 27/02/2003

If (objRs!bprcparam <> "") Then
  StrSql = StrSql & " AND  he.estrnro    = " & objRs!bprcparam
  StrSql = StrSql & " AND  he.ternro     = emp.ternro"
  StrSql = StrSql & " AND  he.htethasta IS NULL"
End If
StrSql = StrSql & " ORDER BY emp.ternro"
' En forma opcional, puede indicar por parametro un argumento que es un numero de
' estructura para trabajar sobre un conjunto menor que todos los empleados que
' puede "ver" el usuario. O.D.A. 25/07/2003
 FLOG.writeline "Levanto los datos "
 OpenRecordset StrSql, objrs_vemple
 
 'carga los nombres de la tablas temporales
 FLOG.writeline "carga los nombres de la tablas temporales"
 Call CargarNombresTablasTemporales
 
 'crea la tabla temporal
 FLOG.writeline "crea la tabla temporal LSTA_EMPLE"
 CreateTempTable ("LSTA_EMPLE")
 
 If Not objrs_vemple.EOF Then
    objrs_vemple.MoveFirst
    
    StrSql = "SELECT estrcodext FROM estructura " & _
             " WHERE tenro = 36 AND estrnro = " & objrs_vemple!ea36estrnro
    OpenRecordset StrSql, objrs_ea36
    
    If Not objrs_ea36.EOF Then
        PathSAP = PathSAP + objrs_ea36!estrcodext + " _ "
    Else
        ' Error. No se encontró la estructura
        FLOG.writeline "Error. No se encontró la estructura"
        GoTo fin
    End If
 End If
 
 LegajoAnterior = 0
 Do While Not objrs_vemple.EOF
 
    If LegajoAnterior <> objrs_vemple!empleg Then
        StrSql = "INSERT INTO " & TTempLstaEmple & " VALUES( " & objrs_vemple!empleg & "," & objrs_vemple!Ternro & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        LegajoAnterior = objrs_vemple!empleg
        
    End If
   objrs_vemple.MoveNext
 Loop

' StrSql = "SELECT estructura.estrdabr, estructura.estrcodext"
' StrSql = StrSql & " FROM usupuedever"
' StrSql = StrSql & " INNER JOIN estructura"
' StrSql = StrSql & " ON    usupuedever.estrnro = estructura.estrnro"
' StrSql = StrSql & " AND   estructura.tenro    = 36 "
' StrSql = StrSql & " WHERE usupuedever.iduser  = '" & objRs!iduser & "'"
' OpenRecordset StrSql, objRs
' El nombre del archivo generado se forma con la estructura del Centro de Empaque
' de alguno de los empleados incluidos en el proceso y no a partir del usuario
' que lo ejecuta. O.D.A. 08/08/2003
 

'FGZ - 25/11/2004
'repnro = 62
Repnro = 53
StrSql = "SELECT * FROM confrep WHERE repnro = " & Repnro
StrSql = StrSql & " AND confnrocol = 40"
OpenRecordset StrSql, rs_Confrep
If Not rs_Confrep.EOF Then
    Estructura_Producto = rs_Confrep!confval
Else
    FLOG.writeline "No esta configurado el Tipo de Estructura Producto"
    Estructura_Producto = 0
End If


'FGZ - 18/11/2004
'Tipos de horas a considerar por el proceso
StrSql = "SELECT * FROM confrep WHERE repnro = " & Repnro
StrSql = StrSql & " AND confnrocol >= 60"
StrSql = StrSql & " AND conftipo = 'TH' "
StrSql = StrSql & " ORDER BY confnrocol"
If rs_Confrep.State = adStateOpen Then rs_Confrep.Close
OpenRecordset StrSql, rs_Confrep
Lista_Thnro = ""
Do While Not rs_Confrep.EOF
    If EsNulo(Lista_Thnro) Then
        Lista_Thnro = rs_Confrep!confval
    Else
        Lista_Thnro = Lista_Thnro & "'" & rs_Confrep!confval
    End If
    
    rs_Confrep.MoveNext
Loop
If EsNulo(Lista_Thnro) Then
    FLOG.writeline "Tipos de Hora a exportar no configurados. Se Usaran los tipos por default: 1,2,3"
    Lista_Thnro = "1,2,3"
End If
If rs_Confrep.State = adStateOpen Then rs_Confrep.Close
Set rs_Confrep = Nothing







' Seteo las variables de progreso
CEmpleadosAProc = objrs_vemple.RecordCount
If CEmpleadosAProc = 0 Then
    CEmpleadosAProc = 1
End If
CDiasAProc = DateDiff("d", fechaDesde, fechaHasta) + 1
If CDiasAProc <= 0 Then
    CDiasAProc = 1
End If
IncPorc = ((100 / CEmpleadosAProc) * (100 / CDiasAProc)) / 100
IncPorcEmpleado = (100 / CDiasAProc)
Progreso = 0
    
 ' Pseudo Vista de Empleado
    Fecha = fechaDesde
    
    Do While Fecha <= fechaHasta
        
        Archivo = PathSAP & "SAP " & Format(Fecha, "DD-MM-YYYY") & ".txt"
        Set fSAP = fs.CreateTextFile(Archivo, True)
        
        'StrSql = " SELECT estruc_actual.estrnro,estructura.estrcodext,Lsta_Emple.ternro,Lsta_Emple.empleg"
        'StrSql = StrSql & " FROM estruc_actual "
        'StrSql = StrSql & " INNER JOIN estructura ON estruc_actual.estrnro = estructura.estrnro "
        'StrSql = StrSql & " INNER JOIN Lsta_Emple ON estruc_actual.ternro = Lsta_Emple.ternro "
        'StrSql = StrSql & " WHERE (estruc_actual.tenro = 37) AND (estrcodext <> 'SDA') "
              
        ' FGZ  - 04/08/2003
        ' Cambio estruc_actual por his_estructura
        StrSql = " SELECT his_estructura.estrnro,estructura.estrcodext," & TTempLstaEmple & ".ternro," & TTempLstaEmple & ".empleg"
        StrSql = StrSql & " FROM his_estructura "
        StrSql = StrSql & " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro "
        StrSql = StrSql & " INNER JOIN " & TTempLstaEmple & " ON his_estructura.ternro = " & TTempLstaEmple & ".ternro "
        StrSql = StrSql & " WHERE (his_estructura.tenro = 37) AND (estructura.estrcodext <> 'SDA') "
        StrSql = StrSql & " AND his_estructura.htetdesde <= " & ConvFecha(fechaHasta)
        StrSql = StrSql & " AND (his_estructura.htethasta >= " & ConvFecha(fechaDesde)
        StrSql = StrSql & " OR his_estructura.htethasta IS NULL) "
        StrSql = StrSql & " ORDER BY " & TTempLstaEmple & ".ternro"
        'StrSql = StrSql & " AND his_estructura.htethasta IS NULL "
        
       ' -------------------------------
        OpenRecordset StrSql, myrs
        
        Do While Not myrs.EOF
        
           Ternro = myrs!Ternro
                
           strFecha = Format(Fecha, "DD/MM/YYYY")
           StrLegajo = Format(myrs!empleg, "000000")
           StrGrupo = Format(myrs!estrcodext, "000")
           
           FLOG.writeline "Exportar_Productividad(Ternro = " & Ternro & ", Fecha=" & Fecha & ")"
           'Call Exportar_Productividad(Ternro, Fecha)
           Call Exportar_Productividad_conDesglose(Ternro, Fecha)
           
           ' Actualizar el Progreso
           FLOG.writeline "Actualizar el Progreso"
           Progreso = Progreso + IncPorc
           StrSql = "UPDATE batch_proceso SET bprcprogreso = " & CInt(Progreso) & " WHERE bpronro = " & nroProceso
           objConn.Execute StrSql, , adExecuteNoRecords
                
           myrs.MoveNext
        Loop
        
        Fecha = DateAdd("d", 1, Fecha)
    Loop
    
    fSAP.Close
    Set fSAP = Nothing
    Set fs = Nothing
        
    
fin:
    FLOG.writeline "Fin"
    StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ",bprcestado = 'Procesado' WHERE bpronro = " & nroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
        
    ' -----------------------------------------------------------------------------------
    'FGZ - 22/09/2003
    'Copio el proceso en el historico de batch_proceso y lo borro de batch_proceso
        StrSql = "SELECT * FROM batch_proceso WHERE bpronro =" & nroProceso
        OpenRecordset StrSql, rs_Batch_Proceso

        
        StrSql = "INSERT INTO His_Batch_Proceso (bpronro,btprcnro,bprcfecha,iduser"
        StrSqlDatos = rs_Batch_Proceso!bpronro & "," & rs_Batch_Proceso!btprcnro & "," & _
                 ConvFecha(rs_Batch_Proceso!bprcfecha) & ",'" & rs_Batch_Proceso!IdUser & "'"
        
        If Not IsNull(rs_Batch_Proceso!bprchora) Then
            StrSql = StrSql & ",bprchora"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprchora & "'"
        End If
        If Not IsNull(rs_Batch_Proceso!bprcempleados) Then
            StrSql = StrSql & ",bprcempleados"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprcempleados & "'"
        End If
        If Not IsNull(rs_Batch_Proceso!bprcfecdesde) Then
            StrSql = StrSql & ",bprcfecdesde"
            StrSqlDatos = StrSqlDatos & "," & ConvFecha(rs_Batch_Proceso!bprcfecdesde)
        End If
        If Not IsNull(rs_Batch_Proceso!bprcfechasta) Then
            StrSql = StrSql & ",bprcfechasta"
            StrSqlDatos = StrSqlDatos & "," & ConvFecha(rs_Batch_Proceso!bprcfechasta)
        End If
        If Not IsNull(rs_Batch_Proceso!bprcestado) Then
            StrSql = StrSql & ",bprcestado"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprcestado & "'"
        End If
        If Not IsNull(rs_Batch_Proceso!bprcparam) Then
            StrSql = StrSql & ",bprcparam"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprcparam & "'"
        End If
        If Not IsNull(rs_Batch_Proceso!bprcprogreso) Then
            StrSql = StrSql & ",bprcprogreso"
            StrSqlDatos = StrSqlDatos & "," & rs_Batch_Proceso!bprcprogreso
        End If
        If Not IsNull(rs_Batch_Proceso!bprcfecfin) Then
            StrSql = StrSql & ",bprcfecfin"
            StrSqlDatos = StrSqlDatos & "," & ConvFecha(rs_Batch_Proceso!bprcfecfin)
        End If
        If Not IsNull(rs_Batch_Proceso!bprchorafin) Then
            StrSql = StrSql & ",bprchorafin"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprchorafin & "'"
        End If
        If Not IsNull(rs_Batch_Proceso!bprctiempo) Then
            StrSql = StrSql & ",bprctiempo"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprctiempo & "'"
        End If
        If Not IsNull(rs_Batch_Proceso!empnro) Then
            StrSql = StrSql & ",empnro"
            StrSqlDatos = StrSqlDatos & "," & rs_Batch_Proceso!empnro
        End If
        If Not IsNull(rs_Batch_Proceso!bprcPid) Then
            StrSql = StrSql & ",bprcPid"
            StrSqlDatos = StrSqlDatos & "," & rs_Batch_Proceso!bprcPid
        End If
        If Not IsNull(rs_Batch_Proceso!bprcfecInicioEj) Then
            StrSql = StrSql & ",bprcfecInicioEj"
            StrSqlDatos = StrSqlDatos & "," & ConvFecha(rs_Batch_Proceso!bprcfecInicioEj)
        End If
        If Not IsNull(rs_Batch_Proceso!bprcfecFinEj) Then
            StrSql = StrSql & ",bprcfecFinEj"
            StrSqlDatos = StrSqlDatos & "," & ConvFecha(rs_Batch_Proceso!bprcfecFinEj)
        End If
        If Not IsNull(rs_Batch_Proceso!bprcUrgente) Then
            StrSql = StrSql & ",bprcUrgente"
            StrSqlDatos = StrSqlDatos & "," & rs_Batch_Proceso!bprcUrgente
        End If
        If Not IsNull(rs_Batch_Proceso!bprcHoraInicioEj) Then
            StrSql = StrSql & ",bprcHoraInicioEj"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprcHoraInicioEj & "'"
        End If
        If Not IsNull(rs_Batch_Proceso!bprcHoraFinEj) Then
            StrSql = StrSql & ",bprcHoraFinEj"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprcHoraFinEj & "'"
        End If

        StrSql = StrSql & ") VALUES (" & StrSqlDatos & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        
        'Reviso que haya copiado
        StrSql = "SELECT * FROM His_batch_proceso WHERE bpronro =" & nroProceso
        OpenRecordset StrSql, rs_His_Batch_Proceso
        
        If Not rs_His_Batch_Proceso.EOF Then
            ' Borro de Batch_proceso
            StrSql = "DELETE FROM Batch_Proceso WHERE bpronro = " & nroProceso
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
    
        If rs_Batch_Proceso.State = adStateOpen Then rs_Batch_Proceso.Close
        If rs_His_Batch_Proceso.State = adStateOpen Then rs_His_Batch_Proceso.Close
    ' FGZ - 22/09/2003
    ' -----------------------------------------------------------------------------------
        
        
        
    If objRs.State = adStateOpen Then objRs.Close

    BorrarTempTable (TTempLstaEmple)

Exit Sub

'Manejador de Error
ME_Main:
    FLOG.writeline "Error: " & Err.Description
    FLOG.writeline "Ultimo SQL: " & StrSql
    
    StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ",bprcestado = 'Error' WHERE bpronro = " & nroProceso
    objConn.Execute StrSql, , adExecuteNoRecords

End Sub
    
    
Public Sub Exportar_Productividad(ByVal NroTer As Long, ByVal Fecha As Date)
' Ultima Modificacion: FGZ
'               Fecha: 25/11/2004
' Descripcion:       : Se agregó el producto en la linea de exportacion
'                       Antes exportaba: Legajo & Grupo & Fecha & Horas
'                       Ahora exporta: Legajo & Grupo & Producto & Fecha & Horas

Dim StrHoras As String
Dim strProducto As String
Dim TipoHora As Integer
Dim rsSuma As New ADODB.Recordset
Dim rs_Producto As New ADODB.Recordset


Set objBTurno.Conexion = objConn
   
        objBTurno.Buscar_Turno Fecha, NroTer, False
        If objBTurno.tiene_turno Then
                
                StrSql = " SELECT SUM(adcanthoras) as SumaHoras FROM gti_acumdiario "
                StrSql = StrSql & " WHERE adfecha = " & ConvFecha(Fecha) & " AND thnro IN (1,2,3) AND ternro = " & NroTer
                OpenRecordset StrSql, rsSuma
                If Not IsNull(rsSuma!sumahoras) Then
                    StrHoras = Replace(Format(rsSuma!sumahoras, "00.00"), ".", ",")
                    
                    'FGZ - 25/11/2004
                    'Busco el producto
                    StrSql = "SELECT estructura.estrcodext FROM estructura "
                    StrSql = StrSql & " INNER JOIN his_estructura ON estructura.estrnro = his_estructura.estrnro"
                    StrSql = StrSql & " WHERE estructura.tenro = " & Estructura_Producto
                    StrSql = StrSql & " AND his_estructura.ternro = " & NroTer
                    StrSql = StrSql & " AND his_estructura.htetdesde <= " & ConvFecha(Fecha)
                    StrSql = StrSql & " AND (his_estructura.htethasta >= " & ConvFecha(Fecha) & " OR htethasta IS NULL)"
                    OpenRecordset StrSql, rs_Producto
                    If Not rs_Producto.EOF Then
                        strProducto = Left(rs_Producto!estrcodext, 4)
                    Else
                        strProducto = "ERRO"
                    End If
                    
                    fSAP.writeline StrLegajo & StrGrupo & strProducto & strFecha & StrHoras
                
                End If
                
        End If
        
    'Cierro y libero
    If rs_Producto.State = adStateOpen Then rs_Producto.Close
    Set rs_Producto = Nothing
End Sub

Public Sub Exportar_Productividad_conDesglose(ByVal NroTer As Long, ByVal Fecha As Date)
' Ultima Modificacion: FGZ
'               Fecha: 25/11/2004
' Descripcion:       : Se agregó el producto en la linea de exportacion
'                       Antes exportaba: Legajo & Grupo & Fecha & Horas
'                       Ahora exporta: Legajo & Grupo & Producto & Fecha & Horas
' Ultima Modificacion: FGZ
'               Fecha: 18/11/2004
' Descripcion:       :

Dim StrHoras As String
Dim strProducto As String
Dim strLinea As String
Dim TipoHora As Integer
Dim Ya_Mostre As Boolean
Dim Sep As String
Dim Aux_cod_Estr As String

Dim rsSuma As New ADODB.Recordset
Dim rs_Producto As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim rs_Desglose As New ADODB.Recordset

Set objBTurno.Conexion = objConn
   
    Sep = "|"
    On Error GoTo ME_Local
    
        objBTurno.Buscar_Turno Fecha, NroTer, False
        If objBTurno.tiene_turno Then
                
                'Busco la estructura grupo de sap
                StrSql = "SELECT estructura.estrcodext FROM estructura "
                StrSql = StrSql & " INNER JOIN his_estructura ON estructura.estrnro = his_estructura.estrnro"
                StrSql = StrSql & " WHERE estructura.tenro = 37 "
                StrSql = StrSql & " AND his_estructura.ternro = " & NroTer
                StrSql = StrSql & " AND his_estructura.htetdesde <= " & ConvFecha(Fecha)
                StrSql = StrSql & " AND (his_estructura.htethasta >= " & ConvFecha(Fecha) & " OR htethasta IS NULL)"
                If rs_Producto.State = adStateOpen Then rs_Producto.Close
                OpenRecordset StrSql, rs_Producto
                If Not rs_Producto.EOF Then
                    StrGrupo = rs_Producto!estrcodext
                Else
                    StrGrupo = " ERROR "
                End If
                
                
                'Busco la estructura producto
                StrSql = "SELECT estructura.estrcodext FROM estructura "
                StrSql = StrSql & " INNER JOIN his_estructura ON estructura.estrnro = his_estructura.estrnro"
                StrSql = StrSql & " WHERE estructura.tenro = " & Estructura_Producto
                StrSql = StrSql & " AND his_estructura.ternro = " & NroTer
                StrSql = StrSql & " AND his_estructura.htetdesde <= " & ConvFecha(Fecha)
                StrSql = StrSql & " AND (his_estructura.htethasta >= " & ConvFecha(Fecha) & " OR htethasta IS NULL)"
                If rs_Producto.State = adStateOpen Then rs_Producto.Close
                OpenRecordset StrSql, rs_Producto
                If Not rs_Producto.EOF Then
                    strProducto = rs_Producto!estrcodext
                Else
                    strProducto = " ERROR "
                End If
                
                'Busco la estructura Linea (12)
                StrSql = "SELECT estructura.estrcodext FROM estructura "
                StrSql = StrSql & " INNER JOIN his_estructura ON estructura.estrnro = his_estructura.estrnro"
                StrSql = StrSql & " WHERE estructura.tenro = 12"
                StrSql = StrSql & " AND his_estructura.ternro = " & NroTer
                StrSql = StrSql & " AND his_estructura.htetdesde <= " & ConvFecha(Fecha)
                StrSql = StrSql & " AND (his_estructura.htethasta >= " & ConvFecha(Fecha) & " OR htethasta IS NULL)"
                If rs_Producto.State = adStateOpen Then rs_Producto.Close
                OpenRecordset StrSql, rs_Producto
                If Not rs_Producto.EOF Then
                    strLinea = rs_Producto!estrcodext
                Else
                    strLinea = " ERROR "
                End If
                
                
                StrSql = " SELECT thnro, SUM(adcanthoras) as SumaHoras FROM gti_acumdiario "
                StrSql = StrSql & " WHERE adfecha = " & ConvFecha(Fecha) & " AND thnro IN (" & Lista_Thnro & ") AND ternro = " & NroTer
                StrSql = StrSql & " GROUP BY thnro"
                OpenRecordset StrSql, rsSuma
                If Not rsSuma.EOF Then
                    If Not IsNull(rsSuma!sumahoras) Then
                        StrHoras = Replace(Format(rsSuma!sumahoras, "00.00"), ".", ",")
                        
                        'Debo buscar si tenia desglose y por la estructura producto ==>
                        StrSql = "SELECT * FROM gti_achdiario "
                        StrSql = StrSql & " WHERE ternro = " & NroTer
                        StrSql = StrSql & " AND achdfecha = " & ConvFecha(Fecha)
                        StrSql = StrSql & " AND thnro = " & rsSuma!thnro
                        If rs.State = adStateOpen Then rs.Close
                        OpenRecordset StrSql, rs
                        If Not rs.EOF Then
                            Ya_Mostre = False
                            Do While Not rs.EOF
                                StrHoras = Replace(Format(rs!achdcanthoras, "00.00"), ".", ",")
                                
                                'Hubo desglose
                                StrSql = "SELECT * FROM gti_achdiario_estr "
                                StrSql = StrSql & " WHERE achdnro = " & rs!achdnro
                                'StrSql = StrSql & " AND tenro = " & Estructura_Producto
                                StrSql = StrSql & " AND achdfecha = " & ConvFecha(Fecha)
                                If rs_Desglose.State = adStateOpen Then rs_Desglose.Close
                                OpenRecordset StrSql, rs_Desglose
                                If Not rs_Desglose.EOF Then
                                    Do While Not rs_Desglose.EOF
                                        'StrHoras = Replace(Format(rs!achdcanthoras, "00.00"), ".", ",")
                                        
                                        'Busco la estructura del desglose
                                        StrSql = "SELECT estructura.tenro, estructura.estrcodext FROM estructura "
                                        StrSql = StrSql & " WHERE estructura.estrnro = " & rs_Desglose!estrnro
                                        If rs_Producto.State = adStateOpen Then rs_Producto.Close
                                        OpenRecordset StrSql, rs_Producto
                                        If Not rs_Producto.EOF Then
                                            'strProducto = Left(rs_Producto!estrcodext, 4)
                                            Aux_cod_Estr = rs_Producto!estrcodext
                                        Else
                                            Aux_cod_Estr = " ERROR "
                                        End If
                                    
                                        Select Case rs_Producto!tenro
                                        Case 12:    'linea
                                            strLinea = Aux_cod_Estr
                                        Case 37:    'Grupo SAP
                                            StrGrupo = Aux_cod_Estr
                                        Case Estructura_Producto:    'Producto
                                            strProducto = Aux_cod_Estr
                                        Case Else
                                        End Select
                                    
                                    
                                        rs_Desglose.MoveNext
                                    Loop
                                    fSAP.writeline StrLegajo & Sep & StrGrupo & Sep & strProducto & Sep & strFecha & Sep & StrHoras & Sep & strLinea
                                Else
                                    'La estructura producto no esta en el desglose ==> uso la estandar y por el total de horas
                                    FLOG.writeline "No Hay estructuras en el desglose ==> uso la estandar y por el total de horas"
                                    If Not Ya_Mostre Then
                                        fSAP.writeline StrLegajo & Sep & StrGrupo & Sep & strProducto & Sep & strFecha & Sep & StrHoras & Sep & strLinea
                                        Ya_Mostre = True
                                    End If
                                End If
                                'Siguiente desglose
                                rs.MoveNext
                            Loop
                        Else
                            'no hubo desglose
                            fSAP.writeline StrLegajo & Sep & StrGrupo & Sep & strProducto & Sep & strFecha & Sep & StrHoras & Sep & strLinea
                        End If
                    End If
            End If
        End If
        
    'Cierro y libero
    If rs_Producto.State = adStateOpen Then rs_Producto.Close
    If rs_Desglose.State = adStateOpen Then rs_Desglose.Close
    If rs.State = adStateOpen Then rs.Close
    
    Set rs_Producto = Nothing
    Set rs = Nothing
    Set rs_Desglose = Nothing
    
Exit Sub
ME_Local:
    FLOG.writeline
    FLOG.writeline "Error Exportando datos."
    FLOG.writeline "Error: " & Err.Description
    FLOG.writeline "Ultimo SQL: " & StrSql
    FLOG.writeline

    'Cierro y libero
    If rs_Producto.State = adStateOpen Then rs_Producto.Close
    If rs_Desglose.State = adStateOpen Then rs_Desglose.Close
    If rs.State = adStateOpen Then rs.Close
    Set rs_Producto = Nothing
    Set rs = Nothing
    Set rs_Desglose = Nothing
End Sub


Public Sub Exportar_Productividad_conDesglose_old(ByVal NroTer As Long, ByVal Fecha As Date)
' Ultima Modificacion: FGZ
'               Fecha: 25/11/2004
' Descripcion:       : Se agregó el producto en la linea de exportacion
'                       Antes exportaba: Legajo & Grupo & Fecha & Horas
'                       Ahora exporta: Legajo & Grupo & Producto & Fecha & Horas
' Ultima Modificacion: FGZ
'               Fecha: 18/11/2004
' Descripcion:       :

Dim StrHoras As String
Dim strProducto As String
Dim strLinea As String
Dim TipoHora As Integer
Dim Ya_Mostre As Boolean

Dim rsSuma As New ADODB.Recordset
Dim rs_Producto As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim rs_Desglose As New ADODB.Recordset

Set objBTurno.Conexion = objConn
   
   
    On Error GoTo ME_Local
    
        objBTurno.Buscar_Turno Fecha, NroTer, False
        If objBTurno.tiene_turno Then
                
                
                'Busco la estructura producto
                StrSql = "SELECT estructura.estrcodext FROM estructura "
                StrSql = StrSql & " INNER JOIN his_estructura ON estructura.estrnro = his_estructura.estrnro"
                StrSql = StrSql & " WHERE estructura.tenro = " & Estructura_Producto
                StrSql = StrSql & " AND his_estructura.ternro = " & NroTer
                StrSql = StrSql & " AND his_estructura.htetdesde <= " & ConvFecha(Fecha)
                StrSql = StrSql & " AND (his_estructura.htethasta >= " & ConvFecha(Fecha) & " OR htethasta IS NULL)"
                If rs_Producto.State = adStateOpen Then rs_Producto.Close
                OpenRecordset StrSql, rs_Producto
                If Not rs_Producto.EOF Then
                    strProducto = Left(rs_Producto!estrcodext, 4)
                Else
                    strProducto = "ERRO"
                End If
                
                'Busco la estructura Linea (12)
                StrSql = "SELECT estructura.estrcodext FROM estructura "
                StrSql = StrSql & " INNER JOIN his_estructura ON estructura.estrnro = his_estructura.estrnro"
                StrSql = StrSql & " WHERE estructura.tenro = 12"
                StrSql = StrSql & " AND his_estructura.ternro = " & NroTer
                StrSql = StrSql & " AND his_estructura.htetdesde <= " & ConvFecha(Fecha)
                StrSql = StrSql & " AND (his_estructura.htethasta >= " & ConvFecha(Fecha) & " OR htethasta IS NULL)"
                If rs_Producto.State = adStateOpen Then rs_Producto.Close
                OpenRecordset StrSql, rs_Producto
                If Not rs_Producto.EOF Then
                    strLinea = Left(rs_Producto!estrcodext, 4)
                Else
                    strLinea = "ERRO"
                End If
                
                
                StrSql = " SELECT thnro, SUM(adcanthoras) as SumaHoras FROM gti_acumdiario "
                StrSql = StrSql & " WHERE adfecha = " & ConvFecha(Fecha) & " AND thnro IN (" & Lista_Thnro & ") AND ternro = " & NroTer
                StrSql = StrSql & " GROUP BY thnro"
                OpenRecordset StrSql, rsSuma
                If Not rsSuma.EOF Then
                    If Not IsNull(rsSuma!sumahoras) Then
                        StrHoras = Replace(Format(rsSuma!sumahoras, "00.00"), ".", ",")
                        
                        'Debo buscar si tenia desglose y por la estructura producto ==>
                        StrSql = "SELECT * FROM gti_achdiario "
                        StrSql = StrSql & " WHERE ternro = " & NroTer
                        StrSql = StrSql & " AND achdfecha = " & ConvFecha(Fecha)
                        StrSql = StrSql & " AND thnro = " & rsSuma!thnro
                        If rs.State = adStateOpen Then rs.Close
                        OpenRecordset StrSql, rs
                        If Not rs.EOF Then
                            Ya_Mostre = False
                            Do While Not rs.EOF
                                'Hubo desglose
                                StrSql = "SELECT * FROM gti_achdiario_estr "
                                StrSql = StrSql & " WHERE achdnro = " & rs!achdnro
                                StrSql = StrSql & " AND tenro = " & Estructura_Producto
                                StrSql = StrSql & " AND achdfecha = " & ConvFecha(Fecha)
                                If rs_Desglose.State = adStateOpen Then rs_Desglose.Close
                                OpenRecordset StrSql, rs_Desglose
                                If Not rs_Desglose.EOF Then
                                    Do While Not rs_Desglose.EOF
                                        StrHoras = Replace(Format(rs!achdcanthoras, "00.00"), ".", ",")
                                    
                                        'Busco el producto del desglose
                                        StrSql = "SELECT estructura.estrcodext FROM estructura "
                                        StrSql = StrSql & " WHERE estructura.estrnro = " & rs_Desglose!estrnro
                                        If rs_Producto.State = adStateOpen Then rs_Producto.Close
                                        OpenRecordset StrSql, rs_Producto
                                        If Not rs_Producto.EOF Then
                                            strProducto = Left(rs_Producto!estrcodext, 4)
                                        Else
                                            strProducto = "ERRO"
                                        End If
                                    
                                        fSAP.writeline StrLegajo & StrGrupo & strProducto & strFecha & StrHoras & strLinea
                                    
                                        rs_Desglose.MoveNext
                                    Loop
                                Else
                                    'La estructura producto no esta en el desglose ==> uso la estandar y por el total de horas
                                    FLOG.writeline "La estructura producto no esta en el desglose ==> uso la estandar y por el total de horas"
                                    If Not Ya_Mostre Then
                                        fSAP.writeline StrLegajo & StrGrupo & strProducto & strFecha & StrHoras & strLinea
                                        Ya_Mostre = True
                                    End If
                                End If
                                'Siguiente desglose
                                rs.MoveNext
                            Loop
                        Else
                            'no hubo desglose
                            fSAP.writeline StrLegajo & StrGrupo & strProducto & strFecha & StrHoras & strLinea
                        End If
                    End If
            End If
        End If
        
    'Cierro y libero
    If rs_Producto.State = adStateOpen Then rs_Producto.Close
    If rs_Desglose.State = adStateOpen Then rs_Desglose.Close
    If rs.State = adStateOpen Then rs.Close
    
    Set rs_Producto = Nothing
    Set rs = Nothing
    Set rs_Desglose = Nothing
    
Exit Sub
ME_Local:
    FLOG.writeline
    FLOG.writeline "Error Exportando datos."
    FLOG.writeline "Error: " & Err.Description
    FLOG.writeline "Ultimo SQL: " & StrSql
    FLOG.writeline

    'Cierro y libero
    If rs_Producto.State = adStateOpen Then rs_Producto.Close
    If rs_Desglose.State = adStateOpen Then rs_Desglose.Close
    If rs.State = adStateOpen Then rs.Close
    Set rs_Producto = Nothing
    Set rs = Nothing
    Set rs_Desglose = Nothing
End Sub


Private Function BuscarTipoHora(NroTurno As Long) As Long
Dim objRs As New ADODB.Recordset

    StrSql = "SELECT * FROM gti_config_tur_hor WHERE turnro = " & NroTurno & " AND gti_config_tur_hor.conhornro = 1 "
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        BuscarTipoHora = objRs!thnro
    Else
        BuscarTipoHora = -1
    End If
    objRs.Close
    Set objRs = Nothing

End Function

Public Function EsNulo(ByVal Objeto) As Boolean
    If IsNull(Objeto) Then
        EsNulo = True
    Else
        If UCase(Objeto) = "NULL" Or UCase(Objeto) = "" Then
            EsNulo = True
        Else
            EsNulo = False
        End If
    End If
End Function

