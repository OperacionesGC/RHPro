Attribute VB_Name = "casamatriz"
' __________________________________________________________________________________________________
' Descripcion: Reporte custom Santillana - Salida para enviar a casa matriz
' Autor      : Gustavo Ring
' Fecha      : 03/12/2008
' Ultima Mod :
' Descripcion:
' ___________________________________________________________________________________________________

Option Explicit


'Global Const Version = "1.00"
'Global Const FechaModificacion = " 03-12-2008 "
'Global Const UltimaModificacion = " " 'Version Inicial Gustavo Ring

'Global Const Version = "1.01"
'Global Const FechaModificacion = " 22-12-2008 "
'Global Const UltimaModificacion = " " 'FGZ

'Global Const Version = "1.02"
'Global Const FechaModificacion = " 02-01-2009 "
'Global Const UltimaModificacion = " " 'Gustavo Ring

'Global Const Version = "1.03"
'Global Const FechaModificacion = " 19-01-2009 "
'Global Const UltimaModificacion = " " 'Gustavo Ring

'Global Const Version = "1.04"
'Global Const FechaModificacion = " 21-01-2009 "
'Global Const UltimaModificacion = " " 'Gustavo Ring


'Global Const Version = "1.05"
'Global Const FechaModificacion = " 31/03/2009 "
'Global Const UltimaModificacion = " " 'FGZ
''                       Encriptacion de string de conexion
''                       correcciones varias

'Global Const Version = "1.06"
'Global Const FechaModificacion = "01/04/2009 "
'Global Const UltimaModificacion = " " 'FGZ
'                       correcciones varias

Global Const Version = "1.07"
Global Const FechaModificacion = "27/05/2009 "
Global Const UltimaModificacion = " " 'Gustavo Ring
'                       Se discriminan los contratos eventuales

' ________________________________________________________________________________________
Global NroProceso As Long
Global Path As String
Global HuboErrores As Boolean


Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset


'NUEVAS
Global EmpErrores As Boolean  ' VERRRRRRRRR

Global filtro As String  ' filtro trae si el empleado es activo o no, y legajo desde -hasta
Global tenro1 As Integer
Global estrnro1 As Integer
Global tenro2 As Integer
Global estrnro2 As Integer
Global tenro3 As Integer
Global estrnro3 As Integer
Global agencia As Integer
Global fecestr As Date
Global repnro As Integer
Global empleados As String
Global Empresas_Filtradas As String
Global contratos_eventuales As String

Dim monnro As Integer
Dim Monto As Double
Dim cargo As Integer
Dim lista As String
Dim Anio As Integer

Dim E1(100) As String
Dim IndiceE1 As Integer
Dim E2(100) As String
Dim IndiceE2 As Integer
Dim E3(100) As String
Dim IndiceE3 As Integer
Dim E5(100) As String
Dim indiceE5 As Integer
Dim E4 As String
Dim AC1 As Long
Dim AC2 As Long

Dim IdUser As String
Dim bpfecha As Date
Dim bphora As String

' Global fecestr As String
'Global TituloRep As String
Global HayDetalleLiq(1 To 12) As Double


Private Sub Main()
Dim strCmdLine As String
Dim Nombre_Arch As String

Dim StrSql As String

Dim PID As String
Dim Parametros As String
Dim ArrParametros
Dim totalEmpleados
Dim cantRegistros

Dim Desde As Date
Dim Hasta As Date
Dim fecestrAnt As Date


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


    TiempoInicialProceso = GetTickCount
    HuboErrores = False
    
    Nombre_Arch = PathFLog & "ReporteCasaMatriz" & "-" & NroProceso & ".log"
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)


    'Abro la conexion
    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If
    OpenConnection strconexion, objconnProgreso
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If
    On Error GoTo 0
    
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
    Flog.writeline
    
    Flog.writeline "Inicio Proceso de Reporte a Casa Matriz : " & Now
    Flog.writeline "Cambio el estado del proceso a Procesando"
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & ", bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline Espacios(Tabulador * 0) & "Obtengo los datos del proceso"
    
    TiempoAcumulado = GetTickCount
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, rs
    
    If Not rs.EOF Then
       'Obtengo los parametros del proceso
        IdUser = rs!IdUser
        bpfecha = rs!bprcfecha
        bphora = rs!bprchora
        Parametros = rs!bprcparam
        
        ArrParametros = Split(Parametros, "@")
             
        Call levantarParametros(ArrParametros)
        
          
        ' Se levantan los datos configurables
        
        Call CargarConfiguracionReporte
              
        cantRegistros = 0
        StrSql = "UPDATE batch_proceso SET bprcprogreso = 0 "
        StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'"
        StrSql = StrSql & ", bprcempleados ='" & CStr(cantRegistros) & "'"
        StrSql = StrSql & " WHERE bpronro = " & NroProceso
        objConn.Execute StrSql, , adExecuteNoRecords
        
        
        'Borrar los datos por si se reprocesa
        'Detalle
        StrSql = "DELETE FROM casa_matriz_det "
        StrSql = StrSql & " WHERE bpronro = " & NroProceso
        objConn.Execute StrSql, , adExecuteNoRecords
        
        
        'Cabecera
        StrSql = "DELETE FROM casa_matriz "
        StrSql = StrSql & " WHERE bpronro = " & NroProceso
        objConn.Execute StrSql, , adExecuteNoRecords
        
        
        ' _____________________________________________________
        ' armar consulta Ppal según Filtro - empls con estructuras activas a la Fecha
        ' ____________________________________________________
        
        Call filtro_empleados(StrSql, fecestr)
        OpenRecordset StrSql, rs1
       
        'seteo de las variables de progreso
        Progreso = 0
        cantRegistros = IndiceE1 + IndiceE2 + IndiceE3
        totalEmpleados = cantRegistros
           
        If cantRegistros = 0 Then
           cantRegistros = 1
           Flog.writeline Espacios(Tabulador * 0) & "No se encontraron empleados para el Filtro."
        End If
        IncPorc = (100 / cantRegistros)
          
        ' Se inicia el proceso
            
        actualizar_progreso (5)
        Call InsertarDatosCab
    
    Else
        Flog.writeline Espacios(Tabulador * 0) & "No se encontraron los datos del proceso nro " & NroProceso
    End If
    
          
    'If rs_Modelo.State = adStateOpen Then rs_Modelo.Close
    'Set rs_Modelo = Nothing

    'Actualizo el estado del proceso
    If Not HuboErrores Then
        StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
        Flog.writeline Espacios(Tabulador * 0) & "Proceso Finalizado Correctamente"
    Else
        StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProceso
        Flog.writeline Espacios(Tabulador * 0) & "Proceso Finalizado Incompleto"
    End If
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    'Flog.writeline "cant open " & Cantidad_de_OpenRecordset

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

' ___________________________________________________________________________________________________
' Procedimiento que inserta la cabecera del reporte.
' Sirve para obtener los datos del proceso
' ___________________________________________________________________________________________________

Sub InsertarDatosCab()

Dim campos As String
Dim valores As String
Dim mondesabr As String
Dim rsdet As New ADODB.Recordset
Dim StrSql As String

On Error GoTo MError


Flog.writeline " "
Flog.writeline Espacios(Tabulador * 1) & "Buscando detalles de la moneda "

' estructura
    StrSql = "SELECT mondesabr FROM moneda "
    StrSql = StrSql & " WHERE monnro = " & monnro
    OpenRecordset StrSql, rsdet
  
  
 If Not rsdet.EOF Then
        mondesabr = rsdet!mondesabr
 Else
        mondesabr = " "
 End If
  
  
Flog.writeline " "
Flog.writeline Espacios(Tabulador * 1) & "Insertar datos de la cabecera  "


campos = " (bpronro,Fecha,Hora,tenro1,estrnro1,tenro2, estrnro2, tenro3, estrnro3, fecestr,anio,moneda,mondesabr)"

valores = "("
valores = valores & NroProceso & "," & ConvFecha(bpfecha) & ",'" & bphora & "',"
valores = valores & tenro1 & "," & estrnro1 & "," & tenro2 & "," & estrnro2 & "," & tenro3 & "," & estrnro3 & ","
valores = valores & ConvFecha(fecestr) & "," & Anio & "," & Monto & ",'" & mondesabr & "')"


StrSql = " INSERT INTO casa_matriz " & campos & " VALUES " & valores
objConn.Execute StrSql, , adExecuteNoRecords

Call InsertarDatosdet

Exit Sub

MError:
    HuboErrores = True
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQL: " & StrSql
End Sub


' ___________________________________________________________________________________________________
' procedimiento que inserta los detalles del acumulador por estructura
' ___________________________________________________________________________________________________

Sub InsertarDatosdet()

Dim campos As String
Dim valores As String
Dim fechadesde As Date
Dim fechahasta As Date
Dim estr1 As Integer
Dim estr2 As Integer
Dim estr3 As Integer
Dim indice As Integer
Dim rsdet As New ADODB.Recordset
Dim montoAcumulador As Double
Dim estrdabr As String
Dim orden As Integer
On Error GoTo MError


campos = " (bpronro,tenro,estrnro,mes,estrdabr,monto,tipest,orden)"
    
orden = 0

' - Seccion 1 Estructura 1

For estr1 = 0 To IndiceE1 - 1
    Flog.writeline " "
    Flog.writeline Espacios(Tabulador * 1) & "Insertar datos del mes de la estructura " & E1(estr1)
  
    'estructura
    StrSql = "SELECT estrdabr FROM estructura "
    StrSql = StrSql & " WHERE estructura.estrnro = " & E1(estr1)
    OpenRecordset StrSql, rsdet
  
    If Not rsdet.EOF Then
        estrdabr = rsdet!estrdabr
    Else
        estrdabr = " "
    End If
    
    For indice = 1 To 12
        fechahasta = ultimo_dia_mes(indice, Anio)
        fechadesde = primer_dia_mes(indice, Anio)
        
        
        Flog.writeline " "
        Flog.writeline Espacios(Tabulador * 1) & "Insertar datos del mes: " & indice
    
'        StrSql = " SELECT proceso.pronro,acu_liq.acunro,acudesabr, almonto, alcant, estructura.estrdabr "
'        StrSql = StrSql & " FROM empleado "
'        StrSql = StrSql & " INNER JOIN cabliq ON cabliq.empleado = empleado.ternro "
'        StrSql = StrSql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
'        StrSql = StrSql & " INNER JOIN acumulador ON acumulador.acunro = acu_liq.acunro "
'        StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro "
'        StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro "
'        StrSql = StrSql & " AND periodo.pliqdesde >=" & ConvFecha(fechadesde) & " AND periodo.pliqhasta <= " & ConvFecha(fechahasta)
'        StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.ternro = empleado.ternro "
'        StrSql = StrSql & " AND his_estructura.htetdesde <= " & ConvFecha(fechahasta) & " AND ((his_estructura.htethasta >=" & ConvFecha(fechahasta) & ") OR his_estructura.htethasta is null)"
'        StrSql = StrSql & " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro"
'        StrSql = StrSql & " WHERE acumulador.acunro= " & AC1
'        StrSql = StrSql & " AND estructura.estrnro = " & E1(estr1)
'        StrSql = StrSql & " AND empleado.ternro IN " & empleados
        
        
        
        StrSql = " SELECT distinct empleado.empleg, empleado.ternro, ammonto monto"
        StrSql = StrSql & " FROM empleado "
        StrSql = StrSql & " INNER JOIN fases ON empleado.ternro = fases.empleado"
        StrSql = StrSql & " INNER JOIN his_estructura as emp ON emp.ternro = empleado.ternro AND emp.tenro = 10 "
        StrSql = StrSql & " INNER JOIN his_estructura as he1 ON he1.ternro = empleado.ternro "
        StrSql = StrSql & " AND he1.htetdesde <= " & ConvFecha(fechahasta) & " AND ((he1.htethasta >=" & ConvFecha(fechahasta) & ") OR he1.htethasta is null)"
        'StrSql = StrSql & " LEFT JOIN his_estructura as he2 ON he2.ternro = empleado.ternro AND he2.tenro = 21 "
        'StrSql = StrSql & " AND he2.htetdesde <= " & ConvFecha(fechahasta) & " AND ((he2.htethasta >=" & ConvFecha(fechahasta) & ") OR he2.htethasta is null)"
        StrSql = StrSql & " INNER JOIN acu_mes ON acu_mes.ternro = empleado.ternro "
        StrSql = StrSql & " WHERE acunro= " & AC1
        StrSql = StrSql & " AND ammes = " & Month(fechahasta)
        StrSql = StrSql & " AND amanio = " & Year(fechahasta)
        StrSql = StrSql & " AND he1.estrnro = " & E1(estr1)
        StrSql = StrSql & " AND empleado.ternro IN " & empleados
        StrSql = StrSql & " AND (fases.altfec <= " & ConvFecha(fechahasta) & " AND ( fases.bajfec >= " & ConvFecha(fechahasta) & " OR fases.bajfec IS NULL))"
        If Empresas_Filtradas <> "0" Then
            StrSql = StrSql & " AND emp.estrnro IN (" & Empresas_Filtradas & ")"
            StrSql = StrSql & " AND emp.htetdesde <= " & ConvFecha(fechahasta) & " AND ((emp.htethasta >=" & ConvFecha(fechahasta) & ") OR emp.htethasta is null)"
        End If
        StrSql = StrSql & " AND empleado.ternro NOT IN (SELECT tercero.ternro FROM tercero "
        StrSql = StrSql & " INNER JOIN his_estructura as eventual ON eventual.ternro = tercero.ternro "
        StrSql = StrSql & " AND eventual.tenro = 18 AND eventual.estrnro IN (" & contratos_eventuales & ")"
        StrSql = StrSql & " AND eventual.htetdesde <= " & ConvFecha(fechahasta) & " AND ((eventual.htethasta >=" & ConvFecha(fechahasta) & ") OR eventual.htethasta is null))"

        
        OpenRecordset StrSql, rsdet
        
        orden = orden + 1
        montoAcumulador = 0
        While Not rsdet.EOF
            montoAcumulador = montoAcumulador + rsdet!Monto
            rsdet.MoveNext
        Wend
        ' Divido por el valor de la moneda
        
        montoAcumulador = montoAcumulador / Monto
        
        Flog.writeline Espacios(Tabulador * 1) & "Monto E1: " & montoAcumulador
        valores = "(" & NroProceso & "," & tenro1 & "," & E1(estr1) & "," & indice & ",'" & estrdabr & "'," & montoAcumulador & ",'1'," & orden & ")"
        
        StrSql = " INSERT INTO casa_matriz_det " & campos & " VALUES " & valores
        objConn.Execute StrSql, , adExecuteNoRecords
    Next indice
Next estr1

' Seccion 1 Estructura 2
For estr2 = 0 To IndiceE2 - 1
  Flog.writeline " "
  Flog.writeline Espacios(Tabulador * 1) & "Insertar datos del mes de la estructura " & E2(estr2)
  
  ' estructura
    StrSql = "SELECT estrdabr FROM estructura "
    StrSql = StrSql & " WHERE estructura.estrnro = " & E2(estr2)
    OpenRecordset StrSql, rsdet
  
    If Not rsdet.EOF Then
        estrdabr = rsdet!estrdabr
    Else
        estrdabr = " "
    End If
  
 
    For indice = 1 To 12
        fechahasta = ultimo_dia_mes(indice, Anio)
        fechadesde = primer_dia_mes(indice, Anio)
      
        
        Flog.writeline " "
        Flog.writeline Espacios(Tabulador * 1) & "Insertar datos del mes: " & indice
    
'        StrSql = " SELECT proceso.pronro,acu_liq.acunro,acudesabr, almonto, alcant, estructura.estrdabr "
'        StrSql = StrSql & " FROM empleado "
'        StrSql = StrSql & " INNER JOIN cabliq ON cabliq.empleado = empleado.ternro "
'        StrSql = StrSql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
'        StrSql = StrSql & " INNER JOIN acumulador ON acumulador.acunro = acu_liq.acunro "
'        StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro "
'        StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro "
'        StrSql = StrSql & " AND periodo.pliqdesde >=" & ConvFecha(fechadesde) & " AND periodo.pliqhasta <= " & ConvFecha(fechahasta)
'        StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.ternro = empleado.ternro "
'        StrSql = StrSql & " AND his_estructura.htetdesde <= " & ConvFecha(fechahasta) & " AND ((his_estructura.htethasta >=" & ConvFecha(fechahasta) & ") OR his_estructura.htethasta is null)"
'        StrSql = StrSql & " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro"
'        StrSql = StrSql & " WHERE acumulador.acunro= " & AC1
'        StrSql = StrSql & " AND estructura.estrnro = " & E2(estr2)
'        StrSql = StrSql & " AND empleado.ternro IN " & empleados
        
        StrSql = " SELECT distinct empleado.empleg, empleado.ternro, ammonto monto"
        StrSql = StrSql & " FROM empleado "
        StrSql = StrSql & " INNER JOIN fases ON empleado.ternro = fases.empleado"
        StrSql = StrSql & " INNER JOIN his_estructura as emp ON emp.ternro = empleado.ternro AND emp.tenro = 10 "
        StrSql = StrSql & " INNER JOIN his_estructura as he1 ON he1.ternro = empleado.ternro "
        StrSql = StrSql & " AND he1.htetdesde <= " & ConvFecha(fechahasta) & " AND ((he1.htethasta >=" & ConvFecha(fechahasta) & ") OR he1.htethasta is null)"
        'StrSql = StrSql & " LEFT JOIN his_estructura as he2 ON he2.ternro = empleado.ternro AND he2.tenro = 21 "
        'StrSql = StrSql & " AND he2.htetdesde <= " & ConvFecha(fechahasta) & " AND ((he2.htethasta >=" & ConvFecha(fechahasta) & ") OR he2.htethasta is null)"
        StrSql = StrSql & " INNER JOIN acu_mes ON acu_mes.ternro = empleado.ternro "
        StrSql = StrSql & " WHERE acunro= " & AC1
        StrSql = StrSql & " AND ammes = " & Month(fechahasta)
        StrSql = StrSql & " AND amanio = " & Year(fechahasta)
        StrSql = StrSql & " AND he1.estrnro = " & E2(estr2)
        StrSql = StrSql & " AND empleado.ternro IN " & empleados
        StrSql = StrSql & " AND (fases.altfec <= " & ConvFecha(fechahasta) & " AND ( fases.bajfec >= " & ConvFecha(fechahasta) & " OR fases.bajfec IS NULL))"
        If Empresas_Filtradas <> "0" Then
            StrSql = StrSql & " AND emp.estrnro IN (" & Empresas_Filtradas & ")"
            StrSql = StrSql & " AND emp.htetdesde <= " & ConvFecha(fechahasta) & " AND ((emp.htethasta >=" & ConvFecha(fechahasta) & ") OR emp.htethasta is null)"
        End If
        StrSql = StrSql & " AND empleado.ternro NOT IN (SELECT tercero.ternro FROM tercero "
        StrSql = StrSql & " INNER JOIN his_estructura as eventual ON eventual.ternro = tercero.ternro "
        StrSql = StrSql & " AND eventual.tenro = 18 AND eventual.estrnro IN (" & contratos_eventuales & ")"
        StrSql = StrSql & " AND eventual.htetdesde <= " & ConvFecha(fechahasta) & " AND ((eventual.htethasta >=" & ConvFecha(fechahasta) & ") OR eventual.htethasta is null))"

        
        OpenRecordset StrSql, rsdet
        
        orden = orden + 1
        montoAcumulador = 0
        While Not rsdet.EOF
            montoAcumulador = montoAcumulador + rsdet!Monto
            rsdet.MoveNext
        Wend
        'Divido por el valor de la moneda
        montoAcumulador = montoAcumulador / Monto
        Flog.writeline Espacios(Tabulador * 1) & "Monto E2: " & montoAcumulador
    
        valores = "(" & NroProceso & "," & tenro1 & "," & E2(estr2) & "," & indice & ",'" & estrdabr & "'," & montoAcumulador & ",'2'," & orden & ")"
        
        StrSql = " INSERT INTO casa_matriz_det " & campos & " VALUES " & valores
        objConn.Execute StrSql, , adExecuteNoRecords
    Next indice
Next estr2

'Sección 1 Estructura 3
For estr3 = 0 To IndiceE3 - 1
    Flog.writeline " "
    Flog.writeline Espacios(Tabulador * 1) & "Insertar datos del mes de la estructura " & E3(estr3)
  
    'estructura
    StrSql = "SELECT estrdabr FROM estructura "
    StrSql = StrSql & " WHERE estructura.estrnro = " & E3(estr3)
    OpenRecordset StrSql, rsdet
    If Not rsdet.EOF Then
        estrdabr = rsdet!estrdabr
    Else
        estrdabr = " "
    End If
    
    
    
    For indice = 1 To 12
        fechahasta = ultimo_dia_mes(indice, Anio)
        fechadesde = primer_dia_mes(indice, Anio)
        
        Flog.writeline " "
        Flog.writeline Espacios(Tabulador * 1) & "Insertar datos del mes: " & indice
    
'        StrSql = " SELECT proceso.pronro,acu_liq.acunro,acudesabr, almonto, alcant, estructura.estrdabr "
'        StrSql = StrSql & " FROM empleado "
'        StrSql = StrSql & " INNER JOIN cabliq ON cabliq.empleado = empleado.ternro "
'        StrSql = StrSql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
'        StrSql = StrSql & " INNER JOIN acumulador ON acumulador.acunro = acu_liq.acunro "
'        StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro "
'        StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro "
'        StrSql = StrSql & " AND periodo.pliqdesde >=" & ConvFecha(fechadesde) & " AND periodo.pliqhasta <= " & ConvFecha(fechahasta)
'        StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.ternro = empleado.ternro "
'        StrSql = StrSql & " AND his_estructura.htetdesde <= " & ConvFecha(fechahasta) & " AND ((his_estructura.htethasta >=" & ConvFecha(fechahasta) & ") OR his_estructura.htethasta is null)"
'        StrSql = StrSql & " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro"
'        StrSql = StrSql & " WHERE acumulador.acunro= " & AC1
'        StrSql = StrSql & " AND estructura.estrnro = " & E3(estr3)
'        StrSql = StrSql & " AND empleado.ternro IN " & empleados
        
        StrSql = " SELECT distinct empleado.empleg, empleado.ternro, ammonto monto"
        StrSql = StrSql & " FROM empleado "
        StrSql = StrSql & " INNER JOIN fases ON empleado.ternro = fases.empleado"
        StrSql = StrSql & " INNER JOIN his_estructura as emp ON emp.ternro = empleado.ternro AND emp.tenro = 10 "
        StrSql = StrSql & " INNER JOIN his_estructura as he1 ON he1.ternro = empleado.ternro "
        StrSql = StrSql & " AND he1.htetdesde <= " & ConvFecha(fechahasta) & " AND ((he1.htethasta >=" & ConvFecha(fechahasta) & ") OR he1.htethasta is null)"
        'StrSql = StrSql & " LEFT JOIN his_estructura as he2 ON he2.ternro = empleado.ternro AND he2.tenro = 21 "
        'StrSql = StrSql & " AND he2.htetdesde <= " & ConvFecha(fechahasta) & " AND ((he2.htethasta >=" & ConvFecha(fechahasta) & ") OR he2.htethasta is null)"
        StrSql = StrSql & " INNER JOIN acu_mes ON acu_mes.ternro = empleado.ternro "
        StrSql = StrSql & " WHERE acunro= " & AC1
        StrSql = StrSql & " AND ammes = " & Month(fechahasta)
        StrSql = StrSql & " AND amanio = " & Year(fechahasta)
        StrSql = StrSql & " AND he1.estrnro = " & E3(estr3)
        StrSql = StrSql & " AND empleado.ternro IN " & empleados
        StrSql = StrSql & " AND (fases.altfec <= " & ConvFecha(fechahasta) & " AND ( fases.bajfec >= " & ConvFecha(fechahasta) & " OR fases.bajfec IS NULL))"
        If Empresas_Filtradas <> "0" Then
            StrSql = StrSql & " AND emp.estrnro IN (" & Empresas_Filtradas & ")"
            StrSql = StrSql & " AND emp.htetdesde <= " & ConvFecha(fechahasta) & " AND ((emp.htethasta >=" & ConvFecha(fechahasta) & ") OR emp.htethasta is null)"
        End If
        StrSql = StrSql & " AND empleado.ternro NOT IN (SELECT tercero.ternro FROM tercero "
        StrSql = StrSql & " INNER JOIN his_estructura as eventual ON eventual.ternro = tercero.ternro "
        StrSql = StrSql & " AND eventual.tenro = 18 AND eventual.estrnro IN (" & contratos_eventuales & ")"
        StrSql = StrSql & " AND eventual.htetdesde <= " & ConvFecha(fechahasta) & " AND ((eventual.htethasta >=" & ConvFecha(fechahasta) & ") OR eventual.htethasta is null))"
        
        
        OpenRecordset StrSql, rsdet
        
        orden = orden + 1
        montoAcumulador = 0
        While Not rsdet.EOF
            montoAcumulador = montoAcumulador + rsdet!Monto
            rsdet.MoveNext
        Wend
        
        'Divido por el valor de la moneda
        montoAcumulador = montoAcumulador / Monto
        Flog.writeline Espacios(Tabulador * 1) & "Monto E2: " & montoAcumulador
    
        valores = "(" & NroProceso & "," & tenro1 & "," & E3(estr3) & "," & indice & ",'" & estrdabr & "'," & montoAcumulador & ",'3'," & orden & ")"
                    
        StrSql = " INSERT INTO casa_matriz_det " & campos & " VALUES " & valores
        objConn.Execute StrSql, , adExecuteNoRecords
    Next indice
Next estr3

actualizar_progreso (20)
Call InsertarDatosdetPlantilla
actualizar_progreso (60)
Call totales_seccion
actualizar_progreso (70)
Call calcular_plantilla_activa
actualizar_progreso (80)
Call calcular_tipo_contrato
actualizar_progreso (85)
Call altas_bajas
actualizar_progreso (90)
Call alta_contrato_eventual
actualizar_progreso (93)
Call eventuales
actualizar_progreso (100)
Exit Sub

MError:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql

End Sub


Sub InsertarDatosdet_old()

Dim campos As String
Dim valores As String
Dim fechadesde As Date
Dim fechahasta As Date
Dim estr1 As Integer
Dim estr2 As Integer
Dim estr3 As Integer
Dim indice As Integer
Dim rsdet As New ADODB.Recordset
Dim montoAcumulador As Double
Dim estrdabr As String
Dim orden As Integer
On Error GoTo MError


campos = " (bpronro,tenro,estrnro,mes,estrdabr,monto,tipest,orden)"
    
orden = 0

' - Seccion 1 Estructura 1

For estr1 = 0 To IndiceE1 - 1
    Flog.writeline " "
    Flog.writeline Espacios(Tabulador * 1) & "Insertar datos del mes de la estructura " & E1(estr1)
  
    'estructura
    StrSql = "SELECT estrdabr FROM estructura "
    StrSql = StrSql & " WHERE estructura.estrnro = " & E1(estr1)
    OpenRecordset StrSql, rsdet
  
    If Not rsdet.EOF Then
        estrdabr = rsdet!estrdabr
    Else
        estrdabr = " "
    End If
    
    For indice = 1 To 12
        fechahasta = ultimo_dia_mes(indice, Anio)
        fechadesde = primer_dia_mes(indice, Anio)
        
        
        Flog.writeline " "
        Flog.writeline Espacios(Tabulador * 1) & "Insertar datos del mes: " & indice
    
        StrSql = " SELECT proceso.pronro,acu_liq.acunro,acudesabr, almonto, alcant, estructura.estrdabr "
        StrSql = StrSql & " FROM empleado "
        StrSql = StrSql & " INNER JOIN cabliq ON cabliq.empleado = empleado.ternro "
        StrSql = StrSql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
        StrSql = StrSql & " INNER JOIN acumulador ON acumulador.acunro = acu_liq.acunro "
        StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro "
        StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro "
        StrSql = StrSql & " AND periodo.pliqdesde >=" & ConvFecha(fechadesde) & " AND periodo.pliqhasta <= " & ConvFecha(fechahasta)
        StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.ternro = empleado.ternro "
        StrSql = StrSql & " AND his_estructura.htetdesde <= " & ConvFecha(fechahasta) & " AND ((his_estructura.htethasta >=" & ConvFecha(fechahasta) & ") OR his_estructura.htethasta is null)"
        StrSql = StrSql & " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro"
        StrSql = StrSql & " WHERE acumulador.acunro= " & AC1
        StrSql = StrSql & " AND estructura.estrnro = " & E1(estr1)
        StrSql = StrSql & " AND empleado.ternro IN " & empleados
        OpenRecordset StrSql, rsdet
        
        orden = orden + 1
        
        
        montoAcumulador = 0
        While Not rsdet.EOF
            montoAcumulador = montoAcumulador + rsdet!almonto
            rsdet.MoveNext
        Wend
        ' Divido por el valor de la moneda
        
        montoAcumulador = montoAcumulador / Monto
        
        Flog.writeline Espacios(Tabulador * 1) & "Monto E1: " & montoAcumulador
    
        valores = "(" & NroProceso & "," & tenro1 & "," & E1(estr1) & "," & indice & ",'" & estrdabr & "'," & montoAcumulador & ",'1'," & orden & ")"
        
        
        StrSql = " INSERT INTO casa_matriz_det " & campos & " VALUES " & valores
        objConn.Execute StrSql, , adExecuteNoRecords
    Next indice
Next estr1

' Seccion 1 Estructura 2

For estr2 = 0 To IndiceE2 - 1
  Flog.writeline " "
  Flog.writeline Espacios(Tabulador * 1) & "Insertar datos del mes de la estructura " & E2(estr2)
  
  ' estructura
    StrSql = "SELECT estrdabr FROM estructura "
    StrSql = StrSql & " WHERE estructura.estrnro = " & E2(estr2)
    OpenRecordset StrSql, rsdet
  
    If Not rsdet.EOF Then
        estrdabr = rsdet!estrdabr
    Else
        estrdabr = " "
    End If
  
 
    For indice = 1 To 12
  
    
    fechahasta = ultimo_dia_mes(indice, Anio)
    fechadesde = primer_dia_mes(indice, Anio)
  
    
    
    Flog.writeline " "
    Flog.writeline Espacios(Tabulador * 1) & "Insertar datos del mes: " & indice

    StrSql = " SELECT proceso.pronro,acu_liq.acunro,acudesabr, almonto, alcant, estructura.estrdabr "
    StrSql = StrSql & " FROM empleado "
    StrSql = StrSql & " INNER JOIN cabliq ON cabliq.empleado = empleado.ternro "
    StrSql = StrSql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
    StrSql = StrSql & " INNER JOIN acumulador ON acumulador.acunro = acu_liq.acunro "
    StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro "
    StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro "
    StrSql = StrSql & " AND periodo.pliqdesde >=" & ConvFecha(fechadesde) & " AND periodo.pliqhasta <= " & ConvFecha(fechahasta)
    StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.ternro = empleado.ternro "
    StrSql = StrSql & " AND his_estructura.htetdesde <= " & ConvFecha(fechahasta) & " AND ((his_estructura.htethasta >=" & ConvFecha(fechahasta) & ") OR his_estructura.htethasta is null)"
    StrSql = StrSql & " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro"
    StrSql = StrSql & " WHERE acumulador.acunro= " & AC1
    StrSql = StrSql & " AND estructura.estrnro = " & E2(estr2)
    StrSql = StrSql & " AND empleado.ternro IN " & empleados
    OpenRecordset StrSql, rsdet
    
    orden = orden + 1
    
    
    montoAcumulador = 0
    While Not rsdet.EOF
        montoAcumulador = montoAcumulador + rsdet!almonto
        rsdet.MoveNext
    Wend
    ' Divido por el valor de la moneda
    
    montoAcumulador = montoAcumulador / Monto
    
    Flog.writeline Espacios(Tabulador * 1) & "Monto E2: " & montoAcumulador

    valores = "(" & NroProceso & "," & tenro1 & "," & E2(estr2) & "," & indice & ",'" & estrdabr & "'," & montoAcumulador & ",'2'," & orden & ")"
    
    
    
    
    StrSql = " INSERT INTO casa_matriz_det " & campos & " VALUES " & valores
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Next
    
Next

' Sección 1 Estructura 3

For estr3 = 0 To IndiceE3 - 1

  Flog.writeline " "
  Flog.writeline Espacios(Tabulador * 1) & "Insertar datos del mes de la estructura " & E3(estr3)
  
  ' estructura
    StrSql = "SELECT estrdabr FROM estructura "
    StrSql = StrSql & " WHERE estructura.estrnro = " & E3(estr3)
    OpenRecordset StrSql, rsdet
  
    If Not rsdet.EOF Then
        estrdabr = rsdet!estrdabr
    Else
        estrdabr = " "
    End If
    
    
    
    For indice = 1 To 12
  
    
    fechahasta = ultimo_dia_mes(indice, Anio)
    fechadesde = primer_dia_mes(indice, Anio)
  
    
    
    Flog.writeline " "
    Flog.writeline Espacios(Tabulador * 1) & "Insertar datos del mes: " & indice

    StrSql = " SELECT proceso.pronro,acu_liq.acunro,acudesabr, almonto, alcant, estructura.estrdabr "
    StrSql = StrSql & " FROM empleado "
    StrSql = StrSql & " INNER JOIN cabliq ON cabliq.empleado = empleado.ternro "
    StrSql = StrSql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
    StrSql = StrSql & " INNER JOIN acumulador ON acumulador.acunro = acu_liq.acunro "
    StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro "
    StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro "
    StrSql = StrSql & " AND periodo.pliqdesde >=" & ConvFecha(fechadesde) & " AND periodo.pliqhasta <= " & ConvFecha(fechahasta)
    StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.ternro = empleado.ternro "
    StrSql = StrSql & " AND his_estructura.htetdesde <= " & ConvFecha(fechahasta) & " AND ((his_estructura.htethasta >=" & ConvFecha(fechahasta) & ") OR his_estructura.htethasta is null)"
    StrSql = StrSql & " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro"
    StrSql = StrSql & " WHERE acumulador.acunro= " & AC1
    StrSql = StrSql & " AND estructura.estrnro = " & E3(estr3)
    StrSql = StrSql & " AND empleado.ternro IN " & empleados
    OpenRecordset StrSql, rsdet
    
    orden = orden + 1
    
    
    montoAcumulador = 0
    While Not rsdet.EOF
        montoAcumulador = montoAcumulador + rsdet!almonto
        rsdet.MoveNext
    Wend
    ' Divido por el valor de la moneda
    
    montoAcumulador = montoAcumulador / Monto
    
    Flog.writeline Espacios(Tabulador * 1) & "Monto E2: " & montoAcumulador

    valores = "(" & NroProceso & "," & tenro1 & "," & E3(estr3) & "," & indice & ",'" & estrdabr & "'," & montoAcumulador & ",'3'," & orden & ")"
                
    StrSql = " INSERT INTO casa_matriz_det " & campos & " VALUES " & valores
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Next
    
Next
actualizar_progreso (20)
Call InsertarDatosdetPlantilla
actualizar_progreso (60)
Call totales_seccion
actualizar_progreso (70)
Call calcular_plantilla_activa
actualizar_progreso (80)
Call calcular_tipo_contrato
actualizar_progreso (90)
Call altas_bajas
actualizar_progreso (100)

Exit Sub

MError:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql

End Sub


' ___________________________________________________________________________________________________
' procedimiento que inserta los detalles del acumulador por estructura según cantidad horas
' ___________________________________________________________________________________________________

Sub InsertarDatosdetPlantilla()
Dim campos As String
Dim valores As String
Dim fechadesde As Date
Dim fechahasta As Date
Dim estr1 As Integer
Dim estr2 As Integer
Dim estr3 As Integer
Dim indice As Integer
Dim rsdet As New ADODB.Recordset
Dim rsdet1 As New ADODB.Recordset
Dim montoAcumulador As Double
Dim estrdabr As String
Dim orden As Integer
Dim horas As Integer
Dim dias As Integer
'Dim porcentaje As Integer
Dim porcentaje As Double

On Error GoTo MError


campos = " (bpronro,tenro,estrnro,mes,estrdabr,monto,tipest,orden)"
orden = 0

'--------------------------------- Comienzo de la Seccion 2 Estructura 1 ------------------------
Flog.writeline Espacios(Tabulador * 1) & "Sección 2"

For estr1 = 0 To IndiceE1 - 1
  Flog.writeline " "
  Flog.writeline Espacios(Tabulador * 1) & "Insertar datos del mes de la estructura " & E1(estr1)
 
    StrSql = "SELECT estrdabr FROM estructura "
    StrSql = StrSql & " WHERE estructura.estrnro = " & E1(estr1)
    OpenRecordset StrSql, rsdet1
    If Not rsdet1.EOF Then
        estrdabr = rsdet1!estrdabr
    Else
        estrdabr = " "
    End If
    
    For indice = 1 To 12
        fechahasta = ultimo_dia_mes(indice, Anio)
        fechadesde = primer_dia_mes(indice, Anio)
        
        Flog.writeline " "
        Flog.writeline Espacios(Tabulador * 2) & "Insertar datos del mes: " & indice
    
        StrSql = " SELECT distinct empleado.empleg, empleado.ternro, ammonto dias, e2.estrcodext horas"
        StrSql = StrSql & " FROM empleado "
        StrSql = StrSql & " INNER JOIN fases ON empleado.ternro = fases.empleado"
        StrSql = StrSql & " INNER JOIN his_estructura as emp ON emp.ternro = empleado.ternro AND emp.tenro = 10"
        StrSql = StrSql & " INNER JOIN his_estructura as he1 ON he1.ternro = empleado.ternro "
        StrSql = StrSql & " AND he1.htetdesde <= " & ConvFecha(fechahasta) & " AND ((he1.htethasta >=" & ConvFecha(fechahasta) & ") OR he1.htethasta is null)"
        StrSql = StrSql & " LEFT JOIN his_estructura as he2 ON he2.ternro = empleado.ternro AND he2.tenro = 21 "
        StrSql = StrSql & " AND he2.htetdesde <= " & ConvFecha(fechahasta) & " AND ((he2.htethasta >=" & ConvFecha(fechahasta) & ") OR he2.htethasta is null)"
        StrSql = StrSql & " INNER JOIN estructura e2 ON he2.estrnro = e2.estrnro"
        StrSql = StrSql & " INNER JOIN acu_mes ON acu_mes.ternro = empleado.ternro "
        StrSql = StrSql & " WHERE acunro= " & AC2
        StrSql = StrSql & " AND ammes = " & Month(fechahasta)
        StrSql = StrSql & " AND amanio = " & Year(fechahasta)
        StrSql = StrSql & " AND he1.estrnro = " & E1(estr1)
        StrSql = StrSql & " AND empleado.ternro IN " & empleados
        StrSql = StrSql & " AND (fases.altfec <= " & ConvFecha(fechahasta) & " AND ( fases.bajfec >= " & ConvFecha(fechahasta) & " OR fases.bajfec IS NULL))"
        If Empresas_Filtradas <> "0" Then
            StrSql = StrSql & " AND emp.estrnro IN (" & Empresas_Filtradas & ")"
            StrSql = StrSql & " AND emp.htetdesde <= " & ConvFecha(fechahasta) & " AND ((emp.htethasta >=" & ConvFecha(fechahasta) & ") OR emp.htethasta is null)"
        End If
        StrSql = StrSql & " AND empleado.ternro NOT IN (SELECT tercero.ternro FROM tercero "
        StrSql = StrSql & " INNER JOIN his_estructura as eventual ON eventual.ternro = tercero.ternro "
        StrSql = StrSql & " AND eventual.tenro = 18 AND eventual.estrnro IN (" & contratos_eventuales & ")"
        StrSql = StrSql & " AND eventual.htetdesde <= " & ConvFecha(fechahasta) & " AND ((eventual.htethasta >=" & ConvFecha(fechahasta) & ") OR eventual.htethasta is null))"

        
        
        OpenRecordset StrSql, rsdet
        
        orden = orden + 1
        montoAcumulador = 0
        porcentaje = 0
        
        If Not rsdet.EOF Then
            Flog.writeline Espacios(Tabulador * 3) & " Se encontraron " & rsdet.RecordCount & " empleados"
        Else
            Flog.writeline Espacios(Tabulador * 3) & " No se encontraron empleados"
        End If
        Flog.writeline
        While Not rsdet.EOF
            Flog.writeline Espacios(Tabulador * 4) & " empleado: " & rsdet!empleg
            'reviso la cantidad de hs diarias segun el regimen horario
            If EsNulo(rsdet!horas) Then
                horas = 0
                Flog.writeline Espacios(Tabulador * 5) & " No se configuro el regimen horario del empleado:" & rsdet("ternro")
            Else
                horas = rsdet!horas
            End If
            Flog.writeline Espacios(Tabulador * 4) & " horas diarias: " & horas
            
            'Topeo la cantidad de dias
            If rsdet!dias > 30 Then
                dias = 30
            Else
                dias = rsdet!dias
            End If
            Flog.writeline Espacios(Tabulador * 4) & " dias: " & dias
            porcentaje = porcentaje + Round((1 * (dias) / 30) * (horas / 8), 2)
        
            rsdet.MoveNext
        Wend
        
        'Calculo el porcentaje de horas
        Flog.writeline Espacios(Tabulador * 1) & "Porcentaje E1: " & porcentaje
        Flog.writeline Espacios(Tabulador * 1) & "Sql: " & StrSql
        valores = "(" & NroProceso & "," & tenro1 & "," & E1(estr1) & "," & indice & ",'" & estrdabr & "'," & porcentaje & ",7," & orden & ")"
        StrSql = " INSERT INTO casa_matriz_det " & campos & " VALUES " & valores
        objConn.Execute StrSql, , adExecuteNoRecords
    
    Next indice
Next estr1

actualizar_progreso (50)
'------------------------------- Fin sección 2 estructura 1 ----------------------------------

'------------------------------- Comienzo de Sección 2 - Estructura 2 ------------------------
Flog.writeline Espacios(Tabulador * 1) & "Sección 2.. Estructuras 2 - Hombres"
For estr2 = 0 To IndiceE2 - 1
    Flog.writeline " "
    Flog.writeline Espacios(Tabulador * 1) & "Insertar datos del mes de la estructura " & E2(estr2)
 
    StrSql = "SELECT estrdabr FROM estructura "
    StrSql = StrSql & " WHERE estructura.estrnro = " & E2(estr2)
    OpenRecordset StrSql, rsdet1
    If Not rsdet1.EOF Then
        estrdabr = rsdet1!estrdabr
    Else
        estrdabr = " "
    End If
    
    For indice = 1 To 12
        fechahasta = ultimo_dia_mes(indice, Anio)
        fechadesde = primer_dia_mes(indice, Anio)
      
        Flog.writeline " "
        Flog.writeline Espacios(Tabulador * 1) & "Insertar datos del mes: " & indice
    
        StrSql = " SELECT distinct empleado.empleg, empleado.ternro, ammonto dias, e2.estrcodext horas"
        StrSql = StrSql & " FROM empleado "
        StrSql = StrSql & " INNER JOIN tercero ON empleado.ternro = tercero.ternro"
        StrSql = StrSql & " INNER JOIN fases ON empleado.ternro = fases.empleado"
        StrSql = StrSql & " INNER JOIN his_estructura as emp ON emp.ternro = empleado.ternro AND emp.tenro = 10"
        StrSql = StrSql & " INNER JOIN his_estructura as he1 ON he1.ternro = empleado.ternro "
        StrSql = StrSql & " AND he1.htetdesde <= " & ConvFecha(fechahasta) & " AND ((he1.htethasta >=" & ConvFecha(fechahasta) & ") OR he1.htethasta is null)"
        StrSql = StrSql & " LEFT JOIN his_estructura as he2 ON he2.ternro = empleado.ternro AND he2.tenro = 21 "
        StrSql = StrSql & " AND he2.htetdesde <= " & ConvFecha(fechahasta) & " AND ((he2.htethasta >=" & ConvFecha(fechahasta) & ") OR he2.htethasta is null)"
        StrSql = StrSql & " INNER JOIN estructura e2 ON he2.estrnro = e2.estrnro"
        StrSql = StrSql & " INNER JOIN acu_mes ON acu_mes.ternro = empleado.ternro "
        StrSql = StrSql & " WHERE acunro= " & AC2
        StrSql = StrSql & " AND ammes = " & Month(fechahasta)
        StrSql = StrSql & " AND amanio = " & Year(fechahasta)
        StrSql = StrSql & " AND he1.estrnro = " & E2(estr2)
        StrSql = StrSql & " AND empleado.ternro IN " & empleados
        StrSql = StrSql & " AND (fases.altfec <= " & ConvFecha(fechahasta) & " AND ( fases.bajfec >= " & ConvFecha(fechahasta) & " OR fases.bajfec IS NULL))"
        If Empresas_Filtradas <> "0" Then
            StrSql = StrSql & " AND emp.estrnro IN (" & Empresas_Filtradas & ")"
            StrSql = StrSql & " AND emp.htetdesde <= " & ConvFecha(fechahasta) & " AND ((emp.htethasta >=" & ConvFecha(fechahasta) & ") OR emp.htethasta is null)"
        End If
        StrSql = StrSql & " AND tersex = -1 "
        StrSql = StrSql & " AND empleado.ternro NOT IN (SELECT tercero.ternro FROM tercero "
        StrSql = StrSql & " INNER JOIN his_estructura as eventual ON eventual.ternro = tercero.ternro "
        StrSql = StrSql & " AND eventual.tenro = 18 AND eventual.estrnro IN (" & contratos_eventuales & ")"
        StrSql = StrSql & " AND eventual.htetdesde <= " & ConvFecha(fechahasta) & " AND ((eventual.htethasta >=" & ConvFecha(fechahasta) & ") OR eventual.htethasta is null))"

        OpenRecordset StrSql, rsdet
        
        orden = orden + 1
        montoAcumulador = 0
        porcentaje = 0
        
        If Not rsdet.EOF Then
            Flog.writeline Espacios(Tabulador * 3) & " Se encontraron " & rsdet.RecordCount & " empleados"
        Else
            Flog.writeline Espacios(Tabulador * 3) & " No se encontraron empleados"
        End If
        Flog.writeline
        While Not rsdet.EOF
            Flog.writeline Espacios(Tabulador * 4) & " empleado: " & rsdet!empleg
            'reviso la cantidad de hs diarias segun el regimen horario
            If EsNulo(rsdet!horas) Then
                horas = 0
                Flog.writeline Espacios(Tabulador * 5) & " No se configuro el regimen horario del empleado:" & rsdet("ternro")
            Else
                horas = rsdet!horas
            End If
            Flog.writeline Espacios(Tabulador * 4) & " horas diarias: " & horas
            
            'Topeo la cantidad de dias
            If rsdet!dias > 30 Then
                dias = 30
            Else
                dias = rsdet!dias
            End If
            Flog.writeline Espacios(Tabulador * 4) & " dias: " & dias
            porcentaje = porcentaje + Round((1 * (dias) / 30) * (horas / 8), 2)
        
            rsdet.MoveNext
        Wend
        
        'Calculo el porcentaje de horas
        Flog.writeline Espacios(Tabulador * 1) & "Porcentaje E2: " & porcentaje
        Flog.writeline Espacios(Tabulador * 1) & "Sql: " & StrSql
        valores = "(" & NroProceso & "," & tenro2 & "," & E2(estr2) & "," & indice & ",'" & estrdabr & "'," & porcentaje & ",8," & orden & ")"
        StrSql = " INSERT INTO casa_matriz_det " & campos & " VALUES " & valores
        objConn.Execute StrSql, , adExecuteNoRecords
    Next indice
Next estr2

'------------------------------- Fin sección 2 estructura 2 Hombres --------------------------------
'------------------------------- Comienzo de Sección 2 - Estructura 2 Mujeres ----------------------

'--------------------------------- Comienzo de la Seccion 2 Estructura 2 Mujeres -------------------

Flog.writeline Espacios(Tabulador * 1) & "Sección 2.. Estructuras 2 Mujeres "
For estr2 = 0 To IndiceE2 - 1
    Flog.writeline " "
    Flog.writeline Espacios(Tabulador * 1) & "Insertar datos del mes de la estructura " & E2(estr2)
 
    StrSql = "SELECT estrdabr FROM estructura "
    StrSql = StrSql & " WHERE estructura.estrnro = " & E2(estr2)
    OpenRecordset StrSql, rsdet1
    If Not rsdet1.EOF Then
        estrdabr = rsdet1!estrdabr
    Else
        estrdabr = " "
    End If
    
    For indice = 1 To 12
        fechahasta = ultimo_dia_mes(indice, Anio)
        fechadesde = primer_dia_mes(indice, Anio)
      
        Flog.writeline " "
        Flog.writeline Espacios(Tabulador * 1) & "Insertar datos del mes: " & indice
    
        StrSql = " SELECT distinct empleado.empleg, empleado.ternro, ammonto dias, e2.estrcodext horas"
        StrSql = StrSql & " FROM empleado "
        StrSql = StrSql & " INNER JOIN tercero ON empleado.ternro = tercero.ternro"
        StrSql = StrSql & " INNER JOIN fases ON empleado.ternro = fases.empleado"
        StrSql = StrSql & " INNER JOIN his_estructura as emp ON emp.ternro = empleado.ternro AND emp.tenro = 10"
        StrSql = StrSql & " INNER JOIN his_estructura as he1 ON he1.ternro = empleado.ternro "
        StrSql = StrSql & " AND he1.htetdesde <= " & ConvFecha(fechahasta) & " AND ((he1.htethasta >=" & ConvFecha(fechahasta) & ") OR he1.htethasta is null)"
        StrSql = StrSql & " LEFT JOIN his_estructura as he2 ON he2.ternro = empleado.ternro AND he2.tenro = 21 "
        StrSql = StrSql & " AND he2.htetdesde <= " & ConvFecha(fechahasta) & " AND ((he2.htethasta >=" & ConvFecha(fechahasta) & ") OR he2.htethasta is null)"
        StrSql = StrSql & " INNER JOIN estructura e2 ON he2.estrnro = e2.estrnro"
        StrSql = StrSql & " INNER JOIN acu_mes ON acu_mes.ternro = empleado.ternro "
        StrSql = StrSql & " WHERE acunro= " & AC2
        StrSql = StrSql & " AND ammes = " & Month(fechahasta)
        StrSql = StrSql & " AND amanio = " & Year(fechahasta)
        StrSql = StrSql & " AND he1.estrnro = " & E2(estr2)
        StrSql = StrSql & " AND empleado.ternro IN " & empleados
        StrSql = StrSql & " AND (fases.altfec <= " & ConvFecha(fechahasta) & " AND ( fases.bajfec >= " & ConvFecha(fechahasta) & " OR fases.bajfec IS NULL))"
        If Empresas_Filtradas <> "0" Then
            StrSql = StrSql & " AND emp.estrnro IN (" & Empresas_Filtradas & ")"
            StrSql = StrSql & " AND emp.htetdesde <= " & ConvFecha(fechahasta) & " AND ((emp.htethasta >=" & ConvFecha(fechahasta) & ") OR emp.htethasta is null)"
        End If
        StrSql = StrSql & " AND tersex = 0 "
        StrSql = StrSql & " AND empleado.ternro NOT IN (SELECT tercero.ternro FROM tercero "
        StrSql = StrSql & " INNER JOIN his_estructura as eventual ON eventual.ternro = tercero.ternro "
        StrSql = StrSql & " AND eventual.tenro = 18 AND eventual.estrnro IN (" & contratos_eventuales & ")"
        StrSql = StrSql & " AND eventual.htetdesde <= " & ConvFecha(fechahasta) & " AND ((eventual.htethasta >=" & ConvFecha(fechahasta) & ") OR eventual.htethasta is null))"
        
        OpenRecordset StrSql, rsdet
        
        orden = orden + 1
        montoAcumulador = 0
        porcentaje = 0
        
        If Not rsdet.EOF Then
            Flog.writeline Espacios(Tabulador * 3) & " Se encontraron " & rsdet.RecordCount & " empleados"
        Else
            Flog.writeline Espacios(Tabulador * 3) & " No se encontraron empleados"
        End If
        Flog.writeline
        While Not rsdet.EOF
            Flog.writeline Espacios(Tabulador * 4) & " empleado: " & rsdet!empleg
            'reviso la cantidad de hs diarias segun el regimen horario
            If EsNulo(rsdet!horas) Then
                horas = 0
                Flog.writeline Espacios(Tabulador * 5) & " No se configuro el regimen horario del empleado:" & rsdet("ternro")
            Else
                horas = rsdet!horas
            End If
            Flog.writeline Espacios(Tabulador * 4) & " horas diarias: " & horas
            
            'Topeo la cantidad de dias
            If rsdet!dias > 30 Then
                dias = 30
            Else
                dias = rsdet!dias
            End If
            Flog.writeline Espacios(Tabulador * 4) & " dias: " & dias
            porcentaje = porcentaje + Round((1 * (dias) / 30) * (horas / 8), 2)
        
            rsdet.MoveNext
        Wend
        
        'Calculo el porcentaje de horas
        Flog.writeline Espacios(Tabulador * 1) & "Porcentaje E2: " & porcentaje
        Flog.writeline Espacios(Tabulador * 1) & "Sql: " & StrSql
        valores = "(" & NroProceso & "," & tenro2 & "," & E2(estr2) & "," & indice & ",'" & estrdabr & "'," & porcentaje & ",81," & orden & ")"
        StrSql = " INSERT INTO casa_matriz_det " & campos & " VALUES " & valores
        objConn.Execute StrSql, , adExecuteNoRecords
    Next indice
Next estr2

actualizar_progreso (55)
'------------------------------- Fin sección 2 estructura 2 --------------------------------

'--------------------------------- Comienzo de la Seccion 2 Estructura 3 -------------------
Flog.writeline Espacios(Tabulador * 1) & "Sección 2.. Estructuras 3"
For estr3 = 0 To IndiceE3 - 1
    Flog.writeline " "
    Flog.writeline Espacios(Tabulador * 1) & "Insertar datos del mes de la estructura " & E3(estr3)
    
    StrSql = "SELECT estrdabr FROM estructura "
    StrSql = StrSql & " WHERE estructura.estrnro = " & E3(estr3)
    OpenRecordset StrSql, rsdet1
    
    If Not rsdet1.EOF Then
        estrdabr = rsdet1!estrdabr
    Else
        estrdabr = " "
    End If
    
    For indice = 1 To 12
        fechahasta = ultimo_dia_mes(indice, Anio)
        fechadesde = primer_dia_mes(indice, Anio)
        
        Flog.writeline " "
        Flog.writeline Espacios(Tabulador * 1) & "Insertar datos del mes: " & indice
    
        StrSql = " SELECT distinct empleado.empleg, empleado.ternro, ammonto dias, e2.estrcodext horas"
        StrSql = StrSql & " FROM empleado "
        StrSql = StrSql & " INNER JOIN fases ON empleado.ternro = fases.empleado"
        StrSql = StrSql & " INNER JOIN his_estructura as emp ON emp.ternro = empleado.ternro AND emp.tenro = 10"
        StrSql = StrSql & " INNER JOIN his_estructura as he1 ON he1.ternro = empleado.ternro "
        StrSql = StrSql & " AND he1.htetdesde <= " & ConvFecha(fechahasta) & " AND ((he1.htethasta >=" & ConvFecha(fechahasta) & ") OR he1.htethasta is null)"
        StrSql = StrSql & " LEFT JOIN his_estructura as he2 ON he2.ternro = empleado.ternro AND he2.tenro = 21 "
        StrSql = StrSql & " AND he2.htetdesde <= " & ConvFecha(fechahasta) & " AND ((he2.htethasta >=" & ConvFecha(fechahasta) & ") OR he2.htethasta is null)"
        StrSql = StrSql & " INNER JOIN estructura e2 ON he2.estrnro = e2.estrnro"
        StrSql = StrSql & " INNER JOIN acu_mes ON acu_mes.ternro = empleado.ternro "
        StrSql = StrSql & " WHERE acunro= " & AC2
        StrSql = StrSql & " AND ammes = " & Month(fechahasta)
        StrSql = StrSql & " AND amanio = " & Year(fechahasta)
        StrSql = StrSql & " AND he1.estrnro = " & E3(estr3)
        StrSql = StrSql & " AND empleado.ternro IN " & empleados
        StrSql = StrSql & " AND (fases.altfec <= " & ConvFecha(fechahasta) & " AND ( fases.bajfec >= " & ConvFecha(fechahasta) & " OR fases.bajfec IS NULL))"
        If Empresas_Filtradas <> "0" Then
            StrSql = StrSql & " AND emp.estrnro IN (" & Empresas_Filtradas & ")"
            StrSql = StrSql & " AND emp.htetdesde <= " & ConvFecha(fechahasta) & " AND ((emp.htethasta >=" & ConvFecha(fechahasta) & ") OR emp.htethasta is null)"
        End If
        StrSql = StrSql & " AND empleado.ternro NOT IN (SELECT tercero.ternro FROM tercero "
        StrSql = StrSql & " INNER JOIN his_estructura as eventual ON eventual.ternro = tercero.ternro "
        StrSql = StrSql & " AND eventual.tenro = 18 AND eventual.estrnro IN (" & contratos_eventuales & ")"
        StrSql = StrSql & " AND eventual.htetdesde <= " & ConvFecha(fechahasta) & " AND ((eventual.htethasta >=" & ConvFecha(fechahasta) & ") OR eventual.htethasta is null))"

        
        OpenRecordset StrSql, rsdet
        
        orden = orden + 1
        montoAcumulador = 0
        porcentaje = 0
        If Not rsdet.EOF Then
            Flog.writeline Espacios(Tabulador * 3) & " Se encontraron " & rsdet.RecordCount & " empleados"
        Else
            Flog.writeline Espacios(Tabulador * 3) & " No se encontraron empleados"
        End If
        Flog.writeline
        While Not rsdet.EOF
            Flog.writeline Espacios(Tabulador * 4) & " empleado: " & rsdet!empleg
            'reviso la cantidad de hs diarias segun el regimen horario
            If EsNulo(rsdet!horas) Then
                horas = 0
                Flog.writeline Espacios(Tabulador * 5) & " No se configuro el regimen horario del empleado:" & rsdet("ternro")
            Else
                horas = rsdet!horas
            End If
            Flog.writeline Espacios(Tabulador * 4) & " horas diarias: " & horas
            
            'Topeo la cantidad de dias
            If rsdet!dias > 30 Then
                dias = 30
            Else
                dias = rsdet!dias
            End If
            Flog.writeline Espacios(Tabulador * 4) & " dias: " & dias
            porcentaje = porcentaje + Round((1 * (dias) / 30) * (horas / 8), 2)
        
            rsdet.MoveNext
        Wend
        
        'Calculo el porcentaje de horas
        Flog.writeline Espacios(Tabulador * 1) & "Porcentaje E2: " & porcentaje
        Flog.writeline Espacios(Tabulador * 1) & "Sql: " & StrSql
        valores = "(" & NroProceso & "," & tenro2 & "," & E3(estr3) & "," & indice & ",'" & estrdabr & "'," & porcentaje & ",9," & orden & ")"
        StrSql = " INSERT INTO casa_matriz_det " & campos & " VALUES " & valores
        objConn.Execute StrSql, , adExecuteNoRecords
    Next indice
Next estr3


Exit Sub

MError:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql

End Sub


Sub InsertarDatosdetPlantilla_old()
Dim campos As String
Dim valores As String
Dim fechadesde As Date
Dim fechahasta As Date
Dim estr1 As Integer
Dim estr2 As Integer
Dim estr3 As Integer
Dim indice As Integer
Dim rsdet As New ADODB.Recordset
Dim rsdet1 As New ADODB.Recordset
Dim montoAcumulador As Double
Dim estrdabr As String
Dim orden As Integer
Dim horas As Integer
Dim dias As Integer
Dim porcentaje As Integer


On Error GoTo MError


campos = " (bpronro,tenro,estrnro,mes,estrdabr,monto,tipest,orden)"
orden = 0

'--------------------------------- Comienzo de la Seccion 2 Estructura 1 ------------------------
Flog.writeline Espacios(Tabulador * 1) & "Sección 2"

For estr1 = 0 To IndiceE1 - 1
  Flog.writeline " "
  Flog.writeline Espacios(Tabulador * 1) & "Insertar datos del mes de la estructura " & E1(estr1)
 
    StrSql = "SELECT estrdabr FROM estructura "
    StrSql = StrSql & " WHERE estructura.estrnro = " & E1(estr1)
    OpenRecordset StrSql, rsdet1
    If Not rsdet1.EOF Then
        estrdabr = rsdet1!estrdabr
    Else
        estrdabr = " "
    End If
    
    For indice = 1 To 12
        fechahasta = ultimo_dia_mes(indice, Anio)
        fechadesde = primer_dia_mes(indice, Anio)
      
        
        Flog.writeline " "
        Flog.writeline Espacios(Tabulador * 2) & "Insertar datos del mes: " & indice
    
        'FGZ - 26/03/2009 ---
        ' este query esta mal y complejo al pedo, lo cambio y reedito
'        StrSql = " SELECT proceso.pronro,acu_liq.acunro,acudesabr, almonto, alcant, estructura.estrdabr,empleado.ternro "
'        StrSql = StrSql & " FROM empleado "
'        StrSql = StrSql & " INNER JOIN cabliq ON cabliq.empleado = empleado.ternro "
'        StrSql = StrSql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
'        StrSql = StrSql & " INNER JOIN acumulador ON acumulador.acunro = acu_liq.acunro "
'        StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro "
'        StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro "
'        StrSql = StrSql & " AND periodo.pliqdesde >=" & ConvFecha(fechadesde) & " AND periodo.pliqhasta <= " & ConvFecha(fechahasta)
'        StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.ternro = empleado.ternro "
'        StrSql = StrSql & " AND his_estructura.htetdesde <= " & ConvFecha(fechahasta) & " AND ((his_estructura.htethasta >=" & ConvFecha(fechahasta) & ") OR his_estructura.htethasta is null)"
'        StrSql = StrSql & " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro"
'        StrSql = StrSql & " WHERE acumulador.acunro= " & AC2
'        StrSql = StrSql & " AND estructura.estrnro = " & E1(estr1)
'        StrSql = StrSql & " AND empleado.ternro IN " & empleados
        
        StrSql = " SELECT distinct empleado.empleg, empleado.ternro, ammonto dias, e2.estrcodext horas"
        StrSql = StrSql & " FROM empleado "
        StrSql = StrSql & " INNER JOIN fases ON empleado.ternro = fases.empleado"
        StrSql = StrSql & " INNER JOIN his_estructura as emp ON emp.ternro = empleado.ternro AND emp.tenro = 10"
        StrSql = StrSql & " INNER JOIN his_estructura as he1 ON he1.ternro = empleado.ternro "
        StrSql = StrSql & " AND he1.htetdesde <= " & ConvFecha(fechahasta) & " AND ((he1.htethasta >=" & ConvFecha(fechahasta) & ") OR he1.htethasta is null)"
        StrSql = StrSql & " LEFT JOIN his_estructura as he2 ON he2.ternro = empleado.ternro AND he2.tenro = 21 "
        StrSql = StrSql & " AND he2.htetdesde <= " & ConvFecha(fechahasta) & " AND ((he2.htethasta >=" & ConvFecha(fechahasta) & ") OR he2.htethasta is null)"
        StrSql = StrSql & " INNER JOIN estructura e2 ON he2.estrnro = e2.estrnro"
        StrSql = StrSql & " INNER JOIN acu_mes ON acu_mes.ternro = empleado.ternro "
        StrSql = StrSql & " WHERE acunro= " & AC2
        StrSql = StrSql & " AND ammes = " & Month(fechahasta)
        StrSql = StrSql & " AND amanio = " & Year(fechahasta)
        StrSql = StrSql & " AND he1.estrnro = " & E1(estr1)
        StrSql = StrSql & " AND empleado.ternro IN " & empleados
        StrSql = StrSql & " AND (fases.altfec <= " & ConvFecha(fechahasta) & " AND ( fases.bajfec >= " & ConvFecha(fechahasta) & " OR fases.bajfec IS NULL))"
        If Empresas_Filtradas <> "0" Then
            StrSql = StrSql & " AND emp.estrnro IN (" & Empresas_Filtradas & ")"
        End If
        OpenRecordset StrSql, rsdet
        
        orden = orden + 1
        montoAcumulador = 0
        porcentaje = 0
        
        If Not rsdet.EOF Then
            Flog.writeline Espacios(Tabulador * 3) & " Se encontraron " & rsdet.RecordCount & " empleados"
        Else
            Flog.writeline Espacios(Tabulador * 3) & " No se encontraron empleados"
        End If
        Flog.writeline
        While Not rsdet.EOF
            Flog.writeline Espacios(Tabulador * 4) & " empleado: " & rsdet!empleg
            'reviso la cantidad de hs diarias segun el regimen horario
            If EsNulo(rsdet!horas) Then
                horas = 0
                Flog.writeline Espacios(Tabulador * 5) & " No se configuro el regimen horario del empleado:" & rsdet("ternro")
            Else
                horas = rsdet!horas
            End If
            Flog.writeline Espacios(Tabulador * 4) & " horas diarias: " & horas
            
            'Topeo la cantidad de dias
            If rsdet!dias > 30 Then
                dias = 30
            Else
                dias = rsdet!dias
            End If
            Flog.writeline Espacios(Tabulador * 4) & " dias: " & dias
            porcentaje = porcentaje + Round((1 * (dias) / 30) * (horas / 8), 2)
        
            rsdet.MoveNext
        Wend
        
        'Calculo el porcentaje de horas
        Flog.writeline Espacios(Tabulador * 1) & "Porcentaje E1: " & porcentaje
        Flog.writeline Espacios(Tabulador * 1) & "Sql: " & StrSql
        valores = "(" & NroProceso & "," & tenro1 & "," & E1(estr1) & "," & indice & ",'" & estrdabr & "'," & porcentaje & ",7," & orden & ")"
        StrSql = " INSERT INTO casa_matriz_det " & campos & " VALUES " & valores
        objConn.Execute StrSql, , adExecuteNoRecords
    
    Next indice
Next estr1

actualizar_progreso (50)
'------------------------------- Fin sección 2 estructura 1 ----------------------------------


'------------------------------- Comienzo de Sección 2 - Estructura 2 ------------------------
Flog.writeline Espacios(Tabulador * 1) & "Sección 2.. Estructuras 2"
For estr2 = 0 To IndiceE2 - 1
  Flog.writeline " "
  Flog.writeline Espacios(Tabulador * 1) & "Insertar datos del mes de la estructura " & E2(estr2)
 
    
    StrSql = "SELECT estrdabr FROM estructura "
    StrSql = StrSql & " WHERE estructura.estrnro = " & E2(estr2)
    OpenRecordset StrSql, rsdet1
    
    If Not rsdet1.EOF Then
        estrdabr = rsdet1!estrdabr
    Else
        estrdabr = " "
    End If
    
    For indice = 1 To 12
  
    
    fechahasta = ultimo_dia_mes(indice, Anio)
    fechadesde = primer_dia_mes(indice, Anio)
  
    
    Flog.writeline " "
    Flog.writeline Espacios(Tabulador * 1) & "Insertar datos del mes: " & indice

    StrSql = " SELECT proceso.pronro,acu_liq.acunro,acudesabr, almonto, alcant, estructura.estrdabr,empleado.ternro "
    StrSql = StrSql & " FROM empleado "
    StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro = empleado.ternro "
    StrSql = StrSql & " INNER JOIN cabliq ON cabliq.empleado = empleado.ternro "
    StrSql = StrSql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
    StrSql = StrSql & " INNER JOIN acumulador ON acumulador.acunro = acu_liq.acunro "
    StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro "
    StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro "
    StrSql = StrSql & " AND periodo.pliqdesde >=" & ConvFecha(fechadesde) & " AND periodo.pliqhasta <= " & ConvFecha(fechahasta)
    StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.ternro = empleado.ternro "
    StrSql = StrSql & " AND his_estructura.htetdesde <= " & ConvFecha(fechahasta) & " AND ((his_estructura.htethasta >=" & ConvFecha(fechahasta) & ") OR his_estructura.htethasta is null)"
    StrSql = StrSql & " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro"
    StrSql = StrSql & " WHERE acumulador.acunro= " & AC2
    StrSql = StrSql & " AND estructura.estrnro = " & E2(estr2)
    StrSql = StrSql & " AND empleado.ternro IN " & empleados
    StrSql = StrSql & " AND tersex = -1 "
    OpenRecordset StrSql, rsdet
    
    orden = orden + 1
    montoAcumulador = 0
    porcentaje = 0
    
    ' estructura regimen horario
    
    While Not rsdet.EOF
        
        StrSql = "SELECT estrdabr,estrcodext FROM estructura "
        StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.tenro = 21 "
        StrSql = StrSql & " AND his_estructura.estrnro = estructura.estrnro "
        StrSql = StrSql & " AND his_estructura.ternro = " & rsdet!ternro
        StrSql = StrSql & " AND htetdesde <= " & ConvFecha(fechahasta) & " AND (htethasta >= " & ConvFecha(fechahasta) & " OR htethasta is null)"
        'StrSql = StrSql & " WHERE estructura.estrnro = " & E1(estr1)
        OpenRecordset StrSql, rsdet1
  
        horas = 0
        If Not rsdet1.EOF Then
              horas = rsdet1!estrcodext
        Else
              Flog.writeline Espacios(Tabulador * 1) & " No se configuro el regimen horario del empleado:" & rsdet("ternro")
        End If
    
        porcentaje = porcentaje + Round((1 * (rsdet!almonto) / 30) * (horas / 8), 2)
    
        dias = montoAcumulador
        rsdet.MoveNext
        
    Wend
    
    ' Calculo el porcentaje de horas
        
    Flog.writeline Espacios(Tabulador * 1) & "Porcentaje E2: " & porcentaje
    Flog.writeline Espacios(Tabulador * 1) & "Sql: " & StrSql
    valores = "(" & NroProceso & "," & tenro2 & "," & E2(estr2) & "," & indice & ",'" & estrdabr & "'," & porcentaje & ",8," & orden & ")"
    StrSql = " INSERT INTO casa_matriz_det " & campos & " VALUES " & valores
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Next
    
Next

'------------------------------- Fin sección 2 estructura 2 Hombres --------------------------------
'------------------------------- Comienzo de Sección 2 - Estructura 2 Mujeres ----------------------

'--------------------------------- Comienzo de la Seccion 2 Estructura 2 Mujeres -------------------

Flog.writeline Espacios(Tabulador * 1) & "Sección 2.. Estructuras 2 Mujeres "

For estr2 = 0 To IndiceE2 - 1

  Flog.writeline " "
  Flog.writeline Espacios(Tabulador * 1) & "Insertar datos del mes de la estructura " & E2(estr2)
 
    
    StrSql = "SELECT estrdabr FROM estructura "
    StrSql = StrSql & " WHERE estructura.estrnro = " & E2(estr2)
    OpenRecordset StrSql, rsdet1
    
    If Not rsdet1.EOF Then
        estrdabr = rsdet1!estrdabr
    Else
        estrdabr = " "
    End If
    
    For indice = 1 To 12
  
    
    fechahasta = ultimo_dia_mes(indice, Anio)
    fechadesde = primer_dia_mes(indice, Anio)
  
    
    Flog.writeline " "
    Flog.writeline Espacios(Tabulador * 1) & "Insertar datos del mes: " & indice

    StrSql = " SELECT proceso.pronro,acu_liq.acunro,acudesabr, almonto, alcant, estructura.estrdabr,empleado.ternro "
    StrSql = StrSql & " FROM empleado "
    StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro = empleado.ternro "
    StrSql = StrSql & " INNER JOIN cabliq ON cabliq.empleado = empleado.ternro "
    StrSql = StrSql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
    StrSql = StrSql & " INNER JOIN acumulador ON acumulador.acunro = acu_liq.acunro "
    StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro "
    StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro "
    StrSql = StrSql & " AND periodo.pliqdesde >=" & ConvFecha(fechadesde) & " AND periodo.pliqhasta <= " & ConvFecha(fechahasta)
    StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.ternro = empleado.ternro "
    StrSql = StrSql & " AND his_estructura.htetdesde <= " & ConvFecha(fechahasta) & " AND ((his_estructura.htethasta >=" & ConvFecha(fechahasta) & ") OR his_estructura.htethasta is null)"
    StrSql = StrSql & " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro"
    StrSql = StrSql & " WHERE acumulador.acunro= " & AC2
    StrSql = StrSql & " AND estructura.estrnro = " & E2(estr2)
    StrSql = StrSql & " AND empleado.ternro IN " & empleados
    StrSql = StrSql & " AND tersex = 0 "
    OpenRecordset StrSql, rsdet
    
    orden = orden + 1
    montoAcumulador = 0
    porcentaje = 0
    
    ' estructura regimen horario
    
    While Not rsdet.EOF
        
        StrSql = "SELECT estrdabr,estrcodext FROM estructura "
        StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.tenro = 21 "
        StrSql = StrSql & " AND his_estructura.estrnro = estructura.estrnro "
        StrSql = StrSql & " AND his_estructura.ternro = " & rsdet!ternro
        StrSql = StrSql & " AND htetdesde <= " & ConvFecha(fechahasta) & " AND (htethasta >= " & ConvFecha(fechahasta) & " OR htethasta is null)"
        'StrSql = StrSql & " WHERE estructura.estrnro = " & E1(estr1)
        OpenRecordset StrSql, rsdet1
  
        horas = 0
        If Not rsdet1.EOF Then
              horas = rsdet1!estrcodext
        Else
              Flog.writeline Espacios(Tabulador * 1) & " No se configuro el regimen horario del empleado:" & rsdet("ternro")
        End If
    
        porcentaje = porcentaje + Round((1 * (rsdet!almonto) / 30) * (horas / 8), 2)
    
        dias = montoAcumulador
        rsdet.MoveNext
        
    Wend
    
    ' Calculo el porcentaje de horas
        
    Flog.writeline Espacios(Tabulador * 1) & "Porcentaje E2: " & porcentaje
    Flog.writeline Espacios(Tabulador * 1) & "Sql: " & StrSql
    valores = "(" & NroProceso & "," & tenro2 & "," & E2(estr2) & "," & indice & ",'" & estrdabr & "'," & porcentaje & ",81," & orden & ")"
    StrSql = " INSERT INTO casa_matriz_det " & campos & " VALUES " & valores
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Next
    
Next

actualizar_progreso (40)
'------------------------------- Fin sección 2 estructura 2 Mujeres --------------------------------
'------------------------------- Comienzo de Sección 2 - Estructura 3 ------------------------------

'------------------------------- Fin sección 2 estructura 2 Hombres --------------------------------
'------------------------------- Comienzo de Sección 2 - Estructura 2 Mujeres ----------------------

'--------------------------------- Comienzo de la Seccion 2 Estructura 2 Mujeres -------------------

Flog.writeline Espacios(Tabulador * 1) & "Sección 2.. Estructuras 3 "

For estr3 = 0 To IndiceE3 - 1

  Flog.writeline " "
  Flog.writeline Espacios(Tabulador * 1) & "Insertar datos del mes de la estructura " & E3(estr3)
 
    
    StrSql = "SELECT estrdabr FROM estructura "
    StrSql = StrSql & " WHERE estructura.estrnro = " & E3(estr3)
    OpenRecordset StrSql, rsdet1
    
    If Not rsdet1.EOF Then
        estrdabr = rsdet1!estrdabr
    Else
        estrdabr = " "
    End If
    
    For indice = 1 To 12
  
    
    fechahasta = ultimo_dia_mes(indice, Anio)
    fechadesde = primer_dia_mes(indice, Anio)
  
    
    Flog.writeline " "
    Flog.writeline Espacios(Tabulador * 1) & "Insertar datos del mes: " & indice

    StrSql = " SELECT proceso.pronro,acu_liq.acunro,acudesabr, almonto, alcant, estructura.estrdabr,empleado.ternro "
    StrSql = StrSql & " FROM empleado "
    StrSql = StrSql & " INNER JOIN cabliq ON cabliq.empleado = empleado.ternro "
    StrSql = StrSql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
    StrSql = StrSql & " INNER JOIN acumulador ON acumulador.acunro = acu_liq.acunro "
    StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro "
    StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro "
    StrSql = StrSql & " AND periodo.pliqdesde >=" & ConvFecha(fechadesde) & " AND periodo.pliqhasta <= " & ConvFecha(fechahasta)
    StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.ternro = empleado.ternro "
    StrSql = StrSql & " AND his_estructura.htetdesde <= " & ConvFecha(fechahasta) & " AND ((his_estructura.htethasta >=" & ConvFecha(fechahasta) & ") OR his_estructura.htethasta is null)"
    StrSql = StrSql & " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro"
    StrSql = StrSql & " WHERE acumulador.acunro= " & AC2
    StrSql = StrSql & " AND estructura.estrnro = " & E3(estr3)
    StrSql = StrSql & " AND empleado.ternro IN " & empleados
    OpenRecordset StrSql, rsdet
    
    orden = orden + 1
    montoAcumulador = 0
    porcentaje = 0
    
    ' estructura regimen horario
    
    
    While Not rsdet.EOF
        
        StrSql = "SELECT estrdabr,estrcodext FROM estructura "
        StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.tenro = 21 "
        StrSql = StrSql & " AND his_estructura.estrnro = estructura.estrnro "
        StrSql = StrSql & " AND his_estructura.ternro = " & rsdet!ternro
        StrSql = StrSql & " AND htetdesde <= " & ConvFecha(fechahasta) & " AND (htethasta >= " & ConvFecha(fechahasta) & " OR htethasta is null)"
        'StrSql = StrSql & " WHERE estructura.estrnro = " & E1(estr1)
        OpenRecordset StrSql, rsdet1
  
        horas = 0
        If Not rsdet1.EOF Then
              horas = rsdet1!estrcodext
        Else
              Flog.writeline Espacios(Tabulador * 1) & " No se configuro el regimen horario del empleado:" & rsdet("ternro")
        End If
    
        porcentaje = porcentaje + Round((1 * (rsdet!almonto) / 30) * (horas / 8), 2)
    
        dias = montoAcumulador
        rsdet.MoveNext
        
    Wend
    
    ' Calculo el porcentaje de horas
        
    Flog.writeline Espacios(Tabulador * 1) & "Porcentaje E2: " & porcentaje
    Flog.writeline Espacios(Tabulador * 1) & "Sql: " & StrSql
    valores = "(" & NroProceso & "," & tenro2 & "," & E3(estr3) & "," & indice & ",'" & estrdabr & "'," & porcentaje & ",9," & orden & ")"
    StrSql = " INSERT INTO casa_matriz_det " & campos & " VALUES " & valores
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Next
    
Next

Exit Sub

MError:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql

End Sub


Sub totales_seccion()

Dim rstot As New ADODB.Recordset
Dim indice As Integer
Dim campos As String
Dim valores As String

' Totales est 1 secc 1

 For indice = 1 To 12
 
    Flog.writeline " "
    Flog.writeline Espacios(Tabulador * 1) & "Calculando Totales sección 1: Tipo de estructura 1.Mes:" & indice

    StrSql = " SELECT sum(monto) montoTotal "
    StrSql = StrSql & " FROM casa_matriz_det WHERE tipest = '1' and bpronro = " & NroProceso
    StrSql = StrSql & " AND mes = " & indice
    OpenRecordset StrSql, rstot
    
    Flog.writeline Espacios(Tabulador * 1) & "Monto total E1: " & rstot!montoTotal

    campos = " (bpronro,mes,monto,tipest)"
    valores = "(" & NroProceso & "," & indice & "," & IIf(Not EsNulo(rstot!montoTotal), rstot!montoTotal, 0) & ",'4')"
    
    StrSql = " INSERT INTO casa_matriz_det " & campos & " VALUES " & valores
    objConn.Execute StrSql, , adExecuteNoRecords
    
    
    If Not EsNulo(rstot!montoTotal) Then
        HayDetalleLiq(indice) = rstot!montoTotal
    End If
    
 Next
 
' Totales est 2 secc 1

For indice = 1 To 12
 
    Flog.writeline " "
    Flog.writeline Espacios(Tabulador * 1) & "Calculando Totales sección 1: Tipo de estructura 2.Mes:" & indice

    StrSql = " SELECT sum(monto) montoTotal "
    StrSql = StrSql & " FROM casa_matriz_det WHERE tipest = '2' and bpronro = " & NroProceso
    StrSql = StrSql & " AND mes = " & indice
    OpenRecordset StrSql, rstot
    
    Flog.writeline Espacios(Tabulador * 1) & "Monto total E1: " & rstot!montoTotal

    campos = " (bpronro,mes,monto,tipest)"
    valores = "(" & NroProceso & "," & indice & "," & IIf(Not EsNulo(rstot!montoTotal), rstot!montoTotal, 0) & ",'5')"
    
    StrSql = " INSERT INTO casa_matriz_det " & campos & " VALUES " & valores
    objConn.Execute StrSql, , adExecuteNoRecords
 Next
 
' Totales est 3 secc 1

For indice = 1 To 12
    Flog.writeline " "
    Flog.writeline Espacios(Tabulador * 1) & "Calculando Totales sección 1: Tipo de estructura 3.Mes:" & indice

    StrSql = " SELECT sum(monto) montoTotal "
    StrSql = StrSql & " FROM casa_matriz_det WHERE tipest = '3' and bpronro = " & NroProceso
    StrSql = StrSql & " AND mes = " & indice
    OpenRecordset StrSql, rstot
    
    Flog.writeline Espacios(Tabulador * 1) & "Monto total E1: " & rstot!montoTotal

    campos = " (bpronro,mes,monto,tipest)"
    valores = "(" & NroProceso & "," & indice & "," & IIf(Not EsNulo(rstot!montoTotal), rstot!montoTotal, 0) & ",'6')"
    
    StrSql = " INSERT INTO casa_matriz_det " & campos & " VALUES " & valores
    objConn.Execute StrSql, , adExecuteNoRecords
 Next

' Calculo totales estructuras E1 sec 2
 For indice = 1 To 12
 
    Flog.writeline " "
    Flog.writeline Espacios(Tabulador * 1) & "Calculando Totales sección 2: Tipo de estructura 1.Mes:" & indice

    StrSql = " SELECT sum(monto) montoTotal "
    StrSql = StrSql & " FROM casa_matriz_det WHERE tipest = '7' and bpronro = " & NroProceso
    StrSql = StrSql & " AND mes = " & indice
    OpenRecordset StrSql, rstot
    
    Flog.writeline Espacios(Tabulador * 1) & "Monto total E1: " & rstot!montoTotal

    campos = " (bpronro,mes,monto,tipest)"
    valores = "(" & NroProceso & "," & indice & "," & IIf(Not EsNulo(rstot!montoTotal), rstot!montoTotal, 0) & ",'11')"
    
    StrSql = " INSERT INTO casa_matriz_det " & campos & " VALUES " & valores
    objConn.Execute StrSql, , adExecuteNoRecords
 Next

' Calculo totales estructuras E2 sec 2

 For indice = 1 To 12
 
    Flog.writeline " "
    Flog.writeline Espacios(Tabulador * 1) & "Calculando Totales sección 2: Tipo de estructura 2.Mes:" & indice

    StrSql = " SELECT sum(monto) montoTotal "
    StrSql = StrSql & " FROM casa_matriz_det WHERE (tipest = '8' OR tipest = '81') and bpronro = " & NroProceso
    StrSql = StrSql & " AND mes = " & indice
    OpenRecordset StrSql, rstot
    
    Flog.writeline Espacios(Tabulador * 1) & "Monto total E2: " & rstot!montoTotal

    campos = " (bpronro,mes,monto,tipest)"
    valores = "(" & NroProceso & "," & indice & "," & IIf(Not EsNulo(rstot!montoTotal), rstot!montoTotal, 0) & ",'12')"
    
    
    StrSql = " INSERT INTO casa_matriz_det " & campos & " VALUES " & valores
    objConn.Execute StrSql, , adExecuteNoRecords
 Next

' Calculo totales estructuras E3 sec 2
 For indice = 1 To 12
 
    Flog.writeline " "
    Flog.writeline Espacios(Tabulador * 1) & "Calculando Totales sección 2: Tipo de estructura 3.Mes:" & indice

    StrSql = " SELECT sum(monto) montoTotal "
    StrSql = StrSql & " FROM casa_matriz_det WHERE tipest = '9' and bpronro = " & NroProceso
    StrSql = StrSql & " AND mes = " & indice
    OpenRecordset StrSql, rstot
    
    Flog.writeline Espacios(Tabulador * 1) & "Monto total E3: " & rstot!montoTotal

    campos = " (bpronro,mes,monto,tipest)"
    valores = "(" & NroProceso & "," & indice & "," & IIf(Not EsNulo(rstot!montoTotal), rstot!montoTotal, 0) & ",'13')"
    
    
    StrSql = " INSERT INTO casa_matriz_det " & campos & " VALUES " & valores
    objConn.Execute StrSql, , adExecuteNoRecords
  
 Next

Exit Sub

MError:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql

End Sub

Sub calcular_plantilla_activa()

Dim rsplantilla As New ADODB.Recordset
Dim indice As Integer
Dim Fecha As String
Dim Cantidad As Integer
Dim campos As String
Dim valores As String

Flog.writeline " "

For indice = 1 To 12
    Fecha = ultimo_dia_mes(indice, Anio)
    Flog.writeline Espacios(Tabulador * 1) & "Calculando la plantilla activa al:" & Fecha
    
    StrSql = " SELECT count(distinct empleado) cantidad "
    StrSql = StrSql & " FROM fases "
    StrSql = StrSql & " INNER JOIN his_estructura as emp ON emp.ternro = fases.empleado AND emp.tenro = 10"
    StrSql = StrSql & " WHERE altfec <= " & ConvFecha(Fecha)
    StrSql = StrSql & " AND (bajfec is null or bajfec >=" & ConvFecha(Fecha)
    StrSql = StrSql & " ) AND empleado in " & empleados
    If Empresas_Filtradas <> "0" Then
        StrSql = StrSql & " AND emp.estrnro IN (" & Empresas_Filtradas & ")"
        StrSql = StrSql & " AND emp.htetdesde <= " & ConvFecha(Fecha) & " AND ((emp.htethasta >=" & ConvFecha(Fecha) & ") OR emp.htethasta is null)"
    End If
    OpenRecordset StrSql, rsplantilla
    If Not rsplantilla.EOF Then
        Cantidad = rsplantilla!Cantidad
    Else
        Cantidad = 0
    End If
    
    If HayDetalleLiq(indice) = 0 Then
        Cantidad = 0
    End If
    
    
    Flog.writeline Espacios(Tabulador * 1) & "Monto Plantilla activa: " & Cantidad
    Flog.writeline Espacios(Tabulador * 1) & StrSql
    
    campos = " (bpronro,mes,monto,tipest)"
    valores = "(" & NroProceso & "," & indice & "," & Cantidad & ",'10')"
        
    StrSql = " INSERT INTO casa_matriz_det " & campos & " VALUES " & valores
    objConn.Execute StrSql, , adExecuteNoRecords
Next indice
 
 
Exit Sub

MError:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql

End Sub

Sub altas_bajas()
Dim rsplantilla As New ADODB.Recordset
Dim indice As Integer
Dim Fecha As String
Dim fechadesde As String
Dim Cantidad As Integer
Dim campos As String
Dim valores As String

Flog.writeline " "


For indice = 1 To 12
    Fecha = ultimo_dia_mes(indice, Anio)
    fechadesde = primer_dia_mes(indice, Anio)
    
    'Flog.writeline Espacios(Tabulador * 1) & "Altas menusales al:" & Fecha
    StrSql = " SELECT Count(distinct empleado) as cantidad "
    StrSql = StrSql & " FROM fases "
    StrSql = StrSql & " INNER JOIN his_estructura as emp ON emp.ternro = fases.empleado AND emp.tenro = 10"
    StrSql = StrSql & " WHERE altfec <= " & ConvFecha(Fecha)
    StrSql = StrSql & " AND altfec >= " & ConvFecha(fechadesde)
    StrSql = StrSql & "  AND empleado IN " & empleados
    If Empresas_Filtradas <> "0" Then
        StrSql = StrSql & " AND emp.estrnro IN (" & Empresas_Filtradas & ")"
        StrSql = StrSql & " AND emp.htetdesde <= " & ConvFecha(Fecha) & " AND ((emp.htethasta >=" & ConvFecha(Fecha) & ") OR emp.htethasta is null)"
    End If
    OpenRecordset StrSql, rsplantilla
    
    Cantidad = 0
    If Not rsplantilla.EOF Then
        Cantidad = rsplantilla!Cantidad
    Else
        Cantidad = 0
    End If
    Flog.writeline Espacios(Tabulador * 1) & "Cantidad de altas: " & Cantidad

    campos = " (bpronro,mes,monto,tipest)"
    valores = "(" & NroProceso & "," & indice & "," & Cantidad & ",'17')"
        
    StrSql = " INSERT INTO casa_matriz_det " & campos & " VALUES " & valores
    objConn.Execute StrSql, , adExecuteNoRecords
  
    'Flog.writeline Espacios(Tabulador * 1) & "Cantidad de Bajas al:" & Fecha
    StrSql = " SELECT count(distinct empleado) as cantidad "
    StrSql = StrSql & " FROM fases "
    StrSql = StrSql & " INNER JOIN his_estructura as emp ON emp.ternro = fases.empleado AND emp.tenro = 10"
    StrSql = StrSql & " WHERE bajfec <= " & ConvFecha(Fecha)
    StrSql = StrSql & " AND bajfec >= " & ConvFecha(fechadesde)
    StrSql = StrSql & " AND empleado IN " & empleados
    If Empresas_Filtradas <> "0" Then
        StrSql = StrSql & " AND emp.estrnro IN (" & Empresas_Filtradas & ")"
        StrSql = StrSql & " AND emp.htetdesde <= " & ConvFecha(Fecha) & " AND ((emp.htethasta >=" & ConvFecha(Fecha) & ") OR emp.htethasta is null)"
    End If
    OpenRecordset StrSql, rsplantilla
    
    Cantidad = 0
    If Not rsplantilla.EOF Then
        Cantidad = rsplantilla!Cantidad
    Else
        Cantidad = 0
    End If
    Flog.writeline Espacios(Tabulador * 1) & "Cantidad de bajas " & Cantidad

    campos = " (bpronro,mes,monto,tipest)"
    valores = "(" & NroProceso & "," & indice & "," & Cantidad & ",'18')"
        
    StrSql = " INSERT INTO casa_matriz_det " & campos & " VALUES " & valores
    objConn.Execute StrSql, , adExecuteNoRecords
Next indice
 
Exit Sub

MError:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql

End Sub

Sub alta_contrato_eventual()

Dim rsplantilla As New ADODB.Recordset
Dim indice As Integer
Dim Fecha As String
Dim fechadesde As String
Dim Cantidad As Integer
Dim campos As String
Dim valores As String

Flog.writeline " "


For indice = 1 To 12
    Fecha = ultimo_dia_mes(indice, Anio)
    fechadesde = primer_dia_mes(indice, Anio)
    
    ' Calculo los contratos eventuales que se generaron en el mes
    
    StrSql = " SELECT Count(distinct ternro) as cantidad "
    StrSql = StrSql & " FROM his_estructura eventual "
    StrSql = StrSql & " WHERE ternro IN " & empleados
    StrSql = StrSql & " AND eventual.estrnro IN (" & contratos_eventuales & ")"
    StrSql = StrSql & " AND eventual.htetdesde <= " & ConvFecha(Fecha) & " AND (eventual.htetdesde >=" & ConvFecha(fechadesde) & ") "
    
    OpenRecordset StrSql, rsplantilla
    
    Cantidad = 0
    If Not rsplantilla.EOF Then
        Cantidad = rsplantilla!Cantidad
    Else
        Cantidad = 0
    End If
    Flog.writeline Espacios(Tabulador * 1) & "Cantidad de contratos eventuales nuevos: " & Cantidad

    campos = " (bpronro,mes,monto,tipest)"
    valores = "(" & NroProceso & "," & indice & "," & Cantidad & ",'19')"
        
    StrSql = " INSERT INTO casa_matriz_det " & campos & " VALUES " & valores
    objConn.Execute StrSql, , adExecuteNoRecords

Next indice
 
Exit Sub

MError:
    Flog.writeline "    Error: " & Err.Description & " .El error se produjo en alta_contrato_eventual."
    Flog.writeline "    SQL Ejecutado: " & StrSql

End Sub

Sub eventuales()

Dim campos As String
Dim valores As String
Dim fechadesde As Date
Dim fechahasta As Date
Dim estr5 As Integer
Dim indice As Integer
Dim rsdet As New ADODB.Recordset
Dim rstot As New ADODB.Recordset
Dim montoAcumulador As Double
Dim estrdabr As String
Dim orden As Integer
On Error GoTo MError


campos = " (bpronro,tenro,estrnro,mes,estrdabr,monto,tipest,orden)"
    
orden = 0

' - Seccion Eventuales

For estr5 = 0 To indiceE5 - 1
    Flog.writeline " "
    Flog.writeline Espacios(Tabulador * 1) & "Insertar datos del mes de los contratos eventuales: " & E5(estr5)
  
    'estructura
    StrSql = "SELECT estrdabr FROM estructura "
    StrSql = StrSql & " WHERE estructura.estrnro = " & E5(estr5)
    OpenRecordset StrSql, rsdet
  
    If Not rsdet.EOF Then
        estrdabr = rsdet!estrdabr
    Else
        estrdabr = " "
    End If
    
    For indice = 1 To 12
        fechahasta = ultimo_dia_mes(indice, Anio)
        fechadesde = primer_dia_mes(indice, Anio)
        
        
        Flog.writeline " "
        Flog.writeline Espacios(Tabulador * 1) & "Insertar datos del mes: " & indice
        
        StrSql = " SELECT distinct empleado.empleg, empleado.ternro, ammonto monto"
        StrSql = StrSql & " FROM empleado "
        StrSql = StrSql & " INNER JOIN fases ON empleado.ternro = fases.empleado"
        StrSql = StrSql & " INNER JOIN his_estructura as emp ON emp.ternro = empleado.ternro AND emp.tenro = 10 "
        StrSql = StrSql & " INNER JOIN his_estructura as he1 ON he1.ternro = empleado.ternro "
        StrSql = StrSql & " AND he1.htetdesde <= " & ConvFecha(fechahasta) & " AND ((he1.htethasta >=" & ConvFecha(fechahasta) & ") OR he1.htethasta is null)"
        StrSql = StrSql & " INNER JOIN acu_mes ON acu_mes.ternro = empleado.ternro "
        StrSql = StrSql & " WHERE acunro= " & AC1
        StrSql = StrSql & " AND ammes = " & Month(fechahasta)
        StrSql = StrSql & " AND amanio = " & Year(fechahasta)
        StrSql = StrSql & " AND he1.estrnro = " & E5(estr5)
        StrSql = StrSql & " AND empleado.ternro IN " & empleados
        StrSql = StrSql & " AND (fases.altfec <= " & ConvFecha(fechahasta) & " AND ( fases.bajfec >= " & ConvFecha(fechahasta) & " OR fases.bajfec IS NULL))"
        If Empresas_Filtradas <> "0" Then
            StrSql = StrSql & " AND emp.estrnro IN (" & Empresas_Filtradas & ")"
            StrSql = StrSql & " AND emp.htetdesde <= " & ConvFecha(fechahasta) & " AND ((emp.htethasta >=" & ConvFecha(fechahasta) & ") OR emp.htethasta is null)"
        End If
        
        
        OpenRecordset StrSql, rsdet
        
        orden = orden + 1
        montoAcumulador = 0
        While Not rsdet.EOF
            montoAcumulador = montoAcumulador + rsdet!Monto
            rsdet.MoveNext
        Wend
        ' Divido por el valor de la moneda
        
        montoAcumulador = montoAcumulador / Monto
        
        Flog.writeline Espacios(Tabulador * 1) & "Monto CE: " & montoAcumulador
        valores = "(" & NroProceso & "," & "18" & "," & E5(estr5) & "," & indice & ",'" & estrdabr & "'," & montoAcumulador & ",'21'," & orden & ")"
        
        StrSql = " INSERT INTO casa_matriz_det " & campos & " VALUES " & valores
        objConn.Execute StrSql, , adExecuteNoRecords
    Next indice
Next estr5

' Calculo los totales

For indice = 1 To 12
 
    Flog.writeline " "
    Flog.writeline Espacios(Tabulador * 1) & "Calculando Totales sección eventuales:Mes:" & indice

    StrSql = " SELECT sum(monto) montoTotal "
    StrSql = StrSql & " FROM casa_matriz_det WHERE tipest = '21' and bpronro = " & NroProceso
    StrSql = StrSql & " AND mes = " & indice
    OpenRecordset StrSql, rstot
    
    Flog.writeline Espacios(Tabulador * 1) & "Monto total CE: " & rstot!montoTotal

    campos = " (bpronro,mes,monto,tipest)"
    valores = "(" & NroProceso & "," & indice & "," & IIf(Not EsNulo(rstot!montoTotal), rstot!montoTotal, 0) & ",'20')"
    
    StrSql = " INSERT INTO casa_matriz_det " & campos & " VALUES " & valores
    objConn.Execute StrSql, , adExecuteNoRecords
    
    
    If Not EsNulo(rstot!montoTotal) Then
        HayDetalleLiq(indice) = rstot!montoTotal
    End If
    
 Next

Exit Sub

MError:
    Flog.writeline "    Error: " & Err.Description & " .El error se produjo en el calculo de eventualuales."
    Flog.writeline "    SQL Ejecutado: " & StrSql

End Sub

Sub altas_bajas_old()
Dim rsplantilla As New ADODB.Recordset
Dim indice As Integer
Dim Fecha As String
Dim fechadesde As String
Dim Cantidad As Integer
Dim campos As String
Dim valores As String

Flog.writeline " "


For indice = 1 To 12
    Fecha = ultimo_dia_mes(indice, Anio)
    fechadesde = primer_dia_mes(indice, Anio)
    Flog.writeline Espacios(Tabulador * 1) & "Altas menusales al:" & Fecha
    StrSql = " SELECT distinct empleado "
    StrSql = StrSql & " FROM fases "
    StrSql = StrSql & " WHERE altfec <= " & ConvFecha(Fecha)
    StrSql = StrSql & " AND altfec >= " & ConvFecha(fechadesde)
    StrSql = StrSql & "  AND empleado IN " & empleados
    OpenRecordset StrSql, rsplantilla
    
    Cantidad = 0
    While Not rsplantilla.EOF
    
        Cantidad = Cantidad + 1
        rsplantilla.MoveNext
    
    Wend
    
    Flog.writeline Espacios(Tabulador * 1) & "Monto altas: " & Cantidad

    campos = " (bpronro,mes,monto,tipest)"
    
    valores = "(" & NroProceso & "," & indice & "," & Cantidad & ",'17')"
        
    
    StrSql = " INSERT INTO casa_matriz_det " & campos & " VALUES " & valores
    objConn.Execute StrSql, , adExecuteNoRecords
  
  Flog.writeline Espacios(Tabulador * 1) & "Bajas menusales al:" & Fecha
    StrSql = " SELECT distinct empleado "
    StrSql = StrSql & " FROM fases "
    StrSql = StrSql & " WHERE bajfec <= " & ConvFecha(Fecha)
    StrSql = StrSql & " AND bajfec >= " & ConvFecha(fechadesde)
    StrSql = StrSql & " AND empleado in " & empleados
    OpenRecordset StrSql, rsplantilla
    
    Cantidad = 0
    While Not rsplantilla.EOF
    
        Cantidad = Cantidad + 1
        rsplantilla.MoveNext
    
    Wend
    
    Flog.writeline Espacios(Tabulador * 1) & "Monto altas: " & Cantidad

    campos = " (bpronro,mes,monto,tipest)"
    
    valores = "(" & NroProceso & "," & indice & "," & Cantidad & ",'18')"
        
    
    StrSql = " INSERT INTO casa_matriz_det " & campos & " VALUES " & valores
    objConn.Execute StrSql, , adExecuteNoRecords
  
 Next
 
 
Exit Sub

MError:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql

End Sub


Sub calcular_tipo_contrato()
Dim rsplantilla As New ADODB.Recordset
Dim indice As Integer
Dim Fecha As String
Dim fechadesde As Date
Dim fechahasta As Date
Dim Cantidad As Double
Dim cantidad2 As Double
Dim campos As String
Dim valores As String
Dim suma As Integer
Dim listaemp As String
Dim rsdet As New ADODB.Recordset
Dim horas As Integer
Dim dias As Integer
Dim UltimoLegajo As String


Flog.writeline " "
Flog.writeline "------------------------------------------------------------------------"
Flog.writeline " Contratos "
    For indice = 1 To 12
        fechahasta = ultimo_dia_mes(indice, Anio)
        fechadesde = primer_dia_mes(indice, Anio)
        
        
        Flog.writeline Espacios(Tabulador * 1) & "----> Calculando empleados por contrato Fijo y eventuales: " & fechahasta
        
        StrSql = " SELECT distinct empleado.empleg, empleado.ternro, ammonto dias, e2.estrcodext horas"
        StrSql = StrSql & " FROM empleado "
        StrSql = StrSql & " INNER JOIN fases ON empleado.ternro = fases.empleado"
        StrSql = StrSql & " INNER JOIN his_estructura as emp ON emp.ternro = empleado.ternro AND emp.tenro = 10"
        StrSql = StrSql & " INNER JOIN his_estructura as he1 ON he1.ternro = empleado.ternro "
        StrSql = StrSql & " AND he1.htetdesde <= " & ConvFecha(fechahasta) & " AND ((he1.htethasta >=" & ConvFecha(fechahasta) & ") OR he1.htethasta is null)"
        StrSql = StrSql & " LEFT JOIN his_estructura as he2 ON he2.ternro = empleado.ternro AND he2.tenro = 21 "
        StrSql = StrSql & " AND he2.htetdesde <= " & ConvFecha(fechahasta) & " AND ((he2.htethasta >=" & ConvFecha(fechahasta) & ") OR he2.htethasta is null)"
        StrSql = StrSql & " INNER JOIN estructura e2 ON he2.estrnro = e2.estrnro"
        StrSql = StrSql & " INNER JOIN acu_mes ON acu_mes.ternro = empleado.ternro "
        StrSql = StrSql & " WHERE acunro= " & AC2
        StrSql = StrSql & " AND ammes = " & Month(fechahasta)
        StrSql = StrSql & " AND amanio = " & Year(fechahasta)
        StrSql = StrSql & " AND he1.tenro = 18 AND (he1.estrnro IN (" & E4 & ") OR he1.estrnro IN (" & contratos_eventuales & "))"
        StrSql = StrSql & " AND empleado.ternro IN " & empleados
        StrSql = StrSql & " AND (fases.altfec <= " & ConvFecha(fechahasta) & " AND ( fases.bajfec >= " & ConvFecha(fechahasta) & " OR fases.bajfec IS NULL))"
        If Empresas_Filtradas <> "0" Then
            StrSql = StrSql & " AND emp.estrnro IN (" & Empresas_Filtradas & ")"
            StrSql = StrSql & " AND emp.htetdesde <= " & ConvFecha(fechahasta) & " AND ((emp.htethasta >=" & ConvFecha(fechahasta) & ") OR emp.htethasta is null)"
        End If
        OpenRecordset StrSql, rsdet
        
        Cantidad = 0
        If Not rsdet.EOF Then
            Flog.writeline Espacios(Tabulador * 3) & " Se encontraron " & rsdet.RecordCount & " empleados"
        Else
            Flog.writeline Espacios(Tabulador * 3) & " No se encontraron empleados"
        End If
        Flog.writeline
        UltimoLegajo = 0
        While Not rsdet.EOF
            If UltimoLegajo <> rsdet!empleg Then
                UltimoLegajo = rsdet!empleg
                Flog.writeline Espacios(Tabulador * 4) & " empleado: " & rsdet!empleg
                'reviso la cantidad de hs diarias segun el regimen horario
                If EsNulo(rsdet!horas) Then
                    horas = 0
                    Flog.writeline Espacios(Tabulador * 5) & " No se configuro el regimen horario del empleado:" & rsdet("ternro")
                Else
                    horas = rsdet!horas
                End If
                Flog.writeline Espacios(Tabulador * 4) & " horas diarias: " & horas
                
                'Topeo la cantidad de dias
                If rsdet!dias > 30 Then
                    dias = 30
                Else
                    dias = rsdet!dias
                End If
                Flog.writeline Espacios(Tabulador * 4) & " dias: " & dias
                Cantidad = Cantidad + Round((1 * (dias) / 30) * (horas / 8), 2)
            Else
                Flog.writeline Espacios(Tabulador * 5) & "Legajo repetido: " & rsdet!empleg
            End If
            rsdet.MoveNext
        Wend
        
        'Inserto los valores
        Flog.writeline Espacios(Tabulador * 1) & "Cantidad contrato Fijo: " & Cantidad

        campos = " (bpronro,mes,monto,tipest)"
        valores = "(" & NroProceso & "," & indice & "," & Cantidad & ",'14')"
    
        StrSql = " INSERT INTO casa_matriz_det " & campos & " VALUES " & valores
        objConn.Execute StrSql, , adExecuteNoRecords
  
        '--------------------------------------------------------------------------------
        'Calculo los temporales
        Flog.writeline Espacios(Tabulador * 1) & "----> Calculando empleados contrato temporal: " & Fecha
        
        StrSql = " SELECT distinct empleado.empleg, empleado.ternro, ammonto dias, e2.estrcodext horas"
        StrSql = StrSql & " FROM empleado "
        StrSql = StrSql & " INNER JOIN fases ON empleado.ternro = fases.empleado"
        StrSql = StrSql & " INNER JOIN his_estructura as emp ON emp.ternro = empleado.ternro AND emp.tenro = 10"
        StrSql = StrSql & " INNER JOIN his_estructura as he1 ON he1.ternro = empleado.ternro "
        StrSql = StrSql & " AND he1.htetdesde <= " & ConvFecha(fechahasta) & " AND ((he1.htethasta >=" & ConvFecha(fechahasta) & ") OR he1.htethasta is null)"
        StrSql = StrSql & " LEFT JOIN his_estructura as he2 ON he2.ternro = empleado.ternro AND he2.tenro = 21 "
        StrSql = StrSql & " AND he2.htetdesde <= " & ConvFecha(fechahasta) & " AND ((he2.htethasta >=" & ConvFecha(fechahasta) & ") OR he2.htethasta is null)"
        StrSql = StrSql & " INNER JOIN estructura e2 ON he2.estrnro = e2.estrnro"
        StrSql = StrSql & " INNER JOIN acu_mes ON acu_mes.ternro = empleado.ternro "
        StrSql = StrSql & " WHERE acunro= " & AC2
        StrSql = StrSql & " AND ammes = " & Month(fechahasta)
        StrSql = StrSql & " AND amanio = " & Year(fechahasta)
        StrSql = StrSql & " AND he1.tenro = 18 AND he1.estrnro NOT IN (" & E4 & ")"
        StrSql = StrSql & " AND empleado.ternro IN " & empleados
        StrSql = StrSql & " AND (fases.altfec <= " & ConvFecha(fechahasta) & " AND ( fases.bajfec >= " & ConvFecha(fechahasta) & " OR fases.bajfec IS NULL))"
        If Empresas_Filtradas <> "0" Then
            StrSql = StrSql & " AND emp.estrnro IN (" & Empresas_Filtradas & ")"
            StrSql = StrSql & " AND emp.htetdesde <= " & ConvFecha(fechahasta) & " AND ((emp.htethasta >=" & ConvFecha(fechahasta) & ") OR emp.htethasta is null)"
        End If
        StrSql = StrSql & " AND empleado.ternro NOT IN (SELECT tercero.ternro FROM tercero "
        StrSql = StrSql & " INNER JOIN his_estructura as eventual ON eventual.ternro = tercero.ternro "
        StrSql = StrSql & " AND eventual.tenro = 18 AND eventual.estrnro IN (" & contratos_eventuales & ")"
        StrSql = StrSql & " AND eventual.htetdesde <= " & ConvFecha(fechahasta) & " AND ((eventual.htethasta >=" & ConvFecha(fechahasta) & ") OR eventual.htethasta is null))"

        
        OpenRecordset StrSql, rsdet
        
        cantidad2 = 0
        If Not rsdet.EOF Then
            Flog.writeline Espacios(Tabulador * 3) & " Se encontraron " & rsdet.RecordCount & " empleados"
        Else
            Flog.writeline Espacios(Tabulador * 3) & " No se encontraron empleados"
        End If
        Flog.writeline
        UltimoLegajo = 0
        While Not rsdet.EOF
            If UltimoLegajo <> rsdet!empleg Then
                UltimoLegajo = rsdet!empleg
                Flog.writeline Espacios(Tabulador * 4) & " empleado: " & rsdet!empleg
                'reviso la cantidad de hs diarias segun el regimen horario
                If EsNulo(rsdet!horas) Then
                    horas = 0
                    Flog.writeline Espacios(Tabulador * 5) & " No se configuro el regimen horario del empleado:" & rsdet("ternro")
                Else
                    horas = rsdet!horas
                End If
                Flog.writeline Espacios(Tabulador * 4) & " horas diarias: " & horas
                
                'Topeo la cantidad de dias
                If rsdet!dias > 30 Then
                    dias = 30
                Else
                    dias = rsdet!dias
                End If
                Flog.writeline Espacios(Tabulador * 4) & " dias: " & dias
                cantidad2 = cantidad2 + Round((1 * (dias) / 30) * (horas / 8), 2)
            Else
                Flog.writeline Espacios(Tabulador * 5) & "Legajo repetido: " & rsdet!empleg
            End If
            rsdet.MoveNext
        Wend
        
        'Inserto los valores
        Flog.writeline Espacios(Tabulador * 1) & "Cantidad contrato Temporal: " & cantidad2
        
        campos = " (bpronro,mes,monto,tipest)"
        valores = "(" & NroProceso & "," & indice & "," & cantidad2 & ",'15')"
        
        StrSql = " INSERT INTO casa_matriz_det " & campos & " VALUES " & valores
        objConn.Execute StrSql, , adExecuteNoRecords
 
        Flog.writeline Espacios(Tabulador * 1) & "Cantidad empleados con contratos fijos y temporales: " & Cantidad + cantidad2
        campos = " (bpronro,mes,monto,tipest)"
        valores = "(" & NroProceso & "," & indice & "," & Cantidad + cantidad2 & ",'16')"
    
        StrSql = " INSERT INTO casa_matriz_det " & campos & " VALUES " & valores
        objConn.Execute StrSql, , adExecuteNoRecords
    Next indice


'------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------
'For indice = 1 To 12
'
'    Fecha = ultimo_dia_mes(indice, Anio)
'    Flog.writeline Espacios(Tabulador * 1) & "Calculando empleados por contrato :" & Fecha
'
'    StrSql = " SELECT count(distinct empleado) as cantidad "
'    StrSql = StrSql & " FROM fases "
'    StrSql = StrSql & " INNER JOIN his_estructura as emp ON emp.ternro = fases.empleado AND emp.tenro = 10"
'    StrSql = StrSql & " INNER JOIN his_estructura as he ON he.ternro = fases.empleado "
'    StrSql = StrSql & " WHERE altfec <= " & ConvFecha(Fecha)
'    StrSql = StrSql & " AND (bajfec is null or bajfec >=" & ConvFecha(Fecha)
'    StrSql = StrSql & " ) AND he.htetdesde <= " & ConvFecha(Fecha)
'    StrSql = StrSql & " AND (he.htethasta is null or he.htethasta >=" & ConvFecha(Fecha)
'    StrSql = StrSql & " ) AND fases.empleado IN " & empleados
'    StrSql = StrSql & " AND he.tenro = 18 AND he.estrnro IN (" & E4 & ")"
'    If Empresas_Filtradas <> "0" Then
'        StrSql = StrSql & " AND emp.estrnro IN (" & Empresas_Filtradas & ")"
'        StrSql = StrSql & " AND emp.htetdesde <= " & ConvFecha(Fecha) & " AND ((emp.htethasta >=" & ConvFecha(Fecha) & ") OR emp.htethasta is null)"
'    End If
'    OpenRecordset StrSql, rsplantilla
'
'    Cantidad = 0
'    If Not rsplantilla.EOF Then
'        Cantidad = rsplantilla!Cantidad
'    Else
'        Cantidad = 0
'    End If
'    Flog.writeline Espacios(Tabulador * 1) & "Cantidad por contrato: " & Cantidad
'    Flog.writeline Espacios(Tabulador * 1) & StrSql
'
'    campos = " (bpronro,mes,monto,tipest)"
'    valores = "(" & NroProceso & "," & indice & "," & Cantidad & ",'14')"
'
'    StrSql = " INSERT INTO casa_matriz_det " & campos & " VALUES " & valores
'    objConn.Execute StrSql, , adExecuteNoRecords
'
'    Flog.writeline Espacios(Tabulador * 1) & "Calculando empleados sin contrato :" & Fecha
'
'    StrSql = " SELECT count(distinct empleado) as cantidad "
'    StrSql = StrSql & " FROM fases "
'    StrSql = StrSql & " INNER JOIN his_estructura as emp ON emp.ternro = fases.empleado AND emp.tenro = 10"
'    StrSql = StrSql & " INNER JOIN his_estructura as he ON he.ternro = fases.empleado "
'    StrSql = StrSql & " WHERE altfec <= " & ConvFecha(Fecha)
'    StrSql = StrSql & " AND (bajfec is null or bajfec >=" & ConvFecha(Fecha)
'    StrSql = StrSql & " ) AND he.htetdesde <= " & ConvFecha(Fecha)
'    StrSql = StrSql & " AND (he.htethasta is null or he.htethasta >=" & ConvFecha(Fecha)
'    StrSql = StrSql & " ) AND fases.empleado IN " & empleados
'    StrSql = StrSql & " AND he.tenro = 18 AND he.estrnro NOT IN (" & E4 & ")"
'    If Empresas_Filtradas <> "0" Then
'        StrSql = StrSql & " AND emp.estrnro IN (" & Empresas_Filtradas & ")"
'        StrSql = StrSql & " AND emp.htetdesde <= " & ConvFecha(Fecha) & " AND ((emp.htethasta >=" & ConvFecha(Fecha) & ") OR emp.htethasta is null)"
'    End If
'    OpenRecordset StrSql, rsplantilla
'
'    cantidad2 = 0
'    If Not rsplantilla.EOF Then
'        cantidad2 = rsplantilla!Cantidad
'    Else
'        cantidad2 = 0
'    End If
'
'    Flog.writeline Espacios(Tabulador * 1) & "Cantidad sin contrato: " & cantidad2
'    campos = " (bpronro,mes,monto,tipest)"
'    valores = "(" & NroProceso & "," & indice & "," & cantidad2 & ",'15')"
'
'     StrSql = " INSERT INTO casa_matriz_det " & campos & " VALUES " & valores
'    objConn.Execute StrSql, , adExecuteNoRecords
'
'
'    Flog.writeline Espacios(Tabulador * 1) & "Cantidad empleados con y sin contrato: " & cantidad2
'    campos = " (bpronro,mes,monto,tipest)"
'    valores = "(" & NroProceso & "," & indice & "," & Cantidad + cantidad2 & ",'16')"
'
'    StrSql = " INSERT INTO casa_matriz_det " & campos & " VALUES " & valores
'    objConn.Execute StrSql, , adExecuteNoRecords
'
'Next indice
 
Exit Sub

MError:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql

End Sub
 
Sub calcular_tipo_contrato_old()
Dim rsplantilla As New ADODB.Recordset
Dim indice As Integer
Dim Fecha As String
Dim Cantidad As Integer
Dim cantidad2 As Integer
Dim campos As String
Dim valores As String
Dim suma As Integer
Dim listaemp As String

Flog.writeline " "


For indice = 1 To 12
    
    Fecha = ultimo_dia_mes(indice, Anio)
    Flog.writeline Espacios(Tabulador * 1) & "Calculando empleados por contrato :" & Fecha
    StrSql = " SELECT distinct empleado "
    StrSql = StrSql & " FROM fases "
    StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = fases.empleado"
    StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.ternro = empleado.ternro "
    'StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro AND estructura.estrnro = " & E4
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro"
    StrSql = StrSql & " WHERE altfec <= " & ConvFecha(Fecha)
    StrSql = StrSql & " AND (bajfec is null or bajfec >=" & ConvFecha(Fecha)
    StrSql = StrSql & " ) AND htetdesde <= " & ConvFecha(Fecha)
    StrSql = StrSql & " AND (htethasta is null or htethasta >=" & ConvFecha(Fecha)
    StrSql = StrSql & " ) AND empleado in " & empleados
    StrSql = StrSql & " AND estructura.estrnro IN (" & E4 & ")"
    OpenRecordset StrSql, rsplantilla
    
    Cantidad = 0
    listaemp = "(0"
    While Not rsplantilla.EOF
        listaemp = listaemp & "," & rsplantilla!Empleado
        Cantidad = Cantidad + 1
        
        rsplantilla.MoveNext
    Wend
    listaemp = listaemp & ")"
        
    Flog.writeline Espacios(Tabulador * 1) & "Monto por contrato: " & Cantidad
    Flog.writeline Espacios(Tabulador * 1) & StrSql

    campos = " (bpronro,mes,monto,tipest)"
    
    valores = "(" & NroProceso & "," & indice & "," & Cantidad & ",'14')"
        
    
    StrSql = " INSERT INTO casa_matriz_det " & campos & " VALUES " & valores
    objConn.Execute StrSql, , adExecuteNoRecords
  
    Flog.writeline Espacios(Tabulador * 1) & "Calculando empleados sin contrato :" & Fecha
    StrSql = " SELECT distinct empleado "
    StrSql = StrSql & " FROM fases "
    StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = fases.empleado"
    StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.ternro = empleado.ternro "
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro"
    StrSql = StrSql & " WHERE altfec <= " & ConvFecha(Fecha)
    StrSql = StrSql & " AND (bajfec is null or bajfec >=" & ConvFecha(Fecha)
    StrSql = StrSql & " ) AND htetdesde <= " & ConvFecha(Fecha)
    StrSql = StrSql & " AND (htethasta is null or htethasta >=" & ConvFecha(Fecha)
    StrSql = StrSql & " ) AND empleado in " & empleados
    StrSql = StrSql & " AND empleado NOT IN " & listaemp
    StrSql = StrSql & " AND estructura.estrnro NOT IN (" & E4 & ")"
    OpenRecordset StrSql, rsplantilla
    
    cantidad2 = 0
    While Not rsplantilla.EOF
    
        cantidad2 = cantidad2 + 1
        rsplantilla.MoveNext
    
    Wend
          
    Flog.writeline Espacios(Tabulador * 1) & "Cantidad sin contrato: " & cantidad2

    campos = " (bpronro,mes,monto,tipest)"
    
    valores = "(" & NroProceso & "," & indice & "," & cantidad2 & ",'15')"
        
    
    StrSql = " INSERT INTO casa_matriz_det " & campos & " VALUES " & valores
    objConn.Execute StrSql, , adExecuteNoRecords
 
 Flog.writeline Espacios(Tabulador * 1) & "Cantidad empleados con y sin contrato: " & cantidad2

    campos = " (bpronro,mes,monto,tipest)"
    
    valores = "(" & NroProceso & "," & indice & "," & Cantidad + cantidad2 & ",'16')"
        
    
    StrSql = " INSERT INTO casa_matriz_det " & campos & " VALUES " & valores
    objConn.Execute StrSql, , adExecuteNoRecords
 
 
 Next
 
 
Exit Sub

MError:
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql

End Sub
 
 
' ____________________________________________________________
' procedimiento que levanta los parametros
' ____________________________________________________________
Sub levantarParametros(ArrParametros)

On Error GoTo ME_param


    filtro = ArrParametros(0)
    tenro1 = CInt(ArrParametros(1))
    estrnro1 = CInt(ArrParametros(2))
    tenro2 = CInt(ArrParametros(3))
    estrnro2 = CInt(ArrParametros(4))
    tenro3 = CInt(ArrParametros(5))
    estrnro3 = CInt(ArrParametros(6))
    agencia = CInt(ArrParametros(7))
    fecestr = CStr(ArrParametros(8))
    cargo = CInt(ArrParametros(9))
    'repnro = ArrParametros(10)
    monnro = CInt(ArrParametros(11))
    Monto = CDbl(ArrParametros(12))
    Anio = CInt(ArrParametros(13))
    Flog.writeline Espacios(Tabulador * 0) & "PARAMETROS"
    Flog.writeline Espacios(Tabulador * 0) & "Filtro: " & filtro
    Flog.writeline Espacios(Tabulador * 0) & "Tenro1: " & tenro1
    Flog.writeline Espacios(Tabulador * 0) & "Estrnro1: " & estrnro1
    Flog.writeline Espacios(Tabulador * 0) & "Tenro2: " & tenro2
    Flog.writeline Espacios(Tabulador * 0) & "Estrnro2: " & estrnro2
    Flog.writeline Espacios(Tabulador * 0) & "Tenro3: " & tenro3
    Flog.writeline Espacios(Tabulador * 0) & "Estrnro3: " & estrnro3
    Flog.writeline Espacios(Tabulador * 0) & "Agencia: " & agencia
    Flog.writeline Espacios(Tabulador * 0) & "Fecha p/Estruct: " & fecestr
    Flog.writeline Espacios(Tabulador * 0) & "Cargo: " & cargo
    Flog.writeline Espacios(Tabulador * 0) & "Nro Reporte: " & repnro
    Flog.writeline Espacios(Tabulador * 0) & "Tipo de Moneda: " & monnro
    Flog.writeline Espacios(Tabulador * 0) & "Valor de la Moneda: " & Monto
    Flog.writeline Espacios(Tabulador * 0) & "Año: " & Anio

Exit Sub

ME_param:
    Flog.writeline "    Error: Error en la carga de Parametros "
    
End Sub

Sub CargarConfiguracionReporte()

'Dim I 'Dim columnaActual
'Dim Nro_col 'Dim Valor As Long
Dim rs2 As New ADODB.Recordset

On Error GoTo ME_conf

Flog.writeline Espacios(Tabulador * 1) & "Buscar la configuracion del Reporte - confrep  "

StrSql = " SELECT * FROM confrep WHERE confrep.repnro= 248 "
StrSql = StrSql & " ORDER BY confnrocol "
OpenRecordset StrSql, rs2

IndiceE1 = 0
IndiceE2 = 0
IndiceE3 = 0
E4 = "0"
indiceE5 = 0
Empresas_Filtradas = "0"
contratos_eventuales = "0"

If rs2.EOF Then
    Flog.writeline Espacios(Tabulador * 0) & "Error. Se debe configurar el confrep. Nro de confrep:" & repnro
    Exit Sub
Else

    Do While Not rs2.EOF
       
       Select Case Trim(rs2!conftipo)
            Case "E1"
                E1(IndiceE1) = rs2!confval
                IndiceE1 = IndiceE1 + 1
            Case "E2"
                E2(IndiceE2) = rs2!confval
                IndiceE2 = IndiceE2 + 1
            Case "E3"
                E3(IndiceE3) = rs2!confval
                IndiceE3 = IndiceE3 + 1
            Case "AC1"
                AC1 = rs2!confval
            Case "AC2"
                AC2 = rs2!confval
            Case "E4"
                If E4 = "0" Then
                    E4 = rs2!confval
                Else
                    E4 = E4 & "," & rs2!confval
                End If
            Case "FIL":
                Empresas_Filtradas = Empresas_Filtradas & "," & rs2!confval
            Case "CE":
                E5(indiceE5) = rs2!confval
                indiceE5 = indiceE5 + 1
                contratos_eventuales = contratos_eventuales & "," & rs2!confval
       End Select
       
       rs2.MoveNext
    Loop
End If


    
Exit Sub

ME_conf:
    ' Flog.Writeline "    Error - Empleado: " & Empleado
    Flog.writeline "    Error: " & Err.Description
    Flog.writeline "    SQL Ejecutado: " & StrSql

End Sub

'---------------------------------------------------------------------------------------------------
' procedimiento que busca los empleados que cumplen con lo seleccionado en el filtro
'---------------------------------------------------------------------------------------------------

Sub filtro_empleados(ByRef StrSql As String, ByVal Fecha As Date)

Dim StrAgencia As String
Dim StrSelect As String
Dim strjoin As String
Dim StrOrder As String
Dim fecdes As String
Dim fechas As String
Dim rsfiltro As New ADODB.Recordset

On Error GoTo ME_armarsql

StrSql = ""
StrSelect = ""
strjoin = ""
StrOrder = ""

' Busco todos los empleados que cumplen con el filtro

StrAgencia = "" ' cuando queremos todos los empleados

If agencia = "-1" Then
    StrAgencia = " AND empleado.ternro NOT IN (SELECT ternro FROM his_estructura agencia "
    StrAgencia = StrAgencia & " WHERE agencia.tenro=28 AND (agencia.htetdesde<=" & ConvFecha(Fecha)
    StrAgencia = StrAgencia & "     AND (agencia.htethasta IS NULL OR agencia.htethasta>=" & ConvFecha(Fecha) & ")) )"
Else
    If agencia = "-2" Then
        StrAgencia = " AND empleado.ternro IN (SELECT ternro FROM his_estructura agencia "
        StrAgencia = StrAgencia & " WHERE agencia.tenro=28 AND (agencia.htetdesde<=" & ConvFecha(Fecha)
        StrAgencia = StrAgencia & " AND (agencia.htethasta IS NULL OR agencia.htethasta>=" & ConvFecha(Fecha) & ")) )"
    Else
        If agencia <> "0" Then 'este caso se da cuando selecionamos una agencia determinada
            StrAgencia = " AND empleado.ternro in (SELECT ternro from his_estructura agencia "
            StrAgencia = StrAgencia & " WHERE agencia.tenro=28 and agencia.estrnro=" & agencia
            StrAgencia = StrAgencia & "  AND (agencia.htetdesde<=" & ConvFecha(Fecha)
            StrAgencia = StrAgencia & "  AND (agencia.htethasta IS NULL OR agencia.htethasta>=" & ConvFecha(Fecha) & ")) )"
        End If
    End If
End If
 
 
 
If tenro1 <> 0 Then  ' Cuando solo selecionamos el primer nivel
    
    
    strjoin = strjoin & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro  AND estact1.tenro  = " & tenro1
    strjoin = strjoin & "  AND (estact1.htetdesde<=" & ConvFecha(Fecha) & " AND (estact1.htethasta IS NULL OR estact1.htethasta>=" & ConvFecha(Fecha) & "))"
    If estrnro1 <> 0 Then
        strjoin = strjoin & " AND estact1.estrnro =" & estrnro1
    End If
    strjoin = strjoin & " INNER JOIN estructura estructura1 ON estructura1.estrnro=estact1.estrnro "
    
    If StrOrder <> "" Then
        StrOrder = StrOrder & ", "
    End If
    StrOrder = StrOrder & " tenro1, estrnro1 "

End If

If tenro2 <> 0 Then  ' ocurre cuando se selecciono hasta el segundo nivel

    
    strjoin = strjoin & " INNER JOIN his_estructura estact2 ON empleado.ternro = estact2.ternro  AND estact2.tenro  = " & tenro2
    strjoin = strjoin & " AND (estact2.htetdesde<=" & ConvFecha(Fecha) & " AND (estact2.htethasta IS NULL OR estact2.htethasta>=" & ConvFecha(Fecha) & "))"
    If estrnro2 <> 0 Then
        strjoin = strjoin & " AND estact2.estrnro =" & estrnro2
    End If
    strjoin = strjoin & " INNER JOIN estructura estructura2 ON estructura2.estrnro=estact2.estrnro "
    
    If StrOrder <> "" Then
        StrOrder = StrOrder & ", "
    End If
    StrOrder = StrOrder & " tenro2, estrnro2 "

End If

If tenro3 <> 0 Then  ' esto ocurre solo cuando se seleccionan los tres niveles


    strjoin = strjoin & " INNER JOIN his_estructura estact3 ON empleado.ternro = estact3.ternro  AND estact3.tenro =" & tenro3
    strjoin = strjoin & "   AND (estact3.htetdesde<=" & ConvFecha(Fecha) & " AND (estact3.htethasta IS NULL OR  estact3.htethasta>=" & ConvFecha(Fecha) & "))"
    If estrnro3 <> 0 Then 'cuando se le asigna un valor al nivel 3
        strjoin = strjoin & " AND estact3.estrnro =" & estrnro3
    End If
    strjoin = strjoin & " INNER JOIN estructura estructura3 ON estructura3.estrnro=estact3.estrnro "

    If StrOrder <> "" Then
        StrOrder = StrOrder & ", "
    End If
    StrOrder = StrOrder & " tenro3, estrnro3 "
    
End If

                      
StrSql = " SELECT DISTINCT empleado.ternro  "   '  empleado.empest, tercero.tersex,
StrSql = StrSql & " FROM empleado "
StrSql = StrSql & " INNER JOIN tercero ON empleado.ternro = tercero.ternro "
StrSql = StrSql & strjoin
StrSql = StrSql & " WHERE " & filtro & StrAgencia

OpenRecordset StrSql, rsfiltro

empleados = "(0"

While Not rsfiltro.EOF
    empleados = empleados & "," & rsfiltro!ternro
    rsfiltro.MoveNext
Wend
 
empleados = empleados & ")"


Exit Sub


ME_armarsql:
    Flog.writeline " Error: Armar consulta del Filtro.- " & Err.Description
    Flog.writeline " Búsqueda de empleados filtrados: " & StrSql
    
    
End Sub

Function primer_dia_mes(mes As Integer, Anio As Integer) As Date
Dim aux As String
     
    primer_dia_mes = C_Date("01/" & mes & "/" & Anio)
    
End Function



Function ultimo_dia_mes(mes As Integer, Anio As Integer) As Date

Dim mes_sgt As Integer
Dim anio_sgt As Integer

    If mes = 12 Then
        mes_sgt = 1
        anio_sgt = Anio + 1
    Else
        mes_sgt = mes + 1
        anio_sgt = Anio
    End If
    
    ultimo_dia_mes = DateAdd("d", -1, primer_dia_mes(mes_sgt, anio_sgt))
    
End Function


Sub actualizar_progreso(Progreso As Integer)

TiempoAcumulado = GetTickCount

StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'"
StrSql = StrSql & " WHERE bpronro = " & NroProceso
objConn.Execute StrSql, , adExecuteNoRecords
    
End Sub

