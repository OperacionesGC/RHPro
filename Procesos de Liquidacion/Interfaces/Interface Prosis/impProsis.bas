Attribute VB_Name = "impProsis"
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
Global fechaDesde
Global fechaHasta
Global concTipoDia(100)
Global paramTipoDia(100)
Global concReintegros
Global paramReintegros
Global concHN
Global paramHN
Global concHB
Global paramHB
Global concHE
Global paramHE
Global concHEB
Global paramHEB
Global topeHEB
Global empresaCod

Global ObtReintegros As Boolean
Global ObtLicencias As Boolean
Global ObtHoras As Boolean
Global GenerarNov As Boolean

Dim topeHorasLic
Dim basicoHoras

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
Dim objRs3 As New ADODB.Recordset
Dim tipoDepuracion
Dim historico As Boolean
Dim param
Dim i
Dim totalEmpleados
Dim cantRegistros
Dim PID As String
Dim ArrParametros
Dim parametros As String

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
    
    On Error GoTo CE
    
    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas
    
    TiempoInicialProceso = GetTickCount
    OpenConnection strconexion, objConn
    
    HuboErrores = False
    
    Nombre_Arch = PathFLog & "InterfaceProsis" & "-" & NroProceso & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    Flog.Writeline "Inicio Interface Prosis : " & Now
    Flog.Writeline "Cambio el estado del proceso a Procesando"
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.Writeline "PID = " & PID
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Flog.Writeline "Obtengo los datos del proceso"
    
    TiempoAcumulado = GetTickCount
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, objRs
    
    If Not objRs.EOF Then
    
       'empresaCod = "POP"
       
       'Obtengo los parametros del proceso
       parametros = objRs!bprcparam
       ArrParametros = Split(parametros, "@")
       
       'Obtengo la fecha desde
       fechaDesde = CDate(objRs!bprcfecdesde)
       
       'Obtengo la fecha hasta
       fechaHasta = CDate(objRs!bprcfechasta)
       
       'Obtengo el codigo de la empresa a buscar
       'empresaCod = ArrParametros(0)
       
       ObtReintegros = CBool(ArrParametros(0))
       ObtLicencias = CBool(ArrParametros(1))
       ObtHoras = CBool(ArrParametros(2))
       GenerarNov = CBool(ArrParametros(3))
       
       'EMPIEZA EL PROCESO
       
       'Busco en el confrep las asociaciones para generar los parametros
       StrSql = " SELECT * FROM confrep "
       StrSql = StrSql & " WHERE repnro = 130 "
      
       OpenRecordset StrSql, objRs2
       
       If objRs2.EOF Then
          Flog.Writeline "No esta configurado el ConfRep"
          Exit Sub
       End If
       
       Flog.Writeline "Obtengo los datos del confrep"
       
       For i = 1 To 100
          concTipoDia(i) = 0
          paramTipoDia(i) = 0
       Next
       
       concReintegros = 0
       paramReintegros = 0
       topeHorasLic = 260
       basicoHoras = 192
       concHN = 0
       paramHN = 0
       concHB = 0
       paramHB = 0
       concHE = 0
       paramHE = 0
       concHEB = 0
       paramHEB = 0
       topeHEB = 0

       Do Until objRs2.EOF
          'Concepto de refrigerios
          If CLng(objRs2!confnrocol) = 1 Then
             Flog.Writeline "Confrep columna " & objRs2!confnrocol
             concReintegros = buscarConcepto(objRs3, objRs2!confval2)
          End If
          
          'Parametro de refrigerios
          If CLng(objRs2!confnrocol) = 2 Then
             Flog.Writeline "Confrep columna " & objRs2!confnrocol
             paramReintegros = objRs2!confval
          End If
          
          'Valor de topeHorasLic
          If CLng(objRs2!confnrocol) = 3 Then
             Flog.Writeline "Confrep columna " & objRs2!confnrocol
             topeHorasLic = objRs2!confval
          End If
          
          'Valor de basico de horas
          If CLng(objRs2!confnrocol) = 4 Then
             Flog.Writeline "Confrep columna " & objRs2!confnrocol
             basicoHoras = objRs2!confval
          End If
          
          'Concepto de horas normales
          If CLng(objRs2!confnrocol) = 5 Then
             Flog.Writeline "Confrep columna " & objRs2!confnrocol
             concHN = buscarConcepto(objRs3, objRs2!confval2)
          End If
          
          'Parametro de horas normales
          If CLng(objRs2!confnrocol) = 6 Then
             Flog.Writeline "Confrep columna " & objRs2!confnrocol
             paramHN = objRs2!confval
          End If
          
          'Concepto de horas B
          If CLng(objRs2!confnrocol) = 7 Then
             Flog.Writeline "Confrep columna " & objRs2!confnrocol
             concHB = buscarConcepto(objRs3, objRs2!confval2)
          End If
          
          'Parametro de horas B
          If CLng(objRs2!confnrocol) = 8 Then
             Flog.Writeline "Confrep columna " & objRs2!confnrocol
             paramHB = objRs2!confval
          End If
          
          'Concepto de horas extras
          If CLng(objRs2!confnrocol) = 9 Then
             Flog.Writeline "Confrep columna " & objRs2!confnrocol
             concHE = buscarConcepto(objRs3, objRs2!confval2)
          End If
          
          'Parametro de horas extras
          If CLng(objRs2!confnrocol) = 10 Then
             Flog.Writeline "Confrep columna " & objRs2!confnrocol
             paramHE = objRs2!confval
          End If
          
          'Concepto de horas extras B
          If CLng(objRs2!confnrocol) = 11 Then
             Flog.Writeline "Confrep columna " & objRs2!confnrocol
             concHEB = buscarConcepto(objRs3, objRs2!confval2)
          End If
          
          'Parametro de horas extras B
          If CLng(objRs2!confnrocol) = 12 Then
             Flog.Writeline "Confrep columna " & objRs2!confnrocol
             paramHEB = objRs2!confval
          End If
          
          'Parametro de tope horas extras B
          If CLng(objRs2!confnrocol) = 13 Then
             Flog.Writeline "Confrep columna " & objRs2!confnrocol
             topeHEB = objRs2!confval
          End If
          
          
          'Concepto por tipo de hora
          If CStr(objRs2!conftipo) = "TD" Then
             Flog.Writeline "Confrep columna " & objRs2!confnrocol
             concTipoDia(CLng(objRs2!confval)) = buscarConcepto(objRs3, objRs2!confval2)
          End If
          
          'Parametro por tipo de hora
          If CStr(objRs2!conftipo) = "TDP" Then
             Flog.Writeline "Confrep columna " & objRs2!confnrocol
             paramTipoDia(CLng(objRs2!confval)) = CLng(objRs2!confval2)
          End If
       
          objRs2.MoveNext
       Loop
       
       objRs2.Close

       If ObtReintegros Then
          'Generar Reintegros
          Call generarReintegros
       End If
       
       If ObtLicencias Then
          'Borrar las licencias en el rango de fechas
          Call borrarDatos(fechaDesde, fechaHasta, 1)

          'Guardo en una tabla las licencias
          Call guardarLicencias
       End If
       
       If ObtHoras Then
          'Borrar las horas pactadas en el rango de fechas
          Call borrarDatos(fechaDesde, fechaHasta, 2)
       
          'Guardo en una tabla las horas pactadas
          Call guardarHorasPactadas
       End If
       
       If GenerarNov Then
          'Genero las novedades
          Call generarNovedades
       End If
    
    Else
        Exit Sub
    End If
   
    'Actualizo el estado del proceso
    If Not HuboErrores Then
       StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
    Else
       StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProceso
    End If
    
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Flog.Writeline "Fin :" & Now
    Flog.Close

    Exit Sub
    
CE:
    HuboErrores = True
    Flog.Writeline " Error: " & Err.Description & Now

End Sub

Function controlNull(Str)
  If Trim(Str) = "" Then
     controlNull = "null"
  Else
     controlNull = "'" & Str & "'"
  End If
End Function

Function sinDatos(Str)

  If IsNull(Str) Then
     sinDatos = True
  Else
     If Trim(Str) = "" Then
        sinDatos = True
     Else
        sinDatos = False
     End If
  End If

End Function


Function buscarConcepto(ByRef rsConcepto As ADODB.Recordset, ByVal concCod)

On Error GoTo MError

    'Busco el concnro del concepto
    StrSql = "SELECT concnro FROM concepto WHERE conccod = '" & concCod & "'"
    
    Flog.Writeline "Buscando Concepto: " & StrSql
    
    OpenRecordset StrSql, rsConcepto
    
    If rsConcepto.EOF Then
       buscarConcepto = 0
    Else
       buscarConcepto = rsConcepto!concnro
    End If
    
    rsConcepto.Close
    
    Exit Function
    
MError:
    Flog.Writeline "Error en buscarEmpleado: " & Err.Description
    HuboErrores = True
    
End Function

Function buscarEmpleado(ByRef rsEmpleado As ADODB.Recordset, ByVal empleg)

On Error GoTo MError

    Flog.Writeline "Buscando si se encuentra al empleado " & empleg

    'Busco el ternro del empleado
    StrSql = "SELECT ternro FROM empleado WHERE empleg = " & empleg
    
    Flog.Writeline "Buscando Empleado: " & StrSql
    
    OpenRecordset StrSql, rsEmpleado
    
    If rsEmpleado.EOF Then
       buscarEmpleado = 0
    Else
       buscarEmpleado = rsEmpleado!ternro
    End If
    
    rsEmpleado.Close
    
    Exit Function
    
MError:
    Flog.Writeline "Error en buscarEmpleado: " & Err.Description
    HuboErrores = True

End Function

Function buscarTipoLicencia(ByRef rsTipoLic As ADODB.Recordset, ByVal codigo)

On Error GoTo MError

    Flog.Writeline "Buscando el tipo de licencia:" & codigo

    'Busco el tdnro de la licencia
    StrSql = "SELECT tdnro FROM tipdia WHERE tdsigla LIKE '" & Mid(codigo, 1, 2) & "%'"
    
    Flog.Writeline "Buscando Tipo Licencia: " & StrSql
    
    OpenRecordset StrSql, rsTipoLic
    
    If rsTipoLic.EOF Then
       buscarTipoLicencia = 0
    Else
       buscarTipoLicencia = rsTipoLic!tdnro
    End If
    
    rsTipoLic.Close
    
    Exit Function
    
MError:
    Flog.Writeline "Error en buscarTipoLicencia: " & Err.Description
    HuboErrores = True

End Function
                 
                  
Sub generarReintegros()

Dim wsClient As New MSSOAPLib30.SoapClient30
Dim auxi As String
Dim objXMLdsPubs As IXMLDOMSelection

Dim rsEmpleado As New ADODB.Recordset
Dim rsConsulta As New ADODB.Recordset
Dim ternro
Dim total
Dim i
Dim totalReg
    
On Error GoTo MError

    Flog.Writeline "Buscando si se encuentran reintegros"
    
    'Actualizo el progreso
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = 0 " & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
             " WHERE bpronro = " & NroProceso
             
    objConn.Execute StrSql, , adExecuteNoRecords
    
    ' Conecxion al Web Service
    Flog.Writeline "Conectandose al WebService"
    wsClient.MSSoapInit "http://192.168.106.17/WSSittMGOExt/rhpro.asmx?WSDL", "RHPro", "RHProSoap"

    ' Parametros auxiliares del Servicio
    'auxi = "<Params>" & _
    '"<Empresa>" & empresaCod & "</Empresa>" & _
    '"<FechaDesde>" & fechaToXML(fechaDesde) & "</FechaDesde>" & _
    '"<FechaHasta>" & fechaToXML(fechaHasta) & "</FechaHasta>" & _
    '"</Params>"
    
    auxi = "<Params>" & _
    "<Empresa></Empresa>" & _
    "<FechaDesde>" & fechaToXML(fechaDesde) & "</FechaDesde>" & _
    "<FechaHasta>" & fechaToXML(fechaHasta) & "</FechaHasta>" & _
    "</Params>"

    ' Executa el metodo execute con la accion SittInterface.Reintegros
    ' y parametros de la variable auxi
    Flog.Writeline "Ejecutando la accion del WebService"
    Set objXMLdsPubs = wsClient.Execute("SittInterface.Reintegros", auxi)

    ' Verifica si el resultado del envio de datos fue OK
    If objXMLdsPubs.Item(0).selectNodes("Result").Item(0).selectSingleNode("IsOk").Text = "true" Then

       'Debug.Print "Empresa;Dia;ReintegroNro;Legajo;FaltantesRendicion;Total"
    
       ' Recorre todo el XML correspondiente solo a los Datos
       totalReg = objXMLdsPubs.Item(0).selectNodes("Reintegros").length
       For i = 0 To objXMLdsPubs.Item(0).selectNodes("Reintegros").length - 1
          If objXMLdsPubs.Item(0).selectNodes("Reintegros").Item(i).selectSingleNode("Empresa").Text = "FBU" Then
               
               Flog.Writeline "Obteniendo Reintegro para el empleado de Sitt: " & objXMLdsPubs.Item(0).selectNodes("Reintegros").Item(i).selectSingleNode("Legajo").Text
               ternro = buscarEmpleado(rsEmpleado, objXMLdsPubs.Item(0).selectNodes("Reintegros").Item(i).selectSingleNode("Legajo").Text)
    
               If ternro <> 0 Then
                  total = objXMLdsPubs.Item(0).selectNodes("Reintegros").Item(i).selectSingleNode("Total").Text
                  
                  StrSql = " SELECT * FROM novemp WHERE empleado = " & ternro
                  StrSql = StrSql & " AND concnro = " & concReintegros
                  StrSql = StrSql & " AND tpanro = " & paramReintegros
                  
                  OpenRecordset StrSql, rsConsulta
                  
                  If rsConsulta.EOF Then
                  
                     StrSql = "INSERT INTO novemp "
                     StrSql = StrSql & "(empleado, concnro, tpanro, nevalor, nevigencia, nedesde, nehasta, neretro, nepliqdesde, nepliqhasta)"
                     StrSql = StrSql & " values (" & ternro & ", " & concReintegros & ", " & paramReintegros & ", " & numberForSQL(total) & ", "
                     StrSql = StrSql & "0,null,null,null,null,null) "
                    
                  Else
                  
                     StrSql = "UPDATE novemp "
                     StrSql = StrSql & " SET nevalor=" & numberForSQL(total)
                     StrSql = StrSql & " WHERE nenro = " & rsConsulta!nenro
                  
                  End If
                  
                  Flog.Writeline "Insertando reintegros para el tercero " & ternro
                  
                  rsConsulta.Close
              
                  objConn.Execute StrSql, , adExecuteNoRecords
               
               End If
               
               'auxi = objXMLdsPubs.Item(0).selectNodes("Reintegros").Item(I).selectSingleNode("Empresa").Text & ";"
               'auxi = auxi & objXMLdsPubs.Item(0).selectNodes("Reintegros").Item(I).selectSingleNode("Dia").Text & ";"
               'auxi = auxi & objXMLdsPubs.Item(0).selectNodes("Reintegros").Item(I).selectSingleNode("ReintegoNumero").Text & ";"
               'auxi = auxi & objXMLdsPubs.Item(0).selectNodes("Reintegros").Item(I).selectSingleNode("Legajo").Text & ";"
               'auxi = auxi & objXMLdsPubs.Item(0).selectNodes("Reintegros").Item(I).selectSingleNode("FaltantesRendicion").Text & ";"
               'auxi = auxi & objXMLdsPubs.Item(0).selectNodes("Reintegros").Item(I).selectSingleNode("Total").Text
               'Debug.Print auxi
               
          End If
               
          'Actualizo el estado del proceso
          TiempoAcumulado = GetTickCount
          StrSql = "UPDATE batch_proceso SET bprcprogreso = " & (((i + 1) * 25) / totalReg) & _
                ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                " WHERE bpronro = " & NroProceso
                         
          objConn.Execute StrSql, , adExecuteNoRecords
        
          Next i
    Else
       Flog.Writeline "No se encontraron datos de reintegros"
       Flog.Writeline "Error: " & objXMLdsPubs.Item(0).xml
    End If
    
    Set objXMLdsPubs = Nothing
    Set wsClient = Nothing
    Set rsEmpleado = Nothing
    Set rsConsulta = Nothing
    
    'Actualizo el estado del proceso
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso =  25 " & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
             " WHERE bpronro = " & NroProceso
             
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Flog.Writeline "Se termino de buscar los reintegros"

    Exit Sub
    
MError:
    Flog.Writeline "Error generarReintegros: " & Err.Description
    HuboErrores = True
    Exit Sub

End Sub


Sub guardarLicencias()

Dim rsEmpleado As New ADODB.Recordset
Dim rsTipoLic As New ADODB.Recordset
Dim rsConsulta As New ADODB.Recordset
Dim ternro
Dim total
Dim tdnro
Dim totalReg

Dim wsClient As New MSSOAPLib30.SoapClient30
Dim auxi As String
Dim objXMLdsPubs As IXMLDOMSelection
Dim i

On Error GoTo MError

    Flog.Writeline "Buscando si se encuentran licencias"
    
    'Actualizo el progreso
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = 25 " & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
             " WHERE bpronro = " & NroProceso
             
    objConn.Execute StrSql, , adExecuteNoRecords
    
    ' Conecxion al Web Service
    Flog.Writeline "Conectandose al WebService"
    wsClient.MSSoapInit "http://192.168.106.17/WSSittMGOExt/rhpro.asmx?WSDL", "RHPro", "RHProSoap"

    ' Parametros auxiliares del Servicio
    'auxi = "<Params>" & _
    '"<Empresa>" & empresaCod & "</Empresa>" & _
    '"<FechaDesde>" & fechaToXML(fechaDesde) & "</FechaDesde>" & _
    '"<FechaHasta>" & fechaToXML(fechaHasta) & "</FechaHasta>" & _
    '"</Params>"
    
    auxi = "<Params>" & _
    "<Empresa></Empresa>" & _
    "<FechaDesde>" & fechaToXML(fechaDesde) & "</FechaDesde>" & _
    "<FechaHasta>" & fechaToXML(fechaHasta) & "</FechaHasta>" & _
    "</Params>"
    
    ' Executa el metodo execute con la accion SittInterface.LicenciasEstados
    ' y parametros de la variable auxi
    Flog.Writeline "Ejecutando la accion del WebService"
    Set objXMLdsPubs = wsClient.Execute("SittInterface.LicenciasEstados", auxi)

    ' Verifica si el resultado del envio de datos fue OK
    If objXMLdsPubs.Item(0).selectNodes("Result").Item(0).selectSingleNode("IsOk").Text = "true" Then

       'Debug.Print "Empresa;Legajo;Estado;Fecha"
    
       ' Recorre todo el XML correspondiente solo a los Datos
       totalReg = objXMLdsPubs.Item(0).selectNodes("LicenciasEstados").length
       For i = 0 To objXMLdsPubs.Item(0).selectNodes("LicenciasEstados").length - 1
           
         If objXMLdsPubs.Item(0).selectNodes("LicenciasEstados").Item(i).selectSingleNode("Empresa").Text = "FBU" Then
             ternro = buscarEmpleado(rsEmpleado, objXMLdsPubs.Item(0).selectNodes("LicenciasEstados").Item(i).selectSingleNode("Legajo").Text)
             
             If ternro <> 0 Then
                
                tdnro = buscarTipoLicencia(rsTipoLic, objXMLdsPubs.Item(0).selectNodes("LicenciasEstados").Item(i).selectSingleNode("Estado").Text)
                
                If tdnro <> 0 Then
                   StrSql = "INSERT INTO prosis_licencias "
                   StrSql = StrSql & "(bpronro,ternro,tdnro,fecha)"
                   StrSql = StrSql & " values (" & NroProceso & "," & ternro & ", " & tdnro & ", " & ConvFecha(objXMLdsPubs.Item(0).selectNodes("LicenciasEstados").Item(i).selectSingleNode("Fecha").Text) & ")"
                   
                   Flog.Writeline "Insertando licencias para el tercero " & ternro
            
                   objConn.Execute StrSql, , adExecuteNoRecords
                
                End If
             
             End If
             
             'auxi = objXMLdsPubs.Item(0).selectNodes("LicenciasEstados").Item(I).selectSingleNode("Empresa").Text & ";"
             'auxi = auxi & objXMLdsPubs.Item(0).selectNodes("LicenciasEstados").Item(I).selectSingleNode("Legajo").Text & ";"
             'auxi = auxi & objXMLdsPubs.Item(0).selectNodes("LicenciasEstados").Item(I).selectSingleNode("Estado").Text & ";"
             'auxi = auxi & objXMLdsPubs.Item(0).selectNodes("LicenciasEstados").Item(I).selectSingleNode("Fecha").Text
             'Debug.Print auxi
             
         End If
         
        'Actualizo el estado del proceso
         TiempoAcumulado = GetTickCount
         StrSql = "UPDATE batch_proceso SET bprcprogreso = " & ((((i + 1) * 25) / totalReg) + 25) & _
                  ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                  " WHERE bpronro = " & NroProceso
                     
         objConn.Execute StrSql, , adExecuteNoRecords
           
       Next i
    Else
       Flog.Writeline "No se encontraron datos de licencias"
       Flog.Writeline "Error: " & objXMLdsPubs.Item(0).xml
       'Flog.Writeline "WebService Error Code : " & objXMLdsPubs.Item(0).selectNodes("Errors").Item(0).selectSingleNode("Code").Text
       'Flog.Writeline "WebService Error Description : " & objXMLdsPubs.Item(0).selectNodes("Errors").Item(0).selectSingleNode("Description").Text
    End If
    
    Set objXMLdsPubs = Nothing
    Set wsClient = Nothing
    Set rsEmpleado = Nothing
    Set rsTipoLic = Nothing
    Set rsConsulta = Nothing
    
    'Actualizo el estado del proceso
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso =  50 " & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
             " WHERE bpronro = " & NroProceso
             
    objConn.Execute StrSql, , adExecuteNoRecords
    
    
    Flog.Writeline "Se termino de buscar las licencias"
    
    Exit Sub
    
MError:
    Flog.Writeline "Error en guardarLicencias: " & Err.Description
    HuboErrores = True
    Exit Sub

End Sub


Sub guardarHorasPactadas()

Dim rsEmpleado As New ADODB.Recordset
Dim rsTipoLic As New ADODB.Recordset
Dim rsConsulta As New ADODB.Recordset
Dim ternro
Dim total
Dim tdnro
Dim i
Dim totalReg
Dim pepe


On Error GoTo MError

    Dim wsClient As New MSSOAPLib30.SoapClient30
    Dim auxi As String
    Dim objXMLdsPubs As IXMLDOMSelection
    
    Flog.Writeline "Buscando si se encuentran horas pactadas"
    
    'Actualizo el progreso
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = 50 " & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
             " WHERE bpronro = " & NroProceso
             
    objConn.Execute StrSql, , adExecuteNoRecords
    
    ' Conecxion al Web Service
    Flog.Writeline "Conectandose al WebService"
    wsClient.MSSoapInit "http://192.168.106.17/WSSittMGOExt/rhpro.asmx?WSDL", "RHPro", "RHProSoap"

    ' Parametros auxiliares del Servicio
    'auxi = "<Params>" & _
    '"<Empresa>" & empresaCod & "</Empresa>" & _
    '"<FechaDesde>" & fechaToXML(fechaDesde) & "</FechaDesde>" & _
    '"<FechaHasta>" & fechaToXML(fechaHasta) & "</FechaHasta>" & _
    '"</Params>"
    
    auxi = "<Params>" & _
    "<Empresa></Empresa>" & _
    "<FechaDesde>" & fechaToXML(fechaDesde) & "</FechaDesde>" & _
    "<FechaHasta>" & fechaToXML(fechaHasta) & "</FechaHasta>" & _
    "</Params>"

    ' Executa el metodo execute con la accion SittInterface.HorasPactadas
    ' y parametros de la variable auxi
    Flog.Writeline "Ejecutando la accion del WebService"
    Set objXMLdsPubs = wsClient.Execute("SittInterface.HorasPactadas", auxi)

    ' Verifica si el resultado del envio de datos fue OK
    If objXMLdsPubs.Item(0).selectNodes("Result").Item(0).selectSingleNode("IsOk").Text = "true" Then

       'Debug.Print "Empresa;Dia;Legajo;HorasPactadas"
    
       ' Recorre todo el XML correspondiente solo a los Datos
       totalReg = objXMLdsPubs.Item(0).selectNodes("HorasPactadas").length
       For i = 0 To objXMLdsPubs.Item(0).selectNodes("HorasPactadas").length - 1
           'If CInt(objXMLdsPubs.Item(0).selectNodes("HorasPactadas").Item(i).selectSingleNode("Legajo").Text) = CInt(422) Then
           '    pepe = 1
           'End If
          If objXMLdsPubs.Item(0).selectNodes("HorasPactadas").Item(i).selectSingleNode("EmpresaLegajo").Text = "FBU" Then
               ternro = buscarEmpleado(rsEmpleado, objXMLdsPubs.Item(0).selectNodes("HorasPactadas").Item(i).selectSingleNode("Legajo").Text)
               
               If ternro <> 0 Then
                   total = objXMLdsPubs.Item(0).selectNodes("HorasPactadas").Item(i).selectSingleNode("HorasPactadas").Text
                     
                   StrSql = "INSERT INTO prosis_horas_pact "
                   StrSql = StrSql & "(bpronro,ternro,horas,fecha)"
                   StrSql = StrSql & " values (" & NroProceso & "," & ternro & ", " & numberForSQL(total) & ", " & ConvFecha(objXMLdsPubs.Item(0).selectNodes("HorasPactadas").Item(i).selectSingleNode("Dia").Text) & ")"
                   
                   Flog.Writeline "Guardando horas pactadas para el tercero " & ternro
            
                   objConn.Execute StrSql, , adExecuteNoRecords
                  
               End If
          End If
          
           'auxi = objXMLdsPubs.Item(0).selectNodes("HorasPactadas").Item(I).selectSingleNode("Empresa").Text & ";"
           'auxi = auxi & objXMLdsPubs.Item(0).selectNodes("HorasPactadas").Item(I).selectSingleNode("Dia").Text & ";"
           'auxi = auxi & objXMLdsPubs.Item(0).selectNodes("HorasPactadas").Item(I).selectSingleNode("Legajo").Text & ";"
           'auxi = auxi & objXMLdsPubs.Item(0).selectNodes("HorasPactadas").Item(I).selectSingleNode("HorasPactadas").Text
           'Debug.Print auxi
           
            'Actualizo el estado del proceso
          TiempoAcumulado = GetTickCount
          StrSql = "UPDATE batch_proceso SET bprcprogreso = " & ((((i + 1) * 25) / totalReg) + 50) & _
                     ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                     " WHERE bpronro = " & NroProceso
                     
          objConn.Execute StrSql, , adExecuteNoRecords

       Next i
    Else
        Flog.Writeline "No se encontraron datos de horas pactadas"
        Flog.Writeline "Error: " & objXMLdsPubs.Item(0).xml
       'Flog.Writeline "WebService Error Code : " & objXMLdsPubs.Item(0).selectNodes("Errors").Item(0).selectSingleNode("Code").Text
       'Flog.Writeline "WebService Error Description : " & objXMLdsPubs.Item(0).selectNodes("Errors").Item(0).selectSingleNode("Description").Text
    End If
    
    Set objXMLdsPubs = Nothing
    Set wsClient = Nothing
    Set rsEmpleado = Nothing
    Set rsTipoLic = Nothing
    Set rsConsulta = Nothing
    
    'Actualizo el estado del proceso
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso =  75 " & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
             " WHERE bpronro = " & NroProceso
             
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Flog.Writeline "Se termino de buscar las horas pactadas"
    
    Exit Sub

MError:
    Flog.Writeline "Error en guardarHorasPactadas: " & Err.Description
    HuboErrores = True
    Exit Sub

End Sub

Sub borrarDatos(fDesde, fHasta, tipo As Integer)

On Error GoTo MError

  If tipo = 1 Then
    'Borro las licencias
    StrSql = "DELETE FROM prosis_licencias "
    StrSql = StrSql & " WHERE fecha  >= " & ConvFecha(fDesde)
    StrSql = StrSql & "   AND fecha  <= " & ConvFecha(fHasta)
    
    objConn.Execute StrSql, , adExecuteNoRecords
  End If
        
  If tipo = 2 Then
    'Borro las horas pactadas
    StrSql = "DELETE FROM prosis_horas_pact "
    StrSql = StrSql & " WHERE fecha  >= " & ConvFecha(fDesde)
    StrSql = StrSql & "   AND fecha  <= " & ConvFecha(fHasta)
          
    objConn.Execute StrSql, , adExecuteNoRecords
  End If
  
  Exit Sub
  
MError:
    Flog.Writeline "Error al borrar los datos, Error: " & Err.Description
    Exit Sub
   
End Sub


Function numberForSQL(Str)
   
  numberForSQL = Replace(Str, ",", ".")

End Function

Function fechaToXML(Fecha)
  
  Fecha = CDate(Fecha)
  
  fechaToXML = Month(Fecha) & "/" & Day(Fecha) & "/" & Year(Fecha)

End Function

Sub generarNovedades()

Dim rsEmpleados As New ADODB.Recordset
Dim rsLic As New ADODB.Recordset
Dim rsTrab As New ADODB.Recordset
Dim tieneLic As Boolean
Dim tieneLicVac As Boolean
Dim tieneLicOtras As Boolean
Dim horasLic
Dim horasLicOtras
Dim horasLicVac
Dim horasTrab
Dim Valor
Dim valorAux
Dim totalReg
Dim i
Dim nuevoTope

On Error GoTo MError

   Flog.Writeline "Empezando a generar las novedades para los empleados"
   
    'Actualizo el progreso
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = 75 " & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
             " WHERE bpronro = " & NroProceso
             
    objConn.Execute StrSql, , adExecuteNoRecords

   'Busco que empleados tiene horas pactadas
   'StrSql = "SELECT DISTINCT ternro FROM prosis_horas_pact "
   'StrSql = StrSql & " WHERE fecha >= " & ConvFecha(fechaDesde)
   'StrSql = StrSql & " AND fecha <= " & ConvFecha(fechaHasta)
   
   
   StrSql = " SELECT DISTINCT ternro "
   StrSql = StrSql & " From Empleado "
   StrSql = StrSql & " WHERE empleado.ternro IN (SELECT prosis_horas_pact.ternro FROM prosis_horas_pact WHERE prosis_horas_pact.fecha >= " & ConvFecha(fechaDesde) & " AND prosis_horas_pact.fecha <= " & ConvFecha(fechaHasta) & " ) "
   StrSql = StrSql & "    OR empleado.ternro IN (SELECT prosis_licencias.ternro FROM prosis_licencias WHERE prosis_licencias.fecha >= " & ConvFecha(fechaDesde) & " AND prosis_licencias.fecha <= " & ConvFecha(fechaHasta) & " ) "
   
   OpenRecordset StrSql, rsEmpleados
   
   totalReg = rsEmpleados.RecordCount
   i = 0
   
   Do Until rsEmpleados.EOF
      'Busco si el empleado tiene licencias
      
      Flog.Writeline "Generando las novedades para el tercero " & rsEmpleados!ternro

      StrSql = "SELECT ternro, tipdia.tdnro, tipdia.tdcanthoras FROM prosis_licencias "
      StrSql = StrSql & " INNER JOIN tipdia ON prosis_licencias.tdnro = tipdia.tdnro "
      StrSql = StrSql & " WHERE fecha >= " & ConvFecha(fechaDesde)
      StrSql = StrSql & " AND fecha <= " & ConvFecha(fechaHasta)
      StrSql = StrSql & " AND ternro = " & rsEmpleados!ternro
      
      OpenRecordset StrSql, rsLic
      
      tieneLic = False
      tieneLicVac = False
      tieneLicOtras = False
      horasLic = 0
      horasLicOtras = 0
      horasLicVac = 0
      Flog.Writeline "Buscando licencias"
      
      Do Until rsLic.EOF
         
         tieneLic = True
         horasLic = horasLic + CDbl(rsLic!tdcanthoras)
         
         If CInt(rsLic!tdnro) = 2 Then
            tieneLicVac = True
            horasLicVac = horasLicVac + CDbl(rsLic!tdcanthoras)
         Else
            tieneLicOtras = True
            horasLicOtras = horasLicOtras + CDbl(rsLic!tdcanthoras)
         End If
         
         rsLic.MoveNext
      Loop
      
      rsLic.Close
      
      'Busco que empleados tiene horas pactadas
      StrSql = "SELECT sum(horas) AS Total FROM prosis_horas_pact "
      StrSql = StrSql & " WHERE fecha >= " & ConvFecha(fechaDesde)
      StrSql = StrSql & " AND fecha <= " & ConvFecha(fechaHasta)
      StrSql = StrSql & " AND ternro = " & rsEmpleados!ternro
    
      OpenRecordset StrSql, rsTrab
      
      Flog.Writeline "Buscando horas pactadas"
      
      horasTrab = 0
      If Not rsTrab.EOF Then
         If Not IsNull(rsTrab!total) Then
            horasTrab = CDbl(rsTrab!total)
         End If
      End If
      
      rsTrab.Close
      
      'Controlo en que caso entro
      'topeHorasLic = 260
      'basicoHoras = 192
      
      Flog.Writeline "Controlando algoritmo"
      
      If tieneLic Then '1
         
         If tieneLicVac Then '2
            
            If tieneLicOtras Then '3
            
               Flog.Writeline "Algoritmo, Hito 1"
               
'               'Genera las novedades para las licencias
                Call pagarHorasLicCompletas(rsEmpleados!ternro)
                
                If horasLicVac >= basicoHoras Then '4
                
                    'genera la novedad para las HB
                    Flog.Writeline "Algoritmo, Hito 7a"
                    Valor = horasTrab
                    Call generarNovedad(rsEmpleados!ternro, concHB, paramHB, Valor)
                
                Else
                    nuevoTope = basicoHoras - horasLicVac
                    
                    If horasLicOtras > nuevoTope Then '5
                    
                        'genera la novedad para las HB
                        Flog.Writeline "Algoritmo, Hito 7b"
                        Valor = horasTrab
                        Call generarNovedad(rsEmpleados!ternro, concHB, paramHB, Valor)
                        
                    Else
                    
                        If (horasLicOtras + horasTrab) > nuevoTope Then '6
                        
                            'genera la novedad para las HN
                            Flog.Writeline "Algoritmo, Hito 7c"
                            Valor = nuevoTope - horasLicOtras
                            Call generarNovedad(rsEmpleados!ternro, concHN, paramHN, Valor)
                            
                            'genera la novedad para las HB
                            Flog.Writeline "Algoritmo, Hito 7d"
                            Valor = horasTrab - Valor
                            Call generarNovedad(rsEmpleados!ternro, concHB, paramHB, Valor)
                            
                        Else
                        
                            'genera la novedad para las HN
                            Flog.Writeline "Algoritmo, Hito 7e"
                            Valor = horasTrab
                            Call generarNovedad(rsEmpleados!ternro, concHN, paramHN, Valor)
                        
                        End If '6
                    
                    End If '5
                    
                End If '4
               
            'si no tiene otras licencias
            Else
            
               If (basicoHoras - horasLicVac) >= horasTrab Then '4
            
                    'genera la novedad para las HN
                    Flog.Writeline "Algoritmo, Hito 7"
                    Valor = horasTrab
                    Call generarNovedad(rsEmpleados!ternro, concHN, paramHN, Valor)
               
               Else
               
                    'genera la novedad para las HN
                    Flog.Writeline "Algoritmo, Hito 8"
                    Valor = basicoHoras - horasLicVac
                    Call generarNovedad(rsEmpleados!ternro, concHN, paramHN, Valor)
               
                    'genera la novedad para las HB
                    Flog.Writeline "Algoritmo, Hito 9"
                    Valor = horasTrab - (basicoHoras - horasLicVac)
                    Call generarNovedad(rsEmpleados!ternro, concHB, paramHB, Valor)
               
               End If '4
            
            End If '3
         
         'Si no tiene licencias por vacaciones
         'topeHorasLic = 260
         'basicoHoras = 192
         Else
            
            'Genera las novedades para las licencias
            Flog.Writeline "Algoritmo, Hito 10"
            Call pagarHorasLicCompletas(rsEmpleados!ternro)
            
            If horasLic > basicoHoras Then '3
                
                'genera la novedad para las HN
                Flog.Writeline "Algoritmo, Hito 11"
                'Valor = horasLic
                'Call generarNovedad(rsEmpleados!ternro, concHN, paramHN, Valor)
            
            Else
               
               nuevoTope = basicoHoras - horasLic
               
               If horasTrab > nuevoTope Then '4
                    
                    'genera la novedad para las HN
                    Flog.Writeline "Algoritmo, Hito 11"
                    Valor = nuevoTope
                    Call generarNovedad(rsEmpleados!ternro, concHN, paramHN, Valor)
                    
                    If (horasTrab + horasLic) > topeHorasLic Then '5
                    
                        'genera la novedad para las HE
                        Flog.Writeline "Algoritmo, Hito 11"
                        Valor = topeHorasLic - basicoHoras
                        Call generarNovedad(rsEmpleados!ternro, concHE, paramHE, Valor)
                    
                        'genera la novedad para las HB
                        Flog.Writeline "Algoritmo, Hito 11"
                        Valor = (horasTrab + horasLic) - topeHorasLic
                        Call generarNovedad(rsEmpleados!ternro, concHB, paramHB, Valor)
                       
                    Else
                    
                        'genera la novedad para las HE
                        Flog.Writeline "Algoritmo, Hito 11"
                        Valor = (horasTrab + horasLic) - basicoHoras
                        Call generarNovedad(rsEmpleados!ternro, concHE, paramHE, Valor)
                    
                    End If '5
                  
               Else
               
                    'genera la novedad para las HN
                    Flog.Writeline "Algoritmo, Hito 12"
                    Valor = horasTrab
                    Call generarNovedad(rsEmpleados!ternro, concHN, paramHN, Valor)
               
               End If '4
            
           End If '3
         
        End If '2
         
      'si no tiene licencias
      Else
      
        If horasTrab > basicoHoras Then
           
           'genera la novedad para las HN
           Flog.Writeline "Algoritmo, Hito 16"
           Valor = basicoHoras
           Call generarNovedad(rsEmpleados!ternro, concHN, paramHN, Valor)
           
           If horasTrab > topeHorasLic Then
           
                'genera la novedad para las HE
                Flog.Writeline "Algoritmo, Hito 17"
                Valor = topeHorasLic - basicoHoras
                Call generarNovedad(rsEmpleados!ternro, concHE, paramHE, Valor)
                
                If (horasTrab - topeHorasLic) > topeHEB Then
                
                    'genera la novedad para las HEB
                    Flog.Writeline "Algoritmo, Hito 18A"
                    Valor = topeHEB
                    Call generarNovedad(rsEmpleados!ternro, concHEB, paramHEB, Valor)
                
                    'genera la novedad para las HB
                    Flog.Writeline "Algoritmo, Hito 18B"
                    Valor = (horasTrab - topeHorasLic) - topeHEB
                    Call generarNovedad(rsEmpleados!ternro, concHB, paramHB, Valor)
                
                Else
                
                    'genera la novedad para las HEB
                    Flog.Writeline "Algoritmo, Hito 18C"
                    Valor = (horasTrab - topeHorasLic)
                    Call generarNovedad(rsEmpleados!ternro, concHB, paramHB, Valor)
                
                End If
           
           Else
           
                'genera la novedad para las HE
                Flog.Writeline "Algoritmo, Hito 19"
                Valor = horasTrab - basicoHoras
                Call generarNovedad(rsEmpleados!ternro, concHE, paramHE, Valor)
           
           End If
           
        Else
           
           'genera la novedad para las HN
           Flog.Writeline "Algoritmo, Hito 20"
           Valor = basicoHoras
           Call generarNovedad(rsEmpleados!ternro, concHN, paramHN, Valor)
        
        End If
      
      End If
   
      'Actualizo el estado del proceso
      TiempoAcumulado = GetTickCount
      StrSql = "UPDATE batch_proceso SET bprcprogreso = " & ((((i + 1) * 25) / totalReg) + 75) & _
                 ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                 " WHERE bpronro = " & NroProceso
                 
      objConn.Execute StrSql, , adExecuteNoRecords
   
      rsEmpleados.MoveNext
      i = i + 1
   Loop
   
   rsEmpleados.Close
   
   Set rsEmpleados = Nothing
   Set rsLic = Nothing
   Set rsTrab = Nothing
   
   Flog.Writeline "Termina de generar las novedades"
   
   Exit Sub

MError:
    Flog.Writeline "Error en generarNovedades, Error: " & Err.Description
    Exit Sub

End Sub


Sub pagarHorasLicCompletas(ternro)

Dim rsEmpleados As New ADODB.Recordset
Dim rsLic As New ADODB.Recordset
Dim rsTrab As New ADODB.Recordset
Dim totalTipoDia(100)
Dim j

On Error GoTo MError

    For j = 0 To 100
       totalTipoDia(j) = 0
    Next
  
    StrSql = "SELECT ternro, tipdia.tdnro, tipdia.tdcanthoras FROM prosis_licencias "
    StrSql = StrSql & " INNER JOIN tipdia ON prosis_licencias.tdnro = tipdia.tdnro "
    StrSql = StrSql & " WHERE fecha >= " & ConvFecha(fechaDesde)
    StrSql = StrSql & " AND fecha <= " & ConvFecha(fechaHasta)
    StrSql = StrSql & " AND ternro = " & ternro
      
    OpenRecordset StrSql, rsLic
     
    Do Until rsLic.EOF
       totalTipoDia(CLng(rsLic!tdnro)) = totalTipoDia(CLng(rsLic!tdnro)) + CDbl(rsLic!tdcanthoras)
        
       rsLic.MoveNext
    Loop
      
    rsLic.Close
  
    For j = 0 To 100
       If totalTipoDia(j) <> 0 Then
          Call generarNovedad(ternro, concTipoDia(j), paramTipoDia(j), totalTipoDia(j))
       End If
    Next
    
    Set rsEmpleados = Nothing
    Set rsLic = Nothing
    Set rsTrab = Nothing
    
    Exit Sub
    
MError:
    Flog.Writeline "Error en pagarHorasLicCompletas para el ternro: " & ternro & " Error: " & Err.Description
    Exit Sub

End Sub


Sub generarNovedad(ternro, concnro, tpanro, Valor)

Dim rsConsulta As New ADODB.Recordset

On Error GoTo MError

    If concnro <> 0 And tpanro <> 0 Then
    
        StrSql = " SELECT * FROM novemp WHERE empleado = " & ternro
        StrSql = StrSql & " AND concnro = " & concnro
        StrSql = StrSql & " AND tpanro = " & tpanro
        
        OpenRecordset StrSql, rsConsulta
        
        If rsConsulta.EOF Then
        
           StrSql = "INSERT INTO novemp "
           StrSql = StrSql & "(empleado, concnro, tpanro, nevalor, nevigencia, nedesde, nehasta, neretro, nepliqdesde, nepliqhasta)"
           StrSql = StrSql & " values (" & ternro & ", " & concnro & ", " & tpanro & ", " & numberForSQL(Valor) & ", "
           StrSql = StrSql & "0,null,null,null,null,null) "
          
        Else
        
           StrSql = "UPDATE novemp "
           StrSql = StrSql & " SET nevalor = nevalor + " & numberForSQL(Valor)
           StrSql = StrSql & " WHERE nenro = " & rsConsulta!nenro
        
        End If
        
        rsConsulta.Close
        
        Flog.Writeline "Generando novedad: " & StrSql
    
        objConn.Execute StrSql, , adExecuteNoRecords
    
    End If
    
    Set rsConsulta = Nothing
    
Exit Sub

MError:
    Flog.Writeline "Error en generarNovedades para el ternro: " & ternro & " Error: " & Err.Description
    Exit Sub

End Sub

