Attribute VB_Name = "genTick"
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
Global listaEmpleados
Global listaTickets
Global TodosEmpl
Global ConcNovTick
Global ParamNovTick
Global TipoGruLiq
Global TipoGruTick
Global TipoRegHor


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
    
    Nombre_Arch = PathFLog & "GenNovTick" & "-" & NroProceso & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    Flog.Writeline "Inicio Generación Novedades Tickets: " & Now
    Flog.Writeline "Cambio el estado del proceso a Procesando"
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.Writeline "PID = " & PID
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
    
    TiempoAcumulado = GetTickCount
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, objRs
    
    If Not objRs.EOF Then
    
       Flog.Writeline "Obtengo los parámetros del proceso"
       
       'Obtengo los parametros del proceso
       parametros = objRs!bprcparam
       ArrParametros = Split(parametros, "@")
       
       'Obtengo la fecha desde
       fechaDesde = CDate(objRs!bprcfecdesde)
       
       'Obtengo la fecha hasta
       fechaHasta = CDate(objRs!bprcfechasta)
       
       'Obtengo la lista de empleados a los cuales se les generará novedades de tickets
       'listaEmpleados = ArrParametros(0)
       
       'Obtengo la lista de los grupos de tickets
       listaTickets = ArrParametros(0)
       
       'Obtengo el valor para ver si seleccionaron todos los empleados
       TodosEmpl = CBool(ArrParametros(1) = "3")
              
       'EMPIEZA EL PROCESO
       
       'Busco en el confrep las asociaciones para generar los parametros
       StrSql = " SELECT * FROM confrep "
       StrSql = StrSql & " WHERE repnro = 140 "
      
       OpenRecordset StrSql, objRs2
       
       If objRs2.EOF Then
          Flog.Writeline "No esta configurado el ConfRep"
          Exit Sub
       End If
       
       Flog.Writeline "Obtengo los datos del confrep"
       
       ConcNovTick = 0
       ParamNovTick = 0
       

       Do Until objRs2.EOF
          'Concepto de Novedad de Tickets
          If CLng(objRs2!confnrocol) = 1 Then
             Flog.Writeline "Confrep columna " & objRs2!confnrocol
             ConcNovTick = buscarConcepto(objRs3, objRs2!confval2)
          End If
          
          'Parametro de Novedad de Tickets
          If CLng(objRs2!confnrocol) = 2 Then
             Flog.Writeline "Confrep columna " & objRs2!confnrocol
             ParamNovTick = objRs2!confval
          End If
          
          'Parametro de Tipo de Grupo de Liquidación
          If CLng(objRs2!confnrocol) = 3 Then
             Flog.Writeline "Confrep columna " & objRs2!confnrocol
             TipoGruLiq = objRs2!confval
          End If
          
          'Parametro de Tipo de Grupo de Tickets
          If CLng(objRs2!confnrocol) = 4 Then
             Flog.Writeline "Confrep columna " & objRs2!confnrocol
             TipoGruTick = objRs2!confval
          End If
          
          'Parametro de Tipo de Regimen Horario
          If CLng(objRs2!confnrocol) = 5 Then
             Flog.Writeline "Confrep columna " & objRs2!confnrocol
             TipoRegHor = objRs2!confval
          End If
          
          objRs2.MoveNext
       Loop
       
       objRs2.Close

       Call generarNovedadTickets
           
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
Function numberForSQL(Str)
   
  numberForSQL = Replace(Str, ",", ".")

End Function

Function fechaToXML(Fecha)
  
  Fecha = CDate(Fecha)
  
  fechaToXML = Month(Fecha) & "/" & Day(Fecha) & "/" & Year(Fecha)

End Function
Sub generarNovedadTickets()

Dim rsEmpleados As New ADODB.Recordset
Dim rsAux As New ADODB.Recordset
Dim rsEstr As New ADODB.Recordset
Dim rsRegHor As New ADODB.Recordset
Dim rsEmplic As New ADODB.Recordset
Dim rsTipCon As New ADODB.Recordset
Dim noLiq As Boolean
Dim noTick As Boolean
Dim l_i
Dim TipoSelec As Integer
Dim CantDias As Double
Dim Encontre As Boolean
Dim ArrTick
Dim ArrAux
Dim TdnroAnt As Long
Dim FecIniCalc As Date
Dim FecFinCalc As Date
Dim FechaAux As Date
Dim DiasLic As Double
Dim NeValor As Double
Dim totalReg
Dim i
Dim HayDatos As Boolean


On Error GoTo MError

   Flog.Writeline "Empezando a generar las novedades de tickets para los empleados"
   
   'Busco los empleados que vienen en la lista como parámetro al proceso
   StrSql = " SELECT DISTINCT Empleado.ternro, empleg "
   StrSql = StrSql & " From Empleado "
   If TodosEmpl = False Then
      StrSql = StrSql & " INNER JOIN batch_empleado ON batch_empleado.ternro = empleado.ternro AND batch_empleado.bpronro=" & NroProceso
   End If
   
   OpenRecordset StrSql, rsEmpleados
   
   totalReg = rsEmpleados.RecordCount
   i = totalReg
   
   
   Do While Not rsEmpleados.EOF
            
      Flog.Writeline "Busco si el empleado " & rsEmpleados!empleg & " esta en algun grupo de liquidación. "
      
      noLiq = False
      noTick = False
      
      StrSql = "SELECT estrnro FROM his_estructura "
      StrSql = StrSql & " WHERE tenro = " & TipoGruLiq 'grupo liquidacion
      StrSql = StrSql & " AND ternro = " & rsEmpleados!ternro
      StrSql = StrSql & " AND (( htetdesde <= " & ConvFecha(fechaDesde) & ") AND ( htethasta >= " & ConvFecha(fechaDesde) & " OR htethasta is null))"
      
      OpenRecordset StrSql, rsEstr
      
      'Si no esta en ningun grupo de liquidacion paso al proximo empleado
      If rsEstr.EOF Then
         Flog.Writeline "El empleado " & rsEmpleados!empleg & " no esta en ningun grupo de liquidación. Por lo tanto paso al siguiente empleado. "
         noLiq = True
      End If
      rsEstr.Close
      
      Flog.Writeline "Busco si el empleado " & rsEmpleados!empleg & " esta en algun grupo de tickets. "
      
      StrSql = "SELECT his_estructura.estrnro FROM his_estructura "
      StrSql = StrSql & " WHERE tenro = " & TipoGruTick 'grupo tickets
      StrSql = StrSql & " AND ternro = " & rsEmpleados!ternro
      StrSql = StrSql & " AND (( htetdesde <= " & ConvFecha(fechaDesde) & ") AND ( htethasta >= " & ConvFecha(fechaDesde) & " OR htethasta is null))"
      
      OpenRecordset StrSql, rsEstr
      
      'Si no esta en ningun grupo de tickets paso al proximo empleado
      If rsEstr.EOF Then
         Flog.Writeline "El empleado " & rsEmpleados!empleg & " no esta en ningun grupo de tickets. Por lo tanto paso al siguiente empleado. "
         noTick = True
      End If
      
      If noLiq = False And noTick = False Then
               
        'Tiene asignado un grupo de liquidacion y un grupo de tickets
        
        Encontre = False
        l_i = 1
        ArrTick = Split(listaTickets, ",")
        
        'Busco la cantidad de dias y el tipo de seleccion asignado al grupo de tickets que
        'tiene el emplado
        
        Flog.Writeline "Busco la cantidad de dias y el tipo de seleccion asignado al grupo de tickets que tiene el emplado " & rsEmpleados!empleg & ". "
        
        Do While Encontre = False And l_i <= UBound(ArrTick)
           ArrAux = Split(ArrTick(l_i), "$")
           If rsEstr!estrnro = CInt(ArrAux(0)) Then
              TipoSelec = ArrAux(2)
              CantDias = CDbl(ArrAux(1))
              Encontre = True
           End If
           l_i = l_i + 1
        Loop
        rsEstr.Close
        
        'Si la cantidad de dias para el grupo de tickets es igual a 0 paso al proximo empleado
        If CantDias = 0 Then
           Encontre = False
        End If
        
        If Encontre = True Then
           If TipoSelec = 5 Then
           
              StrSql = "SELECT empleg, nrocod FROM empleado "
              StrSql = StrSql & " INNER JOIN his_estructura he ON he.tenro = " & TipoRegHor 'Regimen Horario
              StrSql = StrSql & " AND he.ternro = empleado.ternro "
              StrSql = StrSql & " INNER JOIN estr_cod ON estr_cod.estrnro = he.estrnro "
              StrSql = StrSql & " WHERE he.ternro = " & rsEmpleados!ternro
              StrSql = StrSql & " AND (( htetdesde <= " & ConvFecha(fechaDesde) & ") AND ( htethasta >= " & ConvFecha(fechaDesde) & " OR htethasta is null))"
        
              OpenRecordset StrSql, rsRegHor
              If rsRegHor.EOF Then
                 Flog.Writeline "El empleado " & rsRegHor!empleg & ", no tiene asignado regimen horario. Por lo tanto paso al proximo empleado. "
                 Encontre = False
              End If
            End If
            If Encontre = True Then
                StrSql = "SELECT nenro FROM novemp "
                StrSql = StrSql & " WHERE empleado = " & rsEmpleados!ternro
                StrSql = StrSql & " AND concnro = " & ConcNovTick
                StrSql = StrSql & " AND tpanro = " & ParamNovTick
          
                OpenRecordset StrSql, rsAux
                If rsAux.EOF Then
                   Flog.Writeline " No se encontro la novedad para el tercero " & rsEmpleados!ternro & ". Entonces se creará. "
                   
                   StrSql = "INSERT INTO novemp "
                   StrSql = StrSql & "(empleado, concnro, tpanro, nevalor) "
                   StrSql = StrSql & " VALUES (" & rsEmpleados!ternro & "," & ConcNovTick & "," & ParamNovTick & "," & numberForSQL(CantDias) & ")"
                   objConn.Execute StrSql, , adExecuteNoRecords
                Else
                    
                   StrSql = "UPDATE novemp "
                   StrSql = StrSql & " SET nevalor = " & numberForSQL(CantDias)
                   StrSql = StrSql & " WHERE nenro = " & rsAux!nenro
                   objConn.Execute StrSql, , adExecuteNoRecords
                End If
                rsAux.Close
             
                'Busco las licencias para los empleados dentro de las fechas seleccionadas
                Flog.Writeline "Busco la licencias para el tercero " & rsEmpleados!ternro & " entre las fechas seleccionadas."
                
                StrSql = "SELECT elfechadesde, elfechahasta, tdnro, elcanthrs FROM emp_lic "
                StrSql = StrSql & " WHERE elfechadesde <= " & ConvFecha(fechaHasta)
                StrSql = StrSql & " AND elfechahasta >= " & ConvFecha(fechaDesde)
                StrSql = StrSql & " AND empleado = " & rsEmpleados!ternro
                StrSql = StrSql & " ORDER BY tdnro "
            
                OpenRecordset StrSql, rsEmplic
                TdnroAnt = 0
                HayDatos = False
                
                Do While Not rsEmplic.EOF
             
                    HayDatos = True
                    If TdnroAnt <> rsEmplic!tdnro Then
                       TdnroAnt = rsEmplic!tdnro
                       DiasLic = 0
                    End If
                    
                    If rsEmplic!elfechadesde < fechaDesde Then
                       FecIniCalc = fechaDesde
                    Else
                       FecIniCalc = rsEmplic!elfechadesde
                    End If
                    
                    If rsEmplic!elfechahasta > fechaHasta Then
                       FecFinCalc = fechaHasta
                    Else
                       FecFinCalc = rsEmplic!elfechahasta
                    End If
                
                    Select Case TipoSelec
                      Case 1
                         FechaAux = FecIniCalc
                         Do While FechaAux <= FecFinCalc
                         
                            If Weekday(FechaAux) <> 1 And Weekday(FechaAux) <> 7 Then
                               If esFeriado(FechaAux) = False Then
                                  DiasLic = DiasLic + 1
                               End If
                            End If
                            
                            FechaAux = FechaAux + 1
                            
                         Loop
                      
                      Case 2
                         FechaAux = FecIniCalc
                         Do While FechaAux <= FecFinCalc
                         
                            If Weekday(FechaAux) <> 1 Then
                               If esFeriado(FechaAux) = False Then
                                  DiasLic = DiasLic + 1
                               End If
                            End If
                            
                            FechaAux = FechaAux + 1
                            
                         Loop
                      Case 3
                         FechaAux = FecIniCalc
                         Do While FechaAux <= FecFinCalc
                         
                            If (CDbl(FechaAux) Mod 2) = (CDbl(fechaDesde) Mod 2) Then
                               DiasLic = DiasLic + 1
                            End If
                            
                            FechaAux = FechaAux + 1
                            
                         Loop
                      Case 4
                         FechaAux = FecIniCalc
                         Do While FechaAux <= FecFinCalc
                         
                            If (CDbl(FechaAux) Mod 2) <> (CDbl(fechaDesde) Mod 2) Then
                               DiasLic = DiasLic + 1
                            End If
                            
                            FechaAux = FechaAux + 1
                            
                         Loop
                      Case 5
                         If Not IsNull(rsEmplic!elcanthrs) And rsEmplic!elcanthrs <> 0 Then
                            DiasLic = Fix((rsEmplic!elcanthrs / rsRegHor!nrocod))
                         Else
                            DiasLic = 0
                         End If
                    End Select
                
                  rsEmplic.MoveNext
                  
                  If Not rsEmplic.EOF Then
                     If TdnroAnt <> rsEmplic!tdnro Then
                     
                        StrSql = "SELECT tdnro, tdsuma "
                        StrSql = StrSql & " FROM tipd_con "
                        StrSql = StrSql & " WHERE tipd_con.tdnro = " & rsEmplic!tdnro
                        StrSql = StrSql & " AND tipd_con.concnro = " & ConcNovTick
                              
                        OpenRecordset StrSql, rsTipCon
                        
                        Do While Not rsTipCon.EOF
                           
                           'Generar la novedad para el concepto y parametro configurado
                           Flog.Writeline "Generando las novedades de tickets para el tercero " & rsEmpleados!ternro
                  
                           StrSql = "SELECT nenro, nevalor FROM novemp "
                           StrSql = StrSql & " WHERE empleado = " & rsEmpleados!ternro
                           StrSql = StrSql & " AND concnro = " & ConcNovTick
                           StrSql = StrSql & " AND tpanro = " & ParamNovTick
            
                           OpenRecordset StrSql, rsAux
                           
                           If rsAux.EOF Then
                                                        
                              If rsTipCon!tdsuma = -1 Then
                                 StrSql = "INSERT INTO novemp "
                                 StrSql = StrSql & "(empleado, concnro, tpanro, nevalor) "
                                 StrSql = StrSql & " VALUES (" & rsEmpleados!ternro & "," & ConcNovTick & "," & ParamNovTick & "," & (CantDias + DiasLic) & ")"
                                 objConn.Execute StrSql, , adExecuteNoRecords
                              Else
                                                              
                                 StrSql = "INSERT INTO novemp "
                                 StrSql = StrSql & "(empleado, concnro, tpanro, nevalor) "
                                 StrSql = StrSql & " VALUES (" & rsEmpleados!ternro & "," & ConcNovTick & "," & ParamNovTick & "," & (CantDias - DiasLic) & ")"
                                 objConn.Execute StrSql, , adExecuteNoRecords
                              End If
                           Else
                                
                              NeValor = rsAux!NeValor
                                                          
                              If rsTipCon!tdsuma = -1 Then
                                 NeValor = NeValor + DiasLic
                                 
                                 StrSql = "UPDATE novemp "
                                 StrSql = StrSql & " SET nevalor = " & numberForSQL(NeValor)
                                 StrSql = StrSql & " WHERE nenro = " & rsAux!nenro
                                 objConn.Execute StrSql, , adExecuteNoRecords
                              Else
                                 NeValor = NeValor - DiasLic
                                 
                                 StrSql = "UPDATE novemp "
                                 StrSql = StrSql & " SET nevalor = " & numberForSQL(NeValor)
                                 StrSql = StrSql & " WHERE nenro = " & rsAux!nenro
                                 objConn.Execute StrSql, , adExecuteNoRecords
                              End If
                           End If
                           rsAux.Close
                           
                           rsTipCon.MoveNext
                           
                        Loop ' tipd_con
                        
                        rsTipCon.Close
                        
                     End If 'Cambio el tdnro
                  End If 'NOT EOF
               Loop ' emp_lic
               
               rsEmplic.Close
               
               If HayDatos = True Then
                 StrSql = "SELECT tdnro, tdsuma "
                 StrSql = StrSql & " FROM tipd_con "
                 StrSql = StrSql & " WHERE tipd_con.tdnro = " & TdnroAnt
                 StrSql = StrSql & " AND tipd_con.concnro = " & ConcNovTick
                                
                 OpenRecordset StrSql, rsTipCon
                          
                 Do While Not rsTipCon.EOF
                          
                    'Generar la novedad para el concepto y parametro configurado
                    Flog.Writeline "Generando las novedades de tickets para el tercero " & rsEmpleados!ternro
                  
                    StrSql = "SELECT nenro, nevalor FROM novemp "
                    StrSql = StrSql & " WHERE empleado = " & rsEmpleados!ternro
                    StrSql = StrSql & " AND concnro = " & ConcNovTick
                    StrSql = StrSql & " AND tpanro = " & ParamNovTick
              
                    OpenRecordset StrSql, rsAux
                             
                    If rsAux.EOF Then
                                                          
                       If rsTipCon!tdsuma = -1 Then
                          StrSql = "INSERT INTO novemp "
                          StrSql = StrSql & "(empleado, concnro, tpanro, nevalor) "
                          StrSql = StrSql & " VALUES (" & rsEmpleados!ternro & "," & ConcNovTick & "," & ParamNovTick & "," & (CantDias + DiasLic) & ")"
                          objConn.Execute StrSql, , adExecuteNoRecords
                       Else
                                                                
                          StrSql = "INSERT INTO novemp "
                          StrSql = StrSql & "(empleado, concnro, tpanro, nevalor) "
                          StrSql = StrSql & " VALUES (" & rsEmpleados!ternro & "," & ConcNovTick & "," & ParamNovTick & "," & (CantDias - DiasLic) & ")"
                          objConn.Execute StrSql, , adExecuteNoRecords
                       End If
                          
                    Else
                                
                       NeValor = rsAux!NeValor
                                
                       If rsTipCon!tdsuma = -1 Then
                          NeValor = NeValor + DiasLic
                          
                          StrSql = "UPDATE novemp "
                          StrSql = StrSql & " SET nevalor = " & numberForSQL(NeValor)
                          StrSql = StrSql & " WHERE nenro = " & rsAux!nenro
                          objConn.Execute StrSql, , adExecuteNoRecords
                          
                       Else
                          NeValor = NeValor - DiasLic
                          
                          StrSql = "UPDATE novemp "
                          StrSql = StrSql & " SET nevalor = " & numberForSQL(NeValor)
                          StrSql = StrSql & " WHERE nenro = " & rsAux!nenro
                          objConn.Execute StrSql, , adExecuteNoRecords
                          
                       End If
                    End If
                    rsAux.Close
                             
                    rsTipCon.MoveNext
                          
                 Loop ' tipd_con
                 
                 rsTipCon.Close
                 
               End If 'haydatos = true
            End If 'if encontre = true
        End If 'if encontre = true
        
      End If 'if noLiq = False And noTick = False
      
      'Actualizo el estado del proceso
      i = i - 1
      TiempoAcumulado = GetTickCount
      
      StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Fix(((totalReg - i) * 100) / totalReg) & _
               ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
               ", bprcempleados ='" & CStr(i) & "' WHERE bpronro = " & NroProceso
                
      objConn.Execute StrSql, , adExecuteNoRecords
   
      rsEmpleados.MoveNext
      
   Loop
   
   rsEmpleados.Close
      
   Set rsEmpleados = Nothing
   Set rsTipCon = Nothing
   Set rsEmplic = Nothing
   Set rsAux = Nothing
   Set rsEstr = Nothing
   Set rsRegHor = Nothing
      
   Flog.Writeline "Termina de generar las novedades"
   
   Exit Sub

MError:
    Flog.Writeline "Error en generarNovedadTickets, Error: " & Err.Description
    Exit Sub

End Sub
Function esFeriado(Fecha As Date)
Dim rsPais As New ADODB.Recordset
Dim rsFeriado As New ADODB.Recordset
Dim StrSql As String
Dim Pais
Dim Salida

    'Obtengo el pais en el que estoy
    StrSql = " SELECT * FROM pais "
    StrSql = StrSql & " WHERE paisdef = -1 "
    
    OpenRecordset StrSql, rsPais
    
    If Not rsPais.EOF Then
       Pais = rsPais!paisnro
    Else
       Pais = 0
    End If
    
    rsPais.Close
    
    Salida = False
    
    'Busco si la fecha es un feriado
    StrSql = " SELECT * FROM feriado "
    StrSql = StrSql & " WHERE feriado.ferifecha = " & ConvFecha(Fecha)
      
    OpenRecordset StrSql, rsFeriado
      
    If Not rsFeriado.EOF Then
         Salida = ((CInt(rsFeriado!tipferinro) = 1) And (CInt(rsFeriado!fericodext) = CInt(Pais)))
      End If
      rsFeriado.Close
        
      esFeriado = Salida

End Function

