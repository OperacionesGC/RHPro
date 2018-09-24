Attribute VB_Name = "becasAJornal"
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

Global repNro As Integer
Global conceptos As String
Global acumuladores As String
Global procesos As String
Global fechaIngreso As Date
Global tipoModelo As Integer
Global idUser As String

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
Dim tipoDepuracion
Dim historico As Boolean
Dim param
Dim ternro
Dim empleg
Dim rsEmpl As New ADODB.Recordset
Dim i
Dim totalEmpleados
Dim cantRegistros
Dim PID As String
Dim parametros As String
Dim ArrParametros

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
    
    Nombre_Arch = PathFLog & "BecasAJornal" & "-" & NroProceso & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    'Obtengo la cantidad de empledos a procesar
    StrSql = "SELECT count(*) AS total FROM batch_empleado WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, objRs
    
    cantRegistros = CInt(objRs!total)
    totalEmpleados = cantRegistros
    
    objRs.Close
   
    Flog.writeline "Inicio Proceso de Control Pagos : " & Now
    Flog.writeline "Cambio el estado del proceso a Procesando"
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "PID = " & PID
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Obtengo los datos del proceso"
    
    TiempoAcumulado = GetTickCount
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, objRs
    
    If Not objRs.EOF Then
       
       'Obtengo la fecha de ingreso
       fechaIngreso = objRs!bprcfecdesde
       
       'EMPIEZA EL PROCESO

       'Obtengo los empleados sobre los que tengo que generar los recibos
       CargarEmpleados NroProceso, rsEmpl
       
       StrSql = "UPDATE batch_proceso SET bprcprogreso = 0 " & _
                   ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                   ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
       
       objConn.Execute StrSql, , adExecuteNoRecords
       
       'Genero por cada empleado un registro
       Do Until rsEmpl.EOF
          EmpErrores = False
          ternro = rsEmpl!ternro
                    
          'Genero los datos del empleado
          Call cambiarBecasAJornal(ternro)
                
          'Actualizo el estado del proceso
          TiempoAcumulado = GetTickCount
          
          cantRegistros = cantRegistros - 1
          
          StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Fix(((totalEmpleados - cantRegistros) * 100) / totalEmpleados) & _
                   ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                   ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
             
          objConn.Execute StrSql, , adExecuteNoRecords
          
          'Si se generaron todos los datos del empleado correctamente lo borro
          If Not EmpErrores Then
              StrSql = " DELETE FROM batch_empleado "
              StrSql = StrSql & " WHERE bpronro = " & NroProceso
              StrSql = StrSql & " AND ternro = " & ternro
    
              objConn.Execute StrSql, , adExecuteNoRecords
          End If
          
          rsEmpl.MoveNext
       Loop
    
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
    
    Flog.writeline "Fin :" & Now
    Flog.Close

    Exit Sub
    
CE:
    HuboErrores = True
    Flog.writeline " Error: " & Err.Description & Now

End Sub

Function controlNull(Str)
  If Trim(Str) = "" Then
     controlNull = "null"
  Else
     controlNull = "'" & Str & "'"
  End If
End Function

'--------------------------------------------------------------------
' Se encarga de pasar a los empleados que estan en becas a jornalizados
'--------------------------------------------------------------------
Sub cambiarBecasAJornal(ByVal ternro)

Dim StrSql As String
Dim rsConsult As New ADODB.Recordset
Dim rsConsult2 As New ADODB.Recordset
Dim fechaBaja As Date
Dim fechaBajaPrev As Date
Dim causaBaja As Integer
Dim cantDias As Integer
Dim estrnro As String

On Error GoTo MError

fechaBaja = fechaIngreso - 1

'------------------------------------------------------------------
'Busco el tipo modelo y la causa de baja configurado en la confrep
'------------------------------------------------------------------
StrSql = " SELECT confval, conftipo "
StrSql = StrSql & " FROM confrep"
StrSql = StrSql & " WHERE confrep.repnro = 121 "

OpenRecordset StrSql, rsConsult

cantDias = -1
Do While Not rsConsult.EOF

   If rsConsult!conftipo = "MO" Then
      tipoModelo = rsConsult!confval
      ElseIf rsConsult!conftipo = "CB" Then
             causaBaja = rsConsult!confval
             ElseIf rsConsult!conftipo = "VAL" Then
                 cantDias = rsConsult!confval
   End If
   
   rsConsult.MoveNext
   
Loop

rsConsult.Close


'------------------------------------------------------------------
'Determino la fecha prevista
'------------------------------------------------------------------

fechaBajaPrev = DateAdd("d", cantDias, fechaIngreso)

'--------------------------------------------------------------------------------------------
'Actualizo el empleado si esta confifurado la cantidad de días en la configuración de reporte
'--------------------------------------------------------------------------------------------

If cantDias > -1 Then
   StrSql = " UPDATE empleado SET "
   StrSql = StrSql & " empfbajaprev =" & ConvFecha(fechaBajaPrev)
   StrSql = StrSql & " ,tplatenro =" & tipoModelo
   StrSql = StrSql & " ,empfaltagr =" & ConvFecha(fechaIngreso)
   StrSql = StrSql & " WHERE empleado.ternro =" & ternro
   
   Else: StrSql = " UPDATE empleado SET "
         StrSql = StrSql & " ,tplatenro =" & tipoModelo
         StrSql = StrSql & " ,empfaltagr =" & ConvFecha(fechaIngreso)
         StrSql = StrSql & " WHERE empleado.ternro =" & ternro
         
End If
objConn.Execute StrSql, , adExecuteNoRecords


'------------------------------------------------------------------
'Cierro la fase del empleado
'------------------------------------------------------------------

StrSql = " UPDATE fases SET "
StrSql = StrSql & " bajfec =" & ConvFecha(fechaBaja) & ","
StrSql = StrSql & " caunro =" & causaBaja & ","
StrSql = StrSql & " sueldo = 0,"
StrSql = StrSql & " vacaciones = 0,"
StrSql = StrSql & " indemnizacion = 0,"
StrSql = StrSql & " real = 0,"
StrSql = StrSql & " fasrecofec = 0,"
StrSql = StrSql & " estado = 0"
StrSql = StrSql & " WHERE fases.empleado =" & ternro

objConn.Execute StrSql, , adExecuteNoRecords

'------------------------------------------------------------------
'Abro la nueva fase para el empleado
'------------------------------------------------------------------

StrSql = " INSERT INTO fases "
StrSql = StrSql & " (empleado, altfec, sueldo, vacaciones, indemnizacion, real, fasrecofec)"
StrSql = StrSql & " VALUES "
StrSql = StrSql & "(" & ternro
StrSql = StrSql & "," & ConvFecha(fechaIngreso)
StrSql = StrSql & ",-1"
StrSql = StrSql & ",-1"
StrSql = StrSql & ",-1"
StrSql = StrSql & ",-1"
StrSql = StrSql & ",-1)"
    
objConn.Execute StrSql, , adExecuteNoRecords

'------------------------------------------------------------------
'Cierro todas las estrucutras asociada al empleado
'------------------------------------------------------------------

StrSql = " SELECT adptte_estr.tenro, adptte_estr.tplaestrnro, tedabr, tplaestroblig "
StrSql = StrSql & " FROM adptemplate "
StrSql = StrSql & " INNER JOIN adptte_estr ON adptte_estr.tplatenro = adptemplate.tplatenro "
StrSql = StrSql & " INNER JOIN tipoestructura ON tipoestructura.tenro = adptte_estr.tenro "
StrSql = StrSql & " WHERE adptemplate.tplatenro = " & tipoModelo

OpenRecordset StrSql, rsConsult

Do While Not rsConsult.EOF
   
   If rsConsult!tplaestroblig = -1 Then
      StrSql = " UPDATE his_estructura SET "
      StrSql = StrSql & " htethasta =" & ConvFecha(fechaBaja) & ""
      StrSql = StrSql & " WHERE his_estructura.ternro =" & ternro & " AND htethasta is null "
      StrSql = StrSql & " AND his_estructura.tenro =" & rsConsult!tenro

      objConn.Execute StrSql, , adExecuteNoRecords
   End If

   rsConsult.MoveNext
   
Loop
'-------------------------------------------------------------------------------
'Abro las nuevas estructuras para el empleado dependiendo del modelo configurado
'-------------------------------------------------------------------------------

StrSql = " SELECT adptte_estr.tenro, adptte_estr.tplaestrnro, tedabr "
StrSql = StrSql & " FROM adptemplate "
StrSql = StrSql & " INNER JOIN adptte_estr ON adptte_estr.tplatenro = adptemplate.tplatenro "
StrSql = StrSql & " INNER JOIN tipoestructura ON tipoestructura.tenro = adptte_estr.tenro "
StrSql = StrSql & " WHERE adptemplate.tplatenro = " & tipoModelo

OpenRecordset StrSql, rsConsult

Do While Not rsConsult.EOF
   If IsNull(rsConsult!tplaestrnro) Or rsConsult!tplaestrnro = 0 Then
      'Flog.writeline "El tipo de estucutra: '" & rsConsult!tedabr & "' no tiene configurada ninguna estructura por defecto en el modelo."
      
      StrSql = " SELECT estrnro "
      StrSql = StrSql & " FROM his_estructura "
      StrSql = StrSql & " WHERE his_estructura.ternro =" & ternro
      StrSql = StrSql & " AND tenro =" & rsConsult!tenro
      StrSql = StrSql & " AND htethasta =" & ConvFecha(fechaBaja)
      OpenRecordset StrSql, rsConsult2
      
      If Not rsConsult2.EOF Then
         StrSql = " INSERT INTO his_estructura "
         StrSql = StrSql & " (ternro , tenro, estrnro, htetdesde)"
         StrSql = StrSql & " VALUES "
         StrSql = StrSql & "(" & ternro
         StrSql = StrSql & "," & rsConsult!tenro
         StrSql = StrSql & "," & rsConsult2!estrnro
         StrSql = StrSql & "," & ConvFecha(fechaIngreso) & ")"
      
         objConn.Execute StrSql, , adExecuteNoRecords
      End If
      rsConsult2.Close
   
   Else
         StrSql = " INSERT INTO his_estructura "
         StrSql = StrSql & " (ternro , tenro, estrnro, htetdesde)"
         StrSql = StrSql & " VALUES "
         StrSql = StrSql & "(" & ternro
         StrSql = StrSql & "," & rsConsult!tenro
         StrSql = StrSql & "," & rsConsult!tplaestrnro
         StrSql = StrSql & "," & ConvFecha(fechaIngreso) & ")"
      
         objConn.Execute StrSql, , adExecuteNoRecords
               
   End If

   rsConsult.MoveNext
   
Loop

rsConsult.Close

Exit Sub

MError:
    Flog.writeline "Error en el tercero " & ternro & " Error: " & Err.Description
    HuboErrores = True
    EmpErrores = True
    Exit Sub
End Sub

'--------------------------------------------------------------------
' Se encarga de generar un ResultSet de los empleados a cambiar
' si el RS es vacio significa que hay que aplicarlo sobre todos
'--------------------------------------------------------------------
Sub CargarEmpleados(NroProc, ByRef rsEmpl As ADODB.Recordset)

Dim StrEmpl As String

    StrEmpl = "SELECT * FROM batch_empleado WHERE bpronro = " & NroProc
    
    OpenRecordset StrEmpl, rsEmpl
End Sub

Function numberForSQL(Str)
   
  numberForSQL = Replace(Str, ",", ".")

End Function


Function strForSQL(Str)
   
  If IsNull(Str) Then
     strForSQL = "NULL"
  Else
     strForSQL = Str
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


