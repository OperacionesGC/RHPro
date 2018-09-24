Attribute VB_Name = "repControlF931"
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

Global tenro1 As Integer
Global estrnro1 As Integer
Global tenro2 As Integer
Global estrnro2 As Integer
Global tenro3 As Integer
Global estrnro3 As Integer
Global fecEstr As String

Global tipoHora(9) As Integer
Global acumulador As Integer
Global escala(9) As Integer


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
Dim fechadesde
Dim fechahasta
Dim tipoDepuracion
Dim historico As Boolean
Dim param
Dim listapronro
Dim pronro
Dim ternro
Dim pliqnro
Dim rsEmpl As New ADODB.Recordset
Dim acunroSueldo
Dim I
Dim totalEmpleados
Dim cantRegistros
Dim PID As String
Dim parametros As String
Dim ArrParametros
Dim strTempo As String
Dim orden

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
    
    Nombre_Arch = PathFLog & "ReporteControlF931" & "-" & NroProceso & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    'Obtengo la cantidad de empledos a procesar
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, objRs
    
    cantRegistros = CInt(objRs!bprcempleados)
    totalEmpleados = cantRegistros
    
    objRs.Close
   
    Flog.writeline "Inicio Proceso de Control F931 : " & Now
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
       
       'Obtengo los parametros del proceso
       parametros = objRs!bprcparam
       ArrParametros = Split(parametros, "@")
       
       'Obtengo la lista de procesos
       pliqnro = ArrParametros(0)
       
       'Obtengo los cortes de estructura
       tenro1 = CInt(ArrParametros(1))
       estrnro1 = CInt(ArrParametros(2))
       tenro2 = CInt(ArrParametros(3))
       estrnro2 = CInt(ArrParametros(4))
       tenro3 = CInt(ArrParametros(5))
       estrnro3 = CInt(ArrParametros(6))
              
       'EMPIEZA EL PROCESO
       
       'Busco en el confrep el numero de cuenta que se va a usar para
       ' buscar el valor de la escala, tipo de hora y acumulador
       StrSql = " SELECT * FROM confrep "
       StrSql = StrSql & " WHERE repnro = 86 "
      
       OpenRecordset StrSql, objRs2
       
       For I = 1 To 8
           tipoHora(I) = 0
           escala(I) = 0
       Next
       
       acumulador = 0
       
       If objRs2.EOF Then
          Flog.writeline "No esta configurado el ConfRep nro. 86 "
          Exit Sub
       End If
       
       Flog.writeline "Obtengo los datos del confrep"
       
       Do Until objRs2.EOF
       
          Select Case objRs2!confnrocol
             
             Case 1
                  acumulador = objRs2!confval
             Case 2
                  tipoHora(1) = CInt(objRs2!confval)
                  escala(1) = CDbl(objRs2!confval2)
             Case 3
                  tipoHora(2) = CInt(objRs2!confval)
                  escala(2) = CDbl(objRs2!confval2)
             Case 4
                  tipoHora(3) = CInt(objRs2!confval)
                  escala(3) = CDbl(objRs2!confval2)
             Case 5
                  tipoHora(4) = CInt(objRs2!confval)
                  escala(4) = CDbl(objRs2!confval2)
             Case 6
                  tipoHora(5) = CInt(objRs2!confval)
                  escala(5) = CDbl(objRs2!confval2)
             Case 7
                  tipoHora(6) = CInt(objRs2!confval)
                  escala(6) = CDbl(objRs2!confval2)
             Case 8
                  tipoHora(7) = CInt(objRs2!confval)
                  escala(7) = CDbl(objRs2!confval2)
             Case 9
                  tipoHora(8) = CInt(objRs2!confval)
                  escala(8) = CDbl(objRs2!confval2)
                  
          End Select
       
          objRs2.MoveNext
       Loop

       'Obtengo los empleados sobre los que tengo que generar los recibos
       CargarEmpleados NroProceso, rsEmpl
       
       StrSql = "UPDATE batch_proceso SET bprcprogreso = 0 " & _
                   ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                   ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
       
       'Genero por cada empleado un registro
       Do Until rsEmpl.EOF
          EmpErrores = False
          ternro = rsEmpl!ternro
          orden = rsEmpl!estado
          
          Flog.writeline "Generando datos empleado " & ternro & " para el periodo " & pliqnro
             
          Call generarDatosEmpleado01(pliqnro, ternro, orden)
          
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
' Se encarga de generar los datos
'--------------------------------------------------------------------
Sub generarDatosEmpleado01(pliqnro, ternro, orden)

Dim StrSql As String
Dim rsConsult As New ADODB.Recordset

'Variables donde se guardan los datos del INSERT final

Dim Legajo As Integer
Dim apellido As String
Dim apellido2 As String
Dim nombre As String
Dim nombre2 As String
Dim pliqmes
Dim pliqanio
Dim pliqdesde
Dim pliqhasta
Dim pliqdesc
Dim Cuil As String
Dim estado As String

Dim Tipo As Integer
Dim sistema
Dim f931
Dim estrnomb1
Dim estrnomb2
Dim estrnomb3
Dim tenomb1
Dim tenomb2
Dim tenomb3
Dim tieneAcumDiario
Dim maletas
Dim auxStr
Dim I

Dim sql As String

On Error GoTo MError

estrnomb1 = ""
estrnomb2 = ""
estrnomb3 = ""
tenomb1 = ""
tenomb2 = ""
tenomb3 = ""
tieneAcumDiario = False

'------------------------------------------------------------------
'Busco los datos del empleado
'------------------------------------------------------------------
StrSql = " SELECT empleg,terape,terape2,ternom,ternom2,empfaltagr,empest "
StrSql = StrSql & " FROM empleado "
StrSql = StrSql & " WHERE ternro= " & ternro
       
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   nombre = rsConsult!ternom
   If IsNull(rsConsult!ternom2) Then
      nombre2 = ""
   Else
      nombre2 = rsConsult!ternom2
   End If
   apellido = rsConsult!terape
   If IsNull(rsConsult!terape2) Then
      apellido2 = ""
   Else
      apellido2 = rsConsult!terape2
   End If
   Legajo = rsConsult!empleg
   If CInt(rsConsult!empest) = -1 Then
      estado = "Activo"
   Else
      estado = "Inactivo"
   End If
   
Else
   Flog.writeline "Error al obtener los datos del empleado"
'   GoTo MError
End If

rsConsult.Close

'------------------------------------------------------------------
'Busco los datos del periodo actual
'------------------------------------------------------------------
StrSql = " SELECT * FROM periodo "
StrSql = StrSql & " WHERE pliqnro=" & pliqnro
       
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   pliqmes = rsConsult!pliqmes
   pliqanio = rsConsult!pliqanio
   pliqdesde = rsConsult!pliqdesde
   pliqhasta = rsConsult!pliqhasta
   pliqdesc = rsConsult!pliqdesc
   fecEstr = pliqhasta
Else
   Flog.writeline "Error al obtener los datos del periodo"
   GoTo MError
End If

rsConsult.Close
 
'------------------------------------------------------------------
'Busco los datos del tipos de estructura 1
'------------------------------------------------------------------

If tenro1 <> 0 Then
    
    StrSql = " SELECT estrdabr, tedabr "
    StrSql = StrSql & " FROM his_estructura "
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
    StrSql = StrSql & "    AND his_estructura.ternro = " & ternro & " AND his_estructura.tenro = " & tenro1
    StrSql = StrSql & "    AND (htetdesde<=" & ConvFecha(fecEstr) & " AND (htethasta is null or htethasta>=" & ConvFecha(fecEstr) & "))"
    
    If estrnro1 <> 0 Then
        StrSql = StrSql & " AND estructura.estrnro = " & estrnro1
    End If
    
    StrSql = StrSql & "  INNER JOIN tipoestructura ON tipoestructura.tenro = his_estructura.tenro "
           
    OpenRecordset StrSql, rsConsult
    
    If Not rsConsult.EOF Then
       estrnomb1 = rsConsult!estrdabr
       tenomb1 = rsConsult!tedabr
    End If
End If


'------------------------------------------------------------------
'Busco los datos del tipos de estructura 2
'------------------------------------------------------------------

If tenro2 <> 0 Then
    
    StrSql = " SELECT estrdabr, tedabr "
    StrSql = StrSql & " FROM his_estructura "
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
    StrSql = StrSql & "    AND his_estructura.ternro = " & ternro & " AND his_estructura.tenro = " & tenro2
    StrSql = StrSql & "    AND (htetdesde<=" & ConvFecha(fecEstr) & " AND (htethasta is null or htethasta>=" & ConvFecha(fecEstr) & "))"
    
    If estrnro2 <> 0 Then
        StrSql = StrSql & " AND estructura.estrnro = " & estrnro2
    End If
    
    StrSql = StrSql & "  INNER JOIN tipoestructura ON tipoestructura.tenro = his_estructura.tenro "
           
    OpenRecordset StrSql, rsConsult
    
    If Not rsConsult.EOF Then
       estrnomb2 = rsConsult!estrdabr
       tenomb2 = rsConsult!tedabr
    End If
End If


'------------------------------------------------------------------
'Busco los datos del tipos de estructura 3
'------------------------------------------------------------------

If tenro3 <> 0 Then
    
    StrSql = " SELECT estrdabr, tedabr "
    StrSql = StrSql & " FROM his_estructura "
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
    StrSql = StrSql & "    AND his_estructura.ternro = " & ternro & " AND his_estructura.tenro = " & tenro3
    StrSql = StrSql & "    AND (htetdesde<=" & ConvFecha(fecEstr) & " AND (htethasta is null or htethasta>=" & ConvFecha(fecEstr) & "))"
    
    If estrnro3 <> 0 Then
        StrSql = StrSql & " AND estructura.estrnro = " & estrnro3
    End If
    
    StrSql = StrSql & "  INNER JOIN tipoestructura ON tipoestructura.tenro = his_estructura.tenro "
           
    OpenRecordset StrSql, rsConsult
    
    If Not rsConsult.EOF Then
       estrnomb3 = rsConsult!estrdabr
       tenomb3 = rsConsult!tedabr
    End If
End If
 
 
'------------------------------------------------------------------
'Busco el valor del cuil
'------------------------------------------------------------------

sql = " SELECT nrodoc "
sql = sql & " FROM ter_doc "
sql = sql & " WHERE tidnro=10 AND ternro= " & ternro

OpenRecordset sql, rsConsult

Cuil = ""

If Not rsConsult.EOF Then
   Cuil = rsConsult!nrodoc
End If

rsConsult.Close

'------------------------------------------------------------------
'Busco los datos de la cantidad de maletas
'------------------------------------------------------------------

auxStr = "0"

For I = 1 To 8
   auxStr = auxStr & "," & tipoHora(I)
Next

sql = " SELECT sum(adcanthoras) AS suma, thnro "
sql = sql & " FROM gti_acumdiario "
sql = sql & " WHERE ternro = " & ternro
sql = sql & "   AND  thnro IN (" & auxStr & ") "
sql = sql & "   AND adfecha >= " & ConvFecha(pliqdesde)
sql = sql & "   AND adfecha <= " & ConvFecha(pliqhasta)
sql = sql & " GROUP BY thnro "

OpenRecordset sql, rsConsult

maletas = 0
tieneAcumDiario = False

Do Until rsConsult.EOF
  If Not IsNull(rsConsult!suma) Then
     tieneAcumDiario = True
     maletas = maletas + valorReal(rsConsult!thnro, CDbl(rsConsult!suma))
  End If

  rsConsult.MoveNext
Loop

rsConsult.Close

sistema = maletas

'------------------------------------------------------------------
'Busco los datos de los acumuladores
'------------------------------------------------------------------

'Basico
sql = " SELECT acagmonto "
sql = sql & " FROM acu_age"
sql = sql & " WHERE acunro = " & acumulador
sql = sql & " AND pliqnro =  " & pliqnro
sql = sql & " AND empage = " & ternro

OpenRecordset sql, rsConsult

Tipo = 0

If rsConsult.EOF Then
   f931 = 0
   'Si no tiene acu_age y tiene datos en sistema entonces es de tipo 1
   If tieneAcumDiario Then
      Tipo = 1
   End If
Else
   f931 = rsConsult!acagmonto
End If

rsConsult.Close

'------------------------------------------------------------------
'Armo la SQL para guardar los datos
'------------------------------------------------------------------

StrSql = " INSERT INTO rep931 "
StrSql = StrSql & " (bpronro,ternro,tipo,pliqnro, "
StrSql = StrSql & " pliqdesc, legajo, apellido, apellido2, "
StrSql = StrSql & " nombre, nombre2, cuil, sistema, "
StrSql = StrSql & " f931, estado, orden, "
StrSql = StrSql & " estrdabr1,estrdabr2,estrdabr3,tedabr1,tedabr2,tedabr3) "
StrSql = StrSql & " VALUES "
StrSql = StrSql & "(" & NroProceso
StrSql = StrSql & "," & ternro
StrSql = StrSql & "," & Tipo
StrSql = StrSql & "," & pliqnro
StrSql = StrSql & ",'" & pliqdesc & "'"
StrSql = StrSql & "," & Legajo
StrSql = StrSql & ",'" & apellido & "'"
StrSql = StrSql & ",'" & apellido2 & "'"
StrSql = StrSql & ",'" & nombre & "'"
StrSql = StrSql & ",'" & nombre2 & "'"
StrSql = StrSql & ",'" & Cuil & "'"
StrSql = StrSql & "," & numberForSQL(sistema)
StrSql = StrSql & "," & numberForSQL(f931)
StrSql = StrSql & ",'" & estado & "'"
StrSql = StrSql & "," & orden
StrSql = StrSql & "," & controlNull(estrnomb1)
StrSql = StrSql & "," & controlNull(estrnomb2)
StrSql = StrSql & "," & controlNull(estrnomb3)
StrSql = StrSql & "," & controlNull(tenomb1)
StrSql = StrSql & "," & controlNull(tenomb2)
StrSql = StrSql & "," & controlNull(tenomb3)
StrSql = StrSql & ")"

'------------------------------------------------------------------
'Guardo los datos en la BD
'------------------------------------------------------------------

objConn.Execute StrSql, , adExecuteNoRecords

Exit Sub

MError:
    Flog.writeline "Error en empleado: " & Legajo & " Error: " & Err.Description
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

Function valorReal(thnro As Integer, maletas As Double)
   Dim j
   Dim Salida
   
   Salida = 0
   
   For j = 1 To 8
      If thnro = tipoHora(j) Then
         Salida = maletas * escala(j)
      End If
   Next
   
   valorReal = Salida
End Function
