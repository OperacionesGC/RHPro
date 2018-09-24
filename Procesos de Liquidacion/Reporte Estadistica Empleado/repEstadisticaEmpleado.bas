Attribute VB_Name = "repEstadisticaEmpleado"
Option Explicit

'Global Const Version = "1.00"
'Global Const FechaModificacion = "27/09/2012"
'Global Const UltimaModificacion = "Sebastian Stremel - Inicial"

'Global Const Version = "1.01"
'Global Const FechaModificacion = "17/10/2012"
'Global Const UltimaModificacion = "Sebastian Stremel - correccion de errores - caso 16251"

'Global Const Version = "1.02"
'Global Const FechaModificacion = "23/10/2012"
'Global Const UltimaModificacion = "Sebastian Stremel - se agrego que busco los conceptos de un acumulador - caso 16251"

'Global Const Version = "1.03"
'Global Const FechaModificacion = "24/10/2012"
'Global Const UltimaModificacion = "Sebastian Stremel - correcciones - CAS-16251 - NGA - AZ - Reporte Estadística por Empleado"

'Global Const Version = "1.04"
'Global Const FechaModificacion = "12/11/2012"
'Global Const UltimaModificacion = "Sebastian Stremel - correcciones - CAS-16251 - NGA - AZ - Reporte Estadística por Empleado"

'Global Const Version = "1.05"
'Global Const FechaModificacion = "27/11/2012"
'Global Const UltimaModificacion = "Sebastian Stremel - Se corrigio suma de acumulador - CAS-16251 - NGA - AZ - Reporte Estadística por Empleado. "

'Global Const Version = "1.06"
'Global Const FechaModificacion = "28/11/2012"
'Global Const UltimaModificacion = "Sebastian Stremel - Se corrigio error de division por cero - CAS-16251 - Reporte estadistica por empleado - NGA-(CAS-15298) "

'Global Const Version = "1.07"
'Global Const FechaModificacion = "05/06/2013"
'Global Const UltimaModificacion = "Sebastian Stremel - se corrigio error en la suma del acumulador en los casos que hay mas de un mes CAS-19945 - NORTHGATEARINSO - ERROR EN REPORTE ESTADISTICO POR EMPLEADO"

Global Const Version = "1.08"
Global Const FechaModificacion = "14/06/2013"
Global Const UltimaModificacion = "Sebastian Stremel - se corrigio que si se emite por varios meses traiga el resultado correcto - CAS-19945 - NORTHGATEARINSO - ERROR EN REPORTE ESTADISTICO POR EMPLEADO"


'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
Dim fs, f
'Global Flog

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
Global Formato As Integer
Global Modelo As Long
Global CantColumnas As Integer
Global CodCols(200)
Global CodNovCols(200)
Global CodNov
Global CodDirCols(200)
Global CodDir
Global TitCols(200)
Global TipoCols(200)
Global EsMonto(200)
Global TipoColumna(200)
Global NroCols(200)
Global ValCols(200)
Global CharCols(200)
Global TituloRep As String
Global descDesde
Global descHasta
Global FechaHasta
Global FechaDesde
Global Nro_Col
Global listaPer
Global concAnt
Global Desde
Global Hasta
Global nomape
Global prog As Integer


Private Sub Main()

Dim NombreArchivo As String
Dim directorio As String
Dim CArchivos
Dim archivo
Dim Folder
Dim strCmdLine As String
Dim Nombre_Arch As String

Dim StrSql As String
Dim objRs As New ADODB.Recordset
Dim rsPeriodos As New ADODB.Recordset
Dim pliqdesde
Dim pliqhasta
Dim tipoDepuracion
Dim historico As Boolean
Dim param
Dim listapronro
Dim proNro
Dim Ternro
Dim arrpronro
Dim Periodos
Dim rsEmpl As New ADODB.Recordset
Dim I
Dim totalEmpleados
Dim cantRegistros
Dim PID As String
Dim tituloReporte As String
Dim parametros As String
Dim ArrParametros
Dim strTempo As String
Dim orden

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


    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    OpenConnection strconexion, objconnProgreso
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    On Error GoTo 0

    On Error GoTo CE
    
    HuboErrores = False
    
    Nombre_Arch = PathFLog & "ReporteEstadisticaEmp" & "-" & NroProceso & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    
    
    'Obtengo la cantidad de empledos a procesar
    StrSql = "SELECT count(*) AS total FROM batch_empleado WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, objRs
    
    cantRegistros = CInt(objRs!total)
    totalEmpleados = cantRegistros
    
    objRs.Close
    
     ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Version = " & Version
    Flog.writeline "Modificacion = " & UltimaModificacion
    Flog.writeline "Fecha = " & FechaModificacion
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline
    Flog.writeline "PID = " & PID
   
    ' Obtengo el Process ID
    'PID = GetCurrentProcessId
    'Flog.Writeline "PID = " & PID
    
    Flog.writeline "Inicio Proceso de Conceptos y Acumuladores : " & Now
    Flog.writeline "Cambio el estado del proceso a Procesando"
    
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
       
       Flog.writeline "Lista de Parametros = " & parametros
       
       ArrParametros = Split(parametros, "@")
              
       'Obtengo la lista de procesos
       Flog.writeline "Obtengo la Lista de Procesos"
       
       listapronro = ArrParametros(0)
       
       Flog.writeline "Lista de Procesos = " & listapronro
       
       'Obtengo el modelo a usar
       Flog.writeline "Obtengo el Modelo a Usar"
       
       Modelo = CLng(ArrParametros(1))
       
       Flog.writeline "Modelo = " & Modelo
       
       
       'Obtengo el periodo desde
       Flog.writeline "Obtengo el Período Desde"
       
       pliqdesde = CLng(ArrParametros(2))
       
       Flog.writeline "Período Desde = " & pliqdesde
       
       'Obtengo el periodo hasta
       Flog.writeline "Obtengo el Período Hasta"
       
       pliqhasta = CLng(ArrParametros(3))
       
       Flog.writeline "Período Hasta = " & pliqhasta
       
       'Obtengo los cortes de estructura
       Flog.writeline "Obtengo los cortes de estructuras"
       
       Flog.writeline "Obtengo estructura 1"
       tenro1 = CInt(ArrParametros(4))
       estrnro1 = CInt(ArrParametros(5))
       
       Flog.writeline "Corte 1 = " & tenro1 & " - " & estrnro1
       
       Flog.writeline "Obtengo estructura 2"
       tenro2 = CInt(ArrParametros(6))
       estrnro2 = CInt(ArrParametros(7))
       
       Flog.writeline "Corte 2 = " & tenro2 & " - " & estrnro2
       
       Flog.writeline "Obtengo estructura 3"
       
       tenro3 = CInt(ArrParametros(8))
       estrnro3 = CInt(ArrParametros(9))
       
       Flog.writeline "Corte 3 = " & tenro3 & " - " & estrnro3
       
       Flog.writeline "Obtengo la Fecha"
       
       fecEstr = ArrParametros(10)
       
       Flog.writeline "Fecha = " & fecEstr
       
       'Obtengo el formato de salida del reporte
       Flog.writeline "Obtengo el Formato del Reporte"
       
       Formato = CLng(ArrParametros(11)) 'OJO POR AHORA NO ANDA
       
       Flog.writeline "Formato del Reporte = " & Formato
       
       'EMPIEZA EL PROCESO
       'Busco el periodo desde
       StrSql = "SELECT * FROM periodo WHERE pliqnro = " & pliqdesde
       OpenRecordset StrSql, objRs
        
       If Not objRs.EOF Then
          FechaDesde = objRs!pliqdesde
          descDesde = objRs!pliqDesc
       Else
          Flog.writeline "No se encontro el periodo desde."
          Exit Sub
       End If
        
       objRs.Close
       
       'Busco el periodo hasta
       StrSql = "SELECT * FROM periodo WHERE pliqnro = " & pliqhasta
       OpenRecordset StrSql, objRs
        
       If Not objRs.EOF Then
          FechaHasta = objRs!pliqhasta
          descHasta = objRs!pliqDesc
       Else
          Flog.writeline "No se encontro el periodo hasta."
          Exit Sub
       End If
        
       objRs.Close
       
       'Cargo la configuracion del reporte
       Flog.writeline "Cargo la Configuración del Reporte"
       Call CargarConfiguracionReporte(Modelo)
      
       'Obtengo los empleados sobre los que tengo que generar los recibos
       Flog.writeline "Cargo los Empleados "
       Call CargarEmpleados(NroProceso, rsEmpl)
       
       'Guardo en la BD el encabezado
       Flog.writeline "Genero el encabezado del Reporte"
       Call GenerarEncabezadoReporte
       
       StrSql = "UPDATE batch_proceso SET bprcprogreso = 0 " & _
                   ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                   ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
       
       objConn.Execute StrSql, , adExecuteNoRecords
       prog = 0
       If (rsEmpl.RecordCount <> 0) Then
            Progreso = 100 / rsEmpl.RecordCount
       End If
        'sebastian stremel
        Do Until rsEmpl.EOF
            
            StrSql = "SELECT pliqnro FROM periodo WHERE "
            StrSql = StrSql & " pliqdesde >= " & ConvFecha(FechaDesde)
            StrSql = StrSql & " AND pliqhasta <= " & ConvFecha(FechaHasta)
            
            OpenRecordset StrSql, rsPeriodos
            
            EmpErrores = False
            Ternro = rsEmpl!Ternro
            orden = rsEmpl!estado
            
            'armo una lista con los periodos
            listaPer = "0"
            Do Until rsPeriodos.EOF
                listaPer = listaPer & ", " & rsPeriodos!pliqNro
            rsPeriodos.MoveNext
            Loop
            'hasta aca
            
            Call GenerarDatosEmpleadoPeriodo(listapronro, listaPer, Ternro, orden)
            
            'Actualizo el estado del proceso
            TiempoAcumulado = GetTickCount
            
            cantRegistros = cantRegistros - 1
            
            StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Fix(((totalEmpleados - cantRegistros) * 100) / totalEmpleados) & _
                     ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                     ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
               
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
            
            rsEmpl.MoveNext
            rsPeriodos.Close

        
        
        Loop
        'hasta aca
       
       
       If CInt(Formato) = 1 Then
           '-------------------------------------------------------------------
           'Genero la salida para el formato Procesos
           
           'Genero por cada empleado un registro
           Do Until rsEmpl.EOF
              arrpronro = Split(listapronro, ",")
              EmpErrores = False
              Ternro = rsEmpl!Ternro
              orden = rsEmpl!estado
              
              'Genero una entrada para el empleado por cada proceso
              For I = 0 To UBound(arrpronro)
                 proNro = arrpronro(I)
                 Flog.writeline "Generando datos empleado " & Ternro & " para el proceso " & proNro
                 
                 Call GenerarDatosEmpleadoProceso(proNro, Ternro, orden)
                 
              Next
              
              'Actualizo el estado del proceso
              TiempoAcumulado = GetTickCount
              
              cantRegistros = cantRegistros - 1
              
              StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Fix(((totalEmpleados - cantRegistros) * 100) / totalEmpleados) & _
                       ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                       ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
                 
              objconnProgreso.Execute StrSql, , adExecuteNoRecords
              
              'Si se generaron todos los datos del empleado correctamente lo borro
              If Not EmpErrores Then
                  StrSql = " DELETE FROM batch_empleado "
                  StrSql = StrSql & " WHERE bpronro = " & NroProceso
                  StrSql = StrSql & " AND ternro = " & Ternro
        
                  objConn.Execute StrSql, , adExecuteNoRecords
              End If
              
              rsEmpl.MoveNext
           Loop
        
        Else
        
           '-------------------------------------------------------------------
           'Genero la salida para el formato Periodos
        
           'Genero por cada empleado un registro
           Do Until rsEmpl.EOF
              
              StrSql = "SELECT pliqnro FROM periodo WHERE "
              StrSql = StrSql & " pliqdesde >= " & ConvFecha(FechaDesde)
              StrSql = StrSql & " AND pliqhasta <= " & ConvFecha(FechaHasta)
              
              OpenRecordset StrSql, rsPeriodos
              
              EmpErrores = False
              Ternro = rsEmpl!Ternro
              orden = rsEmpl!estado
              
              'Genero una entrada para el empleado por cada periodo
              Do Until rsPeriodos.EOF
                 Flog.writeline "Generando datos empleado " & Ternro & " para el periodo " & rsPeriodos!pliqNro
                 
                 Call GenerarDatosEmpleadoPeriodo(listapronro, rsPeriodos!pliqNro, Ternro, orden)
              
                 rsPeriodos.MoveNext
              Loop
              
              rsPeriodos.Close
              
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
                  StrSql = StrSql & " AND ternro = " & Ternro
        
                  objConn.Execute StrSql, , adExecuteNoRecords
              End If
              
              rsEmpl.MoveNext
           Loop
        
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
    
    Flog.writeline "Fin :" & Now
    Flog.Close

    Exit Sub
    
CE:
    HuboErrores = True
    Flog.writeline
    Flog.writeline " Error: " & Err.Description & Now
    Flog.writeline " Ultimo SQL: " & StrSql
    Flog.writeline
    
End Sub

Function controlNull(Str)
  If Trim(Str) = "" Then
     controlNull = "null"
  Else
     controlNull = "'" & Str & "'"
  End If
End Function


'--------------------------------------------------------------------
' Se encarga de generar los datos para el empleado por cada proceso
'--------------------------------------------------------------------

Sub GenerarDatosEmpleadoProceso(proNro, Ternro, orden)

Dim StrSql As String
Dim rsConsult As New ADODB.Recordset

'Variables donde se guardan los datos del INSERT final

Dim Legajo As Long
Dim apellido As String
Dim apellido2 As String
Dim nombre As String
Dim nombre2 As String

Dim estrnomb1
Dim estrnomb2
Dim estrnomb3
Dim proDesc
Dim pliqDesc
Dim pliqNro
Dim pliqFecha
Dim I

On Error GoTo MError

estrnomb1 = ""
estrnomb2 = ""
estrnomb3 = ""

'------------------------------------------------------------------
'Controlo si el empleado esta en el proceso
'------------------------------------------------------------------
StrSql = " SELECT * "
StrSql = StrSql & " FROM cabliq"
StrSql = StrSql & " WHERE empleado = " & Ternro
StrSql = StrSql & "   AND pronro   = " & proNro

OpenRecordset StrSql, rsConsult

If rsConsult.EOF Then
   'Si el empleado no esta en el proceso entonces paso al proximo
   rsConsult.Close
   
   Exit Sub
End If

rsConsult.Close

'------------------------------------------------------------------
'Busco los datos del empleado
'------------------------------------------------------------------
Flog.writeline "Buscando datos del empleado"
StrSql = " SELECT empleg,terape,terape2,ternom,ternom2,empfaltagr,empremu "
StrSql = StrSql & " FROM empleado "
StrSql = StrSql & " WHERE ternro= " & Ternro
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
Else
   Flog.writeline "Error al obtener los datos del empleado"
'   GoTo MError
End If

rsConsult.Close

'------------------------------------------------------------------
'Busco los datos del tipos de estructura 1
'------------------------------------------------------------------

'---LOG---
Flog.writeline "Buscando datos estructura 1"

If tenro1 <> 0 Then
    
    StrSql = " SELECT estrdabr "
    StrSql = StrSql & " FROM his_estructura "
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
    StrSql = StrSql & "    AND his_estructura.ternro = " & Ternro & " AND his_estructura.tenro = " & tenro1
    StrSql = StrSql & "    AND (htetdesde<=" & ConvFecha(fecEstr) & " AND (htethasta is null or htethasta>=" & ConvFecha(fecEstr) & "))"
    
    If estrnro1 <> 0 Then
        StrSql = StrSql & " AND estructura.estrnro = " & estrnro1
    End If
               
    OpenRecordset StrSql, rsConsult
    
    If Not rsConsult.EOF Then
       estrnomb1 = rsConsult!estrdabr
    End If
End If

'------------------------------------------------------------------
'Busco los datos del tipos de estructura 2
'------------------------------------------------------------------

'---LOG---
Flog.writeline "Buscando datos estructura 2"

If tenro2 <> 0 Then
    
    StrSql = " SELECT estrdabr "
    StrSql = StrSql & " FROM his_estructura "
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
    StrSql = StrSql & "    AND his_estructura.ternro = " & Ternro & " AND his_estructura.tenro = " & tenro2
    StrSql = StrSql & "    AND (htetdesde<=" & ConvFecha(fecEstr) & " AND (htethasta is null or htethasta>=" & ConvFecha(fecEstr) & "))"
    
    If estrnro2 <> 0 Then
        StrSql = StrSql & " AND estructura.estrnro = " & estrnro2
    End If
           
    OpenRecordset StrSql, rsConsult
    
    If Not rsConsult.EOF Then
       estrnomb2 = rsConsult!estrdabr
    End If
End If

'------------------------------------------------------------------
'Busco los datos del tipos de estructura 3
'------------------------------------------------------------------

'---LOG---
Flog.writeline "Buscando datos estructura 3"

If tenro3 <> 0 Then
    
    StrSql = " SELECT estrdabr "
    StrSql = StrSql & " FROM his_estructura "
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
    StrSql = StrSql & "    AND his_estructura.ternro = " & Ternro & " AND his_estructura.tenro = " & tenro3
    StrSql = StrSql & "    AND (htetdesde<=" & ConvFecha(fecEstr) & " AND (htethasta is null or htethasta>=" & ConvFecha(fecEstr) & "))"
    
    If estrnro3 <> 0 Then
        StrSql = StrSql & " AND estructura.estrnro = " & estrnro3
    End If
    
    OpenRecordset StrSql, rsConsult
    
    If Not rsConsult.EOF Then
       estrnomb3 = rsConsult!estrdabr
    End If
End If

'------------------------------------------------------------------
'Busco los datos del proceso
'------------------------------------------------------------------
Flog.writeline "Busco los datos del proceso"
StrSql = " SELECT * FROM proceso WHERE pronro = " & proNro
OpenRecordset StrSql, rsConsult
proDesc = ""
If Not rsConsult.EOF Then
   proDesc = rsConsult!proDesc
   pliqNro = rsConsult!pliqNro
End If
rsConsult.Close

'------------------------------------------------------------------
'Busco los datos del periodo
'------------------------------------------------------------------
Flog.writeline "Busco los datos del periodo"
StrSql = " SELECT * FROM periodo WHERE pliqnro = " & pliqNro
OpenRecordset StrSql, rsConsult
pliqDesc = ""
If Not rsConsult.EOF Then
   pliqDesc = rsConsult!pliqDesc
   pliqFecha = rsConsult!pliqdesde
End If
rsConsult.Close

'------------------------------------------------------------------
'Busco los valores de los conceptos y acumuladores
'------------------------------------------------------------------
Flog.writeline "Busco los valores de los conceptos y acumuladores"
StrSql = " SELECT 'CO', detliq.concnro, detliq.dlicant, detliq.dlimonto  "
StrSql = StrSql & " FROM cabliq "
StrSql = StrSql & " INNER JOIN detliq   ON cabliq.cliqnro = detliq.cliqnro  AND cabliq.empleado = " & Ternro & " AND cabliq.pronro = " & proNro
StrSql = StrSql & " AND ( 1=0 "
For I = 1 To Nro_Col
   If TipoCols(I) = "CO" Then
        StrSql = StrSql & " OR detliq.concnro = " & CodCols(I)
   End If
Next
StrSql = StrSql & " ) "
    
StrSql = StrSql & " UNION "
StrSql = StrSql & " SELECT 'AC', acu_liq.acunro, acu_liq.alcant, acu_liq.almonto "
StrSql = StrSql & " FROM cabliq "
StrSql = StrSql & " INNER JOIN acu_liq  ON cabliq.cliqnro = acu_liq.cliqnro AND cabliq.empleado = " & Ternro & " AND cabliq.pronro = " & proNro
StrSql = StrSql & " AND ( 1=0 "
For I = 1 To Nro_Col
    If TipoCols(I) = "AC" Then
       StrSql = StrSql & " OR acu_liq.acunro = " & CodCols(I)
    End If
Next
StrSql = StrSql & " ) "

'Obtengo los datos de los conceptos y acumuladores
Flog.writeline "borrarValores"
Call borrarValores

OpenRecordset StrSql, rsConsult
Do Until rsConsult.EOF
   Call agregarValor(rsConsult(0), rsConsult(1), rsConsult(3), rsConsult(2))
   rsConsult.MoveNext
Loop
rsConsult.Close

'------------------------------------------------------------------
'Busco los valores de los tipo estructura - Descripción
'------------------------------------------------------------------
Call borrarChar
Flog.writeline "borrarChar"
Flog.writeline "Busco los valores de los tipo estructura - Descripción"
For I = 1 To Nro_Col
   If TipoCols(I) = "TE" Then
      StrSql = " SELECT 'TE', estrdabr "
      StrSql = StrSql & " FROM his_estructura "
      StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
      StrSql = StrSql & "    AND his_estructura.ternro = " & Ternro
      StrSql = StrSql & "    AND (htetdesde<=" & ConvFecha(fecEstr) & " AND (htethasta is null or htethasta>=" & ConvFecha(fecEstr) & "))"
      StrSql = StrSql & " AND his_estructura.tenro = " & CodCols(I)
      OpenRecordset StrSql, rsConsult
      Do Until rsConsult.EOF
         Call agregarChar(I, rsConsult(1))
         rsConsult.MoveNext
      Loop
      rsConsult.Close
   End If
Next


'------------------------------------------------------------------
'Busco los valores de los tipos de fechas
'------------------------------------------------------------------
Flog.writeline "Busco los valores de los tipos de fechas"
For I = 1 To Nro_Col
    If TipoCols(I) = "TF" Then
        Select Case CodCols(I)
        Case 1: 'Fecha de nacimiento
            StrSql = "SELECT 'TF', terfecnac FROM tercero "
            StrSql = StrSql & " WHERE ternro = " & Ternro
        Case 2: 'Fecha de alta reconocida
            StrSql = "SELECT 'TF', altfec FROM fases "
            StrSql = StrSql & " WHERE empleado = " & Ternro
            StrSql = StrSql & " AND fasrecofec = -1 "
            StrSql = StrSql & " ORDER BY altfec "
        Case 3: 'Fecha fase mas antigua
            StrSql = "SELECT 'TF', altfec FROM fases "
            StrSql = StrSql & " WHERE empleado = " & Ternro
            StrSql = StrSql & " ORDER BY altfec "
        Case 4: 'fecha fase mas nueva
            StrSql = "SELECT 'TF', altfec FROM fases "
            StrSql = StrSql & " WHERE empleado = " & Ternro
            StrSql = StrSql & " ORDER BY altfec desc "
        Case 5: 'fecha baja
            StrSql = "SELECT 'TF', bajfec FROM fases "
            StrSql = StrSql & " WHERE empleado = " & Ternro
            StrSql = StrSql & " ORDER BY bajfec desc "
        End Select
        OpenRecordset StrSql, rsConsult
        
      If Not rsConsult.EOF Then
         Call agregarChar(I, rsConsult(1))
      End If
      rsConsult.Close
   End If
Next


'------------------------------------------------------------------
'Busco los valores de los tipo de documentos
'------------------------------------------------------------------
Flog.writeline "Busco los valores de los tipo de documentos"
For I = 1 To Nro_Col

   If TipoCols(I) = "TD" Then
   
      If CodCols(I) = 1 Then
         StrSql = " SELECT 'TD', nrodoc "
         StrSql = StrSql & " FROM ter_doc "
         StrSql = StrSql & " WHERE ter_doc.ternro = " & Ternro
         StrSql = StrSql & " AND ter_doc.tidnro <= 4 "
      Else
         StrSql = " SELECT 'TD', nrodoc "
         StrSql = StrSql & " FROM ter_doc "
         StrSql = StrSql & " WHERE ter_doc.ternro = " & Ternro
         StrSql = StrSql & " AND ter_doc.tidnro = " & CodCols(I)
      End If
      
      OpenRecordset StrSql, rsConsult

      Do Until rsConsult.EOF
         Call agregarChar(I, rsConsult(1))
         rsConsult.MoveNext
      Loop
      rsConsult.Close
   End If
Next


'------------------------------------------------------------------
'Obtengo la fecha desde y hasta del periodo
'------------------------------------------------------------------
Flog.writeline "Obtengo la fecha desde y hasta del periodo"
Dim FechaDesde As Date
Dim FechaHasta As Date
Dim AnioProceso As Integer
Dim MesProceso As Integer
Dim dias As Integer
Dim Aux_Fecha_Desde As Date
Dim Aux_Fecha_Hasta As Date


StrSql = " SELECT profecini, profecfin "
StrSql = StrSql & " FROM proceso "
StrSql = StrSql & " WHERE proceso.pronro = " & proNro
            
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   FechaDesde = rsConsult!profecini
   FechaHasta = rsConsult!profecfin
   MesProceso = Month(rsConsult!profecini)
   AnioProceso = Year(rsConsult!profecini)
End If
rsConsult.Close

'------------------------------------------------------------------
'Busco los valores de las novedades individuales
'------------------------------------------------------------------
Flog.writeline "Busco los valores de las novedades individuales"
StrSql = " SELECT 'NOV', SUM(nevalor), novemp.concnro, novemp.tpanro "
StrSql = StrSql & " FROM novemp "
StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = novemp.concnro "
StrSql = StrSql & " WHERE novemp.empleado = " & Ternro
StrSql = StrSql & " AND ((novemp.nedesde >= " & ConvFecha(FechaDesde)
StrSql = StrSql & " AND  (novemp.nehasta <= " & ConvFecha(FechaHasta) & " OR novemp.nehasta IS NULL))"
StrSql = StrSql & " OR   novemp.nevigencia = 0 )"
StrSql = StrSql & " AND ( 1=0 "
For I = 1 To Nro_Col
   CodNov = Split(CodNovCols(I), "@")
   If TipoCols(I) = "NOV" Then
      StrSql = StrSql & " OR (concepto.concnro = " & CodNov(0) & " AND novemp.tpanro = " & CodNov(1) & ")"
   End If
Next
StrSql = StrSql & " ) "
StrSql = StrSql & " GROUP BY novemp.concnro, novemp.tpanro "
OpenRecordset StrSql, rsConsult

Do Until rsConsult.EOF
   Call agregarValorNov(rsConsult(0), rsConsult(2) & "@" & rsConsult(3), rsConsult(1))
   rsConsult.MoveNext
Loop
rsConsult.Close

'------------------------------------------------------------------
'Busco los valores de las novedades ajuste
'------------------------------------------------------------------
Flog.writeline "Busco los valores de las novedades de ajuste"
StrSql = " SELECT 'NAJ', SUM(navalor), novaju.concnro "
StrSql = StrSql & " FROM novaju "
StrSql = StrSql & " WHERE novaju.empleado = " & Ternro
StrSql = StrSql & " AND ((novaju.nadesde >= " & ConvFecha(FechaDesde)
StrSql = StrSql & " AND  (novaju.nahasta <= " & ConvFecha(FechaHasta) & " OR novaju.nahasta IS NULL))"
StrSql = StrSql & " OR   novaju.navigencia = 0 )"
StrSql = StrSql & " AND ( 1=0 "
For I = 1 To Nro_Col
   If TipoCols(I) = "NAJ" Then
      StrSql = StrSql & " OR (novaju.concnro = " & CodCols(I) & ")"
   End If
Next
StrSql = StrSql & " ) "
StrSql = StrSql & " GROUP BY novaju.concnro "
OpenRecordset StrSql, rsConsult
Do Until rsConsult.EOF
   Call agregarValor(rsConsult(0), rsConsult(2), rsConsult(1), 0)
   rsConsult.MoveNext
Loop
rsConsult.Close

'------------------------------------------------------------------
'Busco los valores de las licencias
'------------------------------------------------------------------
Flog.writeline "Busco los valores de las licencias"
'StrSql = " SELECT 'LIC', SUM(elcantdias), emp_lic.tdnro "
'StrSql = StrSql & " FROM emp_lic "
'StrSql = StrSql & " WHERE emp_lic.empleado = " & ternro
'StrSql = StrSql & " AND (emp_lic.elfechadesde >= " & ConvFecha(FechaDesde)
'StrSql = StrSql & " AND  emp_lic.elfechahasta <= " & ConvFecha(FechaHasta) & ")"
'StrSql = StrSql & " AND ( 1=0 "
'For i = 1 To Nro_Col
'   If TipoCols(i) = "LIC" Then
'      StrSql = StrSql & " OR (emp_lic.tdnro = " & CodCols(i) & ")"
'   End If
'Next
'StrSql = StrSql & " ) "
'StrSql = StrSql & " GROUP BY emp_lic.tdnro "
'OpenRecordset StrSql, rsConsult
'Do Until rsConsult.EOF
'   Call agregarValor(rsConsult(0), rsConsult(2), rsConsult(1), 0)
'   rsConsult.MoveNext
'Loop
'rsConsult.Close

'Martin Ferraro - 13/03/2006 - nueva version
StrSql = "SELECT 'LIC', elcantdias, emp_lic.tdnro, elfechadesde, elfechahasta "
StrSql = StrSql & " FROM emp_lic WHERE (empleado = " & Ternro
StrSql = StrSql & " AND elfechadesde <=" & ConvFecha(FechaHasta)
StrSql = StrSql & " AND elfechahasta >= " & ConvFecha(FechaDesde)
StrSql = StrSql & " AND ( 1=0 "
For I = 1 To Nro_Col
   If TipoCols(I) = "LIC" Then
      StrSql = StrSql & " OR (emp_lic.tdnro = " & CodCols(I) & ") "
   End If
Next
StrSql = StrSql & " ) )"
'StrSql = StrSql & " GROUP BY emp_lic.tdnro, elcantdias "
OpenRecordset StrSql, rsConsult
dias = 0
Do While Not rsConsult.EOF
    Aux_Fecha_Desde = rsConsult!elfechadesde
    Aux_Fecha_Hasta = rsConsult!elfechahasta

    If Aux_Fecha_Hasta > FechaHasta Then
        Aux_Fecha_Hasta = FechaHasta
    End If
    dias = CantidadDeDias(FechaDesde, FechaHasta, Aux_Fecha_Desde, Aux_Fecha_Hasta)
    
    Call agregarValor(rsConsult(0), rsConsult(2), dias, 0)
    rsConsult.MoveNext
Loop
rsConsult.Close

'------------------------------------------------------------------
'Busco los valores de los préstamos
'------------------------------------------------------------------
Flog.writeline "Busco los valores de los préstamos"
StrSql = " SELECT 'PRE', SUM(cuototal), prestamo.lnprenro "
StrSql = StrSql & " FROM pre_cuota "
StrSql = StrSql & " INNER JOIN prestamo ON prestamo.prenro = pre_cuota.prenro "
StrSql = StrSql & " WHERE prestamo.ternro = " & Ternro
StrSql = StrSql & " AND pre_cuota.cuomes = " & MesProceso
StrSql = StrSql & " AND  pre_cuota.cuoano = " & AnioProceso
StrSql = StrSql & " AND ( 1=0 "
For I = 1 To Nro_Col
   If TipoCols(I) = "PRE" Then
      StrSql = StrSql & " OR (prestamo.lnprenro = " & CodCols(I) & ")"
   End If
Next
StrSql = StrSql & " ) "
StrSql = StrSql & " GROUP BY prestamo.lnprenro "
OpenRecordset StrSql, rsConsult
Do Until rsConsult.EOF
   Call agregarValor(rsConsult(0), rsConsult(2), rsConsult(1), 0)
   rsConsult.MoveNext
Loop
rsConsult.Close

'------------------------------------------------------------------
'Busco los valores de los embargos
'------------------------------------------------------------------
Flog.writeline "Busco los valores de los embargos"
StrSql = " SELECT 'EMB', SUM(embcimp), embargo.tpenro "
StrSql = StrSql & " FROM embcuota "
StrSql = StrSql & " INNER JOIN embargo ON embargo.embnro = embcuota.embnro "
StrSql = StrSql & " WHERE embargo.ternro = " & Ternro
StrSql = StrSql & " AND embcuota.embcmes = " & MesProceso
StrSql = StrSql & " AND  embcuota.embcanio = " & AnioProceso
StrSql = StrSql & " AND ( 1=0 "
For I = 1 To Nro_Col
   If TipoCols(I) = "EMB" Then
      StrSql = StrSql & " OR (embargo.tpenro = " & CodCols(I) & ")"
   End If
Next
StrSql = StrSql & " ) "
StrSql = StrSql & " GROUP BY embargo.tpenro "
OpenRecordset StrSql, rsConsult
Do Until rsConsult.EOF
   Call agregarValor(rsConsult(0), rsConsult(2), rsConsult(1), 0)
   rsConsult.MoveNext
Loop
rsConsult.Close

'------------------------------------------------------------------
'Busco los valores de los vales
'------------------------------------------------------------------
Flog.writeline "Busco los valores de los vales"
StrSql = " SELECT 'VAL', SUM(valmonto), vales.tvalenro "
StrSql = StrSql & " FROM vales "
StrSql = StrSql & " WHERE vales.empleado = " & Ternro
StrSql = StrSql & " AND vales.valfecped >= " & ConvFecha(FechaDesde)
StrSql = StrSql & " AND vales.valfecped <= " & ConvFecha(FechaHasta)
StrSql = StrSql & " AND ( 1=0 "
For I = 1 To Nro_Col
   If TipoCols(I) = "VAL" Then
      StrSql = StrSql & " OR (vales.tvalenro = " & CodCols(I) & ")"
   End If
Next
StrSql = StrSql & " ) "
StrSql = StrSql & " GROUP BY vales.tvalenro "
OpenRecordset StrSql, rsConsult
Do Until rsConsult.EOF
   Call agregarValor(rsConsult(0), rsConsult(2), rsConsult(1), 0)
   rsConsult.MoveNext
Loop
rsConsult.Close

'------------------------------------------------------------------
'Busco los valores de la Dirección
'------------------------------------------------------------------
Flog.writeline "Busco los valores de la Dirección"
'Call borrarChar
Dim TipoDomi
Dim Datos(8) As String
Dim j

For I = 1 To Nro_Col
    If TipoCols(I) = "DIR" Then
        CodNov = Split(CodNovCols(I), "@")
        TipoDomi = CodNov(0)
        
        'Calle, Nro, Piso, Dpto, Localidad, Provincia, País
        StrSql = " SELECT 'DIR', calle, nro, piso, oficdepto, locdesc, provdesc, paisdesc "
        StrSql = StrSql & " FROM cabdom "
        StrSql = StrSql & " INNER JOIN detdom ON detdom.domnro = cabdom.domnro "
        StrSql = StrSql & " INNER JOIN localidad ON localidad.locnro = detdom.locnro "
        StrSql = StrSql & " INNER JOIN provincia ON provincia.provnro = detdom.provnro "
        StrSql = StrSql & " INNER JOIN pais ON pais.paisnro = detdom.paisnro "
        StrSql = StrSql & " WHERE cabdom.ternro = " & Ternro & " AND cabdom.tidonro = " & TipoDomi
            
        OpenRecordset StrSql, rsConsult
        
        If Not rsConsult.EOF Then
           For j = 1 To 7
               If IsNull(rsConsult(j)) Then
                  Datos(j) = ""
               Else
                  Datos(j) = rsConsult(j)
               End If
           Next
           Call agregarValorDir(I, CodNov(1), Datos)
        End If
        rsConsult.Close
    End If
Next

'------------------------------------------------------------------
'Busco los valores de las cuentas bancarias
'------------------------------------------------------------------
Flog.writeline "Busco los valores de las cuentas bancarias"
For I = 1 To Nro_Col

   If TipoCols(I) = "CTA" Then
   
      StrSql = " SELECT 'CTA', ctabnro "
      StrSql = StrSql & " FROM ctabancaria "
      StrSql = StrSql & " WHERE ctabancaria.ternro = " & Ternro
      If CodCols(I) = -1 Then
         StrSql = StrSql & " AND ctabancaria.ctabestado = -1 "
      Else
         StrSql = StrSql & " AND ctabancaria.fpagnro = " & CodCols(I)
      End If
      OpenRecordset StrSql, rsConsult

      Do Until rsConsult.EOF
         Call agregarChar(I, rsConsult(1))
         rsConsult.MoveNext
      Loop
      rsConsult.Close
   End If
Next

'------------------------------------------------------------------
'Busco los Tipo Sigla
'------------------------------------------------------------------
Flog.writeline "Busco los valores de las cuentas bancarias"
For I = 1 To Nro_Col

   If TipoCols(I) = "TIPSIG" Then
         StrSql = " SELECT tipodocu.tidsigla "
         StrSql = StrSql & " From Tercero"
         StrSql = StrSql & " LEFT JOIN ter_doc docu ON (docu.ternro= tercero.ternro and docu.tidnro>0 and docu.tidnro<5) "
         StrSql = StrSql & " LEFT JOIN tipodocu ON tipodocu.tidnro= docu.tidnro "
         StrSql = StrSql & " WHERE tercero.ternro = " & Ternro
        
         'StrSql = "select tipodocu.tidsigla from empleado "
         'StrSql = StrSql & "inner join ter_doc on ter_doc.ternro=empleado.ternro "
         'StrSql = StrSql & "inner join tipodocu on tipodocu.tidnro=ter_doc.tidnro "
         'StrSql = StrSql & "Where Empleado.Ternro = " & Ternro
         OpenRecordset StrSql, rsConsult

        Do Until rsConsult.EOF
             Call agregarChar(I, rsConsult(0))
             rsConsult.MoveNext
        Loop
         rsConsult.Close
   End If
Next


'------------------------------------------------------------------
'Busco los valores de las datos
'------------------------------------------------------------------
Flog.writeline "Busco los valores de los datos"
For I = 1 To Nro_Col

   If TipoCols(I) = "DAT" Then
   
      Select Case CodCols(I)
        Case 1: 'Causa Baja
            StrSql = "SELECT 'DAT', caudes FROM fases "
            StrSql = StrSql & " INNER JOIN causa ON causa.caunro = fases.caunro "
            StrSql = StrSql & " WHERE empleado = " & Ternro
            StrSql = StrSql & " AND fases.bajfec <= " & ConvFecha(FechaHasta)
            StrSql = StrSql & " ORDER BY bajfec DESC "
        Case 2: 'Email Interno
            StrSql = "SELECT 'DAT', empemail FROM empleado "
            StrSql = StrSql & " WHERE ternro = " & Ternro
        Case 3: 'Estado del Empleado
            StrSql = "SELECT 'DAT', empest FROM empleado "
            StrSql = StrSql & " WHERE ternro = " & Ternro
        Case 4: 'Estado Civil
            StrSql = "SELECT 'DAT', estcivdesabr FROM tercero "
            StrSql = StrSql & " INNER JOIN estcivil ON estcivil.estcivnro = tercero.estcivnro "
            StrSql = StrSql & " WHERE ternro = " & Ternro
        Case 5: 'Nacionalidad
            StrSql = "SELECT 'DAT', nacionaldes FROM tercero "
            StrSql = StrSql & " INNER JOIN nacionalidad ON nacionalidad.nacionalnro = tercero.nacionalnro "
            StrSql = StrSql & " WHERE ternro = " & Ternro
        Case 6: 'Reporta A
            StrSql = "SELECT 'DAT', e2.empleg, e2.terape, e2.terape2, e2.ternom, e2.ternom2 "
            StrSql = StrSql & " FROM empleado e1 "
            StrSql = StrSql & " INNER JOIN empleado e2 ON e2.ternro = e1.empreporta  "
            StrSql = StrSql & " WHERE e1.ternro = " & Ternro
        Case 7: 'Sexo
            StrSql = "SELECT 'DAT', tersex FROM tercero "
            StrSql = StrSql & " WHERE ternro = " & Ternro
      End Select
      OpenRecordset StrSql, rsConsult

      Do Until rsConsult.EOF
         If CodCols(I) = 3 Then
            If rsConsult(1) = "-1" Then
               Call agregarChar(I, "Activo")
            Else
               Call agregarChar(I, "Inactivo")
            End If
         ElseIf CodCols(I) = 7 Then
                If rsConsult(1) = "-1" Then
                   Call agregarChar(I, "Masculino")
                Else
                   Call agregarChar(I, "Femenino")
                End If
         ElseIf CodCols(I) = 6 Then
                Call agregarChar(I, rsConsult(1) & " - " & rsConsult(2) & " " & rsConsult(3) & ", " & rsConsult(4) & " " & rsConsult(5))
         Else
            Call agregarChar(I, rsConsult(1))
         End If
         
         rsConsult.MoveNext
      Loop
      rsConsult.Close
   End If
Next

'------------------------------------------------------------------
'Busco los valores de los tipo estructura - Código Externo
'------------------------------------------------------------------
Flog.writeline "Busco los valores de los tipo estructura - Código Externo"
For I = 1 To Nro_Col

   If TipoCols(I) = "TCE" Then
   
      StrSql = " SELECT 'TCE', estrcodext "
      StrSql = StrSql & " FROM his_estructura "
      StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
      StrSql = StrSql & "    AND his_estructura.ternro = " & Ternro
      StrSql = StrSql & "    AND (htetdesde<=" & ConvFecha(fecEstr) & " AND (htethasta is null or htethasta>=" & ConvFecha(fecEstr) & "))"
      StrSql = StrSql & " AND his_estructura.tenro = " & CodCols(I)
            
      OpenRecordset StrSql, rsConsult

      Do Until rsConsult.EOF
         Call agregarChar(I, rsConsult(1))
         rsConsult.MoveNext
      Loop
      rsConsult.Close
   End If
Next

'------------------------------------------------------------------
'Busco la edad del empleado
'------------------------------------------------------------------
Flog.writeline "Busco la edad del empleado"
Dim Edad As Long
Dim FechaNacimiento As String
Dim FechaInicio As Date

For I = 1 To Nro_Col

   If TipoCols(I) = "EDA" Then

      If CodCols(I) = 1 Then
         FechaInicio = FechaDesde
      ElseIf CodCols(I) = 2 Then
             FechaInicio = FechaDesde
      ElseIf CodCols(I) = 3 Then
             FechaInicio = FechaHasta
      ElseIf CodCols(I) = 4 Then
             FechaInicio = FechaHasta
      End If
      
      StrSql = " SELECT terfecnac "
      StrSql = StrSql & " FROM tercero "
      StrSql = StrSql & " WHERE tercero.ternro = " & Ternro
            
      OpenRecordset StrSql, rsConsult

      If Not rsConsult.EOF Then
         FechaNacimiento = rsConsult(0)
      End If

      If IsNull(FechaNacimiento) Or FechaNacimiento = "" Then
         Edad = 0
      Else
           If (Month(FechaInicio) > Month(CDate(FechaNacimiento))) Then
               Edad = DateDiff("yyyy", CDate(FechaNacimiento), FechaInicio)
           Else
               If (Month(FechaInicio) = Month(CDate(FechaNacimiento))) And (Day(FechaInicio) >= Day(CDate(FechaNacimiento))) Then
                  Edad = DateDiff("yyyy", CDate(FechaNacimiento), FechaInicio)
               Else
                  Edad = DateDiff("yyyy", CDate(FechaNacimiento), FechaInicio) - 1
               End If
           End If
      End If
      rsConsult.Close
      
      Call agregarValorEdad(I, Edad)
   End If
Next

'------------------------------------------------------------------
'Busco la antiguedad del empleado
'------------------------------------------------------------------
Flog.writeline "Busco la antiguedad del empleado"
Dim Texto As String
Dim antdia As Integer
Dim antmes As Integer
Dim antanio As Integer
Dim q As Integer

For I = 1 To Nro_Col

   If TipoCols(I) = "ANT" Then

      If CodCols(I) = 1 Then
            FechaInicio = FechaDesde
      ElseIf CodCols(I) = 2 Then
             FechaInicio = FechaDesde
      ElseIf CodCols(I) = 3 Then
             FechaInicio = FechaHasta
      ElseIf CodCols(I) = 4 Then
             FechaInicio = FechaHasta
      ElseIf CodCols(I) = 5 Then
             FechaInicio = C_Date(fecEstr) 'ConvFecha(fecEstr)
      End If

      'Calcula la antiguedad en dias, meses y años
      Call Antiguedad(Ternro, "REAL", FechaInicio, antdia, antmes, antanio, q)
      If antanio = 0 Then
         If antmes = 0 Then
            Texto = antdia & " día/s."
         Else
            Texto = antmes & " mes/es "
            If antdia <> 0 Then
               Texto = Texto & antdia & " día/s."
            End If
         End If
      Else
          Texto = antanio & " año/s "
          If antmes = 0 Then
             If antdia <> 0 Then
                Texto = Texto & antdia & " día/s."
             End If
          Else
             Texto = Texto & antmes & " mes/es "
             If antdia <> 0 Then
                Texto = Texto & antdia & " día/s."
             End If
          End If
      End If
      
      Call agregarChar(I, Texto)
      
   End If
Next

'------------------------------------------------------------------
'Busco los valores del telefono
'------------------------------------------------------------------
Flog.writeline "Busco los valores del telefono"
'Dim TipoDomi
Dim NombreCampo As String

For I = 1 To Nro_Col
    If TipoCols(I) = "TEL" Then
        CodNov = Split(CodNovCols(I), "@")
        TipoDomi = CodNov(0) '
        
        Select Case CodNov(1)
          Case 1: 'Telefono Principal
                 NombreCampo = "telefono.teldefault"
          Case 2: 'Telefono Celular
                 NombreCampo = "telefono.telcelular"
          Case 3: 'Telefono Fax
                 NombreCampo = "telefono.telfax"
        End Select
                    
        StrSql = " SELECT 'TEL', telnro "
        StrSql = StrSql & " FROM cabdom "
        StrSql = StrSql & " INNER JOIN telefono ON telefono.domnro = cabdom.domnro AND " & NombreCampo & " = -1 "
        StrSql = StrSql & " WHERE cabdom.ternro = " & Ternro & " AND cabdom.tidonro = " & TipoDomi
        
        OpenRecordset StrSql, rsConsult
        
        If Not rsConsult.EOF Then
          Call agregarChar(I, rsConsult(1))
        End If
        rsConsult.Close
   End If
Next

'------------------------------------------------------------------
'Busco los valores del Analisis Detallado
'------------------------------------------------------------------
Flog.writeline "Busco del Analisis Detallado"


StrSql = " SELECT distinct 'ADE', traza.concnro, traza.tpanro, traza.travalor "
StrSql = StrSql & " FROM cabliq "
StrSql = StrSql & " INNER JOIN proceso  ON cabliq.pronro = proceso.pronro  AND proceso.pliqnro = " & pliqNro & " AND cabliq.pronro = " & proNro
StrSql = StrSql & " INNER JOIN traza   ON cabliq.cliqnro = traza.cliqnro  AND cabliq.empleado = " & Ternro & " AND cabliq.pronro = proceso.pronro "
StrSql = StrSql & " AND ( 1=0 "
For I = 1 To Nro_Col
   CodNov = Split(CodNovCols(I), "@")
   If TipoCols(I) = "ADE" Then
      StrSql = StrSql & " OR (traza.concnro = " & CodNov(0) & " AND traza.tpanro = " & CodNov(1) & ")"
   End If
Next
StrSql = StrSql & " ) "
'StrSql = StrSql & " GROUP BY traza.concnro, traza.tpanro "

OpenRecordset StrSql, rsConsult

Do Until rsConsult.EOF
   Call agregarValorNov(rsConsult(0), rsConsult(1) & "@" & rsConsult(2), rsConsult(3))
   rsConsult.MoveNext
Loop
rsConsult.Close


'------------------------------------------------------------------
'Busco los valores del proceso
'------------------------------------------------------------------
Flog.writeline "Busco los valores del proceso"
Dim textoCampo As String

For I = 1 To Nro_Col
    If TipoCols(I) = "PRO" Then
        'CodNov = Split(CodNovCols(I), "@")
        'TipoDomi = CodNov(0) '
          
        Select Case CodCols(I)
            Case 1: 'Fecha de Inicio
                NombreCampo = "profecini"
            Case 2: 'Fecha de Fin
                NombreCampo = "profecfin"
            Case 3: 'Fecha Pago
                NombreCampo = "profecpago"
            Case 4: 'Fecha Planeada
                NombreCampo = "profecplan"
            Case 5: 'Descripcion
                NombreCampo = "prodesc"
            Case 6: 'Modelo
                NombreCampo = "tprocdesc"
            Case 7: 'Periodo MMM-AA
                NombreCampo = " pliqmes, pliqanio"
            Case 8: 'Periodo MM/AAAA
                NombreCampo = " pliqmes, pliqanio"
            Case 9: 'Periodo MM/AA
                NombreCampo = " pliqmes, pliqanio"
        End Select
        
        StrSql = " SELECT 'PRO', " & NombreCampo
        StrSql = StrSql & " FROM proceso "
        StrSql = StrSql & " INNER JOIN periodo ON proceso.pliqnro = periodo.pliqnro "
        StrSql = StrSql & " INNER JOIN tipoproc ON proceso.tprocnro = tipoproc.tprocnro "
        StrSql = StrSql & " WHERE proNro = " & proNro
        OpenRecordset StrSql, rsConsult
        
        If Not rsConsult.EOF Then
            Select Case CodCols(I)
                Case 1: 'Fecha de Inicio
                    textoCampo = rsConsult(1)
                Case 2: 'Fecha de Fin
                    textoCampo = rsConsult(1)
                Case 3: 'Fecha Pago
                    textoCampo = rsConsult(1)
                Case 4: 'Fecha Planeada
                    textoCampo = rsConsult(1)
                Case 5: 'Descripcion
                    textoCampo = rsConsult(1)
                Case 6: 'Modelo
                    textoCampo = rsConsult(1)
                Case 7: 'Periodo MMM-AA
                    textoCampo = Format(rsConsult(1) & "-" & rsConsult(2), "MMM-yy")
                Case 8: 'Periodo MM/AAAA
                    textoCampo = Format(rsConsult(1) & "-" & rsConsult(2), "MM-yyyy")
                Case 9: 'Periodo MM/AA
                    textoCampo = Format(rsConsult(1) & "-" & rsConsult(2), "MM-yy")
            End Select
          Call agregarChar(I, textoCampo)
        End If
        rsConsult.Close
   End If
Next



'-------------------------------------------------------------------------------
'Inserto los datos en la BD
'-------------------------------------------------------------------------------
Flog.writeline "Inserto los datos en la BD - Legajo = " & Legajo & " TERNRO " & Ternro & " proceso " & proNro & " periodo " & pliqNro
StrSql = " INSERT INTO rep_estadisticaEmpleado_det ( " & _
         " bpronro , legajo, ternro, apellido, apellido2, " & _
         " nombre , nombre2, pronro, pliqnro,pliqfecha, " & _
         " prodesc , pliqdesc, estrdabr1, estrdabr2, estrdabr3, " & _
         " colval1 , colval2, colval3, colval4, colval5, " & _
         " colval6 , colval7, colval8, colval9, colval10," & _
         " colval11 , colval12, colval13, colval14, colval15, " & _
         " colval16 , colval17, colval18, colval19, colval20, " & _
         " colval21 , colval22, colval23, colval24, colval25, " & _
         " colval26 , colval27, colval28, colval29, colval30, " & _
         " colval31 , colval32, colval33, colval34, colval35, " & _
         " colval36 , colval37, colval38, colval39, colval40, " & _
         " colval41 , colval42, colval43, colval44, colval45, " & _
         " colval46 , colval47, colval48, colval49, colval50, " & _
         " colchar1 , colchar2, colchar3, colchar4, colchar5, " & _
         " colchar6 , colchar7, colchar8, colchar9, colchar10," & _
         " colchar11 , colchar12, colchar13, colchar14, colchar15, " & _
         " colchar16 , colchar17, colchar18, colchar19, colchar20, " & _
         " colchar21 , colchar22, colchar23, colchar24, colchar25, " & _
         " colchar26 , colchar27, colchar28, colchar29, colchar30, " & _
         " colchar31 , colchar32, colchar33, colchar34, colchar35, " & _
         " colchar36 , colchar37, colchar38, colchar39, colchar40, " & _
         " colchar41 , colchar42, colchar43, colchar44, colchar45, " & _
         " colchar46 , colchar47, colchar48, colchar49, colchar50, " & _
         " orden ) VALUES ( "

        StrSql = StrSql & NroProceso & ","
        StrSql = StrSql & Legajo & ","
        StrSql = StrSql & Ternro & ","
        StrSql = StrSql & "'" & apellido & "',"
        StrSql = StrSql & "'" & apellido2 & "',"
        StrSql = StrSql & "'" & nombre & "',"
        StrSql = StrSql & "'" & nombre2 & "',"
        StrSql = StrSql & proNro & ","
        StrSql = StrSql & pliqNro & ","
        StrSql = StrSql & ConvFecha(pliqFecha) & ","
        StrSql = StrSql & "'" & proDesc & "',"
        StrSql = StrSql & "'" & pliqDesc & "',"
        StrSql = StrSql & controlNull(estrnomb1) & ","
        StrSql = StrSql & controlNull(estrnomb2) & ","
        StrSql = StrSql & controlNull(estrnomb3) & ","
        
        For I = 1 To 50
            StrSql = StrSql & numberForSQL(ValCols(I)) & ","
        Next
        
        For I = 1 To 50
            StrSql = StrSql & "'" & Mid(CharCols(I), 1, 60) & "',"
        Next
        
        StrSql = StrSql & orden & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
Exit Sub

MError:
    Flog.writeline "Error en empleado: " & Legajo & " Error: " & Err.Description
    Flog.writeline "SQL: " & StrSql
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
     
  If Not IsNull(Str) Then
     If Len(Str) = 0 Then
        numberForSQL = 0
     Else
        numberForSQL = Replace(Str, ",", ".")
     End If
  End If

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


Sub CargarConfiguracionReporte(ByVal Modelo As Long)

    Dim objRs As New ADODB.Recordset
    Dim StrSql As String
    Dim I
    Dim columnaActual
    
    StrSql = " SELECT * FROM repliqcolsEstadisticaEmpleados "
    StrSql = StrSql & " INNER JOIN repliqEstadisticaEmpleado ON repliqEstadisticaEmpleado.rliqnro = repliqcolsEstadisticaEmpleados.rliqnro "
    StrSql = StrSql & " WHERE repliqEstadisticaEmpleado.rliqnro=" & Modelo
    StrSql = StrSql & " ORDER BY rliqnrocol "
    
    OpenRecordset StrSql, objRs
    
    Nro_Col = 0
    CantColumnas = 0
    
    For I = 0 To 50
        TipoColumna(I) = 0
    Next
    
    Do Until objRs.EOF
       
       Nro_Col = Nro_Col + 1
       
       If CLng(objRs!rliqtipo) <> 7 And CLng(objRs!rliqtipo) <> 13 And CLng(objRs!rliqtipo) <> 18 And CLng(objRs!rliqtipo) <> 20 Then
          CodCols(Nro_Col) = objRs!rliqval
       Else
          CodNovCols(Nro_Col) = objRs!rliqval & "@" & objRs!rliqval2
       End If
       columnaActual = CLng(objRs!rliqnrocol)
       
       If CLng(objRs!rliqtipo) = 0 Then     'Monto Concepto
          TipoCols(Nro_Col) = "CO"
          TipoColumna(columnaActual) = 0
          EsMonto(Nro_Col) = True
       ElseIf CLng(objRs!rliqtipo) = 1 Then 'Monto Acumulador
           TipoCols(Nro_Col) = "AC"
           TipoColumna(columnaActual) = 1
           EsMonto(Nro_Col) = True
       ElseIf CLng(objRs!rliqtipo) = 2 Then 'Tipo Estructura por Descripción
           TipoCols(Nro_Col) = "TE"
           TipoColumna(columnaActual) = 2
       ElseIf CLng(objRs!rliqtipo) = 3 Then 'Tipo Docuemento
           TipoCols(Nro_Col) = "TD"
           TipoColumna(columnaActual) = 3
       ElseIf CLng(objRs!rliqtipo) = 4 Then 'Fecha
           TipoCols(Nro_Col) = "TF"
           TipoColumna(columnaActual) = 4
       ElseIf CLng(objRs!rliqtipo) = 5 Then 'Cantidad Concepto
           TipoCols(Nro_Col) = "CO"
           TipoColumna(columnaActual) = 0
           EsMonto(Nro_Col) = False
       ElseIf CLng(objRs!rliqtipo) = 6 Then 'Cantidad Acumulador
           TipoCols(Nro_Col) = "AC"
           TipoColumna(columnaActual) = 1
           EsMonto(Nro_Col) = False
       ElseIf CLng(objRs!rliqtipo) = 7 Then 'Novedades
           TipoCols(Nro_Col) = "NOV"
           TipoColumna(columnaActual) = 0
           EsMonto(Nro_Col) = True
       ElseIf CLng(objRs!rliqtipo) = 8 Then 'Novedades Ajuste
           TipoCols(Nro_Col) = "NAJ"
           TipoColumna(columnaActual) = 0
           EsMonto(Nro_Col) = True
       ElseIf CLng(objRs!rliqtipo) = 9 Then 'Licencias
           TipoCols(Nro_Col) = "LIC"
           TipoColumna(columnaActual) = 0
           EsMonto(Nro_Col) = True
       ElseIf CLng(objRs!rliqtipo) = 10 Then 'Préstamos
           TipoCols(Nro_Col) = "PRE"
           TipoColumna(columnaActual) = 0
           EsMonto(Nro_Col) = True
       ElseIf CLng(objRs!rliqtipo) = 11 Then 'Embargos
           TipoCols(Nro_Col) = "EMB"
           TipoColumna(columnaActual) = 0
           EsMonto(Nro_Col) = True
       ElseIf CLng(objRs!rliqtipo) = 12 Then 'Vales
           TipoCols(Nro_Col) = "VAL"
           TipoColumna(columnaActual) = 0
           EsMonto(Nro_Col) = True
       ElseIf CLng(objRs!rliqtipo) = 13 Then 'Dirección
           TipoCols(Nro_Col) = "DIR"
           TipoColumna(columnaActual) = 2
       ElseIf CLng(objRs!rliqtipo) = 14 Then 'Tipo Estrucutra Cód. Externo
           TipoCols(Nro_Col) = "TCE"
           TipoColumna(columnaActual) = 2
       ElseIf CLng(objRs!rliqtipo) = 15 Then 'Cuenta Bancaria
           TipoCols(Nro_Col) = "CTA"
           TipoColumna(columnaActual) = 2
       ElseIf CLng(objRs!rliqtipo) = 16 Then 'Datos: Causa Baja, Estado, Estado Civil, email
           TipoCols(Nro_Col) = "DAT"         '       nacionalidad, reporta a, sexo
           TipoColumna(columnaActual) = 2
       ElseIf CLng(objRs!rliqtipo) = 17 Then 'Edad
           TipoCols(Nro_Col) = "EDA"
           TipoColumna(columnaActual) = 2
       ElseIf CLng(objRs!rliqtipo) = 18 Then 'Teléfono
           TipoCols(Nro_Col) = "TEL"
           TipoColumna(columnaActual) = 2
       ElseIf CLng(objRs!rliqtipo) = 19 Then 'Antiguedad
           TipoCols(Nro_Col) = "ANT"
           TipoColumna(columnaActual) = 2
       ElseIf CLng(objRs!rliqtipo) = 20 Then 'Analisis Detallado
           TipoCols(Nro_Col) = "ADE"
           TipoColumna(columnaActual) = 0
           EsMonto(Nro_Col) = True
       ElseIf CLng(objRs!rliqtipo) = 21 Then 'Proceso
           TipoCols(Nro_Col) = "PRO"
           TipoColumna(columnaActual) = 2
           'EsMonto(Nro_Col) = True
        ElseIf CLng(objRs!rliqtipo) = 22 Then 'Periodo
           TipoCols(Nro_Col) = "PER"
           TipoColumna(columnaActual) = 2
           'EsMonto(Nro_Col) = True
        ElseIf CLng(objRs!rliqtipo) = 23 Then 'CANTIDAD TIPO CONCEPTO
           TipoCols(Nro_Col) = "CTC"
           TipoColumna(columnaActual) = 2
           EsMonto(Nro_Col) = False
        ElseIf CLng(objRs!rliqtipo) = 24 Then 'MONTO TIPO CONCEPTO MTC
           TipoCols(Nro_Col) = "CTC"
           TipoColumna(columnaActual) = 2
           EsMonto(Nro_Col) = True
        ElseIf CLng(objRs!rliqtipo) = 25 Then 'CANTIDAD TIPO ACUMULADOR CTA
           TipoCols(Nro_Col) = "CTA"
           TipoColumna(columnaActual) = 2
           EsMonto(Nro_Col) = False
        ElseIf CLng(objRs!rliqtipo) = 26 Then 'MONTO TIPO ACUMULADOR CTA
           TipoCols(Nro_Col) = "CTA"
           TipoColumna(columnaActual) = 2
           EsMonto(Nro_Col) = True
        ElseIf CLng(objRs!rliqtipo) = 27 Then 'CANT CONCEPTOS ACUMULADOR
           TipoCols(Nro_Col) = "CCA"
           TipoColumna(columnaActual) = 2
           EsMonto(Nro_Col) = False
        ElseIf CLng(objRs!rliqtipo) = 28 Then 'CANT CONCEPTOS ACUMULADOR
           TipoCols(Nro_Col) = "CCA"
           TipoColumna(columnaActual) = 2
           EsMonto(Nro_Col) = True
      End If
       
       TitCols(columnaActual) = objRs!rliqetiq
       NroCols(Nro_Col) = CLng(objRs!rliqnrocol)
       CantColumnas = CLng(objRs!rliqcantcol)
       TituloRep = objRs!rliqdesc
    
       objRs.MoveNext
    Loop
    
    objRs.Close

End Sub

Sub GenerarEncabezadoReporte()

Dim teNomb1
Dim teNomb2
Dim teNomb3
Dim I
Dim rsConsult As New ADODB.Recordset

On Error GoTo MError

teNomb1 = ""
teNomb2 = ""
teNomb3 = ""

If tenro1 <> 0 Then
    StrSql = " SELECT * "
    StrSql = StrSql & " FROM tipoestructura "
    StrSql = StrSql & "  WHERE tipoestructura.tenro = " & tenro1
           
    OpenRecordset StrSql, rsConsult
    
    If Not rsConsult.EOF Then
       teNomb1 = rsConsult!tedabr
    Else
       teNomb1 = ""
    End If
End If

If tenro2 <> 0 Then
    StrSql = " SELECT * "
    StrSql = StrSql & " FROM tipoestructura "
    StrSql = StrSql & "  WHERE tipoestructura.tenro = " & tenro2
           
    OpenRecordset StrSql, rsConsult
    
    If Not rsConsult.EOF Then
       teNomb2 = rsConsult!tedabr
    Else
       teNomb2 = ""
    End If
End If

If tenro3 <> 0 Then
    StrSql = " SELECT * "
    StrSql = StrSql & " FROM tipoestructura "
    StrSql = StrSql & "  WHERE tipoestructura.tenro = " & tenro3
           
    OpenRecordset StrSql, rsConsult
    
    If Not rsConsult.EOF Then
       teNomb3 = rsConsult!tedabr
    Else
       teNomb3 = ""
    End If
End If

StrSql = " INSERT INTO rep_estadisticaEmpleados ( " & _
         " bpronro , Formato, rliqdesc, pliqDesde, pliqHasta, " & _
         " tedabr1 , tedabr2, tedabr3, " & _
         " coletiq1 , coletiq2, coletiq3, coletiq4, coletiq5, " & _
         " coletiq6 , coletiq7, coletiq8, coletiq9, coletiq10, " & _
         " coletiq11 , coletiq12, coletiq13, coletiq14, coletiq15, " & _
         " coletiq16 , coletiq17, coletiq18, coletiq19, coletiq20, " & _
         " coletiq21 , coletiq22, coletiq23, coletiq24, coletiq25, " & _
         " coletiq26 , coletiq27, coletiq28, coletiq29, coletiq30, " & _
         " coletiq31 , coletiq32, coletiq33, coletiq34, coletiq35, " & _
         " coletiq36 , coletiq37, coletiq38, coletiq39, coletiq40, " & _
         " coletiq41 , coletiq42, coletiq43, coletiq44, coletiq45, " & _
         " coletiq46 , coletiq47, coletiq48, coletiq49, coletiq50, " & _
         " coltipo1  , coltipo2, coltipo3, coltipo4, coltipo5, " & _
         " coltipo6  , coltipo7, coltipo8, coltipo9, coltipo10, " & _
         " coltipo11 , coltipo12, coltipo13, coltipo14, coltipo15, " & _
         " coltipo16 , coltipo17, coltipo18, coltipo19, coltipo20, " & _
         " coltipo21 , coltipo22, coltipo23, coltipo24, coltipo25, " & _
         " coltipo26 , coltipo27, coltipo28, coltipo29, coltipo30, " & _
         " coltipo31 , coltipo32, coltipo33, coltipo34, coltipo35, " & _
         " coltipo36 , coltipo37, coltipo38, coltipo39, coltipo40, " & _
         " coltipo41 , coltipo42, coltipo43, coltipo44, coltipo45, " & _
         " coltipo46 , coltipo47, coltipo48, coltipo49, coltipo50, " & _
         " cantcols ) VALUES ( "

StrSql = StrSql & NroProceso & ","
StrSql = StrSql & Formato & ","
StrSql = StrSql & "'" & TituloRep & "',"
StrSql = StrSql & "'" & descDesde & "',"
StrSql = StrSql & "'" & descHasta & "',"
StrSql = StrSql & controlNull(teNomb1) & ","
StrSql = StrSql & controlNull(teNomb2) & ","
StrSql = StrSql & controlNull(teNomb3) & ","

For I = 1 To 50
    StrSql = StrSql & "'" & TitCols(I) & "',"
Next

For I = 1 To 50
  StrSql = StrSql & TipoColumna(I) & ","
Next

StrSql = StrSql & CantColumnas & ")"

objConn.Execute StrSql, , adExecuteNoRecords

Exit Sub

MError:
    Flog.writeline "Error al cargar los datos del reporte. Error: " & Err.Description
    HuboErrores = True
    EmpErrores = True
    Exit Sub

End Sub


Sub borrarValores()
  
  Dim j
  
  For j = j To CantColumnas
    ValCols(j) = 0
  Next
End Sub 'borrarValores
Sub borrarChar()
  
  Dim j
  
  For j = 1 To CantColumnas
    CharCols(j) = ""
  Next
End Sub 'borrarChar

Sub agregarValor(tipo, codigo, Monto, Cantidad)
  Dim j
  
  For j = 1 To Nro_Col
    If TipoCols(j) = tipo Then 'And CodCols(j) = codigo Then
       'If EsMonto(j) Then
          ValCols(CInt(NroCols(j))) = CDbl(ValCols(CInt(NroCols(j)))) + CDbl(Monto)
       'Else
          ValCols(CInt(NroCols(j))) = CDbl(ValCols(CInt(NroCols(j)))) + CDbl(Cantidad)
       'End If
    End If
  Next
End Sub 'agregarValor(tipo, codigo, valor)
Sub agregarValorEdad(Columna, Valor)
  Dim j
  
     CharCols(Columna) = Valor
  
End Sub 'agregarValorEdad(columna,valor)

Sub agregarValorNov(tipo, codigo, Monto)
  Dim j
  
  For j = 1 To Nro_Col
    If TipoCols(j) = tipo And CodNovCols(j) = codigo Then
       If EsMonto(j) Then
          ValCols(CInt(NroCols(j))) = CDbl(ValCols(CInt(NroCols(j)))) + CDbl(Monto)
       End If
    End If
  Next
End Sub 'agregarValor(tipo, codigo, valor)
Sub agregarValorDir(Columna, tipo, Datos() As String)
  Dim j
  
  If Datos(3) <> "" Then
     CharCols(Columna) = Datos(1) & " " & Datos(2) & " Piso: " & Datos(3)
  Else
     CharCols(Columna) = Datos(1) & " " & Datos(2)
  End If
            
  If Datos(4) <> "" Then
     CharCols(Columna) = CharCols(Columna) & " Dpto: " & Datos(4)
  End If
         
  If tipo = 2 Then
     CharCols(Columna) = CharCols(Columna) & " - " & Datos(5)
  ElseIf tipo = 3 Then
         CharCols(Columna) = CharCols(Columna) & " - " & Datos(5)
         CharCols(Columna) = CharCols(Columna) & " - " & Datos(6)
         
  ElseIf tipo = 4 Then
         CharCols(Columna) = CharCols(Columna) & " - " & Datos(5)
         CharCols(Columna) = CharCols(Columna) & " - " & Datos(6)
         CharCols(Columna) = CharCols(Columna) & " - " & Datos(7)
  End If
End Sub 'agregarValor(tipo, codigo, valor)
Sub agregarValorDirOld(Columna, tipo, Valor1, Valor2, Valor3, Valor4, Valor5, Valor6, Valor7)
  Dim j
  
  If Valor3 <> 0 Then
     CharCols(Columna) = Valor1 & " " & Valor2 & " Piso: " & Valor3
  Else
     CharCols(Columna) = Valor1 & " " & Valor2
  End If
            
  If Valor4 <> 0 Then
     CharCols(Columna) = CharCols(Columna) & " Dpto: " & Valor4
  End If
         
  If tipo = 2 Then
     CharCols(Columna) = CharCols(Columna) & " - " & Valor5
  ElseIf tipo = 3 Then
         CharCols(Columna) = CharCols(Columna) & " - " & Valor5
         CharCols(Columna) = CharCols(Columna) & " - " & Valor6
         
  ElseIf tipo = 4 Then
         CharCols(Columna) = CharCols(Columna) & " - " & Valor5
         CharCols(Columna) = CharCols(Columna) & " - " & Valor6
         CharCols(Columna) = CharCols(Columna) & " - " & Valor7
  End If
End Sub 'agregarValor(tipo, codigo, valor)

Sub agregarChar(Columna, Valor)
  
    CharCols(CInt(NroCols(Columna))) = Valor
  
End Sub 'agregarChar(tipo, codigo, valor)



'--------------------------------------------------------------------
' Se encarga de generar los datos para el empleado para un periodo
'--------------------------------------------------------------------
Sub GenerarDatosEmpleadoPeriodo(ListaProcesos, pliqNro, Ternro, orden)

Dim StrSql As String
Dim rsConsult As New ADODB.Recordset

'Variables donde se guardan los datos del INSERT final

Dim Legajo As Long
Dim apellido As String
Dim apellido2 As String
Dim nombre As String
Dim nombre2 As String

Dim estrnomb1
Dim estrnomb2
Dim estrnomb3
Dim pliqDesc
Dim pliqFecha
Dim I
Dim proNro

On Error GoTo MError

estrnomb1 = ""
estrnomb2 = ""
estrnomb3 = ""
proNro = 0

'------------------------------------------------------------------
'Controlo si el empleado tiene algun proceso en el periodo
'------------------------------------------------------------------
StrSql = " SELECT * "
StrSql = StrSql & " FROM proceso "
StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro AND proceso.pliqnro IN (" & pliqNro & ")"
StrSql = StrSql & " WHERE empleado = " & Ternro
StrSql = StrSql & "   AND proceso.pliqnro IN (" & pliqNro & ")"
StrSql = StrSql & "   AND proceso.pronro IN (" & ListaProcesos & ")"

OpenRecordset StrSql, rsConsult

If rsConsult.EOF Then
   'Si el empleado no tiene procesos en el periodo paso al siguiente
   rsConsult.Close
   
   Exit Sub
End If

rsConsult.Close

'------------------------------------------------------------------
'Busco los datos del empleado
'------------------------------------------------------------------
StrSql = " SELECT empleg,terape,terape2,ternom,ternom2,empfaltagr,empremu "
StrSql = StrSql & " FROM empleado "
StrSql = StrSql & " WHERE ternro= " & Ternro

Flog.writeline "Buscando datos del empleado"
       
OpenRecordset StrSql, rsConsult
nomape = ""
If Not rsConsult.EOF Then
   nombre = rsConsult!ternom
   nomape = nombre
   If IsNull(rsConsult!ternom2) Then
      nombre2 = ""
   Else
      nombre2 = rsConsult!ternom2
      nomape = nomape & " " & nombre2
   End If
   apellido = rsConsult!terape
   nomape = nomape & " " & apellido
   If IsNull(rsConsult!terape2) Then
      apellido2 = ""
   Else
      apellido2 = rsConsult!terape2
      nomape = nomape & " " & apellido2
   End If
   Legajo = rsConsult!empleg
Else
   Flog.writeline "Error al obtener los datos del empleado"
'   GoTo MError
End If

rsConsult.Close

'------------------------------------------------------------------
'Busco los datos del tipos de estructura 1
'------------------------------------------------------------------

'---LOG---
Flog.writeline "Buscando datos estructura 1"

If tenro1 <> 0 Then
    
    StrSql = " SELECT estrdabr "
    StrSql = StrSql & " FROM his_estructura "
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
    StrSql = StrSql & "    AND his_estructura.ternro = " & Ternro & " AND his_estructura.tenro = " & tenro1
    StrSql = StrSql & "    AND (htetdesde<=" & ConvFecha(fecEstr) & " AND (htethasta is null or htethasta>=" & ConvFecha(fecEstr) & "))"
    
    If estrnro1 <> 0 Then
        StrSql = StrSql & " AND estructura.estrnro = " & estrnro1
    End If
               
    OpenRecordset StrSql, rsConsult
    
    If Not rsConsult.EOF Then
       estrnomb1 = rsConsult!estrdabr
    End If
End If

'------------------------------------------------------------------
'Busco los datos del tipos de estructura 2
'------------------------------------------------------------------

'---LOG---
Flog.writeline "Buscando datos estructura 2"

If tenro2 <> 0 Then
    
    StrSql = " SELECT estrdabr "
    StrSql = StrSql & " FROM his_estructura "
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
    StrSql = StrSql & "    AND his_estructura.ternro = " & Ternro & " AND his_estructura.tenro = " & tenro2
    StrSql = StrSql & "    AND (htetdesde<=" & ConvFecha(fecEstr) & " AND (htethasta is null or htethasta>=" & ConvFecha(fecEstr) & "))"
    
    If estrnro2 <> 0 Then
        StrSql = StrSql & " AND estructura.estrnro = " & estrnro2
    End If
           
    OpenRecordset StrSql, rsConsult
    
    If Not rsConsult.EOF Then
       estrnomb2 = rsConsult!estrdabr
    End If
End If

'------------------------------------------------------------------
'Busco los datos del tipos de estructura 3
'------------------------------------------------------------------

'---LOG---
Flog.writeline "Buscando datos estructura 3"

If tenro3 <> 0 Then
    
    StrSql = " SELECT estrdabr "
    StrSql = StrSql & " FROM his_estructura "
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
    StrSql = StrSql & "    AND his_estructura.ternro = " & Ternro & " AND his_estructura.tenro = " & tenro3
    StrSql = StrSql & "    AND (htetdesde<=" & ConvFecha(fecEstr) & " AND (htethasta is null or htethasta>=" & ConvFecha(fecEstr) & "))"
    
    If estrnro3 <> 0 Then
        StrSql = StrSql & " AND estructura.estrnro = " & estrnro3
    End If
    
    OpenRecordset StrSql, rsConsult
    
    If Not rsConsult.EOF Then
       estrnomb3 = rsConsult!estrdabr
    End If
End If

'------------------------------------------------------------------
'Busco los datos del periodo
'------------------------------------------------------------------
'StrSql = " SELECT * FROM periodo WHERE pliqnro = " & pliqNro
StrSql = " SELECT MIN(pliqanio) anioDesde, max(pliqanio) anioHasta FROM periodo WHERE pliqnro in(" & pliqNro & ")"
OpenRecordset StrSql, rsConsult

'pliqDesc = ""
If Not rsConsult.EOF Then
   Desde = rsConsult!anioDesde
   Hasta = rsConsult!anioHasta
End If

rsConsult.Close

Dim cont

For cont = 1 To Nro_Col

    Select Case TipoCols(cont)
        Case "CO":
            Call buscarConc(Desde, Hasta, pliqNro, ListaProcesos, Ternro, cont, Legajo)
        
        Case "AC":
            Call buscarAcu(Desde, Hasta, pliqNro, ListaProcesos, Ternro, cont, Legajo)
        
        Case "CTC"
            Call conceptosImpri(Desde, Hasta, pliqNro, ListaProcesos, Ternro, Legajo)
        
        Case "MTC"
            Call conceptosNoImpri(Desde, Hasta, pliqNro, ListaProcesos, Ternro, Legajo)
        
        Case "CTA"
            Call acumuladoresImpri(Desde, Hasta, pliqNro, ListaProcesos, Ternro, Legajo)
        
        Case "MTA"
            Call acumuladoresNoImpri(Desde, Hasta, pliqNro, ListaProcesos, Ternro, Legajo)
        
        Case "CCA"
            Call CantConceptosAcumulador(Desde, Hasta, pliqNro, ListaProcesos, Ternro, cont, Legajo)
    End Select
    
Next
Exit Sub


'------------------------------------------------------------------
'Busco los valores de la cant/monto de tipo de conceptos/acumuladores imprimibles
'sebastian stremel 20/09/2012
'------------------------------------------------------------------

StrSql = " SELECT 'CTC', detliq.concnro, sum(detliq.dlicant) AS cant, sum(detliq.dlimonto) AS monto, periodo.pliqnro,periodo.pliqmes,periodo.pliqanio,concepto.concabr  "
StrSql = StrSql & " FROM cabliq "
StrSql = StrSql & " INNER JOIN proceso  ON cabliq.pronro = proceso.pronro  AND proceso.pliqnro IN (" & pliqNro & ") AND cabliq.pronro IN (" & ListaProcesos & ") "
StrSql = StrSql & " INNER JOIN detliq   ON cabliq.cliqnro = detliq.cliqnro  AND cabliq.empleado = " & Ternro & " AND cabliq.pronro IN (" & ListaProcesos & ") "
StrSql = StrSql & " INNER JOIN concepto on concepto.concnro=detliq.concnro "
StrSql = StrSql & " INNER JOIN periodo on periodo.pliqnro = proceso.pliqnro "
'StrSql = StrSql & " AND ( 1=0 "
  
For I = 1 To Nro_Col
   If TipoCols(I) = "CTC" Then
        If CodCols(I) <> 1 Then
            StrSql = StrSql & " AND concepto.concimp = " & CodCols(I)
        End If
   End If
Next
    
'StrSql = StrSql & " ) "
StrSql = StrSql & " GROUP BY detliq.concnro, periodo.pliqnro,periodo.pliqmes ,periodo.pliqanio,concepto.concabr   "
StrSql = StrSql & " ORDER BY detliq.concnro, periodo.pliqnro,periodo.pliqmes,periodo.pliqanio,concepto.concabr  "
    
'StrSql = StrSql & " UNION "
    
'StrSql = StrSql & " SELECT 'CTA', acu_liq.acunro, sum(acu_liq.alcant) AS cant, sum(acu_liq.almonto) AS monto "
'StrSql = StrSql & " FROM cabliq "
'StrSql = StrSql & " INNER JOIN proceso  ON cabliq.pronro = proceso.pronro  AND proceso.pliqnro IN (" & pliqNro & ") AND cabliq.pronro IN (" & ListaProcesos & ") "
'StrSql = StrSql & " INNER JOIN acu_liq  ON cabliq.cliqnro = acu_liq.cliqnro AND cabliq.empleado = " & Ternro & " AND cabliq.pronro = proceso.pronro "
'StrSql = StrSql & " INNER JOIN acumulador on acumulador.acunro=acu_liq.acunro "
'StrSql = StrSql & " AND ( 1=0 "
    
'For I = 1 To Nro_Col
'    If TipoCols(I) = "CTA" Then
'        If CodCols(I) <> 1 Then
'            StrSql = StrSql & " AND acumulador.acuimpri= " & CodCols(I)
'        End If
'    End If
'Next
    
'StrSql = StrSql & " ) "
'StrSql = StrSql & " GROUP BY acu_liq.acunro "
Flog.writeline "borrarChar"
Call borrarValores

OpenRecordset StrSql, rsConsult
concAnt = ""
Do Until rsConsult.EOF
    'Call agregarValor(rsConsult(0), rsConsult(1), rsConsult(3), rsConsult(2))
    If rsConsult!ConcNro <> concAnt Then
        StrSql = " INSERT INTO repEstadisticaDet "
        StrSql = StrSql & " (bpronro,ternro,empleg,nombre,apellido,detnro,detalle, "
        StrSql = StrSql & " periodo" & rsConsult!pliqmes & ")"
        StrSql = StrSql & " VALUES "
        StrSql = StrSql & "(" & NroProceso & "," & Ternro & "," & Ternro & ",'',''," & rsConsult!ConcNro & "," & rsConsult!concabr & ", "
        StrSql = StrSql & Replace(rsConsult!Monto, ",", ".") & " )"
        concAnt = rsConsult!ConcNro
    Else
    End If
   
   rsConsult.MoveNext
Loop
objConn.Execute StrSql, , adExecuteNoRecords
rsConsult.Close

'------------------------------------------------------------------




'------------------------------------------------------------------
'Busco los valores de los conceptos y acumuladores
'------------------------------------------------------------------

StrSql = " SELECT 'CO', detliq.concnro, sum(detliq.dlicant) AS cant, sum(detliq.dlimonto) AS monto  "
StrSql = StrSql & " FROM cabliq "
StrSql = StrSql & " INNER JOIN proceso  ON cabliq.pronro = proceso.pronro  AND proceso.pliqnro = " & pliqNro & " AND cabliq.pronro IN (" & ListaProcesos & ") "
StrSql = StrSql & " INNER JOIN detliq   ON cabliq.cliqnro = detliq.cliqnro  AND cabliq.empleado = " & Ternro & " AND cabliq.pronro = proceso.pronro "
StrSql = StrSql & " AND ( 1=0 "
  
For I = 1 To Nro_Col
   If TipoCols(I) = "CO" Then
      StrSql = StrSql & " OR detliq.concnro = " & CodCols(I)
   End If
Next
    
StrSql = StrSql & " ) "
StrSql = StrSql & " GROUP BY detliq.concnro "
    
StrSql = StrSql & " UNION "
    
StrSql = StrSql & " SELECT 'AC', acu_liq.acunro, sum(acu_liq.alcant) AS cant, sum(acu_liq.almonto) AS monto "
StrSql = StrSql & " FROM cabliq "
StrSql = StrSql & " INNER JOIN proceso  ON cabliq.pronro = proceso.pronro  AND proceso.pliqnro = " & pliqNro & " AND cabliq.pronro IN (" & ListaProcesos & ") "
StrSql = StrSql & " INNER JOIN acu_liq  ON cabliq.cliqnro = acu_liq.cliqnro AND cabliq.empleado = " & Ternro & " AND cabliq.pronro = proceso.pronro "
StrSql = StrSql & " AND ( 1=0 "
    
For I = 1 To Nro_Col
    If TipoCols(I) = "AC" Then
       StrSql = StrSql & " OR acu_liq.acunro = " & CodCols(I)
    End If
Next
    
StrSql = StrSql & " ) "
StrSql = StrSql & " GROUP BY acu_liq.acunro "
Flog.writeline "borrarChar"
Call borrarValores

OpenRecordset StrSql, rsConsult

Do Until rsConsult.EOF
   Call agregarValor(rsConsult(0), rsConsult(1), rsConsult(3), rsConsult(2))
   rsConsult.MoveNext
Loop
rsConsult.Close

'------------------------------------------------------------------
'Busco los valores de los tipo estructura - Descripción
'------------------------------------------------------------------
Flog.writeline "borrarChar"
Call borrarChar
Flog.writeline "Busco los valores de los tipo estructura - Descripción"
For I = 1 To Nro_Col

   If TipoCols(I) = "TE" Then
   
      StrSql = " SELECT 'TE', estrdabr "
      StrSql = StrSql & " FROM his_estructura "
      StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
      StrSql = StrSql & "    AND his_estructura.ternro = " & Ternro
      StrSql = StrSql & "    AND (htetdesde<=" & ConvFecha(fecEstr) & " AND (htethasta is null or htethasta>=" & ConvFecha(fecEstr) & "))"
      StrSql = StrSql & " AND his_estructura.tenro = " & CodCols(I)
            
      OpenRecordset StrSql, rsConsult

      Do Until rsConsult.EOF
         Call agregarChar(I, rsConsult(1))
         rsConsult.MoveNext
      Loop
      rsConsult.Close
   End If
Next

'------------------------------------------------------------------
'Busco los valores de los tipos de fechas
'------------------------------------------------------------------
Flog.writeline "Busco los valores de los tipos de fechas"
For I = 1 To Nro_Col
    If TipoCols(I) = "TF" Then
        Select Case CodCols(I)
        Case 1: 'Fecha de nacimiento
            StrSql = "SELECT 'TF', terfecnac FROM tercero "
            StrSql = StrSql & " WHERE ternro = " & Ternro
        Case 2: 'Fecha de alta reconocida
            StrSql = "SELECT 'TF', altfec FROM fases "
            StrSql = StrSql & " WHERE empleado = " & Ternro
            StrSql = StrSql & " AND fasrecofec = -1 "
            StrSql = StrSql & " ORDER BY altfec "
        Case 3: 'Fecha fase mas antigua
            StrSql = "SELECT 'TF', altfec FROM fases "
            StrSql = StrSql & " WHERE empleado = " & Ternro
            StrSql = StrSql & " ORDER BY altfec "
        Case 4: 'fecha fase mas nueva
            StrSql = "SELECT 'TF', altfec FROM fases "
            StrSql = StrSql & " WHERE empleado = " & Ternro
            StrSql = StrSql & " ORDER BY altfec desc "
        Case 5: 'fecha baja
            StrSql = "SELECT 'TF', bajfec FROM fases "
            StrSql = StrSql & " WHERE empleado = " & Ternro
            StrSql = StrSql & " ORDER BY bajfec desc "
        End Select
        OpenRecordset StrSql, rsConsult
        
      If Not rsConsult.EOF Then
         Call agregarChar(I, rsConsult(1))
      End If
      rsConsult.Close
   End If
Next

'------------------------------------------------------------------
'Busco los valores de los tipo de documentos
'------------------------------------------------------------------
Flog.writeline "Busco los valores de los tipo de documentos"
For I = 1 To Nro_Col

   If TipoCols(I) = "TD" Then
   
      If CodCols(I) = 1 Then
         StrSql = " SELECT 'TD', nrodoc "
         StrSql = StrSql & " FROM ter_doc "
         StrSql = StrSql & " WHERE ter_doc.ternro = " & Ternro
         StrSql = StrSql & " AND ter_doc.tidnro <= 4 "
      Else
         StrSql = " SELECT 'TD', nrodoc "
         StrSql = StrSql & " FROM ter_doc "
         StrSql = StrSql & " WHERE ter_doc.ternro = " & Ternro
         StrSql = StrSql & " AND ter_doc.tidnro = " & CodCols(I)
      End If
      
      OpenRecordset StrSql, rsConsult

      Do Until rsConsult.EOF
         Call agregarChar(I, rsConsult(1))
         rsConsult.MoveNext
      Loop
      rsConsult.Close
   End If
Next

'------------------------------------------------------------------
'Obtengo la fecha desde y hasta del periodo
'------------------------------------------------------------------
Flog.writeline "Obtengo la fecha desde y hasta del periodo"
Dim FechaDesde As Date
Dim FechaHasta As Date
Dim AnioPeriodo As Integer
Dim MesPeriodo As Integer
Dim dias As Integer
Dim Aux_Fecha_Desde As Date
Dim Aux_Fecha_Hasta As Date


StrSql = " SELECT pliqdesde, pliqhasta, pliqmes, pliqanio "
StrSql = StrSql & " FROM periodo "
StrSql = StrSql & " WHERE periodo.pliqnro = " & pliqNro
OpenRecordset StrSql, rsConsult
If Not rsConsult.EOF Then
   FechaDesde = rsConsult!pliqdesde
   FechaHasta = rsConsult!pliqhasta
   MesPeriodo = rsConsult!pliqmes
   AnioPeriodo = rsConsult!pliqanio
End If
rsConsult.Close

'------------------------------------------------------------------
'Busco los valores de las novedades individuales
'------------------------------------------------------------------
Flog.writeline "Busco los valores de las novedades individuales"
StrSql = " SELECT 'NOV', SUM(nevalor), novemp.concnro, novemp.tpanro "
StrSql = StrSql & " FROM novemp "
StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = novemp.concnro "
StrSql = StrSql & " WHERE novemp.empleado = " & Ternro
StrSql = StrSql & " AND ((novemp.nedesde >= " & ConvFecha(FechaDesde)
StrSql = StrSql & " AND  (novemp.nehasta <= " & ConvFecha(FechaHasta) & " OR novemp.nehasta IS NULL))"
StrSql = StrSql & " OR   novemp.nevigencia = 0 )"
StrSql = StrSql & " AND ( 1=0 "
For I = 1 To Nro_Col
   CodNov = Split(CodNovCols(I), "@")
   If TipoCols(I) = "NOV" Then
      StrSql = StrSql & " OR (concepto.concnro = " & CodNov(0) & " AND novemp.tpanro = " & CodNov(1) & ")"
   End If
Next
StrSql = StrSql & " ) "
StrSql = StrSql & " GROUP BY novemp.concnro, novemp.tpanro "
OpenRecordset StrSql, rsConsult
Do Until rsConsult.EOF
   Call agregarValorNov(rsConsult(0), rsConsult(2) & "@" & rsConsult(3), rsConsult(1))
   rsConsult.MoveNext
Loop
rsConsult.Close

'------------------------------------------------------------------
'Busco los valores de las novedades ajuste
'------------------------------------------------------------------
Flog.writeline "Busco los valores de las novedades ajuste"
StrSql = " SELECT 'NAJ', SUM(navalor), novaju.concnro "
StrSql = StrSql & " FROM novaju "
StrSql = StrSql & " WHERE novaju.empleado = " & Ternro
StrSql = StrSql & " AND ((novaju.nadesde >= " & ConvFecha(FechaDesde)
StrSql = StrSql & " AND  (novaju.nahasta <= " & ConvFecha(FechaHasta) & " OR novaju.nahasta IS NULL))"
StrSql = StrSql & " OR   novaju.navigencia = 0 )"
StrSql = StrSql & " AND ( 1=0 "
For I = 1 To Nro_Col
   If TipoCols(I) = "NAJ" Then
      StrSql = StrSql & " OR (novaju.concnro = " & CodCols(I) & ")"
   End If
Next
StrSql = StrSql & " ) "
StrSql = StrSql & " GROUP BY novaju.concnro "
OpenRecordset StrSql, rsConsult
Do Until rsConsult.EOF
   Call agregarValor(rsConsult(0), rsConsult(2), rsConsult(1), 0)
   rsConsult.MoveNext
Loop
rsConsult.Close

'------------------------------------------------------------------
'Busco los valores de las licencias
'------------------------------------------------------------------
Flog.writeline "Busco los valores de las Licencias"
'StrSql = " SELECT 'LIC', SUM(elcantdias), emp_lic.tdnro "
'StrSql = StrSql & " FROM emp_lic "
'StrSql = StrSql & " WHERE emp_lic.empleado = " & ternro
'StrSql = StrSql & " AND (emp_lic.elfechadesde >= " & ConvFecha(FechaDesde)
'StrSql = StrSql & " AND  emp_lic.elfechahasta <= " & ConvFecha(FechaHasta) & ")"
'StrSql = StrSql & " AND ( 1=0 "
'For i = 1 To Nro_Col
'   If TipoCols(i) = "LIC" Then
'      StrSql = StrSql & " OR (emp_lic.tdnro = " & CodCols(i) & ")"
'   End If
'Next
'StrSql = StrSql & " ) "
'StrSql = StrSql & " GROUP BY emp_lic.tdnro "
'OpenRecordset StrSql, rsConsult
'Do Until rsConsult.EOF
'   Call agregarValor(rsConsult(0), rsConsult(2), rsConsult(1), 0)
'   rsConsult.MoveNext
'Loop
'rsConsult.Close


'Martin Ferraro - 13/03/2006 - nueva version
StrSql = "SELECT 'LIC', elcantdias, emp_lic.tdnro, elfechadesde, elfechahasta "
StrSql = StrSql & " FROM emp_lic WHERE (empleado = " & Ternro
StrSql = StrSql & " AND elfechadesde <=" & ConvFecha(FechaHasta)
StrSql = StrSql & " AND elfechahasta >= " & ConvFecha(FechaDesde)
StrSql = StrSql & " AND ( 1=0 "
For I = 1 To Nro_Col
   If TipoCols(I) = "LIC" Then
      StrSql = StrSql & " OR (emp_lic.tdnro = " & CodCols(I) & ") "
   End If
Next
StrSql = StrSql & " ) )"
'StrSql = StrSql & " GROUP BY emp_lic.tdnro, elcantdias "
OpenRecordset StrSql, rsConsult
dias = 0
Do While Not rsConsult.EOF
    Aux_Fecha_Desde = rsConsult!elfechadesde
    Aux_Fecha_Hasta = rsConsult!elfechahasta

    If Aux_Fecha_Hasta > FechaHasta Then
        Aux_Fecha_Hasta = FechaHasta
    End If
    dias = CantidadDeDias(FechaDesde, FechaHasta, Aux_Fecha_Desde, Aux_Fecha_Hasta)
    
    Call agregarValor(rsConsult(0), rsConsult(2), dias, 0)
    rsConsult.MoveNext
Loop
rsConsult.Close


'------------------------------------------------------------------
'Busco los valores de los préstamos
'------------------------------------------------------------------
Flog.writeline "Busco los valores de los prestamos"
StrSql = " SELECT 'PRE', SUM(cuototal), prestamo.lnprenro "
StrSql = StrSql & " FROM pre_cuota "
StrSql = StrSql & " INNER JOIN prestamo ON prestamo.prenro = pre_cuota.prenro "
StrSql = StrSql & " WHERE prestamo.ternro = " & Ternro
StrSql = StrSql & " AND pre_cuota.cuomes = " & MesPeriodo
StrSql = StrSql & " AND  pre_cuota.cuoano = " & AnioPeriodo
StrSql = StrSql & " AND ( 1=0 "
For I = 1 To Nro_Col
   If TipoCols(I) = "PRE" Then
      StrSql = StrSql & " OR (prestamo.lnprenro = " & CodCols(I) & ")"
   End If
Next
StrSql = StrSql & " ) "
StrSql = StrSql & " GROUP BY prestamo.lnprenro "
OpenRecordset StrSql, rsConsult
Do Until rsConsult.EOF
   Call agregarValor(rsConsult(0), rsConsult(2), rsConsult(1), 0)
   rsConsult.MoveNext
Loop
rsConsult.Close

'------------------------------------------------------------------
'Busco los valores de los embargos
'------------------------------------------------------------------
Flog.writeline "Busco los valores de los embargos"
StrSql = " SELECT 'EMB', SUM(embcimp), embargo.tpenro "
StrSql = StrSql & " FROM embcuota "
StrSql = StrSql & " INNER JOIN embargo ON embargo.embnro = embcuota.embnro "
StrSql = StrSql & " WHERE embargo.ternro = " & Ternro
StrSql = StrSql & " AND embcuota.embcmes = " & MesPeriodo
StrSql = StrSql & " AND  embcuota.embcanio = " & AnioPeriodo
StrSql = StrSql & " AND ( 1=0 "
For I = 1 To Nro_Col
   If TipoCols(I) = "EMB" Then
      StrSql = StrSql & " OR (embargo.tpenro = " & CodCols(I) & ")"
   End If
Next
StrSql = StrSql & " ) "
StrSql = StrSql & " GROUP BY embargo.tpenro "
OpenRecordset StrSql, rsConsult
Do Until rsConsult.EOF
   Call agregarValor(rsConsult(0), rsConsult(2), rsConsult(1), 0)
   rsConsult.MoveNext
Loop
rsConsult.Close

'------------------------------------------------------------------
'Busco los valores de los vales
'------------------------------------------------------------------
Flog.writeline "Busco los valores de los vales"
StrSql = " SELECT 'VAL', SUM(valmonto), vales.tvalenro "
StrSql = StrSql & " FROM vales "
StrSql = StrSql & " WHERE vales.empleado = " & Ternro
StrSql = StrSql & " AND  vales.pliqdto = " & pliqNro
StrSql = StrSql & " AND ( 1=0 "
For I = 1 To Nro_Col
   If TipoCols(I) = "VAL" Then
      StrSql = StrSql & " OR (vales.tvalenro = " & CodCols(I) & ")"
   End If
Next
StrSql = StrSql & " ) "
StrSql = StrSql & " GROUP BY vales.tvalenro "
OpenRecordset StrSql, rsConsult
Do Until rsConsult.EOF
   Call agregarValor(rsConsult(0), rsConsult(2), rsConsult(1), 0)
   rsConsult.MoveNext
Loop
rsConsult.Close

'------------------------------------------------------------------
'Busco los valores de la Dirección
'------------------------------------------------------------------
Flog.writeline "Busco los valores de la Dirección"
'Call borrarChar
Dim TipoDomi
Dim Datos(8) As String
Dim j

For I = 1 To Nro_Col
    If TipoCols(I) = "DIR" Then
        CodNov = Split(CodNovCols(I), "@")
        TipoDomi = CodNov(0)
        
        'Calle, Nro, Piso, Dpto, Localidad, Provincia, País
        StrSql = " SELECT 'DIR', calle, nro, piso, oficdepto, locdesc, provdesc, paisdesc "
        StrSql = StrSql & " FROM cabdom "
        StrSql = StrSql & " INNER JOIN detdom ON detdom.domnro = cabdom.domnro "
        StrSql = StrSql & " INNER JOIN localidad ON localidad.locnro = detdom.locnro "
        StrSql = StrSql & " INNER JOIN provincia ON provincia.provnro = detdom.provnro "
        StrSql = StrSql & " INNER JOIN pais ON pais.paisnro = detdom.paisnro "
        StrSql = StrSql & " WHERE cabdom.ternro = " & Ternro & " AND cabdom.tidonro = " & TipoDomi
            
        OpenRecordset StrSql, rsConsult
        
        If Not rsConsult.EOF Then
           For j = 1 To 7
               If IsNull(rsConsult(j)) Then
                  Datos(j) = ""
               Else
                  Datos(j) = rsConsult(j)
               End If
           Next
           Call agregarValorDir(I, CodNov(1), Datos)
        End If
        rsConsult.Close
    End If
Next

'------------------------------------------------------------------
'Busco los valores de las cuentas bancarias
'------------------------------------------------------------------
Flog.writeline "Busco los valores de las cuentas bancarias"
For I = 1 To Nro_Col

   If TipoCols(I) = "CTA" Then
   
      StrSql = " SELECT 'CTA', ctabnro "
      StrSql = StrSql & " FROM ctabancaria "
      StrSql = StrSql & " WHERE ctabancaria.ternro = " & Ternro
      If CodCols(I) = -1 Then
         StrSql = StrSql & " AND ctabancaria.ctabestado = -1 "
      Else
         StrSql = StrSql & " AND ctabancaria.fpagnro = " & CodCols(I)
      End If
      OpenRecordset StrSql, rsConsult

      Do Until rsConsult.EOF
         Call agregarChar(I, rsConsult(1))
         rsConsult.MoveNext
      Loop
      rsConsult.Close
   End If
Next

'------------------------------------------------------------------
'Busco los Tipo Sigla
'------------------------------------------------------------------
Flog.writeline "Busco los valores de las cuentas bancarias"
For I = 1 To Nro_Col

   If TipoCols(I) = "TIPSIG" Then
         StrSql = " SELECT tipodocu.tidsigla "
         StrSql = StrSql & " From Tercero"
         StrSql = StrSql & " LEFT JOIN ter_doc docu ON (docu.ternro= tercero.ternro and docu.tidnro>0 and docu.tidnro<5) "
         StrSql = StrSql & " LEFT JOIN tipodocu ON tipodocu.tidnro= docu.tidnro "
         StrSql = StrSql & " WHERE tercero.ternro = " & Ternro
        
         'StrSql = "select tipodocu.tidsigla from empleado "
         'StrSql = StrSql & "inner join ter_doc on ter_doc.ternro=empleado.ternro "
         'StrSql = StrSql & "inner join tipodocu on tipodocu.tidnro=ter_doc.tidnro "
         'StrSql = StrSql & "Where Empleado.Ternro = " & Ternro
         OpenRecordset StrSql, rsConsult

        Do Until rsConsult.EOF
             Call agregarChar(I, rsConsult(0))
             rsConsult.MoveNext
        Loop
         rsConsult.Close
   End If
Next

'------------------------------------------------------------------
'Busco los valores de las datos
'------------------------------------------------------------------
Flog.writeline "Busco los valores de las datos"
For I = 1 To Nro_Col

   If TipoCols(I) = "DAT" Then
   
      Select Case CodCols(I)
        Case 1: 'Causa Baja
            StrSql = "SELECT 'DAT', caudes FROM fases "
            StrSql = StrSql & " INNER JOIN causa ON causa.caunro = fases.caunro "
            StrSql = StrSql & " WHERE empleado = " & Ternro
            StrSql = StrSql & " AND fases.bajfec <= " & ConvFecha(FechaHasta)
            StrSql = StrSql & " ORDER BY bajfec DESC "
        Case 2: 'Email Interno
            StrSql = "SELECT 'DAT', empemail FROM empleado "
            StrSql = StrSql & " WHERE ternro = " & Ternro
        Case 3: 'Estado del Empleado
            StrSql = "SELECT 'DAT', empest FROM empleado "
            StrSql = StrSql & " WHERE ternro = " & Ternro
        Case 4: 'Estado Civil
            StrSql = "SELECT 'DAT', estcivdesabr FROM tercero "
            StrSql = StrSql & " INNER JOIN estcivil ON estcivil.estcivnro = tercero.estcivnro "
            StrSql = StrSql & " WHERE ternro = " & Ternro
        Case 5: 'Nacionalidad
            StrSql = "SELECT 'DAT', nacionaldes FROM tercero "
            StrSql = StrSql & " INNER JOIN nacionalidad ON nacionalidad.nacionalnro = tercero.nacionalnro "
            StrSql = StrSql & " WHERE ternro = " & Ternro
        Case 6: 'Reporta A
            StrSql = "SELECT 'DAT', e2.empleg, e2.terape, e2.terape2, e2.ternom, e2.ternom2 "
            StrSql = StrSql & " FROM empleado e1 "
            StrSql = StrSql & " INNER JOIN empleado e2 ON e2.ternro = e1.empreporta  "
            StrSql = StrSql & " WHERE e1.ternro = " & Ternro
        Case 7: 'Sexo
            StrSql = "SELECT 'DAT', tersex FROM tercero "
            StrSql = StrSql & " WHERE ternro = " & Ternro
      End Select
      OpenRecordset StrSql, rsConsult

      Do Until rsConsult.EOF
         If CodCols(I) = 3 Then
            If rsConsult(1) = "-1" Then
               Call agregarChar(I, "Activo")
            Else
               Call agregarChar(I, "Inactivo")
            End If
         ElseIf CodCols(I) = 7 Then
                If rsConsult(1) = "-1" Then
                   Call agregarChar(I, "Masculino")
                Else
                   Call agregarChar(I, "Femenino")
                End If
         ElseIf CodCols(I) = 6 Then
                Call agregarChar(I, rsConsult(1) & " - " & rsConsult(2) & " " & rsConsult(3) & ", " & rsConsult(4) & " " & rsConsult(5))
         Else
            Call agregarChar(I, rsConsult(1))
         End If
         
         rsConsult.MoveNext
      Loop
      rsConsult.Close
   End If
Next

'------------------------------------------------------------------
'Busco los valores de los tipo estructura - Código Externo
'------------------------------------------------------------------
Flog.writeline "Busco los valores de los tipo estructura - Código Externo"
For I = 1 To Nro_Col

   If TipoCols(I) = "TCE" Then
   
      StrSql = " SELECT 'TCE', estrcodext "
      StrSql = StrSql & " FROM his_estructura "
      StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
      StrSql = StrSql & "    AND his_estructura.ternro = " & Ternro
      StrSql = StrSql & "    AND (htetdesde<=" & ConvFecha(fecEstr) & " AND (htethasta is null or htethasta>=" & ConvFecha(fecEstr) & "))"
      StrSql = StrSql & " AND his_estructura.tenro = " & CodCols(I)
            
      OpenRecordset StrSql, rsConsult

      Do Until rsConsult.EOF
         Call agregarChar(I, rsConsult(1))
         rsConsult.MoveNext
      Loop
      rsConsult.Close
   End If
Next

'------------------------------------------------------------------
'Busco la edad del empleado
'------------------------------------------------------------------
Flog.writeline "Busco la edad del empleado"
Dim Edad As Long
Dim FechaNacimiento As String
Dim FechaInicio As Date

For I = 1 To Nro_Col

   If TipoCols(I) = "EDA" Then

      If CodCols(I) = 1 Then
         FechaInicio = FechaDesde
      ElseIf CodCols(I) = 2 Then
             FechaInicio = FechaDesde
      ElseIf CodCols(I) = 3 Then
             FechaInicio = FechaHasta
      ElseIf CodCols(I) = 4 Then
             FechaInicio = FechaHasta
      End If
      
      StrSql = " SELECT terfecnac "
      StrSql = StrSql & " FROM tercero "
      StrSql = StrSql & " WHERE tercero.ternro = " & Ternro
            
      OpenRecordset StrSql, rsConsult

      If Not rsConsult.EOF Then
         FechaNacimiento = rsConsult(0)
      End If

      If IsNull(FechaNacimiento) Or FechaNacimiento = "" Then
         Edad = 0
      Else
           If (Month(FechaInicio) > Month(CDate(FechaNacimiento))) Then
               Edad = DateDiff("yyyy", CDate(FechaNacimiento), FechaInicio)
           Else
               If (Month(FechaInicio) = Month(CDate(FechaNacimiento))) And (Day(FechaInicio) >= Day(CDate(FechaNacimiento))) Then
                  Edad = DateDiff("yyyy", CDate(FechaNacimiento), FechaInicio)
               Else
                  Edad = DateDiff("yyyy", CDate(FechaNacimiento), FechaInicio) - 1
               End If
           End If
      End If
      rsConsult.Close
      
      Call agregarValorEdad(I, Edad)
   End If
Next

'------------------------------------------------------------------
'Busco la antiguedad del empleado
'------------------------------------------------------------------
Flog.writeline "Busco la antiguedad del empleado"
Dim Texto As String
Dim antdia As Integer
Dim antmes As Integer
Dim antanio As Integer
Dim q As Integer

For I = 1 To Nro_Col

   If TipoCols(I) = "ANT" Then

      If CodCols(I) = 1 Then
         FechaInicio = FechaDesde
      ElseIf CodCols(I) = 2 Then
             FechaInicio = FechaDesde
      ElseIf CodCols(I) = 3 Then
             FechaInicio = FechaHasta
      ElseIf CodCols(I) = 4 Then
             FechaInicio = FechaHasta
      ElseIf CodCols(I) = 5 Then
             FechaInicio = C_Date(fecEstr)
      End If

      'Calcula la antiguedad en dias, meses y años
      Call Antiguedad(Ternro, "REAL", FechaInicio, antdia, antmes, antanio, q)
      If antanio = 0 Then
         If antmes = 0 Then
            Texto = antdia & " día/s."
         Else
            Texto = antmes & " mes/es "
            If antdia <> 0 Then
               Texto = Texto & antdia & " día/s."
            End If
         End If
      Else
          Texto = antanio & " año/s "
          If antmes = 0 Then
             If antdia <> 0 Then
                Texto = Texto & antdia & " día/s."
             End If
          Else
             Texto = Texto & antmes & " mes/es "
             If antdia <> 0 Then
                Texto = Texto & antdia & " día/s."
             End If
          End If
      End If
      
      Call agregarChar(I, Texto)
      
   End If
Next

'------------------------------------------------------------------
'Busco los valores del telefono
'------------------------------------------------------------------
Flog.writeline "Busco los valores del telefono"
'Dim TipoDomi
Dim NombreCampo As String

For I = 1 To Nro_Col
    If TipoCols(I) = "TEL" Then
        CodNov = Split(CodNovCols(I), "@")
        TipoDomi = CodNov(0) '
        
        Select Case CodNov(1)
          Case 1: 'Telefono Principal
                 NombreCampo = "telefono.teldefault"
          Case 2: 'Telefono Celular
                 NombreCampo = "telefono.telcelular"
          Case 3: 'Telefono Fax
                 NombreCampo = "telefono.telfax"
        End Select
                    
        StrSql = " SELECT 'TEL', telnro "
        StrSql = StrSql & " FROM cabdom "
        StrSql = StrSql & " INNER JOIN telefono ON telefono.domnro = cabdom.domnro AND " & NombreCampo & " = -1 "
        StrSql = StrSql & " WHERE cabdom.ternro = " & Ternro & " AND cabdom.tidonro = " & TipoDomi
        
        OpenRecordset StrSql, rsConsult
        
        If Not rsConsult.EOF Then
          Call agregarChar(I, rsConsult(1))
        End If
        rsConsult.Close
        
   End If
Next

        
'------------------------------------------------------------------
'Busco los valores del Analisis Detallado
'------------------------------------------------------------------
Flog.writeline "Busco del Analisis Detallado"


StrSql = " SELECT 'ADE', traza.concnro, traza.tpanro, sum(traza.travalor) "
StrSql = StrSql & " FROM cabliq "
StrSql = StrSql & " INNER JOIN proceso  ON cabliq.pronro = proceso.pronro  AND proceso.pliqnro = " & pliqNro & " AND cabliq.pronro IN (" & ListaProcesos & ") "
StrSql = StrSql & " INNER JOIN traza   ON cabliq.cliqnro = traza.cliqnro  AND cabliq.empleado = " & Ternro & " AND cabliq.pronro = proceso.pronro "
StrSql = StrSql & " AND ( 1=0 "
For I = 1 To Nro_Col
   CodNov = Split(CodNovCols(I), "@")
   If TipoCols(I) = "ADE" Then
      StrSql = StrSql & " OR (traza.concnro = " & CodNov(0) & " AND traza.tpanro = " & CodNov(1) & ")"
   End If
Next
StrSql = StrSql & " ) "
StrSql = StrSql & " GROUP BY traza.concnro, traza.tpanro "

OpenRecordset StrSql, rsConsult

Do Until rsConsult.EOF
   Call agregarValorNov(rsConsult(0), rsConsult(1) & "@" & rsConsult(2), rsConsult(3))
   rsConsult.MoveNext
Loop
rsConsult.Close
        
'------------------------------------------------------------------
'Busco los valores del Periodo
'------------------------------------------------------------------
Flog.writeline "Busco los valores del periodo"
Dim textoCampo As String

For I = 1 To Nro_Col
    If TipoCols(I) = "PER" Then
          
        Select Case CodCols(I)
            Case 1: 'Fecha de Inicio
                NombreCampo = "pliqdesde"
            Case 2: 'Fecha de Fin
                NombreCampo = "pliqhasta"
            Case 3: 'Descripcion
                NombreCampo = "pliqdesc"
            Case 4: 'Periodo MMM-AA
                NombreCampo = " pliqmes, pliqanio"
            Case 5: 'Periodo MM/AAAA
                NombreCampo = " pliqmes, pliqanio"
            Case 6: 'Periodo MM/AA
                NombreCampo = " pliqmes, pliqanio"
        End Select
        
        'StrSql = " SELECT 'PER', " & NombreCampo
        'StrSql = StrSql & " FROM proceso "
        'StrSql = StrSql & " INNER JOIN periodo ON proceso.pliqnro = periodo.pliqnro "
        'StrSql = StrSql & " INNER JOIN tipoproc ON proceso.tprocnro = tipoproc.tprocnro "
        'StrSql = StrSql & " WHERE proNro = " & proNro
        
        StrSql = " SELECT 'PER', " & NombreCampo
        StrSql = StrSql & " FROM periodo "
        StrSql = StrSql & " WHERE pliqnro = " & pliqNro
        
        OpenRecordset StrSql, rsConsult
        
        If Not rsConsult.EOF Then
            Select Case CodCols(I)
                Case 1: 'Fecha de Inicio
                    textoCampo = rsConsult(1)
                Case 2: 'Fecha de Fin
                    textoCampo = rsConsult(1)
                Case 3: 'Descripcion
                    textoCampo = rsConsult(1)
                Case 4: 'Periodo MMM-AA
                    textoCampo = Format(rsConsult(1) & "-" & rsConsult(2), "MMM-yy")
                Case 5: 'Periodo MM/AAAA
                    textoCampo = Format(rsConsult(1) & "-" & rsConsult(2), "MM-yyyy")
                Case 6: 'Periodo MM/AA
                    textoCampo = Format(rsConsult(1) & "-" & rsConsult(2), "MM-yy")
            End Select
          Call agregarChar(I, textoCampo)
        End If
        rsConsult.Close
   End If
Next


'-------------------------------------------------------------------------------
'Inserto los datos en la BD
'-------------------------------------------------------------------------------
Flog.writeline "Inserto los datos en la BD - Legajo = " & Legajo & " TERNRO " & Ternro & " proceso " & proNro & " periodo " & pliqNro
StrSql = " INSERT INTO rep_estadisticaEmpleado_det ( " & _
         " bpronro , legajo, ternro, apellido, apellido2, " & _
         " nombre , nombre2, pronro, pliqnro,pliqfecha, " & _
         " prodesc , pliqdesc, estrdabr1, estrdabr2, estrdabr3, " & _
         " colval1 , colval2, colval3, colval4, colval5, " & _
         " colval6 , colval7, colval8, colval9, colval10," & _
         " colval11 , colval12, colval13, colval14, colval15, " & _
         " colval16 , colval17, colval18, colval19, colval20, " & _
         " colval21 , colval22, colval23, colval24, colval25, " & _
         " colval26 , colval27, colval28, colval29, colval30, " & _
         " colval31 , colval32, colval33, colval34, colval35, " & _
         " colval36 , colval37, colval38, colval39, colval40, " & _
         " colval41 , colval42, colval43, colval44, colval45, " & _
         " colval46 , colval47, colval48, colval49, colval50, " & _
         " colchar1 , colchar2, colchar3, colchar4, colchar5, " & _
         " colchar6 , colchar7, colchar8, colchar9, colchar10," & _
         " colchar11 , colchar12, colchar13, colchar14, colchar15, " & _
         " colchar16 , colchar17, colchar18, colchar19, colchar20, " & _
         " colchar21 , colchar22, colchar23, colchar24, colchar25, " & _
         " colchar26 , colchar27, colchar28, colchar29, colchar30, " & _
         " colchar31 , colchar32, colchar33, colchar34, colchar35, " & _
         " colchar36 , colchar37, colchar38, colchar39, colchar40, " & _
         " colchar41 , colchar42, colchar43, colchar44, colchar45, " & _
         " colchar46 , colchar47, colchar48, colchar49, colchar50, " & _
         " orden ) VALUES ( "

        StrSql = StrSql & NroProceso & ","
        StrSql = StrSql & Legajo & ","
        StrSql = StrSql & Ternro & ","
        StrSql = StrSql & "'" & apellido & "',"
        StrSql = StrSql & "'" & apellido2 & "',"
        StrSql = StrSql & "'" & nombre & "',"
        StrSql = StrSql & "'" & nombre2 & "',"
        StrSql = StrSql & proNro & ","
        StrSql = StrSql & pliqNro & ","
        StrSql = StrSql & ConvFecha(pliqFecha) & ","
        StrSql = StrSql & "null,"
        StrSql = StrSql & "'" & pliqDesc & "',"
        StrSql = StrSql & controlNull(estrnomb1) & ","
        StrSql = StrSql & controlNull(estrnomb2) & ","
        StrSql = StrSql & controlNull(estrnomb3) & ","
        For I = 1 To 50
            StrSql = StrSql & numberForSQL(ValCols(I)) & ","
        Next
        
        For I = 1 To 50
            StrSql = StrSql & "'" & Mid(CharCols(I), 1, 60) & "',"
        Next
        StrSql = StrSql & orden & ")"
        objConn.Execute StrSql, , adExecuteNoRecords

Exit Sub

MError:
    Flog.writeline "Error en empleado: " & Legajo & " Error: " & Err.Description
    Flog.writeline "SQL: " & StrSql
    HuboErrores = True
    EmpErrores = True
    Exit Sub
End Sub
Public Sub conceptosImpri(Desde, Hasta, pliqNro, ListaProcesos, Ternro, leg)
Dim concAnt
Dim I
Dim StrSql As String
Dim rsConsult As New ADODB.Recordset
Dim pliqnroAnt

StrSql = " SELECT 'CTC', detliq.concnro, sum(detliq.dlicant) AS cant, sum(detliq.dlimonto) AS monto, periodo.pliqnro,periodo.pliqmes,periodo.pliqanio,concepto.concabr,concepto.conccod  "
StrSql = StrSql & " FROM cabliq "
StrSql = StrSql & " INNER JOIN proceso  ON cabliq.pronro = proceso.pronro  AND proceso.pliqnro IN (" & pliqNro & ") AND cabliq.pronro IN (" & ListaProcesos & ") "
StrSql = StrSql & " INNER JOIN detliq   ON cabliq.cliqnro = detliq.cliqnro  AND cabliq.empleado = " & Ternro & " AND cabliq.pronro IN (" & ListaProcesos & ") "
StrSql = StrSql & " INNER JOIN concepto on concepto.concnro=detliq.concnro "
StrSql = StrSql & " INNER JOIN periodo on periodo.pliqnro = proceso.pliqnro "
  
For I = 1 To Nro_Col
   If TipoCols(I) = "CTC" Then
        If CodCols(I) <> 1 Then
            If CodCols(I) = 2 Then
                CodCols(I) = 0
            End If
            StrSql = StrSql & " AND concepto.concimp = " & CodCols(I)
        End If
   End If
Next
    
StrSql = StrSql & " GROUP BY detliq.concnro, periodo.pliqnro,periodo.pliqmes ,periodo.pliqanio,concepto.concabr,concepto.conccod   "
StrSql = StrSql & " ORDER BY detliq.concnro, periodo.pliqnro,periodo.pliqmes,periodo.pliqanio,concepto.concabr,concepto.conccod  "
    
    
Flog.writeline "borrarChar"
Call borrarValores

OpenRecordset StrSql, rsConsult
concAnt = ""
pliqnroAnt = ""
Do While Not rsConsult.EOF
    If (rsConsult!ConcNro <> concAnt) Then
        StrSql = " INSERT INTO repEstadisticaDet "
        StrSql = StrSql & " (bpronro,ternro,empleg,nombre,conccod,detnro,detalle, "
        If rsConsult!pliqanio = Desde Then
            StrSql = StrSql & " periodo" & rsConsult!pliqmes & ")"
        Else
            StrSql = StrSql & " periodo" & (rsConsult!pliqmes + 12) & ")"
        End If
        StrSql = StrSql & " VALUES "
        StrSql = StrSql & "(" & NroProceso & "," & Ternro & "," & leg & ",'" & nomape & "'," & rsConsult!Conccod & "," & rsConsult!ConcNro & ",'" & rsConsult!concabr & "', "
        If EsMonto(Nro_Col) Then
            StrSql = StrSql & Replace(rsConsult!Monto, ",", ".") & " )"
        Else
            StrSql = StrSql & Replace(rsConsult!cant, ",", ".") & " )"
        End If
        concAnt = rsConsult!ConcNro
        pliqnroAnt = rsConsult!pliqNro
    Else
        StrSql = " UPDATE repEstadisticaDet SET "
        If rsConsult!pliqanio = Desde Then
            StrSql = StrSql & " periodo" & rsConsult!pliqmes & "="
        Else
            StrSql = StrSql & " periodo" & (rsConsult!pliqmes + 12) & "="
        End If
        
        If EsMonto(Nro_Col) Then
            StrSql = StrSql & Replace(rsConsult!Monto, ",", ".")
        Else
            StrSql = StrSql & Replace(rsConsult!cant, ",", ".")
        End If
        StrSql = StrSql & " WHERE ternro=" & Ternro
        StrSql = StrSql & " AND bpronro=" & NroProceso
        StrSql = StrSql & " AND detnro=" & rsConsult!ConcNro
    End If
    objConn.Execute StrSql, , adExecuteNoRecords
   rsConsult.MoveNext
Loop
rsConsult.Close
End Sub
Public Sub conceptosNoImpri(Desde, Hasta, pliqNro, ListaProcesos, Ternro, leg)
Dim concAnt
Dim StrSql As String
Dim rsConsult As New ADODB.Recordset
Dim I

StrSql = " SELECT 'CTC', detliq.concnro, sum(detliq.dlicant) AS cant, sum(detliq.dlimonto) AS monto, periodo.pliqnro,periodo.pliqmes,periodo.pliqanio,concepto.concabr, concepto.conccod  "
StrSql = StrSql & " FROM cabliq "
StrSql = StrSql & " INNER JOIN proceso  ON cabliq.pronro = proceso.pronro  AND proceso.pliqnro IN (" & pliqNro & ") AND cabliq.pronro IN (" & ListaProcesos & ") "
StrSql = StrSql & " INNER JOIN detliq   ON cabliq.cliqnro = detliq.cliqnro  AND cabliq.empleado = " & Ternro & " AND cabliq.pronro IN (" & ListaProcesos & ") "
StrSql = StrSql & " INNER JOIN concepto on concepto.concnro=detliq.concnro "
StrSql = StrSql & " INNER JOIN periodo on periodo.pliqnro = proceso.pliqnro "
  
For I = 1 To Nro_Col
   If TipoCols(I) = "CTC" Then
        If CodCols(I) <> 1 Then
            If CodCols(I) = 2 Then
                CodCols(I) = 0
            End If
            StrSql = StrSql & " AND concepto.concimp = " & CodCols(I)
        End If
   End If
Next
    
StrSql = StrSql & " GROUP BY detliq.concnro, periodo.pliqnro,periodo.pliqmes ,periodo.pliqanio,concepto.concabr,concepto.conccod   "
StrSql = StrSql & " ORDER BY detliq.concnro, periodo.pliqnro,periodo.pliqmes,periodo.pliqanio,concepto.concabr,concepto.conccod  "
    
    
Flog.writeline "borrarChar"
Call borrarValores

OpenRecordset StrSql, rsConsult
concAnt = ""
Do While Not rsConsult.EOF
    If rsConsult!ConcNro <> concAnt Then
        StrSql = " INSERT INTO repEstadisticaDet "
        StrSql = StrSql & " (bpronro,ternro,empleg,nombre,conccod,detnro,detalle, "
        If rsConsult!pliqanio = Desde Then
            StrSql = StrSql & " periodo" & rsConsult!pliqmes & ")"
        Else
            StrSql = StrSql & " periodo" & (rsConsult!pliqmes + 12) & ")"
        End If
        StrSql = StrSql & " VALUES "
        StrSql = StrSql & "(" & NroProceso & "," & Ternro & "," & leg & ",'" & nomape & "'," & rsConsult!Conccod & "," & rsConsult!ConcNro & "," & rsConsult!concabr & ", "
        If EsMonto(Nro_Col) Then
            StrSql = StrSql & Replace(rsConsult!Monto, ",", ".") & " )"
        Else
            StrSql = StrSql & Replace(rsConsult!cant, ",", ".") & " )"
        End If
        concAnt = rsConsult!ConcNro
    Else
        StrSql = " UPDATE repEstadisticaDet  SET"
        If rsConsult!pliqanio = Desde Then
            StrSql = StrSql & " periodo" & rsConsult!pliqmes & "="
        Else
            StrSql = StrSql & " periodo" & (rsConsult!pliqmes + 12) & "="
        End If
        If EsMonto(Nro_Col) Then
            StrSql = StrSql & Replace(rsConsult!Monto, ",", ".")
        Else
            StrSql = StrSql & Replace(rsConsult!cant, ",", ".")
        End If
        StrSql = StrSql & " WHERE ternro=" & Ternro
        StrSql = StrSql & " AND bpronro=" & NroProceso
        StrSql = StrSql & " AND detnro=" & rsConsult!ConcNro
    End If
    objConn.Execute StrSql, , adExecuteNoRecords
    rsConsult.MoveNext
Loop
rsConsult.Close

End Sub
Public Sub acumuladoresImpri(Desde, Hasta, pliqNro, ListaProcesos, Ternro, leg)
Dim concAnt
Dim StrSql As String
Dim rsConsult As New ADODB.Recordset
Dim I

StrSql = " SELECT 'CTA', acu_liq.acunro, sum(acu_liq.alcant) AS cant, sum(acu_liq.almonto) AS monto, periodo.pliqnro,periodo.pliqmes,periodo.pliqanio,acumulador.acudesabr  "
StrSql = StrSql & " FROM cabliq "
StrSql = StrSql & " INNER JOIN proceso  ON cabliq.pronro = proceso.pronro  AND proceso.pliqnro IN (" & pliqNro & ") AND cabliq.pronro IN (" & ListaProcesos & ") "
StrSql = StrSql & " INNER JOIN acu_liq  ON cabliq.cliqnro = acu_liq.cliqnro AND cabliq.empleado = " & Ternro & " AND cabliq.pronro IN (" & ListaProcesos & ") "
StrSql = StrSql & " INNER JOIN acumulador on acumulador.acunro=acu_liq.acunro "
StrSql = StrSql & " INNER JOIN periodo on periodo.pliqnro = proceso.pliqnro "
  
For I = 1 To Nro_Col
   If TipoCols(I) = "CTA" Then
        If CodCols(I) <> 1 Then
            If CodCols(I) = 2 Then
                CodCols(I) = 0
            End If
            StrSql = StrSql & " AND acumulador.acuimpri = " & CodCols(I)
        End If
   End If
Next
    
StrSql = StrSql & " GROUP BY acu_liq.acunro, periodo.pliqnro,periodo.pliqmes ,periodo.pliqanio,acumulador.acudesabr   "
StrSql = StrSql & " ORDER BY acu_liq.acunro, periodo.pliqnro,periodo.pliqmes,periodo.pliqanio,acumulador.acudesabr  "
    
    
Flog.writeline "borrarChar"
Call borrarValores

OpenRecordset StrSql, rsConsult
concAnt = ""
Do While Not rsConsult.EOF
    If rsConsult!acuNro <> concAnt Then
        StrSql = " INSERT INTO repEstadisticaDet "
        StrSql = StrSql & " (bpronro,ternro,empleg,nombre,conccod,detnro,detalle, "
        If rsConsult!pliqanio = Desde Then
            StrSql = StrSql & " periodo" & rsConsult!pliqmes & ")"
        Else
            StrSql = StrSql & " periodo" & (rsConsult!pliqmes + 12) & ")"
        End If
        StrSql = StrSql & " VALUES "
        StrSql = StrSql & "(" & NroProceso & "," & Ternro & "," & leg & ",'" & nomape & "'," & rsConsult!acuNro & "," & rsConsult!acuNro & ",'" & rsConsult!acudesabr & "', "
        If EsMonto(Nro_Col) Then
            StrSql = StrSql & Replace(rsConsult!Monto, ",", ".") & " )"
        Else
            StrSql = StrSql & Replace(rsConsult!cant, ",", ".") & " )"
        End If
        concAnt = rsConsult!acuNro
    Else
        StrSql = " UPDATE repEstadisticaDet SET "
        If rsConsult!pliqanio = Desde Then
            StrSql = StrSql & " periodo" & rsConsult!pliqmes & "="
        Else
            StrSql = StrSql & " periodo" & (rsConsult!pliqmes + 12) & "="
        End If
        If EsMonto(Nro_Col) Then
            StrSql = StrSql & Replace(rsConsult!Monto, ",", ".")
        Else
            StrSql = StrSql & Replace(rsConsult!cant, ",", ".")
        End If
        StrSql = StrSql & " WHERE ternro=" & Ternro
        StrSql = StrSql & " AND bpronro=" & NroProceso
        StrSql = StrSql & " AND detnro=" & rsConsult!acuNro
    End If
    objConn.Execute StrSql, , adExecuteNoRecords
    rsConsult.MoveNext
Loop
rsConsult.Close

End Sub
Public Sub acumuladoresNoImpri(Desde, Hasta, pliqNro, ListaProcesos, Ternro, leg)

End Sub

Public Sub buscarConc(Desde, Hasta, pliqNro, ListaProcesos, Ternro, cont, leg)
Dim I
Dim concAnt
Dim StrSql As String
Dim rsConsult As New ADODB.Recordset
'------------------------------------------------------------------
'Busco los valores de los conceptos y acumuladores
'------------------------------------------------------------------

StrSql = " SELECT 'CO', detliq.concnro, sum(detliq.dlicant) AS cant, sum(detliq.dlimonto) AS monto, periodo.pliqnro,periodo.pliqmes,periodo.pliqanio,concepto.concabr,concepto.conccod   "
StrSql = StrSql & " FROM cabliq "
StrSql = StrSql & " INNER JOIN proceso  ON cabliq.pronro = proceso.pronro  AND proceso.pliqnro IN (" & pliqNro & ") AND cabliq.pronro IN (" & ListaProcesos & ") "
StrSql = StrSql & " INNER JOIN detliq   ON cabliq.cliqnro = detliq.cliqnro  AND cabliq.empleado = " & Ternro & " AND cabliq.pronro IN (" & ListaProcesos & ") "
StrSql = StrSql & " INNER JOIN concepto on concepto.concnro=detliq.concnro "
StrSql = StrSql & " INNER JOIN periodo on periodo.pliqnro = proceso.pliqnro "
  
'For I = 1 To Nro_Col
   If TipoCols(cont) = "CO" Then
      StrSql = StrSql & " AND detliq.concnro in (" & CodCols(cont) & ")"
   End If
'Next
    

StrSql = StrSql & " GROUP BY detliq.concnro, periodo.pliqnro,periodo.pliqmes ,periodo.pliqanio,concepto.concabr,concepto.conccod "
StrSql = StrSql & " ORDER BY periodo.pliqnro,periodo.pliqmes,periodo.pliqanio,concepto.conccod"
Flog.writeline "borrarChar"

Call borrarValores

OpenRecordset StrSql, rsConsult
concAnt = ""
Do While Not rsConsult.EOF
    If rsConsult!ConcNro <> concAnt Then
        StrSql = " INSERT INTO repEstadisticaDet "
        StrSql = StrSql & " (bpronro,ternro,empleg,nombre,conccod,detnro,detalle, "
        If rsConsult!pliqanio = Desde Then
            StrSql = StrSql & " periodo" & rsConsult!pliqmes & ")"
        Else
            StrSql = StrSql & " periodo" & (rsConsult!pliqmes + 12) & ")"
        End If
        StrSql = StrSql & " VALUES "
        StrSql = StrSql & "(" & NroProceso & "," & Ternro & "," & leg & ",'" & nomape & "'," & rsConsult!Conccod & "," & rsConsult!ConcNro & ",'" & rsConsult!concabr & "', "
        If EsMonto(cont) Then
            StrSql = StrSql & Replace(rsConsult!Monto, ",", ".") & " )"
        Else
            StrSql = StrSql & Replace(rsConsult!cant, ",", ".") & " )"
        End If
        concAnt = rsConsult!ConcNro
    Else
        StrSql = " UPDATE repEstadisticaDet SET "
        If rsConsult!pliqanio = Desde Then
            StrSql = StrSql & " periodo" & rsConsult!pliqmes & "="
        Else
            StrSql = StrSql & " periodo" & (rsConsult!pliqmes + 12) & "="
        End If
        If EsMonto(cont) Then
            StrSql = StrSql & Replace(rsConsult!Monto, ",", ".")
        Else
            StrSql = StrSql & Replace(rsConsult!cant, ",", ".")
        End If
        StrSql = StrSql & " WHERE ternro=" & Ternro
        StrSql = StrSql & " AND bpronro=" & NroProceso
        StrSql = StrSql & " AND detnro=" & rsConsult!ConcNro
    End If
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.writeline "strsql:" & StrSql
    rsConsult.MoveNext
Loop
rsConsult.Close

End Sub
Public Sub buscarAcu(Desde, Hasta, pliqNro, ListaProcesos, Ternro, cont, leg)
Dim I
Dim concAnt
Dim StrSql As String
Dim rsConsult As New ADODB.Recordset
'------------------------------------------------------------------
'Busco los valores de los  acumuladores
'------------------------------------------------------------------

StrSql = StrSql & " SELECT 'AC', acu_liq.acunro concnro, sum(acu_liq.alcant) AS cant, sum(acu_liq.almonto) AS monto,periodo.pliqnro, periodo.pliqmes,periodo.pliqanio,acumulador.acudesabr concabr "
StrSql = StrSql & " FROM cabliq "
StrSql = StrSql & " INNER JOIN proceso  ON cabliq.pronro = proceso.pronro  AND proceso.pliqnro IN (" & pliqNro & ") AND cabliq.pronro IN (" & ListaProcesos & ") "
StrSql = StrSql & " INNER JOIN acu_liq  ON cabliq.cliqnro = acu_liq.cliqnro AND cabliq.empleado = " & Ternro & " AND cabliq.pronro IN (" & ListaProcesos & ") "
StrSql = StrSql & " INNER JOIN acumulador on acumulador.acunro=acu_liq.acunro "
StrSql = StrSql & " INNER JOIN periodo on periodo.pliqnro = proceso.pliqnro "
'For I = 1 To Nro_Col
    If TipoCols(cont) = "AC" Then
       StrSql = StrSql & " AND acu_liq.acunro in (" & CodCols(cont) & ")"
    End If
'Next
StrSql = StrSql & " GROUP BY acu_liq.acunro, periodo.pliqnro,periodo.pliqmes ,periodo.pliqanio,acumulador.acudesabr "
StrSql = StrSql & " ORDER BY periodo.pliqnro,periodo.pliqmes,periodo.pliqanio"
Flog.writeline "borrarChar"

Call borrarValores

OpenRecordset StrSql, rsConsult
concAnt = ""
Do While Not rsConsult.EOF
    If rsConsult!ConcNro <> concAnt Then
        StrSql = " INSERT INTO repEstadisticaDet "
        StrSql = StrSql & " (bpronro,ternro,empleg,nombre,conccod,detnro,detalle, "
        If rsConsult!pliqanio = Desde Then
            StrSql = StrSql & " periodo" & rsConsult!pliqmes & ")"
        Else
            StrSql = StrSql & " periodo" & (rsConsult!pliqmes + 12) & ")"
        End If
        StrSql = StrSql & " VALUES "
        StrSql = StrSql & "(" & NroProceso & "," & Ternro & "," & leg & ",'" & nomape & "'," & rsConsult!ConcNro & "," & rsConsult!ConcNro & ",'" & rsConsult!concabr & "', "
        If EsMonto(cont) Then
            StrSql = StrSql & Replace(rsConsult!Monto, ",", ".") & " )"
        Else
            StrSql = StrSql & Replace(rsConsult!cant, ",", ".") & " )"
        End If
        concAnt = rsConsult!ConcNro
    Else
        StrSql = " UPDATE repEstadisticaDet SET "
        If rsConsult!pliqanio = Desde Then
            StrSql = StrSql & " periodo" & rsConsult!pliqmes & "="
        Else
            StrSql = StrSql & " periodo" & (rsConsult!pliqmes + 12) & "="
        End If
        If EsMonto(cont) Then
            StrSql = StrSql & Replace(rsConsult!Monto, ",", ".")
        Else
            StrSql = StrSql & Replace(rsConsult!cant, ",", ".")
        End If
        StrSql = StrSql & " WHERE ternro=" & Ternro
        StrSql = StrSql & " AND bpronro=" & NroProceso
        StrSql = StrSql & " AND detnro=" & rsConsult!ConcNro
    End If
    objConn.Execute StrSql, , adExecuteNoRecords
    rsConsult.MoveNext
Loop
rsConsult.Close

End Sub

Public Sub CantConceptosAcumulador(Desde, Hasta, pliqNro, ListaProcesos, Ternro, cont, leg)
Dim I
Dim concAnt
Dim StrSql As String

Dim rsConsult As New ADODB.Recordset
Dim rsConsult2 As New ADODB.Recordset
Dim rsConsult3 As New ADODB.Recordset
Dim strsqlAux As String
Dim strsqlAux2 As String
Dim strsql3 As String
Dim valorAnt
Dim valorAnt2 As Double
Dim StrSql2 As String
'------------------------------------------------------------------
'Busco los valores de los  acumuladores
'------------------------------------------------------------------

'StrSql = StrSql & " SELECT 'AC', acu_liq.acunro concnro, sum(acu_liq.alcant) AS cant, sum(acu_liq.almonto) AS monto,periodo.pliqnro, periodo.pliqmes,periodo.pliqanio,acumulador.acudesabr concabr "
StrSql = StrSql & " SELECT 'AC', acu_liq.acunro concnro, sum(acu_liq.alcant) AS cant, sum(acu_liq.almonto) AS monto,periodo.pliqnro, periodo.pliqmes,periodo.pliqanio,acumulador.acudesabr concabr, proceso.pronro "
StrSql = StrSql & " FROM cabliq "
StrSql = StrSql & " INNER JOIN proceso  ON cabliq.pronro = proceso.pronro  AND proceso.pliqnro IN (" & pliqNro & ") AND cabliq.pronro IN (" & ListaProcesos & ") "
StrSql = StrSql & " INNER JOIN acu_liq  ON cabliq.cliqnro = acu_liq.cliqnro AND cabliq.empleado = " & Ternro & " AND cabliq.pronro IN (" & ListaProcesos & ") "
StrSql = StrSql & " INNER JOIN acumulador on acumulador.acunro=acu_liq.acunro "
StrSql = StrSql & " INNER JOIN periodo on periodo.pliqnro = proceso.pliqnro "
'For I = 1 To Nro_Col
    If TipoCols(cont) = "CCA" Then
       StrSql = StrSql & " AND acu_liq.acunro in (" & CodCols(cont) & ")"
    End If
'Next
'StrSql = StrSql & " GROUP BY acu_liq.acunro, periodo.pliqnro,periodo.pliqmes ,periodo.pliqanio,acumulador.acudesabr "
StrSql = StrSql & " GROUP BY acu_liq.acunro, periodo.pliqnro,periodo.pliqmes ,periodo.pliqanio,acumulador.acudesabr,proceso.pronro "
StrSql = StrSql & " ORDER BY periodo.pliqnro,periodo.pliqmes,periodo.pliqanio"
Flog.writeline "borrarChar"

Call borrarValores

OpenRecordset StrSql, rsConsult
concAnt = ""
Do While Not rsConsult.EOF
    If rsConsult!ConcNro <> concAnt Then
        StrSql = " INSERT INTO repEstadisticaDet "
        StrSql = StrSql & " (bpronro,ternro,empleg,nombre,conccod,detnro,detalle, "
        If rsConsult!pliqanio = Desde Then
            StrSql = StrSql & " periodo" & rsConsult!pliqmes & ")"
        Else
            StrSql = StrSql & " periodo" & (rsConsult!pliqmes + 12) & ")"
        End If
        StrSql = StrSql & " VALUES "
        StrSql = StrSql & "(" & NroProceso & "," & Ternro & "," & leg & ",'" & nomape & "'," & rsConsult!ConcNro & "," & rsConsult!ConcNro & ",'" & rsConsult!concabr & "', "
        If EsMonto(cont) Then
            StrSql = StrSql & Replace(rsConsult!Monto, ",", ".") & " )"
        Else
            StrSql = StrSql & Replace(rsConsult!cant, ",", ".") & " )"
        End If
        concAnt = rsConsult!ConcNro
        
        'codigo sebastian stremel 23/10/2012
        'strsqlAux = " SELECT concepto.concabr, concepto.conccod, concepto.concnro, concepto.concimp, detliq.dlicant, detliq.dlimonto "
        strsqlAux = "SELECT concepto.concabr, concepto.conccod, concepto.concnro, concepto.concimp, sum(detliq.dlicant) cant, sum(detliq.dlimonto) monto"
        strsqlAux = strsqlAux & " FROM cabliq "
        strsqlAux = strsqlAux & " INNER JOIN detliq ON cabliq.cliqnro = detliq.cliqnro AND cabliq.pronro IN (" & rsConsult!proNro & ") AND cabliq.empleado = " & Ternro & ""
        'ANTES INNER ESTAS 2 LINEAS SIG
        strsqlAux = strsqlAux & " INNER JOIN concepto ON concepto.concnro = detliq.concnro "
        strsqlAux = strsqlAux & " INNER JOIN con_acum ON detliq.concnro = con_acum.concnro AND con_acum.acunro in (" & CodCols(cont) & ")"
        strsqlAux = strsqlAux & " GROUP BY concepto.concabr, concepto.conccod, concepto.concnro, concepto.concimp"
        strsqlAux = strsqlAux & " ORDER BY concepto.conccod "
        OpenRecordset strsqlAux, rsConsult2
        Do While Not rsConsult2.EOF
            strsqlAux2 = " INSERT INTO repEstadisticaDet "
            strsqlAux2 = strsqlAux2 & " (bpronro,ternro,empleg,nombre,conccod,detnro,detalle, "
            If rsConsult!pliqanio = Desde Then
                strsqlAux2 = strsqlAux2 & " periodo" & rsConsult!pliqmes & ")"
            Else
                strsqlAux2 = strsqlAux2 & " periodo" & (rsConsult!pliqmes + 12) & ")"
            End If
            strsqlAux2 = strsqlAux2 & " VALUES "
            strsqlAux2 = strsqlAux2 & "(" & NroProceso & "," & Ternro & "," & leg & ",'" & nomape & "','" & rsConsult2!Conccod & "'," & rsConsult!ConcNro & ",'" & rsConsult2!concabr & "', "
            If EsMonto(cont) Then
                strsqlAux2 = strsqlAux2 & Replace(rsConsult2!Monto, ",", ".") & " )"
            Else
                strsqlAux2 = strsqlAux2 & Replace(rsConsult2!cant, ",", ".") & " )"
            End If
                    
        objConn.Execute strsqlAux2, , adExecuteNoRecords
        rsConsult2.MoveNext
        Loop
        rsConsult2.Close
        objConn.Execute StrSql, , adExecuteNoRecords
        'hasta aca
    Else
'seba 24/10
        'valorAnt = "periodo" & (rsConsult!pliqmes)
        StrSql = " UPDATE repEstadisticaDet SET "
        If rsConsult!pliqanio = Desde Then
        
            'busco el valor insertado para sumarlo - 14/06/2013
            StrSql2 = " SELECT periodo" & rsConsult!pliqmes & " valor FROM repEstadisticaDet "
            StrSql2 = StrSql2 & " WHERE ternro=" & Ternro
            StrSql2 = StrSql2 & " AND bpronro=" & NroProceso
            StrSql2 = StrSql2 & " AND conccod=" & rsConsult!ConcNro
            StrSql2 = StrSql2 & " AND detnro=" & CodCols(cont)
            OpenRecordset StrSql2, rsConsult2
            If Not rsConsult2.EOF Then
                If EsNulo(rsConsult2!Valor) Then
                    valorAnt2 = 0
                Else
                    valorAnt2 = rsConsult2!Valor
                End If
            Else
                valorAnt2 = 0
            End If
            rsConsult2.Close
            'hasta aca
        
            'StrSql = StrSql & " periodo" & rsConsult!pliqmes & "= periodo" & rsConsult!pliqmes & "+"
            StrSql = StrSql & " periodo" & rsConsult!pliqmes & "="
            'valorAnt = (rsConsult!Monto)
            'valorAnt2 = (rsConsult!Monto)
            'cantAnt = (rsConsult!cant)
        Else
            'StrSql = StrSql & " periodo" & (rsConsult!pliqmes + 12) & "= periodo" & (rsConsult!pliqmes + 12) & "+"
            StrSql = StrSql & " periodo" & (rsConsult!pliqmes + 12) & "="
            'valorAnt = (rsConsult!Monto)
            'valorAnt2 = (rsConsult!Monto)
            'cantAnt = (rsConsult!cant)
        End If
        If EsMonto(cont) Then
            StrSql = StrSql & Replace(rsConsult!Monto + valorAnt2, ",", ".")
            'valorAnt = (rsConsult!Monto)
           
           'busco el monto a sumar
           
           
            'StrSql = StrSql & Replace(rsConsult!Monto + valorAnt2, ",", ".")
            'StrSql = StrSql & Replace(rsConsult!Monto + valorAnt, ",", ".")
        Else
            'valorAnt = (rsConsult!Monto)
            'StrSql = StrSql & Replace(rsConsult!cant, ",", ".")
            StrSql = StrSql & Replace(rsConsult!Monto + valorAnt2, ",", ".")
        End If
        StrSql = StrSql & " WHERE ternro=" & Ternro
        StrSql = StrSql & " AND bpronro=" & NroProceso
        StrSql = StrSql & " AND conccod=" & rsConsult!ConcNro
        StrSql = StrSql & " AND detnro=" & CodCols(cont) 'sebastian stremel - 05/06/2013 CAS-19945 - NORTHGATEARINSO - ERROR EN REPORTE ESTADISTICO POR EMPLEADO
        objConn.Execute StrSql, , adExecuteNoRecords
'hasta aca
    
        'codigo sebastian stremel 23/10/2012
        'strsqlAux = " SELECT concepto.concabr, concepto.conccod, concepto.concnro, concepto.concimp, detliq.dlicant, detliq.dlimonto "
        strsqlAux = "SELECT concepto.concabr, concepto.conccod, concepto.concnro, concepto.concimp, sum(detliq.dlicant) cant, sum(detliq.dlimonto) monto"
        strsqlAux = strsqlAux & " FROM cabliq "
        strsqlAux = strsqlAux & " INNER JOIN detliq ON cabliq.cliqnro = detliq.cliqnro AND cabliq.pronro IN (" & rsConsult!proNro & ") AND cabliq.empleado = " & Ternro & ""
        'antes inner
        strsqlAux = strsqlAux & " INNER JOIN concepto ON concepto.concnro = detliq.concnro "
        strsqlAux = strsqlAux & " INNER JOIN con_acum ON detliq.concnro = con_acum.concnro AND con_acum.acunro in (" & CodCols(cont) & ")"
        strsqlAux = strsqlAux & " GROUP BY concepto.concabr, concepto.conccod, concepto.concnro, concepto.concimp"
        strsqlAux = strsqlAux & " ORDER BY concepto.conccod "
        OpenRecordset strsqlAux, rsConsult2
        Do While Not rsConsult2.EOF
            'SEBA 06/11/2012
            'si no existe el concepto entonces lo agrego
            strsqlAux2 = " SELECT * FROM repEstadisticaDet "
            strsqlAux2 = strsqlAux2 & "WHERE conccod=" & rsConsult2!Conccod
            strsqlAux2 = strsqlAux2 & "AND BPRONRO=" & NroProceso
            strsqlAux2 = strsqlAux2 & "AND ternro=" & Ternro
            strsqlAux2 = strsqlAux2 & "AND detnro =" & CodCols(cont) 'sebastian stremel - 05/06/2013 CAS-19945 - NORTHGATEARINSO - ERROR EN REPORTE ESTADISTICO POR EMPLEADO
            OpenRecordset strsqlAux2, rsConsult3
            If rsConsult3.EOF Then
                'si no existe lo inserto
                strsqlAux2 = " INSERT INTO repEstadisticaDet "
                strsqlAux2 = strsqlAux2 & " (bpronro,ternro,empleg,nombre,conccod,detnro,detalle, "
                If rsConsult!pliqanio = Desde Then
                    strsqlAux2 = strsqlAux2 & " periodo" & rsConsult!pliqmes & ")"
                Else
                    strsqlAux2 = strsqlAux2 & " periodo" & (rsConsult!pliqmes + 12) & ")"
                End If
                strsqlAux2 = strsqlAux2 & " VALUES "
                strsqlAux2 = strsqlAux2 & "(" & NroProceso & "," & Ternro & "," & leg & ",'" & nomape & "','" & rsConsult2!Conccod & "'," & rsConsult!ConcNro & ",'" & rsConsult2!concabr & "', "
                If EsMonto(cont) Then
                    strsqlAux2 = strsqlAux2 & Replace(rsConsult2!Monto, ",", ".") & " )"
                Else
                    strsqlAux2 = strsqlAux2 & Replace(rsConsult2!cant, ",", ".") & " )"
                End If
                    
                objConn.Execute strsqlAux2, , adExecuteNoRecords
                rsConsult2.MoveNext
                'hasta aca
            Else

                'HASTA ACA
                StrSql = " UPDATE repEstadisticaDet SET "
                If rsConsult!pliqanio = Desde Then
                    StrSql = StrSql & " periodo" & rsConsult!pliqmes & "="
                    valorAnt = (rsConsult3("periodo" & rsConsult!pliqmes))
                    If EsNulo(valorAnt) Then
                        valorAnt = 0
                    End If
                Else
                    StrSql = StrSql & " periodo" & (rsConsult!pliqmes + 12) & "="
                    
                    'valorAnt = ("rsConsult3!periodo" & rsConsult!pliqmes + 12)
                    valorAnt = (rsConsult3("periodo" & rsConsult!pliqmes + 12))
                    If EsNulo(valorAnt) Then
                        valorAnt = 0
                    End If
                End If
                If EsMonto(cont) Then
                    StrSql = StrSql & Replace((rsConsult2!Monto + valorAnt), ",", ".")
                Else
                    StrSql = StrSql & Replace((rsConsult2!cant + valorAnt), ",", ".")
                End If
                StrSql = StrSql & " WHERE ternro=" & Ternro
                StrSql = StrSql & " AND bpronro=" & NroProceso
                StrSql = StrSql & " AND conccod=" & rsConsult2!Conccod
                StrSql = StrSql & " AND detnro=" & CodCols(cont)
                objConn.Execute StrSql, , adExecuteNoRecords
                rsConsult2.MoveNext
            End If
            rsConsult3.Close
        Loop
        rsConsult2.Close
    End If
    'objConn.Execute StrSql, , adExecuteNoRecords
    'objConn.Execute StrSql, , adExecuteNoRecords
    rsConsult.MoveNext
Loop
rsConsult.Close

End Sub


Public Sub Antiguedad(ByVal Ternro As Long, ByVal TipoAnt As String, ByVal FechaFin As Date, ByRef Dia As Integer, ByRef Mes As Integer, ByRef Anio As Integer, ByRef DiasHabiles As Integer)
' ---------------------------------------------------------------------------------------------
' Descripcion: Calcula la antiguedad al dia de hoy de un empleado en :
'              dias hábiles(si es menor que un año) o en dias, meses y años en caso contrario.
'              antiguedad.p
' Autor      : JMH
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim aux1 As Long
Dim aux2 As Long
Dim aux3 As Long
Dim fecalta As Date
Dim fecbaja As Date
Dim Seguir As Date
Dim q As Integer

Dim NombreCampo As String

Dim rs_Fases As New ADODB.Recordset


NombreCampo = ""
DiasHabiles = 0

Select Case UCase(TipoAnt)
Case "SUELDO":
    NombreCampo = "sueldo"
Case "INDEMNIZACION":
    NombreCampo = "indemnizacion"
Case "VACACIONES":
    NombreCampo = "vacaciones"
Case "REAL":
    NombreCampo = "real"
Case Else
End Select

Dia = 0
Mes = 0
Anio = 0

StrSql = "SELECT * FROM fases WHERE empleado = " & Ternro & _
         " AND " & NombreCampo & " = -1 " & _
         " AND not altfec is null " & _
         " AND not (bajfec is null AND estado = 0)" & _
         " AND altfec <= " & ConvFecha(FechaFin)
OpenRecordset StrSql, rs_Fases

Do While Not rs_Fases.EOF
    fecalta = rs_Fases!altfec
  
    ' Verificar si se trata de un registro completo (alta/baja) o solo de un alta
    If rs_Fases!estado Then
        fecbaja = FechaFin ' solo es un alta, tomar el fecha-fin
    ElseIf rs_Fases!bajfec <= FechaFin Then
        fecbaja = rs_Fases!bajfec  ' se trata de un registro completo
    Else
        fecbaja = FechaFin ' hasta la fecha ingresada
    End If
    
    Flog.writeline "fase de " & fecalta & " a " & fecbaja
            
    Call DIF_FECHAS2(fecalta, fecbaja, aux1, aux2, aux3)
    If rs_Fases.RecordCount = 1 Then
        Dia = Dia + aux1
        Mes = Mes + aux2 + Int(Dia / 30)
        Anio = Anio + aux3 + Int(Mes / 12)
        Dia = Dia Mod 30
        Mes = Mes Mod 12
    Else
        Dia = Dia + aux1
        Mes = Mes + aux2 + Int(Dia / 30)
        Anio = Anio + aux3 + Int(Mes / 12)
        Dia = Dia Mod 30
        Mes = Mes Mod 12
    End If
        
    If Anio = 0 Then
        Call DiasTrab(Ternro, fecalta, fecbaja, aux1)
        DiasHabiles = DiasHabiles + aux1
    End If
    
siguiente:
    rs_Fases.MoveNext
Loop

If Anio <> 0 Then
    DiasHabiles = 0
End If


' Cierro todo y Libero
If rs_Fases.State = adStateOpen Then rs_Fases.Close
Set rs_Fases = Nothing

End Sub

Public Sub DiasTrab(ByVal Ternro As Long, ByVal Desde As Date, ByVal Hasta As Date, ByRef DiasH As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Calcula la cantidad de dias trabajados de acuerdo al turno en que se trabaja y
'              de acuerdo a los dias que figuran como feriados en la tabla de feriados.
' Autor      : FGZ
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim d1 As Integer
Dim d2 As Integer
Dim aux As Integer
Dim aux2 As Integer
Dim dxsem As Integer

Dim rs_pais As New ADODB.Recordset
Dim rs_feriados As New ADODB.Recordset

    dxsem = 5
    
    d1 = Weekday(Desde)
    d2 = Weekday(Hasta)
    
    aux = DateDiff("d", Desde, Hasta) + 1
    If aux < 7 Then
        DiasH = Minimo(aux, dxsem)
    Else
        If aux = 7 Then
            DiasH = dxsem
        Else
            aux2 = 8 - d1 + d2
            If aux2 < 7 Then
                aux2 = Minimo(aux2, dxsem)
            Else
                If aux2 = 7 Then
                    aux2 = dxsem
                End If
            End If
            
            If aux2 >= 7 Then
                aux2 = Abs(aux2 - 7) + Int(aux2 / 7) * dxsem
            Else
                aux2 = aux2 + Int((aux2 - aux2) / 7) * dxsem
            End If
        End If
    End If
    
    aux = 0
    
    StrSql = "SELECT * FROM pais INNER JOIN tercero ON tercero.paisnro = pais.paisnro WHERE tercero.ternro = " & Ternro
    OpenRecordset StrSql, rs_pais
    
    If Not rs_pais.EOF Then
        ' Resto los Feriados Nacionales
        StrSql = "SELECT * FROM feriado WHERE tipferinro = 2 " & _
                 " AND fericodext = '" & rs_pais!paisnro & "'" & _
                 " AND ferifecha >= " & ConvFecha(Desde) & _
                 " AND ferifecha < " & ConvFecha(Hasta)
        OpenRecordset StrSql, rs_feriados
        
        Do While Not rs_feriados.EOF
            If Weekday(rs_feriados!ferifecha) > 1 Then
                DiasH = DiasH - 1
            End If
            
            ' Siguiente Feriado
            rs_feriados.MoveNext
        Loop
    End If


    ' Resto los feriados por Convenio
    StrSql = "SELECT * FROM empleado INNER JOIN his_estructura ON empleado.ternro = his_estructura.ternro " & _
             " INNER JOIN fer_estr ON fer_estr.tenro = his_estructura.tenro " & _
             " INNER JOIN feriado ON fer_estr.ferinro = feriado.ferinro " & _
             " WHERE empleado.ternro = " & Ternro & _
             " AND feriado.tipferinro = 2" & _
             " AND feriado.ferifecha >= " & ConvFecha(Desde) & _
             " AND feriado.ferifecha < " & ConvFecha(Hasta)
    OpenRecordset StrSql, rs_feriados
    
    Do While Not rs_feriados.EOF
        If Weekday(rs_feriados!ferifecha) > 1 Then
            DiasH = DiasH - 1
        End If
        
        ' Siguiente Feriado
        rs_feriados.MoveNext
    Loop
    
    
    ' cierro todo y libero
    If rs_pais.State = adStateOpen Then rs_pais.Close
    If rs_feriados.State = adStateOpen Then rs_feriados.Close
        
    Set rs_feriados = Nothing
    Set rs_pais = Nothing

End Sub



