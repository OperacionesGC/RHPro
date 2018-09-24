Attribute VB_Name = "repListadoFam"
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

Global Pagina As Long
Global tipoModelo As Integer
Global arrTipoConc(1000) As Integer

Global tenro1 As Integer
Global estrnro1 As Integer
Global tenro2 As Integer
Global estrnro2 As Integer
Global tenro3 As Integer
Global estrnro3 As Integer
Global fecEstr As String
Global tipoFamiliares
Global NroBusqProg As Long
Global NroBusqProg2 As Long
Global NroBusqProg3 As Long
Global TipDocNroInscr As Long
Global empresa
Global emprNro
Global emprActiv
Global emprTer
Global emprDire
Global emprDire2
Global emprCuit
Global zonaDomicilio
Global concFamiliar01
Global concFamiliar02
Global concFamiliar03
Global param_empresa
Global listapronro
Global l_orden
Global filtro
Global totalEmpleados
Global cantRegistros
Global acumSueldoJornal

Global acumGrupo1(100) As Long
Global acumGrupo2(100) As Long
Global acumGrupo3(100) As Long
Global cantAcumGrupo1 As Long
Global cantAcumGrupo2 As Long
Global cantAcumGrupo3 As Long

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
Dim fechadesde
Dim fechahasta
Dim tipoDepuracion
Dim historico As Boolean
Dim param
Dim proNro
Dim ternro
Dim arrpronro
Dim rsEmpl As New ADODB.Recordset
Dim rsEmpresas As New ADODB.Recordset
Dim acunroSueldo
Dim i
Dim PID As String
Dim tituloReporte As String
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
    
    Nombre_Arch = PathFLog & "ReporteListadoFamiliar" & "-" & NroProceso & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    Flog.Writeline "Inicio Proceso Listado Familiar : " & Now
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
       
       'Obtengo los parametros del proceso
       parametros = objRs!bprcparam
       ArrParametros = Split(parametros, "@")
       
       'Obtengo la lista de procesos
       listapronro = ArrParametros(0)
       
       'Obtengo el tipo de familiares
       tipoFamiliares = ArrParametros(1)
       
       'Obtengo los cortes de estructura
       tenro1 = CInt(ArrParametros(2))
       estrnro1 = CInt(ArrParametros(3))
       tenro2 = CInt(ArrParametros(4))
       estrnro2 = CInt(ArrParametros(5))
       tenro3 = CInt(ArrParametros(6))
       estrnro3 = CInt(ArrParametros(7))
       fecEstr = ArrParametros(8)
       
       'Obtengo el titulo del reporte
       tituloReporte = ArrParametros(9)
       
       'Obtengo el nro. de Empresa
       param_empresa = ArrParametros(10)
       
       'EMPIEZA EL PROCESO
       
       'Busco en el confrep el numero de cuenta que se va a usar para
       ' buscar el valor del sueldo
       StrSql = " SELECT * FROM confrep "
       StrSql = StrSql & " WHERE repnro = 61 "
      
       OpenRecordset StrSql, objRs2
       
       If objRs2.EOF Then
          Flog.Writeline "No esta configurado el ConfRep"
          Exit Sub
       End If
       
       Flog.Writeline "Obtengo los datos del confrep"
       
       NroBusqProg = 0
       NroBusqProg2 = 0
       NroBusqProg3 = 0
       concFamiliar01 = 0
       concFamiliar02 = 0
       concFamiliar03 = 0
       
       Do Until objRs2.EOF
       
          Select Case objRs2!confnrocol
             Case 10
                  NroBusqProg = objRs2!confval
             Case 12
                  NroBusqProg2 = objRs2!confval
             Case 16
                  'Busco el concnro del concepto
                  StrSql = "SELECT concnro FROM concepto WHERE conccod = " & objRs2!confval
                  StrSql = StrSql & " OR conccod = '" & objRs2!confval2 & "'"
                  
                  OpenRecordset StrSql, objRs3
                  
                  If objRs3.EOF Then
                     concFamiliar01 = 0
                  Else
                     concFamiliar01 = objRs3!concnro
                  End If
                  
                  objRs3.Close
                  
             Case 17
                  'Busco el concnro del concepto
                  StrSql = "SELECT concnro FROM concepto WHERE conccod = " & objRs2!confval
                  StrSql = StrSql & " OR conccod = '" & objRs2!confval2 & "'"
                  
                  OpenRecordset StrSql, objRs3
                  
                  If objRs3.EOF Then
                     concFamiliar02 = 0
                  Else
                     concFamiliar02 = objRs3!concnro
                  End If
                  
                  objRs3.Close
                  
             Case 62
                  NroBusqProg3 = objRs2!confval
             
             Case 63
                  'Busco el concnro del concepto
                  StrSql = "SELECT concnro FROM concepto WHERE conccod = " & objRs2!confval
                  StrSql = StrSql & " OR conccod = '" & objRs2!confval2 & "'"
                  
                  OpenRecordset StrSql, objRs3
                  
                  If objRs3.EOF Then
                     concFamiliar03 = 0
                  Else
                     concFamiliar03 = objRs3!concnro
                  End If
                  
                  objRs3.Close
                  

          End Select
       
          objRs2.MoveNext
       Loop

       'Obtengo los empleados sobre los que tengo que buscar los familiares
       StrSql = "SELECT * FROM batch_empleado WHERE bpronro = " & NroProceso
       StrSql = StrSql & " ORDER BY progreso,estado"
       OpenRecordset StrSql, rsEmpl
       
       cantRegistros = rsEmpl.RecordCount
       totalEmpleados = cantRegistros

       
       StrSql = "UPDATE batch_proceso SET bprcprogreso = 0 " & _
                ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
            
       objConn.Execute StrSql, , adExecuteNoRecords
       
       'Obtengo la lista de procesos
       arrpronro = Split(listapronro, ",")
       
       'Obtengo los datos de la empresa
       '---LOG---
       Flog.Writeline "Buscando datos de la empresa"
    
       If param_empresa = 0 Then
           Flog.Writeline "No se encontró la empresa"
           Exit Sub
       Else
           'Consulta para obtener la direccion de la empresa
           StrSql = "SELECT empresa.empnro,empresa.empnom,empresa.empactiv,detdom.calle,detdom.nro,detdom.codigopostal,localidad.locdesc " & _
                    " FROM empresa" & _
                    " LEFT JOIN cabdom ON cabdom.ternro = empresa.ternro" & _
                    " LEFT JOIN detdom ON detdom.domnro = cabdom.domnro" & _
                    " LEFT JOIN localidad ON detdom.locnro = localidad.locnro " & _
                    " WHERE empresa.estrnro = " & param_empresa
        
           '---LOG---
           Flog.Writeline "Buscando datos de la direccion de la empresa"
        
           OpenRecordset StrSql, objRs2
        
           If objRs2.EOF Then
                Flog.Writeline "No se encontró el domicilio de la empresa"
                emprDire = "   "
           Else
                emprNro = objRs2!Empnro
                empresa = objRs2!empnom
                emprDire = objRs2!calle & " " & objRs2!Nro
                emprDire2 = objRs2!codigopostal & " " & objRs2!locdesc
                emprActiv = objRs2!empactiv
           End If
           
           If objRs2.State = adStateOpen Then objRs2.Close
           
       End If
       
       'Genero por cada empleado un registro
       arrpronro = Split(listapronro, ",")
       orden = 0
       
       Do Until rsEmpl.EOF
            EmpErrores = False
            ternro = rsEmpl!ternro
            
            'Genero una entrada para el empleado por cada proceso
            For i = 0 To UBound(arrpronro)
                proNro = arrpronro(i)
                Flog.Writeline "Generando datos empleado " & ternro & " para el proceso " & proNro
                orden = orden + 1
                
                Call generarDatosEmpleado01(proNro, ternro, tituloReporte, orden)
                     
            Next
                  
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


'--------------------------------------------------------------------
' Se encarga de generar los datos para el Standar y Deloitte
'--------------------------------------------------------------------
Sub generarDatosEmpleado01(proNro, ternro, descripcion, orden)

Dim StrSql As String
Dim rsConsult As New ADODB.Recordset

'Variables donde se guardan los datos del INSERT final

Dim Legajo As Long
Dim apellido As String
Dim apellido2 As String
Dim nombre As String
Dim nombre2 As String
Dim pliqnro
Dim pliqmes
Dim pliqanio
Dim prodesc
Dim pliqdesc
Dim pliqfecdep
Dim pliqbco
Dim terfecnac
Dim cliqnro
Dim GuardarFam As Boolean
Dim pliqdesde As Date
Dim pliqhasta As Date
Dim guardar_encabezado As Boolean

Dim EmpTernro

Dim estrnomb1
Dim estrnomb2
Dim estrnomb3
Dim tenomb1
Dim tenomb2
Dim tenomb3

Dim sql As String

On Error GoTo MError

estrnomb1 = ""
estrnomb2 = ""
estrnomb3 = ""
tenomb1 = ""
tenomb2 = ""
tenomb3 = ""

'------------------------------------------------------------------
'Busco los datos del periodo actual
'------------------------------------------------------------------
StrSql = " SELECT periodo.*, proceso.profecpago, proceso.prodesc, cabliq.cliqnro FROM periodo "
StrSql = StrSql & " INNER JOIN proceso ON proceso.pliqnro = periodo.pliqnro "
StrSql = StrSql & " AND proceso.pronro= " & proNro
StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
StrSql = StrSql & " AND cabliq.empleado= " & ternro

'---LOG---
Flog.Writeline "Buscando datos del periodo"

OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   pliqnro = rsConsult!pliqnro
   pliqmes = rsConsult!pliqmes
   pliqanio = rsConsult!pliqanio
   prodesc = rsConsult!prodesc
   pliqdesc = rsConsult!pliqdesc
   cliqnro = rsConsult!cliqnro
   pliqdesde = rsConsult!pliqdesde
   pliqhasta = rsConsult!pliqhasta
Else
   Flog.Writeline "El empleado no se encuentra en el proceso"
   Exit Sub
End If

rsConsult.Close

'------------------------------------------------------------------
'Busco los datos del empleado
'------------------------------------------------------------------
StrSql = " SELECT empleg,terape,terape2,ternom,ternom2,empfaltagr,empremu "
StrSql = StrSql & " FROM empleado "
StrSql = StrSql & " WHERE ternro= " & ternro

Flog.Writeline "Buscando datos del empleado"
       
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
   Flog.Writeline "Error al obtener los datos del empleado"
'   GoTo MError
End If

rsConsult.Close

'------------------------------------------------------------------
'Busco los datos del tipos de estructura 1
'------------------------------------------------------------------

'---LOG---
Flog.Writeline "Buscando datos estructura 1"

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

'---LOG---
Flog.Writeline "Buscando datos estructura 2"

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

'---LOG---
Flog.Writeline "Buscando datos estructura 3"

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
'Busco los datos de los familiares
'------------------------------------------------------------------

If (tipoFamiliares = "1") Or (tipoFamiliares = "2") Then

    StrSql = " SELECT docu.nrodoc AS nrodoc, tipodocu.tidsigla AS sigladoc,tercero.ternro,tercero.terape, tercero.ternom,tercero.terape2, tercero.ternom2, tercero.terfecnac, tercero.tersex, famest,famtrab,faminc,famDGIdesde " & _
          " famsalario, famfecvto, famCargaDGI, famDGIdesde, famDGIhasta, famemergencia, " & _
          " paredesc " & _
          " FROM  tercero INNER JOIN familiar ON tercero.ternro=familiar.ternro " & _
          " LEFT JOIN parentesco ON familiar.parenro=parentesco.parenro " & _
          " LEFT JOIN ter_doc docu ON (docu.ternro= familiar.ternro and docu.tidnro>0 and docu.tidnro<5) " & _
          " LEFT JOIN tipodocu ON tipodocu.tidnro= docu.tidnro " & _
          " WHERE familiar.famest = -1 AND familiar.empleado = " & ternro

    '---LOG---
    Flog.Writeline "Buscando datos de los familiares"
    
    OpenRecordset StrSql, rsConsult
       
    guardar_encabezado = True
    
    Do Until rsConsult.EOF
        
          GuardarFam = True
        
          If (tipoFamiliares = "2") Then
              GuardarFam = HayAsignacionesFliares(pliqdesde, pliqhasta, ternro, rsConsult!ternro, NroBusqProg, cliqnro, concFamiliar01)
              If Not GuardarFam Then
                 GuardarFam = HayAsignacionesFliares(pliqdesde, pliqhasta, ternro, rsConsult!ternro, NroBusqProg2, cliqnro, concFamiliar02)
                 If Not GuardarFam Then
                    GuardarFam = HayAsignacionesFliares(pliqdesde, pliqhasta, ternro, rsConsult!ternro, NroBusqProg3, cliqnro, concFamiliar03)
                 End If
              End If
          End If
            
          If GuardarFam Then
              If guardar_encabezado Then
                    '------------------------------------------------------------------
                    'Armo la SQL para guardar los datos
                    '------------------------------------------------------------------
                    StrSql = " INSERT INTO rep_list_fam (bpronro,legajo,ternro,pronro,prodesc,apellido,apellido2,nombre,"
                    StrSql = StrSql & "nombre2,empresa,emprdire,emprdire2,empractiv,emprNro,pliqnro,descripcion,pliqdesc,pliqmes,"
                    StrSql = StrSql & "pliqanio,orden,tedabr1,tedabr2,tedabr3,estrdabr1,estrdabr2,estrdabr3,tipofam)"
                    StrSql = StrSql & " VALUES "
                    StrSql = StrSql & "(" & NroProceso
                    StrSql = StrSql & "," & Legajo
                    StrSql = StrSql & "," & ternro
                    StrSql = StrSql & "," & proNro
                    StrSql = StrSql & ",'" & prodesc & "'"
                    StrSql = StrSql & ",'" & apellido & "'"
                    StrSql = StrSql & ",'" & apellido2 & "'"
                    StrSql = StrSql & ",'" & nombre & "'"
                    StrSql = StrSql & ",'" & nombre2 & "'"
                    StrSql = StrSql & ",'" & empresa & "'"
                    StrSql = StrSql & ",'" & emprDire & "'"
                    StrSql = StrSql & ",'" & emprDire2 & "'"
                    StrSql = StrSql & ",'" & emprActiv & "'"
                    StrSql = StrSql & "," & emprNro
                    StrSql = StrSql & "," & pliqnro
                    StrSql = StrSql & ",'" & Mid(descripcion, 1, 100) & "'"
                    StrSql = StrSql & ",'" & pliqdesc & "'"
                    StrSql = StrSql & "," & pliqmes
                    StrSql = StrSql & "," & pliqanio
                    StrSql = StrSql & "," & orden
                    StrSql = StrSql & "," & controlNull(tenomb1)
                    StrSql = StrSql & "," & controlNull(tenomb2)
                    StrSql = StrSql & "," & controlNull(tenomb3)
                    StrSql = StrSql & "," & controlNull(estrnomb1)
                    StrSql = StrSql & "," & controlNull(estrnomb2)
                    StrSql = StrSql & "," & controlNull(estrnomb3)
                    StrSql = StrSql & "," & tipoFamiliares
                    StrSql = StrSql & ")"
                     
                    objConn.Execute StrSql, , adExecuteNoRecords
                    
                    guardar_encabezado = False
              End If
              
              If IsNull(rsConsult!terfecnac) Then
                 terfecnac = "NULL"
              Else
                 terfecnac = ConvFecha(rsConsult!terfecnac)
              End If
            
              StrSql = " INSERT INTO rep_list_fam_det (bpronro,ternro,pronro,nrodoc,"
              StrSql = StrSql & "sigladoc,ternrofam,terape,ternom,terape2,ternom2,terfecnac,"
              StrSql = StrSql & "tersex,faminc,paredesc) "
              StrSql = StrSql & " VALUES "
              StrSql = StrSql & "(" & NroProceso
              StrSql = StrSql & "," & ternro
              StrSql = StrSql & "," & proNro
              StrSql = StrSql & ",'" & rsConsult!nrodoc & "'"
              StrSql = StrSql & ",'" & rsConsult!sigladoc & "'"
              StrSql = StrSql & "," & rsConsult!ternro
              StrSql = StrSql & ",'" & rsConsult!terape & "'"
              StrSql = StrSql & ",'" & rsConsult!ternom & "'"
              StrSql = StrSql & ",'" & rsConsult!terape2 & "'"
              StrSql = StrSql & ",'" & rsConsult!ternom2 & "'"
              StrSql = StrSql & "," & terfecnac
              StrSql = StrSql & "," & strForSQL(rsConsult!tersex)
              StrSql = StrSql & "," & strForSQL(rsConsult!faminc)
              StrSql = StrSql & ",'" & rsConsult!paredesc & "'"
              StrSql = StrSql & ")"
              
              objConn.Execute StrSql, , adExecuteNoRecords
              
          End If
        
          rsConsult.MoveNext
          
    Loop
    
    rsConsult.Close

End If

Exit Sub

MError:
    Flog.Writeline "Error en empleado: " & Legajo & " Error: " & Err.Description
    HuboErrores = True
    EmpErrores = True
    Exit Sub
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



Public Function HayAsignacionesFliares(pliqdesde As Date, pliqhasta As Date, ternro, ternroFamiliar, Busqueda, cliqnro, concepto)
' ---------------------------------------------------------------------------------------------
' Descripcion: Asignaciones Familiares
' Autor      : FGZ
' Fecha      : 15/04/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Incapacitado As Boolean         '-1 (True) / 0 (False)
Dim Edad As Integer                 'cant de años o nulo o vacio
Dim Sexo As Integer                 '1(Masc) / 2 (Fem) / 3 (Todos)
Dim Estudia As Integer              '1(si) / 2 (no) / 3 (indefinido)
Dim Ayuda_Escolar As Integer        '1(si) / 2 (no) / 3 (indefinido)
Dim Suma_FliaNumerosa As Integer    'Siempre viene 1. No se usa mas
Dim Paga_FliaNumerosa As Integer    'siempre viene 1. No se usa mas
Dim Trabaja_Conyuge As Integer      '1(si) / 2 (no) / 3 (no importa)
Dim Retroactivo_Prenatal As Boolean '-1(TRUE) / 0(FALSE)
Dim Nivel_Estudio As String         'nivnro,nivnro,....
Dim Periodo_Escolar As Integer      'nro del periodo escolar
Dim Parentesco As Integer           'codigo del parentesco
                            
Dim Fam_niv_est     As Integer
Dim Fam_peri_escol  As Integer
Dim Fam_estudia     As Integer
Dim Fecha_vto_asig  As Date
Dim Fin_periodo_liq As Date
Dim Par_asig        As Integer

Dim suma_fn As Integer
Dim paga_fn As Integer
Dim edad_f As Integer
Dim conyuge As Integer
Dim pagaxhijo As Boolean
Dim niv_est_interno As Boolean
Dim interesa_estu As Boolean
Dim sexo_conyuge As Integer
Dim conyuge_trabaja As Integer
Dim Opcion_Fecha_Hasta
Dim Fecha_Hasta_Edad

Dim Param_cur As New ADODB.Recordset
Dim rs_PeriodoEsc As New ADODB.Recordset
Dim rs_Familiar As New ADODB.Recordset
Dim rs_Tercero As New ADODB.Recordset
Dim rs_estudio_actual As New ADODB.Recordset
Dim rs_Nivest As New ADODB.Recordset
Dim rs_Concepto As New ADODB.Recordset

Dim Salida As Boolean
  
'inicializo
Par_asig = 31
conyuge = False
pagaxhijo = False
suma_fn = 0
paga_fn = 0

    Bien = False
    HayAsignacionesFliares = False
    Salida = False
    
    ' Obtener los parametros de la Busqueda
    StrSql = "SELECT * FROM programa WHERE progarchest = -1 AND prognro = " & CStr(Busqueda)
    OpenRecordset StrSql, Param_cur
    
    If Not Param_cur.EOF Then
        Incapacitado = Param_cur!Auxint4
        Edad = Param_cur!Auxint1
        Sexo = Param_cur!Auxint5
        Estudia = Param_cur!Auxchar2
        Ayuda_Escolar = Param_cur!Auxchar3
        Suma_FliaNumerosa = True
        Paga_FliaNumerosa = True
        Trabaja_Conyuge = Param_cur!Auxchar5
        Retroactivo_Prenatal = Param_cur!Auxint2
        Nivel_Estudio = IIf(EsNulo(Param_cur!Auxchar1), 0, Param_cur!Auxchar1)
        Periodo_Escolar = Param_cur!Auxint3
        Parentesco = Param_cur!Auxchar4
        
        Opcion_Fecha_Hasta = Param_cur!Auxchar
        
        If IsNull(Opcion_Fecha_Hasta) Then
           Opcion_Fecha_Hasta = 1
        End If
        
        Select Case CInt(Opcion_Fecha_Hasta)
        Case 1:
            'a mes actual
            Fecha_Hasta_Edad = pliqhasta
            
        Case 2:
            'al mes anterior
            Fecha_Hasta_Edad = DateAdd("d", -1, pliqdesde)
            
        Case 3:
            'a fin de año
            Fecha_Hasta_Edad = CDate("31/12/" & Year(pliqhasta))
            
        Case 4:
            'a principio de año
            Fecha_Hasta_Edad = CDate("01/01/" & Year(pliqhasta))
            
        Case Else:
            'Default. a fin de mes
            Fecha_Hasta_Edad = pliqhasta
            
        End Select
        
    Else
        Flog.Writeline "Error: No se encuentra o no esta generada la busqueda de familiares nro " & Busqueda & "."
        Exit Function
    End If

    ' VALIDAR SI AL EMPLEADO CORRESPONDE PAGARLE ASIGNACIONES FAMILIARES
    ' en funcion : dias trabajados en el mes
    
    'AYUDA ESCOLAR
    If Ayuda_Escolar = 1 Then
        StrSql = "SELECT * FROM edu_periodoesc WHERE edu_periodoesc.perescnro =" & Periodo_Escolar
        OpenRecordset StrSql, rs_PeriodoEsc
        If rs_PeriodoEsc.EOF Then
            Flog.Writeline "Error: Periodo Escolar Incorrecto."
            Exit Function
        End If
    End If
        
    'FECHAS LIMITES DE VENCIMENTOS DE CERTIFICADOS
    Fecha_vto_asig = pliqdesde
    Fin_periodo_liq = pliqhasta


    StrSql = "SELECT * FROM familiar INNER JOIN tercero ON tercero.ternro = familiar.ternro"
    StrSql = StrSql & " WHERE (familiar.empleado =" & ternro
    StrSql = StrSql & " AND familiar.parenro = " & Parentesco
    StrSql = StrSql & " AND familiar.famest = -1"
    StrSql = StrSql & " AND familiar.famsalario = -1"
    StrSql = StrSql & " AND familiar.ternro = " & ternroFamiliar & ")"
    StrSql = StrSql & " AND (familiar.famfecvto >=" & ConvFecha(Fecha_vto_asig) & " OR familiar.famfecvto is null)"
    StrSql = StrSql & " Order by tercero.terfecnac DESC"
    OpenRecordset StrSql, rs_Familiar
             
    Do While Not rs_Familiar.EOF
        If rs_Familiar!parenro = 2 Then 'hijo
            'calculo la edad del familiar
             edad_f = Calcular_Edad(rs_Familiar!terfecnac, Fecha_Hasta_Edad)
             
            conyuge = 3
            sexo_conyuge = 3
            conyuge_trabaja = 3
        End If
                             
        'en los conyuges no interesa si estudia o no
        interesa_estu = True  'el default es que interese
        
        If rs_Familiar!parenro = 3 Then 'conyuge
            sexo_conyuge = IIf(CBool(rs_Familiar!tersex), 1, 2)
            conyuge_trabaja = IIf(CBool(rs_Familiar!famtrab), 1, 2)
            If (rs_Familiar!tersex) And (rs_Familiar!faminc) And (Not rs_Familiar!famtrab) Then
                conyuge = False
            Else
                If (Not rs_Familiar!tersex) And rs_Familiar!famtrab Then
                    conyuge = True
                Else
                    conyuge = False
                End If
            End If
        
            If (sexo_conyuge = 1 And conyuge_trabaja = 1) Then 'Conyuge masculino y trabaja
               GoTo SiguienteFamiliar
            End If
            interesa_estu = False
            edad_f = 0
        Else
            conyuge = 3
            sexo_conyuge = 3
            conyuge_trabaja = 3
        End If
        
        'buscar el nivel de estudio
        Fam_estudia = False
        StrSql = "SELECT * FROM estudio_actual WHERE ternro = " & rs_Familiar!ternro
        OpenRecordset StrSql, rs_estudio_actual
        If Not rs_estudio_actual.EOF Then
            If Not EsNulo(rs_estudio_actual!nivnro) Then
                StrSql = "SELECT * FROM nivest WHERE nivnro =" & rs_estudio_actual!nivnro
                OpenRecordset StrSql, rs_Nivest
                If Not rs_Nivest.EOF Then
                    niv_est_interno = rs_Nivest!nivsist
                    Fam_niv_est = rs_Nivest!nivnro
                    Fam_estudia = IIf(EsNulo(rs_estudio_actual!nivnro), 2, 1)
                Else
                    niv_est_interno = False
                End If
            End If
        End If
            
            
        'ACA SE PODRIA VALIDAR LA VALIDEZ DEL CERTIFICADO ESCOLAR
        'INICIO
        'FIN
        'SI NO ES VALIDO EL CERTIFICADO, ASIGNAR A FAM-NIV-EST = ?
        'FAM-ESTUDIA = (estudio_actual.nivnro <> ?)
                   
        If (CBool(Incapacitado) = CBool(rs_Familiar!faminc)) Then
            If (Parentesco = rs_Familiar!parenro) And (edad_f <= Edad Or EsNulo(Edad)) And _
                ((Sexo = 1 And CBool(rs_Familiar!tersex)) Or (Sexo = 2 And Not CBool(rs_Familiar!tersex)) Or Sexo = 3) Then
                If (conyuge_trabaja = Trabaja_Conyuge Or Trabaja_Conyuge = 3) Then
                    If (Estudia = Fam_estudia Or Estudia = 3 Or (Not interesa_estu)) Then
                        If (InStr(1, Nivel_Estudio, Fam_niv_est) <> 0 Or EsNulo(Nivel_Estudio) Or (Not niv_est_interno)) Then
                            If (((Ayuda_Escolar = 1) And InStr(1, Nivel_Estudio, Fam_niv_est) <> 0) Or ((Ayuda_Escolar = 3) And _
                                    Not Retroactivo_Prenatal)) Then
                                Salida = True
                            End If
                        End If
                    End If
                End If
            End If
        End If
       
       
    If Salida Then
       'Controlo si tiene el concepto de familiares liquidado
       StrSql = " SELECT * FROM detliq WHERE cliqnro = " & cliqnro
       StrSql = StrSql & " AND concnro = " & concepto
       
       OpenRecordset StrSql, rs_Concepto
       
       Salida = Not rs_Concepto.EOF
        
       rs_Concepto.Close
    End If
       
    HayAsignacionesFliares = Salida
    
SiguienteFamiliar:
        rs_Familiar.MoveNext
    Loop
    
    Bien = True
    
'Cierro todo y libero
If Param_cur.State = adStateOpen Then Param_cur.Close
Set Param_cur = Nothing

If rs_estudio_actual.State = adStateOpen Then rs_estudio_actual.Close
Set rs_estudio_actual = Nothing

If rs_PeriodoEsc.State = adStateOpen Then rs_PeriodoEsc.Close
Set rs_PeriodoEsc = Nothing

If rs_Familiar.State = adStateOpen Then rs_Familiar.Close
Set rs_Familiar = Nothing

If rs_Nivest.State = adStateOpen Then rs_Nivest.Close
Set rs_Nivest = Nothing

If rs_Tercero.State = adStateOpen Then rs_Tercero.Close
Set rs_Tercero = Nothing

End Function


Public Function Calcular_Edad(ByVal Fecha As Date, ByVal Hasta As Date) As Integer
'...........................................................................
' Archivo       : edad.i                              fecha ini. : 20/01/92
' Nombre progr. :
' tipo programa : FGZ
' Descripcion   :
'...........................................................................
Dim años  As Integer
Dim ALaFecha As Date

    ALaFecha = C_Date(Hasta)
    
    años = Year(ALaFecha) - Year(Fecha)
    If Month(ALaFecha) < Month(Fecha) Then
       años = años - 1
    Else
        If Month(ALaFecha) = Month(Fecha) Then
            If Day(ALaFecha) < Day(Fecha) Then
                años = años - 1
            End If
        End If
    End If
    Calcular_Edad = años
End Function



